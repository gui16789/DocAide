
$ErrorActionPreference = "Stop"
$script:JobStopwatch = [Diagnostics.Stopwatch]::StartNew()
$script:CurrentWordPid = $null
$script:JobCompleted = $false
$script:JobFailed = $false
$script:JobFailureDetail = ""

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class RedheadNativeMethods {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Get-RuleValue {
  param($Object, [string]$Path, $Default)
  $current = $Object
  foreach ($part in $Path.Split(".")) {
    if ($null -eq $current) { return $Default }
    $property = $current.PSObject.Properties[$part]
    if ($null -eq $property) { return $Default }
    $current = $property.Value
  }
  if ($null -eq $current) { return $Default }
  if (($current -is [string]) -and [string]::IsNullOrWhiteSpace($current)) { return $Default }
  return $current
}

function Convert-MmToPt {
  param([double]$Mm)
  return $Mm * 72 / 25.4
}

function Convert-HexToWordColor {
  param([string]$Hex)
  $clean = ($Hex -replace "#", "").Trim()
  if ($clean.Length -ne 6) { return 0 }
  $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
  return $r + ($g * 256) + ($b * 65536)
}

function Convert-LineWidthPtToWordEnum {
  param([double]$Pt)
  $map = @(
    @{ Pt = 0.25; Enum = 2 },
    @{ Pt = 0.5; Enum = 4 },
    @{ Pt = 0.75; Enum = 6 },
    @{ Pt = 1.0; Enum = 8 },
    @{ Pt = 1.5; Enum = 12 },
    @{ Pt = 2.25; Enum = 18 },
    @{ Pt = 3.0; Enum = 24 },
    @{ Pt = 4.5; Enum = 36 },
    @{ Pt = 6.0; Enum = 48 }
  )
  $best = $map[0]
  foreach ($item in $map) {
    if ([Math]::Abs([double]$item.Pt - $Pt) -lt [Math]::Abs([double]$best.Pt - $Pt)) {
      $best = $item
    }
  }
  return [int]$best.Enum
}

function Convert-WordLineWidthToPt {
  param([int]$LineWidth)
  switch ($LineWidth) {
    2 { return 0.25 }
    4 { return 0.5 }
    6 { return 0.75 }
    8 { return 1.0 }
    12 { return 1.5 }
    18 { return 2.25 }
    24 { return 3.0 }
    36 { return 4.5 }
    48 { return 6.0 }
    default { return 0.0 }
  }
}

function Get-PageSizeSpec {
  param($Rules)
  $sizeName = ([string](Get-RuleValue $Rules "page.size" "A4")).Trim().ToUpperInvariant()
  switch ($sizeName) {
    "A4" {
      return @{
        Name = "A4"
        WidthMm = 210
        HeightMm = 297
        PaperSize = 7
      }
    }
    default {
      return @{
        Name = "A4"
        WidthMm = 210
        HeightMm = 297
        PaperSize = 7
      }
    }
  }
}

function Test-Near {
  param([double]$Actual, [double]$Expected, [double]$Tolerance = 0.75)
  return ([Math]::Abs($Actual - $Expected) -le $Tolerance)
}

function New-Check {
  param([string]$Label, [bool]$Passed, [string]$Detail = "")
  $status = if ($Passed) { "pass" } else { "fail" }
  if ([string]::IsNullOrWhiteSpace($Detail)) {
    return @{ Label = $Label; Status = $status }
  }
  return @{ Label = $Label; Status = $status; Detail = $Detail }
}

function New-Warn {
  param([string]$Label, [string]$Detail = "")
  if ([string]::IsNullOrWhiteSpace($Detail)) {
    return @{ Label = $Label; Status = "warn" }
  }
  return @{ Label = $Label; Status = "warn"; Detail = $Detail }
}

function Get-WordProcessId {
  param($WordApplication)
  try {
    $processId = [uint32]0
    $hwnd = [IntPtr]$WordApplication.Hwnd
    [void][RedheadNativeMethods]::GetWindowThreadProcessId($hwnd, [ref]$processId)
    if ($processId -gt 0) { return [int]$processId }
  } catch {}
  return $null
}

function Get-RecentWordProcessId {
  param([datetime]$StartedAt)
  try {
    $process = Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
      Where-Object { $_.StartTime -ge $StartedAt.AddSeconds(-5) } |
      Sort-Object StartTime -Descending |
      Select-Object -First 1
    if ($null -ne $process) { return [int]$process.Id }
  } catch {}
  return $null
}

function Write-ProcessPidFile {
  param([string]$Path, [int]$WordPid)
  if ([string]::IsNullOrWhiteSpace($Path)) { return }
  @{
    powershellPid = $PID
    wordPid = $WordPid
    createdAt = (Get-Date).ToString("o")
  } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Write-JobStatus {
  param([string]$Stage, [string]$Detail = "", [hashtable]$Extra = @{})
  if ([string]::IsNullOrWhiteSpace($StatusPath)) { return }
  try {
    $payload = @{
      stage = $Stage
      detail = $Detail
      updatedAt = (Get-Date).ToString("o")
      elapsedMs = [int64]$script:JobStopwatch.ElapsedMilliseconds
      powershellPid = $PID
      wordPid = $script:CurrentWordPid
    }
    if ($null -ne $Extra) {
      foreach ($key in $Extra.Keys) {
        $payload[$key] = $Extra[$key]
      }
    }
    $dir = Split-Path -Parent $StatusPath
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
      New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    $payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $StatusPath -Encoding UTF8
  } catch {
    # Status reporting must not interrupt document generation.
  }
}

function Get-ParagraphText {
  param($Paragraph)
  return ($Paragraph.Range.Text -replace "[`r`a]", "").Trim()
}

function Get-ParagraphRawText {
  param($Paragraph)
  return ($Paragraph.Range.Text -replace "[`r`a]", "")
}

function Get-TableCellText {
  param($Cell)
  return ($Cell.Range.Text -replace "[`r`a]", "")
}

function Repeat-Text {
  param([string]$Text, [int]$Count)
  if ($Count -le 0) { return "" }
  return $Text * $Count
}

function Convert-FullWidthDigits {
  param([string]$Text)
  $builder = [System.Text.StringBuilder]::new()
  foreach ($char in $Text.ToCharArray()) {
    $code = [int][char]$char
    if ($code -ge 0xFF10 -and $code -le 0xFF19) {
      [void]$builder.Append([char]($code - 0xFEE0))
    } else {
      [void]$builder.Append($char)
    }
  }
  return $builder.ToString()
}

function Normalize-SignatureText {
  param([string]$Text)
  if ($null -eq $Text) { return "" }
  $normalized = Convert-FullWidthDigits $Text
  $normalized = $normalized -replace "[\s　\u00A0]", ""
  $normalized = $normalized -replace "（", "(" -replace "）", ")"
  return $normalized.Trim()
}

function Test-DateParagraphText {
  param([string]$Text)
  $normalized = Normalize-SignatureText $Text
  return $normalized -match "^\d{4}年\d{1,2}月\d{1,2}日$"
}

function Test-DocumentNoText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
  $normalized = Convert-FullWidthDigits $Text
  $normalized = $normalized -replace "[\s　\u00A0]", ""
  if ($normalized -match "第") { return $false }
  if ($normalized -match "[\[\]【】（）()]") { return $false }
  if ($normalized -notmatch "^.+〔\d{4}〕[1-9]\d*号$") { return $false }
  return $true
}

function Get-ApproxTextWidthPt {
  param([string]$Text, [double]$FontSizePt)
  $units = 0.0
  foreach ($char in $Text.ToCharArray()) {
    $code = [int][char]$char
    if ($code -le 127) {
      $units += 0.5
    } else {
      $units += 1.0
    }
  }
  return $units * $FontSizePt
}

function Get-SignatureTextWidthPt {
  param([string]$Text, [double]$FontSizePt)
  if ($null -eq $Text) { return 0.0 }
  $normalized = $Text -replace "[\s　\u00A0]", ""
  $width = 0.0
  $previousClass = ""
  foreach ($char in $normalized.ToCharArray()) {
    $code = [int][char]$char
    $currentClass = if ($code -le 127) { "latin" } else { "cjk" }
    if (($previousClass -ne "") -and ($previousClass -ne $currentClass)) {
      $width += 0.25 * $FontSizePt
    }
    if ($currentClass -eq "latin") {
      $width += 0.5 * $FontSizePt
    } else {
      $width += 1.025 * $FontSizePt
    }
    $previousClass = $currentClass
  }
  return $width
}

function Get-TextCharUnits {
  param([string]$Text)
  if ($null -eq $Text) { return 0.0 }
  $units = 0.0
  foreach ($char in $Text.ToCharArray()) {
    if ([int][char]$char -le 127) {
      $units += 0.5
    } else {
      $units += 1.0
    }
  }
  return $units
}

function Resolve-ConfiguredPath {
  param([string]$ConfiguredPath)
  if ([string]::IsNullOrWhiteSpace($ConfiguredPath)) { return "" }
  if ([IO.Path]::IsPathRooted($ConfiguredPath)) { return $ConfiguredPath }
  $toolRoot = Split-Path -Parent $PSScriptRoot
  return Join-Path $toolRoot $ConfiguredPath
}

function Set-RangeFont {
  param(
    $Range,
    [string]$Font,
    [double]$SizePt,
    [string]$LatinFont = "",
    [int]$Color = -1,
    [double]$ScalePercent = 100,
    [double]$SpacingPt = 0
  )
  if ($Font) {
    $Range.Font.NameFarEast = $Font
    $Range.Font.Name = $Font
  }
  if ($LatinFont) {
    $Range.Font.NameAscii = $LatinFont
    $Range.Font.NameOther = $LatinFont
  }
  if ($SizePt -gt 0) { $Range.Font.Size = $SizePt }
  if ($Color -ge 0) { $Range.Font.Color = $Color }
  if ($ScalePercent -gt 0) { $Range.Font.Scaling = $ScalePercent }
  $Range.Font.Spacing = $SpacingPt
  $Range.Font.Bold = 0
  $Range.Font.Italic = 0
}

function Clear-ParagraphSpacingUnits {
  param($Paragraph)
  try {
    $Paragraph.Range.ParagraphFormat.SpaceBeforeAuto = 0
    $Paragraph.Range.ParagraphFormat.SpaceAfterAuto = 0
    $Paragraph.Range.ParagraphFormat.LineUnitBefore = 0
    $Paragraph.Range.ParagraphFormat.LineUnitAfter = 0
  } catch {}
}

function Set-BodyParagraph {
  param($Paragraph, $Rules)
  $bodyFont = Get-RuleValue $Rules "body.font" "仿宋_GB2312"
  $latinFont = Get-RuleValue $Rules "body.latinFont" "Times New Roman"
  $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $bodyColor = Convert-HexToWordColor (Get-RuleValue $Rules "body.color" "000000")
  $lineSpacing = [double](Get-RuleValue $Rules "body.lineSpacingPt" 28)
  $firstIndent = [double](Get-RuleValue $Rules "body.firstLineIndentPt" 32)

  Set-RangeFont $Paragraph.Range $bodyFont $bodySize $latinFont $bodyColor
  Clear-ParagraphSpacingUnits $Paragraph
  $Paragraph.Range.ParagraphFormat.Alignment = 0
  $Paragraph.Range.ParagraphFormat.LineSpacingRule = 4
  $Paragraph.Range.ParagraphFormat.LineSpacing = $lineSpacing
  $Paragraph.Range.ParagraphFormat.FirstLineIndent = $firstIndent
  $Paragraph.Range.ParagraphFormat.LeftIndent = 0
  $Paragraph.Range.ParagraphFormat.SpaceBefore = 0
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
}

function Set-TitleParagraph {
  param($Paragraph, $Rules)
  $titleFont = Get-RuleValue $Rules "title.font" "方正小标宋简体"
  $titleSize = [double](Get-RuleValue $Rules "title.sizePt" 22)
  $titleColor = Convert-HexToWordColor (Get-RuleValue $Rules "title.color" "000000")
  $lineSpacing = [double](Get-RuleValue $Rules "title.lineSpacingPt" 32)

  Set-RangeFont $Paragraph.Range $titleFont $titleSize "" $titleColor
  Clear-ParagraphSpacingUnits $Paragraph
  $Paragraph.Range.ParagraphFormat.Alignment = 1
  $Paragraph.Range.ParagraphFormat.LineSpacingRule = 4
  $Paragraph.Range.ParagraphFormat.LineSpacing = $lineSpacing
  $Paragraph.Range.ParagraphFormat.FirstLineIndent = 0
  $Paragraph.Range.ParagraphFormat.LeftIndent = 0
  $Paragraph.Range.ParagraphFormat.SpaceBefore = 0
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
}

function Set-MainRecipientParagraph {
  param($Paragraph, $Rules)
  $font = Get-RuleValue $Rules "mainRecipient.font" (Get-RuleValue $Rules "body.font" "仿宋_GB2312")
  $latinFont = Get-RuleValue $Rules "body.latinFont" "Times New Roman"
  $size = [double](Get-RuleValue $Rules "mainRecipient.sizePt" (Get-RuleValue $Rules "body.sizePt" 16))
  $lineSpacing = [double](Get-RuleValue $Rules "mainRecipient.lineSpacingPt" (Get-RuleValue $Rules "body.lineSpacingPt" 28))
  $color = Convert-HexToWordColor (Get-RuleValue $Rules "body.color" "000000")

  Set-RangeFont $Paragraph.Range $font $size $latinFont $color
  Clear-ParagraphSpacingUnits $Paragraph
  $Paragraph.Range.ParagraphFormat.Alignment = 0
  $Paragraph.Range.ParagraphFormat.LineSpacingRule = 4
  $Paragraph.Range.ParagraphFormat.LineSpacing = $lineSpacing
  $Paragraph.Range.ParagraphFormat.FirstLineIndent = 0
  $Paragraph.Range.ParagraphFormat.LeftIndent = 0
  $Paragraph.Range.ParagraphFormat.RightIndent = 0
  $Paragraph.Range.ParagraphFormat.SpaceBefore = 0
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
}

function Normalize-AttachmentExplanationText {
  param([string]$Text)
  if ($null -eq $Text) { return "" }
  $normalized = ($Text -replace "[`r`a]", "").Trim()
  $normalized = $normalized -replace "^附件\s*[:：]\s*", "附件："
  $normalized = $normalized -replace "^(\d+)\s*[．.]\s*", '$1.'
  $normalized = $normalized -replace "[。；;，,、]\s*$", ""
  return $normalized
}

function Get-AttachmentPrefixText {
  param([string]$Text, [bool]$FirstLine)
  $normalized = Normalize-AttachmentExplanationText $Text
  if ($FirstLine) {
    $afterLabel = $normalized -replace "^附件：", ""
    if ($afterLabel -match "^(\d+\.)") {
      return "附件：$($Matches[1])"
    }
    return "附件："
  }
  if ($normalized -match "^(\d+\.)") {
    return $Matches[1]
  }
  return ""
}

function Get-AttachmentFirstLineStartPrefix {
  param([string]$Text, [bool]$FirstLine)
  if ($FirstLine) { return "" }
  $normalized = Normalize-AttachmentExplanationText $Text
  if ($normalized -match "^\d+\.") { return "附件：" }
  return ""
}

function Get-AttachmentIndentSpec {
  param([string]$Text, $Rules, [bool]$IsFirst)
  $size = [double](Get-RuleValue $Rules "attachment.sizePt" (Get-RuleValue $Rules "body.sizePt" 16))
  $leftBlankChars = [double](Get-RuleValue $Rules "attachment.leftBlankChars" 2)
  $normalized = Normalize-AttachmentExplanationText $Text
  $prefix = Get-AttachmentPrefixText $normalized $IsFirst
  $firstLineStartPrefix = Get-AttachmentFirstLineStartPrefix $normalized $IsFirst
  $leftStart = $size * $leftBlankChars
  $firstLineStart = $leftStart + (Get-ApproxTextWidthPt $firstLineStartPrefix $size)
  $hangWidth = Get-ApproxTextWidthPt $prefix $size
  return @{
    LeftIndent = [double]($firstLineStart + $hangWidth)
    FirstLineIndent = [double](-1 * $hangWidth)
    FirstLineStart = [double]$firstLineStart
    HangWidth = [double]$hangWidth
    CharacterLeftIndent = [double]($leftBlankChars + (Get-TextCharUnits $firstLineStartPrefix) + (Get-TextCharUnits $prefix))
    CharacterFirstLineIndent = [double](-1 * (Get-TextCharUnits $prefix))
  }
}

function Find-AttachmentExplanationParagraphs {
  param($Document)
  $items = @()
  $startIndex = $null
  for ($i = 3; $i -le $Document.Paragraphs.Count; $i++) {
    $text = Get-ParagraphText $Document.Paragraphs.Item($i)
    if ($text -match "^\s*附件\s*[:：]") {
      $startIndex = $i
      break
    }
  }
  if ($null -eq $startIndex) { return @() }

  $items += [pscustomobject]@{ Paragraph = $Document.Paragraphs.Item($startIndex); Index = $startIndex; IsFirst = $true }
  for ($i = $startIndex + 1; $i -le $Document.Paragraphs.Count; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    $text = Get-ParagraphText $paragraph
    if ([string]::IsNullOrWhiteSpace($text)) { break }
    if (Test-DateParagraphText $text) { break }
    if ($text -match "^\s*\d+\s*[．.]") {
      $items += [pscustomobject]@{ Paragraph = $paragraph; Index = $i; IsFirst = $false }
      continue
    }
    break
  }
  return @($items)
}

function Set-AttachmentExplanationParagraph {
  param($Document, [int]$Index, $Rules, [bool]$IsFirst, [bool]$HasPreviousBlank)
  $font = Get-RuleValue $Rules "attachment.font" (Get-RuleValue $Rules "body.font" "仿宋_GB2312")
  $latinFont = Get-RuleValue $Rules "body.latinFont" "Times New Roman"
  $size = [double](Get-RuleValue $Rules "attachment.sizePt" (Get-RuleValue $Rules "body.sizePt" 16))
  $lineSpacing = [double](Get-RuleValue $Rules "attachment.lineSpacingPt" (Get-RuleValue $Rules "body.lineSpacingPt" 28))
  $spaceBeforeLines = [double](Get-RuleValue $Rules "attachment.spaceBeforeLines" 1)

  $Paragraph = $Document.Paragraphs.Item($Index)
  $normalized = Normalize-AttachmentExplanationText (Get-ParagraphText $Paragraph)
  if (-not [string]::IsNullOrWhiteSpace($normalized)) {
    $start = $Paragraph.Range.Start
    $Paragraph.Range.Text = "$normalized`r"
    $Paragraph = $Document.Range($start, $start).Paragraphs.Item(1)
  }

  $indentSpec = Get-AttachmentIndentSpec $normalized $Rules $IsFirst

  Set-RangeFont $Paragraph.Range $font $size $latinFont 0
  $Paragraph.Range.ParagraphFormat.Alignment = 0
  $Paragraph.Range.ParagraphFormat.LineSpacingRule = 4
  $Paragraph.Range.ParagraphFormat.LineSpacing = $lineSpacing
  try {
    $Paragraph.Range.ParagraphFormat.CharacterUnitLeftIndent = 0
    $Paragraph.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
  } catch {}
  $Paragraph.Range.ParagraphFormat.LeftIndent = [single]$indentSpec.LeftIndent
  $Paragraph.Range.ParagraphFormat.FirstLineIndent = [single]$indentSpec.FirstLineIndent
  $attachmentSpaceBefore = [double]0
  if ($IsFirst -and -not $HasPreviousBlank) {
    $attachmentSpaceBefore = [double]($lineSpacing * $spaceBeforeLines)
  }
  $Paragraph.Range.ParagraphFormat.SpaceBefore = [single]$attachmentSpaceBefore
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
}

function Apply-AttachmentExplanationStyle {
  param($Document, $Rules)
  $items = @(Find-AttachmentExplanationParagraphs $Document)
  if ($items.Count -eq 0) { return }

  for ($i = 0; $i -lt $items.Count; $i++) {
    $item = $items[$i]
    $hasPreviousBlank = $false
    if ([int]$item.Index -gt 1) {
      $hasPreviousBlank = Test-BlankParagraph $Document.Paragraphs.Item([int]$item.Index - 1)
    }
    Set-AttachmentExplanationParagraph $Document ([int]$item.Index) $Rules ([bool]$item.IsFirst) $hasPreviousBlank
  }
}

function Find-DocumentStructure {
  param($Document)
  $titleItems = @()
  $mainRecipient = $null
  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  $startIndex = [Math]::Max(3, $docNoIndex + 1)
  $limit = [Math]::Min($Document.Paragraphs.Count, 30)
  for ($i = $startIndex; $i -le $limit; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    $text = Get-ParagraphText $paragraph
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    if ($text -match "[：:]$") {
      $mainRecipient = [pscustomobject]@{
        Paragraph = $paragraph
        Index = $i
        Text = $text
      }
      break
    }
    if ($titleItems.Count -lt 6) {
      $titleItems += [pscustomobject]@{
        Paragraph = $paragraph
        Index = $i
        Text = $text
      }
    }
  }

  return @{
    TitleItems = @($titleItems)
    TitleParagraph = if ($titleItems.Count -gt 0) { $titleItems[0].Paragraph } else { $null }
    MainRecipient = $mainRecipient
  }
}

function Find-ParagraphIndexByText {
  param($Document, [string]$Text, [int]$StartIndex = 1, [int]$Limit = 30)
  $max = [Math]::Min($Document.Paragraphs.Count, $Limit)
  for ($i = $StartIndex; $i -le $max; $i++) {
    if ((Get-ParagraphText $Document.Paragraphs.Item($i)) -eq $Text) {
      return $i
    }
  }
  return $null
}

function Find-DocumentNoParagraphIndex {
  param($Document)
  $limit = [Math]::Min($Document.Paragraphs.Count, 20)
  for ($i = 1; $i -le $limit; $i++) {
    $text = Get-ParagraphText $Document.Paragraphs.Item($i)
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    $docNoPart = ($text -split "`t|签发人[:：]", 2)[0]
    if (($text -match "签发人[:：]") -or (Test-DocumentNoText $docNoPart)) {
      return $i
    }
  }
  return 2
}

function Clear-HeadersFooters {
  param($Document)
  foreach ($section in $Document.Sections) {
    foreach ($kind in 1..3) {
      try {
        $section.Headers.Item($kind).LinkToPrevious = $false
        $section.Headers.Item($kind).Range.Text = ""
        $section.Footers.Item($kind).LinkToPrevious = $false
        $section.Footers.Item($kind).Range.Text = ""
      } catch {
        # Some legacy documents expose only primary header/footer; ignore unavailable variants.
      }
    }
  }
}

function Get-VisibleText {
  param([string]$Text)
  if ($null -eq $Text) { return "" }
  return ($Text -replace "[\s　\u00A0\x07\x0B\x0C]", "")
}

function Test-BlankParagraph {
  param($Paragraph)
  try { if ($Paragraph.Range.Tables.Count -gt 0) { return $false } } catch {}
  try { if ($Paragraph.Range.InlineShapes.Count -gt 0) { return $false } } catch {}
  try { if ($Paragraph.Range.ShapeRange.Count -gt 0) { return $false } } catch {}
  return [string]::IsNullOrWhiteSpace((Get-VisibleText $Paragraph.Range.Text))
}

function Remove-TrailingBreaksFromParagraph {
  param($Paragraph)
  $text = $Paragraph.Range.Text
  if ([string]::IsNullOrEmpty($text) -or $text -notmatch "[\x0B\x0C]") { return $false }

  $lastVisibleIndex = -1
  for ($i = $text.Length - 1; $i -ge 0; $i--) {
    if (-not [string]::IsNullOrWhiteSpace((Get-VisibleText ([string]$text[$i])))) {
      $lastVisibleIndex = $i
      break
    }
  }
  if ($lastVisibleIndex -lt 0 -or $lastVisibleIndex -ge ($text.Length - 1)) { return $false }

  $tail = $text.Substring($lastVisibleIndex + 1)
  if ($tail -notmatch "[\x0B\x0C]") { return $false }

  $Paragraph.Range.Text = $text.Substring(0, $lastVisibleIndex + 1).TrimEnd() + "`r"
  return $true
}

function Remove-TrailingBlankParagraph {
  param($Document)
  if ($Document.Paragraphs.Count -le 1) {
    return @{ Changed = $false; Deleted = 0 }
  }

  $lastParagraph = $Document.Paragraphs.Item($Document.Paragraphs.Count)
  if (-not (Test-BlankParagraph $lastParagraph)) {
    return @{ Changed = $false; Deleted = 0 }
  }

  $beforeCount = $Document.Paragraphs.Count
  $lastParagraph.Range.Delete() | Out-Null
  if ($Document.Paragraphs.Count -lt $beforeCount) {
    return @{ Changed = $true; Deleted = 1 }
  }

  if ($Document.Paragraphs.Count -le 1) {
    return @{ Changed = $false; Deleted = 0 }
  }

  $previousParagraph = $Document.Paragraphs.Item($Document.Paragraphs.Count - 1)
  if (-not (Test-BlankParagraph $previousParagraph)) {
    return @{ Changed = $false; Deleted = 0 }
  }

  $previousText = $previousParagraph.Range.Text
  if ($previousText -notmatch "[\x0B\x0C]") {
    return @{ Changed = $false; Deleted = 0 }
  }

  $beforeCount = $Document.Paragraphs.Count
  $previousParagraph.Range.Delete() | Out-Null
  return @{ Changed = ($Document.Paragraphs.Count -lt $beforeCount); Deleted = 1 }
}

function Trim-TrailingBlankContent {
  param($Document, $Rules)
  if (-not [bool](Get-RuleValue $Rules "cleanup.trimTrailingBlankPages" $true)) {
    return @{
      Enabled = $false
      BlankParagraphsDeleted = 0
      TrailingBreaksDeleted = 0
      PagesBefore = $Document.ComputeStatistics(2)
      PagesAfter = $Document.ComputeStatistics(2)
      PagesRemoved = 0
    }
  }

  $Document.Repaginate()
  $pagesBefore = [int]$Document.ComputeStatistics(2)
  $blankParagraphsDeleted = 0
  $trailingBreaksDeleted = 0
  $guard = 0
  $changed = $true

  while ($changed -and $guard -lt 500) {
    $guard++
    $changed = $false

    while ($Document.Paragraphs.Count -gt 1) {
      $result = Remove-TrailingBlankParagraph $Document
      if (-not [bool]$result.Changed) { break }
      $blankParagraphsDeleted += [int]$result.Deleted
      $changed = $true
    }

    if ($Document.Paragraphs.Count -gt 0) {
      $lastParagraph = $Document.Paragraphs.Item($Document.Paragraphs.Count)
      if (Remove-TrailingBreaksFromParagraph $lastParagraph) {
        $trailingBreaksDeleted++
        $changed = $true
      }
    }
  }

  $Document.Repaginate()
  $pagesAfter = [int]$Document.ComputeStatistics(2)
  return @{
    Enabled = $true
    BlankParagraphsDeleted = $blankParagraphsDeleted
    TrailingBreaksDeleted = $trailingBreaksDeleted
    PagesBefore = $pagesBefore
    PagesAfter = $pagesAfter
    PagesRemoved = [Math]::Max(0, $pagesBefore - $pagesAfter)
  }
}

function Apply-PageSetup {
  param($Document, $Rules)
  $pageSpec = Get-PageSizeSpec $Rules
  $top = Convert-MmToPt ([double](Get-RuleValue $Rules "page.topMarginMm" 37))
  $bottom = Convert-MmToPt ([double](Get-RuleValue $Rules "page.bottomMarginMm" 35))
  $left = Convert-MmToPt ([double](Get-RuleValue $Rules "page.leftMarginMm" 28))
  $right = Convert-MmToPt ([double](Get-RuleValue $Rules "page.rightMarginMm" 26))
  $pageNumberDistance = Convert-MmToPt ([double](Get-RuleValue $Rules "pageNumber.distanceBelowContentMm" 17.5))
  $footerDistance = [Math]::Max(0, $bottom - $pageNumberDistance)

  foreach ($section in $Document.Sections) {
    $section.PageSetup.PaperSize = [int]$pageSpec.PaperSize
    $section.PageSetup.PageWidth = Convert-MmToPt ([double]$pageSpec.WidthMm)
    $section.PageSetup.PageHeight = Convert-MmToPt ([double]$pageSpec.HeightMm)
    $section.PageSetup.TopMargin = $top
    $section.PageSetup.BottomMargin = $bottom
    $section.PageSetup.LeftMargin = $left
    $section.PageSetup.RightMargin = $right
    $section.PageSetup.FooterDistance = $footerDistance
    $section.PageSetup.OddAndEvenPagesHeaderFooter = $true
    $section.PageSetup.DifferentFirstPageHeaderFooter = $false
    $section.PageSetup.VerticalAlignment = 0
  }
}

function Replace-InitialTitle {
  param($Document, [string]$Title)
  if ([string]::IsNullOrWhiteSpace($Title)) { return }

  $start = $null
  $end = $null
  $limit = [Math]::Min($Document.Paragraphs.Count, 18)
  for ($i = 1; $i -le $limit; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    $text = Get-ParagraphText $paragraph
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    if ($null -eq $start) { $start = $paragraph.Range.Start }
    if ($text -match "[：:]$" -and $end -ne $null) { break }
    $end = $paragraph.Range.End
  }

  if ($null -ne $start -and $null -ne $end -and $end -gt $start) {
    $range = $Document.Range($start, $end)
    $range.Text = "$Title`r"
  }
}

function Test-FontNameMatches {
  param([string]$Actual, [string]$Expected)
  if ([string]::IsNullOrWhiteSpace($Actual) -or [string]::IsNullOrWhiteSpace($Expected)) { return $false }
  $actualClean = ($Actual -replace "[-_ ]?GB2312", "").Trim()
  $expectedClean = ($Expected -replace "[-_ ]?GB2312", "").Trim()
  return ($Actual -eq $Expected) -or ($actualClean -eq $expectedClean)
}

function Get-PageNumberAlignment {
  param($Rules, [string]$Path, [int]$Default)
  $value = ([string](Get-RuleValue $Rules $Path "")).Trim().ToLowerInvariant()
  switch ($value) {
    "left" { return 0 }
    "center" { return 1 }
    "centre" { return 1 }
    "right" { return 2 }
    "左" { return 0 }
    "居左" { return 0 }
    "居中" { return 1 }
    "右" { return 2 }
    "居右" { return 2 }
    default { return $Default }
  }
}

function Test-PageNumberIndentForAlignment {
  param($Footer, [int]$Alignment, [double]$ExpectedIndent)
  if ($Alignment -eq 2) {
    return (Test-Near ([double]$Footer.Range.ParagraphFormat.RightIndent) $ExpectedIndent 1.0) -and
      (Test-Near ([double]$Footer.Range.ParagraphFormat.LeftIndent) 0 1.0)
  }
  if ($Alignment -eq 0) {
    return (Test-Near ([double]$Footer.Range.ParagraphFormat.LeftIndent) $ExpectedIndent 1.0) -and
      (Test-Near ([double]$Footer.Range.ParagraphFormat.RightIndent) 0 1.0)
  }
  return (Test-Near ([double]$Footer.Range.ParagraphFormat.LeftIndent) 0 1.0) -and
    (Test-Near ([double]$Footer.Range.ParagraphFormat.RightIndent) 0 1.0)
}

function Add-PageFieldFooter {
  param($Document, $Footer, [int]$Alignment, $Rules)
  $Footer.LinkToPrevious = $false
  try { $Footer.PageNumbers.NumberStyle = 0 } catch {}
  $format = Get-RuleValue $Rules "pageNumber.format" "- {page} -"
  $parts = $format -split "\{page\}", 2
  $prefix = if ($parts.Count -ge 1) { $parts[0] } else { "- " }
  $suffix = if ($parts.Count -ge 2) { $parts[1] } else { " -" }
  $Footer.Range.Text = "$prefix$suffix"
  $range = $Footer.Range
  if ($range.End -gt $range.Start) {
    $range.End = $range.End - 1
  }
  $range.ParagraphFormat.Alignment = $Alignment
  $pageFont = Get-RuleValue $Rules "pageNumber.font" "宋体"
  $pageSize = [double](Get-RuleValue $Rules "pageNumber.sizePt" 14)
  $blankChars = [double](Get-RuleValue $Rules "pageNumber.blankChars" 1)
  $range.ParagraphFormat.LeftIndent = 0
  $range.ParagraphFormat.RightIndent = 0
  if ($Alignment -eq 2) {
    $range.ParagraphFormat.RightIndent = [single]($pageSize * $blankChars)
  } elseif ($Alignment -eq 0) {
    $range.ParagraphFormat.LeftIndent = [single]($pageSize * $blankChars)
  }
  Set-RangeFont $range $pageFont $pageSize $pageFont 0

  $fieldRange = $Footer.Range.Duplicate
  $fieldRange.Start = $fieldRange.Start + $prefix.Length
  $fieldRange.End = $fieldRange.Start
  $null = $Footer.Range.Fields.Add($fieldRange, -1, "PAGE \* Arabic", $false)
  Set-RangeFont $Footer.Range $pageFont $pageSize $pageFont 0
}

function Apply-BodyPageNumbers {
  param($Document, $ImprintInfo, $Rules)
  if (-not [bool](Get-RuleValue $Rules "pageNumber.enabled" $true)) { return }
  $bodySectionCount = if (($null -ne $ImprintInfo) -and $ImprintInfo.ContainsKey("BodySectionCount")) {
    [int]$ImprintInfo.BodySectionCount
  } else {
    $Document.Sections.Count
  }
  $blankInserted = (($null -ne $ImprintInfo) -and $ImprintInfo.ContainsKey("BlankInserted") -and [bool]$ImprintInfo.BlankInserted)
  $sectionCount = if ($blankInserted) { $bodySectionCount } else { $Document.Sections.Count }
  $oddAlignment = Get-PageNumberAlignment $Rules "pageNumber.oddAlign" 2
  $evenAlignment = Get-PageNumberAlignment $Rules "pageNumber.evenAlign" 0
  for ($i = 1; $i -le $sectionCount; $i++) {
    $section = $Document.Sections.Item($i)
    $section.PageSetup.OddAndEvenPagesHeaderFooter = $true
    foreach ($kind in 1..3) {
      try {
        $section.Footers.Item($kind).LinkToPrevious = $false
        $section.Footers.Item($kind).Range.Text = ""
      } catch {}
    }
    $section.Footers.Item(1).PageNumbers.RestartNumberingAtSection = $false
    if ($i -eq 1) {
      $section.Footers.Item(1).PageNumbers.RestartNumberingAtSection = $true
      $section.Footers.Item(1).PageNumbers.StartingNumber = 1
    }
    Add-PageFieldFooter $Document $section.Footers.Item(1) $oddAlignment $Rules
    Add-PageFieldFooter $Document $section.Footers.Item(3) $evenAlignment $Rules
  }
}

function Clear-SectionPageFurniture {
  param($Section)
  foreach ($kind in 1..3) {
    try {
      $Section.Headers.Item($kind).LinkToPrevious = $false
      $Section.Headers.Item($kind).Range.Text = ""
      $Section.Footers.Item($kind).LinkToPrevious = $false
      $Section.Footers.Item($kind).Range.Text = ""
    } catch {}
  }
}

function Set-CollapsedParagraphAfterRange {
  param($Document, $Range)
  try {
    $paragraphRange = $Document.Range($Range.End, $Range.End)
    $paragraph = $paragraphRange.Paragraphs.Item(1)
    $paragraph.Range.Font.Size = 1
    $paragraph.Range.Font.Hidden = 1
    $paragraph.Range.ParagraphFormat.LineSpacingRule = 4
    $paragraph.Range.ParagraphFormat.LineSpacing = 1
    $paragraph.Range.ParagraphFormat.SpaceBefore = 0
    $paragraph.Range.ParagraphFormat.SpaceAfter = 0
    $paragraph.Range.ParagraphFormat.FirstLineIndent = 0
  } catch {}
}

function Remove-NamedShapes {
  param($Document, [string]$Name)
  $deleteShapes = @()
  foreach ($shape in $Document.Shapes) {
    try {
      if ($shape.Name -eq $Name) { $deleteShapes += $shape }
    } catch {}
  }
  foreach ($shape in $deleteShapes) {
    try { $shape.Delete() | Out-Null } catch {}
  }
}

function Get-ExpectedWordColorHex {
  param([int]$WordColor)
  $r = $WordColor -band 255
  $g = ($WordColor -shr 8) -band 255
  $b = ($WordColor -shr 16) -band 255
  return ("{0:X2}{1:X2}{2:X2}" -f $r, $g, $b)
}

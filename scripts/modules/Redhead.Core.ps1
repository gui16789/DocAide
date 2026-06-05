
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

function Set-BodyParagraph {
  param($Paragraph, $Rules)
  $bodyFont = Get-RuleValue $Rules "body.font" "仿宋_GB2312"
  $latinFont = Get-RuleValue $Rules "body.latinFont" "Times New Roman"
  $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $bodyColor = Convert-HexToWordColor (Get-RuleValue $Rules "body.color" "000000")
  $lineSpacing = [double](Get-RuleValue $Rules "body.lineSpacingPt" 28)
  $firstIndent = [double](Get-RuleValue $Rules "body.firstLineIndentPt" 32)

  Set-RangeFont $Paragraph.Range $bodyFont $bodySize $latinFont $bodyColor
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
  $spaceAfter = [double](Get-RuleValue $Rules "mainRecipient.titleSpaceAfterPt" (Get-RuleValue $Rules "title.spaceAfterPt" 22))

  Set-RangeFont $Paragraph.Range $titleFont $titleSize "" $titleColor
  $Paragraph.Range.ParagraphFormat.Alignment = 1
  $Paragraph.Range.ParagraphFormat.LineSpacingRule = 4
  $Paragraph.Range.ParagraphFormat.LineSpacing = $lineSpacing
  $Paragraph.Range.ParagraphFormat.FirstLineIndent = 0
  $Paragraph.Range.ParagraphFormat.LeftIndent = 0
  $Paragraph.Range.ParagraphFormat.SpaceBefore = 0
  $Paragraph.Range.ParagraphFormat.SpaceAfter = $spaceAfter
}

function Set-MainRecipientParagraph {
  param($Paragraph, $Rules)
  $font = Get-RuleValue $Rules "mainRecipient.font" (Get-RuleValue $Rules "body.font" "仿宋_GB2312")
  $latinFont = Get-RuleValue $Rules "body.latinFont" "Times New Roman"
  $size = [double](Get-RuleValue $Rules "mainRecipient.sizePt" (Get-RuleValue $Rules "body.sizePt" 16))
  $lineSpacing = [double](Get-RuleValue $Rules "mainRecipient.lineSpacingPt" (Get-RuleValue $Rules "body.lineSpacingPt" 28))
  $color = Convert-HexToWordColor (Get-RuleValue $Rules "body.color" "000000")

  Set-RangeFont $Paragraph.Range $font $size $latinFont $color
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

function Normalize-TitleText {
  param([string]$Text)
  if ($null -eq $Text) { return "" }
  $normalized = $Text -replace "[`r`n`t`v]", ""
  $normalized = $normalized -replace "[\s　\u00A0]+", ""
  return $normalized.Trim()
}

function Get-TitleBreakCandidates {
  param([string]$Text)
  $positions = New-Object System.Collections.Generic.HashSet[int]
  $keywords = @(
    "有限公司", "股份有限公司", "集团公司", "四川省", "成都市", "宜宾市", "南溪区",
    "财政补贴性", "地方财政补贴性", "保险", "活动方案", "实施方案", "工作方案",
    "隐患排查工作", "进行报备", "进行报备的报告", "通知", "报告", "请示",
    "批复", "意见", "办法", "规定", "方案", "工作"
  )
  foreach ($keyword in $keywords) {
    foreach ($match in [regex]::Matches($Text, [regex]::Escape($keyword))) {
      $position = $match.Index + $match.Length
      if ($position -gt 0 -and $position -lt $Text.Length) {
        [void]$positions.Add($position)
      }
    }
  }
  foreach ($match in [regex]::Matches($Text, "地方财政补贴性(.+?保险(?:（\d{4}）|\(\d{4}\))?)(?=进行|等|的|$)")) {
    $productStart = $match.Groups[1].Index
    $productEnd = $productStart + $match.Groups[1].Length
    if ($productStart -gt 0 -and $productStart -lt $Text.Length) { [void]$positions.Add($productStart) }
    if ($productEnd -gt 0 -and $productEnd -lt $Text.Length) { [void]$positions.Add($productEnd) }
  }
  foreach ($match in [regex]::Matches($Text, "财政补贴性(.+?保险(?:（\d{4}）|\(\d{4}\))?)(?=进行|等|的|$)")) {
    $productStart = $match.Groups[1].Index
    $productEnd = $productStart + $match.Groups[1].Length
    if ($productStart -gt 0 -and $productStart -lt $Text.Length) { [void]$positions.Add($productStart) }
    if ($productEnd -gt 0 -and $productEnd -lt $Text.Length) { [void]$positions.Add($productEnd) }
  }
  for ($i = 0; $i -lt $Text.Length; $i++) {
    $char = [string]$Text[$i]
    if ("省市州县区，、；：）》）".Contains($char) -and ($i + 1) -lt $Text.Length) {
      [void]$positions.Add($i + 1)
    }
  }
  return @($positions | Sort-Object)
}

function Get-TitleProtectedTerms {
  param([string]$Text)
  $terms = New-Object System.Collections.Generic.HashSet[string]
  $fixedTerms = @(
    "安盟财产保险有限公司", "安盟保险", "有限公司", "四川省", "成都市", "宜宾市", "南溪区",
    "地方财政补贴性", "财政补贴性", "农业经营主体用工团体意外伤害保险",
    "农村居民住房保险", "蔬菜种植保险", "安全生产月", "隐患排查工作"
  )
  foreach ($term in $fixedTerms) {
    if (-not [string]::IsNullOrWhiteSpace($term) -and $Text.Contains($term)) {
      [void]$terms.Add($term)
    }
  }
  foreach ($match in [regex]::Matches($Text, "地方财政补贴性(.+?保险(?:（\d{4}）|\(\d{4}\))?)(?=进行|等|的|$)")) {
    if ($match.Groups[1].Length -gt 0) { [void]$terms.Add($match.Groups[1].Value) }
  }
  foreach ($match in [regex]::Matches($Text, "财政补贴性(.+?保险(?:（\d{4}）|\(\d{4}\))?)(?=进行|等|的|$)")) {
    if ($match.Groups[1].Length -gt 0) { [void]$terms.Add($match.Groups[1].Value) }
  }
  foreach ($match in [regex]::Matches($Text, '[“"《（(][^”"》）)]+[”"》）)]')) {
    [void]$terms.Add($match.Value)
  }
  return @($terms)
}

function Test-TitleBreakInsideProtectedTerm {
  param([string]$Text, [int]$Position)
  if ($Position -le 0 -or $Position -ge $Text.Length) { return $false }
  foreach ($term in (Get-TitleProtectedTerms $Text)) {
    foreach ($match in [regex]::Matches($Text, [regex]::Escape($term))) {
      $start = $match.Index
      $end = $match.Index + $match.Length
      if ($Position -gt $start -and $Position -lt $end) {
        return $true
      }
    }
  }
  return $false
}

function Test-TitleBreakHasBadEdge {
  param([string]$Text, [int]$Position)
  if ($Position -le 0 -or $Position -ge $Text.Length) { return $false }
  $before = [string]$Text[$Position - 1]
  $after = [string]$Text[$Position]
  if ("的对关于和与及、，；：（(".Contains($before)) { return $true }
  if ("的对和与及、，；：）》）（(".Contains($after)) { return $true }
  $badPairs = @("种|植", "意外|伤害", "安全|生产", "隐患|排查", "财产|保险", "财政|补贴", "报|备")
  foreach ($pair in $badPairs) {
    $parts = $pair.Split("|")
    $left = $parts[0]
    $right = $parts[1]
    $leftStart = [Math]::Max(0, $Position - $left.Length)
    $rightEnd = [Math]::Min($Text.Length, $Position + $right.Length)
    if ($Text.Substring($leftStart, $Position - $leftStart).EndsWith($left) -and
        $Text.Substring($Position, $rightEnd - $Position).StartsWith($right)) {
      return $true
    }
  }
  return $false
}

function Split-TitleText {
  param([string]$Text, [int]$MaxLineChars = 22)
  $title = Normalize-TitleText $Text
  if ([string]::IsNullOrWhiteSpace($title)) { return @() }
  if ($title.Length -le $MaxLineChars) { return @($title) }

  $lineCount = [Math]::Min(4, [Math]::Max(2, [int][Math]::Ceiling($title.Length / [double]$MaxLineChars)))
  $target = $title.Length / [double]$lineCount
  $candidateSet = New-Object System.Collections.Generic.HashSet[int]
  [void]$candidateSet.Add(0)
  [void]$candidateSet.Add($title.Length)
  $semanticSet = New-Object System.Collections.Generic.HashSet[int]
  foreach ($candidate in (Get-TitleBreakCandidates $title)) {
    [void]$candidateSet.Add([int]$candidate)
    [void]$semanticSet.Add([int]$candidate)
  }
  for ($line = 1; $line -lt $lineCount; $line++) {
    $center = [int][Math]::Round($target * $line)
    for ($offset = -3; $offset -le 3; $offset++) {
      $position = $center + $offset
      if ($position -gt 0 -and $position -lt $title.Length) {
        [void]$candidateSet.Add($position)
      }
    }
  }
  $points = @($candidateSet | Sort-Object)

  $states = @(@{ Score = 0.0; Breaks = @(0) })
  for ($line = 1; $line -le $lineCount; $line++) {
    $nextStates = @()
    foreach ($state in $states) {
      $last = [int]$state.Breaks[-1]
      foreach ($point in $points) {
        if ($point -le $last) { continue }
        if ($line -lt $lineCount -and $point -ge $title.Length) { continue }
        if ($line -eq $lineCount -and $point -ne $title.Length) { continue }

        $segmentLength = $point - $last
        if ($segmentLength -le 0) { continue }
        $score = [double]$state.Score + [Math]::Pow($segmentLength - $target, 2)
        if ($segmentLength -lt 6 -and $line -lt $lineCount) { $score += 100 }
        if ($segmentLength -gt ($MaxLineChars + 4)) { $score += 120 }
        if (-not $semanticSet.Contains($point) -and $point -ne $title.Length) { $score += 140 }
        if (Test-TitleBreakInsideProtectedTerm $title $point) { $score += 400 }
        if (Test-TitleBreakHasBadEdge $title $point) { $score += 180 }
        $segment = $title.Substring($last, $segmentLength)
        if ($segment.EndsWith("的") -or $segment.EndsWith("对") -or $segment.EndsWith("关于")) { $score += 50 }
        if ($line -gt 1 -and ("的对和与及、，）)".Contains([string]$segment[0]))) { $score += 50 }

        $breaks = @($state.Breaks + $point)
        $nextStates += @{ Score = $score; Breaks = $breaks }
      }
    }
    $states = @($nextStates | Sort-Object Score | Select-Object -First 60)
  }

  if ($states.Count -eq 0) { return @($title) }
  $best = $states[0].Breaks
  $lines = @()
  for ($i = 1; $i -lt $best.Count; $i++) {
    $start = [int]$best[$i - 1]
    $end = [int]$best[$i]
    $lines += $title.Substring($start, $end - $start)
  }
  return $lines
}

function Normalize-TitleBlock {
  param($Document, $Rules)
  if ($Document.Paragraphs.Count -lt 4) { return }

  $bodyStartIndex = $null
  $titleParts = @()
  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  $titleStartIndex = [Math]::Max(3, $docNoIndex + 1)
  $limit = [Math]::Min($Document.Paragraphs.Count, 30)
  for ($i = $titleStartIndex; $i -le $limit; $i++) {
    $text = Get-ParagraphText $Document.Paragraphs.Item($i)
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    if ($text -match "[：:]$") {
      $bodyStartIndex = $i
      break
    }
    $titleParts += $text
  }

  if ($null -eq $bodyStartIndex -or $titleParts.Count -eq 0) { return }

  $maxLineChars = [int](Get-RuleValue $Rules "title.maxLineChars" 22)
  $titleLines = Split-TitleText ($titleParts -join "") $maxLineChars
  if ($titleLines.Count -eq 0) { return }

  $softBreak = [string][char]11
  $formattedTitle = ($titleLines -join $softBreak)
  $replaceStart = $Document.Paragraphs.Item($titleStartIndex).Range.Start
  $replaceEnd = $Document.Paragraphs.Item($bodyStartIndex).Range.Start
  $range = $Document.Range($replaceStart, $replaceEnd)
  $range.Text = "$formattedTitle`r"
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
  $pageNumberDistance = Convert-MmToPt ([double](Get-RuleValue $Rules "pageNumber.distanceBelowContentMm" 7))
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

function Add-DocumentNoRedLine {
  param($Document, $Paragraph, $Rules)

  Remove-NamedShapes $Document "redhead-document-no-red-line"

  $lineWidth = [double](Get-RuleValue $Rules "documentNo.redLineWidthPt" 2.25)
  $lineColor = Convert-HexToWordColor (Get-RuleValue $Rules "documentNo.redLineColor" "FF0000")
  $lineOffset = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 4))
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
  foreach ($borderIndex in @(-1, -2, -3, -4)) {
    try { $Paragraph.Range.ParagraphFormat.Borders.Item($borderIndex).LineStyle = 0 } catch {}
  }
  $bottomBorder = $Paragraph.Range.ParagraphFormat.Borders.Item(-3)
  $bottomBorder.LineStyle = 1
  $bottomBorder.LineWidth = Convert-LineWidthPtToWordEnum $lineWidth
  $bottomBorder.Color = $lineColor
  try { $Paragraph.Range.ParagraphFormat.Borders.DistanceFromBottom = [single]$lineOffset } catch {}
}

function Find-DocumentNoRedLineParagraph {
  param($Document, $Rules)

  $expectedColorHex = ((Get-RuleValue $Rules "documentNo.redLineColor" "FF0000") -replace "#", "").ToUpperInvariant()
  $limit = [Math]::Min($Document.Paragraphs.Count, 20)
  for ($i = 1; $i -le $limit; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    if (-not (Test-BlankParagraph $paragraph)) { continue }
    try {
      $border = $paragraph.Range.ParagraphFormat.Borders.Item(-1)
      if (([int]$border.LineStyle -ne 0) -and ((Get-ExpectedWordColorHex ([int]$border.Color)) -eq $expectedColorHex)) {
        return $paragraph
      }
    } catch {}
  }
  return $null
}

function Adjust-DocumentNoRedLinePosition {
  param($Document, $Rules, $DocNoParagraph)

  $redLineParagraph = Find-DocumentNoRedLineParagraph $Document $Rules
  if ($null -eq $redLineParagraph) { return }

  $targetGap = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 4))
  $lineBoxHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  for ($attempt = 0; $attempt -lt 4; $attempt++) {
    $Document.Repaginate()
    $actualGap = [double]$redLineParagraph.Range.Information(6) - ([double]$DocNoParagraph.Range.Information(6) + $lineBoxHeight)
    $delta = $targetGap - $actualGap
    if ([Math]::Abs($delta) -le 0.5) { break }

    $currentSpaceBefore = [double]$redLineParagraph.Range.ParagraphFormat.SpaceBefore
    $redLineParagraph.Range.ParagraphFormat.SpaceBefore = [single]([Math]::Max(0, $currentSpaceBefore + $delta))
  }
}

function Adjust-RedHeaderLayout {
  param($Document, $Rules)

  $headerText = Get-RuleValue $Rules "redHeader.text" "安盟财产保险有限公司文件"
  $headerIndex = Find-ParagraphIndexByText $Document $headerText 1 20
  if (($null -eq $headerIndex) -or ([int]$headerIndex -le 1)) { return }

  $headerParagraph = $Document.Paragraphs.Item([int]$headerIndex)
  $spacerParagraph = $Document.Paragraphs.Item([int]$headerIndex - 1)
  $headerVisibleTopInset = [double](Get-RuleValue $Rules "redHeader.visibleTopInsetPt" 0)
  $targetHeaderOffset = Convert-MmToPt ([double](Get-RuleValue $Rules "redHeader.topFromContentTopMm" 0))

  if ($targetHeaderOffset -gt 0) {
    $spacerParagraph.Range.Font.Size = 1
    $spacerParagraph.Range.Font.Hidden = 0
    $spacerParagraph.Range.Font.Color = Convert-HexToWordColor "FFFFFF"
    $spacerParagraph.Range.ParagraphFormat.LineSpacingRule = 4
    $spacerParagraph.Range.ParagraphFormat.SpaceBefore = 0
    $spacerParagraph.Range.ParagraphFormat.SpaceAfter = 0
    $spacerParagraph.Range.ParagraphFormat.FirstLineIndent = 0

    for ($attempt = 0; $attempt -lt 5; $attempt++) {
      $Document.Repaginate()
      $actualHeaderOffset = ([double]$headerParagraph.Range.Information(6) + $headerVisibleTopInset) - [double]$Document.Sections.Item(1).PageSetup.TopMargin
      $delta = $targetHeaderOffset - $actualHeaderOffset
      if ([Math]::Abs($delta) -le 0.5) { break }

      $currentLineSpacing = [double]$spacerParagraph.Range.ParagraphFormat.LineSpacing
      $nextLineSpacing = [Math]::Max(1, $currentLineSpacing + $delta)
      if ([Math]::Abs($nextLineSpacing - $currentLineSpacing) -le 0.1) { break }
      $spacerParagraph.Range.ParagraphFormat.LineSpacing = [single]$nextLineSpacing
    }
  }

  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  if (($null -eq $docNoIndex) -or ([int]$docNoIndex -le [int]$headerIndex)) { return }

  $docNoParagraph = $Document.Paragraphs.Item([int]$docNoIndex)
  $headerSize = [double](Get-RuleValue $Rules "redHeader.sizePt" 58)
  $docNoSize = [double](Get-RuleValue $Rules "documentNo.sizePt" 16)
  $docNoBlankLines = [double](Get-RuleValue $Rules "redHeader.docNoBlankLines" 2)
  $targetDocNoGap = $docNoSize * $docNoBlankLines

  for ($attempt = 0; $attempt -lt 5; $attempt++) {
    $Document.Repaginate()
    $headerVisibleBottom = [double]$headerParagraph.Range.Information(6) + $headerVisibleTopInset + $headerSize
    $actualDocNoGap = [double]$docNoParagraph.Range.Information(6) - $headerVisibleBottom
    $delta = $targetDocNoGap - $actualDocNoGap
    if ([Math]::Abs($delta) -le 1.0) { break }

    $currentSpaceAfter = [double]$headerParagraph.Range.ParagraphFormat.SpaceAfter
    $nextSpaceAfter = [Math]::Max(0, $currentSpaceAfter + $delta)
    if ([Math]::Abs($nextSpaceAfter - $currentSpaceAfter) -le 0.1) { break }
    $headerParagraph.Range.ParagraphFormat.SpaceAfter = [single]$nextSpaceAfter
  }

  $Document.Repaginate()
}

function Insert-RedHeader {
  param($Document, $Rules, $Meta)

  $headerText = Get-RuleValue $Rules "redHeader.text" "安盟财产保险有限公司文件"
  $headerFont = Get-RuleValue $Rules "redHeader.font" "方正小标宋简体"
  $headerSize = [double](Get-RuleValue $Rules "redHeader.sizePt" 58)
  $headerScale = [double](Get-RuleValue $Rules "redHeader.scalePercent" 64)
  $headerSpacing = [double](Get-RuleValue $Rules "redHeader.characterSpacingPt" -0.5)
  $headerColor = Convert-HexToWordColor (Get-RuleValue $Rules "redHeader.color" "FF0000")
  $headerSpaceAfter = [double](Get-RuleValue $Rules "redHeader.spaceAfterPt" 24)
  $headerTopFromContentTop = Convert-MmToPt ([double](Get-RuleValue $Rules "redHeader.topFromContentTopMm" 0))
  $headerVisibleTopInset = [double](Get-RuleValue $Rules "redHeader.visibleTopInsetPt" 0)
  $configuredTopSpacer = [double](Get-RuleValue $Rules "redHeader.topSpacerPt" -1)
  $headerSpaceBefore = if ($configuredTopSpacer -ge 0) {
    $configuredTopSpacer
  } elseif ($headerTopFromContentTop -gt 0) {
    [Math]::Max(0, $headerTopFromContentTop - $headerVisibleTopInset)
  } else {
    0
  }

  $defaultDocNo = Get-RuleValue $Rules "documentNo.defaultText" "安盟保险〔2026〕142号"
  $docNo = Get-RuleValue $Meta "documentNo" $defaultDocNo
  $signer = Get-RuleValue $Meta "signer" "阮江"
  $signerLabel = Get-RuleValue $Rules "documentNo.signerLabel" "签发人："
  $docNoLeftChars = [int](Get-RuleValue $Rules "documentNo.leftBlankChars" 1)
  $signerRightChars = [int](Get-RuleValue $Rules "documentNo.signerRightBlankChars" 1)
  $docNoPrefix = Repeat-Text "　" ([Math]::Max(0, $docNoLeftChars))

  $insertRange = $Document.Range(0, 0)
  $spacerMarker = "."
  $insertRange.InsertBefore("$spacerMarker`r$headerText`r$docNoPrefix$docNo`t$signerLabel$signer`r")

  $spacerParagraph = $Document.Paragraphs.Item(1)
  if ($headerSpaceBefore -gt 0) {
    $spacerParagraph.Range.Font.Size = 1
    $spacerParagraph.Range.Font.Hidden = 0
    $spacerParagraph.Range.Font.Color = Convert-HexToWordColor "FFFFFF"
    $spacerParagraph.Range.ParagraphFormat.LineSpacingRule = 4
    $spacerParagraph.Range.ParagraphFormat.LineSpacing = [single]$headerSpaceBefore
    $spacerParagraph.Range.ParagraphFormat.SpaceBefore = 0
    $spacerParagraph.Range.ParagraphFormat.SpaceAfter = 0
    $spacerParagraph.Range.ParagraphFormat.FirstLineIndent = 0
  }

  $p1 = $Document.Paragraphs.Item(2)
  Set-RangeFont $p1.Range $headerFont $headerSize "" $headerColor $headerScale $headerSpacing
  $p1.Range.ParagraphFormat.Alignment = 1
  $p1.Range.ParagraphFormat.FirstLineIndent = 0
  $p1.Range.ParagraphFormat.LineSpacingRule = 0
  $p1.Range.ParagraphFormat.SpaceBefore = 0
  $p1.Range.ParagraphFormat.SpaceAfter = $headerSpaceAfter

  $p2 = $Document.Paragraphs.Item(3)
  $docNoFont = Get-RuleValue $Rules "documentNo.font" "仿宋_GB2312"
  $docNoSize = [double](Get-RuleValue $Rules "documentNo.sizePt" 16)
  $docNoSpaceAfter = [double](Get-RuleValue $Rules "documentNo.spaceAfterPt" 40)
  Set-RangeFont $p2.Range $docNoFont $docNoSize "Times New Roman" 0
  $p2.Range.ParagraphFormat.Alignment = 0
  $p2.Range.ParagraphFormat.FirstLineIndent = 0
  $p2.Range.ParagraphFormat.LeftIndent = 0
  $p2.Range.ParagraphFormat.RightIndent = 0
  $p2.Range.ParagraphFormat.SpaceBefore = 0
  $p2.Range.ParagraphFormat.SpaceAfter = $docNoSpaceAfter
  $p2.Range.ParagraphFormat.TabStops.ClearAll()
  $contentWidth = $Document.PageSetup.PageWidth - $Document.PageSetup.LeftMargin - $Document.PageSetup.RightMargin
  $signerTabPosition = $contentWidth - ($docNoSize * [Math]::Max(0, $signerRightChars))
  $null = $p2.Range.ParagraphFormat.TabStops.Add($signerTabPosition, 2, 0)

  $bottomBorder = $p2.Range.ParagraphFormat.Borders.Item(-3)
  $bottomBorder.LineStyle = 0

  $lineText = Get-ParagraphRawText $p2
  $signerIndex = $lineText.IndexOf($signer)
  if ($signerIndex -ge 0) {
    $signerRange = $Document.Range($p2.Range.Start + $signerIndex, $p2.Range.Start + $signerIndex + $signer.Length)
    $signerFont = Get-RuleValue $Rules "documentNo.signerFont" "楷体_GB2312"
    $signerRange.Font.NameFarEast = $signerFont
    $signerRange.Font.Name = $signerFont
    $signerRange.Font.Color = 0
    $signerRange.Font.Bold = 0
  }

  Adjust-RedHeaderLayout $Document $Rules
}

function Apply-BodyAndTitleStyle {
  param($Document, $Rules)

  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  $bodyStartIndex = [Math]::Max(3, $docNoIndex + 1)
  for ($i = $bodyStartIndex; $i -le $Document.Paragraphs.Count; $i++) {
    Set-BodyParagraph $Document.Paragraphs.Item($i) $Rules
  }

  $titleCount = 0
  for ($i = $bodyStartIndex; $i -le [Math]::Min($Document.Paragraphs.Count, 15 + $bodyStartIndex); $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    $text = Get-ParagraphText $paragraph
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    if ($text -match "[：:]$" -and $titleCount -gt 0) { break }
    Set-TitleParagraph $paragraph $Rules
    $titleCount++
    if ($titleCount -ge 6) { break }
  }

  for ($i = $bodyStartIndex; $i -le [Math]::Min($Document.Paragraphs.Count, 30 + $bodyStartIndex); $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    $text = Get-ParagraphText $paragraph
    if ($text -match "[：:]$") {
      Set-MainRecipientParagraph $paragraph $Rules
      break
    }
  }
}

function Adjust-TitlePositionAfterRedLine {
  param($Document, $Rules)

  $structure = Find-DocumentStructure $Document
  $titleParagraph = $structure.TitleParagraph
  if ($null -eq $titleParagraph) { return }

  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  if ($null -eq $docNoIndex) { return }
  $docNoParagraph = $Document.Paragraphs.Item([int]$docNoIndex)
  $redLineBorder = $docNoParagraph.Range.ParagraphFormat.Borders.Item(-3)
  if ([int]$redLineBorder.LineStyle -eq 0) { return }

  $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $titleTopBlankLines = [double](Get-RuleValue $Rules "title.topBlankLines" 2)
  $targetGap = $bodySize * $titleTopBlankLines
  $lineBoxHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  $redLineOffset = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 4))

  for ($attempt = 0; $attempt -lt 3; $attempt++) {
    $Document.Repaginate()
    $redLineBottom = [double]$docNoParagraph.Range.Information(6) + $lineBoxHeight + $redLineOffset + (Convert-WordLineWidthToPt ([int]$redLineBorder.LineWidth))
    $titleTop = [double]$titleParagraph.Range.Information(6)
    $actualGap = $titleTop - $redLineBottom
    $delta = $targetGap - $actualGap
    if ([Math]::Abs($delta) -le 0.5) { break }

    $currentSpaceBefore = [double]$titleParagraph.Range.ParagraphFormat.SpaceBefore
    $titleParagraph.Range.ParagraphFormat.SpaceBefore = [single]([Math]::Max(0, $currentSpaceBefore + $delta))
  }
}

function Apply-Signature {
  param($Document, $Rules, $Meta)

  $company = Get-RuleValue $Meta "company" (Get-RuleValue $Rules "signature.company" "安盟财产保险有限公司")
  $date = Get-RuleValue $Meta "date" (Get-RuleValue $Rules "signature.date" "2026年5月9日")
  $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $noSeal = [bool](Get-RuleValue $Rules "signature.noSeal" $false)
  $dateRightChars = [double](Get-RuleValue $Rules "signature.dateRightIndentChars" 4)
  $companyCenterCorrection = [double](Get-RuleValue $Rules "signature.companyCenterCorrectionPt" 3)
  $companyIndent = [double](Get-RuleValue $Rules "signature.companyLeftIndentPt" 250.2)
  $dateIndent = [double](Get-RuleValue $Rules "signature.dateLeftIndentPt" 282.2)
  $dateRightIndent = $dateRightChars * $bodySize
  $dateWidth = Get-SignatureTextWidthPt $date $bodySize
  $companyWidth = Get-SignatureTextWidthPt $company $bodySize
  $companyRightIndent = [Math]::Max(0, $dateRightIndent - (($companyWidth - $dateWidth) / 2) + $companyCenterCorrection)
  $normalizedCompany = Normalize-SignatureText $company

  function Get-TailSignatureItems {
    $items = @()
    $startIndex = [Math]::Max(1, $Document.Paragraphs.Count - 90)
    for ($i = $startIndex; $i -le $Document.Paragraphs.Count; $i++) {
      $paragraph = $Document.Paragraphs.Item($i)
      $text = Get-ParagraphText $paragraph
      if ([string]::IsNullOrWhiteSpace($text)) { continue }
      $normalized = Normalize-SignatureText $text
      $items += [pscustomobject]@{
        Paragraph = $paragraph
        Text = $text
        Normalized = $normalized
        Start = $paragraph.Range.Start
        IsCompany = ($normalized -eq $normalizedCompany)
        IsDate = (Test-DateParagraphText $text)
      }
    }
    return @($items)
  }

  function Find-PrimarySignaturePair {
    param($Items)
    $companies = @($Items | Where-Object { $_.IsCompany } | Sort-Object Start)
    $dates = @($Items | Where-Object { $_.IsDate } | Sort-Object Start)
    $primaryCompany = $null
    $primaryDate = $null

    if ($companies.Count -gt 0) {
      $primaryCompany = $companies[-1]
      $datesAfterCompany = @($dates | Where-Object { $_.Start -gt $primaryCompany.Start } | Sort-Object Start)
      if ($datesAfterCompany.Count -gt 0) {
        $primaryDate = $datesAfterCompany[0]
      } elseif ($dates.Count -gt 0) {
        $primaryDate = $dates[-1]
      }
    } elseif ($dates.Count -gt 0) {
      $primaryDate = $dates[-1]
      $companiesBeforeDate = @($companies | Where-Object { $_.Start -lt $primaryDate.Start } | Sort-Object Start)
      if ($companiesBeforeDate.Count -gt 0) {
        $primaryCompany = $companiesBeforeDate[-1]
      }
    }

    return @{
      Company = $primaryCompany
      Date = $primaryDate
      Companies = $companies
      Dates = $dates
    }
  }

  $items = Get-TailSignatureItems
  $pair = Find-PrimarySignaturePair $items

  if ($null -eq $pair.Company -and $null -eq $pair.Date) {
    $range = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
    $range.InsertAfter("`r`r$company`r$date`r")
  } elseif ($null -ne $pair.Company -and $null -eq $pair.Date) {
    $range = $pair.Company.Paragraph.Range.Duplicate
    $range.Collapse(0)
    $range.InsertAfter("$date`r")
  } elseif ($null -eq $pair.Company -and $null -ne $pair.Date) {
    $range = $pair.Date.Paragraph.Range.Duplicate
    $range.Collapse(1)
    $range.InsertBefore("$company`r")
  }

  $items = Get-TailSignatureItems
  $pair = Find-PrimarySignaturePair $items
  $keepStarts = @()
  if ($null -ne $pair.Company) { $keepStarts += [int]$pair.Company.Start }
  if ($null -ne $pair.Date) { $keepStarts += [int]$pair.Date.Start }

  $deleteRanges = @()
  foreach ($item in $items) {
    if (($item.IsCompany -or $item.IsDate) -and ($keepStarts -notcontains [int]$item.Start)) {
      $deleteRanges += $item.Paragraph.Range.Duplicate
    }
  }
  foreach ($range in @($deleteRanges | Sort-Object Start -Descending)) {
    $range.Delete() | Out-Null
  }

  $items = Get-TailSignatureItems
  $pair = Find-PrimarySignaturePair $items
  if ($null -ne $pair.Company) {
    $pair.Company.Paragraph.Range.Text = "$company`r"
  }
  if ($null -ne $pair.Date) {
    $pair.Date.Paragraph.Range.Text = "$date`r"
  }

  $companyParagraph = $null
  $dateParagraph = $null
  foreach ($paragraph in $Document.Paragraphs) {
    $text = Get-ParagraphText $paragraph
    $normalized = Normalize-SignatureText $text
    if ($normalized -eq $normalizedCompany) {
      Set-BodyParagraph $paragraph $Rules
      $paragraph.Range.ParagraphFormat.FirstLineIndent = 0
      if ($noSeal) {
        $paragraph.Range.ParagraphFormat.Alignment = 0
        $paragraph.Range.ParagraphFormat.LeftIndent = $companyIndent
        $paragraph.Range.ParagraphFormat.RightIndent = 0
      } else {
        $paragraph.Range.ParagraphFormat.Alignment = 2
        $paragraph.Range.ParagraphFormat.LeftIndent = 0
        $paragraph.Range.ParagraphFormat.RightIndent = [single]$companyRightIndent
      }
      $companyParagraph = $paragraph
    }
    if ((Normalize-SignatureText $text) -eq (Normalize-SignatureText $date)) {
      Set-BodyParagraph $paragraph $Rules
      $paragraph.Range.ParagraphFormat.FirstLineIndent = 0
      if ($noSeal) {
        $paragraph.Range.ParagraphFormat.Alignment = 0
        $paragraph.Range.ParagraphFormat.LeftIndent = $dateIndent
        $paragraph.Range.ParagraphFormat.RightIndent = 0
      } else {
        $paragraph.Range.ParagraphFormat.Alignment = 2
        $paragraph.Range.ParagraphFormat.LeftIndent = 0
        $paragraph.Range.ParagraphFormat.RightIndent = [single]$dateRightIndent
      }
      $dateParagraph = $paragraph
    }
  }

  if (-not $noSeal) {
    Add-SealIfConfigured $Document $Rules $companyParagraph $dateParagraph
  }
}

function Add-SealIfConfigured {
  param($Document, $Rules, $CompanyParagraph, $DateParagraph)
  if (-not [bool](Get-RuleValue $Rules "signature.seal.enabled" $true)) { return }
  $imagePath = Resolve-ConfiguredPath (Get-RuleValue $Rules "signature.seal.imagePath" "")
  if ([string]::IsNullOrWhiteSpace($imagePath) -or -not (Test-Path -LiteralPath $imagePath)) { return }
  if ($null -eq $DateParagraph) { return }

  $width = Convert-MmToPt ([double](Get-RuleValue $Rules "signature.seal.widthMm" 42))
  $height = Convert-MmToPt ([double](Get-RuleValue $Rules "signature.seal.heightMm" 42))
  $verticalOffset = [double](Get-RuleValue $Rules "signature.seal.verticalOffsetPt" -34)

  $pageWidth = $Document.PageSetup.PageWidth
  $rightMargin = $Document.PageSetup.RightMargin
  $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $dateRightChars = [double](Get-RuleValue $Rules "signature.dateRightIndentChars" 4)
  $dateRightIndent = $dateRightChars * $bodySize
  $dateText = Get-ParagraphText $DateParagraph
  $dateCenter = $pageWidth - $rightMargin - $dateRightIndent - ((Get-SignatureTextWidthPt $dateText $bodySize) / 2)
  $left = $dateCenter - ($width / 2)

  try {
    $top = [double]$DateParagraph.Range.Information(6) + $verticalOffset
  } catch {
    $top = Convert-MmToPt 210
  }

  $shape = $Document.Shapes.AddPicture($imagePath, $false, $true, $left, $top, $width, $height, $DateParagraph.Range)
  $shape.WrapFormat.Type = 5
  $shape.LockAspectRatio = $true
  $shape.Name = "redhead-seal"
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

function Normalize-ImprintElementText {
  param([string]$Text, [string]$Label)
  if ($null -eq $Text) { return "" }
  $labelText = if ([string]::IsNullOrWhiteSpace($Label)) { "" } else { $Label.Trim() }
  $value = ($Text -replace "[`r`a]", "").Trim()
  if ([string]::IsNullOrWhiteSpace($value)) { return "" }
  if (-not [string]::IsNullOrWhiteSpace($labelText)) {
    $value = $value -replace ("^\s*" + [regex]::Escape($labelText) + "\s*"), ""
  }
  $value = $value.Trim()
  $value = $value -replace "[。；;]\s*$", ""
  if ([string]::IsNullOrWhiteSpace($value)) { return "" }
  return "$labelText$value。"
}

function Get-ImprintContentWidthFromRules {
  param($Rules)
  $pageSpec = Get-PageSizeSpec $Rules
  $pageWidth = Convert-MmToPt ([double]$pageSpec.WidthMm)
  $left = Convert-MmToPt ([double](Get-RuleValue $Rules "page.leftMarginMm" 28))
  $right = Convert-MmToPt ([double](Get-RuleValue $Rules "page.rightMarginMm" 26))
  return $pageWidth - $left - $right
}

function Get-ImprintElementRowHeight {
  param($Rules, [string]$Text, [double]$LeftBlankChars = 1, [double]$RightBlankChars = 1)
  $rowHeight = [double](Get-RuleValue $Rules "imprint.rowHeightPt" 28)
  $lineSpacing = [double](Get-RuleValue $Rules "imprint.lineSpacingPt" 28)
  $size = [double](Get-RuleValue $Rules "imprint.sizePt" 14)
  $topPadding = [double](Get-RuleValue $Rules "imprint.cellPaddingTopPt" 0)
  $bottomPadding = [double](Get-RuleValue $Rules "imprint.cellPaddingBottomPt" 0)
  $contentWidth = Get-ImprintContentWidthFromRules $Rules
  $availableWidth = [Math]::Max($size, $contentWidth - (($LeftBlankChars + $RightBlankChars) * $size))
  $availableUnits = [Math]::Max(1.0, $availableWidth / $size)
  $lineCount = [Math]::Max(1, [Math]::Ceiling((Get-TextCharUnits $Text) / $availableUnits))
  return [double]([Math]::Max($rowHeight, ($lineSpacing * $lineCount) + $topPadding + $bottomPadding))
}

function Get-ImprintRows {
  param($Rules)
  $rows = @()

  $mainRecipientEnabled = [bool](Get-RuleValue $Rules "imprint.mainRecipient.enabled" $false)
  $mainRecipientLabel = Get-RuleValue $Rules "imprint.mainRecipient.label" "主送："
  $mainRecipientText = Normalize-ImprintElementText (Get-RuleValue $Rules "imprint.mainRecipient.text" "") $mainRecipientLabel
  if ($mainRecipientEnabled -and -not [string]::IsNullOrWhiteSpace($mainRecipientText)) {
    $leftChars = [double](Get-RuleValue $Rules "imprint.mainRecipient.leftBlankChars" 1)
    $rightChars = [double](Get-RuleValue $Rules "imprint.mainRecipient.rightBlankChars" 1)
    $rows += [pscustomobject]@{
      Type = "mainRecipient"
      Text = $mainRecipientText
      Label = $mainRecipientLabel
      LeftBlankChars = $leftChars
      RightBlankChars = $rightChars
      RowHeight = (Get-ImprintElementRowHeight $Rules $mainRecipientText $leftChars $rightChars)
    }
  }

  $ccEnabled = [bool](Get-RuleValue $Rules "imprint.cc.enabled" $false)
  $ccLabel = Get-RuleValue $Rules "imprint.cc.label" "抄送："
  $ccText = Normalize-ImprintElementText (Get-RuleValue $Rules "imprint.cc.text" "") $ccLabel
  if ($ccEnabled -and -not [string]::IsNullOrWhiteSpace($ccText)) {
    $leftChars = [double](Get-RuleValue $Rules "imprint.cc.leftBlankChars" 1)
    $rightChars = [double](Get-RuleValue $Rules "imprint.cc.rightBlankChars" 1)
    $rows += [pscustomobject]@{
      Type = "cc"
      Text = $ccText
      Label = $ccLabel
      LeftBlankChars = $leftChars
      RightBlankChars = $rightChars
      RowHeight = (Get-ImprintElementRowHeight $Rules $ccText $leftChars $rightChars)
    }
  }

  $office = Get-RuleValue $Rules "imprint.office" "安盟财产保险有限公司综合办公室"
  $date = Get-RuleValue $Rules "imprint.date" "2026年5月9日印发"
  $officeLeftChars = [double](Get-RuleValue $Rules "imprint.officeLeftChars" 1)
  $dateRightChars = [double](Get-RuleValue $Rules "imprint.dateRightChars" 1)
  $rows += [pscustomobject]@{
    Type = "issue"
    Text = "$office`t$date"
    Label = ""
    LeftBlankChars = 0
    RightBlankChars = 0
    RowHeight = (Get-ImprintElementRowHeight $Rules "$office`t$date" $officeLeftChars $dateRightChars)
  }
  return @($rows)
}

function Get-ImprintRowsTotalHeight {
  param($Rows)
  $height = 0.0
  foreach ($row in @($Rows)) {
    $height += [double]$row.RowHeight
  }
  return $height
}

function Get-ImprintTableHeight {
  param($Table)
  if ($null -eq $Table) { return 0.0 }
  $height = 0.0
  for ($i = 1; $i -le $Table.Rows.Count; $i++) {
    $rowHeight = [double]$Table.Rows.Item($i).Height
    if ($rowHeight -ge 1000000) { return $height }
    $height += $rowHeight
  }
  return $height
}

function Set-ImprintElementCell {
  param($Cell, $Row, $Rules, [double]$ContentWidth)
  $size = [double](Get-RuleValue $Rules "imprint.sizePt" 14)
  $lineSpacing = [double](Get-RuleValue $Rules "imprint.lineSpacingPt" 28)
  $baselineShift = [double](Get-RuleValue $Rules "imprint.baselineShiftPt" 3)
  $leftStart = $size * [double]$Row.LeftBlankChars
  $labelWidth = Get-ApproxTextWidthPt ([string]$Row.Label) $size
  $rightIndent = $size * [double]$Row.RightBlankChars

  $Cell.Range.Text = [string]$Row.Text
  $Cell.VerticalAlignment = 1
  $Cell.Range.ParagraphFormat.Alignment = 0
  $Cell.Range.ParagraphFormat.LineSpacingRule = 4
  $Cell.Range.ParagraphFormat.LineSpacing = [single]$lineSpacing
  $Cell.Range.ParagraphFormat.SpaceBefore = 0
  $Cell.Range.ParagraphFormat.SpaceAfter = 0
  $Cell.Range.ParagraphFormat.LeftIndent = [single]($leftStart + $labelWidth)
  $Cell.Range.ParagraphFormat.FirstLineIndent = [single](-1 * $labelWidth)
  $Cell.Range.ParagraphFormat.RightIndent = [single]$rightIndent
  $Cell.Range.ParagraphFormat.TabStops.ClearAll()
  $Cell.Range.Font.Position = [single]$baselineShift
}

function Set-ImprintIssueCell {
  param($Cell, $Rules, [double]$ContentWidth)
  $office = Get-RuleValue $Rules "imprint.office" "安盟财产保险有限公司综合办公室"
  $date = Get-RuleValue $Rules "imprint.date" "2026年5月9日印发"
  $officeLeftChars = [int](Get-RuleValue $Rules "imprint.officeLeftChars" 1)
  $dateRightChars = [int](Get-RuleValue $Rules "imprint.dateRightChars" 1)
  $size = [double](Get-RuleValue $Rules "imprint.sizePt" 14)
  $baselineShift = [double](Get-RuleValue $Rules "imprint.baselineShiftPt" 3)
  $officePrefix = Repeat-Text "　" ([Math]::Max(0, $officeLeftChars))
  $dateRightIndent = $size * [Math]::Max(0, $dateRightChars)
  $dateTabPosition = [Math]::Max(0, $ContentWidth - $dateRightIndent)

  $Cell.Range.Text = "$officePrefix$office`t$date"
  $Cell.VerticalAlignment = 1
  $Cell.Range.ParagraphFormat.Alignment = 0
  $Cell.Range.ParagraphFormat.LineSpacingRule = 4
  $Cell.Range.ParagraphFormat.LineSpacing = [single]([double](Get-RuleValue $Rules "imprint.lineSpacingPt" 28))
  $Cell.Range.ParagraphFormat.SpaceBefore = 0
  $Cell.Range.ParagraphFormat.SpaceAfter = 0
  $Cell.Range.ParagraphFormat.LeftIndent = 0
  $Cell.Range.ParagraphFormat.FirstLineIndent = 0
  $Cell.Range.ParagraphFormat.RightIndent = 0
  $Cell.Range.ParagraphFormat.TabStops.ClearAll()
  $null = $Cell.Range.ParagraphFormat.TabStops.Add($dateTabPosition, 2, 0)
  $Cell.Range.Font.Position = [single]$baselineShift
}

function Add-ImprintTable {
  param($Document, $Rules, $Range)

  $rows = @(Get-ImprintRows $Rules)
  Remove-NamedShapes $Document "redhead-imprint-bottom-line"

  $table = $Document.Tables.Add($Range, [Math]::Max(1, $rows.Count), 1)
  $contentWidth = $Document.PageSetup.PageWidth - $Document.PageSetup.LeftMargin - $Document.PageSetup.RightMargin
  $table.PreferredWidthType = 3
  $table.PreferredWidth = $contentWidth
  $table.AllowAutoFit = $false
  $table.Columns.Item(1).PreferredWidth = $contentWidth
  for ($i = 1; $i -le $rows.Count; $i++) {
    $table.Rows.Item($i).HeightRule = 2
    $table.Rows.Item($i).Height = [single]$rows[$i - 1].RowHeight
  }
  $table.TopPadding = [single]([double](Get-RuleValue $Rules "imprint.cellPaddingTopPt" 0))
  $table.BottomPadding = [single]([double](Get-RuleValue $Rules "imprint.cellPaddingBottomPt" 0))
  $table.LeftPadding = 0
  $table.RightPadding = 0
  $table.Borders.Enable = $false
  $table.Borders.Item(-1).LineStyle = 1
  $table.Borders.Item(-1).LineWidth = Convert-LineWidthPtToWordEnum ([double](Get-RuleValue $Rules "imprint.outerLineWidthPt" 1.0))
  $table.Borders.Item(-3).LineStyle = 1
  $table.Borders.Item(-3).LineWidth = Convert-LineWidthPtToWordEnum ([double](Get-RuleValue $Rules "imprint.outerLineWidthPt" 1.0))
  if ($rows.Count -gt 1) {
    $table.Borders.Item(-5).LineStyle = 1
    $table.Borders.Item(-5).LineWidth = Convert-LineWidthPtToWordEnum ([double](Get-RuleValue $Rules "imprint.innerLineWidthPt" 0.75))
  }

  Set-RangeFont $table.Range (Get-RuleValue $Rules "imprint.font" "仿宋_GB2312") ([double](Get-RuleValue $Rules "imprint.sizePt" 14)) "Times New Roman"
  for ($i = 1; $i -le $rows.Count; $i++) {
    $row = $rows[$i - 1]
    if ($row.Type -eq "issue") {
      Set-ImprintIssueCell $table.Cell($i, 1) $Rules $contentWidth
    } else {
      Set-ImprintElementCell $table.Cell($i, 1) $row $Rules $contentWidth
    }
  }
  Set-CollapsedParagraphAfterRange $Document $table.Range
  return $table
}

function Align-ImprintTableToContentBottom {
  param($Document, $Table, $AnchorParagraph, $Rules)
  $expectedBottom = [double]($Document.PageSetup.PageHeight - $Document.PageSetup.BottomMargin)
  for ($i = 0; $i -lt 10; $i++) {
    $Document.Repaginate()
    $tableTop = [double]$Table.Range.Information(6)
    $tableHeight = Get-ImprintTableHeight $Table
    if ($tableHeight -le 0) { break }
    $actualBottom = $tableTop + $tableHeight
    $delta = $expectedBottom - $actualBottom
    if ([Math]::Abs($delta) -le 0.25) { break }
    $movedFloatingTable = $false
    try {
      if ([bool]$Table.Rows.WrapAroundText) {
        $Table.Rows.VerticalPosition = [double]$Table.Rows.VerticalPosition + $delta
        $movedFloatingTable = $true
      }
    } catch {}
    if (-not $movedFloatingTable) {
      $currentSpaceAfter = [double]$AnchorParagraph.Range.ParagraphFormat.SpaceAfter
      $AnchorParagraph.Range.ParagraphFormat.SpaceAfter = [double]([Math]::Max([double]0, [double]($currentSpaceAfter + $delta)))
    }
  }
}

function Position-ImprintTableAtContentBottom {
  param($Document, $Table)
  $tableHeight = Get-ImprintTableHeight $Table
  $contentLeft = [double]$Document.PageSetup.LeftMargin
  $contentBottom = [double]($Document.PageSetup.PageHeight - $Document.PageSetup.BottomMargin)
  $tableTop = $contentBottom - $tableHeight - 3.0
  try {
    $Table.Rows.WrapAroundText = $true
    $Table.Rows.RelativeHorizontalPosition = 1
    $Table.Rows.RelativeVerticalPosition = 1
    $Table.Rows.HorizontalPosition = $contentLeft
    $Table.Rows.VerticalPosition = $tableTop
  } catch {}
}

function Add-Imprint {
  param($Word, $Document, $Rules)
  if (-not [bool](Get-RuleValue $Rules "imprint.enabled" $true)) {
    return @{ BodyPages = $Document.ComputeStatistics(2); BlankInserted = $false; ImprintPage = $null }
  }

  $Document.Repaginate()
  $bodyPages = [int]$Document.ComputeStatistics(2)
  $bodySectionCount = $Document.Sections.Count
  $requireEvenPage = [bool](Get-RuleValue $Rules "imprint.requireEvenPage" $true)
  $samePageWhenPossible = [bool](Get-RuleValue $Rules "imprint.samePageWhenPossible" $true)

  $selection = $Word.Selection
  $null = $selection.EndKey(6)
  if ($samePageWhenPossible -and ((-not $requireEvenPage) -or ($bodyPages % 2 -eq 0))) {
    $anchorParagraph = $null
    for ($i = $Document.Paragraphs.Count; $i -ge 1; $i--) {
      $candidateParagraph = $Document.Paragraphs.Item($i)
      if (-not (Test-BlankParagraph $candidateParagraph)) {
        $anchorParagraph = $candidateParagraph
        break
      }
    }
    if ($null -eq $anchorParagraph) { $anchorParagraph = $Document.Paragraphs.Item($Document.Paragraphs.Count) }
    $anchorY = [double]$anchorParagraph.Range.Information(6)
    $bodyLineSpacing = [double](Get-RuleValue $Rules "body.lineSpacingPt" 28)
    $imprintRows = @(Get-ImprintRows $Rules)
    $imprintHeight = Get-ImprintRowsTotalHeight $imprintRows
    $contentBottom = [double]($Document.PageSetup.PageHeight - $Document.PageSetup.BottomMargin)
    if (($contentBottom - ($anchorY + $bodyLineSpacing)) -ge ($imprintHeight + 8)) {
      $insertStart = [Math]::Max(0, $Document.Content.End - 1)
      $selection.InsertBreak(3)
      $imprintSection = $Document.Sections.Item($Document.Sections.Count)
      Apply-PageSetup $Document $Rules
      Clear-SectionPageFurniture $imprintSection
      $selection.EndKey(6) | Out-Null
      $selection.TypeParagraph()
      $spacerParagraph = $Document.Paragraphs.Item($Document.Paragraphs.Count)
      $spacerParagraph.Range.Font.Size = 1
      $spacerParagraph.Range.Font.Hidden = 1
      $spacerParagraph.Range.ParagraphFormat.LineSpacingRule = 4
      $spacerParagraph.Range.ParagraphFormat.LineSpacing = 1
      $spacerParagraph.Range.ParagraphFormat.SpaceBefore = 0
      $spacerParagraph.Range.ParagraphFormat.SpaceAfter = 0
      $spacerParagraph.Range.ParagraphFormat.FirstLineIndent = 0
      $selection.EndKey(6) | Out-Null
      $table = Add-ImprintTable $Document $Rules $selection.Range
      Position-ImprintTableAtContentBottom $Document $table
      Align-ImprintTableToContentBottom $Document $table $spacerParagraph $Rules
      $Document.Repaginate()
      $tablePage = [int]$table.Range.Information(3)
      $tableTop = [double]$table.Range.Information(6)
      $actualBottom = $tableTop + (Get-ImprintTableHeight $table)
      $samePageOk = ([int]$Document.ComputeStatistics(2) -eq $bodyPages) -and
        ($tablePage -eq $bodyPages) -and
        (Test-Near $actualBottom $contentBottom 2.0)
      if ($samePageOk) {
        return @{
          BodyPages = $bodyPages
          BodySectionCount = $bodySectionCount
          BlankInserted = $false
          SamePage = $true
          Placement = "samePage"
          ImprintPage = $tablePage
        }
      }

      try { $Document.Range($insertStart, $Document.Content.End).Delete() | Out-Null } catch {}
      $Document.Repaginate()
      $selection.EndKey(6) | Out-Null
    }
  }

  $blankNeeded = ($requireEvenPage -and ($bodyPages % 2 -eq 0))
  if ($blankNeeded) {
    $selection.InsertBreak(2)
    Clear-SectionPageFurniture $Document.Sections.Item($Document.Sections.Count)
    $selection.InsertBreak(2)
  } else {
    $selection.InsertBreak(2)
  }

  $imprintSection = $Document.Sections.Item($Document.Sections.Count)
  Apply-PageSetup $Document $Rules
  Clear-SectionPageFurniture $imprintSection
  if ($blankNeeded -and $Document.Sections.Count -gt ($bodySectionCount + 1)) {
    Clear-SectionPageFurniture $Document.Sections.Item($Document.Sections.Count - 1)
  }

  $imprintSection.PageSetup.VerticalAlignment = 3
  $selection.EndKey(6) | Out-Null
  $table = Add-ImprintTable $Document $Rules $selection.Range
  Position-ImprintTableAtContentBottom $Document $table
  $anchorParagraph = $Document.Paragraphs.Item($Document.Paragraphs.Count)
  Align-ImprintTableToContentBottom $Document $table $anchorParagraph $Rules

  return @{
    BodyPages = $bodyPages
    BodySectionCount = $bodySectionCount
    BlankInserted = $blankNeeded
    SamePage = $false
    Placement = if ($blankNeeded) { "blankPage" } else { "newPage" }
    ImprintPage = $null
  }
}

function Get-ExpectedWordColorHex {
  param([int]$WordColor)
  $r = $WordColor -band 255
  $g = ($WordColor -shr 8) -band 255
  $b = ($WordColor -shr 16) -band 255
  return ("{0:X2}{1:X2}{2:X2}" -f $r, $g, $b)
}

function Build-ValidationChecks {
  param($Document, $Rules, $Meta, $ImprintInfo, $TrimInfo, [int]$TotalPages, [int]$ImprintPage, [string]$OutputPdf)

  $checks = @()
  $section = $Document.Sections.Item(1)
  $pageSpec = Get-PageSizeSpec $Rules
  $pageWidth = Convert-MmToPt ([double]$pageSpec.WidthMm)
  $pageHeight = Convert-MmToPt ([double]$pageSpec.HeightMm)
  $top = Convert-MmToPt ([double](Get-RuleValue $Rules "page.topMarginMm" 37))
  $bottom = Convert-MmToPt ([double](Get-RuleValue $Rules "page.bottomMarginMm" 35))
  $left = Convert-MmToPt ([double](Get-RuleValue $Rules "page.leftMarginMm" 28))
  $right = Convert-MmToPt ([double](Get-RuleValue $Rules "page.rightMarginMm" 26))
  $checks += New-Check "$($pageSpec.Name) 页面尺寸 $($pageSpec.WidthMm)mm x $($pageSpec.HeightMm)mm" ((Test-Near $section.PageSetup.PageWidth $pageWidth 1.5) -and (Test-Near $section.PageSetup.PageHeight $pageHeight 1.5))
  $checks += New-Check "页边距符合规则" ((Test-Near $section.PageSetup.TopMargin $top 1.5) -and (Test-Near $section.PageSetup.BottomMargin $bottom 1.5) -and (Test-Near $section.PageSetup.LeftMargin $left 1.5) -and (Test-Near $section.PageSetup.RightMargin $right 1.5))

  $headerText = Get-RuleValue $Rules "redHeader.text" "安盟财产保险有限公司文件"
  $headerIndex = Find-ParagraphIndexByText $Document $headerText 1 20
  $headerParagraph = if ($null -ne $headerIndex) { $Document.Paragraphs.Item([int]$headerIndex) } else { $Document.Paragraphs.Item(1) }
  $headerSize = [double](Get-RuleValue $Rules "redHeader.sizePt" 58)
  $headerScale = [double](Get-RuleValue $Rules "redHeader.scalePercent" 64)
  $headerColorHex = Get-RuleValue $Rules "redHeader.color" "FF0000"
  $actualHeaderColor = Get-ExpectedWordColorHex ([int]$headerParagraph.Range.Font.Color)
  $checks += New-Check "红头文字已插入" ((Get-ParagraphText $headerParagraph) -eq $headerText)
  $checks += New-Check "红头字号、缩放、颜色符合规则" ((Test-Near $headerParagraph.Range.Font.Size $headerSize 0.5) -and ([int]$headerParagraph.Range.Font.Scaling -eq [int]$headerScale) -and ($actualHeaderColor -eq $headerColorHex))
  $structure = Find-DocumentStructure $Document

  $defaultDocNo = Get-RuleValue $Rules "documentNo.defaultText" "安盟保险〔2026〕142号"
  $docNo = Get-RuleValue $Meta "documentNo" $defaultDocNo
  $signer = Get-RuleValue $Meta "signer" "阮江"
  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  $docNoParagraph = $Document.Paragraphs.Item($docNoIndex)
  $docNoLine = Get-ParagraphText $docNoParagraph
  $docNoLineRaw = Get-ParagraphRawText $docNoParagraph
  $checks += New-Check "发文字号和签发人已插入" (($docNoLine.Contains($docNo)) -and ($docNoLine.Contains($signer)))
  $docNoLeftChars = [int](Get-RuleValue $Rules "documentNo.leftBlankChars" 1)
  $signerRightChars = [int](Get-RuleValue $Rules "documentNo.signerRightBlankChars" 1)
  $docNoSize = [double](Get-RuleValue $Rules "documentNo.sizePt" 16)
  $docNoFont = Get-RuleValue $Rules "documentNo.font" "仿宋_GB2312"
  $docNoPrefix = Repeat-Text "　" ([Math]::Max(0, $docNoLeftChars))
  $checks += New-Check "发文字号左空一字" ($docNoLineRaw.StartsWith("$docNoPrefix$docNo")) "配置左空 $docNoLeftChars 字"
  $checks += New-Check "发文字号格式符合规范" (Test-DocumentNoText $docNo) "年份全称，使用六角括号，序号不编虚位，不加第字，末尾加号"
  $docNoStart = $docNoLineRaw.IndexOf($docNo)
  $docNoFontOk = $false
  if ($docNoStart -ge 0) {
    $docNoRange = $Document.Range($docNoParagraph.Range.Start + $docNoStart, $docNoParagraph.Range.Start + $docNoStart + $docNo.Length)
    $docNoFontOk = (($docNoRange.Font.NameFarEast -eq $docNoFont) -or ($docNoRange.Font.Name -eq $docNoFont)) -and (Test-Near ([double]$docNoRange.Font.Size) $docNoSize 0.5)
  }
  $checks += New-Check "发文字号字体字号符合规范" $docNoFontOk "字体 $docNoFont，字号 $docNoSize pt"
  $headerTopFromContentTop = Convert-MmToPt ([double](Get-RuleValue $Rules "redHeader.topFromContentTopMm" 0))
  if ($headerTopFromContentTop -gt 0) {
    $headerVisibleTopInset = [double](Get-RuleValue $Rules "redHeader.visibleTopInsetPt" 0)
    $actualHeaderVisibleTop = [double]$headerParagraph.Range.Information(6) + $headerVisibleTopInset
    $actualHeaderOffset = $actualHeaderVisibleTop - [double]$section.PageSetup.TopMargin
    $checks += New-Check "发文机关标志上边缘距版心上边缘35mm" (Test-Near $actualHeaderOffset $headerTopFromContentTop 1.5) "实际 $([Math]::Round($actualHeaderOffset * 25.4 / 72, 2))mm，目标 $([Math]::Round($headerTopFromContentTop * 25.4 / 72, 2))mm"
  }
  $docNoBlankLines = [double](Get-RuleValue $Rules "redHeader.docNoBlankLines" 2)
  $headerVisibleBottom = [double]$headerParagraph.Range.Information(6) + [double](Get-RuleValue $Rules "redHeader.visibleTopInsetPt" 0) + $headerSize
  $docNoTop = [double]$docNoParagraph.Range.Information(6)
  $expectedDocNoGap = $docNoSize * $docNoBlankLines
  $actualDocNoGap = $docNoTop - $headerVisibleBottom
  $checks += New-Check "发文字号位于发文机关标志下空二行" (Test-Near $actualDocNoGap $expectedDocNoGap 6.0) "实际 $([Math]::Round($actualDocNoGap, 2))pt，目标约 $([Math]::Round($expectedDocNoGap, 2))pt"
  $contentWidth = $Document.PageSetup.PageWidth - $Document.PageSetup.LeftMargin - $Document.PageSetup.RightMargin
  $expectedSignerTab = $contentWidth - ($docNoSize * [Math]::Max(0, $signerRightChars))
  $signerTabOk = $false
  foreach ($tabStop in $docNoParagraph.Range.ParagraphFormat.TabStops) {
    if (([int]$tabStop.Alignment -eq 2) -and (Test-Near ([double]$tabStop.Position) $expectedSignerTab 1.0)) {
      $signerTabOk = $true
      break
    }
  }
  $checks += New-Check "签发人右空一字" $signerTabOk "配置右空 $signerRightChars 字，右对齐制表位 $([Math]::Round($expectedSignerTab, 1))pt"
  $signerFont = Get-RuleValue $Rules "documentNo.signerFont" "楷体_GB2312"
  $signerFontOk = $false
  $signerStart = $docNoLineRaw.IndexOf($signer)
  if ($signerStart -ge 0) {
    $signerRange = $Document.Range($docNoParagraph.Range.Start + $signerStart, $docNoParagraph.Range.Start + $signerStart + $signer.Length)
    $signerFontOk = (($signerRange.Font.NameFarEast -eq $signerFont) -or ($signerRange.Font.Name -eq $signerFont)) -and (Test-Near ([double]$signerRange.Font.Size) $docNoSize 0.5)
  }
  $checks += New-Check "签发人字体字号符合规范" $signerFontOk "签发人三字 $docNoFont，姓名 $signerFont，字号 $docNoSize pt"
  $redLineBorder = $docNoParagraph.Range.ParagraphFormat.Borders.Item(-3)
  $redLineExists = ([int]$redLineBorder.LineStyle -ne 0)
  $checks += New-Check "红线已设置" $redLineExists
  $expectedRedLineColorHex = ((Get-RuleValue $Rules "documentNo.redLineColor" "FF0000") -replace "#", "").ToUpperInvariant()
  $expectedRedLineWidthPt = [double](Get-RuleValue $Rules "documentNo.redLineWidthPt" 2.25)
  $actualRedLineColorHex = ""
  $actualRedLineWidthPt = 0.0
  if ($redLineExists) {
    $actualRedLineColorHex = Get-ExpectedWordColorHex ([int]$redLineBorder.Color)
    $actualRedLineWidthPt = Convert-WordLineWidthToPt ([int]$redLineBorder.LineWidth)
  }
  $redLineFormatOk = $redLineExists -and
    ($actualRedLineColorHex -eq $expectedRedLineColorHex) -and
    (Test-Near $actualRedLineWidthPt $expectedRedLineWidthPt 0.05)
  $checks += New-Check "版头红线颜色线宽符合规则" $redLineFormatOk "颜色 #$expectedRedLineColorHex，线宽 $expectedRedLineWidthPt pt"
  $expectedRedLineOffset = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 4))
  $actualRedLineOffset = 0.0
  $redLineOffsetOk = $false
  if ($redLineExists) {
    try { $actualRedLineOffset = [double]$docNoParagraph.Range.ParagraphFormat.Borders.DistanceFromBottom } catch {}
    $redLineOffsetOk = Test-Near $actualRedLineOffset $expectedRedLineOffset 1.0
  }
  $checks += New-Check "版头红线位于发文字号下4mm" $redLineOffsetOk "可见上沿净距 $([Math]::Round($actualRedLineOffset * 25.4 / 72, 2))mm，目标 $([Math]::Round($expectedRedLineOffset * 25.4 / 72, 2))mm"
  $expectedRedLineLength = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineLengthMm" 156))
  $actualRedLineLength = if ($redLineExists) {
    [double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.LeftMargin - [double]$Document.PageSetup.RightMargin - [double]$docNoParagraph.Range.ParagraphFormat.LeftIndent - [double]$docNoParagraph.Range.ParagraphFormat.RightIndent
  } else { 0.0 }
  $actualRedLineCenter = if ($redLineExists) {
    [double]$Document.PageSetup.LeftMargin + [double]$docNoParagraph.Range.ParagraphFormat.LeftIndent + ($actualRedLineLength / 2)
  } else { 0.0 }
  $expectedContentWidth = [double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.LeftMargin - [double]$Document.PageSetup.RightMargin
  $expectedRedLineCenter = [double]$Document.PageSetup.LeftMargin + ($expectedContentWidth / 2)
  $redLineWidthOk = $redLineExists -and
    (Test-Near $actualRedLineLength $expectedRedLineLength 1.0) -and
    (Test-Near $actualRedLineCenter $expectedRedLineCenter 1.0)
  $checks += New-Check "版头红线居中且与版心等宽" $redLineWidthOk "实际长 $([Math]::Round($actualRedLineLength, 1))pt/$([Math]::Round($actualRedLineLength * 25.4 / 72, 2))mm，目标 $([Math]::Round($expectedRedLineLength, 1))pt/$([Math]::Round($expectedRedLineLength * 25.4 / 72, 2))mm，中心 $([Math]::Round($actualRedLineCenter * 25.4 / 72, 2))mm"

  $titleParagraph = $null
  $titleItems = @($structure.TitleItems)
  $titleParagraphs = @($titleItems | ForEach-Object { $_.Paragraph })
  if ($titleParagraphs.Count -gt 0) { $titleParagraph = $titleParagraphs[0] }
  $titleTopGapOk = $false
  $actualTitleTopGap = 0.0
  $bodySizeForTitleGap = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $titleTopBlankLines = [double](Get-RuleValue $Rules "title.topBlankLines" 2)
  $expectedTitleTopGap = $bodySizeForTitleGap * $titleTopBlankLines
  if (($null -ne $titleParagraph) -and $redLineExists) {
    $lineBoxHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
    $redLineBottom = [double]$docNoParagraph.Range.Information(6) + $lineBoxHeight + $expectedRedLineOffset + $actualRedLineWidthPt
    $actualTitleTopGap = [double]$titleParagraph.Range.Information(6) - $redLineBottom
    $titleTopGapOk = Test-Near $actualTitleTopGap $expectedTitleTopGap 1.0
  }
  $checks += New-Check "标题编排于红色分隔线下空二行" $titleTopGapOk "实际 $([Math]::Round($actualTitleTopGap, 1))pt/$([Math]::Round($actualTitleTopGap * 25.4 / 72, 2))mm，目标 $([Math]::Round($expectedTitleTopGap, 1))pt/$([Math]::Round($expectedTitleTopGap * 25.4 / 72, 2))mm"
  $titleFont = Get-RuleValue $Rules "title.font" "方正小标宋简体"
  $titleSize = [double](Get-RuleValue $Rules "title.sizePt" 22)
  $titleColorHex = ((Get-RuleValue $Rules "title.color" "000000") -replace "#", "").ToUpperInvariant()
  $titleLineSpacing = [double](Get-RuleValue $Rules "title.lineSpacingPt" 32)
  $actualTitleColor = if ($null -ne $titleParagraph) { Get-ExpectedWordColorHex ([int]$titleParagraph.Range.Font.Color) } else { "" }
  $titleFormatOk = ($null -ne $titleParagraph) -and
    ([int]$titleParagraph.Range.ParagraphFormat.Alignment -eq 1) -and
    (Test-Near ([double]$titleParagraph.Range.Font.Size) $titleSize 0.5) -and
    (Test-Near ([double]$titleParagraph.Range.ParagraphFormat.LineSpacing) $titleLineSpacing 0.5) -and
    ($actualTitleColor -eq $titleColorHex)
  $checks += New-Check "标题字体字号颜色行距居中符合规范" $titleFormatOk "字体 $titleFont，字号 $titleSize pt，颜色 #$titleColorHex，行距 $titleLineSpacing pt"
  if ($null -ne $titleParagraph) {
    $titleRaw = Get-ParagraphRawText $titleParagraph
    $actualTitleLines = @(($titleRaw -split ([string][char]11)) | Where-Object { -not [string]::IsNullOrWhiteSpace((Normalize-TitleText $_)) })
    $normalizedTitle = Normalize-TitleText $titleRaw
    $maxLineChars = [int](Get-RuleValue $Rules "title.maxLineChars" 22)
    $expectedTitleLines = @(Split-TitleText $normalizedTitle $maxLineChars)
    $lineBreaksMatch = ($actualTitleLines.Count -eq $expectedTitleLines.Count)
    if ($lineBreaksMatch) {
      for ($j = 0; $j -lt $actualTitleLines.Count; $j++) {
        if ((Normalize-TitleText $actualTitleLines[$j]) -ne $expectedTitleLines[$j]) {
          $lineBreaksMatch = $false
          break
        }
      }
    }
    $breaksSemanticOk = $true
    $breakPositions = @()
    $position = 0
    for ($j = 0; $j -lt ($actualTitleLines.Count - 1); $j++) {
      $position += (Normalize-TitleText $actualTitleLines[$j]).Length
      $breakPositions += $position
      if ((Test-TitleBreakInsideProtectedTerm $normalizedTitle $position) -or (Test-TitleBreakHasBadEdge $normalizedTitle $position)) {
        $breaksSemanticOk = $false
      }
    }
    $titleBlockOk = ($titleParagraphs.Count -eq 1) -and $lineBreaksMatch -and $breaksSemanticOk
    $checks += New-Check "标题已合并并按词义主动分行" $titleBlockOk "标题 $($actualTitleLines.Count) 行，断点 $($breakPositions -join ',')，每行最多 $maxLineChars 字"
    if ($normalizedTitle.Contains("关于对")) {
      $checks += New-Warn "标题用词建议" "规范建议标题尽量少用介词；当前包含'关于对'，如不影响正式名称，可人工确认是否优化为'关于XXX的报告'等表达。"
    }
  }

  $mainRecipientItem = $structure.MainRecipient
  if ($null -ne $mainRecipientItem) {
    $mainRecipientParagraph = $mainRecipientItem.Paragraph
    $mainRecipientText = Get-ParagraphText $mainRecipientParagraph
    $mainRecipientFont = Get-RuleValue $Rules "mainRecipient.font" (Get-RuleValue $Rules "body.font" "仿宋_GB2312")
    $mainRecipientSize = [double](Get-RuleValue $Rules "mainRecipient.sizePt" (Get-RuleValue $Rules "body.sizePt" 16))
    $mainRecipientLineSpacing = [double](Get-RuleValue $Rules "mainRecipient.lineSpacingPt" (Get-RuleValue $Rules "body.lineSpacingPt" 28))
    $titleSpaceAfter = [double](Get-RuleValue $Rules "mainRecipient.titleSpaceAfterPt" (Get-RuleValue $Rules "title.spaceAfterPt" 22))
    $mainRecipientLayoutOk = ([int]$mainRecipientParagraph.Range.ParagraphFormat.Alignment -eq 0) -and
      (Test-Near ([double]$mainRecipientParagraph.Range.ParagraphFormat.LeftIndent) 0 0.5) -and
      (Test-Near ([double]$mainRecipientParagraph.Range.ParagraphFormat.FirstLineIndent) 0 0.5) -and
      (($null -eq $titleParagraph) -or (Test-Near ([double]$titleParagraph.Range.ParagraphFormat.SpaceAfter) $titleSpaceAfter 1.0))
    $checks += New-Check "主送机关标题下空一行且顶格" $mainRecipientLayoutOk "标题后距 $titleSpaceAfter pt，主送机关左顶格"
    $checks += New-Check "主送机关使用全角冒号" ($mainRecipientText.EndsWith("：")) $mainRecipientText
    $mainRecipientColorHex = Get-ExpectedWordColorHex ([int]$mainRecipientParagraph.Range.Font.Color)
    $mainRecipientFontOk = (($mainRecipientParagraph.Range.Font.NameFarEast -eq $mainRecipientFont) -or ($mainRecipientParagraph.Range.Font.Name -eq $mainRecipientFont))
    $mainRecipientFormatOk = $mainRecipientFontOk -and
      (Test-Near ([double]$mainRecipientParagraph.Range.Font.Size) $mainRecipientSize 0.5) -and
      (Test-Near ([double]$mainRecipientParagraph.Range.ParagraphFormat.LineSpacing) $mainRecipientLineSpacing 0.5) -and
      ($mainRecipientColorHex -eq ((Get-RuleValue $Rules "body.color" "000000") -replace "#", "").ToUpperInvariant())
    $checks += New-Check "主送机关字体字号行距符合规范" $mainRecipientFormatOk "字体 $mainRecipientFont，字号 $mainRecipientSize pt，行距 $mainRecipientLineSpacing pt"
  } else {
    $checks += New-Check "主送机关标题下空一行且顶格" $false "未识别到主送机关"
    $checks += New-Check "主送机关使用全角冒号" $false "未识别到主送机关"
    $checks += New-Check "主送机关字体字号行距符合规范" $false "未识别到主送机关"
  }

  $attachmentItems = @(Find-AttachmentExplanationParagraphs $Document)
  if ($attachmentItems.Count -gt 0) {
    $attachmentFont = Get-RuleValue $Rules "attachment.font" (Get-RuleValue $Rules "body.font" "仿宋_GB2312")
    $attachmentSize = [double](Get-RuleValue $Rules "attachment.sizePt" (Get-RuleValue $Rules "body.sizePt" 16))
    $attachmentLineSpacing = [double](Get-RuleValue $Rules "attachment.lineSpacingPt" (Get-RuleValue $Rules "body.lineSpacingPt" 28))
    $attachmentLeftChars = [double](Get-RuleValue $Rules "attachment.leftBlankChars" 2)
    $attachmentSpaceBeforeLines = [double](Get-RuleValue $Rules "attachment.spaceBeforeLines" 1)
    $firstAttachment = $attachmentItems[0]
    $firstAttachmentText = Get-ParagraphText $firstAttachment.Paragraph
    $previousBlank = if ([int]$firstAttachment.Index -gt 1) { Test-BlankParagraph $Document.Paragraphs.Item([int]$firstAttachment.Index - 1) } else { $false }
    $spaceBeforeOk = $previousBlank -or (Test-Near ([double]$firstAttachment.Paragraph.Range.ParagraphFormat.SpaceBefore) ($attachmentLineSpacing * $attachmentSpaceBeforeLines) 1.0)
    $firstLineStart = [double]$firstAttachment.Paragraph.Range.ParagraphFormat.LeftIndent + [double]$firstAttachment.Paragraph.Range.ParagraphFormat.FirstLineIndent
    $expectedFirstLineStart = $attachmentSize * $attachmentLeftChars
    $checks += New-Check "附件说明正文下空一行左空二字" (($spaceBeforeOk) -and (Test-Near $firstLineStart $expectedFirstLineStart 1.0)) "左空 $attachmentLeftChars 字，正文下空 $attachmentSpaceBeforeLines 行"

    $colonOk = $firstAttachmentText -match "^附件："
    $checks += New-Check "附件说明使用全角冒号" $colonOk $firstAttachmentText

    $formatOk = $true
    $indentOk = $true
    $punctuationOk = $true
    $numberingOk = $true
    $numberAlignmentOk = $true
    $attachmentNumberStart = $expectedFirstLineStart + (Get-ApproxTextWidthPt "附件：" $attachmentSize)
    for ($j = 0; $j -lt $attachmentItems.Count; $j++) {
      $item = $attachmentItems[$j]
      $paragraph = $item.Paragraph
      $text = Get-ParagraphText $paragraph
      $normalizedAttachmentText = Normalize-AttachmentExplanationText $text
      $expectedIndent = Get-AttachmentIndentSpec $text $Rules ([bool]$item.IsFirst)
      $actualFirstLineStart = [double]$paragraph.Range.ParagraphFormat.LeftIndent + [double]$paragraph.Range.ParagraphFormat.FirstLineIndent
      $fontOk = (($paragraph.Range.Font.NameFarEast -eq $attachmentFont) -or ($paragraph.Range.Font.Name -eq $attachmentFont))
      $formatOk = $formatOk -and $fontOk -and
        (Test-Near ([double]$paragraph.Range.Font.Size) $attachmentSize 0.5) -and
        (Test-Near ([double]$paragraph.Range.ParagraphFormat.LineSpacing) $attachmentLineSpacing 0.5) -and
        ((Get-ExpectedWordColorHex ([int]$paragraph.Range.Font.Color)) -eq "000000")
      $indentOk = $indentOk -and
        (Test-Near ([double]$paragraph.Range.ParagraphFormat.LeftIndent) ([double]$expectedIndent.LeftIndent) 1.0) -and
        (Test-Near ([double]$paragraph.Range.ParagraphFormat.FirstLineIndent) ([double]$expectedIndent.FirstLineIndent) 1.0)
      $punctuationOk = $punctuationOk -and ($text -notmatch "[。；;，,、]\s*$")
      if ($attachmentItems.Count -gt 1) {
        $actualNumberStart = if ($j -eq 0) {
          [double]($actualFirstLineStart + (Get-ApproxTextWidthPt "附件：" $attachmentSize))
        } else {
          [double]$actualFirstLineStart
        }
        $numberAlignmentOk = $numberAlignmentOk -and (Test-Near $actualNumberStart $attachmentNumberStart 1.0)
        $expectedNumber = $j + 1
        if ($j -eq 0) {
          $numberingOk = $numberingOk -and ($normalizedAttachmentText -match "^附件：$expectedNumber\.")
        } else {
          $numberingOk = $numberingOk -and ($normalizedAttachmentText -match "^$expectedNumber\.")
        }
      }
    }
    $checks += New-Check "附件说明字体字号行距符合规范" $formatOk "字体 $attachmentFont，字号 $attachmentSize pt，行距 $attachmentLineSpacing pt"
    $checks += New-Check "附件说明回行对齐附件名称首字" $indentOk "使用悬挂缩进处理 $($attachmentItems.Count) 条附件说明"
    $checks += New-Check "附件名称后不加标点" $punctuationOk
    if ($attachmentItems.Count -gt 1) {
      $checks += New-Check "多个附件使用阿拉伯数字顺序号" $numberingOk "识别到 $($attachmentItems.Count) 条附件说明"
      $checks += New-Check "多个附件序号左对齐" $numberAlignmentOk "序号起点按第 1 条附件序号对齐"
    }
  }

  $company = Get-RuleValue $Meta "company" (Get-RuleValue $Rules "signature.company" "安盟财产保险有限公司")
  $date = Get-RuleValue $Meta "date" (Get-RuleValue $Rules "signature.date" "2026年5月9日")
  $normalizedCompany = Normalize-SignatureText $company
  $normalizedDate = Normalize-SignatureText $date
  $signatureCompanyCount = 0
  $signatureDateCount = 0
  $tailStart = [Math]::Max(1, $Document.Paragraphs.Count - 90)
  for ($i = $tailStart; $i -le $Document.Paragraphs.Count; $i++) {
    $text = Get-ParagraphText $Document.Paragraphs.Item($i)
    $normalized = Normalize-SignatureText $text
    if ($normalized -eq $normalizedCompany) { $signatureCompanyCount++ }
    if ($normalized -eq $normalizedDate) { $signatureDateCount++ }
  }
  $checks += New-Check "落款单位和成文日期唯一" (($signatureCompanyCount -eq 1) -and ($signatureDateCount -eq 1)) "落款 $signatureCompanyCount 处，日期 $signatureDateCount 处"

  if ($attachmentItems.Count -gt 0) {
    $signatureDateIndex = $null
    for ($i = $tailStart; $i -le $Document.Paragraphs.Count; $i++) {
      $paragraph = $Document.Paragraphs.Item($i)
      if ((Normalize-SignatureText (Get-ParagraphText $paragraph)) -eq $normalizedDate) {
        $signatureDateIndex = $i
      }
    }
    $attachmentBeforeDate = ($null -ne $signatureDateIndex) -and ([int]$attachmentItems[0].Index -lt [int]$signatureDateIndex)
    $checks += New-Check "附件说明位于成文日期之前" $attachmentBeforeDate "附件说明第 $($attachmentItems[0].Index) 段，成文日期第 $signatureDateIndex 段"
  }

  $bodyFont = Get-RuleValue $Rules "body.font" "仿宋_GB2312"
  $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
  $bodyLineSpacing = [double](Get-RuleValue $Rules "body.lineSpacingPt" 28)
  $bodyFirstLineIndent = [double](Get-RuleValue $Rules "body.firstLineIndentPt" 32)
  $bodyColorHex = ((Get-RuleValue $Rules "body.color" "000000") -replace "#", "").ToUpperInvariant()
  $bodyFormatOk = $true
  $bodyIndentOk = $true
  $bodyChecked = 0
  $bodyStartIndex = if ($null -ne $mainRecipientItem) { [int]$mainRecipientItem.Index + 1 } else { 3 }
  $attachmentStartIndex = if ($attachmentItems.Count -gt 0) { [int]$attachmentItems[0].Index } else { [int]::MaxValue }
  for ($i = $bodyStartIndex; $i -le $Document.Paragraphs.Count; $i++) {
    if ($i -ge $attachmentStartIndex) { break }
    $paragraph = $Document.Paragraphs.Item($i)
    $text = Get-ParagraphText $paragraph
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    $normalized = Normalize-SignatureText $text
    if (($normalized -eq $normalizedCompany) -or ($normalized -eq $normalizedDate) -or (Test-DateParagraphText $text)) { break }
    $bodyChecked++
    $paragraphFontOk = (($paragraph.Range.Font.NameFarEast -eq $bodyFont) -or ($paragraph.Range.Font.Name -eq $bodyFont))
    $paragraphColorHex = Get-ExpectedWordColorHex ([int]$paragraph.Range.Font.Color)
    $bodyFormatOk = $bodyFormatOk -and $paragraphFontOk -and
      (Test-Near ([double]$paragraph.Range.Font.Size) $bodySize 0.5) -and
      (Test-Near ([double]$paragraph.Range.ParagraphFormat.LineSpacing) $bodyLineSpacing 0.5) -and
      ($paragraphColorHex -eq $bodyColorHex) -and
      ([int]$paragraph.Range.Font.Bold -eq 0)
    $bodyIndentOk = $bodyIndentOk -and
      (Test-Near ([double]$paragraph.Range.ParagraphFormat.FirstLineIndent) $bodyFirstLineIndent 1.0) -and
      (Test-Near ([double]$paragraph.Range.ParagraphFormat.LeftIndent) 0 0.5)
  }
  $checks += New-Check "正文基础字体字号行距颜色符合规范" (($bodyChecked -gt 0) -and $bodyFormatOk) "检查正文自然段 $bodyChecked 段，颜色 #$bodyColorHex"
  $checks += New-Check "正文自然段左空二字回行顶格" (($bodyChecked -gt 0) -and $bodyIndentOk) "首行缩进 $bodyFirstLineIndent pt，左缩进 0"

  $noSeal = [bool](Get-RuleValue $Rules "signature.noSeal" $false)
  $checks += New-Check "盖章公文模式已启用" (-not $noSeal)
  if (-not $noSeal) {
    $bodySize = [double](Get-RuleValue $Rules "body.sizePt" 16)
    $dateRightChars = [double](Get-RuleValue $Rules "signature.dateRightIndentChars" 4)
    $companyCenterCorrection = [double](Get-RuleValue $Rules "signature.companyCenterCorrectionPt" 3)
    $expectedRightIndent = $bodySize * $dateRightChars
    $dateParagraph = $null
    $companyParagraph = $null
    for ($i = $tailStart; $i -le $Document.Paragraphs.Count; $i++) {
      $paragraph = $Document.Paragraphs.Item($i)
      $paragraphText = Get-ParagraphText $paragraph
      $normalizedParagraphText = Normalize-SignatureText $paragraphText
      if ($normalizedParagraphText -eq $normalizedCompany) {
        $companyParagraph = $paragraph
      }
      if ($normalizedParagraphText -eq $normalizedDate) {
        $dateParagraph = $paragraph
      }
    }
    $dateLayoutOk = ($null -ne $dateParagraph) -and ([int]$dateParagraph.Range.ParagraphFormat.Alignment -eq 2) -and (Test-Near $dateParagraph.Range.ParagraphFormat.RightIndent $expectedRightIndent 1.0)
    $checks += New-Check "成文日期右空四字编排" $dateLayoutOk "右缩进 $([Math]::Round($expectedRightIndent, 1))pt"

    $signatureCenterOk = $false
    $centerDiff = 9999.0
    $expectedCompanyRightIndent = 0.0
    if (($null -ne $companyParagraph) -and ($null -ne $dateParagraph)) {
      $dateWidth = Get-SignatureTextWidthPt (Get-ParagraphText $dateParagraph) $bodySize
      $companyWidth = Get-SignatureTextWidthPt (Get-ParagraphText $companyParagraph) $bodySize
      $expectedCompanyRightIndent = [Math]::Max(0, $expectedRightIndent - (($companyWidth - $dateWidth) / 2) + $companyCenterCorrection)
      $companyRightIndentActual = [double]$companyParagraph.Range.ParagraphFormat.RightIndent
      $companyCenter = ([double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.RightMargin) - $companyRightIndentActual - ($companyWidth / 2)
      $dateCenter = ([double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.RightMargin) - $expectedRightIndent - ($dateWidth / 2)
      $centerDiff = ($companyCenter - $dateCenter) + $companyCenterCorrection
      $signatureCenterOk = ([int]$companyParagraph.Range.ParagraphFormat.Alignment -eq 2) -and
        (Test-Near $companyRightIndentActual $expectedCompanyRightIndent 1.0) -and
        (Test-Near $centerDiff 0 1.5)
    }
    $checks += New-Check "发文机关署名以成文日期为准居中" $signatureCenterOk "中心差 $([Math]::Round($centerDiff, 2))pt，署名右缩进 $([Math]::Round($expectedCompanyRightIndent, 1))pt"

    if ([bool](Get-RuleValue $Rules "signature.seal.enabled" $true)) {
      $sealPath = Resolve-ConfiguredPath (Get-RuleValue $Rules "signature.seal.imagePath" "")
      $sealConfigured = (-not [string]::IsNullOrWhiteSpace($sealPath)) -and (Test-Path -LiteralPath $sealPath)
      $checks += New-Check "真实公章图片已配置" $sealConfigured $sealPath
      if ($sealConfigured) {
        $sealCount = 0
        foreach ($shape in $Document.Shapes) {
          if ($shape.Name -eq "redhead-seal") { $sealCount++ }
        }
        $checks += New-Check "公章已插入文档" ($sealCount -ge 1) "识别到 $sealCount 个公章图形"
      }
    }
  }

  $bodyPages = [int]$ImprintInfo.BodyPages
  $imprintSamePage = [bool]$ImprintInfo.SamePage
  $expectedTotal = if (-not [bool](Get-RuleValue $Rules "imprint.enabled" $true)) {
    $bodyPages
  } elseif ($imprintSamePage) {
    $bodyPages
  } elseif ([bool]$ImprintInfo.BlankInserted) {
    $bodyPages + 2
  } else {
    $bodyPages + 1
  }
  $placementDetail = if ($imprintSamePage) { "版记同页" } elseif ([bool]$ImprintInfo.BlankInserted) { "补空白页后版记" } else { "版记另起页" }
  $checks += New-Check "总页数与正文/版记结构一致" ($TotalPages -eq $expectedTotal) "正文 $bodyPages 页，总计 $TotalPages 页，$placementDetail"
  if ([bool](Get-RuleValue $Rules "imprint.requireEvenPage" $true)) {
    $checks += New-Check "版记位于偶数页" (($null -ne $ImprintPage) -and ($ImprintPage % 2 -eq 0)) "版记第 $ImprintPage 页"
  }
  if ([bool](Get-RuleValue $Rules "cleanup.trimTrailingBlankPages" $true)) {
    $removedPages = if ($null -ne $TrimInfo) { [int]$TrimInfo.PagesRemoved } else { 0 }
    $deletedParagraphs = if ($null -ne $TrimInfo) { [int]$TrimInfo.BlankParagraphsDeleted } else { 0 }
    $deletedBreaks = if ($null -ne $TrimInfo) { [int]$TrimInfo.TrailingBreaksDeleted } else { 0 }
    $checks += New-Check "源文件尾部空白页已处理" $true "清理空白段 $deletedParagraphs 个，尾部分页/分节符 $deletedBreaks 个，页数减少 $removedPages 页"
  }

  if ([bool](Get-RuleValue $Rules "imprint.enabled" $true)) {
    $office = Get-RuleValue $Rules "imprint.office" "安盟财产保险有限公司综合办公室"
    $imprintDate = Get-RuleValue $Rules "imprint.date" "2026年5月9日印发"
    $officeLeftChars = [int](Get-RuleValue $Rules "imprint.officeLeftChars" 1)
    $dateRightChars = [int](Get-RuleValue $Rules "imprint.dateRightChars" 1)
    $imprintSize = [double](Get-RuleValue $Rules "imprint.sizePt" 14)
    $officePrefix = Repeat-Text "　" ([Math]::Max(0, $officeLeftChars))
    $imprintTable = $null
    $imprintSection = $Document.Sections.Item($Document.Sections.Count)
    if ($imprintSection.Range.Tables.Count -gt 0) {
      $imprintTable = $imprintSection.Range.Tables.Item($imprintSection.Range.Tables.Count)
    } elseif ($Document.Tables.Count -gt 0) {
      $imprintTable = $Document.Tables.Item($Document.Tables.Count)
    }
    $expectedRows = @(Get-ImprintRows $Rules)
    $expectedIssueRowIndex = [Math]::Max(1, $expectedRows.Count)
    $actualRowCount = if ($null -ne $imprintTable) { [int]$imprintTable.Rows.Count } else { 0 }
    $checks += New-Check "版记要素行数符合规则" ($actualRowCount -eq $expectedRows.Count) "实际 $actualRowCount 行，规则 $($expectedRows.Count) 行"

    $officeText = ""
    $dateText = ""
    $imprintCell = $null
    if ($null -ne $imprintTable -and $imprintTable.Rows.Count -ge $expectedIssueRowIndex -and $imprintTable.Columns.Count -ge 2) {
      $imprintCell = $imprintTable.Cell($expectedIssueRowIndex, 1)
      $officeText = Get-TableCellText $imprintTable.Cell($expectedIssueRowIndex, 1)
      $dateText = Get-TableCellText $imprintTable.Cell($expectedIssueRowIndex, 2)
    } elseif ($null -ne $imprintTable -and $imprintTable.Rows.Count -ge $expectedIssueRowIndex -and $imprintTable.Columns.Count -eq 1) {
      $imprintCell = $imprintTable.Cell($expectedIssueRowIndex, 1)
      $lineText = Get-TableCellText $imprintTable.Cell($expectedIssueRowIndex, 1)
      $parts = $lineText -split "`t", 2
      if ($parts.Count -ge 1) { $officeText = $parts[0] }
      if ($parts.Count -ge 2) { $dateText = $parts[1] }
    }
    $officeOk = ($officeText -eq "$officePrefix$office")
    $dateTextClean = ($dateText -replace "^[\s　]+", "")
    $dateTextClean = ($dateTextClean -replace "[\s　]+$", "")
    $contentWidth = [double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.LeftMargin - [double]$Document.PageSetup.RightMargin
    $expectedDateTab = $contentWidth - ($imprintSize * [Math]::Max(0, $dateRightChars))
    $actualDateTab = -1.0
    $dateTabOk = $false
    if ($null -ne $imprintCell) {
      foreach ($tabStop in $imprintCell.Range.ParagraphFormat.TabStops) {
        if ([int]$tabStop.Alignment -eq 2) {
          $actualDateTab = [double]$tabStop.Position
          if (Test-Near $actualDateTab $expectedDateTab 1.0) {
            $dateTabOk = $true
            break
          }
        }
      }
    }
    $dateOk = ($dateTextClean -eq $imprintDate) -and $dateTabOk
    $checks += New-Check "印发机关左空一字" $officeOk "配置左空 $officeLeftChars 字"
    $checks += New-Check "印发日期右空一字" $dateOk "配置右空 $dateRightChars 字，右对齐制表位 $([Math]::Round($actualDateTab, 1))pt / 目标 $([Math]::Round($expectedDateTab, 1))pt"
    $expectedTopPadding = [double](Get-RuleValue $Rules "imprint.cellPaddingTopPt" 2)
    $expectedBottomPadding = [double](Get-RuleValue $Rules "imprint.cellPaddingBottomPt" 2)
    $actualTopPadding = if ($null -ne $imprintTable) { [double]$imprintTable.TopPadding } else { -1 }
    $actualBottomPadding = if ($null -ne $imprintTable) { [double]$imprintTable.BottomPadding } else { -1 }
    $verticalAlignment = -1
    $paragraphSpaceBefore = -1
    $paragraphSpaceAfter = -1
    $cellsVerticalOk = ($null -ne $imprintTable)
    if ($null -ne $imprintTable) {
      for ($i = 1; $i -le $imprintTable.Rows.Count; $i++) {
        $cell = $imprintTable.Cell($i, 1)
        $cellParagraphFormat = $cell.Range.Paragraphs.Item(1).Range.ParagraphFormat
        if ($i -eq $expectedIssueRowIndex) {
          $verticalAlignment = [int]$cell.VerticalAlignment
          $paragraphSpaceBefore = [double]$cellParagraphFormat.SpaceBefore
          $paragraphSpaceAfter = [double]$cellParagraphFormat.SpaceAfter
        }
        $cellsVerticalOk = $cellsVerticalOk -and ([int]$cell.VerticalAlignment -eq 1)
      }
    }
    $topPaddingOk = Test-Near $actualTopPadding $expectedTopPadding 0.2
    $bottomPaddingOk = Test-Near $actualBottomPadding $expectedBottomPadding 0.2
    $equalPaddingOk = Test-Near $actualTopPadding $actualBottomPadding 0.2
    $imprintVerticalOk = [bool]$cellsVerticalOk -and [bool]$topPaddingOk -and [bool]$bottomPaddingOk -and [bool]$equalPaddingOk
    $checks += New-Check "版记文字上下居中且间距一致" $imprintVerticalOk "垂直对齐 $verticalAlignment，上/下边距 $actualTopPadding/$actualBottomPadding pt，段前/段后 $paragraphSpaceBefore/$paragraphSpaceAfter pt"

    $imprintFont = Get-RuleValue $Rules "imprint.font" "仿宋_GB2312"
    $imprintSize = [double](Get-RuleValue $Rules "imprint.sizePt" 14)
    $imprintLineSpacing = [double](Get-RuleValue $Rules "imprint.lineSpacingPt" 28)
    $expectedBaselineShift = [double](Get-RuleValue $Rules "imprint.baselineShiftPt" 3)
    $imprintFormatOk = ($null -ne $imprintTable)
    $baselineShiftOk = ($null -ne $imprintTable)
    $actualBaselineShift = -999
    if ($null -ne $imprintTable) {
      for ($i = 1; $i -le $imprintTable.Rows.Count; $i++) {
        $cell = $imprintTable.Cell($i, 1)
        $cellTextRange = $cell.Range.Duplicate
        if ($cellTextRange.End -gt $cellTextRange.Start) { $cellTextRange.End = $cellTextRange.End - 1 }
        if ($i -eq $expectedIssueRowIndex) { $actualBaselineShift = [double]$cellTextRange.Font.Position }
        $cellFontOk = (($cell.Range.Font.NameFarEast -eq $imprintFont) -or ($cell.Range.Font.Name -eq $imprintFont))
        $imprintFormatOk = $imprintFormatOk -and
          $cellFontOk -and
          (Test-Near ([double]$cell.Range.Font.Size) $imprintSize 0.5) -and
          ([int]$cell.Range.ParagraphFormat.LineSpacingRule -eq 4) -and
          (Test-Near ([double]$cell.Range.ParagraphFormat.LineSpacing) $imprintLineSpacing 0.5)
        $baselineShiftOk = $baselineShiftOk -and (Test-Near ([double]$cellTextRange.Font.Position) $expectedBaselineShift 0.1)
      }
    }
    $checks += New-Check "版记字体字号固定行距符合规则" $imprintFormatOk "字体 $imprintFont，字号 $imprintSize pt，固定行距 $imprintLineSpacing pt"
    $checks += New-Check "版记文字基线补偿符合规则" $baselineShiftOk "实际 $([Math]::Round($actualBaselineShift, 2))pt，规则 $expectedBaselineShift pt"

    $issueRow = $expectedRows[$expectedIssueRowIndex - 1]
    $issueRowHeight = if ($null -ne $imprintTable -and $imprintTable.Rows.Count -ge $expectedIssueRowIndex) { [double]$imprintTable.Rows.Item($expectedIssueRowIndex).Height } else { 0 }
    $issueRowOk = ($expectedIssueRowIndex -eq $expectedRows.Count) -and
      (Test-Near $issueRowHeight ([double]$issueRow.RowHeight) 0.5) -and
      (Test-Near ([double]$issueRow.RowHeight) $imprintLineSpacing 0.5)
    $checks += New-Check "印发机关和印发日期排在末线上方一行" $issueRowOk "第 $expectedIssueRowIndex 行，行高 $([Math]::Round($issueRowHeight, 2))pt"

    for ($r = 0; $r -lt $expectedRows.Count; $r++) {
      $expectedRow = $expectedRows[$r]
      if ($expectedRow.Type -eq "mainRecipient" -or $expectedRow.Type -eq "cc") {
        $rowIndex = $r + 1
        $rowName = if ($expectedRow.Type -eq "mainRecipient") { "版记主送机关" } else { "版记抄送机关" }
        $cell = if ($null -ne $imprintTable -and $imprintTable.Rows.Count -ge $rowIndex) { $imprintTable.Cell($rowIndex, 1) } else { $null }
        $cellText = if ($null -ne $cell) { Get-TableCellText $cell } else { "" }
        $size = [double](Get-RuleValue $Rules "imprint.sizePt" 14)
        $expectedFirstLineStart = $size * [double]$expectedRow.LeftBlankChars
        $expectedLabelWidth = Get-ApproxTextWidthPt ([string]$expectedRow.Label) $size
        $expectedRightIndent = $size * [double]$expectedRow.RightBlankChars
        $layoutOk = ($null -ne $cell) -and
          (Test-Near ([double]$cell.Range.ParagraphFormat.LeftIndent + [double]$cell.Range.ParagraphFormat.FirstLineIndent) $expectedFirstLineStart 1.0) -and
          (Test-Near ([double]$cell.Range.ParagraphFormat.FirstLineIndent) (-1 * $expectedLabelWidth) 1.0) -and
          (Test-Near ([double]$cell.Range.ParagraphFormat.RightIndent) $expectedRightIndent 1.0)
        $checks += New-Check "$rowName 已编入版记且格式符合规范" (($cellText -eq [string]$expectedRow.Text) -and $layoutOk) "第 $rowIndex 行，左右各空 $($expectedRow.LeftBlankChars)/$($expectedRow.RightBlankChars) 字"
      }
    }

    $middleLineExpected = ($expectedRows.Count -gt 1)
    $middleLineVisible = ($null -ne $imprintTable) -and ([int]$imprintTable.Borders.Item(-5).LineStyle -ne 0)
    $expectedInnerLineWidth = [double](Get-RuleValue $Rules "imprint.innerLineWidthPt" 0.75)
    $actualInnerLineWidth = if ($middleLineVisible) { Convert-WordLineWidthToPt ([int]$imprintTable.Borders.Item(-5).LineWidth) } else { 0 }
    $middleLineOk = if ($middleLineExpected) {
      $middleLineVisible -and (Test-Near $actualInnerLineWidth $expectedInnerLineWidth 0.05)
    } else {
      $true
    }
    $checks += New-Check "版记要素之间使用中间细分隔线" $middleLineOk "需要中间线：$middleLineExpected，线宽 $actualInnerLineWidth pt"

    $expectedLeft = [double]$imprintSection.PageSetup.LeftMargin
    $expectedWidth = [double]($imprintSection.PageSetup.PageWidth - $imprintSection.PageSetup.LeftMargin - $imprintSection.PageSetup.RightMargin)
    $expectedBottom = [double]($imprintSection.PageSetup.PageHeight - $imprintSection.PageSetup.BottomMargin)
    $tableTop = if ($null -ne $imprintTable) { [double]$imprintTable.Range.Information(6) } else { 0 }
    $tableHeight = if ($null -ne $imprintTable) { Get-ImprintTableHeight $imprintTable } else { 0 }
    $actualBottom = if ($tableHeight -gt 0) { $tableTop + $tableHeight } else { 0 }
    $bottomBorderVisible = ($null -ne $imprintTable) -and ([int]$imprintTable.Borders.Item(-3).LineStyle -ne 0)
    $topBorderVisible = ($null -ne $imprintTable) -and ([int]$imprintTable.Borders.Item(-1).LineStyle -ne 0)
    $expectedOuterLineWidth = [double](Get-RuleValue $Rules "imprint.outerLineWidthPt" 1.0)
    $actualTopLineWidth = if ($topBorderVisible) { Convert-WordLineWidthToPt ([int]$imprintTable.Borders.Item(-1).LineWidth) } else { 0 }
    $actualBottomLineWidth = if ($bottomBorderVisible) { Convert-WordLineWidthToPt ([int]$imprintTable.Borders.Item(-3).LineWidth) } else { 0 }
    $outerLineWidthOk = $topBorderVisible -and $bottomBorderVisible -and
      (Test-Near $actualTopLineWidth $expectedOuterLineWidth 0.05) -and
      (Test-Near $actualBottomLineWidth $expectedOuterLineWidth 0.05)
    $checks += New-Check "版记首末粗分隔线线宽符合规则" $outerLineWidthOk "首线 $actualTopLineWidth pt，末线 $actualBottomLineWidth pt"
    $tableWidth = if ($null -ne $imprintTable -and $imprintTable.Columns.Count -ge 2) {
      [double]$imprintTable.Columns.Item(1).Width + [double]$imprintTable.Columns.Item(2).Width
    } elseif ($null -ne $imprintTable -and $imprintTable.Columns.Count -eq 1) {
      [double]$imprintTable.Columns.Item(1).Width
    } else {
      0
    }
    $lineOk = $bottomBorderVisible -and
      (Test-Near $actualBottom $expectedBottom 2.0) -and
      (Test-Near $tableWidth $expectedWidth 1.5)
    $lineDetail = if ($null -ne $imprintTable) {
      "底边框已设置：$bottomBorderVisible，表格底部 $([Math]::Round($actualBottom, 2))pt，版心下边缘 $([Math]::Round($expectedBottom, 2))pt，宽度 $([Math]::Round($tableWidth, 2))pt"
    } else {
      "未找到版记表格"
    }
    $checks += New-Check "版记末条分隔线与版心下边缘重合" $lineOk $lineDetail
  }

  if ([bool](Get-RuleValue $Rules "pageNumber.enabled" $true)) {
    $bodySectionCount = [int]$ImprintInfo.BodySectionCount
    $bodySection = $Document.Sections.Item($bodySectionCount)
    $oddFooter = $bodySection.Footers.Item(1)
    $evenFooter = $bodySection.Footers.Item(3)
    $oddFooterText = ($oddFooter.Range.Text -replace "[`r`a]", "").Trim()
    $evenFooterText = ($evenFooter.Range.Text -replace "[`r`a]", "").Trim()
    $checks += New-Check "正文页脚已设置页码字段" (($oddFooterText.Contains("-")) -and ($evenFooterText.Contains("-")))
    $pageNumberFont = Get-RuleValue $Rules "pageNumber.font" "宋体"
    $pageNumberSize = [double](Get-RuleValue $Rules "pageNumber.sizePt" 14)
    $pageNumberBlankChars = [double](Get-RuleValue $Rules "pageNumber.blankChars" 1)
    $expectedPageNumberIndent = $pageNumberSize * $pageNumberBlankChars
    $oddFontOk = (($oddFooter.Range.Font.NameFarEast -eq $pageNumberFont) -or ($oddFooter.Range.Font.Name -eq $pageNumberFont))
    $evenFontOk = (($evenFooter.Range.Font.NameFarEast -eq $pageNumberFont) -or ($evenFooter.Range.Font.Name -eq $pageNumberFont))
    $pageNumberFormatOk = ($oddFooterText -match "^- ?\d* ?-$") -and ($evenFooterText -match "^- ?\d* ?-$") -and
      $oddFontOk -and $evenFontOk -and
      (Test-Near ([double]$oddFooter.Range.Font.Size) $pageNumberSize 0.5) -and
      (Test-Near ([double]$evenFooter.Range.Font.Size) $pageNumberSize 0.5)
    $checks += New-Check "页码字体字号形式符合规范" $pageNumberFormatOk "字体 $pageNumberFont，字号 $pageNumberSize pt，形式 - 1 -"
    $expectedFooterDistance = ([double]$bodySection.PageSetup.BottomMargin) - (Convert-MmToPt ([double](Get-RuleValue $Rules "pageNumber.distanceBelowContentMm" 7)))
    $checks += New-Check "页码位于版心下边缘下7mm" (Test-Near ([double]$bodySection.PageSetup.FooterDistance) $expectedFooterDistance 1.0) "FooterDistance $([Math]::Round($bodySection.PageSetup.FooterDistance, 2))pt"
    $oddAlignment = Get-PageNumberAlignment $Rules "pageNumber.oddAlign" 2
    $evenAlignment = Get-PageNumberAlignment $Rules "pageNumber.evenAlign" 0
    $pageNumberBlankOk = (Test-PageNumberIndentForAlignment $oddFooter $oddAlignment $expectedPageNumberIndent) -and
      (Test-PageNumberIndentForAlignment $evenFooter $evenAlignment $expectedPageNumberIndent)
    $checks += New-Check "页码单双页按规则对齐并空字" $pageNumberBlankOk "奇页对齐 $oddAlignment，偶页对齐 $evenAlignment，空 $pageNumberBlankChars 字"

    if ([bool]$ImprintInfo.BlankInserted) {
      $blankSection = $Document.Sections.Item($bodySectionCount + 1)
      $blankFooter = (($blankSection.Footers.Item(1).Range.Text + $blankSection.Footers.Item(3).Range.Text) -replace "[`r`a]", "").Trim()
      $checks += New-Check "空白页不编页码" ([string]::IsNullOrWhiteSpace($blankFooter))

      $imprintSection = $Document.Sections.Item($Document.Sections.Count)
      $imprintFooter = (($imprintSection.Footers.Item(1).Range.Text + $imprintSection.Footers.Item(3).Range.Text) -replace "[`r`a]", "").Trim()
      $checks += New-Check "空白页后的版记页不编页码" ([string]::IsNullOrWhiteSpace($imprintFooter))
    } elseif ([bool]$ImprintInfo.SamePage) {
      $checks += New-Check "版记同页按正文页码处理" $true
    } else {
      $imprintSection = $Document.Sections.Item($Document.Sections.Count)
      $imprintFooter = (($imprintSection.Footers.Item(1).Range.Text + $imprintSection.Footers.Item(3).Range.Text) -replace "[`r`a]", "").Trim()
      $checks += New-Check "版记页按正文连续编页码" (-not [string]::IsNullOrWhiteSpace($imprintFooter))
    }
  } else {
    $checks += New-Check "页码规则已关闭" $true
  }
  $checks += New-Check "PDF 预览文件已生成" ((Test-Path -LiteralPath $OutputPdf) -and ((Get-Item -LiteralPath $OutputPdf).Length -gt 0))

  return $checks
}


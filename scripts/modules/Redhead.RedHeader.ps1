# 版头（发文机关标志、发文字号、版头红线）
# 本模块由 Redhead.Core.ps1 拆分而来，函数定义在 dot-source 后与其它模块共享作用域。

function Set-RedLineBlankParagraph {
  param($Paragraph, $Rules)
  $blankFont = Get-RuleValue $Rules "body.font" "仿宋_GB2312"
  $blankSize = [double](Get-RuleValue $Rules "title.sizePt" 22)
  $lineSpacing = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)

  Set-RangeFont $Paragraph.Range $blankFont $blankSize "" 0
  Clear-ParagraphSpacingUnits $Paragraph
  $Paragraph.Range.ParagraphFormat.Alignment = 1
  $Paragraph.Range.ParagraphFormat.LineSpacingRule = 4
  $Paragraph.Range.ParagraphFormat.LineSpacing = $lineSpacing
  $Paragraph.Range.ParagraphFormat.FirstLineIndent = 0
  $Paragraph.Range.ParagraphFormat.LeftIndent = 0
  $Paragraph.Range.ParagraphFormat.RightIndent = 0
  $Paragraph.Range.ParagraphFormat.SpaceBefore = 0
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
}

function Ensure-BlankParagraphsAfterDocumentNo {
  param($Document, [int]$DocNoIndex, $Rules, [int]$Count = 2)

  $existing = 0
  for ($i = $DocNoIndex + 1; $i -le $Document.Paragraphs.Count; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    if (-not (Test-BlankParagraph $paragraph)) { break }
    $existing++
    if ($existing -ge $Count) { break }
  }

  $missing = [Math]::Max(0, $Count - $existing)
  if ($missing -gt 0) {
    $insertPosition = if (($DocNoIndex + 1) -le $Document.Paragraphs.Count) {
      $Document.Paragraphs.Item($DocNoIndex + 1).Range.Start
    } else {
      $Document.Paragraphs.Item($DocNoIndex).Range.End
    }
    $insertRange = $Document.Range($insertPosition, $insertPosition)
    $insertRange.InsertBefore((Repeat-Text "`r" $missing))
  }

  $blanks = @()
  for ($i = $DocNoIndex + 1; $i -le $Document.Paragraphs.Count; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    if (-not (Test-BlankParagraph $paragraph)) { break }
    Set-RedLineBlankParagraph $paragraph $Rules
    $blanks += $paragraph
    if ($blanks.Count -ge $Count) { break }
  }
  return @($blanks)
}

function Find-DocumentNoRedLineShape {
  param($Document, $Rules)

  $expectedColorHex = ((Get-RuleValue $Rules "documentNo.redLineColor" "FF0000") -replace "#", "").ToUpperInvariant()
  $expectedWidth = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineLengthMm" 156))
  foreach ($shape in $Document.Shapes) {
    try {
      $nameMatch = ($shape.Name -eq "redhead-document-no-red-line")
      $isHorizontalLine = ([Math]::Abs([double]$shape.Height) -le 2.0)
      $colorMatch = ((Get-ExpectedWordColorHex ([int]$shape.Line.ForeColor.RGB)) -eq $expectedColorHex)
      $widthMatch = Test-Near ([double]$shape.Width) $expectedWidth 3.0
      if ($nameMatch -or ($isHorizontalLine -and $colorMatch -and $widthMatch)) {
        return $shape
      }
    } catch {}
  }
  return $null
}

function Get-DocumentNoRedLineInfo {
  param($Document, $Rules, $DocNoParagraph)

  $lineBoxHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  $shape = Find-DocumentNoRedLineShape $Document $Rules
  if ($null -ne $shape) {
    $anchorTop = 0.0
    try { $anchorTop = [double]$shape.Anchor.Information(6) } catch {}
    $absoluteTop = $anchorTop + [double]$shape.Top
    $absoluteLeft = [double]$Document.PageSetup.LeftMargin + [double]$shape.Left
    $lineWidth = [double]$shape.Line.Weight
    return [pscustomobject]@{
      Exists = $true
      Source = "shape"
      Shape = $shape
      Border = $null
      ColorHex = Get-ExpectedWordColorHex ([int]$shape.Line.ForeColor.RGB)
      LineWidthPt = $lineWidth
      LengthPt = [double]$shape.Width
      CenterPt = $absoluteLeft + ([double]$shape.Width / 2)
      TopPt = $absoluteTop
      BottomPt = $absoluteTop + $lineWidth
      OffsetPt = [double]$shape.Top
    }
  }

  $border = $DocNoParagraph.Range.ParagraphFormat.Borders.Item(-3)
  if ([int]$border.LineStyle -ne 0) {
    $lineWidth = Convert-WordLineWidthToPt ([int]$border.LineWidth)
    $offset = 0.0
    try { $offset = [double]$DocNoParagraph.Range.ParagraphFormat.Borders.DistanceFromBottom } catch {}
    $contentWidth = [double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.LeftMargin - [double]$Document.PageSetup.RightMargin
    $length = $contentWidth - [double]$DocNoParagraph.Range.ParagraphFormat.LeftIndent - [double]$DocNoParagraph.Range.ParagraphFormat.RightIndent
    $top = [double]$DocNoParagraph.Range.Information(6) + $lineBoxHeight + $offset
    return [pscustomobject]@{
      Exists = $true
      Source = "border"
      Shape = $null
      Border = $border
      ColorHex = Get-ExpectedWordColorHex ([int]$border.Color)
      LineWidthPt = $lineWidth
      LengthPt = $length
      CenterPt = [double]$Document.PageSetup.LeftMargin + [double]$DocNoParagraph.Range.ParagraphFormat.LeftIndent + ($length / 2)
      TopPt = $top
      BottomPt = $top + $lineWidth
      OffsetPt = $offset
    }
  }

  return [pscustomobject]@{
    Exists = $false
    Source = ""
    Shape = $null
    Border = $null
    ColorHex = ""
    LineWidthPt = 0.0
    LengthPt = 0.0
    CenterPt = 0.0
    TopPt = 0.0
    BottomPt = 0.0
    OffsetPt = 0.0
  }
}

function Add-DocumentNoRedLine {
  param($Document, $Paragraph, $Rules)

  Remove-NamedShapes $Document "redhead-document-no-red-line"

  $lineWidth = [double](Get-RuleValue $Rules "documentNo.redLineWidthPt" 2.25)
  $lineColor = Convert-HexToWordColor (Get-RuleValue $Rules "documentNo.redLineColor" "FF0000")
  $lineOffset = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 1.8))
  $lineLength = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineLengthMm" 156))
  Clear-ParagraphSpacingUnits $Paragraph
  $Paragraph.Range.ParagraphFormat.SpaceBefore = 0
  $Paragraph.Range.ParagraphFormat.SpaceAfter = 0
  foreach ($borderIndex in @(-1, -2, -3, -4)) {
    try { $Paragraph.Range.ParagraphFormat.Borders.Item($borderIndex).LineStyle = 0 } catch {}
  }

  $docNoIndex = Find-DocumentNoParagraphIndex $Document
  $blankParagraphs = @(Ensure-BlankParagraphsAfterDocumentNo $Document ([int]$docNoIndex) $Rules 2)
  if ($blankParagraphs.Count -eq 0) { return }

  $anchorParagraph = $blankParagraphs[0]
  $shape = $Document.Shapes.AddLine(0, 0, $lineLength, 0, $anchorParagraph.Range)
  $shape.Name = "redhead-document-no-red-line"
  $shape.Line.ForeColor.RGB = $lineColor
  $shape.Line.Weight = $lineWidth
  $shape.RelativeHorizontalPosition = 2
  $shape.RelativeVerticalPosition = 2
  $contentWidth = [double]$Document.PageSetup.PageWidth - [double]$Document.PageSetup.LeftMargin - [double]$Document.PageSetup.RightMargin
  $shape.Left = [single](($contentWidth - $lineLength) / 2)
  $shape.Top = [single]$lineOffset
  $shape.LockAspectRatio = 0
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

  $targetGap = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 1.8))
  for ($attempt = 0; $attempt -lt 4; $attempt++) {
    $Document.Repaginate()
    $redLineInfo = Get-DocumentNoRedLineInfo $Document $Rules $DocNoParagraph
    if (-not [bool]$redLineInfo.Exists) { return }
    $actualGap = [double]$redLineInfo.OffsetPt
    $delta = $targetGap - $actualGap
    if ([Math]::Abs($delta) -le 0.5) { break }

    if ([string]$redLineInfo.Source -eq "shape") {
      $redLineInfo.Shape.Top = [single]([Math]::Max(0, [double]$redLineInfo.Shape.Top + $delta))
    } else {
      try { $DocNoParagraph.Range.ParagraphFormat.Borders.DistanceFromBottom = [single]([Math]::Max(0, $actualGap + $delta)) } catch {}
    }
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
    Clear-ParagraphSpacingUnits $spacerParagraph
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
  $docNoLineHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  $docNoBlankLines = [double](Get-RuleValue $Rules "redHeader.docNoBlankLines" 2)
  $targetDocNoGap = $docNoLineHeight * $docNoBlankLines
  $docNoSpacerParagraph = $null
  if ([int]$docNoIndex - [int]$headerIndex -gt 1) {
    for ($i = [int]$headerIndex + 1; $i -lt [int]$docNoIndex; $i++) {
      $candidate = $Document.Paragraphs.Item($i)
      if (Test-BlankParagraph $candidate) {
        $docNoSpacerParagraph = $candidate
        $docNoSpacerParagraph.Range.Font.Size = 1
        $docNoSpacerParagraph.Range.Font.Color = Convert-HexToWordColor "FFFFFF"
        Clear-ParagraphSpacingUnits $docNoSpacerParagraph
        $docNoSpacerParagraph.Range.ParagraphFormat.LineSpacingRule = 4
        $docNoSpacerParagraph.Range.ParagraphFormat.LineSpacing = [single]$docNoLineHeight
        $docNoSpacerParagraph.Range.ParagraphFormat.SpaceBefore = 0
        $docNoSpacerParagraph.Range.ParagraphFormat.SpaceAfter = 0
        $docNoSpacerParagraph.Range.ParagraphFormat.FirstLineIndent = 0
        $docNoSpacerParagraph.Range.ParagraphFormat.LeftIndent = 0
        $docNoSpacerParagraph.Range.ParagraphFormat.RightIndent = 0
        break
      }
    }
  }

  for ($attempt = 0; $attempt -lt 5; $attempt++) {
    $Document.Repaginate()
    $headerVisibleBottom = [double]$headerParagraph.Range.Information(6) + $headerVisibleTopInset + $headerSize
    $actualDocNoGap = [double]$docNoParagraph.Range.Information(6) - $headerVisibleBottom
    $delta = $targetDocNoGap - $actualDocNoGap
    if ([Math]::Abs($delta) -le 1.0) { break }

    if ($null -ne $docNoSpacerParagraph) {
      $currentLineSpacing = [double]$docNoSpacerParagraph.Range.ParagraphFormat.LineSpacing
      $nextLineSpacing = [Math]::Max(1, $currentLineSpacing + $delta)
      if ([Math]::Abs($nextLineSpacing - $currentLineSpacing) -le 0.1) { break }
      $docNoSpacerParagraph.Range.ParagraphFormat.LineSpacing = [single]$nextLineSpacing
    } else {
      $currentSpaceAfter = [double]$headerParagraph.Range.ParagraphFormat.SpaceAfter
      $nextSpaceAfter = [Math]::Max(0, $currentSpaceAfter + $delta)
      if ([Math]::Abs($nextSpaceAfter - $currentSpaceAfter) -le 0.1) { break }
      $headerParagraph.Range.ParagraphFormat.SpaceAfter = [single]$nextSpaceAfter
    }
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

  $insertRange = $Document.Range(0, 0)
  $insertRange.InsertBefore("`r$headerText`r`r$docNo`t$signerLabel$signer`r")

  $spacerParagraph = $Document.Paragraphs.Item(1)
  if ($headerSpaceBefore -gt 0) {
    $spacerParagraph.Range.Font.Size = 1
    $spacerParagraph.Range.Font.Hidden = 0
    $spacerParagraph.Range.Font.Color = Convert-HexToWordColor "FFFFFF"
    Clear-ParagraphSpacingUnits $spacerParagraph
    $spacerParagraph.Range.ParagraphFormat.LineSpacingRule = 4
    $spacerParagraph.Range.ParagraphFormat.LineSpacing = [single]$headerSpaceBefore
    $spacerParagraph.Range.ParagraphFormat.SpaceBefore = 0
    $spacerParagraph.Range.ParagraphFormat.SpaceAfter = 0
    $spacerParagraph.Range.ParagraphFormat.FirstLineIndent = 0
  }

  $p1 = $Document.Paragraphs.Item(2)
  Set-RangeFont $p1.Range $headerFont $headerSize "" $headerColor $headerScale $headerSpacing
  Clear-ParagraphSpacingUnits $p1
  $p1.Range.ParagraphFormat.Alignment = 1
  $p1.Range.ParagraphFormat.FirstLineIndent = 0
  $p1.Range.ParagraphFormat.LineSpacingRule = 0
  $p1.Range.ParagraphFormat.SpaceBefore = 0
  $p1.Range.ParagraphFormat.SpaceAfter = $headerSpaceAfter

  $docNoIndexAfterInsert = Find-DocumentNoParagraphIndex $Document
  if (($null -ne $docNoIndexAfterInsert) -and ([int]$docNoIndexAfterInsert -gt 2)) {
    $docNoSpacerParagraph = $Document.Paragraphs.Item([int]$docNoIndexAfterInsert - 1)
    if (Test-BlankParagraph $docNoSpacerParagraph) {
      $docNoSpacerParagraph.Range.Font.Size = 1
      $docNoSpacerParagraph.Range.Font.Color = Convert-HexToWordColor "FFFFFF"
      Clear-ParagraphSpacingUnits $docNoSpacerParagraph
      $docNoSpacerParagraph.Range.ParagraphFormat.LineSpacingRule = 4
      $docNoSpacerParagraph.Range.ParagraphFormat.LineSpacing = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
      $docNoSpacerParagraph.Range.ParagraphFormat.SpaceBefore = 0
      $docNoSpacerParagraph.Range.ParagraphFormat.SpaceAfter = 0
      $docNoSpacerParagraph.Range.ParagraphFormat.FirstLineIndent = 0
      $docNoSpacerParagraph.Range.ParagraphFormat.LeftIndent = 0
      $docNoSpacerParagraph.Range.ParagraphFormat.RightIndent = 0
    }
  }

  $docNoIndexAfterInsert = Find-DocumentNoParagraphIndex $Document
  if ($null -eq $docNoIndexAfterInsert) { $docNoIndexAfterInsert = 4 }
  $p2 = $Document.Paragraphs.Item([int]$docNoIndexAfterInsert)
  $docNoFont = Get-RuleValue $Rules "documentNo.font" "仿宋_GB2312"
  $docNoSize = [double](Get-RuleValue $Rules "documentNo.sizePt" 16)
  $docNoSpaceAfter = [double](Get-RuleValue $Rules "documentNo.spaceAfterPt" 40)
  Set-RangeFont $p2.Range $docNoFont $docNoSize "Times New Roman" 0
  Clear-ParagraphSpacingUnits $p2
  $p2.Range.ParagraphFormat.Alignment = 0
  $p2.Range.ParagraphFormat.FirstLineIndent = [single]($docNoSize * [Math]::Max(0, $docNoLeftChars))
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
  $redLineInfo = Get-DocumentNoRedLineInfo $Document $Rules $docNoParagraph
  if (-not [bool]$redLineInfo.Exists) { return }

  $titleTopBlankLines = [double](Get-RuleValue $Rules "title.topBlankLines" 2)
  $blankLineHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  $targetGap = if ([string]$redLineInfo.Source -eq "shape") {
    [Math]::Max(0, ($blankLineHeight * $titleTopBlankLines) - [double]$redLineInfo.OffsetPt - [double]$redLineInfo.LineWidthPt)
  } else {
    $blankLineHeight * $titleTopBlankLines
  }

  for ($attempt = 0; $attempt -lt 3; $attempt++) {
    $Document.Repaginate()
    $redLineInfo = Get-DocumentNoRedLineInfo $Document $Rules $docNoParagraph
    $redLineBottom = [double]$redLineInfo.BottomPt
    $titleTop = [double]$titleParagraph.Range.Information(6)
    $actualGap = $titleTop - $redLineBottom
    $delta = $targetGap - $actualGap
    if ([Math]::Abs($delta) -le 0.5) { break }

    $currentSpaceBefore = [double]$titleParagraph.Range.ParagraphFormat.SpaceBefore
    $titleParagraph.Range.ParagraphFormat.SpaceBefore = [single]([Math]::Max(0, $currentSpaceBefore + $delta))
  }
}

function Remove-ExistingRedHead {
  param($Document, $Rules)

  $headerText = ([string](Get-RuleValue $Rules "redHeader.text" "")).Trim() -replace "\s", ""
  $limit = [Math]::Min($Document.Paragraphs.Count, 8)
  $matchedIndexes = @()

  for ($i = 1; $i -le $limit; $i++) {
    $text = Get-ParagraphText $Document.Paragraphs.Item($i)
    $compact = $text -replace "\s", ""
    if ([string]::IsNullOrWhiteSpace($compact)) { continue }
    $isHeader = ($headerText -ne "") -and $compact.Contains($headerText)
    $isDocNo = $compact -match "〔\d{4}〕\d+号"
    if ($isHeader -or $isDocNo) { $matchedIndexes += $i }
  }
  if ($matchedIndexes.Count -eq 0) { return 0 }

  # 版头块 = 首段到最后一个命中段；其中的命中段与空白段一并删除
  $maxMatched = [int]($matchedIndexes | Measure-Object -Maximum).Maximum
  $deleteIndexes = @()
  for ($i = 1; $i -le $maxMatched; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    if (($matchedIndexes -contains $i) -or (Test-BlankParagraph $paragraph)) {
      $deleteIndexes += $i
    }
  }

  # 删除锚定在版头块内的形状（旧版头红线、装饰图形）
  for ($s = $Document.Shapes.Count; $s -ge 1; $s--) {
    $shape = $Document.Shapes.Item($s)
    $anchorStart = $null
    try { $anchorStart = [int]$shape.Anchor.Start } catch { continue }
    foreach ($i in $deleteIndexes) {
      $range = $Document.Paragraphs.Item($i).Range
      if (($anchorStart -ge [int]$range.Start) -and ($anchorStart -lt [int]$range.End)) {
        $shape.Delete()
        break
      }
    }
  }

  for ($i = $deleteIndexes.Count - 1; $i -ge 0; $i--) {
    $Document.Paragraphs.Item([int]$deleteIndexes[$i]).Range.Delete() | Out-Null
  }
  return $deleteIndexes.Count
}

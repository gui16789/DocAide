# 套红格式校验
# 本模块由 Redhead.Core.ps1 拆分而来，函数定义在 dot-source 后与其它模块共享作用域。

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
  $expectedDocNoIndent = $docNoSize * [Math]::Max(0, $docNoLeftChars)
  $docNoPrefixOk = $docNoLineRaw.StartsWith("$docNoPrefix$docNo")
  $docNoIndentOk = Test-Near ([double]$docNoParagraph.Range.ParagraphFormat.FirstLineIndent) $expectedDocNoIndent 1.0
  $checks += New-Check "发文字号左空一字" ($docNoPrefixOk -or $docNoIndentOk) "配置左空 $docNoLeftChars 字，首行缩进 $([Math]::Round($docNoParagraph.Range.ParagraphFormat.FirstLineIndent, 2))pt"
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
  $docNoLineHeight = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  $expectedDocNoGap = $docNoLineHeight * $docNoBlankLines
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
  $signerSuffix = Repeat-Text "　" ([Math]::Max(0, $signerRightChars))
  $signerSuffixOk = ($signerSuffix.Length -gt 0) -and $docNoLineRaw.EndsWith($signerSuffix)
  $checks += New-Check "签发人右空一字" ($signerTabOk -or $signerSuffixOk) "配置右空 $signerRightChars 字，右对齐制表位 $([Math]::Round($expectedSignerTab, 1))pt"
  $signerFont = Get-RuleValue $Rules "documentNo.signerFont" "楷体_GB2312"
  $signerFontOk = $false
  $signerStart = $docNoLineRaw.IndexOf($signer)
  if ($signerStart -ge 0) {
    $signerRange = $Document.Range($docNoParagraph.Range.Start + $signerStart, $docNoParagraph.Range.Start + $signerStart + $signer.Length)
    $signerFontOk = ((Test-FontNameMatches $signerRange.Font.NameFarEast $signerFont) -or (Test-FontNameMatches $signerRange.Font.Name $signerFont)) -and (Test-Near ([double]$signerRange.Font.Size) $docNoSize 0.5)
  }
  $checks += New-Check "签发人字体字号符合规范" $signerFontOk "签发人三字 $docNoFont，姓名 $signerFont，字号 $docNoSize pt"
  $redLineInfo = Get-DocumentNoRedLineInfo $Document $Rules $docNoParagraph
  $redLineExists = [bool]$redLineInfo.Exists
  $checks += New-Check "红线已设置" $redLineExists
  $expectedRedLineColorHex = ((Get-RuleValue $Rules "documentNo.redLineColor" "FF0000") -replace "#", "").ToUpperInvariant()
  $expectedRedLineWidthPt = [double](Get-RuleValue $Rules "documentNo.redLineWidthPt" 2.25)
  $actualRedLineColorHex = if ($redLineExists) { [string]$redLineInfo.ColorHex } else { "" }
  $actualRedLineWidthPt = if ($redLineExists) { [double]$redLineInfo.LineWidthPt } else { 0.0 }
  $redLineFormatOk = $redLineExists -and
    ($actualRedLineColorHex -eq $expectedRedLineColorHex) -and
    (Test-Near $actualRedLineWidthPt $expectedRedLineWidthPt 0.05)
  $checks += New-Check "版头红线颜色线宽符合规则" $redLineFormatOk "颜色 #$expectedRedLineColorHex，线宽 $expectedRedLineWidthPt pt"
  $expectedRedLineOffset = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineOffsetMm" 1.8))
  $actualRedLineOffset = if ($redLineExists) { [double]$redLineInfo.OffsetPt } else { 0.0 }
  $redLineOffsetOk = $redLineExists -and (Test-Near $actualRedLineOffset $expectedRedLineOffset 1.0)
  $checks += New-Check "版头红线浮动位置符合标准模板" $redLineOffsetOk "锚点下移 $([Math]::Round($actualRedLineOffset, 2))pt/$([Math]::Round($actualRedLineOffset * 25.4 / 72, 2))mm，目标 $([Math]::Round($expectedRedLineOffset, 2))pt/$([Math]::Round($expectedRedLineOffset * 25.4 / 72, 2))mm"
  $expectedRedLineLength = Convert-MmToPt ([double](Get-RuleValue $Rules "documentNo.redLineLengthMm" 156))
  $actualRedLineLength = if ($redLineExists) { [double]$redLineInfo.LengthPt } else { 0.0 }
  $actualRedLineCenter = if ($redLineExists) { [double]$redLineInfo.CenterPt } else { 0.0 }
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
  $titleTopBlankLines = [double](Get-RuleValue $Rules "title.topBlankLines" 2)
  $blankLineHeightForTitle = [double](Get-RuleValue $Rules "documentNo.redLineLineBoxHeightPt" 28)
  $expectedTitleTopGap = if ([string]$redLineInfo.Source -eq "shape") {
    [Math]::Max(0, ($blankLineHeightForTitle * $titleTopBlankLines) - [double]$redLineInfo.OffsetPt - [double]$redLineInfo.LineWidthPt)
  } else {
    $blankLineHeightForTitle * $titleTopBlankLines
  }
  if (($null -ne $titleParagraph) -and $redLineExists) {
    $redLineBottom = [double]$redLineInfo.BottomPt
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
    $actualTitleLines = @($titleItems | ForEach-Object { Normalize-TitleText $_.Text } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $normalizedTitle = Normalize-TitleText (($titleItems | ForEach-Object { $_.Text }) -join "")
    $maxLineChars = [int](Get-RuleValue $Rules "title.maxLineChars" 22)
    $lineCountOk = ($actualTitleLines.Count -ge 1) -and ($actualTitleLines.Count -le 4)
    $lineLengthOk = $true
    foreach ($line in $actualTitleLines) {
      if ((Normalize-TitleText $line).Length -gt $maxLineChars) {
        $lineLengthOk = $false
        break
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
    $titleBlockOk = $lineCountOk -and $lineLengthOk -and $breaksSemanticOk
    $checks += New-Check "标题已按词义主动分行" $titleBlockOk "标题 $($actualTitleLines.Count) 行，断点 $($breakPositions -join ',')，每行最多 $maxLineChars 字"
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
    $titleGapOk = $true
    if ($titleItems.Count -gt 0) {
      $lastTitleItem = $titleItems[$titleItems.Count - 1]
      $titleGapOk = Test-Near ([double]$lastTitleItem.Paragraph.Range.ParagraphFormat.SpaceAfter) $titleSpaceAfter 1.0
      $nextIndex = [int]$lastTitleItem.Index + 1
      if (($nextIndex -lt [int]$mainRecipientItem.Index) -and (Test-BlankParagraph $Document.Paragraphs.Item($nextIndex))) {
        $titleGapOk = $true
      }
    }
    $mainRecipientAlignment = [int]$mainRecipientParagraph.Range.ParagraphFormat.Alignment
    $mainRecipientLayoutOk = (($mainRecipientAlignment -eq 0) -or ($mainRecipientAlignment -eq 3)) -and
      (Test-Near ([double]$mainRecipientParagraph.Range.ParagraphFormat.LeftIndent) 0 0.5) -and
      (Test-Near ([double]$mainRecipientParagraph.Range.ParagraphFormat.FirstLineIndent) 0 0.5) -and
      $titleGapOk
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
    $pageNumberDistanceMm = [double](Get-RuleValue $Rules "pageNumber.distanceBelowContentMm" 17.5)
    $expectedFooterDistance = ([double]$bodySection.PageSetup.BottomMargin) - (Convert-MmToPt $pageNumberDistanceMm)
    $checks += New-Check "页码位于版心下边缘下方规则距离" (Test-Near ([double]$bodySection.PageSetup.FooterDistance) $expectedFooterDistance 1.0) "规则 $pageNumberDistanceMm mm，FooterDistance $([Math]::Round($bodySection.PageSetup.FooterDistance, 2))pt"
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

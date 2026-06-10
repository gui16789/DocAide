# 版记表格生成与定位
# 本模块由 Redhead.Core.ps1 拆分而来，函数定义在 dot-source 后与其它模块共享作用域。

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

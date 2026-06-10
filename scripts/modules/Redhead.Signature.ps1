# 落款署名、成文日期和公章
# 本模块由 Redhead.Core.ps1 拆分而来，函数定义在 dot-source 后与其它模块共享作用域。

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

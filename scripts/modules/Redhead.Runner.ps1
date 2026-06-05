Write-JobStatus "initializing" "读取规则和处理参数"
$rules = Get-Content -LiteralPath $RulesPath -Raw -Encoding UTF8 | ConvertFrom-Json
$meta = Get-Content -LiteralPath $MetaPath -Raw -Encoding UTF8 | ConvertFrom-Json

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$baseName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
$outputDocx = Join-Path $OutputDir ($baseName + "-套红.docx")
$outputPdf = Join-Path $OutputDir ($baseName + "-套红.pdf")
Write-JobStatus "output-prepared" "输出目录已准备" @{ outputDocx = $outputDocx; outputPdf = $outputPdf }

$word = $null
$doc = $null

try {
  Write-JobStatus "word-starting" "正在启动 Microsoft Word"
  $wordStartedAt = Get-Date
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $script:CurrentWordPid = Get-WordProcessId $word
  if ($null -eq $script:CurrentWordPid) {
    $script:CurrentWordPid = Get-RecentWordProcessId $wordStartedAt
  }
  Write-ProcessPidFile $PidPath $script:CurrentWordPid
  Write-JobStatus "word-ready" "Microsoft Word 已启动" @{ wordPid = $script:CurrentWordPid }

  $inputExtension = [IO.Path]::GetExtension($InputPath).ToLowerInvariant()
  $convertToDocx = [bool](Get-RuleValue $rules "cleanup.convertDocToDocx" $true)
  if (($inputExtension -eq ".doc") -or (($inputExtension -ne ".docx") -and $convertToDocx)) {
    Write-JobStatus "input-converting" "正在打开源文件并转换为 DOCX" @{ inputPath = $InputPath }
    $doc = $word.Documents.Open($InputPath, $false, $false)
    $doc.Activate()
    $doc.SaveAs2($outputDocx, 12)
    $doc.Close($false)
    $doc = $null
    Write-JobStatus "input-converted" "源文件已转换为 DOCX" @{ outputDocx = $outputDocx }
  } else {
    Write-JobStatus "input-copying" "源文件已是 DOCX，按规则直接复制"
    Copy-Item -LiteralPath $InputPath -Destination $outputDocx -Force
  }

  Write-JobStatus "document-opening" "正在打开待套红 DOCX"
  $doc = $word.Documents.Open($outputDocx, $false, $false)
  $doc.Activate()
  Write-JobStatus "document-opened" "待套红 DOCX 已打开"

  if ([bool](Get-RuleValue $rules "cleanup.removeExistingHeadersFooters" $true)) {
    Write-JobStatus "cleanup-headers-footers" "正在清理源文页眉页脚"
    Clear-HeadersFooters $doc
  }
  Write-JobStatus "page-setup" "正在应用页面和版心设置"
  Apply-PageSetup $doc $rules

  if ([bool](Get-RuleValue $rules "cleanup.fixDoubleAnmeng" $true)) {
    Write-JobStatus "text-cleanup" "正在执行源文文本清理"
    $find = $doc.Content.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()
    $null = $find.Execute("安安盟", $false, $true, $false, $false, $false, $true, 1, $false, "安盟", 2)
  }

  $titleOverride = Get-RuleValue $meta "titleOverride" ""
  if ([bool](Get-RuleValue $meta "replaceTitle" $false)) {
    Write-JobStatus "title-replacing" "正在按处理参数覆盖标题"
    Replace-InitialTitle $doc $titleOverride
  }

  Write-JobStatus "trim-before-redhead" "正在清理源文尾部空白内容"
  $trimInfo = Trim-TrailingBlankContent $doc $rules
  Write-JobStatus "redhead-inserting" "正在插入发文机关标志、发文字号和签发人"
  Insert-RedHeader $doc $rules $meta
  Write-JobStatus "title-normalizing" "正在规范标题分行"
  Normalize-TitleBlock $doc $rules
  Write-JobStatus "body-styling" "正在应用标题、主送机关和正文样式"
  Apply-BodyAndTitleStyle $doc $rules
  Write-JobStatus "redline-adding" "正在设置版头红色分隔线"
  $docNoIndexForRedLine = Find-DocumentNoParagraphIndex $doc
  if ($null -ne $docNoIndexForRedLine) {
    Add-DocumentNoRedLine $doc $doc.Paragraphs.Item([int]$docNoIndexForRedLine) $rules
  }
  Write-JobStatus "title-positioning" "正在校正版头红线与标题距离"
  Adjust-TitlePositionAfterRedLine $doc $rules
  Write-JobStatus "signature-applying" "正在处理落款、成文日期和印章"
  Apply-Signature $doc $rules $meta
  Write-JobStatus "attachment-styling" "正在规范附件说明"
  Apply-AttachmentExplanationStyle $doc $rules
  Write-JobStatus "trim-after-signature" "正在清理落款后的尾部空白内容"
  $trimInfoAfterSignature = Trim-TrailingBlankContent $doc $rules
  $trimInfo["BlankParagraphsDeleted"] = [int]$trimInfo.BlankParagraphsDeleted + [int]$trimInfoAfterSignature.BlankParagraphsDeleted
  $trimInfo["TrailingBreaksDeleted"] = [int]$trimInfo.TrailingBreaksDeleted + [int]$trimInfoAfterSignature.TrailingBreaksDeleted
  $trimInfo["PagesAfter"] = [int]$trimInfoAfterSignature.PagesAfter
  $trimInfo["PagesRemoved"] = [Math]::Max(0, [int]$trimInfo.PagesBefore - [int]$trimInfo.PagesAfter)

  Write-JobStatus "imprint-adding" "正在生成和定位版记"
  $imprintInfo = Add-Imprint $word $doc $rules
  Write-JobStatus "page-numbering" "正在设置页码"
  Apply-BodyPageNumbers $doc $imprintInfo $rules
  Write-JobStatus "fields-updating" "正在更新 Word 字段并重新分页"
  $doc.Fields.Update() | Out-Null
  $doc.Repaginate()

  $totalPages = [int]$doc.ComputeStatistics(2)
  $imprintPage = if ([bool](Get-RuleValue $rules "imprint.enabled" $true)) { $totalPages } else { $null }
  Write-JobStatus "document-saving" "正在保存套红 DOCX" @{ totalPages = $totalPages; imprintPage = $imprintPage }
  $doc.Save()
  Write-JobStatus "document-reopening" "正在重启 Word 并准备稳定导出和校验"
  $doc.Close($true) | Out-Null
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
  $doc = $null
  $word.Quit() | Out-Null
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  $word = $null

  Write-JobStatus "pdf-exporting" "正在通过独立 Word 进程导出 PDF 预览" @{ outputPdf = $outputPdf }
  $pdfExportScriptPath = Join-Path $OutputDir "export-pdf.ps1"
  $pdfExportOut = Join-Path $OutputDir "pdf-export.out.log"
  $pdfExportErr = Join-Path $OutputDir "pdf-export.err.log"
  $escapedDocx = $outputDocx.Replace("'", "''")
  $escapedPdf = $outputPdf.Replace("'", "''")
  $pdfExportScript = @"
`$ErrorActionPreference = "Stop"
`$word = `$null
`$doc = `$null
try {
  `$word = New-Object -ComObject Word.Application
  `$word.Visible = `$false
  `$word.DisplayAlerts = 0
  `$doc = `$word.Documents.Open('$escapedDocx', `$false, `$true)
  `$doc.ExportAsFixedFormat('$escapedPdf', 17)
} finally {
  if (`$null -ne `$doc) {
    try { `$doc.Close(`$false) | Out-Null } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject(`$doc)
  }
  if (`$null -ne `$word) {
    try { `$word.Quit() | Out-Null } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject(`$word)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
"@
  $pdfExportScript | Set-Content -LiteralPath $pdfExportScriptPath -Encoding UTF8
  $pdfProcess = Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $pdfExportScriptPath) -PassThru -WindowStyle Hidden -RedirectStandardOutput $pdfExportOut -RedirectStandardError $pdfExportErr
  if (-not $pdfProcess.WaitForExit(90000)) {
    try { Stop-Process -Id $pdfProcess.Id -Force -ErrorAction SilentlyContinue } catch {}
    Write-JobStatus "pdf-export-timeout" "PDF 导出超时，继续执行 DOCX 校验" @{ outputPdf = $outputPdf }
  } elseif ($pdfProcess.ExitCode -ne 0) {
    Write-JobStatus "pdf-export-failed" "PDF 导出失败，继续执行 DOCX 校验" @{ exitCode = $pdfProcess.ExitCode; outputPdf = $outputPdf }
  }

  Write-JobStatus "document-reopening" "正在重新打开已保存 DOCX 以执行格式校验"
  $wordStartedAt = Get-Date
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $script:CurrentWordPid = Get-WordProcessId $word
  if ($null -eq $script:CurrentWordPid) {
    $script:CurrentWordPid = Get-RecentWordProcessId $wordStartedAt
  }
  Write-ProcessPidFile $PidPath $script:CurrentWordPid
  $doc = $word.Documents.Open($outputDocx, $false, $true)
  $doc.Repaginate()
  $totalPages = [int]$doc.ComputeStatistics(2)
  $imprintPage = if ([bool](Get-RuleValue $rules "imprint.enabled" $true)) { $totalPages } else { $null }

  Write-JobStatus "validating" "正在执行格式校验"
  $checks = Build-ValidationChecks $doc $rules $meta $imprintInfo $trimInfo $totalPages $imprintPage $outputPdf
  $hasFailures = @($checks | Where-Object { $_.Status -eq "fail" }).Count -gt 0

  Write-JobStatus "result-writing" "正在写入处理结果"
  $result = @{
    ok = $true
    validationPassed = -not $hasFailures
    outputDocx = $outputDocx
    outputPdf = $outputPdf
    pages = @{
      body = [int]$imprintInfo.BodyPages
      total = $totalPages
      blankInserted = [bool]$imprintInfo.BlankInserted
      imprint = $imprintPage
      trailingBlankPagesRemoved = [int]$trimInfo.PagesRemoved
    }
    checks = $checks
    warnings = if ($hasFailures) { @("存在未通过的格式校验项，请查看检查结果。") } else { @() }
  }
  $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ResultPath -Encoding UTF8
  $script:JobCompleted = $true
  Write-JobStatus "completed" "处理完成" @{ totalPages = $totalPages; validationPassed = (-not $hasFailures) }
} catch {
  $script:JobFailed = $true
  $script:JobFailureDetail = $_.Exception.Message
  $result = @{
    ok = $false
    error = $_.Exception.Message
    outputDocx = $outputDocx
    outputPdf = $outputPdf
    checks = @()
    warnings = @("处理失败，请确认 Microsoft Word 可用且字体已安装。")
  }
  Write-JobStatus "failed" $_.Exception.Message @{ outputDocx = $outputDocx; outputPdf = $outputPdf }
  $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ResultPath -Encoding UTF8
  throw
} finally {
  Write-JobStatus "word-cleanup" "正在释放 Word COM 资源"
  if ($null -ne $doc) {
    try { $doc.Close($false) | Out-Null } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
  }
  if ($null -ne $word) {
    try { $word.Quit() | Out-Null } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  if ($script:JobCompleted) {
    Write-JobStatus "completed" "处理完成，Word COM 资源已释放"
  } elseif ($script:JobFailed) {
    Write-JobStatus "failed" "处理失败，Word COM 资源已释放：$script:JobFailureDetail"
  } else {
    Write-JobStatus "finished" "Word COM 资源已释放"
  }
}

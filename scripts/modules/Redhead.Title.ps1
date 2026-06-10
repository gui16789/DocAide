# 公文标题规范分行
# 本模块由 Redhead.Core.ps1 拆分而来，函数定义在 dot-source 后与其它模块共享作用域。

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
  param([string]$Text, [int]$Position, [int]$IgnoreLongerThan = 0)
  if ($Position -le 0 -or $Position -ge $Text.Length) { return $false }
  foreach ($term in (Get-TitleProtectedTerms $Text)) {
    # 禁拆词本身超过单行上限时必须在词内断行，不再视为违规
    if ($IgnoreLongerThan -gt 0 -and $term.Length -gt $IgnoreLongerThan) { continue }
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
      $last = [int]$state["Breaks"][-1]
      foreach ($point in $points) {
        if ($point -le $last) { continue }
        if ($line -lt $lineCount -and $point -ge $title.Length) { continue }
        if ($line -eq $lineCount -and $point -ne $title.Length) { continue }

        $segmentLength = $point - $last
        if ($segmentLength -le 0) { continue }
        $score = [double]$state["Score"] + [Math]::Pow($segmentLength - $target, 2)
        if ($segmentLength -lt 6 -and $line -lt $lineCount) { $score += 100 }
        if ($segmentLength -gt $MaxLineChars) { $score += 5000 + 200 * ($segmentLength - $MaxLineChars) }
        if (-not $semanticSet.Contains($point) -and $point -ne $title.Length) { $score += 140 }
        if (Test-TitleBreakInsideProtectedTerm $title $point $MaxLineChars) { $score += 4000 }
        if (Test-TitleBreakHasBadEdge $title $point) { $score += 2500 }
        $segment = $title.Substring($last, $segmentLength)
        if ($segment.EndsWith("的") -or $segment.EndsWith("对") -or $segment.EndsWith("关于")) { $score += 50 }
        if ($line -gt 1 -and ("的对和与及、，）)".Contains([string]$segment[0]))) { $score += 50 }

        $breaks = @($state["Breaks"] + $point)
        $nextStates += @{ Score = $score; Breaks = $breaks }
      }
    }
    $states = @($nextStates | Sort-Object { [double]$_["Score"] } | Select-Object -First 60)
  }

  if ($states.Count -eq 0) { return @($title) }
  $best = $states[0]["Breaks"]
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

  $formattedTitle = ($titleLines -join "`r")
  $replaceStart = $Document.Paragraphs.Item($titleStartIndex).Range.Start
  $replaceEnd = $Document.Paragraphs.Item($bodyStartIndex).Range.Start
  $range = $Document.Range($replaceStart, $replaceEnd)
  $range.Text = "$formattedTitle`r`r"
}

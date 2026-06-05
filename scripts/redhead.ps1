param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputDir,
  [Parameter(Mandatory = $true)][string]$RulesPath,
  [Parameter(Mandatory = $true)][string]$MetaPath,
  [Parameter(Mandatory = $true)][string]$ResultPath,
  [string]$PidPath = "",
  [string]$StatusPath = ""
)

$moduleRoot = Join-Path $PSScriptRoot "modules"
. (Join-Path $moduleRoot "Redhead.Core.ps1")
. (Join-Path $moduleRoot "Redhead.Runner.ps1")

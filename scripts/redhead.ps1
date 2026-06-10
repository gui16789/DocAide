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
. (Join-Path $moduleRoot "Redhead.Title.ps1")
. (Join-Path $moduleRoot "Redhead.RedHeader.ps1")
. (Join-Path $moduleRoot "Redhead.Signature.ps1")
. (Join-Path $moduleRoot "Redhead.Imprint.ps1")
. (Join-Path $moduleRoot "Redhead.Validation.ps1")
. (Join-Path $moduleRoot "Redhead.Runner.ps1")

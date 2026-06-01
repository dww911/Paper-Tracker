# If you are in the web/ folder, delegate to project root
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
& (Join-Path $Root "start-web.ps1")

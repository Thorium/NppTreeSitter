# Deploy TreeSitterLexer to Notepad++ installation
$ErrorActionPreference = "Stop"

$nppDir = "C:\Program Files\Notepad++"
$pluginDir = "$nppDir\plugins\TreeSitterLexer"
$grammarDir = "$pluginDir\TreeSitterGrammars"
$configDir = "$nppDir\plugins\Config"

# Source paths (relative to this script's directory)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllSrc = Join-Path $ScriptDir "build\x64\Release\TreeSitterLexer.dll"
$grammarSrc = Join-Path $ScriptDir "build\x64\Release\TreeSitterGrammars"
$xmlSrc = Join-Path $ScriptDir "config\TreeSitterLexer.xml"

Write-Host "Creating directories..."
New-Item -ItemType Directory -Path $grammarDir -Force | Out-Null

Write-Host "Copying TreeSitterLexer.dll..."
Copy-Item $dllSrc "$pluginDir\TreeSitterLexer.dll" -Force

Write-Host "Copying grammar DLLs and .scm files..."
Copy-Item "$grammarSrc\*.dll" $grammarDir -Force
Copy-Item "$grammarSrc\*.scm" $grammarDir -Force

Write-Host "Copying config XML..."
Copy-Item $xmlSrc "$configDir\TreeSitterLexer.xml" -Force

Write-Host ""
Write-Host "=== Deployment complete ==="
Write-Host "Plugin:   $pluginDir\TreeSitterLexer.dll"
Get-ChildItem "$grammarDir\*.dll" | ForEach-Object { Write-Host "Grammar:  $($_.FullName)" }
Get-ChildItem "$grammarDir\*.scm" | ForEach-Object { Write-Host "Query:    $($_.FullName)" }
Write-Host "Config:   $configDir\TreeSitterLexer.xml"

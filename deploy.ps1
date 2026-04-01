# Deploy TreeSitterLexer to Notepad++ installation
param(
    [ValidateSet("x64","Win32","x86","ARM64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"

$BuildPlatform = switch ($Platform) {
    "x64"   { "x64" }
    "Win32" { "Win32" }
    "x86"   { "Win32" }
    "ARM64" { "ARM64" }
}

$nppDir = "C:\Program Files\Notepad++"
$pluginDir = "$nppDir\plugins\TreeSitterLexer"
$grammarDir = "$pluginDir\TreeSitterGrammars"
$configDir = "$nppDir\plugins\Config"

# Source paths (relative to this script's directory)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllSrc = Join-Path $ScriptDir "build\$BuildPlatform\Release\TreeSitterLexer.dll"
$grammarSrc = Join-Path $ScriptDir "build\$BuildPlatform\Release\TreeSitterGrammars"
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
Write-Host "Platform: $Platform"
Write-Host "Plugin:   $pluginDir\TreeSitterLexer.dll"
Get-ChildItem "$grammarDir\*.dll" | ForEach-Object { Write-Host "Grammar:  $($_.FullName)" }
Get-ChildItem "$grammarDir\*.scm" | ForEach-Object { Write-Host "Query:    $($_.FullName)" }
Write-Host "Config:   $configDir\TreeSitterLexer.xml"

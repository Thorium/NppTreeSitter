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
$bundledGrammarDir = "$pluginDir\BundledTreeSitterGrammars"
$configDir = "$nppDir\plugins\Config"

# Source paths (relative to this script's directory)
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllSrc = Join-Path $ScriptDir "build\$BuildPlatform\Release\TreeSitterLexer.dll"
$grammarSrc = Join-Path $ScriptDir "build\$BuildPlatform\Release\TreeSitterGrammars"
$xmlSrc = Join-Path $ScriptDir "config\TreeSitterLexer.xml"

$OptionalBundledLanguages = @(
    "r",
    "lua",
    "yaml",
    "toml",
    "markdown",
    "cmake",
    "fortran",
    "pascal",
    "matlab"
)

function Copy-GrammarLanguageFiles {
    param(
        [string]$SourceDir,
        [string]$DestinationDir,
        [string]$Language
    )

    Copy-Item (Join-Path $SourceDir "tree-sitter-$Language.dll") $DestinationDir -Force
    Get-ChildItem -Path $SourceDir -Filter "$Language-*.scm" -File -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName $DestinationDir -Force
    }
}

Write-Host "Creating directories..."
New-Item -ItemType Directory -Path $grammarDir -Force | Out-Null
New-Item -ItemType Directory -Path $bundledGrammarDir -Force | Out-Null

Write-Host "Copying TreeSitterLexer.dll..."
Copy-Item $dllSrc "$pluginDir\TreeSitterLexer.dll" -Force

Write-Host "Copying grammar DLLs and .scm files..."
Get-ChildItem -Path $grammarSrc -Filter "tree-sitter-*.dll" -File | ForEach-Object {
    $language = $_.BaseName.Substring("tree-sitter-".Length)
    if ($OptionalBundledLanguages -notcontains $language) {
        Copy-GrammarLanguageFiles -SourceDir $grammarSrc -DestinationDir $grammarDir -Language $language
    }
}

Write-Host "Copying bundled grammar install payload..."
foreach ($language in $OptionalBundledLanguages) {
    $dllPath = Join-Path $grammarSrc "tree-sitter-$language.dll"
    if (Test-Path $dllPath) {
        Copy-GrammarLanguageFiles -SourceDir $grammarSrc -DestinationDir $bundledGrammarDir -Language $language
    }
}

Write-Host "Copying config XML..."
New-Item -ItemType Directory -Path $configDir -Force | Out-Null
Copy-Item $xmlSrc "$configDir\TreeSitterLexer.xml" -Force
Copy-Item $xmlSrc "$pluginDir\TreeSitterLexer.xml" -Force

Write-Host ""
Write-Host "=== Deployment complete ==="
Write-Host "Platform: $Platform"
Write-Host "Plugin:   $pluginDir\TreeSitterLexer.dll"
Get-ChildItem "$grammarDir\*.dll" | ForEach-Object { Write-Host "Grammar:  $($_.FullName)" }
Get-ChildItem "$grammarDir\*.scm" | ForEach-Object { Write-Host "Query:    $($_.FullName)" }
Get-ChildItem "$bundledGrammarDir\*.dll" | ForEach-Object { Write-Host "Bundle:   $($_.FullName)" }
Get-ChildItem "$bundledGrammarDir\*.scm" | ForEach-Object { Write-Host "BundleQ:  $($_.FullName)" }
Write-Host "Config:   $configDir\TreeSitterLexer.xml"
Write-Host "PluginCfg: $pluginDir\TreeSitterLexer.xml"

<#
.SYNOPSIS
    Packages the TreeSitterLexer plugin into a zip suitable for nppPluginList / Plugin Admin.
.DESCRIPTION
    Creates a zip file with the plugin DLL at the root level (not nested in subfolders),
    as required by Notepad++ Plugin Admin and the nppPluginList validator.

    Expected zip structure:
        TreeSitterLexer.dll
        TreeSitterLexer.xml
        Config/TreeSitterLexer.xml
        TreeSitterGrammars/tree-sitter-*.dll
        TreeSitterGrammars/*-*.scm
        BundledTreeSitterGrammars/tree-sitter-*.dll
        BundledTreeSitterGrammars/*-*.scm

    When Plugin Admin installs the plugin, it extracts everything into:
        <Npp>/plugins/TreeSitterLexer/
.PARAMETER Platform
    Target platform: x64, Win32/x86, or ARM64. Default: x64.
.PARAMETER Configuration
    Build configuration: Release or Debug. Default: Release.
.PARAMETER OutDir
    Output directory for the zip file. Default: release/
.EXAMPLE
    .\package-release.ps1
    .\package-release.ps1 -Platform ARM64
#>
param(
    [ValidateSet("x64","Win32","x86","ARM64")]
    [string]$Platform = "x64",
    [ValidateSet("Release","Debug")]
    [string]$Configuration = "Release",
    [string]$OutDir
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildPlatform = switch ($Platform) {
    "x64"   { "x64" }
    "Win32" { "Win32" }
    "x86"   { "Win32" }
    "ARM64" { "ARM64" }
}
$BuildDir  = Join-Path $ScriptDir "build\$BuildPlatform\$Configuration"
$DllPath   = Join-Path $BuildDir "TreeSitterLexer.dll"
$GrammarDir = Join-Path $BuildDir "TreeSitterGrammars"
$ConfigXml = Join-Path $ScriptDir "config\TreeSitterLexer.xml"

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

if (-not $OutDir) {
    $OutDir = Join-Path $ScriptDir "release"
}

# Validate build artifacts exist
if (-not (Test-Path $DllPath)) {
    throw "Plugin DLL not found at $DllPath. Build the project first."
}
if (-not (Test-Path $GrammarDir)) {
    throw "Grammar directory not found at $GrammarDir. Run grammars\build-grammars.ps1 first."
}
if (-not (Test-Path $ConfigXml)) {
    throw "Config XML not found at $ConfigXml."
}

# Determine zip name based on platform
$platformSuffix = switch ($Platform) {
    "x64"   { "x64" }
    "Win32" { "x86" }
    "x86"   { "x86" }
    "ARM64" { "arm64" }
}
$ZipName = "NppTreeSitter-$platformSuffix.zip"
$ZipPath = Join-Path $OutDir $ZipName

# Create a temporary staging directory with the correct flat structure
$StagingDir = Join-Path $OutDir "staging-$platformSuffix"
if (Test-Path $StagingDir) {
    Remove-Item $StagingDir -Recurse -Force
}
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

$StagingGrammarDir = Join-Path $StagingDir "TreeSitterGrammars"
New-Item -ItemType Directory -Path $StagingGrammarDir -Force | Out-Null

$StagingBundledGrammarDir = Join-Path $StagingDir "BundledTreeSitterGrammars"
New-Item -ItemType Directory -Path $StagingBundledGrammarDir -Force | Out-Null

$StagingConfigDir = Join-Path $StagingDir "Config"
New-Item -ItemType Directory -Path $StagingConfigDir -Force | Out-Null

Write-Host "=== Packaging NppTreeSitter Release ==="
Write-Host "Platform: $Platform"
Write-Host "Build dir: $BuildDir"
Write-Host "Source:   $BuildDir"
Write-Host "Output:   $ZipPath"
Write-Host ""

# Copy DLL to staging root (this is the critical part - DLL must be at zip root)
Write-Host "Copying TreeSitterLexer.dll to zip root..."
Copy-Item $DllPath $StagingDir -Force

# Copy config XML to staging root
Write-Host "Copying TreeSitterLexer.xml to zip root..."
Copy-Item $ConfigXml $StagingDir -Force

# Copy config XML to Config/ for direct deployment compatibility
Write-Host "Copying TreeSitterLexer.xml to Config/..."
Copy-Item $ConfigXml $StagingConfigDir -Force

# Copy grammar DLLs and query files
Write-Host "Copying grammar DLLs and query files..."
Get-ChildItem -Path $GrammarDir -Filter "tree-sitter-*.dll" -File | ForEach-Object {
    $language = $_.BaseName.Substring("tree-sitter-".Length)
    if ($OptionalBundledLanguages -notcontains $language) {
        Copy-GrammarLanguageFiles -SourceDir $GrammarDir -DestinationDir $StagingGrammarDir -Language $language
    }
}

# Copy bundled grammar install payload for the in-plugin installer
Write-Host "Copying bundled grammar install payload..."
foreach ($language in $OptionalBundledLanguages) {
    $dllPath = Join-Path $GrammarDir "tree-sitter-$language.dll"
    if (Test-Path $dllPath) {
        Copy-GrammarLanguageFiles -SourceDir $GrammarDir -DestinationDir $StagingBundledGrammarDir -Language $language
    }
}

# Remove old zip if it exists
if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

# Create the zip from the staging directory contents
Write-Host "Creating $ZipName..."
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal

# Clean up staging directory
Remove-Item $StagingDir -Recurse -Force

# Show the zip contents for verification
Write-Host ""
Write-Host "=== Zip contents ==="
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
foreach ($entry in $zip.Entries) {
    Write-Host "  $($entry.FullName)"
}
$zip.Dispose()

# Show the SHA-256 hash (needed for nppPluginList)
$hash = (Get-FileHash $ZipPath -Algorithm SHA256).Hash
Write-Host ""
Write-Host "=== Release Info ==="
Write-Host "File: $ZipPath"
Write-Host "Size: $((Get-Item $ZipPath).Length) bytes"
Write-Host "SHA-256: $hash"
Write-Host ""
Write-Host "Use this hash as the 'id' field in the nppPluginList JSON entry."

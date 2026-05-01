<#
.SYNOPSIS
    Downloads and compiles tree-sitter grammar DLLs for the TreeSitterLexer plugin.
.DESCRIPTION
    Reads grammars.json, downloads release tarballs from GitHub, reads each
    repo's tree-sitter.json for source layout, compiles to DLL via MSVC cl.exe.
    Adapted from the WinMerge build-grammars.ps1 script.
.PARAMETER OutDir
    Output directory for DLLs and .scm files.
.PARAMETER Platform
    Target platform: x64, Win32/x86, or ARM64. Default: x64.
.PARAMETER Configuration
    Build configuration: Release or Debug. Default: Release.
.PARAMETER GrammarFilter
    Optional regex to build only matching grammar names.
.EXAMPLE
    .\build-grammars.ps1
    .\build-grammars.ps1 -GrammarFilter "python|rust"
#>
param(
    [string]$OutDir,
    [ValidateSet("x64","Win32","x86","ARM64")]
    [string]$Platform = "x64",
    [ValidateSet("Release","Debug")]
    [string]$Configuration = "Release",
    [string]$GrammarFilter
)

$ErrorActionPreference = "Continue"

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot   = (Resolve-Path (Join-Path $ScriptDir "..")).Path
$BuildPlatform = switch ($Platform) {
    "x64"   { "x64" }
    "Win32" { "Win32" }
    "x86"   { "Win32" }
    "ARM64" { "ARM64" }
}
if (-not $OutDir) {
    $OutDir = Join-Path $RepoRoot "build\$BuildPlatform\$Configuration\TreeSitterGrammars"
    $OutDir = (New-Item -ItemType Directory -Path $OutDir -Force).FullName
}
$TempBase   = Join-Path $RepoRoot "BuildTmp\grammar-sources"
$ConfigFile = Join-Path $ScriptDir "grammars.json"

# ---- Locate and import MSVC environment ----

function Find-VcVarsAll {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        $vswhere = Join-Path $env:ProgramFiles "Microsoft Visual Studio\Installer\vswhere.exe"
    }
    if (-not (Test-Path $vswhere)) { throw "vswhere.exe not found" }
    $ip = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null
    if (-not $ip) { throw "No VS with C++ tools found" }
    $vcvars = Join-Path $ip "VC\Auxiliary\Build\vcvarsall.bat"
    if (-not (Test-Path $vcvars)) { throw "vcvarsall.bat not found at $vcvars" }
    return $vcvars
}

function Import-VcEnvironment {
    param([string]$Arch)
    $vcvars = Find-VcVarsAll
    Write-Host "Importing MSVC environment ($Arch) ..."
    $envFile = Join-Path $env:TEMP "vcvars_env.txt"
    $cmdLine = 'call "{0}" {1} >nul 2>&1 & set > "{2}"' -f $vcvars, $Arch, $envFile
    cmd.exe /d /c $cmdLine
    if (Test-Path $envFile) {
        foreach ($line in (Get-Content $envFile)) {
            if ($line -match '^([^=]+)=(.*)$') {
                [Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
            }
        }
        Remove-Item $envFile -Force -EA SilentlyContinue
    }
    $cl = Get-Command cl.exe -EA SilentlyContinue
    if (-not $cl) { throw "cl.exe not found after importing vcvars" }
    Write-Host "  cl.exe: $($cl.Source)"
}

function Get-VcArch {
    switch ($Platform) {
        "x64"   { return "amd64" }
        "Win32" { return "x86" }
        "x86"   { return "x86" }
        "ARM64" { return "amd64_arm64" }
        default { return "amd64" }
    }
}

function Get-LibArch {
    switch ($Platform) {
        "x64"   { return "x64" }
        "Win32" { return "x86" }
        "x86"   { return "x86" }
        "ARM64" { return "arm64" }
        default { return "x64" }
    }
}

# ---- Download grammar source ----

function Expand-ArchiveToDirectory {
    param(
        [string]$ArchivePath,
        [string]$ExtractDir
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        foreach ($entry in $zip.Entries) {
            if (-not $entry.FullName) {
                continue
            }

            $relativePath = $entry.FullName
            $slash = $relativePath.IndexOf('/')
            if ($slash -ge 0) {
                $relativePath = $relativePath.Substring($slash + 1)
            }

            if (-not $relativePath) {
                continue
            }

            $destPath = Join-Path $ExtractDir ($relativePath -replace '/', '\')
            if ($entry.FullName.EndsWith('/')) {
                New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                continue
            }

            $destParent = Split-Path -Parent $destPath
            if ($destParent) {
                New-Item -ItemType Directory -Path $destParent -Force | Out-Null
            }

            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
        }
    } finally {
        $zip.Dispose()
    }
}

function Get-GrammarSource {
    param([string]$Repo, [string]$Tag)

    function Test-GrammarSourceCache {
        param([string]$ExtractDir)

        $tsJsonPath = Join-Path $ExtractDir "tree-sitter.json"
        if (-not (Test-Path $tsJsonPath)) {
            return $false
        }

        try {
            $tsJson = Get-Content $tsJsonPath -Raw | ConvertFrom-Json
        } catch {
            return $false
        }

        if (-not $tsJson.grammars) {
            return $false
        }

        foreach ($grammar in $tsJson.grammars) {
            $grammarPath = $grammar.path
            if (-not $grammarPath) {
                $grammarPath = "."
            }

            $grammarDir = if ($grammarPath -eq ".") { $ExtractDir } else { Join-Path $ExtractDir $grammarPath }
            $parserPath = Join-Path $grammarDir "src\parser.c"
            $grammarJsPath = Join-Path $grammarDir "grammar.js"
            if (-not (Test-Path $parserPath) -and -not (Test-Path $grammarJsPath)) {
                return $false
            }
        }

        return $true
    }

    $name = ($Repo -split '/')[-1]
    $extractDir = Join-Path $TempBase $name
    $tagFile = Join-Path $extractDir ".opencode-tag"
    $cachedTag = $null
    $cacheValid = $false
    if (Test-Path $tagFile) {
        $cachedTag = (Get-Content $tagFile -Raw).Trim()
    }
    if (Test-Path $extractDir) {
        $cacheValid = Test-GrammarSourceCache -ExtractDir $extractDir
    }

    if ($cacheValid) {
        if ($cachedTag -eq $Tag) {
            Write-Host "  Cached: $extractDir ($Tag)"
            return $extractDir
        }

        if (-not $cachedTag) {
            Write-Host "  Refreshing cache: missing tag metadata"
        }
    } elseif (Test-Path $extractDir) {
        Write-Host "  Refreshing cache: existing source tree is incomplete"
    }

    if (Test-Path $extractDir) {
        if ($cachedTag) {
            Write-Host "  Refreshing cache: $cachedTag -> $Tag"
        } else {
            Write-Host "  Refreshing cache: unknown tag -> $Tag"
        }
        Remove-Item $extractDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $TempBase -Force | Out-Null
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null

    function Download-SourceZip {
        param([string]$RepoName, [string]$RepoTag)
        $zipUrl = "https://github.com/$Repo/archive/refs/tags/$RepoTag.zip"
        $zipFile = Join-Path $TempBase "$RepoName-$($RepoTag -replace '[^A-Za-z0-9._-]', '_').zip"
        Write-Host "  Downloading source archive for $RepoTag ..."
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing -ErrorAction Stop
        return $zipFile
    }

    $downloaded = $false
    $archivePath = $null
    $archiveType = $null
    foreach ($ext in @("tar.gz", "tar.xz")) {
        $tarUrl  = "https://github.com/$Repo/releases/download/$Tag/$name.$ext"
        $tarFile = Join-Path $TempBase "$name.$ext"
        try {
            Write-Host "  Downloading $name.$ext ..."
            Invoke-WebRequest -Uri $tarUrl -OutFile $tarFile -UseBasicParsing -ErrorAction Stop
            $downloaded = $true
            $archivePath = $tarFile
            $archiveType = "tar"
            break
        } catch {
            # Try next format
        }
    }
    if (-not $downloaded) {
        try {
            $zipFile = Download-SourceZip -RepoName $name -RepoTag $Tag
            $downloaded = $true
            $archivePath = $zipFile
            $archiveType = "zip"
        } catch {
            throw "No downloadable release or source archive found for $Repo $Tag"
        }
    }

    Write-Host "  Extracting ..."
    if ($archiveType -eq "zip") {
        Expand-ArchiveToDirectory -ArchivePath $archivePath -ExtractDir $extractDir
    } else {
        & tar.exe -xf $archivePath -C $extractDir
        if ($LASTEXITCODE -ne 0) {
            Remove-Item $archivePath -Force -EA SilentlyContinue
            Remove-Item $extractDir -Recurse -Force -EA SilentlyContinue
            New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
            try {
                $zipFile = Download-SourceZip -RepoName $name -RepoTag $Tag
                Write-Host "  Extracting fallback source archive ..."
                Expand-ArchiveToDirectory -ArchivePath $zipFile -ExtractDir $extractDir
                $archivePath = $zipFile
                $archiveType = "zip"
            } catch {
                throw "Failed to extract $archivePath"
            }
        }
    }
    Remove-Item $archivePath -Force

    if (-not (Test-GrammarSourceCache -ExtractDir $extractDir)) {
        Remove-Item $extractDir -Recurse -Force -EA SilentlyContinue
        throw "Extracted source for $Repo $Tag is incomplete (missing generated parser.c)"
    }

    Set-Content -Path $tagFile -Value $Tag -Encoding Ascii
    return $extractDir
}

# ---- Compile a grammar to DLL ----

function Build-GrammarDll {
    param(
        [string]$GrammarName,
        [string]$SourceDir,
        [string]$RepoDir,
        [string[]]$HighlightsScm,
        [string[]]$LocalsScm,
        [string[]]$InjectionsScm,
        [string[]]$TagsScm,
        [string]$DllName
    )

    function Ensure-GeneratedParser {
        param([string]$GrammarSourceDir)

        $parserPath = Join-Path $GrammarSourceDir "src\parser.c"
        if (Test-Path $parserPath) {
            return $true
        }

        $grammarJsPath = Join-Path $GrammarSourceDir "grammar.js"
        if (-not (Test-Path $grammarJsPath)) {
            return $false
        }

        Write-Host "  Generating parser sources ..."
        $generateArgs = @("exec", "--yes", "tree-sitter-cli", "--", "generate")
        $process = Start-Process npm -ArgumentList $generateArgs -WorkingDirectory $GrammarSourceDir -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            Write-Error "  tree-sitter generate failed for $GrammarName (exit $($process.ExitCode))"
            return $false
        }

        return (Test-Path $parserPath)
    }

    if (-not (Ensure-GeneratedParser -GrammarSourceDir $SourceDir)) {
        Write-Warning "  parser.c is unavailable and could not be generated for $GrammarName"
        return $false
    }

    $srcDir   = Join-Path $SourceDir "src"
    $parserC  = Join-Path $srcDir "parser.c"
    $scannerC = Join-Path $srcDir "scanner.c"
    if (-not (Test-Path $parserC)) {
        Write-Warning "  parser.c not found at $parserC - skipping $GrammarName"
        return $false
    }
    $sources = @($parserC)
    if (Test-Path $scannerC) { $sources += $scannerC }

    $buildDir = Join-Path $RepoRoot "BuildTmp\grammar-build\$DllName\$Platform\$Configuration"
    New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
    New-Item -ItemType Directory -Path $OutDir   -Force | Out-Null
    $dllPath = Join-Path $OutDir "$DllName.dll"

    $cflags = @("/nologo","/c","/TC","/W3","/D_USRDLL","/D_WINDOWS","/wd4996","/wd4267","/wd4244","/wd4101")
    if ($Configuration -eq "Release") {
        $cflags += @("/O2","/MD","/DNDEBUG","/GL")
    } else {
        $cflags += @("/Od","/MDd","/D_DEBUG","/Zi")
    }

    $objFiles = @()
    foreach ($src in $sources) {
        $objName = [IO.Path]::GetFileNameWithoutExtension($src) + ".obj"
        $objPath = Join-Path $buildDir $objName
        $objFiles += $objPath
        $fileName = [IO.Path]::GetFileName($src)
        Write-Host "  Compiling $fileName ..."
        $allArgs = $cflags + @("/I`"$srcDir`"", "/Fo`"$objPath`"", "`"$src`"")
        $p = Start-Process cl.exe -ArgumentList $allArgs -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -ne 0) {
            Write-Error "  cl.exe failed for $fileName (exit $($p.ExitCode))"
            return $false
        }
    }

    Write-Host "  Linking $DllName.dll ..."
    $linkArgs = @("/nologo","/DLL","/OUT:`"$dllPath`"")
    if ($Configuration -eq "Release") {
        $linkArgs += @("/LTCG","/OPT:REF","/OPT:ICF")
    } else {
        $linkArgs += @("/DEBUG")
    }
    $linkArgs += $objFiles
    $p = Start-Process link.exe -ArgumentList $linkArgs -NoNewWindow -Wait -PassThru
    if ($p.ExitCode -ne 0) {
        Write-Error "  link.exe failed for $DllName (exit $($p.ExitCode))"
        return $false
    }

    function Write-QueryFile {
        param(
            [string]$QueryName,
            [string[]]$QueryPaths
        )

        if (-not $QueryPaths -or $QueryPaths.Count -eq 0) {
            if ($QueryName -eq "highlights") {
                Write-Warning "  No highlights.scm for $GrammarName"
            }
            return
        }

        $dest = Join-Path $OutDir "$GrammarName-$QueryName.scm"
        $builder = New-Object System.Text.StringBuilder
        $first = $true
        foreach ($queryPath in $QueryPaths) {
            if (-not (Test-Path $queryPath)) {
                continue
            }

            if (-not $first) {
                [void]$builder.AppendLine()
            }
            [void]$builder.AppendLine("; Source: $([IO.Path]::GetFileName($queryPath))")
            [void]$builder.Append((Get-Content $queryPath -Raw))
            $first = $false
        }

        if ($first) {
            if ($QueryName -eq "highlights") {
                Write-Warning "  No highlights.scm for $GrammarName"
            }
            return
        }

        [System.IO.File]::WriteAllText($dest, $builder.ToString())
        Write-Host "  Wrote $QueryName -> $GrammarName-$QueryName.scm"
    }

    Write-QueryFile -QueryName "highlights" -QueryPaths $HighlightsScm
    Write-QueryFile -QueryName "locals" -QueryPaths $LocalsScm
    Write-QueryFile -QueryName "injections" -QueryPaths $InjectionsScm
    Write-QueryFile -QueryName "tags" -QueryPaths $TagsScm

    Write-Host "  OK: $dllPath"
    return $true
}

function Resolve-QueryPaths {
    param(
        [object]$QuerySpec,
        [string]$SourceDir,
        [string]$RepoDir,
        [string]$DefaultRelativePath
    )

    $resolved = New-Object System.Collections.Generic.List[string]
    $candidates = @()
    if ($QuerySpec) {
        $candidates = if ($QuerySpec -is [array]) { $QuerySpec } else { @($QuerySpec) }
    } elseif ($DefaultRelativePath) {
        $candidates = @($DefaultRelativePath)
    }

    foreach ($candidate in $candidates) {
        $tryPaths = @()

        if ($candidate -like 'node_modules/*') {
            $relative = $candidate.Substring('node_modules/'.Length)
            $parts = $relative -split '/'
            if ($parts.Length -ge 2) {
                $dependencyRoot = Join-Path $TempBase $parts[0]
                $remaining = ($parts[1..($parts.Length - 1)] -join '\')
                $tryPaths += (Join-Path $dependencyRoot $remaining)
            }
        }

        $tryPaths = @(
            (Join-Path $repoDir $candidate),
            (Join-Path $SourceDir $candidate)
        ) + $tryPaths

        foreach ($tryPath in $tryPaths) {
            if ((Test-Path $tryPath) -and -not $resolved.Contains($tryPath)) {
                $resolved.Add($tryPath)
                break
            }
        }
    }

    return $resolved.ToArray()
}

# ---- Main ----

$succeeded = [int]0
$failed    = [int]0
$skipped   = [int]0

Write-Host "=== TreeSitterLexer Grammar Builder ==="
Write-Host "Platform:      $Platform"
Write-Host "Configuration: $Configuration"
Write-Host "Output:        $OutDir"
Write-Host ""

Import-VcEnvironment -Arch (Get-VcArch)
$crtLibDir = Join-Path $env:VCToolsInstallDir ("lib\" + (Get-LibArch))
if (-not (Test-Path (Join-Path $crtLibDir "msvcrt.lib"))) {
    throw "MSVC runtime libraries for $Platform are not installed at $crtLibDir. Install the corresponding MSVC C++ toolchain before building grammars."
}
Write-Host ""

$config = Get-Content $ConfigFile -Raw | ConvertFrom-Json

foreach ($entry in $config.grammars) {
    $repo     = $entry.repo
    $tag      = $entry.tag
    $repoName = ($repo -split '/')[-1]
    Write-Host "--- $repoName ($tag) ---"

    try {
        $repoDir = Get-GrammarSource -Repo $repo -Tag $tag
    } catch {
        Write-Warning "  Download failed: $_"
        $failed++
        continue
    }

    $tsJsonPath = Join-Path $repoDir "tree-sitter.json"
    if (-not (Test-Path $tsJsonPath)) {
        Write-Warning "  No tree-sitter.json - skipping"
        $skipped++
        continue
    }
    $tsJson = Get-Content $tsJsonPath -Raw | ConvertFrom-Json

    foreach ($g in $tsJson.grammars) {
        $gName = $g.name
        $gPath = $g.path
        if (-not $gPath) { $gPath = "." }

        if ($GrammarFilter -and ($gName -notmatch $GrammarFilter)) {
            Write-Host "  Skipping $gName (filtered)"
            $skipped++
            continue
        }

        if ($Platform -eq "Win32" -and $gName -eq "flow") {
            Write-Warning "  Skipping flow on Win32 because the current x86 MSVC build hits an internal compiler error"
            $skipped++
            continue
        }

        $sourceDir = if ($gPath -eq ".") { $repoDir } else { Join-Path $repoDir $gPath }
        $dllName   = "tree-sitter-$gName"

        $hlScm = Resolve-QueryPaths -QuerySpec $g.highlights -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath "queries\highlights.scm"
        $localsScm = Resolve-QueryPaths -QuerySpec $g.locals -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath "queries\locals.scm"
        $injectionsScm = Resolve-QueryPaths -QuerySpec $g.injections -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath $null
        $tagsScm = Resolve-QueryPaths -QuerySpec $g.tags -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath $null

        Write-Host "  Grammar: $gName (path: $gPath)"
        $ok = Build-GrammarDll -GrammarName $gName -SourceDir $sourceDir -RepoDir $repoDir -HighlightsScm $hlScm -LocalsScm $localsScm -InjectionsScm $injectionsScm -TagsScm $tagsScm -DllName $dllName
        if ($ok) { $succeeded++ } else { $failed++ }
    }
    Write-Host ""
}

Write-Host "=== Done: $succeeded succeeded, $failed failed, $skipped skipped ==="
if ($failed -gt 0) { exit 1 }

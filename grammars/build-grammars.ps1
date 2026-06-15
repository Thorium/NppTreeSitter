<#
.SYNOPSIS
    Downloads and compiles tree-sitter grammar DLLs for the TreeSitterLexer plugin.
.DESCRIPTION
    Reads grammars.json, downloads release tarballs (or GitHub source archives)
    from GitHub, reads each repo's tree-sitter.json for source layout, and
    compiles to DLL via MSVC cl.exe.

    Source downloads are cached (tag-aware). Compilation is incremental (a DLL is
    rebuilt only when its sources or this script are newer) and runs in parallel
    across grammars (bounded by the CPU count). Repos that ship no
    tree-sitter.json can declare a grammar "name" directly in grammars.json.

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
            # No tree-sitter.json: accept the cache if a conventional parser source
            # (or a grammar.js we can generate from) exists at the repo root.
            $parserPath = Join-Path $ExtractDir "src\parser.c"
            $grammarJsPath = Join-Path $ExtractDir "grammar.js"
            return ((Test-Path $parserPath) -or (Test-Path $grammarJsPath))
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

# Generate src/parser.c from grammar.js (via tree-sitter-cli) when a repo ships
# no committed parser sources. Returns $true if parser.c is present afterwards.
function Ensure-GeneratedParser {
    param([string]$GrammarName, [string]$GrammarSourceDir)

    $parserPath = Join-Path $GrammarSourceDir "src\parser.c"
    if (Test-Path $parserPath) {
        return $true
    }

    $grammarJsPath = Join-Path $GrammarSourceDir "grammar.js"
    if (-not (Test-Path $grammarJsPath)) {
        return $false
    }

    Write-Host "  Generating parser sources for $GrammarName ..."
    # Invoke through cmd.exe so PATHEXT resolves npm -> npm.cmd. A bare "npm" via
    # Start-Process can hit the extensionless Unix shell script shipped alongside
    # node ("%1 is not a valid Win32 application").
    Push-Location $GrammarSourceDir
    try {
        & cmd.exe /c npm exec --yes tree-sitter-cli -- generate 2>&1 | Write-Host
        $exit = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    if ($exit -ne 0) {
        Write-Error "  tree-sitter generate failed for $GrammarName (exit $exit)"
        return $false
    }

    return (Test-Path $parserPath)
}

# ---- Query (.scm) bundling ----

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
            (Join-Path $SourceDir $candidate),
            (Join-Path $ScriptDir $candidate)
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

function Write-QueryFile {
    param(
        [string]$GrammarName,
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

# ---- Up-to-date check ----

# A grammar needs (re)building when its DLL is missing or older than any input
# (the parser/scanner sources or this script). Query .scm files are bundled
# separately and aren't compiled into the DLL, so they're not inputs here.
function Test-NeedsBuild {
    param([string]$DllPath, [string[]]$Inputs)
    if (-not (Test-Path $DllPath)) { return $true }
    $dllTime = (Get-Item $DllPath).LastWriteTimeUtc
    foreach ($in in $Inputs) {
        if ($in -and (Test-Path $in)) {
            if ((Get-Item $in).LastWriteTimeUtc -gt $dllTime) { return $true }
        }
    }
    return $false
}

# ---- Parallel compilation ----

# Self-contained compile+link for one grammar. Runs inside a background job, so
# it may use only its $Plan argument, the inherited MSVC environment, and
# built-in cmdlets (no script-level functions/variables).
$BuildGrammarScript = {
    param($Plan)
    New-Item -ItemType Directory -Path $Plan.BuildDir -Force | Out-Null
    $objs = @()
    foreach ($src in $Plan.Sources) {
        $obj = Join-Path $Plan.BuildDir ([IO.Path]::GetFileNameWithoutExtension($src) + ".obj")
        $objs += $obj
        # CFlags forces C with /TC; C++ scanners (.cc/.cpp/.cxx) need /TP instead.
        # cl applies the last /T* flag, so appending /TP overrides /TC for those files.
        $langFlag = @()
        if ($src -match '\.(cc|cpp|cxx)$') { $langFlag = @("/TP") }
        $clArgs = @($Plan.CFlags) + $langFlag + @("/I$($Plan.SrcDir)", "/Fo$obj", "$src")
        $out = & cl.exe @clArgs 2>&1
        if ($LASTEXITCODE -ne 0) {
            return [pscustomobject]@{ Name = $Plan.Name; Ok = $false; Stage = "compile $([IO.Path]::GetFileName($src))"; Output = ($out -join [Environment]::NewLine) }
        }
    }
    $linkArgs = @($Plan.LinkFlags) + @("/OUT:$($Plan.DllPath)") + $objs
    $out = & link.exe @linkArgs 2>&1
    return [pscustomobject]@{ Name = $Plan.Name; Ok = ($LASTEXITCODE -eq 0); Stage = "link"; Output = ($out -join [Environment]::NewLine) }
}

# Run build plans with bounded parallelism using background jobs (PowerShell 5.1
# compatible; ForEach-Object -Parallel needs pwsh 7, which the build doesn't use).
function Invoke-GrammarBuilds {
    param([scriptblock]$Action, [object[]]$Plans, [int]$Throttle)
    $results = @()
    if (-not $Plans -or $Plans.Count -eq 0) { return $results }
    if ($Throttle -lt 1) { $Throttle = 1 }
    $queue = New-Object System.Collections.Queue
    foreach ($pl in $Plans) { [void]$queue.Enqueue($pl) }
    $jobs = New-Object System.Collections.ArrayList
    while ($queue.Count -gt 0 -or $jobs.Count -gt 0) {
        while ($jobs.Count -lt $Throttle -and $queue.Count -gt 0) {
            [void]$jobs.Add((Start-Job -ScriptBlock $Action -ArgumentList $queue.Dequeue()))
        }
        $finished = Wait-Job -Job ($jobs.ToArray()) -Any
        foreach ($j in @($finished)) {
            $results += Receive-Job -Job $j
            Remove-Job -Job $j
            $jobs.Remove($j)
        }
    }
    return $results
}

# ---- Main ----

$succeeded = 0; $failed = 0; $skipped = 0; $cached = 0

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

# Shared compiler/linker flags.
$cflags    = @("/nologo","/c","/TC","/W3","/D_USRDLL","/D_WINDOWS","/wd4996","/wd4267","/wd4244","/wd4101")
$linkFlags = @("/nologo","/DLL")
if ($Configuration -eq "Release") {
    $cflags    += @("/O2","/MD","/DNDEBUG","/GL")
    $linkFlags += @("/LTCG","/OPT:REF","/OPT:ICF")
} else {
    $cflags    += @("/Od","/MDd","/D_DEBUG","/Zi")
    $linkFlags += @("/DEBUG")
}

$config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

# Phase 1 (sequential, fast): download sources, generate parsers when needed,
# bundle query files, and decide which grammars actually need recompiling.
$plans = @()
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

    # Most grammars ship a tree-sitter.json describing their grammar name(s) and
    # source layout. Some older/popular ones (e.g. kotlin, erlang) don't; for
    # those, grammars.json may supply an explicit "name" and we assume the
    # conventional single-grammar layout (sources in ./src, queries in ./queries).
    $tsJsonPath = Join-Path $repoDir "tree-sitter.json"
    if (Test-Path $tsJsonPath) {
        $tsJson = Get-Content $tsJsonPath -Raw | ConvertFrom-Json
        $grammarList = $tsJson.grammars
    } elseif ($entry.name) {
        Write-Host "  No tree-sitter.json; using name '$($entry.name)' from grammars.json"
        $grammarList = @([pscustomobject]@{ name = $entry.name; path = "."; highlights = $null; locals = $null; tags = $null; injections = $null })
    } else {
        Write-Warning "  No tree-sitter.json and no 'name' in grammars.json - skipping"
        $skipped++
        continue
    }

    foreach ($g in $grammarList) {
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
        Write-Host "  Grammar: $gName (path: $gPath)"

        if (-not (Ensure-GeneratedParser -GrammarName $gName -GrammarSourceDir $sourceDir)) {
            Write-Warning "  parser.c is unavailable and could not be generated for $gName"
            $failed++
            continue
        }

        $srcDir   = Join-Path $sourceDir "src"
        $parserC  = Join-Path $srcDir "parser.c"
        if (-not (Test-Path $parserC)) {
            Write-Warning "  parser.c not found at $parserC - skipping $gName"
            $failed++
            continue
        }
        $sources = @($parserC)
        # External scanners ship as scanner.c (C) or scanner.cc/.cpp (C++, e.g. hcl).
        # Compiling the wrong/no extension leaves tree_sitter_<lang>_external_scanner_*
        # unresolved at link time, so pick up whichever variant exists.
        foreach ($scannerName in @("scanner.c", "scanner.cc", "scanner.cpp")) {
            $scannerPath = Join-Path $srcDir $scannerName
            if (Test-Path $scannerPath) { $sources += $scannerPath; break }
        }

        # Resolve & bundle .scm query files (cheap; always refreshed).
        # grammars.json may override query locations per entry - needed for grammars
        # whose upstream repo lacks a usable queries/highlights.scm (we then point at a
        # vendored-queries/*.scm checked into this repo) or keeps queries off the
        # conventional path (e.g. queries/Neovim/highlights.scm). These overrides win
        # over the grammar's own tree-sitter.json paths and also apply to name-fallback
        # grammars (which have no tree-sitter.json at all).
        $hlSpec  = if ($entry.highlights)  { $entry.highlights }  else { $g.highlights }
        $locSpec = if ($entry.locals)      { $entry.locals }      else { $g.locals }
        $injSpec = if ($entry.injections)  { $entry.injections }  else { $g.injections }
        $tagSpec = if ($entry.tags)        { $entry.tags }        else { $g.tags }

        $hlScm = Resolve-QueryPaths -QuerySpec $hlSpec -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath "queries\highlights.scm"
        $localsScm = Resolve-QueryPaths -QuerySpec $locSpec -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath "queries\locals.scm"
        $injectionsScm = Resolve-QueryPaths -QuerySpec $injSpec -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath "queries\injections.scm"
        $tagsScm = Resolve-QueryPaths -QuerySpec $tagSpec -SourceDir $sourceDir -RepoDir $repoDir -DefaultRelativePath "queries\tags.scm"

        Write-QueryFile -GrammarName $gName -QueryName "highlights" -QueryPaths $hlScm
        Write-QueryFile -GrammarName $gName -QueryName "locals" -QueryPaths $localsScm
        Write-QueryFile -GrammarName $gName -QueryName "injections" -QueryPaths $injectionsScm
        Write-QueryFile -GrammarName $gName -QueryName "tags" -QueryPaths $tagsScm

        $dllPath  = Join-Path $OutDir "$dllName.dll"
        $buildDir = Join-Path $RepoRoot "BuildTmp\grammar-build\$dllName\$Platform\$Configuration"
        if (-not (Test-NeedsBuild -DllPath $dllPath -Inputs (@($sources) + @($PSCommandPath)))) {
            Write-Host "  Up to date: $dllName.dll"
            $cached++
            continue
        }

        $plans += @{
            Name      = $dllName
            Sources   = $sources
            SrcDir    = $srcDir
            BuildDir  = $buildDir
            DllPath   = $dllPath
            CFlags    = $cflags
            LinkFlags = $linkFlags
        }
    }
    Write-Host ""
}

# Phase 2 (parallel): compile + link the grammars that need it.
if ($plans.Count -gt 0) {
    $throttle = [Environment]::ProcessorCount
    if ($throttle -lt 1) { $throttle = 1 }
    Write-Host "Building $($plans.Count) grammar(s) with up to $throttle parallel job(s) ..."
    foreach ($r in (Invoke-GrammarBuilds -Action $BuildGrammarScript -Plans $plans -Throttle $throttle)) {
        if ($r.Ok) {
            Write-Host "  OK: $($r.Name).dll"
            $succeeded++
        } else {
            Write-Warning "  FAILED ($($r.Stage)): $($r.Name)"
            if ($r.Output) { Write-Host $r.Output }
            $failed++
        }
    }
}

Write-Host ""
Write-Host "=== Done: $succeeded built, $cached up-to-date, $failed failed, $skipped skipped ==="
if ($failed -gt 0) { exit 1 }

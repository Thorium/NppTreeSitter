# NppTreeSitter

A [Notepad++](https://notepad-plus-plus.org/) plugin that provides syntax highlighting and code folding powered by [tree-sitter](https://tree-sitter.github.io/) grammars.

From the Notepad++ menu:

> Language -> T -> treesitter.*

Instead of hand-written lexer rules, NppTreeSitter uses tree-sitter's incremental parsing to produce a full syntax tree, then maps tree-sitter highlight captures to Scintilla style IDs. This gives accurate, language-aware highlighting that understands the actual structure of the code.

## How it works

The plugin implements Scintilla's `ILexer5` interface and exports the [Lexilla](https://www.scintilla.org/LexillaDoc.html) API (`GetLexerCount`, `GetLexerName`, `CreateLexer`). Notepad++ discovers these at startup and registers each grammar as an external language (e.g. `treesitter.python`) in the Language menu.

Each language needs two files in the `TreeSitterGrammars` directory alongside the plugin DLL:

- **Grammar DLL** (e.g. `tree-sitter-python.dll`) -- compiled tree-sitter grammar exporting `tree_sitter_<lang>()`
- **Highlight query** (e.g. `python-highlights.scm`) -- standard tree-sitter `highlights.scm` with `@capture` names mapped to style categories

The tree-sitter core library is compiled directly into the plugin DLL (not a separate dependency).

## Supported languages

The grammar manifest (`grammars/grammars.json`) currently lists 17 languages:

Python, Rust, Go, C, C++, C#, F#, JavaScript, TypeScript, Java, JSON, HTML, CSS, Bash, Ruby, PHP, XML

Grammar DLLs are built from source using `grammars/build-grammars.ps1`.

## Building

### Prerequisites

- Visual Studio 2022 or newer (v143+ toolset)
- Windows SDK 10.0
- PowerShell (for building grammar DLLs)
- Git (for the tree-sitter submodule)

### Build the plugin

```
git clone --recurse-submodules https://github.com/<owner>/NppTreeSitter.git
cd NppTreeSitter
MSBuild NppTreeSitter.vcxproj -p:Configuration=Release -p:Platform=x64
```

Output: `build\x64\Release\TreeSitterLexer.dll`

### Build grammar DLLs

```powershell
cd grammars
.\build-grammars.ps1
```

This clones each grammar repo listed in `grammars.json`, compiles the parser source into a DLL, and copies the `highlights.scm` query file.

## Installation

Copy files into your Notepad++ installation:

```
<Notepad++>\plugins\TreeSitterLexer\TreeSitterLexer.dll
<Notepad++>\plugins\TreeSitterLexer\TreeSitterGrammars\tree-sitter-python.dll
<Notepad++>\plugins\TreeSitterLexer\TreeSitterGrammars\python-highlights.scm
<Notepad++>\plugins\TreeSitterLexer\TreeSitterGrammars\...  (other languages)
<Notepad++>\plugins\Config\TreeSitterLexer.xml
```

Or run `deploy.ps1` (requires admin privileges for writing to Program Files).

After restarting Notepad++, the tree-sitter languages appear at the bottom of the **Language** menu.

## Style mapping

Tree-sitter highlight captures are mapped to 14 Scintilla style IDs:

| Style ID | Name | Typical captures |
|----------|------|-----------------|
| 0 | Default | (unmarked nodes) |
| 1 | Comment | `@comment` |
| 2 | Comment Doc | `@comment.documentation` |
| 3 | Keyword | `@keyword`, `@keyword.*` |
| 4 | String | `@string`, `@string.*` |
| 5 | Number | `@number`, `@float` |
| 6 | Operator | `@operator` |
| 7 | Function | `@function`, `@function.*`, `@method` |
| 8 | Type | `@type`, `@type.*` |
| 9 | Preprocessor | `@preproc`, `@attribute` |
| 10 | Variable | `@variable`, `@variable.*`, `@parameter` |
| 11 | Constant | `@constant`, `@constant.*`, `@boolean` |
| 12 | Punctuation | `@punctuation.*` |
| 13 | Tag | `@tag`, `@tag.*` |

Colors and fonts for each style are configured in `config/TreeSitterLexer.xml` and can be customized through Notepad++'s **Settings > Style Configurator**.

## Project structure

```
NppTreeSitter/
  src/
    dllmain.cpp          # DLL entry, N++ plugin stubs, Lexilla exports
    TreeSitterLexer.h     # ILexer5 implementation, grammar loader, style map
    TreeSitterLexer.cpp   # Full implementation (~730 lines)
  config/
    TreeSitterLexer.xml   # N++ style/language config for all lexers
  grammars/
    grammars.json         # Grammar manifest (repos + tags)
    build-grammars.ps1    # Build script for grammar DLLs
  external/
    tree-sitter/          # git submodule: tree-sitter core library
    scintilla/include/    # Vendored Scintilla headers (ILexer.h, Sci_Position.h)
  deploy.ps1              # Deployment script
  NppTreeSitter.vcxproj   # MSVC project file
```

## License

GPL-2.0-or-later

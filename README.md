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

The grammar manifest (`grammars/grammars.json`) lists a curated bundled set of common languages and common project/config formats that are intended to ship in release packages.

Release packages currently split them into two groups:

- preinstalled in `TreeSitterGrammars`: Python, Rust, Go, C, C++, C#, F#, JavaScript, TypeScript, Java, JSON, HTML, CSS, Bash, Ruby, PHP, XML
- bundled for install-on-demand in `BundledTreeSitterGrammars`: R, Lua, YAML, TOML, Markdown, CMake, Fortran, Pascal, MATLAB

Grammar DLLs are built from source using `grammars/build-grammars.ps1`.

The wider tree-sitter ecosystem has many more parsers. See the upstream parser list:

- <https://github.com/tree-sitter/tree-sitter/wiki/List-of-parsers>

Not every parser should be bundled by default. The bundled set should stay focused on broadly useful languages and avoid unusually large or niche parser packs.

The practical policy is:

- bundle mainstream languages and common project/config formats
- prefer official or well-maintained parser repositories
- avoid shipping very large parser collections or low-value niche grammars by default
- let additional grammars be added separately when needed

## Adding more grammars

There are two different levels of support for additional grammars:

1. Add a grammar package for manual use
2. Add a grammar package plus auto-detect mapping

### Manual grammar install

The plugin discovers grammars dynamically at runtime from `TreeSitterGrammars`.

If you place these files alongside the plugin:

- `TreeSitterGrammars\tree-sitter-<language>.dll`
- `TreeSitterGrammars\<language>-highlights.scm`
- optional: `TreeSitterGrammars\<language>-locals.scm`
- optional: `TreeSitterGrammars\<language>-tags.scm`
- optional: `TreeSitterGrammars\<language>-injections.scm`

then Notepad++ can expose that lexer as `treesitter.<language>` without rebuilding the plugin DLL.

### Add a grammar to this repository

1. Choose a parser from the upstream list above.
2. Prefer repositories with:
   - a valid `tree-sitter.json`
   - either generated `src/parser.c` or a `grammar.js` that can be generated during the build
   - `queries/highlights.scm`
   - release tags or a stable tag to pin
3. Add the repo and tag to `grammars/grammars.json`.
4. Run `grammars\build-grammars.ps1` for the desired platform.
5. Copy the generated `tree-sitter-*.dll` and `*-*.scm` files into `TreeSitterGrammars`.

### Add auto-detect for a grammar

Dynamic grammar discovery is already supported, but file-extension auto-detect is still configured in code.

To make a newly added grammar auto-detect by file extension, update:

- `src/dllmain.cpp` in `DetectTreeSitterLanguageForFile`
- `src/dllmain.cpp` in `PreferBuiltInLexerForExtension`

This keeps the runtime flexible while still letting bundled languages have predictable defaults.

## Installing bundled grammars

Release packages now include a `BundledTreeSitterGrammars` payload alongside the active `TreeSitterGrammars` directory.

The `Install Missing Bundled Grammar` command copies the matching per-architecture grammar files from that bundled payload into `TreeSitterGrammars`, then activates the lexer for the current buffer.

This is meant to work in an installed Notepad++ plugin directory, not only from a developer checkout.

If a language is already present in `TreeSitterGrammars`, the command stays disabled because there is nothing to install.

## Building

### Prerequisites

- Visual Studio 2022 or newer (v143+ toolset)
- Windows SDK 10.0
- PowerShell (for building grammar DLLs)
- Git (for the tree-sitter submodule)

### Build the plugin

```powershell
git clone --recurse-submodules https://github.com/<owner>/NppTreeSitter.git
cd NppTreeSitter
MSBuild NppTreeSitter.vcxproj -p:Configuration=Release -p:Platform=x64
MSBuild NppTreeSitter.vcxproj -p:Configuration=Release -p:Platform=Win32
MSBuild NppTreeSitter.vcxproj -p:Configuration=Release -p:Platform=ARM64
```

Outputs:

- `build\x64\Release\TreeSitterLexer.dll`
- `build\Win32\Release\TreeSitterLexer.dll`
- `build\ARM64\Release\TreeSitterLexer.dll`

### Build grammar DLLs

```powershell
cd grammars
.\build-grammars.ps1 -Platform x64
.\build-grammars.ps1 -Platform Win32
.\build-grammars.ps1 -Platform ARM64
```

This downloads each grammar repo listed in `grammars.json`, compiles the parser source into a DLL, and copies the query files used by the plugin (`highlights`, `locals`, `injections`, `tags`) when they are available.

The builder now refuses to bless incomplete cached source trees, refreshes broken extractions, and generates `src/parser.c` on demand for grammars that ship `grammar.js` instead of committing generated C sources.

## Installation

Copy files into your Notepad++ installation:

```
<Notepad++>\plugins\TreeSitterLexer\TreeSitterLexer.dll
<Notepad++>\plugins\TreeSitterLexer\TreeSitterGrammars\tree-sitter-python.dll
<Notepad++>\plugins\TreeSitterLexer\TreeSitterGrammars\python-highlights.scm
<Notepad++>\plugins\TreeSitterLexer\TreeSitterGrammars\...  (other languages)
<Notepad++>\plugins\TreeSitterLexer\BundledTreeSitterGrammars\...  (bundled install payload)
<Notepad++>\plugins\Config\TreeSitterLexer.xml
```

Or run `deploy.ps1 -Platform x64`, `deploy.ps1 -Platform Win32`, or `deploy.ps1 -Platform ARM64` (requires admin privileges for writing to Program Files).

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

// TreeSitterLexer - DLL entry point, N++ plugin stubs, and Lexilla exports
//
// This DLL serves as BOTH a Notepad++ plugin AND an external lexer provider.
// N++ requires all plugins to export the full plugin interface (isUnicode,
// setInfo, getName, beNotified, messageProc, getFuncsArray) BEFORE it checks
// for Lexilla exports (GetLexerCount, GetLexerName, CreateLexer).
//
// Without the plugin stubs, N++ rejects the DLL as "ANSI plugin not
// compatible with Unicode Notepad++" and deletes it from the plugins dir.
//
// SPDX-License-Identifier: GPL-2.0-or-later

#include "TreeSitterLexer.h"

#include <tree_sitter/api.h>

#include <algorithm>
#include <cctype>
#include <cstring>
#include <cwctype>
#include <string>
#include <unordered_map>
#include <vector>

#include <shellapi.h>

#include <commctrl.h>
#include <winver.h>

using Scintilla::ILexer5;

// Forward declarations - avoid pulling in full N++ headers
struct SCNotification;

struct Sci_CharacterRangeFull {
    Sci_Position cpMin;
    Sci_Position cpMax;
};

struct Sci_TextRangeFull {
    Sci_CharacterRangeFull chrg;
    char* lpstrText;
};

struct SCNotification {
    NMHDR nmhdr;
    Sci_Position position;
    int ch;
    int modifiers;
    int modificationType;
    const char *text;
    Sci_Position length;
    Sci_Position linesAdded;
    int message;
    uintptr_t wParam;
    intptr_t lParam;
    Sci_Position line;
    int foldLevelNow;
    int foldLevelPrev;
    int margin;
    int listType;
    int x;
    int y;
    int token;
    Sci_Position annotationLinesAdded;
    int updated;
    int listCompletionMethod;
    int characterSource;
};

constexpr UINT NPPMSG = WM_USER + 1000;
constexpr UINT NPPM_GETCURRENTSCINTILLA = NPPMSG + 4;
constexpr UINT NPPM_GETCURRENTBUFFERID = NPPMSG + 60;
constexpr UINT NPPM_GETFULLPATHFROMBUFFERID = NPPMSG + 58;
constexpr UINT NPPM_MENUCOMMAND = NPPMSG + 48;
constexpr UINT NPPM_GETMENUHANDLE = NPPMSG + 25;
constexpr UINT NPPM_SETSTATUSBAR = NPPMSG + 24;
constexpr UINT NPPM_SETMENUITEMCHECK = NPPMSG + 40;

constexpr UINT NPPN_READY = 1001;
constexpr UINT NPPN_BUFFERACTIVATED = 1010;
constexpr UINT NPPN_LANGCHANGED = 1011;

constexpr int SC_UPDATE_SELECTION = 0x2;
constexpr UINT SCN_UPDATEUI = 2007;
constexpr UINT SCN_MODIFIED = 2008;
constexpr int SC_MOD_INSERTTEXT = 0x1;
constexpr int SC_MOD_DELETETEXT = 0x2;

constexpr int STATUSBAR_CUR_POS = 2;

constexpr UINT SCI_GETLENGTH = 2006;
constexpr UINT SCI_GETCURRENTPOS = 2008;
constexpr UINT SCI_GOTOPOS = 2025;
constexpr UINT SCI_LINEFROMPOSITION = 2166;
constexpr UINT SCI_GETSELECTIONSTART = 2143;
constexpr UINT SCI_GETSELECTIONEND = 2145;
constexpr UINT SCI_SETSEL = 2160;
constexpr UINT SCI_CLEARDOCUMENTSTYLE = 2005;
constexpr UINT SCI_COLOURISE = 4003;
constexpr UINT SCI_ENSUREVISIBLE = 2232;
constexpr UINT SCI_GRABFOCUS = 2400;
constexpr UINT SCI_GETTEXTRANGEFULL = 2039;
constexpr UINT SCI_GETLEXERLANGUAGE = 4012;
constexpr UINT SCI_SETILEXER = 4033;
constexpr UINT SCI_GETDOCPOINTER = 2357;
constexpr UINT SCI_INDICSETSTYLE = 2080;
constexpr UINT SCI_INDICSETFORE = 2081;
constexpr UINT SCI_SETINDICATORCURRENT = 2500;
constexpr UINT SCI_INDICATORFILLRANGE = 2504;
constexpr UINT SCI_INDICATORCLEARRANGE = 2505;
constexpr UINT SCI_INDICSETUNDER = 2510;
constexpr UINT SCI_INDICSETALPHA = 2523;
constexpr UINT SCI_INDICSETOUTLINEALPHA = 2558;

constexpr int INDIC_ROUNDBOX = 7;

// Indicator slot for symbol occurrence highlighting. Chosen above the range
// N++ uses for its own URL/search indicators and below the SCE_UNIVERSAL ones.
constexpr int kOccurrenceIndicator = 17;

// Skip per-caret-move reparsing on very large documents.
constexpr Sci_Position kMaxOccurrenceDocBytes = 2 * 1024 * 1024;

constexpr int MAIN_VIEW = 0;
constexpr int SUB_VIEW = 1;
constexpr int NPPMAINMENU = 1;

constexpr wchar_t kGitHubUrl[] = L"https://github.com/Thorium/NppTreeSitter";

// Startup isolation toggle used during crash diagnosis.
constexpr bool kDisableExternalLexersForStartupIsolation = false;

// Minimal reproduction of N++ plugin structs (from PluginInterface.h)
struct NppData
{
    HWND _nppHandle;
    HWND _scintillaMainHandle;
    HWND _scintillaSecondHandle;
};

typedef void (__cdecl * PFUNCPLUGINCMD)();

struct ShortcutKey
{
    bool _isCtrl;
    bool _isAlt;
    bool _isShift;
    UCHAR _key;
};

static const int menuItemSize = 64;
static const int kFuncItemCount = 14;

struct FuncItem
{
    wchar_t _itemName[menuItemSize];
    PFUNCPLUGINCMD _pFunc;
    int _cmdID;
    bool _init2Check;
    ShortcutKey *_pShKey;
};

extern FuncItem g_funcItems[kFuncItemCount];

// ============================================================================
// Globals
// ============================================================================
static HMODULE g_hModule = nullptr;
static NppData g_nppData = {};
static bool g_autoDetectEnabled = true;
static bool g_highlightOccurrencesEnabled = true;

namespace {

// Indices into g_funcItems (separators occupy 2, 7 and 12)
constexpr int kAutoDetectCommandIndex = 0;
constexpr int kInstallBundledGrammarCommandIndex = 1;
constexpr int kGoToDefinitionCommandIndex = 3;
constexpr int kSelectDefinitionCommandIndex = 4;
constexpr int kNextDefinitionCommandIndex = 5;
constexpr int kPrevDefinitionCommandIndex = 6;
constexpr int kExpandSelectionCommandIndex = 8;
constexpr int kShrinkSelectionCommandIndex = 9;
constexpr int kHighlightOccurrencesCommandIndex = 10;
constexpr int kShowSymbolPathCommandIndex = 11;

std::string DetectTreeSitterLanguageForFile(const std::wstring& path);

struct TagEntry {
    std::string name;
    std::string role;
    std::string kind;
    bool hasSymbolRange = false;
    uint32_t symbolStartByte = 0;
    uint32_t symbolEndByte = 0;
    uint32_t startByte = 0;
    uint32_t endByte = 0;
};

struct LocalDef {
    std::string name;
    uint32_t startByte = 0;
    uint32_t endByte = 0;
    uint32_t scopeStartByte = 0;
    uint32_t scopeEndByte = 0;
};

struct LocalScopeInfo {
    uint32_t startByte = 0;
    uint32_t endByte = 0;
    bool inherits = true;
};

struct SymbolOccurrence {
    std::string name;
    uint32_t startByte = 0;
    uint32_t endByte = 0;
    bool valid = false;
};

std::wstring ToLower(const std::wstring& value)
{
    std::wstring lower = value;
    for (auto& ch : lower)
        ch = static_cast<wchar_t>(std::towlower(ch));
    return lower;
}

std::string ToLowerAscii(const std::string& value)
{
    std::string lower = value;
    for (auto& ch : lower)
        ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    return lower;
}

std::wstring NormalizeMenuText(std::wstring text)
{
    text.erase(std::remove(text.begin(), text.end(), L'&'), text.end());
    const auto tab = text.find(L'\t');
    if (tab != std::wstring::npos)
        text.resize(tab);
    return text;
}

bool TryFindMenuCommandByText(HMENU menu, const std::wstring& targetText, UINT& commandId)
{
    if (!menu)
        return false;

    const int itemCount = GetMenuItemCount(menu);
    for (int i = 0; i < itemCount; ++i) {
        wchar_t buffer[256] = {};
        GetMenuStringW(menu, static_cast<UINT>(i), buffer, static_cast<int>(std::size(buffer)), MF_BYPOSITION);
        if (ToLower(NormalizeMenuText(buffer)) == ToLower(targetText)) {
            const UINT id = GetMenuItemID(menu, i);
            if (id != static_cast<UINT>(-1)) {
                commandId = id;
                return true;
            }
        }

        HMENU subMenu = GetSubMenu(menu, i);
        if (subMenu && TryFindMenuCommandByText(subMenu, targetText, commandId))
            return true;
    }

    return false;
}

bool OpenGitHubPage()
{
    const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(
        g_nppData._nppHandle,
        L"open",
        kGitHubUrl,
        nullptr,
        nullptr,
        SW_SHOWNORMAL));
    return result > 32;
}

std::wstring GetModuleVersionString()
{
    if (!g_hModule)
        return {};

    wchar_t modulePath[MAX_PATH] = {};
    const DWORD pathLen = GetModuleFileNameW(g_hModule, modulePath, static_cast<DWORD>(std::size(modulePath)));
    if (pathLen == 0 || pathLen >= std::size(modulePath))
        return {};

    DWORD handle = 0;
    const DWORD versionSize = GetFileVersionInfoSizeW(modulePath, &handle);
    if (versionSize == 0)
        return {};

    std::vector<BYTE> versionData(versionSize);
    if (!GetFileVersionInfoW(modulePath, 0, versionSize, versionData.data()))
        return {};

    struct LangAndCodePage {
        WORD wLanguage;
        WORD wCodePage;
    };

    LangAndCodePage* translate = nullptr;
    UINT translateLen = 0;
    if (!VerQueryValueW(versionData.data(), L"\\VarFileInfo\\Translation",
            reinterpret_cast<LPVOID*>(&translate), &translateLen) ||
        !translate || translateLen < sizeof(LangAndCodePage)) {
        return {};
    }

    wchar_t subBlock[64] = {};
    swprintf_s(subBlock, L"\\StringFileInfo\\%04x%04x\\FileVersion",
        translate[0].wLanguage, translate[0].wCodePage);

    LPWSTR versionText = nullptr;
    UINT versionTextLen = 0;
    if (!VerQueryValueW(versionData.data(), subBlock,
            reinterpret_cast<LPVOID*>(&versionText), &versionTextLen) ||
        !versionText || versionTextLen == 0) {
        return {};
    }

    return std::wstring(versionText);
}

HRESULT CALLBACK AboutDialogCallback(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, LONG_PTR)
{
    if (msg == TDN_HYPERLINK_CLICKED) {
        const auto* link = reinterpret_cast<LPCWSTR>(lParam);
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(
            hwnd,
            L"open",
            link,
            nullptr,
            nullptr,
            SW_SHOWNORMAL));
        return result > 32 ? S_OK : E_FAIL;
    }

    if (msg == TDN_CREATED) {
        SendMessageW(hwnd, TDM_SET_MARQUEE_PROGRESS_BAR, FALSE, 0);
    }

    return S_OK;
}

HWND GetCurrentScintilla()
{
    int which = MAIN_VIEW;
    if (!::SendMessageW(g_nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0,
            reinterpret_cast<LPARAM>(&which))) {
        return g_nppData._scintillaMainHandle;
    }
    return which == SUB_VIEW ? g_nppData._scintillaSecondHandle : g_nppData._scintillaMainHandle;
}

std::string GetCurrentLexerName(HWND hSci)
{
    if (!hSci)
        return {};

    const auto length = static_cast<int>(::SendMessageW(hSci, SCI_GETLEXERLANGUAGE, 0, 0));
    if (length <= 0)
        return {};

    std::string lexerName(static_cast<size_t>(length) + 1, '\0');
    ::SendMessageW(hSci, SCI_GETLEXERLANGUAGE, 0, reinterpret_cast<LPARAM>(lexerName.data()));
    lexerName.resize(static_cast<size_t>(length));
    return lexerName;
}

bool IsGenericLexerName(const std::string& lexerName)
{
    const std::string lower = ToLowerAscii(lexerName);
    return lower.empty() || lower == "null" || lower == "text" || lower == "txt";
}

std::wstring GetCurrentFilePath()
{
    const auto bufferId = static_cast<UINT_PTR>(::SendMessageW(g_nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));
    if (!bufferId)
        return {};

    const auto length = static_cast<int>(::SendMessageW(g_nppData._nppHandle,
        NPPM_GETFULLPATHFROMBUFFERID, static_cast<WPARAM>(bufferId), 0));
    if (length <= 0)
        return {};

    std::wstring path(static_cast<size_t>(length) + 1, L'\0');
    ::SendMessageW(g_nppData._nppHandle, NPPM_GETFULLPATHFROMBUFFERID,
        static_cast<WPARAM>(bufferId), reinterpret_cast<LPARAM>(path.data()));
    path.resize(static_cast<size_t>(length));
    return path;
}

std::wstring GetFileExtensionLower(const std::wstring& path)
{
    const auto slash = path.find_last_of(L"\\/");
    const auto dot = path.find_last_of(L'.');
    if (dot == std::wstring::npos || (slash != std::wstring::npos && dot < slash))
        return {};
    return ToLower(path.substr(dot + 1));
}

std::wstring GetFileNameLower(const std::wstring& path)
{
    const auto slash = path.find_last_of(L"\\/");
    const std::wstring name = slash == std::wstring::npos ? path : path.substr(slash + 1);
    return ToLower(name);
}

std::wstring GetModuleDirectory()
{
    if (!g_hModule)
        return {};

    wchar_t dllPath[MAX_PATH] = {};
    const DWORD pathLen = GetModuleFileNameW(g_hModule, dllPath, static_cast<DWORD>(std::size(dllPath)));
    if (pathLen == 0 || pathLen >= std::size(dllPath))
        return {};

    std::wstring dir(dllPath);
    const auto pos = dir.rfind(L'\\');
    if (pos != std::wstring::npos)
        dir.resize(pos);
    return dir;
}

bool FileExists(const std::wstring& path)
{
    const DWORD attr = GetFileAttributesW(path.c_str());
    return attr != INVALID_FILE_ATTRIBUTES && !(attr & FILE_ATTRIBUTE_DIRECTORY);
}

bool EnsureDirectoryExists(const std::wstring& path)
{
    if (path.empty())
        return false;

    const DWORD attr = GetFileAttributesW(path.c_str());
    if (attr != INVALID_FILE_ATTRIBUTES)
        return (attr & FILE_ATTRIBUTE_DIRECTORY) != 0;

    return CreateDirectoryW(path.c_str(), nullptr) != 0 || GetLastError() == ERROR_ALREADY_EXISTS;
}

bool CopyFileIfExists(const std::wstring& source, const std::wstring& destination)
{
    if (!FileExists(source))
        return false;
    return CopyFileW(source.c_str(), destination.c_str(), FALSE) != 0;
}

std::wstring GetDeployedGrammarDirectory()
{
    TreeSitterRegistry::Instance().Initialize(g_hModule);
    return TreeSitterRegistry::Instance().GetGrammarDir();
}

std::wstring GetBundledGrammarPackageDirectory()
{
    const std::wstring moduleDir = GetModuleDirectory();
    if (moduleDir.empty())
        return {};
    return moduleDir + L"\\BundledTreeSitterGrammars";
}

bool IsGrammarInstalled(const std::string& language)
{
    if (language.empty())
        return false;

    const std::wstring grammarDir = GetDeployedGrammarDirectory();
    if (grammarDir.empty())
        return false;

    const std::wstring languageW(language.begin(), language.end());
    return FileExists(grammarDir + L"\\tree-sitter-" + languageW + L".dll") &&
        FileExists(grammarDir + L"\\" + languageW + L"-highlights.scm");
}

bool IsBundledGrammarBuildAvailable(const std::string& language)
{
    if (language.empty())
        return false;

    const std::wstring bundledGrammarDir = GetBundledGrammarPackageDirectory();
    if (bundledGrammarDir.empty())
        return false;

    const std::wstring languageW(language.begin(), language.end());
    return FileExists(bundledGrammarDir + L"\\tree-sitter-" + languageW + L".dll") &&
        FileExists(bundledGrammarDir + L"\\" + languageW + L"-highlights.scm");
}

bool CanInstallBundledGrammarForCurrentFile(HWND hSci)
{
    if (!hSci)
        return false;

    const std::wstring path = GetCurrentFilePath();
    const std::string language = DetectTreeSitterLanguageForFile(path);
    if (language.empty())
        return false;

    if (IsGrammarInstalled(language))
        return false;

    return IsBundledGrammarBuildAvailable(language);
}

bool ActivateTreeSitterLexer(HWND hSci, const std::string& language)
{
    if (!hSci || language.empty())
        return false;

    TreeSitterRegistry::Instance().Refresh(g_hModule);
    const std::string lexerName = "treesitter." + language;
    Scintilla::ILexer5* lexer = TreeSitterRegistry::Instance().CreateLexer(lexerName.c_str());
    if (!lexer)
        return false;

    ::SendMessageW(hSci, SCI_SETILEXER, 0, reinterpret_cast<LPARAM>(lexer));
    ::SendMessageW(hSci, SCI_CLEARDOCUMENTSTYLE, 0, 0);
    const auto docLen = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETLENGTH, 0, 0));
    ::SendMessageW(hSci, SCI_COLOURISE, 0, static_cast<LPARAM>(docLen));
    return true;
}

bool TryInstallBundledGrammar(const std::string& language)
{
    if (language.empty())
        return false;

    const std::wstring bundledGrammarDir = GetBundledGrammarPackageDirectory();
    const std::wstring deployedGrammarDir = GetDeployedGrammarDirectory();
    if (bundledGrammarDir.empty() || deployedGrammarDir.empty())
        return false;

    if (!EnsureDirectoryExists(deployedGrammarDir))
        return false;

    std::wstring languageW(language.begin(), language.end());
    const std::wstring sourceDll = bundledGrammarDir + L"\\tree-sitter-" + languageW + L".dll";
    const std::wstring targetDll = deployedGrammarDir + L"\\tree-sitter-" + languageW + L".dll";
    if (!FileExists(sourceDll))
        return false;

    bool copiedAny = false;
    copiedAny = CopyFileIfExists(sourceDll, targetDll) || copiedAny;

    const std::wstring queryKinds[] = { L"highlights", L"locals", L"injections", L"tags" };
    for (const auto& kind : queryKinds) {
        const std::wstring sourceQuery = bundledGrammarDir + L"\\" + languageW + L"-" + kind + L".scm";
        const std::wstring targetQuery = deployedGrammarDir + L"\\" + languageW + L"-" + kind + L".scm";
        copiedAny = CopyFileIfExists(sourceQuery, targetQuery) || copiedAny;
    }

    return copiedAny;
}

std::string DetectTreeSitterLanguageForFile(const std::wstring& path)
{
    const std::wstring fileName = GetFileNameLower(path);
    if (fileName == L"dockerfile")
        return "dockerfile";
    if (fileName == L"cmakelists.txt")
        return "cmake";

    const std::wstring extension = GetFileExtensionLower(path);
    static const std::unordered_map<std::wstring, std::string> kExtensionMap = {
        { L"py", "python" },
        { L"pyw", "python" },
        { L"r", "r" },
        { L"rmd", "markdown" },
        { L"rs", "rust" },
        { L"go", "go" },
        { L"c", "c" },
        { L"cc", "cpp" },
        { L"cpp", "cpp" },
        { L"cxx", "cpp" },
        { L"h", "cpp" },
        { L"hh", "cpp" },
        { L"hpp", "cpp" },
        { L"hxx", "cpp" },
        { L"cs", "c-sharp" },
        { L"js", "javascript" },
        { L"mjs", "javascript" },
        { L"cjs", "javascript" },
        { L"jsx", "javascript" },
        { L"ts", "typescript" },
        { L"tsx", "tsx" },
        { L"java", "java" },
        { L"json", "json" },
        { L"htm", "html" },
        { L"html", "html" },
        { L"shtml", "html" },
        { L"xhtml", "html" },
        { L"css", "css" },
        { L"sh", "bash" },
        { L"bash", "bash" },
        { L"zsh", "bash" },
        { L"rb", "ruby" },
        { L"rake", "ruby" },
        { L"gemspec", "ruby" },
        { L"pl", "perl" },
        { L"pm", "perl" },
        { L"t", "perl" },
        { L"php", "php" },
        { L"xml", "xml" },
        { L"xsd", "xml" },
        { L"xsl", "xml" },
        { L"xslt", "xml" },
        { L"svg", "xml" },
        { L"fs", "fsharp" },
        { L"fsx", "fsharp" },
        { L"fsi", "fsharp_signature" },
        { L"lua", "lua" },
        { L"yaml", "yaml" },
        { L"yml", "yaml" },
        { L"toml", "toml" },
        { L"md", "markdown" },
        { L"markdown", "markdown" },
        { L"mkd", "markdown" },
        { L"cmake", "cmake" },
        { L"sql", "sql" },
        { L"f", "fortran" },
        { L"for", "fortran" },
        { L"f90", "fortran" },
        { L"f95", "fortran" },
        { L"f03", "fortran" },
        { L"f08", "fortran" },
        { L"pas", "pascal" },
        { L"pp", "pascal" },
        { L"inc", "pascal" },
    };

    const auto it = kExtensionMap.find(extension);
    if (it != kExtensionMap.end())
        return it->second;
    return {};
}

bool PreferBuiltInLexerForExtension(const std::wstring& extension)
{
    (void)extension;
    return false;
}

std::string TrimSelection(const std::string& value)
{
    const auto start = value.find_first_not_of(" \t\r\n");
    if (start == std::string::npos)
        return {};
    const auto end = value.find_last_not_of(" \t\r\n");
    return value.substr(start, end - start + 1);
}

std::string GetDocumentText(HWND hSci)
{
    if (!hSci)
        return {};

    const auto docEnd = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETLENGTH, 0, 0));
    if (docEnd <= 0)
        return {};

    std::string text(static_cast<size_t>(docEnd) + 1, '\0');
    Sci_TextRangeFull range{};
    range.chrg.cpMin = 0;
    range.chrg.cpMax = docEnd;
    range.lpstrText = text.data();
    ::SendMessageW(hSci, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&range));
    text.resize(static_cast<size_t>(docEnd));
    return text;
}

std::string GetSelectedOrCurrentWord(HWND hSci)
{
    if (!hSci)
        return {};

    auto selStart = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONSTART, 0, 0));
    auto selEnd = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONEND, 0, 0));
    const auto docLen = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETLENGTH, 0, 0));
    if (docLen <= 0)
        return {};

    std::string doc = GetDocumentText(hSci);
    if (doc.empty())
        return {};

    if (selStart < selEnd && selEnd <= docLen) {
        return TrimSelection(doc.substr(static_cast<size_t>(selStart), static_cast<size_t>(selEnd - selStart)));
    }

    auto pos = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETCURRENTPOS, 0, 0));
    if (pos < 0 || pos > docLen)
        return {};

    auto isWordChar = [](unsigned char ch) {
        return std::isalnum(ch) || ch == '_' || ch == '$';
    };

    // If the caret sits just after an identifier, treat that identifier as current.
    if (pos > 0 && (pos == docLen || !isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(pos)])))) {
        const auto prev = static_cast<unsigned char>(doc[static_cast<size_t>(pos - 1)]);
        if (isWordChar(prev))
            --pos;
    }

    auto start = pos;
    while (start > 0 && isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(start - 1)])))
        --start;
    auto end = pos;
    while (end < docLen && isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(end)])))
        ++end;

    if (start >= end)
        return {};
    return doc.substr(static_cast<size_t>(start), static_cast<size_t>(end - start));
}

std::string GetActiveTreeSitterLanguage(HWND hSci)
{
    std::string lexerName = GetCurrentLexerName(hSci);
    const std::string prefix = "treesitter.";
    if (lexerName.compare(0, prefix.size(), prefix) == 0)
        return lexerName.substr(prefix.size());
    return {};
}

const TreeSitterGrammar* GetActiveGrammar(HWND hSci)
{
    const std::string language = GetActiveTreeSitterLanguage(hSci);
    if (language.empty())
        return nullptr;

    TreeSitterRegistry::Instance().Initialize(g_hModule);
    return TreeSitterRegistry::Instance().GetGrammarByName(language);
}

std::string GetSetProperty(const TSQuery* query, uint32_t patternIndex, const std::string& key)
{
    uint32_t stepCount = 0;
    const TSQueryPredicateStep* steps = ts_query_predicates_for_pattern(query, patternIndex, &stepCount);
    if (!steps)
        return {};

    for (uint32_t i = 0; i + 2 < stepCount; ++i) {
        if (steps[i].type != TSQueryPredicateStepTypeString)
            continue;

        uint32_t nameLen = 0;
        const char* name = ts_query_string_value_for_id(query, steps[i].value_id, &nameLen);
        if (!name || std::string(name, nameLen) != "set!")
            continue;

        if (steps[i + 1].type != TSQueryPredicateStepTypeString ||
            steps[i + 2].type != TSQueryPredicateStepTypeString) {
            continue;
        }

        uint32_t keyLen = 0;
        const char* keyStr = ts_query_string_value_for_id(query, steps[i + 1].value_id, &keyLen);
        if (!keyStr || std::string(keyStr, keyLen) != key)
            continue;

        uint32_t valueLen = 0;
        const char* value = ts_query_string_value_for_id(query, steps[i + 2].value_id, &valueLen);
        if (value)
            return std::string(value, valueLen);
    }

    return {};
}

SymbolOccurrence GetCurrentSymbolOccurrence(HWND hSci, const std::string& doc)
{
    SymbolOccurrence occurrence;
    if (!hSci || doc.empty())
        return occurrence;

    auto selStart = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONSTART, 0, 0));
    auto selEnd = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONEND, 0, 0));
    const auto docLen = static_cast<Sci_Position>(doc.size());

    auto isWordChar = [](unsigned char ch) {
        return std::isalnum(ch) || ch == '_' || ch == '$' || ch == '\'';
    };

    if (selStart < selEnd && selEnd <= docLen) {
        while (selStart < selEnd && std::isspace(static_cast<unsigned char>(doc[static_cast<size_t>(selStart)])))
            ++selStart;
        while (selEnd > selStart && std::isspace(static_cast<unsigned char>(doc[static_cast<size_t>(selEnd - 1)])))
            --selEnd;

        if (selStart < selEnd) {
            bool allWordChars = true;
            for (auto i = selStart; i < selEnd; ++i) {
                if (!isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(i)]))) {
                    allWordChars = false;
                    break;
                }
            }
            if (allWordChars) {
                occurrence.name = doc.substr(static_cast<size_t>(selStart), static_cast<size_t>(selEnd - selStart));
                occurrence.startByte = static_cast<uint32_t>(selStart);
                occurrence.endByte = static_cast<uint32_t>(selEnd);
                occurrence.valid = !occurrence.name.empty();
            }
        }

        return occurrence;
    }

    auto pos = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETCURRENTPOS, 0, 0));
    if (pos < 0 || pos > docLen)
        return occurrence;

    if (pos > 0 && (pos == docLen || !isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(pos)])))) {
        const auto prev = static_cast<unsigned char>(doc[static_cast<size_t>(pos - 1)]);
        if (isWordChar(prev))
            --pos;
    }

    auto start = pos;
    while (start > 0 && isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(start - 1)])))
        --start;
    auto end = pos;
    while (end < docLen && isWordChar(static_cast<unsigned char>(doc[static_cast<size_t>(end)])))
        ++end;

    if (start >= end)
        return occurrence;

    occurrence.name = doc.substr(static_cast<size_t>(start), static_cast<size_t>(end - start));
    occurrence.startByte = static_cast<uint32_t>(start);
    occurrence.endByte = static_cast<uint32_t>(end);
    occurrence.valid = true;
    return occurrence;
}

std::vector<LocalDef> CollectLocals(const TreeSitterGrammar* grammar, const std::string& docText)
{
    std::vector<LocalDef> defs;
    if (!grammar || !grammar->GetLanguage() || !grammar->GetLocalsQuery() || docText.empty())
        return defs;

    TSParser* parser = ts_parser_new();
    if (!parser)
        return defs;

    ts_parser_set_language(parser, grammar->GetLanguage());
    TSTree* tree = ts_parser_parse_string(parser, nullptr, docText.c_str(), static_cast<uint32_t>(docText.size()));
    if (!tree) {
        ts_parser_delete(parser);
        return defs;
    }

    const TSQuery* query = grammar->GetLocalsQuery();
    TSQueryCursor* cursor = ts_query_cursor_new();
    if (!cursor) {
        ts_tree_delete(tree);
        ts_parser_delete(parser);
        return defs;
    }

    std::vector<LocalScopeInfo> scopes;
    ts_query_cursor_exec(cursor, query, ts_tree_root_node(tree));

    TSQueryMatch match;
    while (ts_query_cursor_next_match(cursor, &match)) {
        for (uint16_t i = 0; i < match.capture_count; ++i) {
            const TSQueryCapture capture = match.captures[i];
            uint32_t captureLen = 0;
            const char* captureName = ts_query_capture_name_for_id(query, capture.index, &captureLen);
            if (!captureName)
                continue;

            const std::string name(captureName, captureLen);
            const uint32_t nodeStart = ts_node_start_byte(capture.node);
            const uint32_t nodeEnd = ts_node_end_byte(capture.node);

            if (name == "local.scope") {
                LocalScopeInfo scope;
                scope.startByte = nodeStart;
                scope.endByte = nodeEnd;
                scope.inherits = GetSetProperty(query, match.pattern_index, "local.scope-inherits") != "false";
                scopes.push_back(scope);
                continue;
            }

            if (name != "local.definition")
                continue;
            if (nodeStart >= nodeEnd || nodeEnd > docText.size())
                continue;

            LocalDef def;
            def.name = docText.substr(nodeStart, nodeEnd - nodeStart);
            def.startByte = nodeStart;
            def.endByte = nodeEnd;

            const LocalScopeInfo* bestScope = nullptr;
            for (const auto& scope : scopes) {
                if (nodeStart >= scope.startByte && nodeEnd <= scope.endByte) {
                    if (!bestScope ||
                        (scope.endByte - scope.startByte) < (bestScope->endByte - bestScope->startByte)) {
                        bestScope = &scope;
                    }
                }
            }

            if (bestScope) {
                def.scopeStartByte = bestScope->startByte;
                def.scopeEndByte = bestScope->endByte;
            }
            defs.push_back(std::move(def));
        }
    }

    ts_query_cursor_delete(cursor);
    ts_tree_delete(tree);
    ts_parser_delete(parser);
    return defs;
}

const TagEntry* FindLocalDefinitionAtCaret(const TreeSitterGrammar* grammar, const std::string& docText,
    const SymbolOccurrence& symbol)
{
    if (!symbol.valid)
        return nullptr;

    static std::vector<TagEntry> localResults;
    localResults.clear();

    const std::vector<LocalDef> defs = CollectLocals(grammar, docText);
    for (const auto& def : defs) {
        if (def.name != symbol.name)
            continue;
        if (def.startByte == symbol.startByte && def.endByte == symbol.endByte)
            continue;
        if (def.startByte > symbol.startByte)
            continue;
        if (def.scopeStartByte > 0 || def.scopeEndByte > 0) {
            if (symbol.startByte < def.scopeStartByte || symbol.endByte > def.scopeEndByte)
                continue;
        }

        TagEntry entry;
        entry.name = def.name;
        entry.role = "definition";
        entry.kind = "local";
        entry.hasSymbolRange = true;
        entry.symbolStartByte = def.startByte;
        entry.symbolEndByte = def.endByte;
        entry.startByte = def.startByte;
        entry.endByte = def.endByte;
        localResults.push_back(std::move(entry));
    }

    const TagEntry* best = nullptr;
    for (const auto& entry : localResults) {
        if (!best || entry.startByte > best->startByte)
            best = &entry;
    }
    return best;
}

std::vector<TagEntry> CollectTags(const TreeSitterGrammar* grammar, const std::string& docText)
{
    std::vector<TagEntry> tags;
    if (!grammar || !grammar->GetLanguage() || !grammar->GetTagsQuery() || docText.empty())
        return tags;

    TSParser* parser = ts_parser_new();
    if (!parser)
        return tags;

    ts_parser_set_language(parser, grammar->GetLanguage());
    TSTree* tree = ts_parser_parse_string(parser, nullptr, docText.c_str(), static_cast<uint32_t>(docText.size()));
    if (!tree) {
        ts_parser_delete(parser);
        return tags;
    }

    const TSQuery* query = grammar->GetTagsQuery();
    TSQueryCursor* cursor = ts_query_cursor_new();
    if (!cursor) {
        ts_tree_delete(tree);
        ts_parser_delete(parser);
        return tags;
    }

    ts_query_cursor_exec(cursor, query, ts_tree_root_node(tree));

    TSQueryMatch match;
    while (ts_query_cursor_next_match(cursor, &match)) {
        TagEntry entry;
        entry.role = GetSetProperty(query, match.pattern_index, "role");
        entry.kind = GetSetProperty(query, match.pattern_index, "kind");

        uint32_t earliestStart = UINT32_MAX;
        uint32_t latestEnd = 0;

        for (uint16_t i = 0; i < match.capture_count; ++i) {
            const TSQueryCapture capture = match.captures[i];
            uint32_t captureLen = 0;
            const char* captureName = ts_query_capture_name_for_id(query, capture.index, &captureLen);
            if (!captureName)
                continue;

            const std::string name(captureName, captureLen);
            const uint32_t capStart = ts_node_start_byte(capture.node);
            const uint32_t capEnd = ts_node_end_byte(capture.node);

            if (capStart < earliestStart)
                earliestStart = capStart;
            if (capEnd > latestEnd)
                latestEnd = capEnd;

            if (name.rfind("name", 0) == 0) {
                entry.hasSymbolRange = true;
                entry.symbolStartByte = capStart;
                entry.symbolEndByte = capEnd;
                if (entry.symbolStartByte < entry.symbolEndByte && entry.symbolEndByte <= docText.size()) {
                    entry.name = docText.substr(entry.symbolStartByte, entry.symbolEndByte - entry.symbolStartByte);
                }
            }

            const auto dot = name.find('.');
            if (dot != std::string::npos) {
                const std::string captureRole = name.substr(0, dot);
                if (entry.role.empty() && (captureRole == "definition" || captureRole == "reference"))
                    entry.role = captureRole;

                if (entry.kind.empty() && dot + 1 < name.size())
                    entry.kind = name.substr(dot + 1);
            }
        }

        if (!entry.name.empty()) {
            if (entry.role.empty())
                entry.role = "definition";
            if (entry.startByte == 0 && entry.endByte == 0 && earliestStart != UINT32_MAX && earliestStart < latestEnd) {
                entry.startByte = earliestStart;
                entry.endByte = latestEnd;
            }
            if (entry.hasSymbolRange) {
                entry.startByte = entry.symbolStartByte;
                entry.endByte = entry.symbolEndByte;
            }
            tags.push_back(std::move(entry));
        }
    }

    ts_query_cursor_delete(cursor);
    ts_tree_delete(tree);
    ts_parser_delete(parser);
    return tags;
}


void ShowStatus(const std::wstring& text, int part = STATUSBAR_CUR_POS)
{
    ::SendMessageW(g_nppData._nppHandle, NPPM_SETSTATUSBAR, static_cast<WPARAM>(part),
        reinterpret_cast<LPARAM>(text.c_str()));
}

void ShowInstallResultDialog(const wchar_t* title, const std::wstring& message, UINT flags = MB_OK | MB_ICONINFORMATION)
{
    ::MessageBoxW(g_nppData._nppHandle, message.c_str(), title, flags);
}

void NavigateToPosition(HWND hSci, Sci_Position pos)
{
    if (!hSci || pos < 0)
        return;

    ::SendMessageW(hSci, SCI_GOTOPOS, static_cast<WPARAM>(pos), 0);
    ::SendMessageW(hSci, SCI_SETSEL, static_cast<WPARAM>(pos), static_cast<LPARAM>(pos));
    const auto line = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_LINEFROMPOSITION,
        static_cast<WPARAM>(pos), 0));
    if (line >= 0)
        ::SendMessageW(hSci, SCI_ENSUREVISIBLE, static_cast<WPARAM>(line), 0);
    ::SendMessageW(hSci, SCI_GRABFOCUS, 0, 0);
}

void SetCommandEnabled(int commandIndex, bool enabled)
{
    if (commandIndex < 0 || commandIndex >= kFuncItemCount)
        return;

    if (g_funcItems[commandIndex]._cmdID == 0)
        return;

    g_funcItems[commandIndex]._init2Check = enabled;
    ::EnableMenuItem(
        reinterpret_cast<HMENU>(::SendMessageW(g_nppData._nppHandle, NPPM_GETMENUHANDLE, 0, 0)),
        static_cast<UINT>(g_funcItems[commandIndex]._cmdID),
        MF_BYCOMMAND | (enabled ? MF_ENABLED : MF_GRAYED));
}

void SetMenuItemChecked(int commandIndex, bool checked)
{
    if (commandIndex < 0 || commandIndex >= kFuncItemCount)
        return;

    if (g_funcItems[commandIndex]._cmdID == 0)
        return;

    ::SendMessageW(
        g_nppData._nppHandle,
        NPPM_SETMENUITEMCHECK,
        static_cast<WPARAM>(g_funcItems[commandIndex]._cmdID),
        static_cast<LPARAM>(checked ? TRUE : FALSE));
}

const TagEntry* FindDefinitionAtCaret(HWND hSci)
{
    if (!hSci)
        return nullptr;

    const TreeSitterGrammar* grammar = GetActiveGrammar(hSci);
    if (!grammar)
        return nullptr;

    const std::string docText = GetDocumentText(hSci);
    if (docText.empty())
        return nullptr;

    const SymbolOccurrence symbol = GetCurrentSymbolOccurrence(hSci, docText);
    if (!symbol.valid)
        return nullptr;

    const TagEntry* best = FindLocalDefinitionAtCaret(grammar, docText, symbol);
    if (best)
        return best;

    if (!grammar->GetTagsQuery())
        return nullptr;

    static std::vector<TagEntry> cachedTags;
    cachedTags = CollectTags(grammar, docText);
    if (cachedTags.empty())
        return nullptr;

    for (const auto& tag : cachedTags) {
        if (tag.name != symbol.name || tag.role != "definition")
            continue;
        if (tag.startByte == symbol.startByte && tag.endByte == symbol.endByte)
            continue;
        if (!best || tag.startByte < best->startByte)
            best = &tag;
    }

    if (!best) {
        for (const auto& tag : cachedTags) {
            if (tag.name == symbol.name && tag.role == "definition") {
                best = &tag;
                break;
            }
        }
    }

    return best;
}

void UpdateDefinitionCommandState()
{
    HWND hSci = GetCurrentScintilla();
    const bool enabled = FindDefinitionAtCaret(hSci) != nullptr;
    SetCommandEnabled(kGoToDefinitionCommandIndex, enabled);
    SetCommandEnabled(kSelectDefinitionCommandIndex, enabled);
}

void UpdateInstallBundledGrammarCommandState()
{
    HWND hSci = GetCurrentScintilla();
    SetCommandEnabled(kInstallBundledGrammarCommandIndex, CanInstallBundledGrammarForCurrentFile(hSci));
}

void UpdateSyntaxCommandStates()
{
    HWND hSci = GetCurrentScintilla();
    const bool hasTreeSitter = !GetActiveTreeSitterLanguage(hSci).empty();
    SetCommandEnabled(kNextDefinitionCommandIndex, hasTreeSitter);
    SetCommandEnabled(kPrevDefinitionCommandIndex, hasTreeSitter);
    SetCommandEnabled(kExpandSelectionCommandIndex, hasTreeSitter);
    SetCommandEnabled(kShrinkSelectionCommandIndex, hasTreeSitter);
    SetCommandEnabled(kShowSymbolPathCommandIndex, hasTreeSitter);
}

// ============================================================================
// Shared parse helper for the interactive commands below
// ============================================================================
struct ParsedDocument {
    TSParser* parser = nullptr;
    TSTree* tree = nullptr;

    ~ParsedDocument()
    {
        if (tree)
            ts_tree_delete(tree);
        if (parser)
            ts_parser_delete(parser);
    }

    bool Parse(const TreeSitterGrammar* grammar, const std::string& docText)
    {
        if (!grammar || !grammar->GetLanguage() || docText.empty())
            return false;

        parser = ts_parser_new();
        if (!parser)
            return false;

        ts_parser_set_language(parser, grammar->GetLanguage());
        tree = ts_parser_parse_string(parser, nullptr, docText.c_str(),
            static_cast<uint32_t>(docText.size()));
        return tree != nullptr;
    }
};

UINT_PTR GetCurrentBufferId()
{
    return static_cast<UINT_PTR>(::SendMessageW(g_nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));
}

// ============================================================================
// Syntax-aware selection expansion (grow to enclosing AST node) and shrink
// ============================================================================
struct SelectionStep {
    UINT_PTR bufferId = 0;
    Sci_Position prevStart = 0;
    Sci_Position prevEnd = 0;
    Sci_Position newStart = 0;
    Sci_Position newEnd = 0;
};

std::vector<SelectionStep> g_selectionHistory;

void expandSyntaxSelection()
{
    HWND hSci = GetCurrentScintilla();
    const TreeSitterGrammar* grammar = GetActiveGrammar(hSci);
    if (!grammar) {
        ShowStatus(L"TreeSitterLexer: current buffer is not using a tree-sitter lexer");
        return;
    }

    const std::string docText = GetDocumentText(hSci);
    ParsedDocument doc;
    if (!doc.Parse(grammar, docText)) {
        ShowStatus(L"TreeSitterLexer: could not parse the current document");
        return;
    }

    auto selStart = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONSTART, 0, 0));
    auto selEnd = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONEND, 0, 0));
    if (selStart < 0 || selEnd < selStart || selEnd > static_cast<Sci_Position>(docText.size()))
        return;

    TSNode root = ts_tree_root_node(doc.tree);
    TSNode node = ts_node_named_descendant_for_byte_range(root,
        static_cast<uint32_t>(selStart), static_cast<uint32_t>(selEnd));

    // Climb past nodes whose range equals the current selection so each
    // invocation grows the selection by one syntactic level.
    while (!ts_node_is_null(node) &&
           static_cast<Sci_Position>(ts_node_start_byte(node)) == selStart &&
           static_cast<Sci_Position>(ts_node_end_byte(node)) == selEnd) {
        node = ts_node_parent(node);
    }

    if (ts_node_is_null(node)) {
        ShowStatus(L"TreeSitterLexer: nothing larger to select");
        return;
    }

    const auto nodeStart = static_cast<Sci_Position>(ts_node_start_byte(node));
    const auto nodeEnd = static_cast<Sci_Position>(ts_node_end_byte(node));

    const UINT_PTR bufferId = GetCurrentBufferId();
    if (!g_selectionHistory.empty() &&
        (g_selectionHistory.back().bufferId != bufferId ||
         g_selectionHistory.back().newStart != selStart ||
         g_selectionHistory.back().newEnd != selEnd)) {
        g_selectionHistory.clear();
    }
    g_selectionHistory.push_back({ bufferId, selStart, selEnd, nodeStart, nodeEnd });

    ::SendMessageW(hSci, SCI_SETSEL, static_cast<WPARAM>(nodeStart), static_cast<LPARAM>(nodeEnd));

    const char* nodeType = ts_node_type(node);
    std::wstring status = L"TreeSitterLexer: selected ";
    if (nodeType) {
        const std::string typeName(nodeType);
        status.append(typeName.begin(), typeName.end());
    }
    ShowStatus(status);
}

void shrinkSyntaxSelection()
{
    HWND hSci = GetCurrentScintilla();
    if (!hSci)
        return;

    const auto selStart = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONSTART, 0, 0));
    const auto selEnd = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONEND, 0, 0));

    if (g_selectionHistory.empty() ||
        g_selectionHistory.back().bufferId != GetCurrentBufferId() ||
        g_selectionHistory.back().newStart != selStart ||
        g_selectionHistory.back().newEnd != selEnd) {
        g_selectionHistory.clear();
        ShowStatus(L"TreeSitterLexer: no selection expansion to undo");
        return;
    }

    const SelectionStep step = g_selectionHistory.back();
    g_selectionHistory.pop_back();
    ::SendMessageW(hSci, SCI_SETSEL, static_cast<WPARAM>(step.prevStart), static_cast<LPARAM>(step.prevEnd));
    ShowStatus(L"TreeSitterLexer: selection shrunk");
}

// ============================================================================
// Symbol breadcrumb: show the enclosing definition path at the caret
// (e.g. "MyClass.MyMethod"). Ported from WinMerge's GetEnclosingSymbols.
// ============================================================================

// Heuristic: does this AST node type represent a named definition?
// Matches the node-type names used across tree-sitter grammars for functions,
// methods, types and modules (e.g. "function_definition", "class_specifier",
// "impl_item", "namespace_definition"); excludes call/invocation nodes.
static bool IsDefinitionLikeNodeType(const char* type)
{
    if (!type || strstr(type, "call") != nullptr)
        return false;
    static const char* const keywords[] = {
        "function", "method", "class", "struct", "interface",
        "namespace", "module", "impl", "trait", "enum", "constructor",
    };
    for (const char* kw : keywords) {
        if (strstr(type, kw) != nullptr)
            return true;
    }
    return false;
}

void showSymbolPath()
{
    HWND hSci = GetCurrentScintilla();
    const TreeSitterGrammar* grammar = GetActiveGrammar(hSci);
    if (!grammar) {
        ShowStatus(L"TreeSitterLexer: current buffer is not using a tree-sitter lexer");
        return;
    }

    const std::string docText = GetDocumentText(hSci);
    ParsedDocument doc;
    if (!doc.Parse(grammar, docText)) {
        ShowStatus(L"TreeSitterLexer: could not parse the current document");
        return;
    }

    auto caret = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETCURRENTPOS, 0, 0));
    if (caret < 0 || caret > static_cast<Sci_Position>(docText.size()))
        return;

    const uint32_t byteOffset = static_cast<uint32_t>(caret);
    TSNode node = ts_node_descendant_for_byte_range(
        ts_tree_root_node(doc.tree), byteOffset, byteOffset);

    // Collect enclosing definition names, innermost first. Requiring a "body"
    // field filters out references/invocations and forward declarations that
    // match the type-name heuristic but are not real definitions.
    std::vector<std::string> names;
    while (!ts_node_is_null(node)) {
        if (IsDefinitionLikeNodeType(ts_node_type(node)) &&
            !ts_node_is_null(ts_node_child_by_field_name(node, "body", 4))) {
            // Most grammars expose the symbol as a "name" field; C/C++ use a
            // (possibly nested) "declarator" field instead.
            TSNode nameNode = ts_node_child_by_field_name(node, "name", 4);
            if (ts_node_is_null(nameNode)) {
                TSNode declNode = ts_node_child_by_field_name(node, "declarator", 10);
                while (!ts_node_is_null(declNode)) {
                    TSNode inner = ts_node_child_by_field_name(declNode, "declarator", 10);
                    if (ts_node_is_null(inner))
                        break;
                    declNode = inner;
                }
                nameNode = declNode;
            }
            if (!ts_node_is_null(nameNode)) {
                const uint32_t s = ts_node_start_byte(nameNode);
                const uint32_t e = ts_node_end_byte(nameNode);
                if (s < e && e <= docText.size() && e - s <= 256)
                    names.push_back(docText.substr(s, e - s));
            }
        }
        node = ts_node_parent(node);
    }

    if (names.empty()) {
        ShowStatus(L"TreeSitterLexer: not inside a named definition");
        return;
    }

    // Join outermost -> innermost, keeping at most the 3 innermost names.
    const size_t nCount = names.size() < 3 ? names.size() : 3;
    std::string joined;
    for (size_t i = nCount; i-- > 0; ) {
        if (!joined.empty())
            joined += '.';
        joined += names[i];
    }

    std::wstring status = L"TreeSitterLexer: ";
    status.append(joined.begin(), joined.end());  // ASCII-ish symbol names
    ShowStatus(status);
}

// ============================================================================
// Symbol occurrence highlighting - marks identifiers that share both the
// text and the syntax node type of the symbol under the caret, so matches
// inside strings and comments are not highlighted.
// ============================================================================

// Last applied highlight state; caret moves that keep the same symbol in an
// unmodified document skip the reparse entirely.
struct OccurrenceCache {
    LRESULT docPointer = 0;
    std::string name;
    uint32_t startByte = 0;
    uint32_t endByte = 0;
    bool valid = false;
};

OccurrenceCache g_occurrenceCache;
bool g_occurrenceTextDirty = false;

void EnsureOccurrenceIndicatorStyle(HWND hSci)
{
    static std::unordered_set<HWND> styledViews;
    if (!styledViews.insert(hSci).second)
        return;

    ::SendMessageW(hSci, SCI_INDICSETSTYLE, kOccurrenceIndicator, INDIC_ROUNDBOX);
    ::SendMessageW(hSci, SCI_INDICSETFORE, kOccurrenceIndicator, RGB(0x33, 0x99, 0xFF));
    ::SendMessageW(hSci, SCI_INDICSETALPHA, kOccurrenceIndicator, 60);
    ::SendMessageW(hSci, SCI_INDICSETOUTLINEALPHA, kOccurrenceIndicator, 150);
    ::SendMessageW(hSci, SCI_INDICSETUNDER, kOccurrenceIndicator, TRUE);
}

void ClearOccurrenceHighlights(HWND hSci)
{
    g_occurrenceCache.valid = false;
    if (!hSci)
        return;

    const auto docLen = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETLENGTH, 0, 0));
    ::SendMessageW(hSci, SCI_SETINDICATORCURRENT, kOccurrenceIndicator, 0);
    ::SendMessageW(hSci, SCI_INDICATORCLEARRANGE, 0, static_cast<LPARAM>(docLen));
}

void UpdateOccurrenceHighlights(HWND hSci)
{
    if (!hSci)
        return;

    if (!g_highlightOccurrencesEnabled) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    const TreeSitterGrammar* grammar = GetActiveGrammar(hSci);
    if (!grammar) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    const auto docLen = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETLENGTH, 0, 0));
    if (docLen > kMaxOccurrenceDocBytes) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    const std::string docText = GetDocumentText(hSci);
    const SymbolOccurrence symbol = GetCurrentSymbolOccurrence(hSci, docText);
    if (!symbol.valid || symbol.endByte > docText.size()) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    const auto docPointer = static_cast<LRESULT>(::SendMessageW(hSci, SCI_GETDOCPOINTER, 0, 0));
    if (g_occurrenceCache.valid && !g_occurrenceTextDirty &&
        g_occurrenceCache.docPointer == docPointer &&
        g_occurrenceCache.startByte == symbol.startByte &&
        g_occurrenceCache.endByte == symbol.endByte &&
        g_occurrenceCache.name == symbol.name) {
        return;
    }

    ParsedDocument doc;
    if (!doc.Parse(grammar, docText)) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    TSNode root = ts_tree_root_node(doc.tree);
    TSNode caretNode = ts_node_named_descendant_for_byte_range(root, symbol.startByte, symbol.endByte);

    // Only highlight when the caret symbol is itself a syntax node
    // (i.e. not a word inside a string or comment).
    if (ts_node_is_null(caretNode) ||
        ts_node_start_byte(caretNode) != symbol.startByte ||
        ts_node_end_byte(caretNode) != symbol.endByte) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    const char* caretType = ts_node_type(caretNode);
    if (!caretType) {
        ClearOccurrenceHighlights(hSci);
        return;
    }

    const uint32_t symbolLen = symbol.endByte - symbol.startByte;

    std::vector<std::pair<uint32_t, uint32_t>> matches;
    std::vector<TSNode> stack;
    stack.push_back(root);
    while (!stack.empty()) {
        TSNode node = stack.back();
        stack.pop_back();

        const uint32_t nodeStart = ts_node_start_byte(node);
        const uint32_t nodeEnd = ts_node_end_byte(node);
        if (nodeEnd - nodeStart < symbolLen)
            continue;

        if (nodeEnd - nodeStart == symbolLen && ts_node_is_named(node) &&
            std::strcmp(ts_node_type(node), caretType) == 0 &&
            docText.compare(nodeStart, symbolLen, symbol.name) == 0) {
            matches.emplace_back(nodeStart, nodeEnd);
        }

        const uint32_t childCount = ts_node_child_count(node);
        for (uint32_t i = 0; i < childCount; ++i)
            stack.push_back(ts_node_child(node, i));
    }

    EnsureOccurrenceIndicatorStyle(hSci);
    ::SendMessageW(hSci, SCI_SETINDICATORCURRENT, kOccurrenceIndicator, 0);
    ::SendMessageW(hSci, SCI_INDICATORCLEARRANGE, 0, static_cast<LPARAM>(docLen));
    for (const auto& m : matches) {
        ::SendMessageW(hSci, SCI_INDICATORFILLRANGE, static_cast<WPARAM>(m.first),
            static_cast<LPARAM>(m.second - m.first));
    }

    g_occurrenceCache.docPointer = docPointer;
    g_occurrenceCache.name = symbol.name;
    g_occurrenceCache.startByte = symbol.startByte;
    g_occurrenceCache.endByte = symbol.endByte;
    g_occurrenceCache.valid = true;
    g_occurrenceTextDirty = false;
}

void toggleHighlightOccurrences()
{
    g_highlightOccurrencesEnabled = !g_highlightOccurrencesEnabled;
    SetMenuItemChecked(kHighlightOccurrencesCommandIndex, g_highlightOccurrencesEnabled);

    if (g_highlightOccurrencesEnabled) {
        UpdateOccurrenceHighlights(GetCurrentScintilla());
        ShowStatus(L"TreeSitterLexer: symbol occurrence highlighting enabled");
    } else {
        // Indicators live in the documents shown in either view; clear both.
        ClearOccurrenceHighlights(g_nppData._scintillaMainHandle);
        ClearOccurrenceHighlights(g_nppData._scintillaSecondHandle);
        ShowStatus(L"TreeSitterLexer: symbol occurrence highlighting disabled");
    }
}

// ============================================================================
// Definition navigation - jump between tagged definitions (functions,
// methods, classes, ...) from the grammar's tags query.
// ============================================================================
void NavigateDefinition(bool forward)
{
    HWND hSci = GetCurrentScintilla();
    const TreeSitterGrammar* grammar = GetActiveGrammar(hSci);
    if (!grammar) {
        ShowStatus(L"TreeSitterLexer: current buffer is not using a tree-sitter lexer");
        return;
    }
    if (!grammar->GetTagsQuery()) {
        ShowStatus(L"TreeSitterLexer: no tags query available for this language");
        return;
    }

    const std::string docText = GetDocumentText(hSci);
    std::vector<TagEntry> definitions;
    for (auto& tag : CollectTags(grammar, docText)) {
        if (tag.role == "definition")
            definitions.push_back(std::move(tag));
    }
    if (definitions.empty()) {
        ShowStatus(L"TreeSitterLexer: no definitions found in current buffer");
        return;
    }

    std::sort(definitions.begin(), definitions.end(),
        [](const TagEntry& a, const TagEntry& b) { return a.startByte < b.startByte; });

    const auto caretPos = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETCURRENTPOS, 0, 0));

    // First definition strictly after the caret; everything before that
    // (exclusive of definitions at the caret itself) is "behind" the caret.
    const auto firstAfter = std::upper_bound(definitions.begin(), definitions.end(), caretPos,
        [](Sci_Position pos, const TagEntry& def) {
            return pos < static_cast<Sci_Position>(def.startByte);
        });
    const auto firstAtOrAfter = std::lower_bound(definitions.begin(), definitions.end(), caretPos,
        [](const TagEntry& def, Sci_Position pos) {
            return static_cast<Sci_Position>(def.startByte) < pos;
        });

    size_t index = 0;
    bool wrapped = false;
    if (forward) {
        wrapped = firstAfter == definitions.end();
        index = wrapped ? 0 : static_cast<size_t>(firstAfter - definitions.begin());
    } else {
        wrapped = firstAtOrAfter == definitions.begin();
        index = (wrapped ? definitions.size() : static_cast<size_t>(firstAtOrAfter - definitions.begin())) - 1;
    }
    const TagEntry* target = &definitions[index];

    NavigateToPosition(hSci, static_cast<Sci_Position>(target->startByte));

    std::wstring status = L"TreeSitterLexer: ";
    if (!target->kind.empty()) {
        status.append(target->kind.begin(), target->kind.end());
        status += L' ';
    }
    status.append(target->name.begin(), target->name.end());
    if (wrapped)
        status += L" (wrapped)";
    ShowStatus(status);
}

void goToNextDefinition() {
    NavigateDefinition(true);
}

void goToPreviousDefinition() {
    NavigateDefinition(false);
}

void autoDetectTreeSitterLanguage()
{
    HWND hSci = GetCurrentScintilla();
    if (!hSci) {
        ShowStatus(L"TreeSitterLexer: no active editor");
        return;
    }

    const std::wstring path = GetCurrentFilePath();
    const std::string language = DetectTreeSitterLanguageForFile(path);
    if (language.empty()) {
        ShowStatus(L"TreeSitterLexer: no tree-sitter lexer mapped for this extension");
        return;
    }

    const std::wstring extension = GetFileExtensionLower(path);

    TreeSitterRegistry::Instance().Initialize(g_hModule);
    if (!TreeSitterRegistry::Instance().GetGrammarByName(language)) {
        ShowStatus(L"TreeSitterLexer: mapped grammar is not available");
        return;
    }

    const std::string currentLexer = GetCurrentLexerName(hSci);
    const std::string targetLexer = "treesitter." + language;
    if (ToLowerAscii(currentLexer) == ToLowerAscii(targetLexer)) {
        ShowStatus(L"TreeSitterLexer: matching tree-sitter lexer already active");
        return;
    }

    if (!IsGenericLexerName(currentLexer) && PreferBuiltInLexerForExtension(extension)) {
        ShowStatus(L"TreeSitterLexer: keeping the current built-in lexer for this file type");
        return;
    }

    UINT commandId = 0;
    HMENU mainMenu = reinterpret_cast<HMENU>(::SendMessageW(g_nppData._nppHandle, NPPM_GETMENUHANDLE,
        static_cast<WPARAM>(NPPMAINMENU), 0));
    std::wstring targetMenuText(targetLexer.begin(), targetLexer.end());
    if (!TryFindMenuCommandByText(mainMenu, targetMenuText, commandId)) {
        ShowStatus(L"TreeSitterLexer: could not find the language menu entry");
        return;
    }

    if (!::SendMessageW(g_nppData._nppHandle, NPPM_MENUCOMMAND, 0, static_cast<LPARAM>(commandId))) {
        ShowStatus(L"TreeSitterLexer: failed to switch language");
        return;
    }

    std::wstring status = L"TreeSitterLexer: auto-detected ";
    status.append(targetMenuText.begin(), targetMenuText.end());
    ShowStatus(status);

    UpdateDefinitionCommandState();
}

void toggleAutoDetectTreeSitterLanguage()
{
    g_autoDetectEnabled = !g_autoDetectEnabled;
    SetMenuItemChecked(kAutoDetectCommandIndex, g_autoDetectEnabled);

    if (g_autoDetectEnabled) {
        autoDetectTreeSitterLanguage();
    } else {
        ShowStatus(L"TreeSitterLexer: auto-detect disabled");
    }
}

void installMissingBundledGrammar()
{
    HWND hSci = GetCurrentScintilla();
    if (!hSci) {
        ShowStatus(L"TreeSitterLexer: no active editor");
        return;
    }

    const std::wstring path = GetCurrentFilePath();
    const std::string language = DetectTreeSitterLanguageForFile(path);
    if (language.empty()) {
        ShowStatus(L"TreeSitterLexer: no bundled grammar mapping for this file");
        ShowInstallResultDialog(
            L"TreeSitterLexer",
            L"No bundled Tree-sitter grammar is mapped for this file type.\n\n"
            L"For example, `.rst` is reStructuredText, not Rust. Rust files are detected from `.rs`.");
        return;
    }

    TreeSitterRegistry::Instance().Initialize(g_hModule);
    if (TreeSitterRegistry::Instance().GetGrammarByName(language)) {
        if (ActivateTreeSitterLexer(hSci, language)) {
            std::wstring status = L"TreeSitterLexer: grammar already available -> treesitter.";
            status.append(language.begin(), language.end());
            ShowStatus(status);
            std::wstring message = L"The grammar is already installed and has been applied:\n\n treesitter.";
            message.append(language.begin(), language.end());
            ShowInstallResultDialog(L"TreeSitterLexer", message);
            UpdateDefinitionCommandState();
            UpdateInstallBundledGrammarCommandState();
            return;
        }
    }

    if (!TryInstallBundledGrammar(language)) {
        ShowStatus(L"TreeSitterLexer: bundled grammar build output not found for this file type");
        std::wstring message =
            L"The file type is recognized, but no locally built bundled grammar package was found to install.\n\n"
            L"Expected local build output for: treesitter.";
        message.append(language.begin(), language.end());
        ShowInstallResultDialog(L"TreeSitterLexer", message, MB_OK | MB_ICONWARNING);
        return;
    }

    if (!ActivateTreeSitterLexer(hSci, language)) {
        ShowStatus(L"TreeSitterLexer: grammar copied but activation failed; restart Notepad++ if needed");
        std::wstring message =
            L"The grammar files were copied, but the lexer could not be activated immediately.\n\n"
            L"Try restarting Notepad++ and selecting treesitter.";
        message.append(language.begin(), language.end());
        ShowInstallResultDialog(L"TreeSitterLexer", message, MB_OK | MB_ICONWARNING);
        return;
    }

    std::wstring status = L"TreeSitterLexer: installed bundled grammar -> treesitter.";
    status.append(language.begin(), language.end());
    ShowStatus(status);
    std::wstring message = L"Installed and activated bundled grammar:\n\n treesitter.";
    message.append(language.begin(), language.end());
    ShowInstallResultDialog(L"TreeSitterLexer", message);
    UpdateDefinitionCommandState();
    UpdateInstallBundledGrammarCommandState();
}

void GoToDefinitionImpl(bool preferReference)
{
    HWND hSci = GetCurrentScintilla();
    const TagEntry* best = FindDefinitionAtCaret(hSci);
    if (!best) {
        const std::string language = GetActiveTreeSitterLanguage(hSci);
        if (language.empty()) {
            ShowStatus(L"TreeSitterLexer: current buffer is not using a tree-sitter lexer");
        } else {
            ShowStatus(L"TreeSitterLexer: definition not found in current buffer");
        }
        return;
    }

    const std::string symbol = GetSelectedOrCurrentWord(hSci);
    if (symbol.empty()) {
        ShowStatus(L"TreeSitterLexer: current buffer is not using a tree-sitter lexer");
        return;
    }

    NavigateToPosition(hSci, static_cast<Sci_Position>(best->startByte));
    std::wstring status = L"TreeSitterLexer: go to definition -> ";
    status.append(symbol.begin(), symbol.end());
    ShowStatus(status);

    UpdateDefinitionCommandState();
}

void selectCurrentSymbolDefinition() {
    GoToDefinitionImpl(false);
}

void goToCurrentSymbolDefinition() {
    GoToDefinitionImpl(true);
}

} // namespace

// Dummy menu entry - N++ requires at least 1 FuncItem
static void aboutDlg()
{
    // Show version (read from the embedded resource) and a clickable repo link.
    // Uses a TaskDialog so the hyperlink works; AboutDialogCallback opens the
    // URL via ShellExecute on TDN_HYPERLINK_CLICKED. Falls back to a plain
    // MessageBox if the TaskDialog is unavailable.
    const std::wstring version = GetModuleVersionString();
    std::wstring mainInstruction = L"TreeSitterLexer";
    if (!version.empty())
        mainInstruction += L" " + version;

    static const wchar_t* const kRepoUrl = L"https://github.com/Thorium/NppTreeSitter";
    const std::wstring content =
        L"Tree-sitter syntax highlighting for Notepad++.\n"
        L"Provides external lexers powered by tree-sitter grammars.\n\n"
        L"<a href=\"" + std::wstring(kRepoUrl) + L"\">" + kRepoUrl + L"</a>";

    TASKDIALOGCONFIG config = {};
    config.cbSize = sizeof(config);
    config.hwndParent = g_nppData._nppHandle;
    config.hInstance = g_hModule;
    config.dwFlags = TDF_ENABLE_HYPERLINKS | TDF_ALLOW_DIALOG_CANCELLATION;
    config.dwCommonButtons = TDCBF_OK_BUTTON;
    config.pszWindowTitle = L"About TreeSitterLexer";
    config.pszMainIcon = TD_INFORMATION_ICON;
    config.pszMainInstruction = mainInstruction.c_str();
    config.pszContent = content.c_str();
    config.pfCallback = AboutDialogCallback;

    if (FAILED(TaskDialogIndirect(&config, nullptr, nullptr, nullptr))) {
        const std::wstring fallback =
            L"Tree-sitter syntax highlighting for Notepad++\n\n" + std::wstring(kRepoUrl);
        ::MessageBoxW(g_nppData._nppHandle, fallback.c_str(),
            mainInstruction.c_str(), MB_OK | MB_ICONINFORMATION);
    }
}

// Default shortcuts (users can remap them in the Shortcut Mapper)
static ShortcutKey g_nextDefinitionKey   = { true, true, false, VK_DOWN };
static ShortcutKey g_prevDefinitionKey   = { true, true, false, VK_UP };
static ShortcutKey g_expandSelectionKey  = { true, true, false, 'W' };
static ShortcutKey g_shrinkSelectionKey  = { true, true, true,  'W' };

// A null _pFunc renders as a menu separator in Notepad++
FuncItem g_funcItems[] = {
    { L"Auto-detect Tree-sitter Language", toggleAutoDetectTreeSitterLanguage, 0, false, nullptr },
    { L"Install Missing Bundled Grammar", installMissingBundledGrammar, 0, false, nullptr },
    { L"---", nullptr, 0, false, nullptr },
    { L"Go to Definition", goToCurrentSymbolDefinition, 0, false, nullptr },
    { L"Select Definition", selectCurrentSymbolDefinition, 0, false, nullptr },
    { L"Go to Next Definition", goToNextDefinition, 0, false, &g_nextDefinitionKey },
    { L"Go to Previous Definition", goToPreviousDefinition, 0, false, &g_prevDefinitionKey },
    { L"---", nullptr, 0, false, nullptr },
    { L"Expand Selection", expandSyntaxSelection, 0, false, &g_expandSelectionKey },
    { L"Shrink Selection", shrinkSyntaxSelection, 0, false, &g_shrinkSelectionKey },
    { L"Highlight Symbol Occurrences", toggleHighlightOccurrences, 0, true, nullptr },
    { L"Show Symbol Path", showSymbolPath, 0, false, nullptr },
    { L"---", nullptr, 0, false, nullptr },
    { L"About TreeSitterLexer", aboutDlg, 0, false, nullptr }
};

BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID /*reserved*/)
{
    if (reason == DLL_PROCESS_ATTACH) {
        g_hModule = hModule;
        DisableThreadLibraryCalls(hModule);
    }
    return TRUE;
}

// ============================================================================
// Notepad++ Plugin Interface exports (__cdecl)
// These must be present or N++ will reject the DLL before even checking
// for Lexilla exports.
// ============================================================================

extern "C" {

__declspec(dllexport) BOOL __cdecl isUnicode()
{
    return TRUE;
}

__declspec(dllexport) void __cdecl setInfo(NppData nppData)
{
    g_nppData = nppData;
}

__declspec(dllexport) const wchar_t* __cdecl getName()
{
    return L"TreeSitterLexer";
}

__declspec(dllexport) void __cdecl beNotified(SCNotification* notifyCode)
{
    if (!notifyCode)
        return;

    switch (notifyCode->nmhdr.code) {
    case NPPN_READY:
        SetMenuItemChecked(kAutoDetectCommandIndex, g_autoDetectEnabled);
        SetMenuItemChecked(kHighlightOccurrencesCommandIndex, g_highlightOccurrencesEnabled);
        if (g_autoDetectEnabled)
            autoDetectTreeSitterLanguage();
        UpdateDefinitionCommandState();
        UpdateInstallBundledGrammarCommandState();
        UpdateSyntaxCommandStates();
        break;
    case NPPN_BUFFERACTIVATED:
        if (g_autoDetectEnabled)
            autoDetectTreeSitterLanguage();
        UpdateDefinitionCommandState();
        UpdateInstallBundledGrammarCommandState();
        UpdateSyntaxCommandStates();
        UpdateOccurrenceHighlights(GetCurrentScintilla());
        break;
    case NPPN_LANGCHANGED:
        UpdateDefinitionCommandState();
        UpdateSyntaxCommandStates();
        UpdateOccurrenceHighlights(GetCurrentScintilla());
        break;
    case SCN_MODIFIED:
        if (notifyCode->modificationType & (SC_MOD_INSERTTEXT | SC_MOD_DELETETEXT))
            g_occurrenceTextDirty = true;
        break;
    case SCN_UPDATEUI:
        if (notifyCode->updated & SC_UPDATE_SELECTION) {
            UpdateDefinitionCommandState();
            UpdateInstallBundledGrammarCommandState();
            // Use the view that fired the notification, not the focused one;
            // in split view they can differ.
            UpdateOccurrenceHighlights(notifyCode->nmhdr.hwndFrom);
        }
        break;
    }
}

__declspec(dllexport) LRESULT __cdecl messageProc(UINT /*Message*/,
                                                   WPARAM /*wParam*/,
                                                   LPARAM /*lParam*/)
{
    return TRUE;
}

__declspec(dllexport) FuncItem* __cdecl getFuncsArray(int* nbF)
{
    *nbF = kFuncItemCount;
    return g_funcItems;
}

// ============================================================================
// Lexilla API exports (__stdcall)
// These are what N++ uses to discover and create external lexers,
// checked in PluginsManager::loadPlugins() AFTER the plugin interface.
// ============================================================================

__declspec(dllexport) int __stdcall GetLexerCount()
{
    if (kDisableExternalLexersForStartupIsolation)
        return 0;

    // Lazy init on first call
    TreeSitterRegistry::Instance().Initialize(g_hModule);
    return TreeSitterRegistry::Instance().GetLexerCount();
}

__declspec(dllexport) void __stdcall GetLexerName(unsigned int index,
                                                   char* name, int bufLength)
{
    if (kDisableExternalLexersForStartupIsolation) {
        if (name && bufLength > 0)
            name[0] = '\0';
        return;
    }

    TreeSitterRegistry::Instance().GetLexerName(index, name, bufLength);
}

__declspec(dllexport) ILexer5* __stdcall CreateLexer(const char* name)
{
    if (kDisableExternalLexersForStartupIsolation)
        return nullptr;

    return TreeSitterRegistry::Instance().CreateLexer(name);
}

} // extern "C"

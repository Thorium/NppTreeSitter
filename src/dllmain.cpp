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

constexpr int SC_UPDATE_SELECTION = 0x2;
constexpr UINT SCN_UPDATEUI = 2007;

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
static const int kFuncItemCount = 5;

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

namespace {

constexpr int kAutoDetectCommandIndex = 0;
constexpr int kGoToDefinitionCommandIndex = 1;
constexpr int kSelectDefinitionCommandIndex = 2;
constexpr int kInstallBundledGrammarCommandIndex = 3;

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

SymbolOccurrence GetCurrentSymbolOccurrence(HWND hSci)
{
    SymbolOccurrence occurrence;
    if (!hSci)
        return occurrence;

    auto selStart = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONSTART, 0, 0));
    auto selEnd = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETSELECTIONEND, 0, 0));
    const auto docLen = static_cast<Sci_Position>(::SendMessageW(hSci, SCI_GETLENGTH, 0, 0));
    if (docLen <= 0)
        return occurrence;

    std::string doc = GetDocumentText(hSci);
    if (doc.empty())
        return occurrence;

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

void UpdateAutoDetectCheck()
{
    if (g_funcItems[kAutoDetectCommandIndex]._cmdID == 0)
        return;

    ::SendMessageW(
        g_nppData._nppHandle,
        NPPM_SETMENUITEMCHECK,
        static_cast<WPARAM>(g_funcItems[kAutoDetectCommandIndex]._cmdID),
        static_cast<LPARAM>(g_autoDetectEnabled ? TRUE : FALSE));
}

const TagEntry* FindDefinitionAtCaret(HWND hSci)
{
    if (!hSci)
        return nullptr;

    const std::string language = GetActiveTreeSitterLanguage(hSci);
    if (language.empty())
        return nullptr;

    TreeSitterRegistry::Instance().Initialize(g_hModule);
    const TreeSitterGrammar* grammar = TreeSitterRegistry::Instance().GetGrammarByName(language);
    if (!grammar)
        return nullptr;

    const SymbolOccurrence symbol = GetCurrentSymbolOccurrence(hSci);
    if (!symbol.valid)
        return nullptr;

    const std::string docText = GetDocumentText(hSci);
    if (docText.empty())
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
    UpdateAutoDetectCheck();

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
    ::MessageBoxW(g_nppData._nppHandle,
        L"TreeSitterLexer - Tree-sitter syntax highlighting for Notepad++\n\n"
        L"Provides external lexers powered by tree-sitter grammars.",
        L"TreeSitterLexer",
        MB_OK | MB_ICONINFORMATION);
}

FuncItem g_funcItems[] = {
    { L"Auto-detect Tree-sitter Language", toggleAutoDetectTreeSitterLanguage, 0, false, nullptr },
    { L"Go to Definition", goToCurrentSymbolDefinition, 0, false, nullptr },
    { L"Select Definition", selectCurrentSymbolDefinition, 0, false, nullptr },
    { L"Install Missing Bundled Grammar", installMissingBundledGrammar, 0, false, nullptr },
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
        UpdateAutoDetectCheck();
        if (g_autoDetectEnabled)
            autoDetectTreeSitterLanguage();
        UpdateDefinitionCommandState();
        UpdateInstallBundledGrammarCommandState();
        break;
    case NPPN_BUFFERACTIVATED:
        if (g_autoDetectEnabled)
            autoDetectTreeSitterLanguage();
        UpdateDefinitionCommandState();
        UpdateInstallBundledGrammarCommandState();
        break;
    case SCN_UPDATEUI:
        if (notifyCode->updated & SC_UPDATE_SELECTION) {
            UpdateDefinitionCommandState();
            UpdateInstallBundledGrammarCommandState();
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

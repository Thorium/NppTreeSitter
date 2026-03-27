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

using Scintilla::ILexer5;

// Forward declarations - avoid pulling in full N++ headers
struct SCNotification;

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

struct FuncItem
{
    wchar_t _itemName[menuItemSize];
    PFUNCPLUGINCMD _pFunc;
    int _cmdID;
    bool _init2Check;
    ShortcutKey *_pShKey;
};

// ============================================================================
// Globals
// ============================================================================
static HMODULE g_hModule = nullptr;
static NppData g_nppData = {};

// Dummy menu entry - N++ requires at least 1 FuncItem
static void aboutDlg()
{
    ::MessageBoxW(g_nppData._nppHandle,
        L"TreeSitterLexer - Tree-sitter syntax highlighting for Notepad++\n\n"
        L"Provides external lexers powered by tree-sitter grammars.",
        L"TreeSitterLexer",
        MB_OK | MB_ICONINFORMATION);
}

static FuncItem g_funcItems[] = {
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

__declspec(dllexport) void __cdecl beNotified(SCNotification* /*notifyCode*/)
{
    // Nothing to handle
}

__declspec(dllexport) LRESULT __cdecl messageProc(UINT /*Message*/,
                                                   WPARAM /*wParam*/,
                                                   LPARAM /*lParam*/)
{
    return TRUE;
}

__declspec(dllexport) FuncItem* __cdecl getFuncsArray(int* nbF)
{
    *nbF = sizeof(g_funcItems) / sizeof(g_funcItems[0]);
    return g_funcItems;
}

// ============================================================================
// Lexilla API exports (__stdcall)
// These are what N++ uses to discover and create external lexers,
// checked in PluginsManager::loadPlugins() AFTER the plugin interface.
// ============================================================================

__declspec(dllexport) int __stdcall GetLexerCount()
{
    // Lazy init on first call
    TreeSitterRegistry::Instance().Initialize(g_hModule);
    return TreeSitterRegistry::Instance().GetLexerCount();
}

__declspec(dllexport) void __stdcall GetLexerName(unsigned int index,
                                                   char* name, int bufLength)
{
    TreeSitterRegistry::Instance().GetLexerName(index, name, bufLength);
}

__declspec(dllexport) ILexer5* __stdcall CreateLexer(const char* name)
{
    return TreeSitterRegistry::Instance().CreateLexer(name);
}

} // extern "C"

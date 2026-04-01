// TreeSitterLexer - ILexer5 bridge for tree-sitter grammars in Notepad++
//
// Loads tree-sitter grammar DLLs and highlight queries at runtime,
// wraps them in Scintilla's ILexer5 interface for use as an external
// lexer plugin in Notepad++.
//
// Grammar DLLs are discovered in a "TreeSitterGrammars" directory
// alongside the plugin DLL. Each grammar DLL exports tree_sitter_<lang>()
// and has a companion <lang>-highlights.scm query file.
//
// SPDX-License-Identifier: GPL-2.0-or-later

#pragma once

#include <stdint.h>

#include "ILexer.h"
#include "Sci_Position.h"

#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <memory>

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

// Forward declarations for tree-sitter C API types
typedef struct TSParser TSParser;
typedef struct TSTree TSTree;
typedef struct TSQuery TSQuery;
typedef struct TSLanguage TSLanguage;

// ============================================================================
// Style IDs (0-15) assigned to Scintilla by our lexer.
// These are referenced in the companion XML config for colors/fonts.
// ============================================================================
namespace TSStyle {
    enum : char {
        Default       = 0,
        Comment       = 1,
        CommentDoc    = 2,
        Keyword       = 3,
        String        = 4,
        Number        = 5,
        Operator      = 6,
        Function      = 7,
        Type          = 8,
        Preprocessor  = 9,
        Variable      = 10,
        Constant      = 11,
        Punctuation   = 12,
        Tag           = 13,
        Count         = 14,
    };
}

// Style metadata for ILexer5::NameOfStyle / TagsOfStyle / DescriptionOfStyle
struct TSStyleInfo {
    const char* name;
    const char* tags;
    const char* description;
};

extern const TSStyleInfo kStyleInfo[TSStyle::Count];

// ============================================================================
// TreeSitterGrammar - Loads a grammar DLL and its highlight query.
// ============================================================================
class TreeSitterGrammar {
public:
    TreeSitterGrammar();
    ~TreeSitterGrammar();

    TreeSitterGrammar(const TreeSitterGrammar&) = delete;
    TreeSitterGrammar& operator=(const TreeSitterGrammar&) = delete;

    bool Load(const std::wstring& grammarDir, const std::wstring& language);

    const TSLanguage* GetLanguage() const { return m_pLanguage; }
    const TSQuery* GetHighlightQuery() const { return m_pQuery; }
    const TSQuery* GetLocalsQuery() const { return m_pLocalsQuery; }
    const TSQuery* GetInjectionQuery() const { return m_pInjectionQuery; }
    const std::string& GetName() const { return m_name; }

private:
    TSQuery* LoadQuery(const std::wstring& path);

    HMODULE           m_hDll;
    const TSLanguage* m_pLanguage;
    TSQuery*          m_pQuery;
    TSQuery*          m_pLocalsQuery;
    TSQuery*          m_pInjectionQuery;
    std::string       m_name;    // ASCII language name (e.g. "python")
};

// ============================================================================
// TreeSitterStyleMap - Maps tree-sitter capture names to TSStyle IDs.
// ============================================================================
class TreeSitterStyleMap {
public:
    TreeSitterStyleMap();

    char MapCapture(const std::string& captureName) const;

private:
    std::unordered_map<std::string, char> m_map;
};

// ============================================================================
// TreeSitterILexer - ILexer5 implementation wrapping a tree-sitter grammar.
//
// One instance per Scintilla document. Created by CreateLexer().
// Scintilla calls Lex() when text changes; we parse with tree-sitter
// and apply styles via IDocument::StartStyling/SetStyleFor.
// ============================================================================
class TreeSitterILexer : public Scintilla::ILexer5 {
public:
    explicit TreeSitterILexer(const TreeSitterGrammar* grammar);
    ~TreeSitterILexer();

    // ILexer5 interface
    int SCI_METHOD Version() const override;
    void SCI_METHOD Release() override;
    const char* SCI_METHOD PropertyNames() override;
    int SCI_METHOD PropertyType(const char* name) override;
    const char* SCI_METHOD DescribeProperty(const char* name) override;
    Sci_Position SCI_METHOD PropertySet(const char* key, const char* val) override;
    const char* SCI_METHOD DescribeWordListSets() override;
    Sci_Position SCI_METHOD WordListSet(int n, const char* wl) override;
    void SCI_METHOD Lex(Sci_PositionU startPos, Sci_Position lengthDoc,
                        int initStyle, Scintilla::IDocument* pAccess) override;
    void SCI_METHOD Fold(Sci_PositionU startPos, Sci_Position lengthDoc,
                         int initStyle, Scintilla::IDocument* pAccess) override;
    void* SCI_METHOD PrivateCall(int operation, void* pointer) override;
    int SCI_METHOD LineEndTypesSupported() override;
    int SCI_METHOD AllocateSubStyles(int styleBase, int numberStyles) override;
    int SCI_METHOD SubStylesStart(int styleBase) override;
    int SCI_METHOD SubStylesLength(int styleBase) override;
    int SCI_METHOD StyleFromSubStyle(int subStyle) override;
    int SCI_METHOD PrimaryStyleFromStyle(int style) override;
    void SCI_METHOD FreeSubStyles() override;
    void SCI_METHOD SetIdentifiers(int style, const char* identifiers) override;
    int SCI_METHOD DistanceToSecondaryStyles() override;
    const char* SCI_METHOD GetSubStyleBases() override;
    int SCI_METHOD NamedStyles() override;
    const char* SCI_METHOD NameOfStyle(int style) override;
    const char* SCI_METHOD TagsOfStyle(int style) override;
    const char* SCI_METHOD DescriptionOfStyle(int style) override;

    // ILexer5 additions
    const char* SCI_METHOD GetName() override;
    int SCI_METHOD GetIdentifier() override;
    const char* SCI_METHOD PropertyGet(const char* key) override;

private:
    const TreeSitterGrammar* m_grammar;
    TreeSitterStyleMap       m_styleMap;
    TSParser*                m_parser;
    TSTree*                  m_tree;
    std::string              m_lexerName;  // "treesitter.<lang>"

    // Properties
    bool m_fold;
};

// ============================================================================
// TreeSitterRegistry - Global discovery and caching of grammar DLLs.
//
// Scans the grammar directory on first use. Reports available lexers
// via GetLexerCount/GetLexerName. Lazily loads grammars on CreateLexer.
// ============================================================================
class TreeSitterRegistry {
public:
    static TreeSitterRegistry& Instance();

    void Initialize(HMODULE hPluginDll);

    int GetLexerCount() const;
    void GetLexerName(unsigned int index, char* name, int bufLength) const;
    Scintilla::ILexer5* CreateLexer(const char* name);

    const TreeSitterGrammar* GetGrammarByName(const std::string& langName);
    const std::wstring& GetGrammarDir() const { return m_grammarDir; }

private:
    TreeSitterRegistry() = default;

    bool m_initialized = false;
    std::wstring m_grammarDir;

    // Ordered list of discovered language names (for stable indexing)
    std::vector<std::string> m_languageNames;

    // language name -> loaded grammar (lazy)
    std::unordered_map<std::string, std::unique_ptr<TreeSitterGrammar>> m_grammars;

    // Languages that failed to load (don't retry)
    std::unordered_set<std::string> m_failedLanguages;
};

// TreeSitterLexer - ILexer5 bridge implementation
//
// SPDX-License-Identifier: GPL-2.0-or-later

#include "TreeSitterLexer.h"

#include <tree_sitter/api.h>

#include <fstream>
#include <sstream>
#include <algorithm>
#include <cassert>
#include <cstring>
#include <regex>

// ============================================================================
// Style metadata
// ============================================================================

const TSStyleInfo kStyleInfo[TSStyle::Count] = {
    { "default",      "SCE_TS_DEFAULT",      "Default text"           },
    { "comment",      "SCE_TS_COMMENT",      "Comment"                },
    { "commentDoc",   "SCE_TS_COMMENTDOC",   "Documentation comment"  },
    { "keyword",      "SCE_TS_KEYWORD",      "Keyword"                },
    { "string",       "SCE_TS_STRING",       "String literal"         },
    { "number",       "SCE_TS_NUMBER",       "Number literal"         },
    { "operator",     "SCE_TS_OPERATOR",     "Operator"               },
    { "function",     "SCE_TS_FUNCTION",     "Function / method name" },
    { "type",         "SCE_TS_TYPE",         "Type name"              },
    { "preprocessor", "SCE_TS_PREPROCESSOR", "Preprocessor / attribute" },
    { "variable",     "SCE_TS_VARIABLE",     "Variable / property"    },
    { "constant",     "SCE_TS_CONSTANT",     "Constant"               },
    { "punctuation",  "SCE_TS_PUNCTUATION",  "Punctuation"            },
    { "tag",          "SCE_TS_TAG",          "Tag (HTML/XML)"         },
};

// ============================================================================
// TreeSitterStyleMap
// ============================================================================

TreeSitterStyleMap::TreeSitterStyleMap()
{
    // Keywords
    m_map["keyword"]             = TSStyle::Keyword;
    m_map["keyword.function"]    = TSStyle::Keyword;
    m_map["keyword.operator"]    = TSStyle::Keyword;
    m_map["keyword.import"]      = TSStyle::Keyword;
    m_map["keyword.type"]        = TSStyle::Keyword;
    m_map["keyword.modifier"]    = TSStyle::Keyword;
    m_map["keyword.repeat"]      = TSStyle::Keyword;
    m_map["keyword.return"]      = TSStyle::Keyword;
    m_map["keyword.conditional"] = TSStyle::Keyword;
    m_map["keyword.exception"]   = TSStyle::Keyword;
    m_map["keyword.coroutine"]   = TSStyle::Keyword;
    m_map["keyword.directive"]   = TSStyle::Preprocessor;
    m_map["include"]             = TSStyle::Keyword;
    m_map["repeat"]              = TSStyle::Keyword;
    m_map["conditional"]         = TSStyle::Keyword;
    m_map["exception"]           = TSStyle::Keyword;

    // Functions
    m_map["function"]            = TSStyle::Function;
    m_map["function.call"]       = TSStyle::Function;
    m_map["function.builtin"]    = TSStyle::Function;
    m_map["function.macro"]      = TSStyle::Function;
    m_map["method"]              = TSStyle::Function;
    m_map["method.call"]         = TSStyle::Function;
    m_map["constructor"]         = TSStyle::Function;

    // Comments
    m_map["comment"]             = TSStyle::Comment;
    m_map["comment.documentation"] = TSStyle::CommentDoc;

    // Strings
    m_map["string"]              = TSStyle::String;
    m_map["string.documentation"]= TSStyle::String;
    m_map["string.regex"]        = TSStyle::String;
    m_map["string.escape"]       = TSStyle::String;
    m_map["string.special"]      = TSStyle::String;
    m_map["character"]           = TSStyle::String;
    m_map["character.special"]   = TSStyle::String;

    // Numbers
    m_map["number"]              = TSStyle::Number;
    m_map["number.float"]        = TSStyle::Number;
    m_map["float"]               = TSStyle::Number;
    m_map["boolean"]             = TSStyle::Number;

    // Operators and punctuation
    m_map["operator"]            = TSStyle::Operator;
    m_map["punctuation"]         = TSStyle::Punctuation;
    m_map["punctuation.bracket"] = TSStyle::Punctuation;
    m_map["punctuation.delimiter"] = TSStyle::Punctuation;
    m_map["punctuation.special"] = TSStyle::Punctuation;

    // Preprocessor / attributes
    m_map["preproc"]             = TSStyle::Preprocessor;
    m_map["define"]              = TSStyle::Preprocessor;
    m_map["attribute"]           = TSStyle::Preprocessor;
    m_map["attribute.builtin"]   = TSStyle::Preprocessor;

    // Types
    m_map["type"]                = TSStyle::Type;
    m_map["type.builtin"]        = TSStyle::Type;
    m_map["type.definition"]     = TSStyle::Type;
    m_map["type.qualifier"]      = TSStyle::Type;
    m_map["storageclass"]        = TSStyle::Type;

    // Variables / properties
    m_map["variable"]            = TSStyle::Default;
    m_map["variable.builtin"]    = TSStyle::Variable;
    m_map["variable.parameter"]  = TSStyle::Default;
    m_map["variable.member"]     = TSStyle::Variable;
    m_map["property"]            = TSStyle::Variable;
    m_map["field"]               = TSStyle::Variable;
    m_map["constant"]            = TSStyle::Constant;
    m_map["constant.builtin"]    = TSStyle::Constant;
    m_map["constant.macro"]      = TSStyle::Constant;
    m_map["module"]              = TSStyle::Type;
    m_map["namespace"]           = TSStyle::Type;
    m_map["label"]               = TSStyle::Variable;
    m_map["tag"]                 = TSStyle::Tag;
    m_map["tag.attribute"]       = TSStyle::Variable;
    m_map["tag.delimiter"]       = TSStyle::Punctuation;
}

char TreeSitterStyleMap::MapCapture(const std::string& captureName) const
{
    // Exact match
    auto it = m_map.find(captureName);
    if (it != m_map.end())
        return it->second;

    // Prefix fallback: "keyword.control.fsharp" -> "keyword.control" -> "keyword"
    std::string prefix = captureName;
    while (true) {
        auto pos = prefix.rfind('.');
        if (pos == std::string::npos)
            break;
        prefix = prefix.substr(0, pos);
        it = m_map.find(prefix);
        if (it != m_map.end())
            return it->second;
    }

    return TSStyle::Default;
}

// ============================================================================
// TreeSitterGrammar
// ============================================================================

TreeSitterGrammar::TreeSitterGrammar()
    : m_hDll(nullptr), m_pLanguage(nullptr), m_pQuery(nullptr),
      m_pLocalsQuery(nullptr), m_pInjectionQuery(nullptr)
{
}

TreeSitterGrammar::~TreeSitterGrammar()
{
    if (m_pQuery)
        ts_query_delete(m_pQuery);
    if (m_pLocalsQuery)
        ts_query_delete(m_pLocalsQuery);
    if (m_pInjectionQuery)
        ts_query_delete(m_pInjectionQuery);
    if (m_hDll)
        FreeLibrary(m_hDll);
}

TSQuery* TreeSitterGrammar::LoadQuery(const std::wstring& path)
{
    std::ifstream file(path, std::ios::binary);
    if (!file.is_open())
        return nullptr;

    std::string source((std::istreambuf_iterator<char>(file)),
                        std::istreambuf_iterator<char>());

    uint32_t errorOffset = 0;
    TSQueryError errorType = TSQueryErrorNone;
    TSQuery* pQuery = ts_query_new(
        m_pLanguage,
        source.c_str(),
        static_cast<uint32_t>(source.size()),
        &errorOffset,
        &errorType);

    if (errorType != TSQueryErrorNone)
        return nullptr;

    return pQuery;
}

bool TreeSitterGrammar::Load(const std::wstring& grammarDir, const std::wstring& language)
{
    // Convert wide language name to ASCII for m_name
    m_name.clear();
    for (wchar_t ch : language)
        m_name += static_cast<char>(ch);

    // Load grammar DLL: tree-sitter-<language>.dll
    std::wstring dllPath = grammarDir + L"\\tree-sitter-" + language + L".dll";
    m_hDll = LoadLibraryW(dllPath.c_str());
    if (!m_hDll)
        return false;

    // Resolve tree_sitter_<language>() export (hyphens -> underscores)
    std::string funcName = "tree_sitter_";
    for (wchar_t ch : language) {
        funcName += (ch == L'-') ? '_' : static_cast<char>(ch);
    }

    typedef const TSLanguage* (*TSLanguageFunc)();
    auto pfn = reinterpret_cast<TSLanguageFunc>(GetProcAddress(m_hDll, funcName.c_str()));
    if (!pfn) {
        FreeLibrary(m_hDll);
        m_hDll = nullptr;
        return false;
    }

    m_pLanguage = pfn();
    if (!m_pLanguage) {
        FreeLibrary(m_hDll);
        m_hDll = nullptr;
        return false;
    }

    // Load highlight query: <language>-highlights.scm
    std::wstring queryPath = grammarDir + L"\\" + language + L"-highlights.scm";
    m_pQuery = LoadQuery(queryPath);

    // Load locals query (optional): <language>-locals.scm
    std::wstring localsPath = grammarDir + L"\\" + language + L"-locals.scm";
    m_pLocalsQuery = LoadQuery(localsPath);

    // Load injection query (optional): <language>-injections.scm
    std::wstring injectionPath = grammarDir + L"\\" + language + L"-injections.scm";
    m_pInjectionQuery = LoadQuery(injectionPath);

    return true;
}

// ============================================================================
// TreeSitterILexer
// ============================================================================

TreeSitterILexer::TreeSitterILexer(const TreeSitterGrammar* grammar)
    : m_grammar(grammar)
    , m_parser(nullptr)
    , m_tree(nullptr)
    , m_fold(true)
{
    m_lexerName = "treesitter." + grammar->GetName();

    m_parser = ts_parser_new();
    if (m_parser && grammar->GetLanguage())
        ts_parser_set_language(m_parser, grammar->GetLanguage());
}

TreeSitterILexer::~TreeSitterILexer()
{
    if (m_tree)
        ts_tree_delete(m_tree);
    if (m_parser)
        ts_parser_delete(m_parser);
}

int SCI_METHOD TreeSitterILexer::Version() const { return Scintilla::lvRelease5; }
void SCI_METHOD TreeSitterILexer::Release() { delete this; }

const char* SCI_METHOD TreeSitterILexer::PropertyNames() { return "fold\n"; }
int SCI_METHOD TreeSitterILexer::PropertyType(const char* name)
{
    if (strcmp(name, "fold") == 0) return 0; // SC_TYPE_BOOLEAN
    return -1;
}
const char* SCI_METHOD TreeSitterILexer::DescribeProperty(const char* name)
{
    if (strcmp(name, "fold") == 0) return "Enable code folding";
    return "";
}
Sci_Position SCI_METHOD TreeSitterILexer::PropertySet(const char* key, const char* val)
{
    if (strcmp(key, "fold") == 0) {
        m_fold = (val && val[0] == '1');
        return 0;
    }
    return -1;
}
const char* SCI_METHOD TreeSitterILexer::PropertyGet(const char* key)
{
    if (strcmp(key, "fold") == 0) return m_fold ? "1" : "0";
    return "";
}

const char* SCI_METHOD TreeSitterILexer::DescribeWordListSets() { return ""; }
Sci_Position SCI_METHOD TreeSitterILexer::WordListSet(int, const char*) { return -1; }

// ---------------------------------------------------------------------------
// Predicate evaluation helpers
// ---------------------------------------------------------------------------

// Check if a match satisfies all predicates for its pattern.
// Supports #match? / #not-match? (regex) and #eq? / #not-eq? (literal).
static bool EvaluatePredicates(const TSQuery* query, const TSQueryMatch& match,
                               const char* docText, uint32_t docLen)
{
    uint32_t stepCount = 0;
    const TSQueryPredicateStep* steps =
        ts_query_predicates_for_pattern(query, match.pattern_index, &stepCount);
    if (!steps || stepCount == 0)
        return true;  // No predicates = always matches

    // Walk predicate steps. Each predicate is a sequence of steps ending
    // with TSQueryPredicateStepTypeDone.
    uint32_t i = 0;
    while (i < stepCount) {
        // First step should be a string naming the predicate function
        if (steps[i].type != TSQueryPredicateStepTypeString) {
            // Skip to next predicate
            while (i < stepCount && steps[i].type != TSQueryPredicateStepTypeDone) i++;
            if (i < stepCount) i++; // skip Done
            continue;
        }

        uint32_t predNameLen = 0;
        const char* predName = ts_query_string_value_for_id(
            query, steps[i].value_id, &predNameLen);
        std::string pred(predName, predNameLen);
        i++; // move past predicate name

        if (pred == "match?" || pred == "not-match?") {
            // Expect: capture, string(regex)
            if (i + 2 <= stepCount &&
                steps[i].type == TSQueryPredicateStepTypeCapture &&
                steps[i + 1].type == TSQueryPredicateStepTypeString) {

                uint32_t captureIdx = steps[i].value_id;
                uint32_t regexLen = 0;
                const char* regexStr = ts_query_string_value_for_id(
                    query, steps[i + 1].value_id, &regexLen);

                // Find the captured node's text
                std::string nodeText;
                for (uint16_t c = 0; c < match.capture_count; c++) {
                    if (match.captures[c].index == captureIdx) {
                        uint32_t s = ts_node_start_byte(match.captures[c].node);
                        uint32_t e = ts_node_end_byte(match.captures[c].node);
                        if (s < docLen && e <= docLen && s < e)
                            nodeText.assign(docText + s, e - s);
                        break;
                    }
                }

                try {
                    std::regex re(regexStr, regexLen,
                                  std::regex::ECMAScript | std::regex::optimize);
                    bool matched = std::regex_search(nodeText, re);
                    if ((pred == "match?" && !matched) ||
                        (pred == "not-match?" && matched))
                        return false;
                } catch (...) {
                    // Invalid regex — skip predicate
                }

                i += 2;
            }
        } else if (pred == "eq?" || pred == "not-eq?") {
            // Expect: capture, string  OR  capture, capture
            if (i + 2 <= stepCount &&
                steps[i].type == TSQueryPredicateStepTypeCapture) {

                uint32_t captureIdx1 = steps[i].value_id;
                std::string text1;
                for (uint16_t c = 0; c < match.capture_count; c++) {
                    if (match.captures[c].index == captureIdx1) {
                        uint32_t s = ts_node_start_byte(match.captures[c].node);
                        uint32_t e = ts_node_end_byte(match.captures[c].node);
                        if (s < docLen && e <= docLen && s < e)
                            text1.assign(docText + s, e - s);
                        break;
                    }
                }

                std::string text2;
                if (steps[i + 1].type == TSQueryPredicateStepTypeString) {
                    uint32_t sLen = 0;
                    const char* sVal = ts_query_string_value_for_id(
                        query, steps[i + 1].value_id, &sLen);
                    text2.assign(sVal, sLen);
                } else if (steps[i + 1].type == TSQueryPredicateStepTypeCapture) {
                    uint32_t captureIdx2 = steps[i + 1].value_id;
                    for (uint16_t c = 0; c < match.capture_count; c++) {
                        if (match.captures[c].index == captureIdx2) {
                            uint32_t s = ts_node_start_byte(match.captures[c].node);
                            uint32_t e = ts_node_end_byte(match.captures[c].node);
                            if (s < docLen && e <= docLen && s < e)
                                text2.assign(docText + s, e - s);
                            break;
                        }
                    }
                }

                bool equal = (text1 == text2);
                if ((pred == "eq?" && !equal) ||
                    (pred == "not-eq?" && equal))
                    return false;

                i += 2;
            }
        }

        // Skip to Done sentinel
        while (i < stepCount && steps[i].type != TSQueryPredicateStepTypeDone) i++;
        if (i < stepCount) i++; // skip Done
    }

    return true;
}

// ---------------------------------------------------------------------------
// GetSetProperty - Extract #set! predicate property from a query pattern.
//
// #set! predicates like (#set! injection.language "javascript") are encoded
// as: [String "set!", String "key", String "value", Done]
// ---------------------------------------------------------------------------
static std::string GetSetProperty(const TSQuery* query, uint32_t patternIndex,
                                   const std::string& key)
{
    uint32_t stepCount = 0;
    const TSQueryPredicateStep* steps =
        ts_query_predicates_for_pattern(query, patternIndex, &stepCount);
    if (!steps || stepCount == 0)
        return std::string();

    for (uint32_t i = 0; i + 2 < stepCount; i++) {
        if (steps[i].type != TSQueryPredicateStepTypeString)
            continue;

        uint32_t nameLen = 0;
        const char* name = ts_query_string_value_for_id(query, steps[i].value_id, &nameLen);
        if (!name || std::string(name, nameLen) != "set!")
            continue;

        if (i + 1 >= stepCount || steps[i + 1].type != TSQueryPredicateStepTypeString)
            continue;

        uint32_t keyLen = 0;
        const char* keyStr = ts_query_string_value_for_id(query, steps[i + 1].value_id, &keyLen);
        if (!keyStr || std::string(keyStr, keyLen) != key) {
            // Skip to Done sentinel for this predicate
            while (i < stepCount && steps[i].type != TSQueryPredicateStepTypeDone)
                i++;
            continue;
        }

        if (i + 2 >= stepCount || steps[i + 2].type != TSQueryPredicateStepTypeString)
            continue;

        uint32_t valLen = 0;
        const char* valStr = ts_query_string_value_for_id(query, steps[i + 2].value_id, &valLen);
        if (valStr)
            return std::string(valStr, valLen);
    }

    return std::string();
}

// ---------------------------------------------------------------------------
// Highlight entry used during Lex() - defined here so helpers can share it.
// ---------------------------------------------------------------------------
struct LexHighlight {
    uint32_t start;
    uint32_t end;
    char style;
    uint32_t pattern;  // higher pattern index = higher priority
};

// ---------------------------------------------------------------------------
// Lex() - The core: parse document with tree-sitter, apply Scintilla styles
// ---------------------------------------------------------------------------
void SCI_METHOD TreeSitterILexer::Lex(Sci_PositionU startPos, Sci_Position lengthDoc,
                                       int /*initStyle*/, Scintilla::IDocument* pAccess)
{
    if (!m_parser || !m_grammar || !m_grammar->GetHighlightQuery())
        return;

    // Get full document text via IDocument
    Sci_Position docLen = pAccess->Length();
    const char* docText = pAccess->BufferPointer();
    if (!docText || docLen == 0)
        return;

    // Parse the full document
    if (m_tree) {
        ts_tree_delete(m_tree);
        m_tree = nullptr;
    }
    m_tree = ts_parser_parse_string(m_parser, nullptr,
                                     docText, static_cast<uint32_t>(docLen));
    if (!m_tree)
        return;

    // Compute the byte range to style
    Sci_PositionU endPos = startPos + static_cast<Sci_PositionU>(lengthDoc);
    if (endPos > static_cast<Sci_PositionU>(docLen))
        endPos = static_cast<Sci_PositionU>(docLen);

    TSNode rootNode = ts_tree_root_node(m_tree);

    // ---------------------------------------------------------------
    // Step 1: Run locals query (full tree) to build scope/def/ref info
    // ---------------------------------------------------------------
    struct LocalDef {
        std::string name;
        uint32_t startByte;
        uint32_t endByte;
        int highlight;  // TSStyle, -1 if unknown
    };
    struct LocalScope {
        uint32_t startByte;
        uint32_t endByte;
        bool inherits;
        std::vector<LocalDef> defs;
    };
    struct PendingRef {
        std::string name;
        uint32_t startByte;
        uint32_t endByte;
    };

    std::vector<LocalScope> localScopes;
    std::vector<PendingRef> pendingRefs;

    const TSQuery* localsQuery = m_grammar->GetLocalsQuery();
    if (localsQuery) {
        TSQueryCursor* lCursor = ts_query_cursor_new();
        if (lCursor) {
            ts_query_cursor_exec(lCursor, localsQuery, rootNode);

            TSQueryMatch match;
            while (ts_query_cursor_next_match(lCursor, &match)) {
                for (uint16_t i = 0; i < match.capture_count; i++) {
                    TSQueryCapture capture = match.captures[i];
                    TSNode node = capture.node;

                    uint32_t nameLen = 0;
                    const char* captureName = ts_query_capture_name_for_id(
                        localsQuery, capture.index, &nameLen);
                    std::string sCapture(captureName, nameLen);

                    uint32_t nodeStart = ts_node_start_byte(node);
                    uint32_t nodeEnd = ts_node_end_byte(node);

                    if (sCapture == "local.scope") {
                        bool inherits = true;
                        std::string inheritVal = GetSetProperty(
                            localsQuery, match.pattern_index, "local.scope-inherits");
                        if (inheritVal == "false")
                            inherits = false;

                        LocalScope scope;
                        scope.startByte = nodeStart;
                        scope.endByte = nodeEnd;
                        scope.inherits = inherits;
                        localScopes.push_back(std::move(scope));
                    }
                    else if (sCapture == "local.definition") {
                        if (nodeStart < static_cast<uint32_t>(docLen) &&
                            nodeEnd <= static_cast<uint32_t>(docLen)) {
                            std::string defName(docText + nodeStart, nodeEnd - nodeStart);

                            // Find the innermost enclosing scope
                            LocalScope* pBestScope = nullptr;
                            for (auto& scope : localScopes) {
                                if (nodeStart >= scope.startByte && nodeEnd <= scope.endByte) {
                                    if (!pBestScope ||
                                        (scope.endByte - scope.startByte) <
                                        (pBestScope->endByte - pBestScope->startByte)) {
                                        pBestScope = &scope;
                                    }
                                }
                            }
                            if (pBestScope) {
                                LocalDef def;
                                def.name = defName;
                                def.startByte = nodeStart;
                                def.endByte = nodeEnd;
                                def.highlight = -1;
                                pBestScope->defs.push_back(std::move(def));
                            }
                        }
                    }
                    else if (sCapture == "local.reference") {
                        if (nodeStart < static_cast<uint32_t>(docLen) &&
                            nodeEnd <= static_cast<uint32_t>(docLen)) {
                            PendingRef ref;
                            ref.name.assign(docText + nodeStart, nodeEnd - nodeStart);
                            ref.startByte = nodeStart;
                            ref.endByte = nodeEnd;
                            pendingRefs.push_back(std::move(ref));
                        }
                    }
                }
            }
            ts_query_cursor_delete(lCursor);

            // Sort scopes by start byte, then by size (smallest first)
            std::sort(localScopes.begin(), localScopes.end(),
                [](const LocalScope& a, const LocalScope& b) {
                    if (a.startByte != b.startByte)
                        return a.startByte < b.startByte;
                    return (a.endByte - a.startByte) < (b.endByte - b.startByte);
                });
        }
    }

    // ---------------------------------------------------------------
    // Step 2: Run highlight query (full tree) and collect highlights
    // ---------------------------------------------------------------
    const TSQuery* query = m_grammar->GetHighlightQuery();
    TSQueryCursor* cursor = ts_query_cursor_new();
    if (!cursor)
        return;

    // No byte range restriction — we need full-tree highlights for
    // locals resolution (definitions may be outside the styled range).
    ts_query_cursor_exec(cursor, query, rootNode);

    std::vector<LexHighlight> highlights;
    // Map from node byte range to style, for locals definition resolution.
    // Key: (startByte << 32 | endByte). Assumes files < 4 GB.
    std::unordered_map<uint64_t, char> nodeHighlightMap;

    TSQueryMatch match;
    while (ts_query_cursor_next_match(cursor, &match)) {
        if (!EvaluatePredicates(query, match, docText, static_cast<uint32_t>(docLen)))
            continue;

        for (uint16_t i = 0; i < match.capture_count; i++) {
            TSQueryCapture capture = match.captures[i];
            TSNode node = capture.node;

            uint32_t nameLen = 0;
            const char* name = ts_query_capture_name_for_id(
                query, capture.index, &nameLen);

            std::string captureName(name, nameLen);
            char style = m_styleMap.MapCapture(captureName);

            uint32_t nodeStart = ts_node_start_byte(node);
            uint32_t nodeEnd = ts_node_end_byte(node);

            if (nodeStart < nodeEnd) {
                highlights.push_back({ nodeStart, nodeEnd, style, match.pattern_index });

                uint64_t nodeKey = (static_cast<uint64_t>(nodeStart) << 32) |
                                    static_cast<uint64_t>(nodeEnd);
                nodeHighlightMap[nodeKey] = style;
            }
        }
    }
    ts_query_cursor_delete(cursor);

    // ---------------------------------------------------------------
    // Step 3: Resolve locals — scope-aware coloring
    // ---------------------------------------------------------------
    if (!localScopes.empty()) {
        // Pass 1: Assign highlights to definitions from the highlight map
        for (auto& scope : localScopes) {
            for (auto& def : scope.defs) {
                uint64_t defKey = (static_cast<uint64_t>(def.startByte) << 32) |
                                   static_cast<uint64_t>(def.endByte);
                auto it = nodeHighlightMap.find(defKey);
                if (it != nodeHighlightMap.end())
                    def.highlight = it->second;
            }
        }

        // Pass 2: Resolve references — find matching definition in enclosing scopes
        std::unordered_map<uint64_t, char> refHighlights;
        for (const auto& ref : pendingRefs) {
            int resolvedHighlight = -1;

            // Collect scopes enclosing this reference
            std::vector<const LocalScope*> enclosing;
            for (const auto& scope : localScopes) {
                if (ref.startByte >= scope.startByte && ref.endByte <= scope.endByte)
                    enclosing.push_back(&scope);
            }
            // Sort by size (innermost first)
            std::sort(enclosing.begin(), enclosing.end(),
                [](const LocalScope* a, const LocalScope* b) {
                    return (a->endByte - a->startByte) < (b->endByte - b->startByte);
                });

            for (const auto* pScope : enclosing) {
                for (const auto& def : pScope->defs) {
                    if (def.name == ref.name &&
                        def.startByte <= ref.startByte &&
                        def.highlight >= 0) {
                        resolvedHighlight = def.highlight;
                        break;
                    }
                }
                if (resolvedHighlight >= 0)
                    break;
                if (!pScope->inherits)
                    break;
            }

            if (resolvedHighlight >= 0) {
                uint64_t refKey = (static_cast<uint64_t>(ref.startByte) << 32) |
                                   static_cast<uint64_t>(ref.endByte);
                refHighlights[refKey] = static_cast<char>(resolvedHighlight);
            }
        }

        // Pass 3: Apply resolved reference highlights to entries
        for (auto& h : highlights) {
            uint64_t key = (static_cast<uint64_t>(h.start) << 32) |
                            static_cast<uint64_t>(h.end);
            auto it = refHighlights.find(key);
            if (it != refHighlights.end())
                h.style = it->second;
        }
    }

    // ---------------------------------------------------------------
    // Step 4: Run injection query — embedded language highlighting
    // ---------------------------------------------------------------
    const TSQuery* injQuery = m_grammar->GetInjectionQuery();
    if (injQuery) {
        TSQueryCursor* ijCursor = ts_query_cursor_new();
        if (ijCursor) {
            ts_query_cursor_exec(ijCursor, injQuery, rootNode);

            struct InjectionRegion {
                std::string language;
                uint32_t contentStart;
                uint32_t contentEnd;
                TSPoint startPoint;
                TSPoint endPoint;
            };
            std::vector<InjectionRegion> injections;

            TSQueryMatch ijMatch;
            while (ts_query_cursor_next_match(ijCursor, &ijMatch)) {
                std::string language;
                TSNode contentNode = {};
                bool hasContent = false;

                // Check for #set! injection.language
                language = GetSetProperty(injQuery, ijMatch.pattern_index,
                                          "injection.language");

                // Check for #set! injection.self
                if (language.empty()) {
                    std::string selfVal = GetSetProperty(injQuery, ijMatch.pattern_index,
                                                         "injection.self");
                    if (!selfVal.empty())
                        language = m_grammar->GetName();
                }

                for (uint16_t i = 0; i < ijMatch.capture_count; i++) {
                    TSQueryCapture cap = ijMatch.captures[i];
                    uint32_t nLen = 0;
                    const char* cName = ts_query_capture_name_for_id(
                        injQuery, cap.index, &nLen);
                    std::string sCapture(cName, nLen);

                    if (sCapture == "injection.content") {
                        contentNode = cap.node;
                        hasContent = true;
                    } else if (sCapture == "injection.language" && language.empty()) {
                        uint32_t s = ts_node_start_byte(cap.node);
                        uint32_t e = ts_node_end_byte(cap.node);
                        if (s < static_cast<uint32_t>(docLen) &&
                            e <= static_cast<uint32_t>(docLen))
                            language.assign(docText + s, e - s);
                    }
                }

                if (hasContent && !language.empty()) {
                    InjectionRegion region;
                    region.language = language;
                    region.contentStart = ts_node_start_byte(contentNode);
                    region.contentEnd = ts_node_end_byte(contentNode);
                    region.startPoint = ts_node_start_point(contentNode);
                    region.endPoint = ts_node_end_point(contentNode);
                    injections.push_back(std::move(region));
                }
            }
            ts_query_cursor_delete(ijCursor);

            // Process each injection region
            TreeSitterRegistry& registry = TreeSitterRegistry::Instance();
            for (const auto& inj : injections) {
                const TreeSitterGrammar* pInjGrammar =
                    registry.GetGrammarByName(inj.language);
                if (!pInjGrammar || !pInjGrammar->GetHighlightQuery())
                    continue;

                if (inj.contentStart >= static_cast<uint32_t>(docLen) ||
                    inj.contentEnd > static_cast<uint32_t>(docLen))
                    continue;

                std::string injContent(docText + inj.contentStart,
                                        inj.contentEnd - inj.contentStart);
                if (injContent.empty())
                    continue;

                TSParser* pInjParser = ts_parser_new();
                if (!pInjParser)
                    continue;

                ts_parser_set_language(pInjParser, pInjGrammar->GetLanguage());
                TSTree* pInjTree = ts_parser_parse_string(
                    pInjParser, nullptr,
                    injContent.c_str(),
                    static_cast<uint32_t>(injContent.size()));

                if (pInjTree) {
                    const TSQuery* pInjQ = pInjGrammar->GetHighlightQuery();
                    TSNode injRoot = ts_tree_root_node(pInjTree);
                    TSQueryCursor* pInjCur = ts_query_cursor_new();
                    if (pInjCur) {
                        ts_query_cursor_exec(pInjCur, pInjQ, injRoot);

                        TSQueryMatch injM;
                        while (ts_query_cursor_next_match(pInjCur, &injM)) {
                            if (!EvaluatePredicates(pInjQ, injM,
                                    injContent.c_str(),
                                    static_cast<uint32_t>(injContent.size())))
                                continue;

                            for (uint16_t ci = 0; ci < injM.capture_count; ci++) {
                                TSQueryCapture cap = injM.captures[ci];
                                uint32_t cnLen = 0;
                                const char* cn = ts_query_capture_name_for_id(
                                    pInjQ, cap.index, &cnLen);
                                std::string sName(cn, cnLen);
                                char style = m_styleMap.MapCapture(sName);

                                // Map back to parent document byte positions
                                uint32_t capStart = ts_node_start_byte(cap.node);
                                uint32_t capEnd = ts_node_end_byte(cap.node);
                                uint32_t parentStart = inj.contentStart + capStart;
                                uint32_t parentEnd = inj.contentStart + capEnd;

                                if (parentStart < parentEnd) {
                                    highlights.push_back({
                                        parentStart, parentEnd,
                                        style, UINT32_MAX  // highest priority
                                    });
                                }
                            }
                        }
                        ts_query_cursor_delete(pInjCur);
                    }
                    ts_tree_delete(pInjTree);
                }
                ts_parser_delete(pInjParser);
            }
        }
    }

    // ---------------------------------------------------------------
    // Step 5: Sort and apply styles
    // ---------------------------------------------------------------
    std::sort(highlights.begin(), highlights.end(),
        [](const LexHighlight& a, const LexHighlight& b) {
            if (a.start != b.start) return a.start < b.start;
            return a.pattern < b.pattern;
        });

    // Build a flat style buffer for the requested range (initialized to Default).
    Sci_Position rangeLen = static_cast<Sci_Position>(endPos - startPos);
    std::vector<char> styles(rangeLen, TSStyle::Default);

    for (const auto& h : highlights) {
        // Clip to requested range
        uint32_t paintStart = (h.start < startPos) ? static_cast<uint32_t>(startPos) : h.start;
        uint32_t paintEnd = (h.end > endPos) ? static_cast<uint32_t>(endPos) : h.end;
        if (paintStart < paintEnd) {
            for (uint32_t pos = paintStart; pos < paintEnd; pos++)
                styles[pos - startPos] = h.style;
        }
    }

    // Apply the style buffer to Scintilla in a single call
    pAccess->StartStyling(startPos);
    pAccess->SetStyles(rangeLen, styles.data());
}

// ---------------------------------------------------------------------------
// Fold() - Compute fold levels from the AST
// ---------------------------------------------------------------------------
void SCI_METHOD TreeSitterILexer::Fold(Sci_PositionU startPos, Sci_Position lengthDoc,
                                        int /*initStyle*/, Scintilla::IDocument* pAccess)
{
    if (!m_fold || !m_tree)
        return;

    Sci_Position docLen = pAccess->Length();
    Sci_PositionU endPos = startPos + static_cast<Sci_PositionU>(lengthDoc);
    if (endPos > static_cast<Sci_PositionU>(docLen))
        endPos = static_cast<Sci_PositionU>(docLen);

    Sci_Position startLine = pAccess->LineFromPosition(startPos);
    Sci_Position endLine = pAccess->LineFromPosition(endPos);

    // Walk the AST to compute per-line depth.
    // Strategy: for each AST node that spans multiple lines, the first line
    // gets SC_FOLDLEVELHEADERFLAG and subsequent lines get +1 depth.
    // We use a simple recursive walk.
    TSNode root = ts_tree_root_node(m_tree);

    // Compute fold levels for each line in range
    // Simple approach: count the depth of the deepest node starting on each line
    struct FoldInfo {
        int level;
        bool isHeader;
    };
    Sci_Position lineCount = endLine - startLine + 1;
    std::vector<FoldInfo> foldInfo(lineCount, { 0, false });

    // Recursive lambda to walk the tree
    struct Walker {
        Scintilla::IDocument* doc;
        Sci_Position startLine;
        Sci_Position endLine;
        std::vector<FoldInfo>& info;

        void walk(TSNode node, int depth) {
            uint32_t childCount = ts_node_child_count(node);

            TSPoint nodeStart = ts_node_start_point(node);
            TSPoint nodeEnd = ts_node_end_point(node);

            Sci_Position nodeLine = static_cast<Sci_Position>(nodeStart.row);
            Sci_Position nodeEndLine = static_cast<Sci_Position>(nodeEnd.row);

            // If this node spans multiple lines, mark the first line as a fold header
            if (nodeEndLine > nodeLine && childCount > 0) {
                if (nodeLine >= startLine && nodeLine <= endLine) {
                    Sci_Position idx = nodeLine - startLine;
                    info[idx].isHeader = true;
                    if (depth > info[idx].level)
                        info[idx].level = depth;
                }
            }

            // Update level for all lines this node touches
            for (Sci_Position line = nodeLine; line <= nodeEndLine; line++) {
                if (line >= startLine && line <= endLine) {
                    Sci_Position idx = line - startLine;
                    if (depth > info[idx].level)
                        info[idx].level = depth;
                }
            }

            // Recurse into children
            for (uint32_t i = 0; i < childCount; i++) {
                TSNode child = ts_node_child(node, i);
                walk(child, depth + 1);
            }
        }
    };

    Walker walker = { pAccess, startLine, endLine, foldInfo };
    walker.walk(root, 0);

    // Apply fold levels
    constexpr int SC_FOLDLEVELBASE = 0x400;
    constexpr int SC_FOLDLEVELHEADERFLAG = 0x2000;

    for (Sci_Position line = startLine; line <= endLine; line++) {
        Sci_Position idx = line - startLine;
        int level = SC_FOLDLEVELBASE + foldInfo[idx].level;
        if (foldInfo[idx].isHeader)
            level |= SC_FOLDLEVELHEADERFLAG;
        pAccess->SetLevel(line, level);
    }
}

// ---------------------------------------------------------------------------
// Remaining ILexer5 stubs
// ---------------------------------------------------------------------------

void* SCI_METHOD TreeSitterILexer::PrivateCall(int, void*) { return nullptr; }
int SCI_METHOD TreeSitterILexer::LineEndTypesSupported() { return 0; }
int SCI_METHOD TreeSitterILexer::AllocateSubStyles(int, int) { return -1; }
int SCI_METHOD TreeSitterILexer::SubStylesStart(int) { return -1; }
int SCI_METHOD TreeSitterILexer::SubStylesLength(int) { return 0; }
int SCI_METHOD TreeSitterILexer::StyleFromSubStyle(int subStyle) { return subStyle; }
int SCI_METHOD TreeSitterILexer::PrimaryStyleFromStyle(int style) { return style; }
void SCI_METHOD TreeSitterILexer::FreeSubStyles() {}
void SCI_METHOD TreeSitterILexer::SetIdentifiers(int, const char*) {}
int SCI_METHOD TreeSitterILexer::DistanceToSecondaryStyles() { return 0; }
const char* SCI_METHOD TreeSitterILexer::GetSubStyleBases() { return ""; }

int SCI_METHOD TreeSitterILexer::NamedStyles() { return TSStyle::Count; }

const char* SCI_METHOD TreeSitterILexer::NameOfStyle(int style)
{
    if (style >= 0 && style < TSStyle::Count) return kStyleInfo[style].name;
    return "";
}

const char* SCI_METHOD TreeSitterILexer::TagsOfStyle(int style)
{
    if (style >= 0 && style < TSStyle::Count) return kStyleInfo[style].tags;
    return "";
}

const char* SCI_METHOD TreeSitterILexer::DescriptionOfStyle(int style)
{
    if (style >= 0 && style < TSStyle::Count) return kStyleInfo[style].description;
    return "";
}

const char* SCI_METHOD TreeSitterILexer::GetName() { return m_lexerName.c_str(); }
int SCI_METHOD TreeSitterILexer::GetIdentifier() { return 0; }

// ============================================================================
// TreeSitterRegistry
// ============================================================================

TreeSitterRegistry& TreeSitterRegistry::Instance()
{
    static TreeSitterRegistry instance;
    return instance;
}

void TreeSitterRegistry::Initialize(HMODULE hPluginDll)
{
    if (m_initialized)
        return;
    m_initialized = true;

    // Grammar directory: alongside the plugin DLL
    wchar_t dllPath[MAX_PATH] = {};
    GetModuleFileNameW(hPluginDll, dllPath, MAX_PATH);
    std::wstring dir(dllPath);
    auto pos = dir.rfind(L'\\');
    if (pos != std::wstring::npos)
        dir = dir.substr(0, pos);
    m_grammarDir = dir + L"\\TreeSitterGrammars";

    // Check if directory exists
    DWORD attr = GetFileAttributesW(m_grammarDir.c_str());
    if (attr == INVALID_FILE_ATTRIBUTES || !(attr & FILE_ATTRIBUTE_DIRECTORY))
        return;

    // Discover available grammar DLLs (don't load them yet)
    WIN32_FIND_DATAW fd;
    std::wstring pattern = m_grammarDir + L"\\tree-sitter-*.dll";
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE)
        return;

    do {
        std::wstring fileName(fd.cFileName);
        const std::wstring prefix = L"tree-sitter-";
        const std::wstring suffix = L".dll";
        if (fileName.size() > prefix.size() + suffix.size()) {
            std::wstring langW = fileName.substr(
                prefix.size(), fileName.size() - prefix.size() - suffix.size());

            // Also verify the .scm file exists
            std::wstring scmPath = m_grammarDir + L"\\" + langW + L"-highlights.scm";
            DWORD scmAttr = GetFileAttributesW(scmPath.c_str());
            if (scmAttr != INVALID_FILE_ATTRIBUTES) {
                std::string langA;
                for (wchar_t ch : langW) langA += static_cast<char>(ch);
                m_languageNames.push_back(langA);
            }
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);

    // Sort for deterministic ordering
    std::sort(m_languageNames.begin(), m_languageNames.end());
}

int TreeSitterRegistry::GetLexerCount() const
{
    return static_cast<int>(m_languageNames.size());
}

void TreeSitterRegistry::GetLexerName(unsigned int index, char* name, int bufLength) const
{
    if (index >= m_languageNames.size() || !name || bufLength <= 0) {
        if (name && bufLength > 0) name[0] = '\0';
        return;
    }
    // Lexer name format: "treesitter.<language>"
    std::string fullName = "treesitter." + m_languageNames[index];
    int len = static_cast<int>(fullName.size());
    if (len >= bufLength) len = bufLength - 1;
    memcpy(name, fullName.c_str(), len);
    name[len] = '\0';
}

Scintilla::ILexer5* TreeSitterRegistry::CreateLexer(const char* name)
{
    if (!name)
        return nullptr;

    // Parse "treesitter.<language>"
    std::string sName(name);
    const std::string prefix = "treesitter.";
    if (sName.compare(0, prefix.size(), prefix) != 0)
        return nullptr;

    std::string langName = sName.substr(prefix.size());

    // Check if previously failed
    if (m_failedLanguages.count(langName) > 0)
        return nullptr;

    // Check if already loaded
    auto it = m_grammars.find(langName);
    if (it != m_grammars.end()) {
        return new TreeSitterILexer(it->second.get());
    }

    // Lazy load the grammar
    std::wstring langW;
    for (char ch : langName) langW += static_cast<wchar_t>(ch);

    auto grammar = std::make_unique<TreeSitterGrammar>();
    if (!grammar->Load(m_grammarDir, langW) || !grammar->GetHighlightQuery()) {
        m_failedLanguages.insert(langName);
        return nullptr;
    }

    TreeSitterGrammar* pGrammar = grammar.get();
    m_grammars[langName] = std::move(grammar);
    return new TreeSitterILexer(pGrammar);
}

const TreeSitterGrammar* TreeSitterRegistry::GetGrammarByName(const std::string& langName)
{
    if (langName.empty())
        return nullptr;

    // Check if previously failed
    if (m_failedLanguages.count(langName) > 0)
        return nullptr;

    // Check if already loaded
    auto it = m_grammars.find(langName);
    if (it != m_grammars.end())
        return it->second.get();

    // Lazy load the grammar
    std::wstring langW;
    for (char ch : langName) langW += static_cast<wchar_t>(ch);

    auto grammar = std::make_unique<TreeSitterGrammar>();
    if (!grammar->Load(m_grammarDir, langW) || !grammar->GetHighlightQuery()) {
        m_failedLanguages.insert(langName);
        return nullptr;
    }

    TreeSitterGrammar* pGrammar = grammar.get();
    m_grammars[langName] = std::move(grammar);
    return pGrammar;
}

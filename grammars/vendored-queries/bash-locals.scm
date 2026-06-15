; Vendored from nvim-treesitter queries/bash/locals.scm (reliable upstream).
; Unknown node types (grammar-version drift) make tree-sitter skip this whole
; query gracefully, so a mismatch degrades to no locals rather than breaking.
; Scopes
(function_definition) @local.scope

; Definitions
(variable_assignment
  name: (variable_name) @local.definition.var)

(function_definition
  name: (word) @local.definition.function)

; References
(variable_name) @local.reference

(word) @local.reference

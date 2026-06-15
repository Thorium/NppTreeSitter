; Vendored from nvim-treesitter queries/matlab/locals.scm (reliable upstream).
; Unknown node types (grammar-version drift) make tree-sitter skip this whole
; query gracefully, so a mismatch degrades to no locals rather than breaking.
; References
(identifier) @local.reference

; Definitions
(function_definition
  name: (identifier) @local.definition.function
  (function_arguments
    (identifier)* @local.definition.parameter
    (","
      (identifier) @local.definition.parameter)*)?) @local.scope

(assignment
  left: (identifier) @local.definition.var)

(multioutput_variable
  (identifier) @local.definition.var)

(iterator
  .
  (identifier) @local.definition.var)

(lambda
  (arguments
    (identifier) @local.definition.parameter))

(global_operator
  (identifier) @local.definition.var)

(persistent_operator
  (identifier) @local.definition.var)

(catch_clause
  (identifier) @local.definition)

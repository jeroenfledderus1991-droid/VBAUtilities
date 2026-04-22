# VBA Formatting Options — Configurator Spec (with Explanations)

Purpose: all configurable options for formatting VBA code, with clear explanations.  
Use this document as the source for UI/CLI configuration.

---

## 1) Indentation & Alignment

**indent.type**  
Defines whether indentation uses tabs or spaces.
- `tabs` — indent with tab characters.
- `spaces` — indent with spaces.

**indent.size** *(only for spaces)*  
Number of spaces per indent level.
- `2` / `4` / `8`

**indent.continuation**  
How to indent continued lines (line breaks with `_`).
- `align_with_open_paren` — align under the opening `(`.  
- `indent_one_level` — add one extra indent level.  
- `fixed_spaces` — use a fixed number of spaces.

**indent.continuation_size** *(only for fixed_spaces)*  
How many spaces to use for continuation lines.
- `2 | 4 | 8 | 12 | 16`

**indent.label_style** *(GoTo labels)*  
How labels align relative to code.
- `flush_left` — labels at column 1.  
- `indent_with_code` — labels align with code block.

---

## 2) Keyword Style

**keywords.case**  
Casing for VBA keywords (If, Sub, End, etc.).
- `uppercase` — IF, SUB  
- `lowercase` — if, sub  
- `pascal` — If, Sub  
- `preserve` — keep existing casing.

**keywords.enforce**  
Whether to normalize all VBA keywords to the chosen case.
- `true` — force all keywords to chosen case.  
- `false` — do not change keyword casing.

---

## 3) Blocks & Structure

**block.then_style**  
How to format single-line If statements.
- `same_line` — `If x Then DoSomething`  
- `next_line` —  
If x Then
DoSomething
End If
**block.end_if_style**  
Spelling style for `End If`.
- `end_if` — `End If`  
- `endif` — `EndIf`

**block.end_sub_style**  
Spelling style for `End Sub`.
- `end_sub` — `End Sub`  
- `endsub` — `EndSub`

**block.blank_line_between_procedures**  
Number of blank lines between procedures.
- `0 | 1 | 2`

**block.blank_line_after_declarations**  
Blank lines after `Dim/Const/Static` declarations.
- `0 | 1 | 2`

**block.blank_line_between_logical_sections**  
Insert blank lines between logical sections inside a procedure.
- `true` — separate sections (init/process/cleanup).  
- `false` — no automatic separation.

**block.enforce_end_with**  
Ensure all block constructs are properly closed.
- `true` — add missing `End With`, `End Select`, etc.  
- `false` — do not enforce.

---

## 4) Spacing (General)

**spacing.around_operators**  
Add spaces around operators like `=`, `+`, `&`, `-`, `*`.
- `true` — `a + b`  
- `false` — `a+b`

**spacing.after_comma**  
Add a space after commas in argument lists.
- `true` — `Foo(a, b)`  
- `false` — `Foo(a,b)`

**spacing.before_comma**  
Add a space before commas (rare style).
- `true` — `Foo(a , b)`  
- `false` — `Foo(a, b)`

**spacing.after_keyword**  
Ensure a space after keywords like `If`, `For`, `Do`.
- `true` — `If x Then`  
- `false` — `Ifx Then` (not typical)

**spacing.before_then**  
Ensure a space before `Then`.
- `true` — `If x Then`  
- `false` — `If xThen` (not typical)

**spacing.inside_parentheses**  
Spaces inside parentheses.
- `true` — `( a + b )`  
- `false` — `(a + b)`

**spacing.after_type_declaration**  
Space after `As` in declarations.
- `true` — `Dim x As Long`  
- `false` — `Dim x AsLong`

---

## 5) Line Breaks & Length

**line.max_length**  
Maximum line length before wrapping.
- `0` — no limit  
- `80 | 100 | 120 | 140`

**line.wrap_style**  
Where to break lines.
- `none` — never wrap  
- `wrap_at_length` — wrap exactly at max length  
- `wrap_at_operator` — break before operators  
- `wrap_at_comma` — break after commas

**line.continuation_style**  
How to format line continuation.
- `space_underscore` — space before `_`  
- `underscore_only` — no space before `_`

**line.keep_single_line_if_short**  
Keep short one-line statements on a single line.
- `true` — allow `If x Then y`  
- `false` — always expand to multi-line.

---

## 6) Comments

**comment.style**  
How comments are written.
- `apostrophe` — `' comment`  
- `rem` — `Rem comment`  
- `preserve` — keep existing

**comment.placement**  
Where to place comments.
- `above_line` — comment on its own line  
- `inline` — at end of code line  
- `preserve` — keep existing

**comment.align_inline**  
Align inline comments to the same column.
- `true` — align  
- `false` — keep as-is

**comment.blank_line_before_block**  
Insert a blank line before comment blocks.
- `true` — blank line added  
- `false` — no blank line

---

## 7) Declarations & Option Rules

**declarations.option_explicit**  
How to handle `Option Explicit`.
- `require` — ensure it exists  
- `remove` — remove if present  
- `preserve` — leave as-is

**declarations.option_compare**  
How to handle `Option Compare`.
- `none` — no Option Compare  
- `text` — use `Option Compare Text`  
- `binary` — use `Option Compare Binary`  
- `database` — use `Option Compare Database`

**declarations.order**  
Ordering of Option statements.
- `option_explicit_first` — Option Explicit first  
- `option_compare_then_explicit` — Compare before Explicit  
- `preserve` — no reordering

**declarations.const_before_dim**  
Whether `Const` blocks should appear before `Dim`.
- `true` — Const first  
- `false` — keep as-is

**declarations.grouping**  
How to group declarations.
- `separate_by_type` — Const, Dim, Static in blocks  
- `as_found` — keep order as found

---

## 8) Procedures & Headers

**procedure.blank_line_after_header**  
Blank line after procedure declaration line.
- `0 | 1 | 2`

**procedure.align_parameters**  
How to align multi-line parameters.
- `none` — no alignment  
- `vertical` — align parameters in a column  
- `align_with_paren` — align with opening `(`

**procedure.header_comment_template**  
Standard header comment style for procedures.
- `none` — no header  
- `brief` — name + purpose  
- `full` — name, purpose, params, returns, author, date

---

## 9) Select Case Formatting

**selectcase.case_style**  
Indentation of `Case` lines.
- `same_indent` — same indent as `Select Case`  
- `indent_cases` — indent each `Case` one level

**selectcase.blank_line_between_cases**  
Blank lines between cases.
- `0 | 1`

**selectcase.default_case_label**  
Label for default case.
- `case_else` — `Case Else`  
- `else_case` — `Else` (non-standard)

---

## 10) With ... End With

**with.blank_line_inside**  
Blank lines inside With blocks.
- `0 | 1`

**with.indent_inner_block**  
Indent code inside `With`.
- `true` — indent  
- `false` — keep same indent

---

## 11) Error Handling Conventions

**errorhandling.style**  
Preferred error handling pattern.
- `none` — no enforced error handling  
- `goto_handler` — `On Error GoTo ErrHandler`  
- `on_error_resume_next` — `On Error Resume Next`  
- `on_error_goto_0` — `On Error GoTo 0`

**errorhandling.handler_label**  
Label name for error handler block.
- `ErrHandler` (default) or custom

**errorhandling.require_exit_before_handler**  
Require an `Exit Sub/Function` before handler.
- `true` — ensure exit before handler  
- `false` — allow fallthrough

---

## 12) Naming & Casing (Optional)

**naming.variables_case**  
Casing style for variables.
- `camel` — myVar  
- `snake` — my_var  
- `pascal` — MyVar  
- `preserve` — keep existing

**naming.constants_case**  
Casing style for constants.
- `upper` — MAX_SIZE  
- `pascal` — MaxSize  
- `preserve` — keep existing

**naming.procedures_case**  
Casing style for procedures.
- `pascal` — DoWork  
- `camel` — doWork  
- `preserve` — keep existing

---

## 13) Misc & Safety

**misc.remove_trailing_whitespace**  
Remove trailing spaces at line ends.
- `true | false`

**misc.ensure_final_newline**  
Ensure file ends with newline.
- `true | false`

**misc.normalize_line_endings**  
Normalize line endings.
- `crlf` — Windows style  
- `lf` — Unix style  
- `preserve` — keep existing

**misc.keep_blank_lines_max**  
Maximum consecutive blank lines allowed.
- `0 | 1 | 2 | 3`

---

## 14) Presets (Recommended)

**preset.minimal**
- Only whitespace/indent fixes.

**preset.standard**
- Indent + spacing + keywords + blank lines.

**preset.strict**
- Full enforcement including declarations and naming.

---

## 15) Example Config (YAML)

```yaml
indent:
type: spaces
size: 4
continuation: align_with_open_paren

keywords:
case: pascal
enforce: true

block:
then_style: next_line
blank_line_between_procedures: 1
blank_line_after_declarations: 1

spacing:
around_operators: true
after_comma: true
inside_parentheses: false

line:
max_length: 120
wrap_style: wrap_at_operator
continuation_style: space_underscore

comment:
style: apostrophe
placement: above_line

declarations:
option_explicit: require
const_before_dim: true

misc:
remove_trailing_whitespace: true
ensure_final_newline: true
normalize_line_endings: crlf
16) Example Config (JSON){
  "indent": { "type": "spaces", "size": 4, "continuation": "align_with_open_paren" },
  "keywords": { "case": "pascal", "enforce": true },
  "block": { "then_style": "next_line", "blank_line_between_procedures": 1, "blank_line_after_declarations": 1 },
  "spacing": { "around_operators": true, "after_comma": true, "inside_parentheses": false },
  "line": { "max_length": 120, "wrap_style": "wrap_at_operator", "continuation_style": "space_underscore" },
  "comment": { "style": "apostrophe", "placement": "above_line" },
  "declarations": { "option_explicit": "require", "const_before_dim": true },
  "misc": { "remove_trailing_whitespace": true, "ensure_final_newline": true, "normalize_line_endings": "crlf" }
}

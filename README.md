# Visual Basic Metadata (VBM) — Format Specification

```text
Specification:  Visual Basic Metadata (VBM) 1.0
Status:         Draft
File extension: .vbmd  (optional alias: .vbm)
MIME type:      application/vbmd+text
Content header: VERSION 5.00
```

---

## 1. Introduction

**Visual Basic Metadata (VBM)** is a line-oriented, tree-structured, human-readable text format. Its surface syntax is the designer section of a Visual Basic 6 `.frm` / `.ctl` file, generalised so that the same grammar carries any hierarchical metadata, not only form definitions.

A VBM document is a **metadata tree**: a `VERSION` header, exactly one root block of nested `Begin … End` blocks and `BeginProperty … EndProperty` property blocks, followed by a trailing list of `Attribute` declarations. Blocks carry key/value properties; values are opaque raw strings with optional trailing comments.

VBM is a **pure structural format**. It defines the grammar and the document tree; it does **not** define what any identifier means. The words `VB.Form`, `Person`, `Computer`, `Money`, `Caption`, `Age`, `HasSSD` are opaque identifiers carried by the document. Interpretation is the consumer's responsibility — the same relationship as JSON to JSON Schema.

A `.vbmd` file is self-contained. There is no schema file, no registry, no association, no sidecar, no imports. A `.vbmd` file is complete on its own.

### 1.1 Design goals

| Goal | Consequence |
|---|---|
| Byte-for-byte compatibility with the `.frm` designer section | Any `.frm` header is a valid `.vbmd` with zero edits |
| A single grammar for VB6 metadata *and* arbitrary hierarchical data | `VB.Form DemoForm` and `Person Joe_Smith` parse identically |
| Self-contained files | No external file needed to read a `.vbmd` |
| Lossless round-trip | Raw values, comments, casing, whitespace, line endings all preserved |
| A stable grammar for tooling | Written against once; never changes per document type |
| No meaning baked into the format | Interpretation lives in the consumer, exactly as in JSON |

### 1.2 Non-goals

- VBM does **not** define a schema language.
- VBM does **not** define a schema for VB6 forms, projects, or any other content.
- VBM does **not** validate values, cardinalities, required properties, or cross-field constraints.
- VBM does **not** cover the VB6 **code section** that follows the designer section in a real `.frm`; a `.vbmd` file ends where the code section would begin.

---

## 2. File format

### 2.1 Extension and MIME type

| Item | Value |
|---|---|
| Canonical extension | `.vbmd` |
| Alias extension | `.vbm` (accepted; collides with VisualBoyAdvance Movie, so `.vbmd` is preferred) |
| MIME type | `application/vbmd+text` |
| Fallback MIME | `text/x-vbmd` |

### 2.2 Encoding

- **Canonical encoding:** UTF-8 without BOM.
- A reader must accept UTF-8 with a leading byte-order mark and treat it as not part of the document.
- A reader may accept legacy encodings (e.g. code pages) for compatibility with real `.frm` files, but the encoding is not part of the format.
- Non-UTF-8 encodings are out of scope for VBM/1.0.

### 2.3 Line endings

- **Canonical:** CRLF (`\r\n`).
- A reader must accept LF (`\n`) alone.
- Line endings are preserved on lossless re-emission.

### 2.4 Line length

- No maximum is defined. A reader should accept lines of at least 64 KiB.
- Indentation is not significant to the parser.

---

## 3. Reserved words

The following tokens have format-level meaning and are recognised case-insensitively (matching VB6 conventions):

| Reserved word | Role |
|---|---|
| `VERSION` | Content version header, first line of the document |
| `Begin` | Opens a nested block |
| `End` | Closes a `Begin` block |
| `BeginProperty` | Opens a property block |
| `EndProperty` | Closes a `BeginProperty` block |
| `Attribute` | Introduces an attribute declaration after the root block |
| `True` | Recognised boolean literal (in addition to VB6's `-1`) |
| `False` | Recognised boolean literal (in addition to VB6's `0`) |

Reserved words may not be used as **type names**, **block names**, or **property keys** without being treated as reserved tokens. In practice, existing VB6 files never use them as identifiers.

All other identifiers are application-defined and carry no format-level meaning.

---

## 4. Grammar

### 4.1 EBNF

```ebnf
document        = version-line , { blank-line } , root-block ,
                  { attribute-line } , { blank-line } , EOF ;

version-line    = "VERSION" , ws , version-number , [ ws , comment ] , eol ;
version-number  = digits , "." , digits ;

root-block      = block ;

block           = "Begin" , ws , type-name , ws , identifier , eol ,
                  { inner-statement } ,
                  "End" , [ ws ] , eol ;

inner-statement = property-line
                | property-block
                | block
                | blank-line
                | comment-line ;

property-block  = "BeginProperty" , ws , type-name , [ ws , guid ] , eol ,
                  { property-line | blank-line | comment-line } ,
                  "EndProperty" , [ ws ] , eol ;

property-line   = identifier , ws , "=" , ws , raw-value , [ ws , comment ] , eol ;
attribute-line  = "Attribute" , ws , identifier , ws , "=" , ws , raw-value ,
                  [ ws , comment ] , eol ;

comment-line    = ws , comment , eol ;
blank-line      = ws , eol ;

type-name       = identifier , { "." , identifier } ;
identifier      = ident-start , { ident-char } ;
ident-start     = letter | "_" ;
ident-char      = letter | digit | "_" ;

guid            = "{" , 8hex , "-" , 4hex , "-" , 4hex , "-" , 4hex , "-" , 12hex , "}" ;

comment         = "'" , { any-char-except-eol } ;
raw-value       = { any-char-except-eol } ;

ws              = ( space | tab ) , { space | tab } ;
eol             = "\r\n" | "\n" ;
digits          = digit , { digit } ;
letter          = "A".."Z" | "a".."z" ;
digit           = "0".."9" ;
hex             = digit | "A".."F" | "a".."f" ;
```

### 4.2 Reserved words are case-insensitive

`VERSION`, `vErSiOn`, `Begin`, `begin`, `TRUE`, `True` are all accepted. Canonical emission uses the casing shown in §3.

### 4.3 Identifiers are case-preserving, case-insensitive

Identifiers are stored with their original casing for round-trip. Two identifiers that differ only in case refer to the same name; this affects **nothing at the format level** (the format does not compare identifiers), but consumers may rely on it.

---

## 5. Lexical details

### 5.1 Value scanning

Values are opaque text, scanned **quote-aware** from the first non-whitespace character after `=` to end-of-line.

Rules:

1. Scan forward from the first non-`ws` character after `=`.
2. On `"` — a string literal begins. Continue until the next `"`. An unterminated string is a **lexical error**.
3. On `'` **outside** a string literal — a trailing comment begins. Everything from the `'` to end-of-line is the comment; the comment is stripped from the value.
4. On `'` **inside** a string literal — part of the string, not a comment.
5. Trailing whitespace is trimmed from the value after comment stripping.
6. The value is stored exactly as scanned (including internal whitespace and string quotes).

Examples:

| Line | Value | Comment |
|---|---|---|
| `Caption = "It's fine"` | `"It's fine"` | *none* |
| `Caption = "It's fine" ' note` | `"It's fine"` | `note` |
| `Weight = 100 'In kilograms` | `100` | `In kilograms` |
| `Default = -1 'True` | `-1` | `True` |
| `HasSSD = True` | `True` | *none* |
| `OS = "Windows 11"` | `"Windows 11"` | *none* |
| `Empty =` | *(empty string)* | *none* |
| `Empty =  ' trailing` | *(empty string)* | `trailing` |

### 5.2 Trailing-whitespace handling inside strings

Whitespace inside a `"…"` string is preserved as-is. Whitespace outside is trimmed only at the trailing end of the raw value.

### 5.3 Empty values

An empty raw value is syntactically valid: `Key =` yields an empty string value. Whether an empty value is *meaningful* is a consumer question.

### 5.4 Multi-line values

VBM does **not** support multi-line values. A value ends at end-of-line. This matches `.frm` exactly.

---

## 6. Structural rules

A well-formed VBM document satisfies all of the following. Violating any rule is a **structural error**.

1. **One `VERSION` line**, at the top of the file, before any block.
2. **Exactly one root block** — a `Begin … End` pair at depth 0.
3. **`Begin`/`End` balance** — every `Begin` matches exactly one `End` at the same depth.
4. **`BeginProperty`/`EndProperty` balance** — every `BeginProperty` matches exactly one `EndProperty` at the same depth.
5. **`BeginProperty` is flat** — its body contains only property lines, comment lines, and blank lines. No nested `Begin` or `BeginProperty`.
6. **`Begin` names a type and a name** — `Begin <type-name> <identifier>`.
7. **`BeginProperty` names a type** — `BeginProperty <type-name>` with an optional GUID trailer.
8. **`Attribute` placement** — `Attribute` lines appear only after the root block's closing `End`, never inside a block.
9. **No content after `Attribute` section ends** — no `Begin`, `BeginProperty`, or `VERSION` after the first `Attribute` line.
10. **`End` and `EndProperty` take no arguments.** A trailing comment is not permitted on them (matches VB6).
11. **Duplicate keys are permitted syntactically.** Whether duplicates are meaningful is a consumer question (see §9).

### 6.1 What the format deliberately does not enforce

The following are **consumer concerns**, not format concerns. A VBM document is well-formed even if:

- Two blocks share the same name.
- Two properties in the same block share the same key.
- A block's children are out of any "expected" order.
- Properties appear after child blocks.
- Required-looking properties are missing.
- Values look like one type but are another (`-1` is not inherently a boolean).
- An unrecognised `Begin` type appears.

---

## 7. Abstract document model

The canonical model produced by a VBM parser.

### 7.1 Nodes

```text
Document
├── version        : Version        (from VERSION line)
├── root           : Block          (exactly one)
└── attributes     : Attribute[]    (zero or more, ordered)

Block
├── type           : TypeName       (from Begin)
├── name           : Identifier     (from Begin)
├── body           : Statement[]    (ordered; properties, property blocks, blocks, comments, blanks)
├── raw-prefix     : Text           (source text of "Begin …" line, verbatim)
├── raw-suffix     : Text           (source text of "End" line, verbatim)
├── indent         : Text           (leading whitespace of the Begin line)
└── line-ending    : Eol            (CRLF or LF)

PropertyBlock
├── type           : TypeName       (from BeginProperty)
├── guid           : Guid | null    (optional trailer)
├── body           : Property[]     (properties, comments, blanks; no nested blocks)
└── raw-* / indent / line-ending    (same as Block)

Property
├── key            : Identifier
├── value          : RawValue       (verbatim, quotes intact)
├── comment        : string | null  (verbatim text after the `'`)
├── raw            : Text           (whole line, verbatim)
├── indent         : Text
└── line-ending    : Eol

Attribute
├── key            : Identifier
├── value          : RawValue
├── comment        : string | null
├── raw            : Text
├── indent         : Text
└── line-ending    : Eol

Version
├── major          : digits
├── minor          : digits
└── raw            : Text           (whole VERSION line, verbatim)

TypeName         = Identifier ( "." Identifier )*
RawValue         = string         ; the verbatim value text, quotes included
Guid             = "{8-4-4-4-12}"
```

### 7.2 Preservation

Every node keeps:

- its verbatim `raw` text,
- its trailing comment (if any),
- its original indentation,
- its original line ending.

Together these enable **lossless re-emission** (§10.2).

---

## 8. Parsing pipeline

A conforming parser processes a document in the following stages:

```text
1. Tokenize     VERSION, Begin, End, BeginProperty, EndProperty,
                Attribute, identifier, =, raw-value, comment, EOL
                     ↓
2. Parse        untyped tree of Document / Block / PropertyBlock /
                Property / Attribute nodes, with raw text preserved
                     ↓
3. Structural   apply the structural rules of §6
   validation
                     ↓
4. Emit         (a) the abstract tree of §7, or
                (b) canonical text (§10.1), or
                (c) lossless text (§10.2)
```

Stages 1–3 require no external information. Stage 4 is a consumer decision. Type-checking, semantic validation, and schema binding are **outside** VBM.

### 8.1 Error handling

- A **lexical error** (unterminated string, unrecognised reserved word, malformed GUID) aborts parsing.
- A **structural error** (§6) aborts parsing.
- A consumer may choose to apply recovery for a subset of structural errors (e.g. treating an unmatched `End` as the end of the parent block), but recovery is **not part of VBM/1.0**.

---

## 9. Duplicate keys and repeated statements

The format permits:

- Multiple properties with the same `key` in the same body.
- Multiple `BeginProperty` blocks with the same `type` in the same body.
- Multiple `Begin` blocks with the same `type` and even the same `name`.

Meaning is a consumer concern. Typical readings:

| Consumer policy | Effect |
|---|---|
| First-wins | Later duplicates ignored; a warning may be emitted |
| Last-wins | Earlier duplicates overwritten (`.frm` reads scalar properties this way) |
| List | Duplicates collected in order into an ordered list |
| Error | Duplicates rejected |

The format itself imposes no policy.

---

## 10. Serialization

### 10.1 Canonical form

A canonical writer emits:

- three-space indentation per level (VB6 convention),
- `VERSION <major>.<minor>` on the first line,
- reserved words with their canonical casing (§3),
- property lines as `Key = Value` with exactly one space around `=`,
- attribute lines as `Attribute Key = Value` with exactly one space around `=` and after `Attribute`,
- CRLF line endings,
- UTF-8 without BOM.

Comments are preserved in canonical form. Raw values are emitted exactly as parsed (so `"It's fine"` stays as-is, `100` stays as `100`). Canonical emission does **not** re-spell values — `True` stays `True`, and `-1 'True` stays `-1 'True`. Value re-spelling is a consumer concern.

### 10.2 Lossless form

A lossless writer emits the original bytes exactly, using the preserved `raw`, `indent`, and `line-ending` fields of every node. Line endings, casing, whitespace inside `=`, trailing comments, and blank lines are all reproduced verbatim. This is byte-identical round-trip.

### 10.3 Consumer-facing pretty-printing

A consumer that wants to apply its own value conventions (e.g. rewriting `True` → `-1 'True` for VB6 compatibility) does so **after** parsing, then re-serialises using the canonical writer with that consumer's value rendering. This is not part of VBM.

---

## 11. Conformance

### 11.1 Conforming document

A file is a **conforming VBM document** if:

1. It parses according to §4,
2. It satisfies every rule in §6,
3. It uses only UTF-8 (or UTF-8 with BOM) encoding.

### 11.2 Conforming parser

A parser is **conforming** if:

1. It accepts every conforming document.
2. It rejects any document with a lexical error (§8.1) or a structural error (§6).
3. It produces the abstract model of §7.
4. It preserves enough information to reproduce a lossless re-emission.

### 11.3 Conforming writer

- A **canonical writer** emits the form of §10.1.
- A **lossless writer** reproduces the original bytes of §10.2.

---

## 12. Compatibility with `.frm`

A `.frm` file consists of the designer section, followed by a code section. VBM covers **only the designer section**.

- The designer section of any valid `.frm` / `.ctl` file is a conforming VBM document as-is.
- A `.vbmd` file ends at the last `Attribute` line. It does not carry the code section.
- No transformation is required to convert a `.frm` designer section to a `.vbmd` file, and vice versa.

This is the central compatibility claim. It relies entirely on three facts:

1. The VBM grammar is the `.frm` designer grammar.
2. The VBM structural rules are the `.frm` designer structural rules.
3. VBM adds nothing to a document that `.frm` would reject.

---

## 13. Examples

### 13.1 Form document

```vb6
VERSION 5.00
Begin VB.Form DemoForm
   Caption = "Screen Blanker Demo"
   ClientHeight = 960
   ClientWidth = 7470
   BeginProperty Font
      Name = "MS Sans Serif"
      Charset = 0
      Weight = 400
   EndProperty
   Begin VB.CommandButton cmdStartStop
      Caption = "Start Demo"
      Default = -1 'True
      Left = 240
      Top = 120
      Width = 1830
      Height = 390
   End
   Begin VB.Menu mnuOption
      Caption = "&Options"
      Begin VB.Menu mnuExit
         Caption = "E&xit"
      End
   End
End
Attribute VB_Name = "DemoForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
```

### 13.2 Person document

```vb6
VERSION 5.00
Begin Person Joe_Smith
   Name = "Joe Smith"
   Age = 45
   Weight = 100 'In kilograms
   BeginProperty Money
      Currency = "USD"
      Amount = 5000
   EndProperty
   Begin Computer Main_PC
      OS = "Windows 11"
      RAM = 32 'In Gigabytes
      HasSSD = True
   End
End
Attribute VB_Name = "Demo"
Attribute VB_Immutable = False
```

The two documents parse to the **same** abstract model (§7), differing only in identifier strings. A conforming VBM parser does not know, and does not need to know, which is a form and which is a person.

### 13.3 Document that fails structural rules

```vb6
VERSION 5.00
Begin Person Joe_Smith
   Name = "Joe Smith"
End
Attribute VB_Name = "Demo"
Attribute VB_Immutable = False
Begin Computer Main_PC        ' ← structural error: Begin after Attribute
   OS = "Windows 11"
End
```

Rejected by §6 rule 9.

### 13.4 Document with lexical quirks handled correctly

```vb6
VERSION 5.00
Begin Note n1
   Quote = "It's fine" ' real comment
   Hash  = "#not a comment"
   Empty =
   Trail = " padded "
End
```

Values parsed:

| Key | Value | Comment |
|---|---|---|
| `Quote` | `"It's fine"` | `real comment` |
| `Hash`  | `"#not a comment"` | *none* |
| `Empty` | *(empty)* | *none* |
| `Trail` | `" padded "` | *none* |

Note that `#` is not a comment character in VBM (matching `.frm`); only `'` outside a string begins a comment.

---

## 14. Versioning

VBM separates two version numbers, and only one of them lives in the file:

| Version | Where it lives | Purpose |
|---|---|---|
| **Content version** — `VERSION 5.00` | The file itself, first line | Marks the content as `.frm`-compatible |
| **Spec version** — `VBM/1.0` | Documentation, tooling, MIME type | Governs the grammar and rules in this document |

- The content version is fixed at `5.00` for VBM/1.0. That is what makes `.frm` files valid `.vbmd` files.
- The spec version is what consumers refer to when they say "VBM/1.0" or "VBM/1.1".
- A future VBM/2.0 may relax or extend the grammar; it will not change the content version.

---

## Appendix A — Grammar (single-page EBNF)

```ebnf
document        = version-line , { blank-line } , root-block ,
                  { attribute-line } , { blank-line } , EOF ;
version-line    = "VERSION" , ws , digits , "." , digits , [ ws , comment ] , eol ;
root-block      = block ;
block           = "Begin" , ws , type-name , ws , identifier , eol ,
                  { inner-statement } ,
                  "End" , [ ws ] , eol ;
inner-statement = property-line | property-block | block | blank-line | comment-line ;
property-block  = "BeginProperty" , ws , type-name , [ ws , guid ] , eol ,
                  { property-line | blank-line | comment-line } ,
                  "EndProperty" , [ ws ] , eol ;
property-line   = identifier , ws , "=" , ws , raw-value , [ ws , comment ] , eol ;
attribute-line  = "Attribute" , ws , identifier , ws , "=" , ws , raw-value ,
                  [ ws , comment ] , eol ;
comment-line    = ws , comment , eol ;
blank-line      = ws , eol ;
type-name       = identifier , { "." , identifier } ;
identifier      = ident-start , { ident-char } ;
ident-start     = letter | "_" ;
ident-char      = letter | digit | "_" ;
guid            = "{" , 8hex , "-" , 4hex , "-" , 4hex , "-" , 4hex , "-" , 12hex , "}" ;
comment         = "'" , { any-char-except-eol } ;
raw-value       = { any-char-except-eol } ;
ws              = ( space | tab ) , { space | tab } ;
eol             = "\r\n" | "\n" ;
digits          = digit , { digit } ;
letter          = "A".."Z" | "a".."z" ;
digit           = "0".."9" ;
hex             = digit | "A".."F" | "a".."f" ;
```

## Appendix B — Reserved words (single page)

```text
VERSION   Begin   End   BeginProperty   EndProperty   Attribute   True   False
```

Case-insensitive. All other identifiers are application-defined.

## Appendix C — Conformance checklist

- [ ] File is UTF-8 (BOM optional, tolerated).
- [ ] First line is `VERSION <major>.<minor>` (canonical: `VERSION 5.00`).
- [ ] Exactly one root `Begin … End` block.
- [ ] Every `Begin` has a matching `End`; every `BeginProperty` has a matching `EndProperty`.
- [ ] `BeginProperty` bodies contain only properties, comments, and blanks.
- [ ] All `Attribute` lines appear after the root block's `End`.
- [ ] No `Begin`, `BeginProperty`, or `VERSION` after the first `Attribute` line.
- [ ] Line endings are CRLF or LF.
- [ ] Values parse under §5.1 (quote-aware, `'`-trailing comments).
- [ ] Original text, comments, casing, and line endings are preserved for lossless round-trip.

---

*End of specification.*

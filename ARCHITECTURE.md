# Architecture

Developer reference for `pptx_to_accessible_html.py`.

---

## Overview

The converter is a single-file Python script with no framework dependencies beyond `python-pptx`. It reads a `.pptx` file, walks the slide/shape tree, and emits a self-contained HTML file with embedded base64 images and inline CSS.

**Data flow:**

```
.pptx file
  --> python-pptx Presentation object
    --> per-slide: convert_slide() extracts (title, content_parts)
      --> convert_pptx() groups consecutive same-titled slides
        --> merges adjacent lists
          --> wraps in HTML template with inline CSS
            --> writes .html file
```

---

## Key Functions

### Entry Point

| Function | Purpose |
|----------|---------|
| `main()` | CLI argument parsing, batch/single-file dispatch |
| `convert_pptx()` | Orchestrates full conversion: slide iteration, grouping, HTML assembly |

### Slide Processing

| Function | Purpose |
|----------|---------|
| `convert_slide()` | Processes one slide, returns `(title_text, content_html_parts)`. Does NOT wrap in `<section>` -- the caller handles that after grouping. |

`convert_slide` iterates shapes in document order and dispatches by type:

1. **Title placeholders** (type 1, 3) --> extracted as the slide's title string
2. **Slide number placeholders** (type 13) --> skipped
3. **Running headers** (body placeholders matching deck title) --> skipped
4. **Pictures** (`MSO_SHAPE_TYPE.PICTURE`) --> embedded as `<figure><img>` with alt text
5. **Code shapes** (monospace font or name contains "code") --> `<pre><code>`
6. **Text frames** --> `render_text_frame()` for structured HTML
7. **Subtitle placeholders** (type 15) --> `<p class="subtitle">`
8. **Group shapes** --> recursed for nested pictures
9. **Tables** --> `render_table()`
10. **Speaker notes** (if `--include-notes`) --> `<aside>`

### Text Rendering

| Function | Purpose |
|----------|---------|
| `render_text_frame()` | Converts a text frame to HTML with headings, lists, and paragraphs |
| `render_paragraph_runs()` | Converts a single paragraph's runs to HTML, handling `<sup>`, `<sub>`, and citation bracketing |
| `render_code_block()` | Renders a text frame as `<pre><code>`, preserving whitespace |
| `_detect_list_type()` | Returns `"ol"`, `"ul"`, or `""` based on paragraph bullet XML markers |

### Shape Classification

| Function | Purpose |
|----------|---------|
| `is_title_placeholder()` | Placeholder type 1 (TITLE) or 3 (CENTER_TITLE) |
| `is_subtitle_placeholder()` | Placeholder type 15 (SUBTITLE) only |
| `is_slide_number_placeholder()` | Placeholder type 13 |
| `is_running_header()` | Body placeholder (type 2) whose text matches deck title |
| `is_code_shape()` | Shape name contains "code" OR all runs use monospace fonts |
| `shape_has_text()` | Has a text frame with non-empty text |

### Image Handling

| Function | Purpose |
|----------|---------|
| `get_alt_text_and_decorative()` | Returns `(alt_text, is_decorative)` by inspecting `cNvPr` attributes and the decorative extension namespace |
| `get_alt_text_reliable()` | Fallback alt text extractor -- searches all `cNvPr` elements regardless of namespace |
| `image_to_data_uri()` | Converts image bytes to a base64 `data:` URI |

### Post-Processing

| Function | Purpose |
|----------|---------|
| `merge_adjacent_lists()` | Regex-based pass that removes `</ul>\s*<ul>` and `</ol>\s*<ol>` boundaries |

---

## Slide Grouping Algorithm

In `convert_pptx()`, after collecting `(title, parts, slide_number)` for each slide:

1. Iterate slides in order
2. If a slide's title matches the previous slide's title, append its content parts to the existing group
3. Otherwise, start a new group
4. Each group becomes one `<section aria-label="Slide N">` with one `<h2>`
5. `merge_adjacent_lists()` is applied to the combined content of each group

Untitled slides get the fallback title `"Slide N"`, which prevents them from being grouped with other untitled slides (each gets a unique N).

---

## Code Detection Heuristics

A text shape is treated as code (`<pre><code>`) when either:

1. **Shape name**: `shape.name` contains "code" (case-insensitive). This is an explicit author signal.
2. **Font detection**: Every non-empty run in the text frame uses a font from the `MONOSPACE_FONTS` set (30+ common monospace typefaces). The check requires ALL runs to be monospace -- a single proportional run disqualifies the shape.

Font names are compared case-insensitively against the set. If `run.font.name` is `None` (inherited from theme), the run is not counted as monospace.

---

## Citation Detection

In `render_paragraph_runs()`, superscript runs (baseline > 0) are checked against `_CITATION_RE`:

```python
_CITATION_RE = re.compile(r'^[\d,\s\-\u2013]+$')
```

This matches runs containing only digits, commas, spaces, hyphens, and en-dashes -- the pattern used for inline citation numbers like `1,2,3` or `4-6`. Matched runs are wrapped as `<sup>[...]</sup>` instead of plain `<sup>...</sup>`.

Non-matching superscripts (e.g., trademark symbols, ordinals) are left as plain `<sup>`.

Subscripts are never bracketed.

---

## Decorative Image Detection

Two methods are checked, covering different PowerPoint versions:

1. **PowerPoint 365 extension**: `<adec:decorative val="1"/>` inside `<a:extLst>` under `<p:cNvPr>`, using namespace `http://schemas.microsoft.com/office/drawing/2017/decorative`
2. **Legacy/LibreOffice**: `decorative="1"` attribute directly on `<p:cNvPr>`

If either is found, the image is skipped entirely (not embedded in HTML).

---

## List Type Detection

`_detect_list_type()` checks the paragraph's XML properties (`pPr`) for:

- `buAutoNum` --> `<ol>` (auto-numbered list; PowerPoint stores the numbering style, e.g., `arabicPeriod`)
- `buChar` --> `<ul>` (character-bulleted list)
- Neither --> not a list item

When the list type changes between consecutive paragraphs (e.g., `ul` to `ol`), the current list is closed and a new one opened. The `merge_adjacent_lists()` post-processor only merges same-type lists.

---

## OMML to MathML Conversion

PowerPoint equations (inserted via Insert > Equation) are stored as OMML (Office Math Markup Language), a Microsoft XML vocabulary defined in ECMA-376. The converter translates OMML to W3C MathML 3.0 for embedding in HTML5.

### Why OMML shapes are invisible to python-pptx

A text box containing math is wrapped in a `<mc:AlternateContent>` element in the slide XML:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="a14">
    <p:sp>                          <!-- real shape with OMML -->
      <p:txBody>...</p:txBody>
    </p:sp>
  </mc:Choice>
  <mc:Fallback>
    <p:sp>                          <!-- rasterised image fallback -->
      <p:spPr><a:blipFill>...</a:blipFill></p:spPr>
    </p:sp>
  </mc:Fallback>
</mc:AlternateContent>
```

`python-pptx`'s `slide.shapes` iterator walks the `<p:spTree>` looking for `<p:sp>`, `<p:pic>`, etc. as **direct children**. `<mc:AlternateContent>` is a different tag and is silently skipped, so the math shape never appears in the normal shape list.

`convert_slide()` handles this with an explicit secondary pass over `spTree`, finding every `<mc:AlternateContent>` → `<mc:Choice>` → `<p:sp>` and passing the `<p:txBody>` directly to `render_txBody_from_xml()`.

### Inline math within paragraphs

Within an `<a:p>` (DrawingML paragraph), OMML math appears as `<a14:m>` sibling elements interspersed with `<a:r>` text runs:

```xml
<a:p>
  <a:r><a:t>Set of airports </a:t></a:r>
  <a14:m>
    <m:oMath><m:r><m:t>𝒦</m:t></m:r></m:oMath>
  </a14:m>
  <a:endParaRPr/>
</a:p>
```

`_render_para_xml()` walks `<a:p>` children directly — bypassing `python-pptx`'s run abstraction — and dispatches on tag:

- `<a:r>` → text run (with superscript/subscript/citation processing)
- `<a14:m>` → OMML math wrapper, delegates to `omml_to_mathml()`

The `<a14:m>` wrapper contains either `<m:oMath>` (inline) or `<m:oMathPara>` (display/block). The `display` attribute on the emitted `<math>` element is set accordingly.

### `_omml_el` dispatch table

The core recursive converter `_omml_el()` handles each OMML structural element:

| OMML element | MathML output | Notes |
|---|---|---|
| `m:r` + `m:t` | `<mi>`, `<mn>`, or `<mo>` | Classified by `_math_token()` |
| `m:f` | `<mfrac>` | `m:num` → numerator, `m:den` → denominator |
| `m:sSubSup` | `<msubsup>` | `m:e` base, `m:sub`, `m:sup` |
| `m:sSub` | `<msub>` | |
| `m:sSup` | `<msup>` | |
| `m:rad` | `<msqrt>` or `<mroot>` | `<mroot>` when `m:deg` has content |
| `m:d` | `<mrow><mo>(</mo>…<mo>)</mo></mrow>` | Delimiter chars from `m:dPr`; `m:begChr`/`m:endChr`/`m:sepChr` |
| `m:nary` | `<msubsup>` wrapping `<mo largeop>` | Op char from `m:naryPr/m:chr` |
| `m:func` | `<mrow>…<mo>&#x2061;</mo>…</mrow>` | U+2061 = invisible function application |
| `m:limLow` | `<munder>` | |
| `m:limUpp` | `<mover>` | |
| `m:acc` | `<mover>` with accent `<mo>` | Char from `m:accPr/m:chr` |
| `m:bar` | `<mover>` or `<munder>` | `m:barPr/m:pos` = "top" (default) or "bot" |
| `m:groupChr` | `<mover>` or `<munder>` | Char from `m:groupChrPr/m:chr` |
| `m:m` | `<mtable>` | `m:mr` rows, `m:e` cells |
| `m:eqArr` | `<mtable>` | Each `m:e` becomes a `<mtr><mtd>` |
| `m:phant` | `<mphantom>` | |
| `m:box` / `m:borderBox` | `<menclose notation="box">` | |
| `m:sPre` | `<mmultiscripts>` with `<mprescripts/>` | Pre-scripts |
| `m:oMath` / `m:oMathPara` | transparent (recurse into children) | |
| Unknown | recurse into children | Content always preserved |

### Property tag filtering

OMML property elements carry formatting metadata, not renderable content. They are listed in `_M_PROP_TAGS` and return `""` immediately when encountered:

```python
_M_PROP_TAGS = {
    m:rPr, m:sSubSupPr, m:sSubPr, m:sSupPr, m:fPr, m:radPr,
    m:dPr, m:naryPr, m:mPr, m:eqArrPr, m:funcPr, m:limLowPr,
    m:limUppPr, m:accPr, m:barPr, m:groupChrPr, m:ctrlPr, m:phantPr
}
```

Note that `<m:r>` (math run) is handled as a special case, not via its children: `_omml_el()` extracts `<m:t>` directly from the run without recursing, so the `<a:rPr>` DrawingML run-properties element inside `<m:r>` is never processed.

### Math token classification

`_math_token(text, variant)` assigns the correct MathML token element:

1. **`<mn>`** — text matches `\d+(?:\.\d*)?` (a number)
2. **`<mo>`** — single character found in `_MATH_OPS` (operators, relations, punctuation)
3. **`<mi>`** — everything else (identifiers, Greek letters, script/fraktur characters)

Multi-character identifier runs (e.g., `dep`, `arr`, `min`) are emitted as a single `<mi>dep</mi>` — correct for function-name-style subscripts. Single-character math identifiers default to italic in MathML, which is the standard convention.

### `mathvariant` detection

`_get_math_variant(r_el)` inspects `<m:rPr>` inside `<m:r>`:

| OMML | MathML `mathvariant` |
|---|---|
| `<m:nor/>` | `"normal"` (upright) |
| `<m:sty m:val="p"/>` | `"normal"` |
| `<m:sty m:val="b"/>` | `"bold"` |
| `<m:sty m:val="bi"/>` | `"bold-italic"` |
| `<m:scr m:val="cal"/>` | `"script"` |
| `<m:scr m:val="frak"/>` | `"fraktur"` |
| `<m:scr m:val="double-struck"/>` | `"double-struck"` |
| Absent / `<m:sty m:val="i"/>` | `""` (MathML italic default) |

OMML `val` attributes are **namespace-qualified** (`{m_ns}val`), unlike most XML. `_mval()` tries both the qualified and unqualified forms for compatibility across authoring tools.

### Namespace constants

| Python constant | XML namespace |
|---|---|
| `_M_NS` / `_M` | `http://schemas.openxmlformats.org/officeDocument/2006/math` |
| `_A14_NS` / `_A14` | `http://schemas.microsoft.com/office/drawing/2010/main` |
| `_MC_NS` / `_MC` | `http://schemas.openxmlformats.org/markup-compatibility/2006` |
| `_DML_NS` / `_A` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `_PML_NS` / `_P` | `http://schemas.openxmlformats.org/presentationml/2006/main` |

### Graceful degradation

Unknown or future OMML elements fall through to the generic branch which calls `_children_ml(el)`. This recurses into all children, so any text content inside an unrecognised structure is still emitted — as a flat `<mrow>` rather than the correct structure, but never silently lost.

---

## Focusable Math

The `--focusable-math` flag threads a `focusable_math: bool` parameter through the entire call chain down to `omml_to_mathml(focusable=True)`, which adds `tabindex="0"` directly on the `<math>` element:

```html
<math xmlns="http://www.w3.org/1998/Math/MathML" display="inline" tabindex="0">
  …
</math>
```

`tabindex="0"` places the element in the **natural tab order** at its document position, without affecting DOM order. Users pressing Tab reach each equation in reading order.

When this flag is active, `convert_pptx()` also injects companion CSS into the `<style>` block:

```css
math[tabindex="0"] { cursor: default; border-radius: 2px; outline-offset: 3px; }
math[tabindex="0"]:focus { outline: 2px solid #005fcc; background: #f0f4ff; }
```

The CSS is **only emitted when `--focusable-math` is active** — if no `<math>` elements have `tabindex`, the selectors would never match and the bytes would be wasted. Using a conditional variable before the f-string avoids injecting a Python expression inside the large HTML template literal.

---

## MathJax Integration

The `--mathjax` flag causes `convert_pptx()` to inject a single `<script>` tag just before `</head>`:

```html
<script src="https://cdn.jsdelivr.net/npm/mathjax@4/tex-mml-chtml.js" defer></script>
```

**Component choice — `tex-mml-chtml`:** Supports both TeX/LaTeX and MathML input with CommonHTML output. Although the converter produces MathML, this combined component is the recommended MathJax 4 loader and handles any residual TeX notation that may exist in slide text.

**Version pin — `mathjax@4`:** Pinned to the major version so patch and minor updates are picked up automatically but breaking API changes are not.

**`defer`:** Defers script execution until after the document is parsed, equivalent to placing the script at the end of `<body>`. This is preferred over `async` for MathJax because it guarantees the full DOM (including all `<math>` elements) is available before typesetting begins.

**Threading:** `mathjax` does **not** thread past `convert_pptx()`. It only affects the HTML template assembly (a string variable prepended to the `</head>` close tag). Per-equation rendering is unaffected: MathJax discovers and renders all `<math>` elements on page load.

**Browser compatibility without MathJax:** Native MathML is supported in Firefox (all versions), Safari 14+, and Chrome 109+. Without `--mathjax`, the output is fully functional in these browsers with no network requests.

---

## Threading of Parameters

```
main()
  --> convert_pptx(bracket_refs, include_notes, focusable_math, mathjax)
        |
        | [mathjax consumed here — HTML template only]
        |
        --> convert_slide(bracket_refs, focusable_math, deck_title, include_notes)
              --> render_text_frame(bracket_refs, focusable_math)
                    --> render_paragraph_runs(bracket_refs, focusable_math)
                          --> _render_para_xml(bracket_refs, focusable_math)
                                --> omml_to_mathml(display, focusable)
              --> render_txBody_from_xml(bracket_refs, focusable_math)   [AlternateContent path]
                    --> _render_para_xml(bracket_refs, focusable_math)
                          --> omml_to_mathml(display, focusable)
```

`deck_title` is extracted from slide 1's title placeholder in `convert_pptx()` and passed down to `convert_slide()` for running header detection. It does not flow further.

---

## PowerPoint Authoring for Best Conversion Results

### Equations

**Use Insert > Equation, not workarounds.** Equations inserted via the built-in equation editor are stored as OMML and are reliably converted. Equations typed as regular Unicode characters (e.g., copying from a character map) appear as plain `<mi>` tokens at best, or as ordinary text outside any `<math>` element at worst. Screenshots of equations become images that cannot be converted to MathML at all.

**Inline vs. display equations.** Both are handled:
- Inline: an equation inside a text paragraph alongside other text — converted to `<math display="inline">`
- Display: an equation in its own paragraph (or in its own text box) — converted to `<math display="block">`

**AlternateContent detection is automatic.** Every `<mc:AlternateContent>/<mc:Choice Requires="a14">/<p:sp>` in the slide tree is processed. No special shape naming or grouping is required. However, shapes wrapped in non-standard compatibility wrappers (e.g., `Requires="a15"` or other extensions) would not be found by the current spTree walk — if an equation appears missing in the output, check the raw slide XML for unusual wrapper tags.

**Math font encoding.** PowerPoint writes Unicode Mathematical Alphanumeric Symbols (U+1D400–U+1D7FF) directly into `<m:t>` text nodes (e.g., `𝒦` = U+1D4A6, `𝑘` = U+1D458). These are preserved as-is in `<mi>` tokens; the script performs no transliteration. Screen readers that support MathML will announce these correctly.

**Upright vs. italic symbols.** By default, PowerPoint's equation editor sets single letters as italic (conventional for variables). To mark a symbol as upright/roman — conventional for operators, abbreviations, and named constants — select it in the equation editor, then choose *Normal Text* from the equation ribbon. PowerPoint writes `<m:rPr><m:sty m:val="p"/></m:rPr>` which maps to `mathvariant="normal"` on the `<mi>` element.

### Headings and lists in math-containing shapes

Math-containing text boxes are processed through `render_txBody_from_xml()` rather than `render_text_frame()`. The current implementation detects bullet/numbered lists (`buChar` / `buAutoNum`) but does **not** apply the bold+size heading heuristic used for regular text frames. If sub-headings are needed inside a math-containing text box, consider using a separate non-math text box for the heading.

### Code shapes and math

The code-shape detector (`is_code_shape()`) inspects `run.font.name` against a list of monospace faces. Equation editor runs use `Cambria Math` which is not in the monospace list, so math-containing shapes are never misidentified as code blocks.

### Running headers

The running-header filter compares a body placeholder's full text against the deck title from slide 1. Because math content is not returned by `para.text` (python-pptx does not see it), a running header that contains an equation will not be correctly detected and will appear in the output as content. Avoid using equations in running header text boxes.

---

## HTML Output Structure

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <style>/* inline CSS */</style>
</head>
<body>
  <a href="#main-content" class="skip-link">Skip to main content</a>
  <main id="main-content">
    <h1>Deck Title</h1>
    <p class="deck-meta">Converted from ...</p>

    <section aria-label="Slide 1">
      <h2>Slide Title</h2>
      <!-- content: <p>, <ul>, <ol>, <figure>, <table>, <pre>, <aside> -->
    </section>

    <!-- more sections -->
  </main>
</body>
</html>
```

The `<h1>` is derived from the filename. Each `<section>` gets an `aria-label` referencing the first slide number in its group. All images are base64-embedded, making the output fully self-contained.

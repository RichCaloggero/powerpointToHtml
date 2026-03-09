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

## Threading of Parameters

Several options need to flow from CLI through the full call chain:

```
main() --> convert_pptx(bracket_refs, include_notes)
  --> convert_slide(bracket_refs, deck_title, include_notes)
    --> render_text_frame(bracket_refs)
      --> render_paragraph_runs(bracket_refs)
```

`deck_title` is extracted from slide 1's title placeholder in `convert_pptx()` and passed down to `convert_slide()` for running header detection.

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

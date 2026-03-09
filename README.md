# PPTX to Accessible HTML Converter

Converts PowerPoint (`.pptx`) files to accessible HTML, preserving SME-authored alt text, heading structure, images, tables, and lists. Designed for screen reader users, the script strips out decorative images and produces self-contained HTML with embedded images.

---

## Requirements

Python 3.7 or higher and one dependency:

```bash
pip install python-pptx
```

---

## Usage

**Convert a single file:**

```bash
python pptx_to_accessible_html.py presentation.pptx
```

Output will be saved as `presentation.html` in the same folder.

**Specify a custom output filename:**

```bash
python pptx_to_accessible_html.py presentation.pptx -o output.html
```

**Convert an entire folder of `.pptx` files:**

```bash
python pptx_to_accessible_html.py ./my_slides_folder/
```

Each file will produce a matching `.html` file in the same folder.

**Include speaker notes:**

```bash
python pptx_to_accessible_html.py presentation.pptx --include-notes
```

**Disable citation bracket formatting:**

```bash
python pptx_to_accessible_html.py presentation.pptx --no-bracket-refs
```

See [Citation References](#citation-references) for details.

---

## CLI Options

| Flag | Default | Description |
|------|---------|-------------|
| `-o`, `--output` | `<input>.html` | Output file path (single-file mode only) |
| `--include-notes` | off | Include speaker notes as `<aside>` blocks |
| `--no-bracket-refs` | off | Disable `[1,2,3]` bracketing of superscript citation numbers |

---

## What the Script Converts

### Structural Elements

- **Slide titles** --> `<h2>` (one per slide or per group of same-titled slides)
- **Bold/large in-slide text** --> `<h3>`
- **Body text** --> `<p>`
- **Bullet points** (buChar) --> `<ul>` / `<li>`
- **Numbered lists** (buAutoNum) --> `<ol>` / `<li>`
- **Tables** --> `<table>` with `<thead>` and `<th scope="col">`
- **Speaker notes** (optional) --> `<aside class="speaker-notes">`
- **Subtitle placeholders** --> `<p class="subtitle">`

### Images

- **Images with alt text** --> `<img>` with SME-authored `alt` attribute
- **Decorative images** (flagged in PowerPoint) --> skipped entirely
- **Images without alt text** --> embedded with fallback `alt="Image on slide N"` (should be reviewed)
- **Group shapes** --> recursed for nested images

### Code Blocks

Text shapes are detected as code when:
- The shape name contains "code" (case-insensitive), OR
- All text runs use a monospace font (Consolas, Courier New, Fira Mono, etc.)

Code blocks render as `<pre><code>` with original whitespace and indentation preserved.

### Text Formatting

- **Superscript** (baseline > 0) --> `<sup>`
- **Subscript** (baseline < 0) --> `<sub>`

### Citation References

Superscript runs containing only digits, commas, and hyphens (e.g., `1,2,3` or `4`) are recognized as citation numbers. By default, these are wrapped in brackets for clarity:

- PowerPoint: `compare`^`1,2,3` --> HTML: `compare<sup>[1,2,3]</sup>`

This makes the association between inline citations and numbered reference lists clear for screen reader users who cannot see the raised positioning. Disable with `--no-bracket-refs`.

---

## Consecutive Slide Grouping

When multiple consecutive slides share the same title, they are merged into a single `<section>` with one `<h2>`. This avoids redundant headings and enables list merging across slide boundaries. Each slide's content is concatenated within the shared section.

Adjacent lists of the same type (`<ul>` or `<ol>`) within a grouped section are automatically merged into a single list.

---

## Filtered Elements

The following PowerPoint elements are automatically excluded from output:

- **Slide number placeholders** (placeholder type 13)
- **Running headers** -- body placeholders (type 2) whose text matches the deck title from slide 1 (e.g., a repeating course name on every slide)
- **Decorative images** -- images flagged as decorative via PowerPoint's alt text panel

---

## Alt Text and Decorative Images

The script reads alt text exactly as authored in PowerPoint -- no AI generation or modification.

**To add alt text in PowerPoint:**

Right-click an image --> **Edit Alt Text** --> type a description.

**To mark an image as decorative:**

Right-click an image --> **Edit Alt Text** --> check *Mark as decorative*.

The script will skip decorative images entirely and omit them from the HTML output.

Images with neither alt text nor a decorative flag receive a fallback description of:

> "Image on slide N"

These should be reviewed and updated in the source PowerPoint before final conversion.

---

## Image Sizing

Images are rendered at their actual PowerPoint dimensions, converted to pixels at 96 DPI.
A `max-width: 100%` rule ensures they scale down gracefully on narrow screens without distortion.

---

## Accessibility Features

- Skip navigation link ("Skip to main content") at the top of every page
- Semantic HTML5 landmarks (`<main>`, `<section>`, `<aside>`)
- `aria-label` on each slide section
- Table headers use `scope="col"` for screen reader compatibility
- Superscript/subscript preserved as `<sup>`/`<sub>` for correct screen reader announcement
- Citation numbers bracketed for non-visual clarity
- Self-contained output -- images embedded as base64 data URIs, requiring no external assets

---

## Auditing Images Before Converting

Use the companion script `inspect_pptx_images.py` to audit every image in a deck:

```bash
python inspect_pptx_images.py presentation.pptx
```

This prints each image's name, alt text, and decorative status -- useful for catching missing or incomplete alt text before generating the final HTML.

---

## Recommended Workflow

1. SMEs author alt text and mark decorative images in PowerPoint
2. Run `inspect_pptx_images.py` to verify all images are correctly tagged
3. Run `pptx_to_accessible_html.py` to generate the HTML
4. Test with a screen reader

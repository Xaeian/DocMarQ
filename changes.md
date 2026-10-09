# Changes `docmarq`

## `0.4.0` Legal documents

- Breaking: `md_to_docx` and the banner moved to `docmarq.paper`, their fields to `PaperStyle`
- Breaking: the markdown engine works in mm only
- `docmarq.legal`: letters, contracts and language versions side by side
- `docmarq.paper.md_into` writes after a cover
- Core: `doc.columns()`, `doc.into()`, `doc.hyphenation()`, `doc.border()`, Word fields
- Text: `justify`, `hyphenate`, `keep_together`, small capitals, heading sizes and gaps
- Page number as `Page 1/5`, as in `pdfmarq`
- Fix: `1)` printed as `1.`, a title list garbled in metadata, image over `---` lost as a heading
- Fix: two tables in a row merged, repeated headings sharing a bookmark, gutter ignored in widths
- Fix: table borders out of schema order, justified line before a hard break stretched

## `0.3.7` Table widths

- Table columns sized by content, as in `pdfmarq`

## `0.3.6` SVG fonts

- SVG text uses document fonts, via `register_fonts`

## `0.3.5` Math input

- LaTeX math delimiters `\[...\]` and `\(...\)` accepted
- `\tag{...}` becomes a right-aligned equation label, Word style

## `0.3.4` Image sizing

- Block images fill text width down to `image_min_dpi`
- Images alone in a paragraph become a row of figures, wrapping into a grid
- Mermaid diagrams size like any other figure
- Block images centred by default, matching `pdfmarq`
- Fix: `scale=` assumed 96 DPI

## `0.3.3` Numbering & lists

- Ordered lists keep author's numbering, `ordered(start=)`
- SVG rasterizer moved to `docmarq.svg`, usable without markdown
- `mermaid_remote=False` keeps diagram source off the network
- Fix: paragraph properties in OOXML order, blocks inside list items

## `0.3.2` Naming

- Style fields renamed: `font_body`, `banner`, `banner_compact`, `sign`
- Sign scenarios from `sign_labels`

## `0.3.1` Sharp SVG

- SVG → PNG at higher resolution
- `svglib`, `rlPyCairo`, `Pillow` as core deps

## `0.3.0` Caller-owned styling

- Breaking: `render:` block dropped, `page=A4` replaces `width` / `height` / `landscape`, `para_gap_pt` → `para_gap`
- Signature block, `entity` / `address`, duplicate-title skip
- Fixes across Word styles, tables, images and metadata

## `0.2.2` Inline code styling

- Inline `code` honors `mono_family` + code colors
- Warnings: code block in blockquote, missing `PyYAML`

## `0.2.1` SVG deps

- `svglib` + `rlPyCairo` included in `docmarq[md]`

## `0.2.0` Math & gutter

- Math formulas → native Word equations _(OMML)_ + image fallback
- `render:` keys `gutter`, `header`
- Wider paragraph spacing

## `0.1.1` Fonts & mermaid

- `.svg` image conversion
- `head_family` for heading typography
- Frontmatter `logo:` via `base_dir`
- `render:` keys
- Mermaid: info-string DSL, `font_body` labels

## `0.1.0` Initial release

Fluent `.docx` generation.
The `[md]` extra renders markdown with banner headers, mermaid diagrams and GitHub callouts.

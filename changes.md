# Changes `docmarq`

## `0.2.2` Inline code styling

- Inline `code` honors `mono_family` + code colors
- Warnings: code block in blockquote, missing `PyYAML`

## `0.2.1` SVG deps

- `svglib` + `rlPyCairo` included in `docmarq[md]`

## `0.2.0` Math & gutter

- Math formulas → native Word equations _(OMML)_ + image fallback
- `render:` keys `gutter`, `header`
- Wider paragraph spacing _(PDF parity)_

## `0.1.1` Fonts, mermaid & render keys

- `.svg` image conversion
- `head_family` for heading typography
- Frontmatter `logo:` via `base_dir`
- New `render:` keys
- Mermaid: info-string DSL + `font_body` labels

## `0.1.0` Initial release

Initial release of `docmarq`, a companion tool to [`PDFMarQ`](https://github.com/Xaeian/PDFMarQ).
Like `PDFMarQ`, it converts Markdown documents, but targets `.docx` output instead of PDF.
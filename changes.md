# Changes `docmarq`

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

Fluent `.docx` generation. The `[md]` extra renders markdown with banner headers,
mermaid diagrams and GitHub callouts.

# Changelog

All notable changes to this project will be documented in this file.

## [0.3.1] - 2026-06-29

### Fixed

- CLI `--help` now reliably displays usage examples by placing them in the Typer `help=` argument.

## [0.3.0] - 2026-06-29

### Added

- New `docubridge extract-styles` command to turn any `.docx` into a reusable YAML style profile.
- New built-in `default` style profile with Chinese `宋体`, English `Times New Roman`, single line spacing and no first-line indent.
- Support for `line_spacing` / `line_spacing_pt` in style profiles and rendering.
- Added `__version__` and `docubridge --version` / `-v` flag.
- Added structured release artifacts: source distribution (sdist) and wheel.
- Lowered minimum Python version from 3.12 to 3.10.

### Fixed

- Run-level font size, bold and italic are now correctly applied when rendering Markdown to Word.
- Heading color fidelity when using `--template` is preserved; Heading 1/2 stay black when the original style has no color.
- `--template` now strips original document body content and only loads styles/numbering/theme/sections.

### Changed

- Improved `--help` output for every command with detailed descriptions and multiple usage examples.
- Updated `academic` and `business` built-in styles to use `font_ascii` / `font_east_asia` and single line spacing.
- README and README_CN refreshed with `extract-styles`, `default` style and 0.3.0 release notes.

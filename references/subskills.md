# PPT Skill Child Map

The package keeps task-specific child skills under `references/subskills/` so the aggregate entry point remains portable.

| Child | Main use | Modes |
|---|---|---|
| `taskpkg-pptx-official` | PPTX parse/edit/render/schema | full-project, deck-only, qa-only |
| `taskpkg-pptx-community` | competition narrative and presentation quality | full-project, deck-only |
| `taskpkg-cyberppt` | grid, typography, palette, editable dense layouts | full-project, deck-only |
| `taskpkg-docx-official` | DOCX structure, styles, tables, export | full-project, document-only, qa-only |
| `taskpkg-docx-community` | long-form planning document review | full-project, document-only |
| `taskpkg-pdf-official` | PDF export and page/text QA | full-project, document-only, qa-only |
| `taskpkg-imagegen` | generated/editable raster assets | full-project, asset-only |
| `taskpkg-screenshot` | screenshot capture and evidence | full-project, deck-only, qa-only |
| `taskpkg-uiux-pro-max` | visual hierarchy and accessibility-informed polish | full-project, deck-only, qa-only |
| `taskpkg-skill-harness` | offline audit guidance | full-project, qa-only |
| `taskpkg-skill-eval-harness` | paired evaluation guidance | full-project, qa-only |

## Loading Rule

Read only the child references needed for the selected mode. For a full project, read in this order: PPTX official, PPTX community, CyberPPT, DOCX official/community, PDF, imagegen, screenshot, UIUX, then harness. Do not load both official and community child guidance twice from different install roots; `references/subskills/` is the portable copy.

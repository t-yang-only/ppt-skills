# PPT Skill QA Checklist

## Inputs And Traceability

- [ ] Original PPTX/DOCX/PDF and requirements are copied unchanged.
- [ ] SHA-256 or equivalent hashes are recorded.
- [ ] Source page/slide counts, dimensions, and flattened-image limitations are recorded.
- [ ] Every externally sourced or project-material figure has source/year/unit/sample/method or `项目材料工作口径` / `UNMEASURED`.

## PPTX

- [ ] PowerPoint/LibreOffice can open the file.
- [ ] Slide ratio is uniform and content stays within the canvas.
- [ ] Text, shapes, charts, and notes remain editable where native representation exists.
- [ ] No overflow, overlap, clipping, missing glyphs, low-resolution images, placeholder text, or unlabeled chart axes.
- [ ] Every slide has a purposeful visual element and the visual hierarchy changes appropriately across sections.

## DOCX/PDF

- [ ] Headings and table structure are usable; tables do not break across pages unexpectedly.
- [ ] Header/footer/page numbers are stable.
- [ ] DOCX and PDF have matching content and member names.
- [ ] PDF text is selectable where expected, and rendered pages are legible.

## Iteration And Packaging

- [ ] When required, three complete render/check/revise rounds are stored, not just selected screenshots.
- [ ] Every generated asset has a prompt, timestamp, dimensions, intended use, realism/clarity result, and keep/discard decision.
- [ ] Harness results include commands, versions, findings, and explicit `UNMEASURED` items.
- [ ] Final directory contains exactly the requested final artifacts; process directory contains only relevant evidence and no credentials.

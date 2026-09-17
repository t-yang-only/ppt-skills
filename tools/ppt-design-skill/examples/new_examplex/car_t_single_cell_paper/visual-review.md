# Visual review — blue editorial final

## Gate 1: visual effect

**PASS**

The deck now uses a BioTech Blue Research Editorial system: light blue field,
deep navy typography, cobalt structure, and cyan data paths with orange/lime
validation accents. The cover uses a dedicated conceptual CAR-T hero, while the
body uses a four-zone grid so the paper figures support—not replace—the narrative.
The second visual pass enlarged body typography, replaced repeated full-figure
tiles with actual Fig.2/3/4/6 crops, and filled the evidence cards with source-backed
experimental conditions and cohort details.
The third pass removed rounded containers and pink accents, and replaced the cover
with a purpose-built blue scientific visualization.
The fourth pass converted the clinical, study-design, atlas, persistence and in-vivo
pages from explanatory cards into timelines, process bands, state structures, gradients
and experimental sequences. The final merge promoted the approved P1 clinical-
persistence page and P2 study-design pipeline into the full 12-page deck.

## Gate 2: serious defects

| Requirement | Status | Evidence |
|---|---|---|
| R1 | PASS | 12 slides, 13.333 × 7.5 inches, rendered to PNG/PDF |
| R2 | PASS | Paper and figure references appear in source captions |
| R3 | PASS | Cohort, cell counts, BCA strata, cytokine scale, doses and animal model follow the paper |
| R4 | PASS | Clinical, cellular, mechanistic and preclinical evidence are separated |
| R5 | PASS | Most pages use a four-zone information grid with 3–5 evidence units |
| R6 | PASS | Direct paper figures are labeled; editorial summaries do not invent new data |
| R7 | PASS | Blue scientific palette, readable light-background contrast, enlarged body type, no pink or rounded containers |
| R8 | PASS | Figure modules, metrics, method notes, evidence boundaries and mechanism summary vary intentionally |
| R9 | PASS | Final page states the contribution and retains the preclinical boundary |

## QA evidence

- PPTX structural inspection: 12 slides, 16:9, expected picture/text/shape mix;
- PPTX → PDF → PNG render: `rendered/slide01.png` through `slide12.png`;
- All 12 PNGs visually reviewed at 1280 × 720 after the final merge; slides 6
  and 11 were rechecked after removing the rectangular evidence containers.
- Revision 5: P03 output metric spacing was corrected after review found the
  `695,819` value touching the `analyzed cells` label; the PPTX was regenerated,
  re-rendered, and P03 was visually rechecked.

## Acceptance Record pre-review — Revision 6

The full 12-page PNG set was reviewed against the acceptance contract. No hard
overflow, page-number defect, accidental clipping, or evidence-boundary breach
was found. The primary content audit also reconciled the headline cohort and
intervention facts against Bai et al., *Nature* 634, 702–711 (2024).

| Page | Primary evidence check | Visual disposition |
|---|---|---|
| 01 | Cover states the paper, journal, year, and clinical persistence premise. | PASS; conceptual cover is clearly marked. |
| 02 | `8.4 years` is tied to the BCA-L clinical persistence group. | PASS; clinical observation is separated from mechanism. |
| 03 | `82 patients + 6 healthy donors`, `695,819 cells`, and `17 states` are shown with the study pipeline. | PASS; pipeline is legible and editable. |
| 04 | Fig. 1 UMAP and cluster annotation are retained; original figure citation is visible. | TARGETED READABILITY CHECK; dense labels are legible only at presentation scale after zoom. |
| 05 | Five BCA persistence groups and BCA-L `n = 5` are separated from the interpretation. | PASS; duration units and source note are visible. |
| 06 | Fig. 2 evidence supports type-2 signature, secretion, chromatin, and perturbation framing. | TARGETED READABILITY CHECK; multiple cropped panels make this the densest page. |
| 07 | Fig. 3 ligand–receptor result and `13.9%` cluster-2 context are framed as a proposed mechanism. | PASS; causal wording remains bounded. |
| 08 | Fig. 4 serum profiling is labeled with `345 measurements`, `30 cytokines`, and `33 patients`. | TARGETED READABILITY CHECK; figure labels/caption need second-reviewer confirmation. |
| 09 | Fig. 5 mouse expansion, rechallenge, and survival are explicitly labeled preclinical. | PASS; no clinical efficacy claim. |
| 10 | Fig. 6 IL-4 priming/manufacturing conditions and ET2-L/H groups are labeled as design options. | TARGETED READABILITY CHECK; small experimental labels need second-reviewer confirmation. |
| 11 | Association → cellular state → proposed mechanism → preclinical test is separated in an evidence chain. | PASS; evidence levels are explicit. |
| 12 | Contribution and boundary are stated without extending the paper beyond its evidence. | PASS; final takeaway is appropriately bounded. |

Current primary-review conclusion: `BLOCKED` pending second Reviewer review of
P01, the four targeted-readability pages (04/06/08/10), and all MUST rows. This
is a controlled review state, not a claim that the deck is already signed off.

# Adaptive Group — HTML Presentation Brand Guide
*Derived from PowerPoint training template. All measurements reference a 1280×720px canvas.*

---

## CANVAS & GRID

* **Dimensions:** 1280 × 720 px (16:9)
* **Source EMU:** 12 192 000 × 6 858 000
* **Aspect ratio:** always maintained via `padding-top: 56.25%` in HTML
* **Coordinate system:** all positions expressed as % of 1280 (width) and 720 (height)
* **Footer safe zone:** bottom 3.2% (23px) reserved for footer bar on all content/agenda/closing slides
* **Content safe zone:** top 13.5%, left from sidebar edge, bottom 3.2%

---

## COLOR SYSTEM

### Primary Palette

| Name | Hex | Usage |
|---|---|---|
| Primary Gold | `#B8975A` | Dominant brand accent |
| Template Gold | `#B59859` | Exact template value — icon fills, title bands |
| Footer Gold | `#C8B182` | Footer bar, active agenda icon |
| Divider Gold | `#B09247` | Gold divider lines under titles |
| Professional Black | `#000000` | Sidebar, title bars, opening/closing backgrounds |
| Charcoal | `#2C2C2C` | Body text on light slides |
| Slide Background | `#F9F9F9` | Content and agenda slide background |
| Icon Background | `#F5EFE2` | Inactive agenda icon fill |
| Hex Accent | `#F5F2EA` | Bottom-right corner hex decorations (light slides) |

### Accent & Status Colors

| Name | Hex |
|---|---|
| Success Green | `#4CAF50` |
| Error Red | `#F44336` |
| Warning Orange | `#FF9800` |
| Info Blue | `#2196F3` |
| Neutral Gray | `#9E9E9E` |

### Gradients

* **Gold Horizontal:** `linear-gradient(90deg, #D4C5A0 0%, #B8975A 50%, #D4C5A0 100%)`
* **Gold Vertical:** `linear-gradient(180deg, #D4C5A0 0%, #B8975A 100%)`
* **Dark (opening/closing):** `linear-gradient(135deg, #1A1A1A 0%, #000000 100%)`

---

## TYPOGRAPHY

### Font Stack

| Role | Template Font | Web Fallback |
|---|---|---|
| Headings / Title bands | Arial Nova Cond | Copperplate, Georgia, serif |
| Body / Agenda items | Arial Nova Cond | Open Sans, Source Sans 3, sans-serif |

### Size Scale (at 1280px canvas)

| Element | px | vw |
|---|---|---|
| Training Title band | 37px | ~1.6vw |
| Subtitle band | 27px | ~1.1vw |
| AGENDA wordmark | 37px | ~1.5vw |
| Content slide title | 35px | ~1.4vw |
| Body text | 19–21px | ~1.1vw |
| Agenda item text | 21px | ~0.9vw |
| Trainer name / role | 16px | ~0.9vw |
| Footer copyright | 9px | ~0.55vw |

* Minimum font size on any element: **12px**
* Use `vw` units so text scales proportionally with the slide container
* Title bands: font-weight 700, letter-spacing 0.05–0.15em
* Body: font-weight 400–600

---

## OPENING SLIDES

### Background & Structure

* Full-bleed black background `#000000`
* Three gold hexagons cascade diagonally from top-left corner
* Two smaller gold hexagons in top-right corner
* Bottom footer bar `#C8B182`, height 3.2%

### Component Positions

| Component | Left | Top | Width | Height | Color |
|---|---|---|---|---|---|
| Background | 0 | 0 | 100% | 100% | `#000000` |
| Hex TL-1 | 1.5% | 1% | 7.2% | 10.1% | `#B59859` |
| Hex TL-2 | 8.1% | 7.3% | 7.2% | 10.1% | `#B59859` |
| Hex TL-3 | 15.2% | 13.5% | 7.2% | 10.1% | `#B59859` |
| Hex TR-1 | right 4.5% | 1.3% | 3.6% | 5.4% | `#B59859` |
| Hex TR-2 | right 2.5% | 6.8% | 3.1% | 4.4% | `#B59859` |
| Training Title band | 30.5% | 21.5% | 39% | 9% | `#B59859`, white text |
| Subtitle band | 30.5% | 31.5% | 39% | 9% | `#B59859`, white text |
| Date / Location | 30.5% | 44.9% | 39% | 5.7% | white 85% opacity, no fill |
| Adaptive logo block | 34.3% | 58% | 27.8% | 32% | transparent |
| Client logo placeholder | 76.5% | 56.3% | 14.7% | 25.4% | white bg |
| "for" label | 27.4% | 73.9% | 5.9% | 6.7% | `#B59859` |
| "by" label | 66.5% | 74.3% | 5.9% | 6.7% | `#B59859` |
| Trainer Name | 68.9% | 83.2% | 29.9% | 5.8% | white |
| Role | 68.9% | 89.6% | 29.9% | 5.8% | white 55% opacity |
| Footer bar | 0 | bottom 0 | 100% | 3.2% | `#C8B182` |

### Variants

* **Title_Logos** — Adaptive logo (left) + client logo (right) + one trainer
* **1_Title_Logos** — Client logo only (no Adaptive logo area on left), one trainer
* **1_Title_Logos_2Trainers** — Two trainer blocks side by side; title moves up to T 13.5%; two photo placeholders at L 22.1% and L 54.1%, each 22% w × 24% h; "for" and "by" gold labels between them

---

## AGENDA SLIDES

### Variants Overview

| Variant | Items | Black Panel | Icon Size | AGENDA Label |
|---|---|---|---|---|
| Agenda / 3_Agenda | 4 | None (light bg throughout) | 4.9% w × 7.6% h | Top-left dark block, L 20.4% |
| agenda items (inline) | 4 | None | 7.5% w × 11.3% h | Top-left dark block, L 12.3% |
| 5_Agenda5 | 5 | Left panel 58.5% w | 8.9% w × 15.7% h | Right white area, L 70.1% |
| 6_Agenda3 | 3 | Left panel 58.5% w | 8.9% w × 15.7% h | Right white area, L 70.1% |
| 6_Agenda8 | 8 | Left panel 58.5% w | 6.5% w × 11.1% h | Right white area, L 70.1% |

### Standard 4-item Agenda (Agenda / 3_Agenda)

* Background: `#F9F9F9`
* Two small gold hexagons top-left: L 1.5% / T 1% and L 7.1% / T 7%
* AGENDA title block (black): L 20.4%, T 0.4%, W 16.9%, H 22.9% — white Cinzel bold text
* Gold divider line below AGENDA: L 21.1%, T 17.9%, W 7%
* 4 icon+text rows starting at T 27.1%, pitch ~13.3% per item
* AGENDA word repeated right side as decorative hex pair: L 79.9% and L 84.3%
* Bottom-right hex pair (cream `#F5F2EA`): R 7.7% and R 0.4%
* Footer bar `#C8B182`

### Left-Panel Agenda (5_Agenda5, 6_Agenda3, 6_Agenda8)

* Black left panel: L 0, W 58.5%, H 96.8%
* Gold hex icons inside black panel, L 0.8–0.9%, pitch ~18% per item
* Item titles in white text inside the black panel, L 11%, each row H 9.2%
* AGENDA wordmark block (right white area): L 70.1%, T 35%, W 16.9%, H 22.9%
* Gold divider line below wordmark: L 70.8%, T 52.5%
* Top-right gold hex pair: same as content slides

### Agenda Icon — States

| State | Fill | Number Color | When Used |
|---|---|---|---|
| Active | `#C8B182` | `#000000` | Current / first item |
| Inactive | `#F5EFE2` | `#2C2C2C` | All other items |

### Agenda Icon Sizes by Variant

| Variant | Width | Height | Vertical pitch |
|---|---|---|---|
| Agenda / 3_Agenda | 4.9% | 7.6% | ~13.3% |
| 5_Agenda5 / 6_Agenda3 | 8.9% | 15.7% | ~18% |
| 6_Agenda8 | 6.5% | 11.1% | ~11.7% |
| agenda items (inline) | 7.5% | 11.3% | ~12.5% |

---

## CONTENT SLIDES

### Background & Structure

* Background: `#F9F9F9`
* Black left sidebar from top to 96.8% height
* Black title bar spanning from sidebar right edge to ~92% of width
* Gold divider line (1px) directly below title bar
* Gold hex pair in top-right corner
* Footer bar `#C8B182`

### Component Positions

| Component | Position | Size | Color |
|---|---|---|---|
| Background | full | 100% × 100% | `#F9F9F9` |
| Black left sidebar | L 0, T 0 | 10.8% w × 96.8% h | `#000000` |
| Hexagon icons in sidebar | L 0.8–0.9% | 8.9% w, vertical rhythm | `#B59859` |
| Title bar | L 12.3%, T 0 | 79.4% w × 13.5% h | `#000000`, white text |
| Gold divider line | L 13%, T 13.5% | 7% w × 1px | `#B09247` |
| Content safe zone | L 13%, T 15%, R 5%, B 4% | ~82% × ~78% | — |
| Top-right hex-1 | right 4.5%, T 1.3% | 3.6% × 5.4% | `#B59859` |
| Top-right hex-2 | right 8.3%, T 4% | 3.7% × 5.4% | `#B59859` |
| Footer bar | bottom 0 | 100% × 3.2% | `#C8B182` |
| Footer logo zone | right 0, bottom 0 | 8.4% × 3.2% | Adaptive lockup |

### Sidebar Variants

| Layout Name | Sidebar Width | Hexes in Sidebar | Title Bar Starts |
|---|---|---|---|
| 1_Custom Layout3 | 10.8% | 3 hexes | L 12.3% |
| 1_Custom Layout5 | 10.8% | 5 hexes (full height) | L 12.3% |
| Custom Layout8 | 8.6% | 8 smaller hexes (6.5% w) | L 12.3% |
| 1_Section (divider) | 10.8% | 5 hexes | Section title centered L 13.9% |
| 2_Custom Layout | None | None | Title at L 5.5%, top 0 |

### Hexagon Vertical Positions in Sidebar (1_Custom Layout3)

| Hex # | Top |
|---|---|
| 1 | 21.8% |
| 2 | 39.9% |
| 3 | 57.9% |

For 5-hex variant: continue at 76.0% with the same ~18% pitch.

### Section Slide (Divider between topics)

* Background: `#F9F9F9` — same as content
* Black left sidebar identical to content slides
* Section title block: L 13.9%, T 40.3%, W 72.2%, H 19.3% — black fill, white Cinzel bold
* Gold divider line: L 14.8%, T 59.7%
* Bottom-right cream hex pair (`#F5F2EA`)

---

## CLOSING SLIDES

### Thank You — 1 Trainer ("end")

| Component | Position | Size | Color |
|---|---|---|---|
| Background | full | 100% × 100% | `#000000` |
| Hex TL-1/2/3 | same as opening | 7.2% × 10.1% | `#B59859` |
| Photo / logo placeholder | L 10.1%, T 12.7% | 79.7% × 74.6% | transparent |
| **THANK YOU! band** | L 24.8%, T 28.3% | 50.4% × 14.7% | `#B59859`, white Cinzel bold |
| Email icon placeholder | L 1.5%, T 59.7% | 8.1% × 14.4% | transparent |
| Email address text | L 12.3%, T ~63.5% | 36.9% × 6.9% | white |
| LinkedIn placeholder | right ~20%, bottom 12.8% | 12.7% × 21.5% | white/light bg |
| Hex BR-1 | right 6.9%, bottom 12.8% | 3.6% × 5.4% | `#B59859` |
| Hex BR-2 | right 3.1%, bottom 7.3% | 3.1% × 4.4% | `#B59859` |
| Footer bar | bottom 0 | 100% × 3.2% | `#C8B182` |

### Thank You — 2 Trainers ("ThankYou_2Trainers")

* Black top panel covers only 56.5% height (not full)
* THANK YOU band: L 31.5%, T 28.3%, W 36.9%, H 12.5%
* Trainer 1 photo: L 72.1%, T 58.2%, W 12.7%, H 21.5%
* Trainer 2 photo: L 86.4%, T 58.3%, W 12.7%, H 21.5%
* Email icon row 1: L 1.5%, T 59.7%
* Email address row 1: L 11.1%, T 63.5%, W 36.9%
* Email icon row 2: L 1.5%, T 77.6%
* Email address row 2: L 11.1%, T 81.3%, W 36.9%
* Trainer Name 1: L 72.1%, T 81.3% / Trainer Name 2: L 86.4%, T 81.2%

### Questions Slide

* Background: `#F9F9F9` (light)
* "Questions?" text: L 24.5%, T 41%, W 51%, H 18%
* Gold divider line: L 33.8%, T 59%
* Chat icon illustration: right 39%, T -4.5%, W 39% — intentionally overflows top
* Bottom-right cream hex pair (`#F5F2EA`)

### SurveyCode Slide

* Black left half: L 0, W 53.4%, H 96.8%
* Three TL hexagons: same cascade pattern, `#B59859`
* "SCAN & RATE" text: L 9%, T 33.5%, W 36.9%, H 26.2% — gold fill, white Cinzel bold
* QR placeholder: L 42.3%, T 54.8%, W 7%, H 32.2%
* White right panel: L 54.9%, T 8.4%, W 43.5%, H 82.2%
* Two gold hexes BR: `#B59859`

---

## SHARED COMPONENTS

### Hexagon Shape

```
clip-path: polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%);
```

Use `padding-top` to control height proportionally from width. Text inside requires an `position:absolute; inset:0` child with flexbox centering.

### Hexagon Decoration Contexts

| Context | Color | Size | Position |
|---|---|---|---|
| TL cascade (opening/closing) | `#B59859` | 7.2% × 10.1% | 3-step diagonal: L 1.5→8.1→15.2%, T 1→7.3→13.5% |
| TR micro pair (content/agenda) | `#B59859` | 3.6–3.7% × 5.4% | Right 4.5% and 8.3%, T 1.3–4% |
| Sidebar icons (agenda/content) | `#B59859` | 8.9% × 15.7% | L 0.8–0.9%, vertical rhythm ~18% |
| BR decor (light slides) | `#F5F2EA` | 5.6–9.6% × 9.1–15.3% | Right 0.4–7.7%, bottom 4.6–12.2% |
| BR decor (dark slides) | `#B59859` | 3.1–3.6% × 4.4–5.4% | Right 3.1–6.9%, bottom 7.3–12.8% |

### Gold Divider Line

* Width: 7% of canvas (90px)
* Height: 1px
* Color: `#B09247`
* Placement: directly below title bar or AGENDA wordmark block, left-aligned with content text

### Footer Bar

* Height: 3.2% (23px)
* Color: `#C8B182`
* Text: copyright notice, ~0.55vw, `rgba(0,0,0,0.6)`
* Logo zone: right 8.4%, same height — Adaptive wordmark
* Present on: all content, agenda, and closing slides

### AGENDA Wordmark Block

* Size: W 16.9% × H 22.9%
* Fill: `#000000`
* Text: "AGENDA", white, Cinzel Bold 700, letter-spacing 0.12em, centered
* Always followed by gold divider line (`#B09247`) at ~7% below block bottom

---

## SPACING REFERENCE

| Measurement | % | px |
|---|---|---|
| Left sidebar width (standard) | 10.8% | 138px |
| Left sidebar width (narrow) | 8.6% | 110px |
| Title bar height | 13.5% | 97px |
| Title bar left edge | 12.3% | 157px |
| Title bar width | 79.4% | 1017px |
| Content start (top, below divider) | ~15–18% | 108–130px |
| Content start (left) | ~13–16% | 166–205px |
| Black agenda panel width | 58.5% | 749px |
| Agenda item pitch (4-item) | ~13.3% | 96px |
| Agenda item pitch (3/5-item) | ~18% | 130px |
| Footer height | 3.2% | 23px |

---

## DESIGN PRINCIPLES

* Gold used sparingly — as accent on black backgrounds and structural bands, not as a fill for large content areas
* Hexagonal pattern is the core visual motif — repeat consistently, never in isolation
* Black and white are the two structural colors — black for panels/sidebars/bands, white/near-white for content areas
* Generous white space in the content area — the black sidebar creates visual weight so the content zone must breathe
* Shadow: `0px 2px 6px rgba(0,0,0,0.10)` — for card-style elements
* Border radius: 8px for cards, 4px for buttons
* High contrast always — never light text on light background or dark text on dark background

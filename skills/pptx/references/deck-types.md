# Deck types: what changes between consulting, academic, product, and internal decks

Genre decides almost everything mechanical: how dense a slide may be, whether titles state conclusions, how many slides a 20-minute slot holds, and whether the deck must survive being read without you. Decide the genre **before** the outline, and design the design system second.

Two axes drive most differences:

- **Presented vs read.** A talk is performed; a consulting deck is read (circulated, skimmed, opened alone). Read decks need action titles and self-contained slides; presented decks may lean on your voice.
- **Goal.** Decide → convince → teach → report. A deck that tries to do two of these does neither.

## Consulting / strategy decks

**Goal & audience.** Drive a scarce-attention decision-maker to act. Read-first documents: they must stand alone [C2][C3].

- **Spine:** answer first — governing thought → 3–4 MECE arguments → evidence (pyramid principle) [C1][C2], set up with SCQA (Situation–Complication–Question–Answer) [C2][C3]. Standard blocks: cover · 1-page executive summary (the whole answer) · body · recommendation/next steps · appendix [C3].
- **Titles:** action titles — a full sentence stating the takeaway, never a topic like "Market overview". The storyline must be readable from the titles alone, so write the title list first, then build slides [C3][C4].
- **Density:** one message per slide; the body proves only the title's claim; ~20% whitespace; one font, one dominant color, gray shades, one accent; source line and page number on every slide; no animation [C3][C4].
- **Charts by intent:** bar = compare, line = trend, waterfall = drivers, 2×2 = positioning, flow = process, pie = composition (≤ 5 slices) [C4].
- **Pitfalls:** bottom-up narration that buries the answer; topic titles; two messages on one slide; data with no "so what?"; recommendation only at the end; missing executive summary.

**Outline (recommendation deck):** 1 cover · 2 executive summary · 3 situation · 4 complication · 5 key question · 6 recommendation · 7–9 MECE arguments + evidence · 10 impact sizing · 11 risks · 12 roadmap (owners/dates) · 13 next 30 days · 14+ appendix [C3][C4].

## Academic talks (conference / seminar / defense)

**Goal & audience.** A talk is a taster for the work, not the paper. The audience is smart but outside your area; your job is one remembered takeaway [A1].

- **Spine:** Background → Purpose → Methods → Results → Discussion → Conclusion (keep this order) [A4]; build tension with repeated problem→solution ("but" / "therefore") cycles and close by restating the finding [A3]. Skip a contents slide; one mid-talk signpost is enough [A1].
- **Titles:** state the *result*, not the topic [A3].
- **Density:** ≤ 6–7 lines per slide; body ≥ 17–18 pt; figures that read from the back; 3 colors total (base + accent + support); don't rely on red-vs-green alone [A1][A4][J2].
- **Budget:** ~1 slide per minute — a 15-minute talk is ~12–15 slides, a 20-minute talk ~18–22 [A4]. Never overrun; keep backup slides for Q&A [A1].
- **Pitfalls:** staying abstract with no concrete example (the #1 failure); reading slides; opening with an agenda/TOC; hiding known problems; gluing old slides into a new talk; overrunning [A1][A5].

**Outline (15-min conference talk):** 1 title stating the result · 2 motivating problem/example · 3 minimal background · 4 gap ("but") · 5 approach · 6–9 results (one finding per slide) · 10 discussion · 11 conclusion (restate finding) · 12 backup [A1][A3][A4].

## Product / sales / pitch decks

Two different decks share the name:

- **Investor pitch — "why invest".** Sequoia's 10-block arc: purpose · problem · solution · why now · market · competition · product · business model · traction · team (+ ask) [P1][P2]. Kawasaki's 10/20/30 (10 slides, 20 minutes, ≥ 30 pt) is the time-boxed variant [P3]; a seed pitch ≈ 9 slides / ~150 seconds: what we do, how, customers, unique insight, bottom-up market size, progress, growth, team, three powerful facts [P5][P6].
- **Customer/sales narrative — "why change".** Lead with the world changing and the stake (winners & losers), paint the "promised land", then present capabilities as the gifts that get there + evidence [P4]. Never open with your company, logos, or the product.

Notes: competition slide is mandatory ("no competition" is a red flag) [P1]; show the product/output, not the tech [P1]; ≤ 15 slides for investors [P1]; slides are scenography for the spoken pitch — big text is a feature (≥ 30 pt forces one idea) [P3][P5].
**Pitfalls:** starting with the solution; hockey-stick projections with no basis; absurd top-down TAM; feature lists instead of purpose; reading the deck [P1][P3][P4].

**Outline (investor, Sequoia format):** 1 purpose (one line) · 2 problem (quantified) · 3 solution (core insight) · 4 why now · 5 market (TAM/SAM/SOM, sourced) · 6 competition (2×2) · 7 product · 8 business model · 9 traction · 10 team · 11 the ask [P1][P2].

## Internal decks: status & training

- **Status / business review.** Answer-first: status call (on track vs needs attention) on slide 1, then KPIs, plan-vs-actual, risks with responses, decisions needed; color-code status so it reads at a glance; keep it short — it supports a live conversation [C4][I1][I2]. Pitfalls: burying the status, data with no "so what", no owners/dates/asks.
- **Training / workshops.** Change capability, not just inform: define objective and audience first, then the outline, then content, then polish [I3]. Light text, favor demonstration; split session-length decks into modules with exercises.

## Comparison

| Genre | Goal | Structure | Density / visuals | Length | Tone |
|---|---|---|---|---|---|
| Consulting/strategy | Drive a decision | Pyramid: answer → MECE arguments → evidence | 1 message/slide, action titles, ~20% whitespace | Body 50+; summary = 1 slide | Assertive, evidence-led |
| Academic | Communicate one finding | Background→Purpose→Methods→Results→Discussion→Conclusion | ≤ 6–7 lines, ≥ 17–18 pt, big figures | ~1 slide/min | Honest, illustrative |
| Investor pitch | Get the meeting / raise | Sequoia 10 (purpose→…→team→ask) | ≤ 15 slides, ≥ 30 pt, show product | ~20 min | Story-first, quantified |
| Sales narrative | Move a buyer to change | Change → winners/losers → promised land → gifts → proof | Outcome visuals; no logos-first | Deal-stage | Urgent, loss-aversion |
| Internal status | Inform + get decisions | Status first → KPIs → variance → risks → actions | Color-coded status, 1 message/slide | Very short | Direct, factual |
| Training | Build capability | Objectives → modules → exercise → summary | Light text, one idea/slide | Session-length | Instructional |

## Japanese decks

- **Typography:** prefer fonts that exist on the target machine — `BIZ UDGothic` and `Noto Sans JP` render true in this environment; `Yu Gothic` / `Meiryo` get substituted during local QA (different metrics), so leave extra width slack on those [J2]. Sizes: title 36–44 pt, heading 28–32 pt, body 20–24 pt, in-figure ≥ 18 pt [A4]. Two-axis color rules: 3 colors, widen red/green differences (color-vision diversity) [A4][J2].
- **Registration (禁則):** enable kinsoku so `、。」` never orphan at line start and opening brackets never dangle at line end; avoid splitting compounds across lines [J3].
- **Consulting in Japanese:** 結論ファースト + 1スライド1メッセージ + 主張→根拠→示唆 ("So what?"); storyboard the title list (タイトル一覧) before opening the tool [C4].
- **Academic in Japanese:** keep the IMRAD-derived order 背景→目的→方法→結果→考察→結論, ~1 slide/min [A4].

## Markdown-first alternatives (when a deck is generated from text)

- **Slidev:** `slidev export --format pptx` gives image slides; `--format pptx-editable` rebuilds native shapes (SVG/canvas/video stay pictures; fonts not embedded) [M1].
- **Marp CLI:** `--pptx` is image-based; the experimental `--pptx-editable` needs LibreOffice, has lower fidelity and no presenter notes [M2].
- If the deliverable must be editable/native, generate with pptxgenjs (this skill's main path) or convert via LibreOffice.

## Sources

- [C1] Barbara Minto — Pyramid Principle concept — https://www.barbaraminto.com/concept
- [C2] StrategyCase — The Pyramid Principle: a consultant's guide — https://strategycase.com/the-pyramid-principle-case-interview/
- [C3] Slideworks — How McKinsey consultants make presentations — https://slideworks.io/resources/how-mckinsey-consultants-make-presentations
- [C4] スラサク — コンサル資料の作り方 — https://slasaku.jp/column/consul-shiryo-tsukurikata-guide
- [A1] Peyton Jones, Hughes & Launchbury — How to give a good research talk — https://simon.peytonjones.org/assets/pdfs/giving-a-talk.pdf
- [A3] Kirchoff — How to tell a compelling story in scientific presentations (Nature) — https://www.nature.com/articles/d41586-021-03603-2
- [A4] ドキュメントプラス — 分かりやすい学会発表のスライドの作り方 — https://www.documentplus.jp/2025/11/24/前編-分かりやすい学会発表のスライドの作り方/
- [A5] Nature — Boosting the signal in scientific talks — https://www.nature.com/articles/d41586-020-00838-3
- [P1] Winning Presentations — Investor pitch deck template (Sequoia format) — https://winningpresentations.com/investor-pitch-deck-template
- [P2] Sequoia Capital — Pitch deck template — https://research.cs.queensu.ca/softwarestartups/sequoiatemplate.pdf
- [P3] Guy Kawasaki — The 10/20/30 rule of PowerPoint — https://guykawasaki.com/the_102030_rule
- [P4] Andy Raskin — The making of a great strategic narrative — https://medium.com/the-mission/the-making-of-a-great-sales-narrative-978938b3926
- [P5] Y Combinator — Practical design: pitching — https://www.ycombinator.com/blog/practical-design-pitching
- [P6] Y Combinator — How to build your seed round pitch deck — https://www.ycombinator.com/library/2u-how-to-build-your-seed-round-pitch-deck
- [I1] SlidesPilot — Quarterly business review (QBR) deck — https://www.slidespilot.com/blog/quarterly-business-review-qbr-deck
- [I2] Rework — Exec narrative and board prep — https://resources.rework.com/guides/chief-of-staff-playbooks/exec-narrative-board-prep
- [I3] UMU — 研修資料の作り方 — https://www.umu.co/column/training-materials_20210929
- [J2] 北海道大学 小泉研究室 — 研究のプレゼン: フォントと色づかい — https://noah.ees.hokudai.ac.jp/envmi/koizumilab/labs-blog/
- [J3] PREZEN SQUARE — 禁則処理 — https://www.prezen-square.jp/posts/begginer_characters
- [M1] Slidev — Exporting (pptx / pptx-editable) — https://sli.dev/guide/exporting
- [M2] Marp CLI — Convert to PowerPoint — https://github.com/marp-team/marp-cli

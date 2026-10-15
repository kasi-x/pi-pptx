# Charts

Native PowerPoint charts only — everything goes through `slide.addChart()`, never a screenshot.
The general API (mutated option objects, colours without `#`, one `pptxgenjs` per file) is in
`pptxgenjs.md`; palettes, contrast thresholds and chart alt text are in `design-system.md`.

**Evidence.** Every recipe below ran on this machine on **2026-10-15 (pptxgenjs 4.0.1)**:
`writeFile()` → `soffice --convert-to pdf` → `pdftoppm` (`scripts/render.sh`), then the chart part
was read at `ppt/charts/chartN.xml`. In the tables `[v]` = probed that day, `[OOXML]` = rule from
the chart schema (MS-OE376 / ECMA-376, see Sources); PowerPoint is not installed here, so anything
only PowerPoint can reject is flagged.

## 1. Pick the chart from the intent, not from the data shape

| Presentation intent | Chart | pptxgenjs | Non-negotiable |
|---|---|---|---|
| Compare / rank categories | Horizontal bar, sorted by value | `ChartType.bar`, `barDir: "bar"` | Zero baseline; sort by the value you talk about |
| Compare a small grid of categories | Clustered column | `ChartType.bar`, `barDir: "col"` | ≤ 3–4 series, else small multiples |
| Trend over time, few series | Multi-series line | `ChartType.line` | Time on x; one y axis |
| Trend + composition together | Stacked area (100 % stacked) | `ChartType.area`, `barGrouping: "stacked"` | Only top and bottom bands are readable — label them |
| Composition at one moment | Pie / doughnut, **≤ 5 slices** | `ChartType.pie` / `.doughnut` | Never for change over time; collapse the tail into "Other" |
| Composition over periods | Stacked bar (one bar per period) | `ChartType.bar`, `barGrouping: "stacked"` | Parts must sum to a meaningful whole |
| Distribution | Columns of pre-binned values | `ChartType.bar` | Say what the bins are in the axis title |
| Correlation | Scatter | `ChartType.scatter` | X data goes **first** (§2.7); label each cloud |
| Correlation + a third magnitude | Bubble | `ChartType.bubble` | State what the area means (`sizes`) |
| Level + rate on one slide | Combo, secondary value axis | `addChart([{…},{…}], opts)` | Only when the series are causally linked (§4) |
| Process / bridge (open → steps → close) | Waterfall | stacked bar + invisible base (§2.10) | PowerPoint's native waterfall is not in pptxgenjs |

`[v]` The enum is exactly `area, bar, bar3d, bubble, bubble3d, doughnut, line, pie, radar,
scatter`, and `grep -i waterfall node_modules/pptxgenjs/dist/pptxgen.cjs.js` finds nothing — so
native waterfall / treemap / histogram / box-whisker / funnel parts cannot be generated; §2.10 is
the workaround.

## 2. Recipes

**2.0 Shared preamble.** Every recipe below is a slice of one tested file, so the helpers are
consistent throughout.

```js
const C = { ink: "1F2933", paper: "FFFFFF", navy: "16324F", ice: "CADCFC",
            teal: "2A9D8F", grey: "5C6B73", line: "E5E9EC" };
const RAMP = ["16324F", "2E5480", "4E7CA8", "7FA6C9", "B3CBDF"]; // ordered categories only
const BOX = { x: 0.7, y: 1.75, w: 11.93, h: 5.0 };   // headline sits above, 13.33x7.5 canvas
const BASE = { showTitle: false, showLegend: true, legendPos: "b",
               legendFontSize: 12, legendColor: C.grey };
const axes = (extra = {}) => Object.assign({
  catAxisLabelColor: C.grey, valAxisLabelColor: C.grey,
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  catAxisLineColor: C.line, valAxisLineColor: C.line,
  catAxisTitleColor: C.grey, valAxisTitleColor: C.grey,
  catAxisTitleFontSize: 12, valAxisTitleFontSize: 12,
  valGridLine: { color: C.line, size: 1 }, catGridLine: { style: "none" },
}, extra);
```

**2.1 Clustered column** — three series, legend below, titled value axis.
```js
s.addChart(p.ChartType.bar, [
  { name: "EMEA",     labels: ["Q1","Q2","Q3","Q4"], values: [42, 48, 61, 74] },
  { name: "Americas", labels: ["Q1","Q2","Q3","Q4"], values: [38, 41, 45, 52] },
  { name: "APAC",     labels: ["Q1","Q2","Q3","Q4"], values: [30, 33, 32, 31] },
], axes({ ...BOX, barDir: "col", barGapWidthPct: 60, chartColors: [C.navy, C.teal, C.ice],
  ...BASE, valAxisTitle: "Revenue (EUR m)", showValAxisTitle: true }));
```
verified 2026-10-15 (pptxgenjs 4.0.1): 3 objects → 3 `<c:ser>`, `<c:barDir val="col"/>`,
`<c:gapWidth val="60"/>`; rendered per quarter with legend and axis title. `barGrouping` defaults
to `clustered`.

**2.2 Stacked column.** Labels sit **on the fill**, so the fill follows the label colour: ink
labels need light fills (`6D96BC` 4.73:1, `9BBBD8` 7.37:1, `CADCFC` 10.65:1 — WCAG wants 4.5:1 at
11 pt), white labels need dark ones (`16324F` 13.1:1, `2E5480` 7.7:1, `3F6C99` 5.45:1; `4E7CA8` is
4.4:1 and fails).
```js
s.addChart(p.ChartType.bar, [
  { name: "Subscription", labels: ["Q1","Q2","Q3","Q4"], values: [30, 38, 52, 66] },
  { name: "Services",     labels: ["Q1","Q2","Q3","Q4"], values: [22, 24, 20, 21] },
  { name: "Licence",      labels: ["Q1","Q2","Q3","Q4"], values: [18, 16, 13, 11] },
], axes({ ...BOX, barDir: "col", barGrouping: "stacked", barGapWidthPct: 70,
  chartColors: ["6D96BC", "9BBBD8", "CADCFC"], ...BASE, showValue: true,
  dataLabelPosition: "ctr",            // stacked: ctr | inEnd | inBase ONLY
  dataLabelColor: C.ink, dataLabelFontSize: 11, dataLabelFontBold: true }));
```
verified 2026-10-15 (pptxgenjs 4.0.1): `<c:barGrouping val="stacked"/>` + `<c:overlap val="100"/>`
(forced for any grouping containing `stacked`), `<c:dLblPos val="ctr"/>` per series, segment values
(18/22/30 … 11/21/66) inside the bars. `outEnd` is illegal on a stack (§5.6).

**2.3 Multi-series line** — series order = legend order, one y axis.
```js
const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"];
s.addChart(p.ChartType.line, [
  { name: "p99", labels: months, values: [410,425,460,505,512,530,340,288] },
  { name: "p50", labels: months, values: [120,124,131,138,140,142,89,74] },
], axes({ ...BOX, chartColors: [C.navy, C.teal], lineSize: 3,
  lineDataSymbol: "circle", lineDataSymbolSize: 6, ...BASE,
  valAxisTitle: "Response time (ms)", showValAxisTitle: true,
  valAxisMinVal: 0, valAxisMaxVal: 560 }));
```
verified 2026-10-15 (pptxgenjs 4.0.1): 2 `<c:ser>`, `<c:symbol val="circle"/>` per series,
`<c:min val="0"/>`/`<c:max val="560"/>`; lines, markers, legend rendered. `lineDash` is
**chart-wide** (`"dash"` dashed every series: 2 `<a:prstDash val="dash"/>`, verified), so a dashed
forecast beside a solid actual needs the combo form of §2.9 with two `line` groups and no `valAxes`
(verified: one dashed, one solid, one axis pair).

**2.4 Stacked area** — first series at the bottom, no in-chart labels.
```js
const years = ["2021","2022","2023","2024","2025"];
s.addChart(p.ChartType.area, [
  { name: "On-prem", labels: years, values: [78, 66, 52, 39, 28] },
  { name: "Cloud",   labels: years, values: [22, 34, 48, 61, 72] },
], axes({ ...BOX, barGrouping: "stacked", chartColors: [C.ice, C.navy], ...BASE,
  valAxisTitle: "Share of workloads (%)", showValAxisTitle: true,
  valAxisMinVal: 0, valAxisMaxVal: 100 }));
```
verified 2026-10-15 (pptxgenjs 4.0.1): `<c:areaChart>` + `<c:grouping val="stacked"/>`; both bands
rendered, 2025 with the 72 % cloud band on top. Area charts take **no** `dataLabelPosition` (§5.8).

**2.5 Pie, five slices (the maximum)** — labels outside on white, so they stay ink; the tail
becomes "Other".
```js
s.addChart(p.ChartType.pie, [
  { name: "Source", labels: ["Direct","Partner","Marketplace","Events","Other"],
    values: [38, 24, 20, 10, 8] },
], { ...BOX, chartColors: RAMP, ...BASE, legendPos: "r",
  showValue: true, showPercent: false, dataLabelPosition: "outEnd",
  dataLabelFormatCode: '0"%"',          // literal % — not 0'%' and not 0%
  dataLabelColor: C.ink, dataLabelFontSize: 12,
  showLeaderLines: true, firstSliceAng: 0 });
```
verified 2026-10-15 (pptxgenjs 4.0.1): labels are written **per point** — 5 `<c:dLbl>` blocks with
their own `<c:numFmt formatCode="0&quot;%&quot;"/>` and `<c:dLblPos val="outEnd"/>`, plus
`<c:showLeaderLines val="1"/>`; 38 %/24 %/20 %/10 %/8 % rendered outside, slice pixels proportional
to value (38→92 122 px, 8→19 569 px; 2 424–2 446 px per unit).

**2.6 Doughnut** — `holeSize` is a percentage of the outer radius. pptxgenjs **drops**
`dataLabelPosition` on doughnuts (verified: no `<c:dLblPos>` in the part), so labels land on the
ring: light ramp with ink labels, or dark ramp with white labels, never a ramp that crosses both.
```js
s.addChart(p.ChartType.doughnut, [
  { name: "Tickets", labels: ["Sync","Auth","Export","Billing","Other"], values: [52,19,14,9,6] },
], { ...BOX, chartColors: ["6D96BC","7FA6C9","9BBBD8","B3CBDF","CADCFC"], holeSize: 55,
  ...BASE, legendPos: "r", showValue: true, showPercent: false, dataLabelFormatCode: '0"%"',
  dataLabelColor: C.ink, dataLabelFontSize: 12 });
```
verified 2026-10-15 (pptxgenjs 4.0.1): `<c:holeSize val="55"/>`, 5 `<c:dLbl>` blocks; labels on the
ring at 4.73:1–10.65:1 against their fills, slice areas proportional to values.

**2.7 Scatter — `data[0]` is the X axis.** The one chart whose data shape is not "one object per
series" (§5.1): shared numeric X values first, then one object per Y series. The X object is
neither drawn nor listed in the legend.
```js
s.addChart(p.ChartType.scatter, [
  { name: "X", values: [3, 4, 6, 7, 9, 11, 12, 15, 18, 24, 30] },        // shared x
  { name: "Enterprise", values: [42, 47, 51, 55, 62, 68, 66, 74, 96, 118, 131] },
  { name: "Mid-market", values: [12, 16, 19, 22, 25, 28, 27, 31, 37, 44, 49] },
], axes({ ...BOX, chartColors: [C.navy, C.teal],
  lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 9,   // 0 = points, no line
  ...BASE, catAxisTitle: "Deal size (EUR m)", showCatAxisTitle: true,
  valAxisTitle: "Cycle time (days)", showValAxisTitle: true }));
```
verified 2026-10-15 (pptxgenjs 4.0.1): 3 objects → **2** `<c:ser>` (X is dropped), `lineSize: 0` →
`<a:ln><a:noFill/></a:ln>`, two value axes and **no** category axis; two separate clouds rendered,
legend listing only `Enterprise`/`Mid-market`. Per-point text: `showLabel: true` plus a nested
`labels: [["boiler", …]]` on the series and `dataLabelFormatScatter: "custom"` → one `<c:dLbl>` per
point carrying that string (verified).

**2.8 Bubble** — one `sizes` array per series.
```js
s.addChart(p.ChartType.bubble, [
  { name: "X", values: [3, 7, 12, 18, 22] },                              // shared x
  { name: "Incidents", values: [62, 44, 22, 14, 8], sizes: [12, 30, 46, 24, 9] },
], axes({ ...BOX, chartColors: [C.navy], ...BASE, showLegend: false,
  catAxisTitle: "Mean time to resolve (hours)", showCatAxisTitle: true,
  valAxisTitle: "Customer impact score", showValAxisTitle: true }));
```
verified 2026-10-15 (pptxgenjs 4.0.1): 2 objects → 1 `<c:ser>` with `<c:xVal>`, `<c:yVal>`,
`<c:bubbleSize>`; bubble area tracked `sizes` proportionally (299–312 px² per unit, ±2 %).
**Without `sizes` no series is written at all** (§5.2).

**2.9 Combo with a secondary value axis.** The first argument is the array of chart groups;
`valAxes` / `catAxes` hold one entry per axis, and every series on the secondary axis must say so —
omit the flags and `writeFile()` throws (§5.3). Axis titles need the `show…` switch (§5.9).
```js
const years = ["2021","2022","2023","2024"];
s.addChart([
  { type: p.ChartType.bar,
    data: [{ name: "Revenue (EUR m)", labels: years, values: [38, 46, 55, 64] }],
    options: { chartColors: [C.navy], barDir: "col" } },
  { type: p.ChartType.line,
    data: [{ name: "Headcount", labels: years, values: [180, 186, 191, 188] }],
    options: { chartColors: [C.teal], lineSize: 3, lineDataSymbol: "circle",
               secondaryValAxis: true, secondaryCatAxis: true } },   // picks axes 2
], axes({ ...BOX, barGapWidthPct: 80, ...BASE,
  valAxes: [
    { valAxisTitle: "Revenue (EUR m)", showValAxisTitle: true, valAxisTitleColor: C.navy,
      valAxisLabelColor: C.navy, valAxisMinVal: 0, valAxisMaxVal: 80 },
    { valAxisTitle: "Headcount", showValAxisTitle: true, valAxisTitleColor: C.teal,
      valAxisLabelColor: C.teal, valAxisMinVal: 0, valAxisMaxVal: 220,
      valGridLine: { style: "none" } },
  ],
  catAxes: [{ catAxisTitle: "Fiscal year", showCatAxisTitle: true },
            { catAxisHidden: true }] }));     // 2nd category axis exists but is invisible
```
verified 2026-10-15 (pptxgenjs 4.0.1): one part with `<c:barChart>` + `<c:lineChart>`, **2
`<c:valAx>` + 2 `<c:catAx>`**, per-axis ranges, `<c:delete val="1"/>` on the second category axis;
columns rendered against the left axis, the line against the right, both titles, both legend
entries. Group order sets z-order — line group last. A combo with **no** `valAxes`/`catAxes` is
legal too: single-axis bar+line (verified: 1 + 1 axes), which is what a target line over columns
wants.

**2.10 Waterfall — stacked bar with an invisible base series.** PowerPoint has a native waterfall,
pptxgenjs does not (§1), so the bridge is two stacked series: an invisible one carrying the running
total, a visible one carrying each step.

- **The visible bar must always be positive.** A negative value inside a stack is drawn downward
  from *zero*, not from the running total (verified: a −57 step rendered from 0 to −57 below the
  axis instead of floating 142→199). For a decrease, set the base to the *lower* end and the bar to
  `Math.abs(delta)`; direction comes from where the bar sits.
- **The invisible series still draws its labels** (verified: 12 labels for 6 steps), so one XML
  edit switches them off after generation.

```js
const steps = [ { label: "FY24 exit", delta: 120, total: true },
                { label: "New logos", delta: 48 }, { label: "Expansion", delta: 31 },
                { label: "Discounts", delta: -57 }, { label: "Churn", delta: -73 },
                { label: "FY25 exit", delta: 69, total: true } ];
function bridge(steps) {                    // -> { base, bar } pairs for the two series
  let running = 0;
  return steps.map((st) => {
    if (st.total) { running = st.delta; return { ...st, base: 0, bar: st.delta }; }
    const from = running, to = running + st.delta; running = to;
    return { ...st, base: Math.min(from, to), bar: Math.abs(st.delta) };
  });
}
const bars = bridge(steps);
s.addChart(p.ChartType.bar, [
  { name: "_base",  labels: bars.map((d) => d.label), values: bars.map((d) => d.base) },
  { name: "change", labels: bars.map((d) => d.label), values: bars.map((d) => d.bar) },
], axes({ ...BOX, barDir: "col", barGrouping: "stacked", barGapWidthPct: 45,
  chartColors: ["transparent", C.navy],      // 1st series: <a:noFill/>
  ...BASE, showLegend: false, showValue: true, dataLabelPosition: "ctr",
  dataLabelColor: C.paper, dataLabelFontSize: 11, dataLabelFontBold: true,
  valAxisTitle: "ARR (EUR m)", showValAxisTitle: true,
  valAxisMinVal: 0, valAxisMaxVal: 240 }));
```
Then one patch: hide the base labels and give the visible series signed step labels. The two
assertions make a wrong series order fail loudly instead of labelling the wrong series.
```bash
scripts/unpack.sh deck.pptx unpacked/
python3 - <<'PY'
import pathlib
part = pathlib.Path("unpacked/ppt/charts/chart1.xml")
head, sep, rest = part.read_text(encoding="utf-8").partition("<c:ser>")
base, close1, mid = rest.partition("</c:ser>")
assert "<c:v>_base</c:v>" in base          # first series = the invisible base series
step, close2, tail = mid.removeprefix("<c:ser>").partition("</c:ser>")
assert "<c:v>change</c:v>" in step         # second series = the visible step bars
labels = ("+120", "+48", "+31", "-57", "-73", "+69")
dbl = "".join(
    '<c:dLbl><c:idx val="%d"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>'
    '<a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
    '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>'
    % (i, v) for i, v in enumerate(labels))
part.write_text(head + sep + base.replace("<c:dLbls>", '<c:dLbls><c:delete val="1"/>', 1)
                + close1 + "<c:ser>" + step.replace("<c:dLbls>", "<c:dLbls>" + dbl, 1)
                + close2 + tail, encoding="utf-8")
PY
scripts/pack.sh unpacked/ out.pptx
```
verified 2026-10-15 (pptxgenjs 4.0.1): series values `[0,120,168,142,69,0]` and
`[120,48,31,57,73,69]`; after the patch the render shows 6 floating bars proportional to |delta|
(0.0025 of slide height per unit, ±1 %), white labels `+120 +48 +31 −57 −73 +69` inside them, and
no base label. The bash/python above ran verbatim and produced byte-identical XML to the
argument-taking variant used while testing this recipe; per-point `<c:dLbl>` goes right after
`<c:dLbls>` — the schema order `[OOXML]` (`CT_DLbls` = `dLbl*` then group properties).
LibreOffice renders the patched part; PowerPoint was not available to confirm it accepts it.

## 3. Options that matter

### Axes

| Option | Effect in the XML | Notes |
|---|---|---|
| `catAxisTitle` / `valAxisTitle` **+ `show…AxisTitle: true`** | `<c:title>` inside that axis | `[v]` without the switch the title is dropped entirely |
| `valAxisMinVal` / `valAxisMaxVal` | `<c:min>` / `<c:max>` | Never above 0 on bar/column/area (§4) |
| `valAxisMajorUnit` | `<c:majorUnit val="…"/>` | Tick spacing |
| `valAxisLabelFormatCode` | `<c:numFmt>` on the **valAx** | `[v]` `"$#,##0"` → `$1,200`-style ticks |
| `valLabelFormatCode` | `<c:formatCode>` in the data table only | `[v]` it does **not** format axis labels — a common mix-up |
| `valAxisDisplayUnit` + `valAxisDisplayUnitLabel: true` | `<c:builtInUnit val="hundredThousands"/><c:dispUnitsLbl/>` | Keeps 620 000 as `6.2` + a unit label |
| `catAxisLabelRotate: -45` | `rot="-2700000"` in the axis `txPr` | Degrees × 60 000; for long category names |
| `catAxisLabelFrequency: "2"` | `<c:tickLblSkip val="2"/>` | Show every 2nd tick |
| `catAxisLineShow` / `valAxisLineShow: false` | `<a:ln><a:noFill/></a:ln>` | Axis line off, labels stay |
| `catAxisHidden` / `valAxisHidden: true` | `<c:delete val="1"/>` on that axis | How a combo hides its second category axis |
| `valGridLine: { color: "E5E9EC", size: 1 }` | `<c:majorGridlines>` with that line | Default `888888` is too loud |
| `catGridLine: { style: "none" }` | no `<c:majorGridlines>` element | Vertical gridlines usually add nothing |
| `catAxisOrientation: "minMax"` | `<c:orientation val="minMax"/>` | Time-like axes; the other value flips the order |
| `valAxisLogScaleBase` | `<c:logBase val="n"/>` | 2–99; not probed here beyond the type signature |

### Data labels

pptxgenjs strings: `ctr` Center, `inEnd` InsideEnd, `inBase` InsideBase, `outEnd` OutsideEnd,
`bestFit` BestFit, `l`/`r`/`t`/`b` left/right/top/bottom. Legal sets `[OOXML]` (MS-OE376
§5.7.2.48), with pptxgenjs's own behaviour appended.

| Chart | Legal positions | pptxgenjs behaviour |
|---|---|---|
| Clustered bar / column | inBase, inEnd, outEnd, ctr | **`outEnd` is deleted** before writing (`[v]`: no `<c:dLblPos>` in the part) — use `inEnd` |
| Stacked / 100 % stacked bar | inBase, inEnd, ctr | writes `outEnd` anyway, and PowerPoint then rejects the chart (§5.6) |
| Pie / bar-of-pie | bestFit, outEnd, inEnd, ctr | per-point labels; `outEnd` + `showLeaderLines: true` is the readable default |
| Line / scatter / bubble / stock | l, r, t, b, ctr | line/bubble also take `showLabel: true` for per-point custom text |
| Area, doughnut, radar, bar3d | **none** — "shall not be specified" | option deleted silently (`[v]`: doughnut part has no `<c:dLblPos>`) |

| Option | Effect | Notes |
|---|---|---|
| `showValue` | `<c:showVal val="1"/>` in every series' `dLbls` | Chart-wide; no per-series or per-point switch |
| `showPercent` / `showSerName` | `<c:showPercent>` / `<c:showSerName>` | Off unless set; pie/doughnut default the format code to `0%` only while `showPercent` is on |
| `dataLabelFormatCode: '0"%"'` | `<c:numFmt>` on the labels | Literal `%`; a bare `%` multiplies by 100 (§5.10). Default `#,##0` |
| `dataLabelColor` / `Size` / `FontFace` / `FontBold` | the label `txPr` | Default 12 pt black; one colour for the whole chart, so match the fill (§2.2) |
| `dataLabelBkgrdColors: true` | `<c:dLbls><c:spPr>` filled with the **series** colour | Bar path only (`[v]`: a pie wrote an empty `<c:spPr/>`); pointless on an equally coloured bar |
| `showLeaderLines: true` | `<c:showLeaderLines val="1"/>` | Pie/bar-of-pie, needed for outside labels |
| `altText` | `descr` on the frame's `<p:cNvPr>` | The accessibility hook — give the takeaway, not the chart type |

### Legend, title, colour, size

| Option | Effect | Notes |
|---|---|---|
| `showLegend` | writes `<c:legend>` | **Off by default** — set it or colour has no key |
| `legendPos` | `<c:legendPos val="b\|l\|r\|t\|tr"/>` | Default `r`; `b` keeps the plot wide |
| `legendFontSize` / `legendColor` / `legendFontFace` | legend `txPr` | 11–12 pt grey reads as a caption, not as content |
| `showTitle` + `title` (+ `FontSize/Bold/Color/Align`) | `<c:title>` rich text | **Off by default** (`[v]`: 14 pt bold → `sz="1400"`); on a slide prefer a text box as the action title (§4) |
| `chartColors` | one colour per series, or one per **category** | Default is the Office 2007 palette (`C0504D`, `4F81BD`, …; pies differ) — always set it |
| `chartColorsOpacity` | `<a:alpha val="850"/>` on the series fill | Bar/line/scatter/bubble **only** (`[v]`: ignored on a pie) |
| `chartArea` / `plotArea` | `{ fill, border }` → `<c:spPr>` | Chart area is written after `</c:chart>`, plot area after the axes |
| `firstSliceAng` / `holeSize` | `<c:firstSliceAng>` / `<c:holeSize>` | Rotate a pie so small slices neighbour the largest; hole is 1–100 |
| `barGapWidthPct` / `barOverlapPct` | `<c:gapWidth>` / `<c:overlap>` | 0–1000, default 150; 40–70 reads better. Stacked always writes 100 |
| `lineSize` / `lineDash` / `lineSmooth` / `lineDataSymbol*` | series `spPr`, `<c:marker>` | 2 pt default line; marker size 2–72 |
| `x, y, w, h` | frame `<a:off>` / `<a:ext>` | Inches or percent strings; **off-canvas charts are written but invisible** |

## 4. Making a chart say one thing

- **The title is the message.** State the takeaway ("Churn and discounting ate 60 % of new-logo
  growth"); the chart proves it. `showTitle: false` plus a real text box.
- **One message per chart.** Two questions = two charts; a chart that needs a paragraph is a table.
- **Bars start at zero.** `valAxisMinVal` above 0 exaggerates differences (`[v]`: `valAxisMinVal: 40`
  writes `<c:min val="40"/>` and renders a truncated axis). Only position-over-time charts (line,
  scatter) may drop the baseline, and then the axis must be labelled.
- **Ordered data one hue ramp; unordered categories distinct hues.** `RAMP` for quarters and bins,
  deck palette colours for product lines. Never a rainbow.
- **Sort bars by the value in the message**, not alphabetically, unless the categories have a
  natural order.
- **≤ 5 pie/doughnut slices**, tail collapsed into "Other"; if the reader must compare slices, use a
  bar chart — a pie is only for "one part dominates".
- **≤ 3 series per chart** (4 with a legend); more belongs in small multiples or a table.
- **Colour carries one meaning.** Accent for the series the message is about, greys/tints for the
  rest; never red-vs-green alone, never colour as the only signal (`design-system.md` §1, §6).
- **Legend or direct labels, not both**; legend order = series/stack order.
- **Unit, period and source**: axis titles plus a 10 pt source line.
- **`altText` on every chart** with the takeaway. **No 3D, no bar shadows, no gradients** — they
  distort magnitude and date the deck.

## 5. Failure modes

| # | Symptom | Cause | Fix |
|---|---|---|---|
| 1 | Scatter with 2+ series shows one cloud of nonsense; bubble writes **no series at all** | `data[0]` must hold the shared X values and each later object one Y series (`sizes` required for bubble); a `[[x,y],…]` pair-per-point shape is not supported | §2.7, §2.8 |
| 2 | `writeFile()` throws `Secondary axis must be used by one of the multiple charts` | `valAxes` on a single chart, or a combo whose series never set `secondaryValAxis` | Combo form + flags on the secondary series (§2.9) |
| 3 | `writeFile()` throws `There must be the same number of value and category axes.` | `valAxes` and `catAxes` have different lengths | Equal lengths; `catAxes` may be omitted for a single-axis combo |
| 4 | Combo renders locally but the part has 2 `<c:valAx>` and 1 `<c:catAx>` | `catAxes` omitted while two value axes were declared | Supply 2 `catAxes`; PowerPoint behaviour with the orphan axis is untested here |
| 5 | Chart missing from the slide although it was added | Frame off-canvas — written to the XML, never rendered (`[v]`: `off x="12801600"` > slide width `12192000` EMU) | Keep `x + w ≤ 13.33`, `y + h ≤ 7.5` |
| 6 | PowerPoint offers to repair; the part holds `<c:dLblPos val="outEnd"/>` in a stacked `barChart` | `outEnd` is illegal for stacked/100 % stacked bars `[OOXML]`; pptxgenjs writes it anyway | `ctr` / `inEnd` / `inBase` on stacked charts |
| 7 | `outEnd` labels silently vanish on a clustered column | pptxgenjs's position validation is inverted for bar charts (`[v]`: the option is deleted; `ctr`/`inBase`/`inEnd` survive) | Use `inEnd` — same visual default |
| 8 | `dataLabelPosition` has no effect on area / doughnut / radar / bar3d | The schema forbids the element there; pptxgenjs deletes the option (`[v]`) | Accept the automatic position or label outside the chart |
| 9 | Axis titles do not appear | `valAxisTitle`/`catAxisTitle` without the `show…AxisTitle` switch that gates the whole `<c:title>` (`[v]`) | Set both |
| 10 | Label reads `3800'%'` for the value 38 | A bare `%` multiplies by 100 and single quotes are not a literal delimiter (`[v]`: `0"%"` → `38%`; `0.0%` on 0.38 → `38.0%`) | Quote or escape the percent sign, or store fractions |
| 11 | Pie/doughnut transparency ignored | `chartColorsOpacity` is only read on the bar/line/scatter/bubble paths (`[v]`) | Pick colours at the right weight instead |
| 12 | Label background boxes invisible | `dataLabelBkgrdColors` fills the box with the **series** colour, bar path only | Use it for labels outside the bar, with a contrasting `dataLabelColor` |
| 13 | Axis ticks ignore `valLabelFormatCode` | That option is the data-table format code (`[v]` source + XML) | Use `valAxisLabelFormatCode` |
| 14 | One-series bar chart shows differently coloured bars with a one-entry legend | A multi-colour `chartColors` array on a single series writes per-category `<c:dPt>` colours (`[v]`: 4 `<c:dPt>`) while the legend still lists one series | One colour per series |
| 15 | Values in the file differ from "Edit Data in Excel" | The chart renders from `<c:numCache>`; the embedded workbook is a separate part (`[v]`: patching the cache to 300 rendered 300 while the workbook still held 120) | Regenerate the slide; do not hand-edit chart data |
| 16 | Chart is bare: no title, no legend, dated colours | `showTitle`/`showLegend` default to false and the palette defaults to Office 2007 (`[v]`) | Set them, or switch them off on purpose |
| 17 | Code crashes with `tmpData.forEach is not a function` | `addChart(type, data, opts)` needs an **array** of series objects; the library does no validation (`[v]`, validation is commented out) | Pass `[{ name, labels, values }]` |

## 6. Verify a chart before delivery

`scripts/render.sh out.pptx` (then look at every page), then read the parts back:
```bash
scripts/unpack.sh out.pptx unpacked/
python3 - <<'PY'
import re, pathlib
for f in sorted(pathlib.Path("unpacked/ppt/charts").glob("chart*.xml")):
    x = f.read_text()
    print(f.name, re.findall(r"<c:\w+Chart>", x), "ser", x.count("<c:ser>"),
          "valAx", x.count("<c:valAx>"), "catAx", x.count("<c:catAx>"),
          "dLblPos", set(re.findall(r'<c:dLblPos val="(\w+)"/>', x)),
          "numFmt", set(re.findall(r'formatCode="([^"]*)"', x)))
PY
```
Read it against §3: one `<c:*Chart>` element per `addChart` (plus one per combo group), the series
count you expect for the type (scatter/bubble drop the X object), axis counts matching your
`valAxes`/`catAxes`, and no `outEnd` in a stacked part.

## Sources

- [G1] PptxGenJS charts API — option names, combo signature, `holeSize`: https://gitbrent.github.io/PptxGenJS/docs/api-charts/
- [G2] PptxGenJS source and demos — `dist/pptxgen.cjs.js` (option validation, per-point `<c:dLbl>` for pies, scatter/bubble data shape, `dPt`), `demos/modules/demo_chart.mjs`, `types/index.d.ts`: https://github.com/gitbrent/PptxGenJS
- [G3] Issues [#1013](https://github.com/gitbrent/PptxGenJS/issues/1013) (secondary-axis combo), [#285](https://github.com/gitbrent/PptxGenJS/issues/285) (single-series `chartColors`), [#288](https://github.com/gitbrent/PptxGenJS/issues/288) (hiding axis values by format code)
- [O1] MS-OE376 §5.7.2.48 `dLblPos` — legal label positions per chart group: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7dd610ae-d805-40a4-ae3b-c171b6bf08e3
- [O2] ECMA-376 (Office Open XML) chart schema: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- [O3] Microsoft — available chart types in Office (waterfall is native): https://support.microsoft.com/en-us/office/available-chart-types-in-office-a6187218-807e-4103-9e0a-27cdb19afb90
- [O4] Microsoft — number format codes (`%` scales by 100, quoted literals): https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
- [V1] FT Visual Vocabulary — chart types by intent: https://github.com/Financial-Times/chart-doctor/tree/main/visual-vocabulary
- [V2] Claus Wilke — Fundamentals of Data Visualization, §6 (bars start at zero), §10 (proportions, pie limits): https://clauswilke.com/dataviz/visualizing-amounts.html
- [V3] Datawrapper — why not to use two axes: https://blog.datawrapper.de/dualaxis/
- [V4] data.europa academy — guidelines for pie charts: https://data.europa.eu/apps/data-visualisation-guide/guidelines-for-pie-charts

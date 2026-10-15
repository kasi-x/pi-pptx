/**
 * Render a structured outline (JSON) into a deck. Deterministic, no LLM in the loop:
 * the agent designs the outline, this script owns the pixels.
 *
 * Usage:
 *   node outline_to_pptx.js outline.json [out.pptx]
 *
 * Outline schema (all keys optional unless noted):
 * {
 *   "meta":  { "title": "...", "subtitle": "...", "author": "...", "date": "..." },
 *   "palette": "ocean" | "forest" | "clay",           // default: ocean
 *   "slides": [
 *     { "type": "title" },                                       // uses meta
 *     { "type": "section", "title": "Plan" },
 *     { "type": "bullets", "title": "...", "bullets": ["...", "..."], "notes": "..." },
 *     { "type": "stats",   "title": "...", "stats": [{ "value": "+18%", "label": "growth" }] },
 *     { "type": "chart",   "title": "...", "takeaway": "...",
 *       "chart": { "kind": "line|bar|pie|doughnut|area",
 *                  "series": [{ "name": "s", "labels": ["Q1"], "values": [1] }] } },
 *     { "type": "table",   "title": "...", "table": { "headers": ["A"], "rows": [["1"]] } },
 *     { "type": "closing", "title": "Thank you", "subtitle": "..." }
 *   ]
 * }
 */
const fs = require("node:fs");
const path = require("node:path");
const pptxgen = require("pptxgenjs");

const IN = 13.333; // LAYOUT_WIDE width
const M = 0.7; // outer margin
const TITLE = { y: 0.5, h: 0.9 };
const BODY_Y = 1.7;
const BODY_H = 5.2;
const FONT = { head: "Cambria", body: "Calibri" };

const PALETTES = {
	ocean: { paper: "FFFFFF", ink: "1F2933", brand: "16324F", accent: "2A9D8F", soft: "E9F2F0", grey: "5C6B73", line: "E5E9EC", onBrand: "FFFFFF", onBrandSoft: "CADCFC" },
	forest: { paper: "FFFFFF", ink: "1F2933", brand: "2C5F2D", accent: "97BC62", soft: "F0F4E8", grey: "5C6B73", line: "E4E9DA", onBrand: "FFFFFF", onBrandSoft: "DCE8C8" },
	clay: { paper: "FFFFFF", ink: "2B2523", brand: "B85042", accent: "7A9E9F", soft: "F6EFE9", grey: "6B5E57", line: "EADFD6", onBrand: "FFFFFF", onBrandSoft: "F2D9CF" },
};

function fail(msg) {
	console.error(`outline: ${msg}`);
	process.exit(1);
}

function addBase(pres, C, slide) {
	slide.background = { color: C.paper };
	slide.addShape(pres.ShapeType.rect, { x: M, y: 0.62, w: 0.18, h: 0.18, fill: { color: C.accent }, line: { type: "none" } });
}

function addTitle(pres, C, slide, text) {
	addBase(pres, C, slide);
	slide.addText(text, { x: M + 0.35, y: TITLE.y, w: IN - 2 * M - 0.35, h: TITLE.h, fontFace: FONT.head, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: "middle" });
}

function darkSlide(pres, C, s) {
	s.background = { color: C.brand };
	s.addShape(pres.ShapeType.rect, { x: M, y: M, w: 0.18, h: 0.18, fill: { color: C.accent }, line: { type: "none" } });
}

function renderTitle(pres, C, meta) {
	const s = pres.addSlide();
	darkSlide(pres, C, s);
	s.addText(meta.title ?? "Untitled deck", { x: M, y: 2.5, w: IN - 2 * M, h: 0.95, fontFace: FONT.head, fontSize: 42, bold: true, color: C.onBrand, margin: 0 });
	if (meta.subtitle) s.addText(meta.subtitle, { x: M, y: 3.75, w: IN - 2 * M, h: 0.7, fontFace: FONT.body, fontSize: 20, color: C.onBrandSoft, margin: 0 });
	const metaLine = [meta.author, meta.date].filter(Boolean).join("  ·  ");
	if (metaLine) s.addText(metaLine, { x: M, y: 6.4, w: IN - 2 * M, h: 0.4, fontSize: 12, color: C.onBrandSoft, margin: 0 });
	return s;
}

function renderSection(pres, C, def) {
	const s = pres.addSlide();
	darkSlide(pres, C, s);
	s.addText(def.title ?? "Section", { x: M, y: 3.1, w: IN - 2 * M, h: 1.0, fontFace: FONT.head, fontSize: 32, bold: true, color: C.onBrand, margin: 0 });
	return s;
}

function renderBullets(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const bullets = Array.isArray(def.bullets) ? def.bullets.slice(0, 7) : [];
	if (!bullets.length) fail(`bullets slide "${def.title}" has no bullets`);
	s.addText(
		bullets.map((b, i) => ({ text: String(b), options: { bullet: true, breakLine: i < bullets.length - 1, paraSpaceAfter: 10 } })),
		{ x: M, y: BODY_Y, w: IN - 2 * M, h: BODY_H, fontFace: FONT.body, fontSize: 16, color: C.ink, margin: 0, valign: "top" },
	);
	return s;
}

function renderStats(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const stats = Array.isArray(def.stats) ? def.stats.slice(0, 4) : [];
	if (!stats.length) fail(`stats slide "${def.title}" has no stats`);
	const colW = (IN - 2 * M) / stats.length;
	stats.forEach((stat, i) => {
		const x = M + i * colW;
		s.addText(String(stat.value ?? ""), { x, y: 2.3, w: colW - 0.4, h: 1.2, fontFace: FONT.head, fontSize: 50, bold: true, color: C.accent, margin: 0 });
		s.addText(String(stat.label ?? ""), { x, y: 3.5, w: colW - 0.4, h: 0.8, fontSize: 15, color: C.grey, margin: 0 });
	});
	return s;
}

const CHART_TYPES = { bar: "bar", line: "line", pie: "pie", doughnut: "doughnut", area: "area" };

function renderChart(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const chart = def.chart ?? {};
	const kind = CHART_TYPES[chart.kind];
	if (!kind) fail(`chart slide "${def.title}": unknown kind "${chart.kind}" (bar|line|pie|doughnut|area)`);
	const series = (chart.series ?? []).map((ser) => ({ name: String(ser.name ?? "series"), labels: (ser.labels ?? []).map(String), values: (ser.values ?? []).map(Number) }));
	if (!series.length) fail(`chart slide "${def.title}" has no series`);
	const colors = [C.accent, C.brand, C.grey, "9AA5AD"];
	s.addChart(
		kind,
		series,
		{
			x: M, y: BODY_Y, w: def.takeaway ? 8.0 : IN - 2 * M, h: BODY_H,
			showTitle: false, showLegend: series.length > 1,
			showValue: kind === "pie" || kind === "doughnut" ? false : series.length === 1 && chart.kind !== "line",
			chartColors: series.length === 1 ? [C.accent] : colors,
			catAxisLabelColor: C.grey, valAxisLabelColor: C.grey,
			catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
			valGridLine: { color: C.line, size: 1 }, catGridLine: { style: "none" },
			...(kind === "line" ? { lineDataSymbol: "circle" } : {}),
		},
	);
	if (def.takeaway) {
		s.addText(String(def.takeaway), { x: 8.9, y: BODY_Y, w: IN - M - 8.9, h: 3.0, fontSize: 18, color: C.ink, margin: 0, lineSpacingMultiple: 1.2 });
	}
	return s;
}

function renderTable(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const table = def.table ?? {};
	const headers = (table.headers ?? []).map((h) => ({ text: String(h), options: { bold: true, fill: { color: C.brand }, color: C.onBrand } }));
	const rows = (table.rows ?? []).map((row) => row.map((cell) => ({ text: String(cell), options: { color: C.ink } })));
	if (!headers.length || !rows.length) fail(`table slide "${def.title}" needs headers and rows`);
	const n = headers.length;
	const colW = Array(n).fill((IN - 2 * M) / n);
	s.addTable([headers, ...rows], { x: M, y: BODY_Y, w: IN - 2 * M, colW, rowH: 0.5, fontSize: 14, fontFace: FONT.body, border: { type: "solid", color: C.line, pt: 0.5 }, valign: "middle" });
	return s;
}

function renderClosing(pres, C, def) {
	const s = pres.addSlide();
	darkSlide(pres, C, s);
	s.addText(def.title ?? "Thank you", { x: M, y: 3.0, w: IN - 2 * M, h: 0.95, fontFace: FONT.head, fontSize: 34, bold: true, color: C.onBrand, margin: 0 });
	if (def.subtitle) s.addText(String(def.subtitle), { x: M, y: 4.25, w: IN - 2 * M, h: 0.7, fontSize: 18, color: C.onBrandSoft, margin: 0 });
	return s;
}

function main() {
	const [outlinePath, outArg] = process.argv.slice(2);
	if (!outlinePath) fail("usage: node outline_to_pptx.js outline.json [out.pptx]");
	let outline;
	try {
		outline = JSON.parse(fs.readFileSync(outlinePath, "utf8"));
	} catch (err) {
		fail(`cannot read outline ${outlinePath}: ${err.message}`);
	}
	if (!Array.isArray(outline.slides) || !outline.slides.length) fail("outline has no slides[]");

	const C = PALETTES[outline.palette ?? "ocean"] ?? fail(`unknown palette "${outline.palette}" (ocean|forest|clay)`);
	const meta = outline.meta ?? {};
	const pres = new pptxgen();
	pres.layout = "LAYOUT_WIDE";
	pres.author = meta.author ?? "pi-pptx";
	pres.title = meta.title ?? "Deck";

	const renderers = { title: (d) => renderTitle(pres, C, meta), section: (d) => renderSection(pres, C, d), bullets: (d) => renderBullets(pres, C, d), stats: (d) => renderStats(pres, C, d), chart: (d) => renderChart(pres, C, d), table: (d) => renderTable(pres, C, d), closing: (d) => renderClosing(pres, C, d) };

	outline.slides.forEach((def, i) => {
		const render = renderers[def?.type];
		if (!render) fail(`slide ${i + 1}: unknown type "${def?.type}" (${Object.keys(renderers).join("|")})`);
		const slide = render(def);
		if (def.notes) slide.addNotes(String(def.notes));
	});

	const out = outArg ?? path.basename(outlinePath).replace(/\.json$/i, "") + ".pptx";
	pres.writeFile({ fileName: out }).then(() => {
		console.log(`wrote: ${out} (${outline.slides.length} slides)`);
		if (outline.slides.some((d) => d?.type === "chart" && !d.notes)) console.log("hint: add \u201cnotes\u201d to chart slides \u2014 the takeaway belongs in the talk track too");
	});
}

main();

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
 *   "palette": "ocean" | "forest" | "clay" | "consulting",     // default: ocean
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
 *
 * The consulting types (exec_summary, kpi_row, axis_table, premise_conclusion,
 * issue_action, chevron_flow, status_heatmap, two_by_two, bridge, recommendation)
 * and the shared kicker/source/page chrome are documented in
 * references/consulting-layouts.md.
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
	consulting: { paper: "FFFFFF", ink: "1F3864", brand: "002060", accent: "C00000", soft: "EEF2F8", grey: "595959", line: "D9D9D9", onBrand: "FFFFFF", onBrandSoft: "D6E0F0", consult: true },
};

function fail(msg) {
	console.error(`outline: ${msg}`);
	process.exit(1);
}

function addBase(pres, C, slide) {
	slide.background = { color: C.paper };
	if (!C.consult) slide.addShape(pres.ShapeType.rect, { x: M, y: 0.62, w: 0.18, h: 0.18, fill: { color: C.accent }, line: { type: "none" } });
}

function wideUnits(ch) {
	return /[\u3000-\u9fff\uff00-\uffef]/.test(ch) ? 1 : 0.5;
}

// Action titles stay on one or two lines; the unit budget is ~26 em at 23 pt.
function titleLines(text) {
	const segments = String(text ?? "").split("\n");
	const lines = [];
	for (const segment of segments) {
		let current = "";
		let used = 0;
		for (const ch of segment) {
			const units = wideUnits(ch);
			if (used + units > 26 && current) {
				lines.push(current);
				current = "";
				used = 0;
			}
			current += ch;
			used += units;
		}
		if (current) lines.push(current);
	}
	if (lines.length > 2) return [lines[0], lines[1].slice(0, -2) + "\u2026"];
	return lines.length ? lines : [""];
}

function addTitle(pres, C, slide, text) {
	addBase(pres, C, slide);
	if (C.consult) {
		slide.addText(titleLines(text).join("\n"), { x: M, y: TITLE.y, w: 8.9, h: TITLE.h, fontFace: FONT.body, fontSize: 23, bold: true, color: C.ink, margin: 0, valign: "top", lineSpacingMultiple: 1 });
		return;
	}
	slide.addText(text, { x: M + 0.35, y: TITLE.y, w: IN - 2 * M - 0.35, h: TITLE.h, fontFace: FONT.head, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: "middle" });
}

// Shared consulting furniture: kicker top right, source bottom left, page number bottom right.
function addConsultingChrome(C, slide, def, meta, page, total, dark) {
	const ink = dark ? C.onBrandSoft : C.grey;
	const kicker = def.kicker ?? (def.type === "title" ? meta.date : meta.kicker);
	if (kicker) slide.addText(String(kicker), { x: 9.9, y: TITLE.y, w: IN - M - 9.9, h: 0.32, fontFace: FONT.body, fontSize: 11, color: ink, align: "right", margin: 0, valign: "top", charSpacing: 1 });
	const source = def.source ?? meta.source ?? (def.type === "title" ? [meta.author, meta.date].filter(Boolean).join("  \u00b7  ") : "");
	if (source) slide.addText(String(source), { x: M, y: 6.7, w: 9.6, h: 0.3, fontFace: FONT.body, fontSize: 10, color: ink, margin: 0, valign: "top" });
	slide.addText(`${String(page).padStart(2, "0")} / ${String(total).padStart(2, "0")}`, { x: 11.6, y: 6.7, w: IN - M - 11.6, h: 0.3, fontFace: FONT.body, fontSize: 10, color: ink, align: "right", margin: 0, valign: "top" });
}

function darkSlide(pres, C, s) {
	s.background = { color: C.brand };
	if (!C.consult) s.addShape(pres.ShapeType.rect, { x: M, y: M, w: 0.18, h: 0.18, fill: { color: C.accent }, line: { type: "none" } });
}

function renderTitle(pres, C, meta) {
	const s = pres.addSlide();
	darkSlide(pres, C, s);
	s.addText(meta.title ?? "Untitled deck", { x: M, y: 2.5, w: IN - 2 * M, h: 0.95, fontFace: FONT.head, fontSize: 42, bold: true, color: C.onBrand, margin: 0 });
	const metaLine = [meta.author, meta.date].filter(Boolean).join("  \u00b7  ");
	if (metaLine && !C.consult) s.addText(metaLine, { x: M, y: 6.4, w: IN - 2 * M, h: 0.4, fontSize: 12, color: C.onBrandSoft, margin: 0 });
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
	if (def.chart?.kind === "waterfall") return renderBridge(pres, C, def);
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
const NO_BORDER = () => ({ type: "none", pt: 0, color: "FFFFFF" });
const EDGE = (color, pt = 0.75) => ({ type: "solid", color, pt });

function rule(pres, slide, C, x, y, w) {
	slide.addShape(pres.ShapeType.rect, { x, y, w, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
}

function asText(value) {
	if (value && typeof value === "object") return String(value.text ?? value.label ?? value.title ?? "");
	return String(value ?? "");
}

function renderExecSummary(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const points = (Array.isArray(def.points) ? def.points : []).slice(0, 4);
	if (!points.length) fail(`exec_summary slide "${def.title}" has no points`);
	const band = [def.soWhat, def.nextStep].filter(Boolean).join("\n");
	const bottom = band ? 5.35 : 6.2;
	const rowH = (bottom - BODY_Y) / points.length;
	points.forEach((point, i) => {
		const y = BODY_Y + i * rowH;
		const compact = points.length > 3;
		s.addText([
			{ text: `${String(i + 1).padStart(2, "0")} `, options: { fontSize: compact ? 18 : 20, bold: true, color: C.accent } },
			{ text: asText(point.headline ?? point.title), options: { fontSize: 16, bold: true, color: C.ink, breakLine: true } },
			{ text: asText(point.detail ?? point.text), options: { fontSize: 14, color: C.grey } },
		], { x: M, y, w: IN - 2 * M, h: compact ? 0.62 : 0.9, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: compact ? 2 : 6 });
		if (i < points.length - 1) rule(pres, s, C, M, y + rowH - 0.05, IN - 2 * M);
	});
	if (band) s.addText(band, { x: M, y: 5.62, w: IN - 2 * M, h: 0.66, shape: pres.ShapeType.rect, fill: { color: C.soft }, line: { type: "none" }, fontFace: FONT.body, fontSize: 14, bold: true, color: C.ink, margin: [4, 10, 4, 10], valign: "middle" });
	return s;
}

function renderKpiRow(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const kpis = (Array.isArray(def.kpis) ? def.kpis : []).slice(0, 4);
	if (!kpis.length) fail(`kpi_row slide "${def.title}" has no kpis`);
	const gap = 0.4;
	const w = (IN - 2 * M - gap * (kpis.length - 1)) / kpis.length;
	kpis.forEach((kpi, i) => {
		const x = M + i * (w + gap);
		const value = asText(kpi.value);
		const runs = [
			{ text: value, options: { fontSize: Math.max(28, 40 - Math.max(0, value.length - 4) * 4), bold: true, color: C.accent, breakLine: true } },
			{ text: asText(kpi.label), options: { fontSize: 14, color: C.grey, breakLine: true } },
			{ text: asText(kpi.readout), options: { fontSize: 14, color: C.ink, breakLine: !!kpi.source } },
		];
		if (kpi.source) runs.push({ text: String(kpi.source), options: { fontSize: 10, color: C.grey } });
		s.addText(runs, { x, y: 2.1, w: w - 0.1, h: 2.6, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 10 });
		if (i < kpis.length - 1) s.addShape(pres.ShapeType.rect, { x: x + w + gap / 2 - 0.006, y: 2.1, w: 0.012, h: 3.0, fill: { color: C.line }, line: { type: "none" } });
	});
	return s;
}

function renderAxisTable(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const table = def.table ?? {};
	const headers = (table.headers ?? []).map(asText);
	if (!headers.length) fail(`axis_table slide "${def.title}" has no headers`);
	const rows = (table.rows ?? []).slice(0, 7);
	const dataCount = rows.filter((row) => Array.isArray(row)).length;
	if (!dataCount) fail(`axis_table slide "${def.title}" has no rows`);
	const head = headers.map((h, i) => ({ text: h, options: { bold: true, fontSize: 14, color: C.ink, align: i === 0 ? "left" : "center", valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), EDGE(C.ink, 1.25), NO_BORDER()] } }));
	const tableRows = [head];
	let seen = 0;
	rows.forEach((row) => {
		if (row && !Array.isArray(row) && row.group) {
			tableRows.push([{ text: asText(row.group), options: { bold: true, fontSize: 14, color: C.ink, fill: { color: C.soft }, valign: "middle", margin: [4, 8, 4, 8], colspan: headers.length, border: [NO_BORDER(), NO_BORDER(), NO_BORDER(), NO_BORDER()] } }]);
			return;
		}
		seen += 1;
		const last = seen === dataCount;
		tableRows.push(headers.map((_, i) => ({ text: asText(row[i]), options: { fontSize: 14, bold: i === 0, color: C.ink, align: i === 0 ? "left" : "center", valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), last ? NO_BORDER() : EDGE(C.line), NO_BORDER()] } })));
	});
	const weights = table.widths ?? headers.map((_, i) => (i === 0 ? 2 : 1));
	const sum = weights.reduce((a, b) => a + Number(b), 0);
	const colW = weights.map((v) => (IN - 2 * M) * Number(v) / sum);
	const rowH = tableRows.map((row) => (row[0].options?.colspan ? 0.42 : Math.min(0.68, 4.4 / tableRows.length)));
	s.addTable(tableRows, { x: M, y: BODY_Y, w: IN - 2 * M, colW, rowH, fontFace: FONT.body, fontSize: 14, valign: "middle", autoPage: false });
	return s;
}

function renderPremiseConclusion(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const left = def.premise ?? def.left ?? {};
	const right = def.conclusion ?? def.right ?? {};
	const triX = M + 6.6 + 0.35;
	const rightX = triX + 0.4 + 0.35;
	const rightW = IN - M - rightX;
	s.addText(asText(left.header ?? "事実"), { x: M, y: BODY_Y, w: 6.6, h: 0.3, fontFace: FONT.body, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "top" });
	rule(pres, s, C, M, 2.12, 6.6);
	(left.rows ?? []).slice(0, 4).forEach((fact, i) => {
		const pair = Array.isArray(fact) ? fact : [fact.axis, fact.text];
		s.addText([
			{ text: asText(pair[0]), options: { fontSize: 14, bold: true, color: C.grey, breakLine: true } },
			{ text: asText(pair[1]), options: { fontSize: 14, color: C.ink } },
		], { x: M, y: 2.35 + i * 0.95, w: 6.6, h: 0.58, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 4 });
	});
	s.addShape(pres.ShapeType.triangle, { x: triX, y: 3.7, w: 0.4, h: 0.35, rotate: 90, fill: { color: C.accent }, line: { type: "none" } });
	s.addText(asText(right.header ?? "だから、"), { x: rightX, y: BODY_Y, w: rightW, h: 0.3, fontFace: FONT.body, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "top" });
	rule(pres, s, C, rightX, 2.12, rightW);
	const bullets = (right.bullets ?? []).slice(0, 4).map((item) => ({ text: asText(item), options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } }));
	if (bullets.length) s.addText(bullets, { x: rightX, y: 2.35, w: rightW, h: 3.6, fontFace: FONT.body, fontSize: 14, color: C.ink, margin: 0, valign: "top" });
	return s;
}

function renderIssueAction(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const rows = (def.rows ?? []).slice(0, 4);
	if (!rows.length) fail(`issue_action slide "${def.title}" has no rows`);
	const heads = def.headers ?? {};
	const w = (IN - 2 * M - 0.4) / 2;
	const x2 = M + w + 0.4;
	[[M, asText(heads.issue ?? "課題")], [x2, asText(heads.action ?? "打ち手")]].forEach(([x, head]) => {
		s.addText(head, { x, y: BODY_Y, w, h: 0.3, fontFace: FONT.body, fontSize: 15, bold: true, color: C.ink, margin: 0, valign: "top" });
		rule(pres, s, C, x, 2.12, w);
	});
	rows.forEach((row, i) => {
		const y = 2.35 + i * 1.05;
		const issue = row.issue ?? {};
		const action = row.action ?? {};
		s.addText([
			{ text: asText(issue.label ?? issue), options: { fontSize: 14, bold: true, color: C.ink, breakLine: true } },
			{ text: asText(issue.text ?? ""), options: { fontSize: 14, color: C.grey } },
		], { x: M, y, w, h: 0.72, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 4 });
		const meta = [action.owner, action.date].filter(Boolean).map(asText).join("  \u00b7  ");
		const runs = [
			{ text: asText(action.label ?? action), options: { fontSize: 14, bold: true, color: C.ink, breakLine: true } },
			{ text: asText(action.text ?? ""), options: { fontSize: 14, color: C.grey, breakLine: !!meta } },
		];
		if (meta) runs.push({ text: meta, options: { fontSize: 12, color: C.grey } });
		s.addText(runs, { x: x2, y, w, h: 0.72, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 4 });
		if (i < rows.length - 1) rule(pres, s, C, M, y + 0.8, IN - 2 * M);
	});
	return s;
}

function renderChevronFlow(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const steps = (def.steps ?? []).slice(0, 5);
	if (steps.length < 3) fail(`chevron_flow slide "${def.title}" needs 3-5 steps`);
	const gap = 0.32;
	const w = (IN - 2 * M - gap * (steps.length - 1)) / steps.length;
	steps.forEach((step, i) => {
		const runs = [
			{ text: asText(step.label ?? String(i + 1).padStart(2, "0")), options: { fontSize: 12, bold: true, color: C.onBrandSoft, breakLine: true } },
			{ text: asText(step.title), options: { fontSize: 15, bold: true, color: C.onBrand, breakLine: true } },
			{ text: asText(step.text), options: { fontSize: 14, color: C.onBrand } },
		];
		s.addText(runs, {
			x: M + i * (w + gap), y: 2.55, w, h: 1.85,
			shape: i === 0 ? pres.ShapeType.homePlate : pres.ShapeType.chevron,
			fill: { color: C.brand }, line: { type: "none" }, rectRadius: 0.18,
			fontFace: FONT.body, margin: [8, 10, 8, i === 0 ? 14 : 26], valign: "middle", paraSpaceAfter: 4,
		});
	});
	return s;
}

const STATUS = {
	ok: { mark: "\u25cf", label: "\u6b63\u5e38", color: "1B7F4B" },
	warn: { mark: "\u25b2", label: "\u6ce8\u610f", color: "8A5A00" },
	risk: { mark: "\u203b", label: "\u30ea\u30b9\u30af", color: "B3261E" },
};

function statusOf(value, where) {
	const key = String(value ?? "").toLowerCase();
	if (["ok", "good", "green", "\u6b63\u5e38"].includes(key)) return STATUS.ok;
	if (["warn", "warning", "amber", "yellow", "\u6ce8\u610f"].includes(key)) return STATUS.warn;
	if (["risk", "red", "bad", "\u30ea\u30b9\u30af"].includes(key)) return STATUS.risk;
	return fail(`${where}: unknown status "${value}" (ok|warn|risk)`);
}

function renderStatusHeatmap(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const hm = def.heatmap ?? {};
	const comments = Array.isArray(hm.comments) ? hm.comments : null;
	const columns = (hm.columns ?? ["\u9805\u76ee", "\u72b6\u614b"]).map(asText);
	if (comments) columns.push(asText(hm.commentHeader ?? "\u30b3\u30e1\u30f3\u30c8"));
	const rows = (hm.rows ?? []).slice(0, 5);
	if (!rows.length) fail(`status_heatmap slide "${def.title}" has no rows`);
	const stateCount = columns.length - 1 - (comments ? 1 : 0);
	rows.forEach((row) => {
		if ((row.values ?? []).length !== stateCount) fail(`status_heatmap slide "${def.title}": each row needs ${stateCount} status values`);
	});
	const axisW = 2.8;
	const commentW = comments ? 3.4 : 0;
	const stateW = (IN - 2 * M - axisW - commentW) / (columns.length - 1 - (comments ? 1 : 0));
	const colW = columns.map((_, i) => (i === 0 ? axisW : comments && i === columns.length - 1 ? commentW : stateW));
	const head = columns.map((h, i) => ({ text: h, options: { bold: true, fontSize: 14, color: C.ink, align: i === 0 ? "left" : "center", valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), EDGE(C.ink, 1.25), NO_BORDER()] } }));
	const body = rows.map((row, ri) => {
		const bottom = ri === rows.length - 1 ? NO_BORDER() : EDGE(C.line);
		const cells = [{ text: asText(row.label), options: { fontSize: 14, bold: true, color: C.ink, valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), bottom, NO_BORDER()] } }];
		(row.values ?? []).forEach((value) => {
			const state = statusOf(value, `status_heatmap slide "${def.title}"`);
			cells.push({ text: `${state.mark} ${state.label}`, options: { fontSize: 14, bold: true, color: state.color, align: "center", valign: "middle", margin: [4, 6, 4, 6], border: [NO_BORDER(), NO_BORDER(), bottom, NO_BORDER()] } });
		});
		if (comments) cells.push({ text: asText(comments[ri]), options: { fontSize: 14, color: C.ink, valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), bottom, NO_BORDER()] } });
		return cells;
	});
	s.addTable([head, ...body], { x: M, y: BODY_Y, w: IN - 2 * M, colW, rowH: [0.45, ...rows.map(() => 0.65)], fontFace: FONT.body, fontSize: 14, valign: "middle", autoPage: false });
	const legendY = Math.min(6.05, BODY_Y + 0.45 + rows.length * 0.65 + 0.32);
	s.addText([
		{ text: "\u25cf \u6b63\u5e38", options: { color: STATUS.ok.color } },
		{ text: "\u3000\u25b2 \u6ce8\u610f", options: { color: STATUS.warn.color } },
		{ text: "\u3000\u203b \u30ea\u30b9\u30af", options: { color: STATUS.risk.color } },
	], { x: M, y: legendY, w: IN - 2 * M, h: 0.3, fontFace: FONT.body, fontSize: 12, margin: 0, valign: "top" });
	return s;
}

function renderTwoByTwo(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const m = def.matrix ?? {};
	const xAxis = m.xAxis ?? {};
	const yAxis = m.yAxis ?? {};
	const quads = (m.quadrants ?? []).slice(0, 4);
	if (quads.length !== 4) fail(`two_by_two slide "${def.title}" needs 4 quadrants`);
	s.addText(`\u6a2a\u8cc2\uff1a${asText(xAxis.name ?? "\u5f71\u97ff\u5ea6")}\u3000/\u3000\u7e26\u8cc2\uff1a${asText(yAxis.name ?? "\u5b9f\u884c\u5bb9\u6613\u6027")}`, { x: M, y: BODY_Y, w: IN - 2 * M, h: 0.3, fontFace: FONT.body, fontSize: 14, color: C.grey, margin: 0, valign: "top" });
	const axisW = 1.3;
	const qW = (IN - 2 * M - axisW - 0.6) / 2;
	const x1 = M + axisW + 0.3;
	const x2 = x1 + qW + 0.3;
	[[x1, xAxis.low ?? "\u4f4e"], [x2, xAxis.high ?? "\u9ad8"]].forEach(([x, label]) => {
		s.addText(asText(label), { x, y: 2.3, w: qW, h: 0.3, fontFace: FONT.body, fontSize: 14, bold: true, color: C.ink, margin: 0, valign: "top" });
	});
	[[M, yAxis.high ?? "\u9ad8", 2.9], [M, yAxis.low ?? "\u4f4e", 4.5]].forEach(([x, label, y]) => {
		s.addText(asText(label), { x, y, w: axisW, h: 1.3, fontFace: FONT.body, fontSize: 14, bold: true, color: C.grey, margin: 0, valign: "middle" });
	});
	quads.forEach((quad, i) => {
		const items = (quad.items ?? []).slice(0, 3);
		s.addText([
			{ text: asText(quad.label), options: { bold: true, fontSize: 15, color: C.ink, breakLine: true } },
			...items.map((item, j) => ({ text: asText(item), options: { fontSize: 14, color: C.ink, breakLine: j < items.length - 1 } })),
		], { x: i % 2 === 0 ? x1 : x2, y: i < 2 ? 2.9 : 4.5, w: qW, h: 1.3, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 4 });
	});
	rule(pres, s, C, M, 2.75, IN - 2 * M);
	rule(pres, s, C, M, 4.35, IN - 2 * M);
	s.addShape(pres.ShapeType.rect, { x: M + axisW + 0.15, y: 2.75, w: 0.012, h: 3.05, fill: { color: C.line }, line: { type: "none" } });
	s.addShape(pres.ShapeType.rect, { x: x1 + qW + 0.15, y: 2.75, w: 0.012, h: 3.05, fill: { color: C.line }, line: { type: "none" } });
	return s;
}

function renderBridge(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const chart = def.chart ?? def.bridge ?? {};
	const steps = (chart.steps ?? []).slice(0, 8);
	if (steps.length < 2) fail(`bridge slide "${def.title}" needs chart.steps`);
	let running = 0;
	const labels = [];
	const base = [];
	const bar = [];
	steps.forEach((step) => {
		const label = asText(step.label);
		const delta = Number(step.delta ?? 0);
		if (!Number.isFinite(delta)) fail(`bridge slide "${def.title}": non-numeric delta`);
		const format = (v) => Math.abs(v).toLocaleString("en-US");
		if (step.total) {
			running = delta;
			base.push(Math.min(0, delta));
			bar.push(Math.abs(delta));
			labels.push(`${label}\n${format(delta)}`);
		} else {
			const next = running + delta;
			base.push(Math.min(running, next));
			bar.push(Math.abs(delta));
			running = next;
			labels.push(`${label}\n${delta >= 0 ? "+" : "\u2212"}${format(delta)}`);
		}
	});
	const axisMin = Math.min(0, ...base);
	s.addChart(pres.ChartType.bar, [
		{ name: "_base", labels, values: base },
		{ name: "change", labels, values: bar },
	], {
		x: M, y: BODY_Y, w: IN - 2 * M, h: 4.6,
		barDir: "col", barGrouping: "stacked", barGapWidthPct: 45,
		chartColors: ["transparent", C.brand], showTitle: false, showLegend: false, showValue: false,
		catAxisLabelColor: C.grey, valAxisLabelColor: C.grey, catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
		catGridLine: { style: "none" }, valGridLine: { color: C.line, size: 1 },
		valAxisTitle: chart.unit ? String(chart.unit) : "", showValAxisTitle: !!chart.unit, valAxisMinVal: axisMin,
	});
	return s;
}

function renderRecommendation(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const reasons = (def.reasons ?? []).slice(0, 3);
	if (!reasons.length) fail(`recommendation slide "${def.title}" has no reasons`);
	s.addText(asText(def.answer ?? def.recommendation ?? ""), { x: M, y: BODY_Y, w: IN - 2 * M, h: 0.9, shape: pres.ShapeType.rect, fill: { color: C.soft }, line: { type: "none" }, fontFace: FONT.body, fontSize: 18, bold: true, color: C.ink, margin: [8, 12, 8, 12], valign: "middle" });
	const gap = 0.4;
	const w = (IN - 2 * M - gap * 2) / 3;
	reasons.forEach((reason, i) => {
		s.addText([
			{ text: String(i + 1).padStart(2, "0"), options: { fontSize: 16, bold: true, color: C.accent, breakLine: true } },
			{ text: asText(reason), options: { fontSize: 14, color: C.ink } },
		], { x: M + i * (w + gap), y: 2.9, w, h: 1.5, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 6 });
	});
	const steps = (def.nextSteps ?? []).slice(0, 5);
	const rows = steps.map((step) => [asText(step.action ?? step), asText(step.owner), asText(step.date)]);
	if (!rows.length) fail(`recommendation slide "${def.title}" has no nextSteps`);
	const head = ["Next step", "Owner", "Date"].map((h) => ({ text: h, options: { bold: true, fontSize: 14, color: C.ink, border: [NO_BORDER(), NO_BORDER(), EDGE(C.ink, 1.25), NO_BORDER()] } }));
	const body = rows.map((row, i) => row.map((cell) => ({ text: cell, options: { fontSize: 14, color: C.ink, valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), i === rows.length - 1 ? NO_BORDER() : EDGE(C.line), NO_BORDER()] } })));
	s.addTable([head, ...body], { x: M, y: 4.7, w: IN - 2 * M, colW: [6.5, 2.9, 2.53], rowH: [0.4, ...rows.map(() => 0.42)], fontFace: FONT.body, fontSize: 14, valign: "middle", autoPage: false });
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

	const C = PALETTES[outline.palette ?? "ocean"] ?? fail(`unknown palette "${outline.palette}" (ocean|forest|clay|consulting)`);
	const meta = outline.meta ?? {};
	const pres = new pptxgen();
	pres.layout = "LAYOUT_WIDE";
	pres.author = meta.author ?? "pi-pptx";
	pres.title = meta.title ?? "Deck";

	const renderers = {
		title: () => renderTitle(pres, C, meta), section: (d) => renderSection(pres, C, d), bullets: (d) => renderBullets(pres, C, d),
		stats: (d) => renderStats(pres, C, d), chart: (d) => renderChart(pres, C, d), table: (d) => renderTable(pres, C, d), closing: (d) => renderClosing(pres, C, d),
		exec_summary: (d) => renderExecSummary(pres, C, d), kpi_row: (d) => renderKpiRow(pres, C, d), axis_table: (d) => renderAxisTable(pres, C, d),
		premise_conclusion: (d) => renderPremiseConclusion(pres, C, d), issue_action: (d) => renderIssueAction(pres, C, d),
		chevron_flow: (d) => renderChevronFlow(pres, C, d), status_heatmap: (d) => renderStatusHeatmap(pres, C, d),
		two_by_two: (d) => renderTwoByTwo(pres, C, d), bridge: (d) => renderBridge(pres, C, d), recommendation: (d) => renderRecommendation(pres, C, d),
	};

	outline.slides.forEach((def, i) => {
		const render = renderers[def?.type];
		if (!render) fail(`slide ${i + 1}: unknown type "${def?.type}" (${Object.keys(renderers).join("|")})`);
		const slide = render(def);
		if (C.consult) addConsultingChrome(C, slide, def, meta, i + 1, outline.slides.length, ["title", "section", "closing"].includes(def.type));
		const note = def.note ?? def.notes;
		if (note) slide.addNotes(String(note));
	});

	const out = outArg ?? path.basename(outlinePath).replace(/\.json$/i, "") + ".pptx";
	pres.writeFile({ fileName: out }).then(() => {
		console.log(`wrote: ${out} (${outline.slides.length} slides)`);
		if (outline.slides.some((d) => d?.type === "chart" && !d.notes)) console.log("hint: add \u201cnotes\u201d to chart slides \u2014 the takeaway belongs in the talk track too");
	});
}

main();

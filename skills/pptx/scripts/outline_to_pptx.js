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
 *     { "type": "figure",  "title": "...", "image": "assets/flow1.svg",
 *       "caption": "...", "width": 8.0 },
 *     { "type": "closing", "title": "Thank you", "subtitle": "..." }
 *   ]
 * }
 *
 * The consulting types (exec_summary, kpi_row, axis_table, premise_conclusion,
 * issue_action, chevron_flow, status_heatmap, two_by_two, bridge, recommendation)
 * and the shared kicker/source/page chrome are documented in
 * references/consulting-layouts.md. Ten additional palette-independent types
 * (timeline, compare, quote, kpi_dashboard, issue_tree, process_steps,
 * metric_callout, agenda, risk_matrix, team_roster) are documented there too.
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
let assetContext = { dir: process.cwd(), manifest: null };

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

// Figure assets: images rendered by render_assets.py (manifest.json is optional).
function loadAssetContext(dir) {
	const manifestPath = path.join(dir, "assets", "manifest.json");
	if (!fs.existsSync(manifestPath)) return { dir, manifest: null };
	try {
		const entries = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
		if (!entries || typeof entries !== "object" || Array.isArray(entries)) throw new Error("manifest must be an object keyed by asset id");
		return { dir, manifest: { dir: path.dirname(manifestPath), entries } };
	} catch (err) {
		console.error(`outline: ignoring ${manifestPath}: ${err.message}`);
		return { dir, manifest: null };
	}
}

function imagePixels(file) {
	const ext = path.extname(file).toLowerCase();
	if (ext === ".png") {
		const buf = fs.readFileSync(file);
		if (buf.length >= 24 && buf.subarray(1, 4).toString() === "PNG") return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
	}
	if (ext === ".svg") {
		const text = fs.readFileSync(file, "utf8").slice(0, 8192);
		const vb = text.match(/viewBox\s*=\s*"\s*-?[\d.]+\s+-?[\d.]+\s+([\d.]+)\s+([\d.]+)/);
		if (vb && Number(vb[1]) > 0 && Number(vb[2]) > 0) return { w: Number(vb[1]), h: Number(vb[2]) };
		const wm = text.match(/\bwidth\s*=\s*"([\d.]+)px?"/), hm = text.match(/\bheight\s*=\s*"([\d.]+)px?"/);
		if (wm && hm && Number(wm[1]) > 0 && Number(hm[1]) > 0) return { w: Number(wm[1]), h: Number(hm[1]) };
	}
	return null;
}

function resolveFigure(image, def) {
	const raw = String(image).trim();
	let file = path.isAbsolute(raw) ? raw : path.resolve(assetContext.dir, raw);
	const m = assetContext.manifest;
	const entry = m && (m.entries[raw] ?? m.entries[path.basename(raw)] ?? m.entries[path.basename(raw, path.extname(raw))]);
	if (!fs.existsSync(file) && entry) {
		const rel = [entry.png, entry.path].find((p) => typeof p === "string" && fs.existsSync(path.resolve(m.dir, p)));
		if (rel) file = path.resolve(m.dir, rel);
	}
	if (!fs.existsSync(file)) fail(`figure slide "${def.title}": image not found: ${raw}`);
	// pptxgenjs embeds a broken PNG preview for SVG in Node; use the raster companion when present.
	if (path.extname(file).toLowerCase() === ".svg") {
		const raster = file.replace(/\.svg$/i, ".png");
		if (fs.existsSync(raster)) return raster;
	}
	return file;
}

function renderFigure(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	if (typeof def.image !== "string" || !def.image.trim()) fail(`figure slide "${def.title}" has no image`);
	const file = resolveFigure(def.image, def);
	const size = imagePixels(file);
	const want = def.width === undefined ? 8.0 : Number(def.width);
	if (!Number.isFinite(want) || want <= 0) fail(`figure slide "${def.title}": width must be a positive number of inches`);
	let w = Math.min(want, IN - 2 * M);
	let h = size ? w * size.h / size.w : w * 9 / 16;
	const maxH = BODY_H - (C.consult ? 1.25 : 0.75); // consulting source/page chrome starts at y=6.7
	if (h > maxH) {
		h = maxH;
		w = h * (size ? size.w / size.h : 16 / 9);
	}
	const imageOptions = {
		x: (IN - w) / 2, y: BODY_Y, w, h,
		sizing: { type: "contain", w, h },
		altText: String(def.caption ?? def.altText ?? def.title ?? "figure"),
	};
	if (path.extname(file).toLowerCase() === ".svg") {
		s.addImage({ ...imageOptions, data: "image/svg+xml;base64," + fs.readFileSync(file).toString("base64") });
	} else {
		s.addImage({ ...imageOptions, path: file });
	}
	if (def.caption) s.addText(String(def.caption), { x: M, y: BODY_Y + h + 0.15, w: IN - 2 * M, h: 0.5, fontFace: FONT.body, fontSize: 12, color: C.grey, align: "center", margin: 0, valign: "top" });
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

// Additional cross-genre layouts. Like the consulting types above, these use
// only palette tokens and addTitle(), so consulting chrome is added by main().
function renderTimeline(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const events = (Array.isArray(def.events) ? def.events : []).slice(0, 6);
	if (!events.length) fail(`timeline slide "${def.title}" has no events`);
	events.forEach((ev, i) => {
		if (!asText(ev?.date) || !asText(ev?.label)) fail(`timeline slide "${def.title}": event ${i + 1} needs date and label`);
	});
	const axisY = 3.85;
	const cell = (IN - 2 * M) / events.length;
	const pad = 0.18;
	rule(pres, s, C, M, axisY, IN - 2 * M);
	events.forEach((ev, i) => {
		const cx = M + (i + 0.5) * cell;
		const x = cx - cell / 2 + pad;
		const w = cell - 2 * pad;
		s.addShape(pres.ShapeType.ellipse, { x: cx - 0.11, y: axisY - 0.11, w: 0.22, h: 0.22, fill: { color: C.accent }, line: { type: "none" } });
		s.addText(asText(ev.date), { x, y: 3.15, w, h: 0.4, fontFace: FONT.body, fontSize: 13, bold: true, color: C.grey, align: "center", margin: 0, valign: "top" });
		const runs = [{ text: asText(ev.label), options: { fontSize: 14, bold: true, color: C.ink, breakLine: !!ev.text } }];
		if (ev.text) runs.push({ text: asText(ev.text), options: { fontSize: 14, color: C.grey } });
		s.addText(runs, { x, y: 4.2, w, h: 1.6, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 6 });
	});
	return s;
}

function renderCompare(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const t = def.compare ?? def;
	const headers = (t.headers ?? []).map(asText);
	const rows = (t.rows ?? []).slice(0, 7);
	if (headers.length < 2 || headers.length > 3) fail(`compare slide "${def.title}" needs 2-3 headers`);
	if (!rows.length) fail(`compare slide "${def.title}" has no rows`);
	rows.forEach((row, i) => {
		if (!Array.isArray(row) || row.length !== headers.length + 1) fail(`compare slide "${def.title}": row ${i + 1} needs a row label plus ${headers.length} values`);
	});
	const labelW = 2.6;
	const colW = [labelW, ...headers.map(() => (IN - 2 * M - labelW) / headers.length)];
	const head = [asText(t.labelHeader ?? "\u9805\u76ee"), ...headers].map((h, i) => ({
		text: h, options: { bold: true, fontSize: 14, color: C.ink, fill: { color: C.soft }, align: i === 0 ? "left" : "center", valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), EDGE(C.ink, 1.25), NO_BORDER()] },
	}));
	const body = rows.map((row, ri) => row.map((cell, i) => ({
		text: asText(cell), options: { fontSize: 14, bold: i === 0, color: C.ink, align: i === 0 ? "left" : "center", valign: "middle", margin: [4, 6, 4, 8], border: [NO_BORDER(), NO_BORDER(), ri === rows.length - 1 ? NO_BORDER() : EDGE(C.line), NO_BORDER()] },
	})));
	const rowH = [0.55, ...rows.map(() => Math.min(0.72, 4.4 / (rows.length + 1)))];
	s.addTable([head, ...body], { x: M, y: BODY_Y, w: IN - 2 * M, colW, rowH, fontFace: FONT.body, fontSize: 14, valign: "middle", autoPage: false });
	return s;
}

function renderQuote(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const quote = asText(def.quote ?? def.text);
	const attribution = asText(def.attribution ?? def.author ?? "");
	if (!quote) fail(`quote slide "${def.title}" has no quote`);
	if (!attribution) fail(`quote slide "${def.title}" has no attribution`);
	s.addText([
		{ text: "\u201c", options: { fontSize: 44, bold: true, color: C.accent } },
		{ text: quote, options: { fontSize: 24, bold: true, color: C.ink } },
	], { x: M, y: 2.3, w: IN - 2 * M, h: 2.6, fontFace: FONT.body, margin: 0, valign: "top", lineSpacingMultiple: 1.15 });
	rule(pres, s, C, M, 5.25, 2.0);
	s.addText(`\u2014 ${attribution}`, { x: M, y: 5.5, w: IN - 2 * M, h: 0.5, fontFace: FONT.body, fontSize: 16, color: C.grey, margin: 0, valign: "top" });
	return s;
}

function renderKpiDashboard(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const kpis = (Array.isArray(def.kpis) ? def.kpis : []).slice(0, 4);
	if (!kpis.length) fail(`kpi_dashboard slide "${def.title}" has no kpis`);
	kpis.forEach((kpi, i) => {
		if (!asText(kpi?.value) || !asText(kpi?.label)) fail(`kpi_dashboard slide "${def.title}": kpi ${i + 1} needs value and label`);
	});
	const gap = 0.5;
	const w = (IN - 2 * M - gap) / 2;
	const rowH = 1.9;
	kpis.forEach((kpi, i) => {
		const x = M + (i % 2) * (w + gap);
		const y = 1.9 + Math.floor(i / 2) * (rowH + 0.4);
		const delta = asText(kpi.delta);
		const arrow = /^[-|\u2212]/.test(delta) ? "\u25bc" : delta.startsWith("+") ? "\u25b2" : "\u2192";
		const spark = (Array.isArray(kpi.spark) ? kpi.spark : []).slice(0, 8).map(Number).filter(Number.isFinite);
		const runs = [
			{ text: asText(kpi.label), options: { fontSize: 14, bold: true, color: C.grey, breakLine: true } },
			{ text: asText(kpi.value), options: { fontSize: 40, bold: true, color: C.accent, breakLine: !!delta || spark.length >= 2 } },
		];
		if (delta) runs.push({ text: `${arrow} ${delta}`, options: { fontSize: 14, color: C.ink, breakLine: spark.length >= 2 } });
		if (spark.length >= 2) {
			const blocks = ["\u2581", "\u2582", "\u2583", "\u2584", "\u2585", "\u2586", "\u2587", "\u2588"];
			const lo = Math.min(...spark);
			const hi = Math.max(...spark);
			const glyphs = spark.map((v) => blocks[hi === lo ? 4 : Math.max(0, Math.min(7, Math.round((v - lo) / (hi - lo) * 7)))]).join("");
			runs.push({ text: glyphs, options: { fontSize: 16, bold: true, color: C.brand, charSpacing: 2 } });
		}
		s.addText(runs, { x, y, w, h: rowH, shape: pres.ShapeType.rect, fill: { color: C.soft }, line: { type: "none" }, fontFace: FONT.body, margin: [10, 14, 10, 14], valign: "top", paraSpaceAfter: 5 });
	});
	return s;
}

function renderIssueTree(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const root = asText(def.root ?? def.question);
	const branches = (Array.isArray(def.branches) ? def.branches : []).slice(0, 4);
	if (!root) fail(`issue_tree slide "${def.title}" has no root question`);
	if (branches.length < 3) fail(`issue_tree slide "${def.title}" needs 3-4 branches`);
	branches.forEach((b, i) => {
		if (!asText(b?.label) || !asText(b?.text)) fail(`issue_tree slide "${def.title}": branch ${i + 1} needs label and text`);
	});
	const gap = 0.4;
	const w = (IN - 2 * M - gap * (branches.length - 1)) / branches.length;
	const centers = branches.map((_, i) => M + (i + 0.5) * (w + gap));
	const stemY = 2.72;
	centers.forEach((cx) => {
		s.addShape(pres.ShapeType.rect, { x: cx - 0.006, y: stemY, w: 0.012, h: 3.3 - stemY, fill: { color: C.line }, line: { type: "none" } });
	});
	s.addShape(pres.ShapeType.rect, { x: centers[0] - 0.006, y: stemY - 0.006, w: centers[centers.length - 1] - centers[0] + 0.012, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
	s.addShape(pres.ShapeType.rect, { x: IN / 2 - 0.006, y: 2.57, w: 0.012, h: stemY - 2.57, fill: { color: C.line }, line: { type: "none" } });
	s.addText(root, { x: M, y: 1.85, w: IN - 2 * M, h: 0.72, shape: pres.ShapeType.rect, fill: { color: C.brand }, line: { type: "none" }, fontFace: FONT.body, fontSize: 16, bold: true, color: C.onBrand, margin: [8, 12, 8, 12], valign: "middle" });
	branches.forEach((branch, i) => {
		s.addText([
			{ text: asText(branch.label), options: { fontSize: 15, bold: true, color: C.ink, breakLine: true } },
			{ text: asText(branch.text), options: { fontSize: 14, color: C.ink } },
		], { x: M + i * (w + gap), y: 3.3, w, h: 2.2, shape: pres.ShapeType.rect, fill: { color: C.soft }, line: { type: "none" }, fontFace: FONT.body, margin: [12, 12, 12, 12], valign: "top", paraSpaceAfter: 8 });
	});
	return s;
}

function renderProcessSteps(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const steps = (Array.isArray(def.steps) ? def.steps : []).slice(0, 6);
	if (!steps.length) fail(`process_steps slide "${def.title}" has no steps`);
	steps.forEach((step, i) => {
		if (!asText(step?.title) || !asText(step?.text)) fail(`process_steps slide "${def.title}": step ${i + 1} needs title and text`);
	});
	const top = 1.9;
	const rowH = Math.min(0.95, 4.7 / steps.length);
	const dense = steps.length >= 6;
	for (let i = 0; i < steps.length - 1; i += 1) {
		s.addShape(pres.ShapeType.rect, { x: M + 0.216, y: top + i * rowH + 0.49, w: 0.012, h: rowH - 0.47, fill: { color: C.line }, line: { type: "none" } });
	}
	steps.forEach((step, i) => {
		const y = top + i * rowH;
		s.addText(String(i + 1).padStart(2, "0"), { x: M, y: y + 0.02, w: 0.45, h: 0.45, shape: pres.ShapeType.rect, fill: { color: C.soft }, line: { type: "none" }, fontFace: FONT.body, fontSize: 14, bold: true, color: C.accent, align: "center", valign: "middle", margin: 0 });
		s.addText([
			{ text: asText(step.title), options: { fontSize: dense ? 14 : 15, bold: true, color: C.ink, breakLine: true } },
			{ text: asText(step.text), options: { fontSize: 14, color: C.grey } },
		], { x: M + 0.8, y, w: IN - 2 * M - 0.8, h: Math.min(0.7, rowH - 0.32), fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: dense ? 0 : 4 });
	});
	return s;
}

function renderMetricCallout(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const value = asText(def.value);
	const label = asText(def.label);
	if (!value) fail(`metric_callout slide "${def.title}" has no value`);
	if (!label) fail(`metric_callout slide "${def.title}" has no label`);
	const unit = asText(def.unit);
	s.addText([
		{ text: value, options: { fontSize: 88, bold: true, color: C.accent } },
		...(unit ? [{ text: ` ${unit}`, options: { fontSize: 30, bold: true, color: C.ink } }] : []),
	], { x: M, y: 2.0, w: IN - 2 * M, h: 1.6, fontFace: FONT.head, margin: 0, valign: "top" });
	s.addText(label, { x: M, y: 3.9, w: IN - 2 * M, h: 0.55, fontFace: FONT.body, fontSize: 20, bold: true, color: C.ink, margin: 0, valign: "top" });
	if (def.context) s.addText(asText(def.context), { x: M, y: 4.75, w: IN - 2 * M - 1.0, h: 0.9, fontFace: FONT.body, fontSize: 16, color: C.grey, margin: 0, valign: "top" });
	return s;
}

function renderAgenda(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const sections = (Array.isArray(def.sections) ? def.sections : []).slice(0, 6);
	if (!sections.length) fail(`agenda slide "${def.title}" has no sections`);
	sections.forEach((sec, i) => {
		if (!asText(sec?.title)) fail(`agenda slide "${def.title}": section ${i + 1} needs a title`);
	});
	const rowH = Math.min(0.95, 4.8 / sections.length);
	const dense = sections.length >= 6;
	sections.forEach((sec, i) => {
		const y = 1.85 + i * rowH;
		s.addText(String(i + 1).padStart(2, "0"), { x: M, y, w: 1.05, h: Math.min(0.56, rowH - 0.32), fontFace: FONT.head, fontSize: dense ? 26 : 30, bold: true, color: C.accent, margin: 0, valign: "top" });
		const runs = [{ text: asText(sec.title), options: { fontSize: dense ? 16 : 18, bold: true, color: C.ink, breakLine: !!sec.text } }];
		if (sec.text) runs.push({ text: asText(sec.text), options: { fontSize: 14, color: C.grey } });
		s.addText(runs, { x: M + 1.4, y, w: IN - 2 * M - 1.4, h: Math.min(0.7, rowH - 0.31), fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: dense ? 0 : 4 });
	});
	return s;
}

function riskLevel(value, where) {
	const i = ["low", "mid", "high"].indexOf(String(value ?? "").toLowerCase());
	if (i < 0) fail(`${where}: expected low|mid|high, got "${value}"`);
	return i;
}

function renderRiskMatrix(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const risks = (Array.isArray(def.risks) ? def.risks : []).slice(0, 9);
	if (!risks.length) fail(`risk_matrix slide "${def.title}" has no risks`);
	const axes = def.axes ?? {};
	const xName = asText(axes.x?.name ?? "\u767a\u751f\u78ba\u7387");
	const yName = asText(axes.y?.name ?? "\u5f71\u97ff\u5ea6");
	const xLow = asText(axes.x?.low ?? "\u4f4e");
	const xHigh = asText(axes.x?.high ?? "\u9ad8");
	const yLow = asText(axes.y?.low ?? "\u4f4e");
	const yHigh = asText(axes.y?.high ?? "\u9ad8");
	risks.forEach((risk, i) => {
		if (!asText(risk?.label)) fail(`risk_matrix slide "${def.title}": risk ${i + 1} needs a label`);
		riskLevel(risk.impact, `risk_matrix slide "${def.title}": risk ${i + 1} impact`);
		riskLevel(risk.probability, `risk_matrix slide "${def.title}": risk ${i + 1} probability`);
	});
	s.addText(`\u6a2a\u8cc2\uff1a${xName}\u3000\u7e26\u8cc2\uff1a${yName}`, { x: M, y: 1.7, w: IN - 2 * M, h: 0.3, fontFace: FONT.body, fontSize: 14, color: C.grey, margin: 0, valign: "top" });
	const px = M + 1.2;
	const py = 2.35;
	const pw = IN - M - 0.2 - px;
	const ph = 3.5;
	const cw = pw / 3;
	const ch = ph / 3;
	for (let i = 0; i <= 3; i += 1) {
		s.addShape(pres.ShapeType.rect, { x: px + i * cw - 0.006, y: py, w: 0.012, h: ph, fill: { color: C.line }, line: { type: "none" } });
		s.addShape(pres.ShapeType.rect, { x: px, y: py + i * ch - 0.006, w: pw, h: 0.012, fill: { color: C.line }, line: { type: "none" } });
	}
	[[px, xLow, "left"], [px + pw - 1.2, xHigh, "right"]].forEach(([x, label, align]) => {
		s.addText(asText(label), { x, y: 6.1, w: 1.2, h: 0.3, fontFace: FONT.body, fontSize: 13, color: C.grey, align, margin: 0, valign: "top" });
	});
	[[py, yHigh], [py + 2 * ch, yLow]].forEach(([y, label]) => {
		s.addText(asText(label), { x: M, y, w: 1.0, h: 0.4, fontFace: FONT.body, fontSize: 13, color: C.grey, margin: 0, valign: "top" });
	});
	const cells = new Map();
	risks.forEach((risk) => {
		const key = `${riskLevel(risk.probability, "risk_matrix")}:${riskLevel(risk.impact, "risk_matrix")}`;
		cells.set(key, (cells.get(key) ?? 0) + 1);
	});
	cells.forEach((count, key) => {
		if (count > 2) fail(`risk_matrix slide "${def.title}": at most 2 risks per cell (grid ${key})`);
	});
	const seen = new Map();
	risks.forEach((risk) => {
		const prob = riskLevel(risk.probability, "risk_matrix");
		const impact = riskLevel(risk.impact, "risk_matrix");
		const key = `${prob}:${impact}`;
		const j = seen.get(key) ?? 0;
		seen.set(key, j + 1);
		const count = cells.get(key);
		const cx = px + (prob + 0.5) * cw;
		const cy = py + (2 - impact + 0.5) * ch + (j - (count - 1) / 2) * 0.62;
		const score = (impact + 1) * (prob + 1);
		const state = score <= 2 ? STATUS.ok : score <= 4 ? STATUS.warn : STATUS.risk;
		s.addShape(pres.ShapeType.ellipse, { x: cx - 1.15, y: cy - 0.07, w: 0.14, h: 0.14, fill: { color: state.color }, line: { type: "none" } });
		s.addText([
			{ text: `${state.mark} `, options: { fontSize: 14, bold: true, color: state.color } },
			{ text: asText(risk.label), options: { fontSize: 14, color: C.ink } },
		], { x: cx - 0.95, y: cy - 0.15, w: cw - 1.0, h: 0.3, fontFace: FONT.body, margin: 0, valign: "middle" });
	});
	return s;
}

function renderTeamRoster(pres, C, def) {
	const s = pres.addSlide();
	addTitle(pres, C, s, def.title ?? "");
	const members = (Array.isArray(def.members) ? def.members : []).slice(0, 6);
	if (!members.length) fail(`team_roster slide "${def.title}" has no members`);
	members.forEach((m, i) => {
		if (!asText(m?.name) || !asText(m?.role)) fail(`team_roster slide "${def.title}": member ${i + 1} needs name and role`);
	});
	const gap = 0.5;
	const w = (IN - 2 * M - gap) / 2;
	const rowH = 1.2;
	members.forEach((member, i) => {
		const x = M + (i % 2) * (w + gap);
		const y = 1.9 + Math.floor(i / 2) * (rowH + 0.45);
		const name = asText(member.name);
		const initials = asText(member.initials) || name.slice(0, 2);
		s.addText(initials, { x: x + 0.05, y: y + 0.2, w: 0.7, h: 0.7, shape: pres.ShapeType.rect, fill: { color: C.soft }, line: { type: "none" }, fontFace: FONT.body, fontSize: 18, bold: true, color: C.brand, align: "center", valign: "middle", margin: 0 });
		const runs = [
			{ text: name, options: { fontSize: 16, bold: true, color: C.ink, breakLine: true } },
			{ text: asText(member.role), options: { fontSize: 14, color: C.grey, breakLine: !!member.text } },
		];
		if (member.text) runs.push({ text: asText(member.text), options: { fontSize: 14, color: C.ink } });
		s.addText(runs, { x: x + 1.15, y: y + 0.15, w: w - 1.2, h: rowH - 0.25, fontFace: FONT.body, margin: 0, valign: "top", paraSpaceAfter: 4 });
	});
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
	assetContext = loadAssetContext(path.dirname(path.resolve(outlinePath)));

	const C = PALETTES[outline.palette ?? "ocean"] ?? fail(`unknown palette "${outline.palette}" (ocean|forest|clay|consulting)`);
	const meta = outline.meta ?? {};
	const pres = new pptxgen();
	pres.layout = "LAYOUT_WIDE";
	pres.author = meta.author ?? "pi-pptx";
	pres.title = meta.title ?? "Deck";

	const renderers = {
		title: () => renderTitle(pres, C, meta), section: (d) => renderSection(pres, C, d), bullets: (d) => renderBullets(pres, C, d),
		stats: (d) => renderStats(pres, C, d), chart: (d) => renderChart(pres, C, d), figure: (d) => renderFigure(pres, C, d), table: (d) => renderTable(pres, C, d), closing: (d) => renderClosing(pres, C, d),
		exec_summary: (d) => renderExecSummary(pres, C, d), kpi_row: (d) => renderKpiRow(pres, C, d), axis_table: (d) => renderAxisTable(pres, C, d),
		premise_conclusion: (d) => renderPremiseConclusion(pres, C, d), issue_action: (d) => renderIssueAction(pres, C, d),
		chevron_flow: (d) => renderChevronFlow(pres, C, d), status_heatmap: (d) => renderStatusHeatmap(pres, C, d),
		two_by_two: (d) => renderTwoByTwo(pres, C, d), bridge: (d) => renderBridge(pres, C, d), recommendation: (d) => renderRecommendation(pres, C, d),
		timeline: (d) => renderTimeline(pres, C, d), compare: (d) => renderCompare(pres, C, d), quote: (d) => renderQuote(pres, C, d),
		kpi_dashboard: (d) => renderKpiDashboard(pres, C, d), issue_tree: (d) => renderIssueTree(pres, C, d),
		process_steps: (d) => renderProcessSteps(pres, C, d), metric_callout: (d) => renderMetricCallout(pres, C, d),
		agenda: (d) => renderAgenda(pres, C, d), risk_matrix: (d) => renderRiskMatrix(pres, C, d), team_roster: (d) => renderTeamRoster(pres, C, d),
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

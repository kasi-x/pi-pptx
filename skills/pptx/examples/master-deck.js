/**
 * Master-slide starter — branded masters with placeholders, sections, slide numbers.
 * Two masters: "MAIN" (title + body placeholder for text slides) and "DATA"
 * (title only, for table/chart slides so no empty body placeholder lingers).
 * Copy this when a deck must look consistent across many slides.
 *
 * Run: node examples/master-deck.js   (writes ./master-deck.pptx)
 */
const pptxgen = require("pptxgenjs");

const FONT = { head: "Cambria", body: "Calibri" };
const C = { ink: "1F2933", paper: "FFFFFF", navy: "16324F", teal: "2A9D8F", grey: "5C6B73", line: "E5E9EC" };

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "pi-pptx";

// build fresh option objects per use — pptxgenjs mutates them in place (see references/pptxgenjs.md)
const slideNumber = () => ({ x: 12.4, y: 7.05, w: 0.6, h: 0.3, color: "9AA5AD", fontSize: 10, align: "right" });
const motif = () => ({ rect: { x: 0.7, y: 0.62, w: 0.18, h: 0.18, fill: { color: C.teal } } });
const titlePh = () => ({ placeholder: { options: { name: "title", type: "title", x: 1.05, y: 0.5, w: 11.6, h: 0.6, fontFace: FONT.head, fontSize: 26, bold: true, color: C.ink, margin: 0, valign: "middle" } } });

// Text slides: title + a body placeholder.
p.defineSlideMaster({
	title: "MAIN",
	background: { color: C.paper },
	objects: [
		motif(),
		titlePh(),
		{ placeholder: { options: { name: "body", type: "body", x: 0.7, y: 1.5, w: 11.9, h: 5.2, fontFace: FONT.body, fontSize: 16, color: C.ink, margin: 0, valign: "top" } } },
	],
	slideNumber: slideNumber(),
});

// Data slides (tables/charts): title only — no unused body placeholder to collide with.
p.defineSlideMaster({
	title: "DATA",
	background: { color: C.paper },
	objects: [motif(), titlePh()],
	slideNumber: slideNumber(),
});

p.addSection({ title: "Review" });
const s1 = p.addSlide({ masterName: "MAIN", sectionTitle: "Review" });
s1.addText("Progress held; capacity is the next constraint", { placeholder: "title" });
s1.addText(
	[
		{ text: "Revenue +18% with flat headcount", options: { bullet: true, breakLine: true } },
		{ text: "Review cycle 2.4x faster after the automation rollout", options: { bullet: true, breakLine: true } },
		{ text: "Q4 volume crosses planned capacity — decision needed this month", options: { bullet: true } },
	],
	{ placeholder: "body" },
);
s1.addNotes("State the takeaway first, then walk the three bullets.");

p.addSection({ title: "Plan" });
const s2 = p.addSlide({ masterName: "DATA", sectionTitle: "Plan" });
s2.addText("Capacity upgrade pays back within one quarter", { placeholder: "title" });
s2.addTable(
	[
		[
			{ text: "Workstream", options: { bold: true, fill: { color: C.navy }, color: C.paper } },
			{ text: "Owner", options: { bold: true, fill: { color: C.navy }, color: C.paper } },
			{ text: "Done by", options: { bold: true, fill: { color: C.navy }, color: C.paper } },
		],
		[
			{ text: "Capacity upgrade", options: { color: C.ink } },
			{ text: "Platform", options: { color: C.ink } },
			{ text: "Week 3", options: { color: C.ink } },
		],
		[
			{ text: "Review automation", options: { color: C.ink } },
			{ text: "Tooling", options: { color: C.ink } },
			{ text: "Week 5", options: { color: C.ink } },
		],
	],
	{ x: 0.7, y: 1.7, w: 11.9, colW: [6, 3.2, 2.7], rowH: 0.55, fontSize: 15, fontFace: FONT.body, border: { type: "solid", color: C.line, pt: 0.5 }, valign: "middle" },
);
s2.addText("Assumes the Q4 forecast; revisit if volume growth exceeds 20%.", { x: 0.7, y: 4.6, w: 11.9, h: 0.5, fontSize: 12, color: C.grey, margin: 0, italic: true });
s2.addNotes("The payback case, then the assumption to watch.");

const s3 = p.addSlide({ masterName: "DATA", sectionTitle: "Plan" });
s3.addText("Volume keeps climbing", { placeholder: "title" });
s3.addChart(
	p.ChartType.line,
	[{ name: "requests", labels: ["Q1", "Q2", "Q3", "Q4"], values: [120, 180, 260, 410] }],
	{
		x: 0.7, y: 1.7, w: 7.8, h: 5.0,
		showTitle: false, showLegend: false, showValue: false,
		chartColors: [C.teal], lineDataSymbol: "circle",
		catAxisLabelColor: C.grey, valAxisLabelColor: C.grey,
		catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
		valGridLine: { color: C.line, size: 1 }, catGridLine: { style: "none" },
	},
);
s3.addText("Requests grew 3.4x; Q4 is the first quarter above planned capacity.", { x: 8.8, y: 1.7, w: 3.8, h: 2.4, fontSize: 18, color: C.ink, margin: 0, lineSpacingMultiple: 1.2 });
s3.addNotes("Emphasize the capacity inflection, not the raw growth.");

p.writeFile({ fileName: "master-deck.pptx" }).then((f) => console.log("wrote:", f));

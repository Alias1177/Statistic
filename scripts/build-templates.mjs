// Generates two English sample .docx files into public/templates/.
// Run: node scripts/build-templates.mjs
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
} from "docx";
import { writeFile, mkdir } from "node:fs/promises";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(__dirname, "..", "public", "templates");

const fmt = n => n.toLocaleString("en-US");
const pct = (a, b) => "-" + (((a - b) / a) * 100).toFixed(1) + "%";

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };

function cell(text, opts = {}) {
  return new TableCell({
    borders,
    width: { size: opts.w || 2000, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    shading: opts.head ? { fill: "E7EEF7", type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), bold: !!opts.head, size: 20 })],
    })],
  });
}
const row = cells => new TableRow({ children: cells });

function buildObjectsTable({ monthA, monthB, labelA, labelB, objects }) {
  const widths = [3200, 2000, 2000, 1200];
  const total = widths.reduce((s, w) => s + w, 0);
  const totalA = objects.reduce((s, o) => s + o.a, 0);
  const totalB = objects.reduce((s, o) => s + o.b, 0);
  const head = row([
    cell("Site",                           { w: widths[0], head: true }),
    cell(`${monthA} (${labelA}), kWh`,     { w: widths[1], head: true, align: AlignmentType.RIGHT }),
    cell(`${monthB} (${labelB}), kWh`,     { w: widths[2], head: true, align: AlignmentType.RIGHT }),
    cell("Saved",                          { w: widths[3], head: true, align: AlignmentType.RIGHT }),
  ]);
  const rows = objects.map(o => row([
    cell(o.name,         { w: widths[0] }),
    cell(fmt(o.a),       { w: widths[1], align: AlignmentType.RIGHT }),
    cell(fmt(o.b),       { w: widths[2], align: AlignmentType.RIGHT }),
    cell(pct(o.a, o.b),  { w: widths[3], align: AlignmentType.RIGHT }),
  ]));
  const totalRow = row([
    cell("TOTAL",                          { w: widths[0], head: true }),
    cell(fmt(totalA),                      { w: widths[1], head: true, align: AlignmentType.RIGHT }),
    cell(fmt(totalB),                      { w: widths[2], head: true, align: AlignmentType.RIGHT }),
    cell(pct(totalA, totalB),              { w: widths[3], head: true, align: AlignmentType.RIGHT }),
  ]);
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...rows, totalRow] });
}

function buildDaysTable(rows) {
  const widths = [2400, 6000];
  const total = widths.reduce((s, w) => s + w, 0);
  const head = row([
    cell("Day",                  { w: widths[0], head: true }),
    cell("Consumption (kWh)",    { w: widths[1], head: true, align: AlignmentType.RIGHT }),
  ]);
  const data = rows.map(d => row([
    cell(d.day,                  { w: widths[0] }),
    cell(fmt(d.value),           { w: widths[1], align: AlignmentType.RIGHT }),
  ]));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...data] });
}

function buildLoadsTable(rows) {
  const widths = [3600, 2800, 2000];
  const total = widths.reduce((s, w) => s + w, 0);
  const head = row([
    cell("Load type",            { w: widths[0], head: true }),
    cell("Consumption (kWh)",    { w: widths[1], head: true, align: AlignmentType.RIGHT }),
    cell("Share",                { w: widths[2], head: true, align: AlignmentType.RIGHT }),
  ]);
  const data = rows.map(d => row([
    cell(d.name,                 { w: widths[0] }),
    cell(fmt(d.v),               { w: widths[1], align: AlignmentType.RIGHT }),
    cell(d.p,                    { w: widths[2], align: AlignmentType.RIGHT }),
  ]));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...data] });
}

const P = (text, opts = {}) => new Paragraph({
  spacing: { before: 80, after: 80 },
  children: [new TextRun({ text, bold: !!opts.bold, size: opts.size || 22, italics: !!opts.italic })],
});
const H1 = text => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 120 },
  children: [new TextRun({ text, bold: true, size: 30 })],
});
const H2 = text => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, bold: true, size: 26 })],
});
const SP = () => new Paragraph({ children: [new TextRun("")] });

function makeDoc({ title, intro, monthA, monthB, labelA, labelB, objects, daysA, daysB, loadsA, loadsB, source }) {
  return new Document({
    styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      children: [
        H1(title),
        P(intro),
        SP(),
        H2("1. Monthly summary"),
        P("Total consumption per zone before and after the rollout."),
        buildObjectsTable({ monthA, monthB, labelA, labelB, objects }),
        SP(),
        H2("2. Daily breakdown (sample week)"),
        P(`${monthA} (${labelA})`, { bold: true }),
        buildDaysTable(daysA),
        SP(),
        P(`${monthB} (${labelB})`, { bold: true }),
        buildDaysTable(daysB),
        SP(),
        H2("3. Load mix"),
        P(`${monthA} (${labelA})`, { bold: true }),
        buildLoadsTable(loadsA),
        SP(),
        P(`${monthB} (${labelB})`, { bold: true }),
        buildLoadsTable(loadsB),
        SP(),
        P(source, { italic: true, size: 18 }),
      ],
    }],
  });
}

// =================== TEMPLATE 1: University campus ===================
const campus = makeDoc({
  title: "Energy efficiency report — Riverside University Campus",
  intro: "Comparative analysis of electricity consumption between March and April 2026 after rolling out the smart sensor and load-control system.",
  monthA: "March 2026", monthB: "April 2026",
  labelA: "no system",  labelB: "with system",
  objects: [
    { name: "Academic Buildings",    a: 52000, b: 44500 },
    { name: "Dormitories",           a: 38000, b: 34200 },
    { name: "Administration",        a: 12500, b: 10300 },
    { name: "Laboratories",          a: 28000, b: 25800 },
    { name: "Outdoor Lighting",      a:  9500, b:  6900 },
  ],
  daysA: [
    { day: "03/01", value: 4600 }, { day: "03/02", value: 4550 },
    { day: "03/03", value: 4700 }, { day: "03/04", value: 4650 },
    { day: "03/05", value: 4800 }, { day: "03/06", value: 4900 },
    { day: "03/07", value: 4750 },
  ],
  daysB: [
    { day: "04/01", value: 4100 }, { day: "04/02", value: 4050 },
    { day: "04/03", value: 4200 }, { day: "04/04", value: 4150 },
    { day: "04/05", value: 4250 }, { day: "04/06", value: 4300 },
    { day: "04/07", value: 4120 },
  ],
  loadsA: [
    { name: "Lighting",     v: 35000, p: "25%" },
    { name: "HVAC",         v: 45000, p: "32%" },
    { name: "Equipment",    v: 38000, p: "27%" },
    { name: "Other",        v: 22000, p: "16%" },
  ],
  loadsB: [
    { name: "Lighting",     v: 25000, p: "20%" },
    { name: "HVAC",         v: 42000, p: "34%" },
    { name: "Equipment",    v: 36000, p: "30%" },
    { name: "Other",        v: 18700, p: "16%" },
  ],
  source: "Source: campus metering systems · automated report.",
});

// =================== TEMPLATE 2: Shopping mall ===================
const mall = makeDoc({
  title: "Energy efficiency report — Atlas Shopping Mall",
  intro: "Comparative analysis of electricity consumption between June and July 2026 after upgrading the lighting and HVAC control systems.",
  monthA: "June 2026",            monthB: "July 2026",
  labelA: "before upgrade",       labelB: "after upgrade",
  objects: [
    { name: "Retail Gallery",         a: 86000, b: 71200 },
    { name: "Food Court",             a: 42000, b: 36500 },
    { name: "Cinema (5 screens)",     a: 31000, b: 27800 },
    { name: "Underground Parking",    a: 18500, b: 12100 },
    { name: "Technical Rooms",        a: 22500, b: 19800 },
    { name: "Outdoor Lighting",       a: 14000, b:  8400 },
  ],
  daysA: [
    { day: "06/01", value: 7100 }, { day: "06/02", value: 7250 },
    { day: "06/03", value: 7320 }, { day: "06/04", value: 7480 },
    { day: "06/05", value: 7950 }, { day: "06/06", value: 8120 },
    { day: "06/07", value: 7780 },
  ],
  daysB: [
    { day: "07/01", value: 5800 }, { day: "07/02", value: 5920 },
    { day: "07/03", value: 6010 }, { day: "07/04", value: 6180 },
    { day: "07/05", value: 6450 }, { day: "07/06", value: 6520 },
    { day: "07/07", value: 6230 },
  ],
  loadsA: [
    { name: "Lighting",            v: 62000, p: "29%" },
    { name: "HVAC (climate)",      v: 78000, p: "37%" },
    { name: "Escalators & lifts",  v: 35000, p: "16%" },
    { name: "Refrigeration",       v: 25000, p: "12%" },
    { name: "Other",               v: 14000, p: "6%"  },
  ],
  loadsB: [
    { name: "Lighting (LED)",      v: 38000, p: "22%" },
    { name: "HVAC (smart)",        v: 64000, p: "37%" },
    { name: "Escalators & lifts",  v: 31000, p: "18%" },
    { name: "Refrigeration",       v: 24000, p: "14%" },
    { name: "Other",               v: 14800, p: "9%"  },
  ],
  source: "Source: mall metering systems · automated report.",
});

await mkdir(OUT_DIR, { recursive: true });
const buf1 = await Packer.toBuffer(campus);
await writeFile(join(OUT_DIR, "template-campus.docx"), buf1);
console.log("Wrote template-campus.docx", buf1.length, "bytes");

const buf2 = await Packer.toBuffer(mall);
await writeFile(join(OUT_DIR, "template-mall.docx"), buf2);
console.log("Wrote template-mall.docx", buf2.length, "bytes");

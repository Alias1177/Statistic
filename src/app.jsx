const { useState, useEffect, useRef, useMemo, useCallback } = React;

// =================== DEFAULT DATA ===================
const DEFAULT_DATA = {
  place: "PITI Campus",
  monthA: "March 2026",
  monthB: "April 2026",
  monthALabel: "no system",
  monthBLabel: "with system",
  objects: [
    { name: "Academic Buildings",   a: 52000, b: 44500 },
    { name: "Dormitories",          a: 38000, b: 34200 },
    { name: "Administration",       a: 12500, b: 10300 },
    { name: "Laboratories",         a: 28000, b: 25800 },
    { name: "Outdoor Lighting",     a:  9500, b:  6900 },
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
    { name: "Lighting",    value: 35000, color: "#f59e0b" },
    { name: "HVAC",        value: 45000, color: "#3b82f6" },
    { name: "Equipment",   value: 38000, color: "#8b5cf6" },
    { name: "Other",       value: 22000, color: "#6b7280" },
  ],
  loadsB: [
    { name: "Lighting",    value: 25000, color: "#10b981" },
    { name: "HVAC",        value: 42000, color: "#3b82f6" },
    { name: "Equipment",   value: 36000, color: "#8b5cf6" },
    { name: "Other",       value: 18700, color: "#6b7280" },
  ],
};

// =================== UTILS ===================
const fmt  = n => Math.round(n).toLocaleString("en-US");
const fmt1 = n => n.toFixed(1);
const parseNum = s => {
  if (typeof s !== "string") return NaN;
  const cleaned = s.replace(/[\s ]/g, "").replace(",", ".").replace(/[^0-9.\-]/g, "");
  return cleaned === "" ? NaN : parseFloat(cleaned);
};

function useInView(threshold = 0.12) {
  const ref = useRef(null);
  const [inView, setInView] = useState(false);
  useEffect(() => {
    if (!ref.current) return;
    const obs = new IntersectionObserver(
      ([e]) => { if (e.isIntersecting) { setInView(true); obs.disconnect(); } },
      { threshold }
    );
    obs.observe(ref.current);
    return () => obs.disconnect();
  }, [threshold]);
  return [ref, inView];
}

function useCountUp(target, trigger, duration = 1200) {
  const [val, setVal] = useState(0);
  useEffect(() => {
    if (!trigger) return;
    let start = null;
    const ease = t => 1 - Math.pow(1 - t, 3);
    const step = ts => {
      if (!start) start = ts;
      const p = Math.min((ts - start) / duration, 1);
      setVal(target * ease(p));
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }, [trigger, target]);
  return val;
}

// =================== DOCX PARSER (bilingual: English + Russian) ===================
const PALETTE = ["#f59e0b", "#3b82f6", "#8b5cf6", "#10b981", "#ec4899", "#6b7280"];
// Header keywords in EN and RU
const HEADER_WORDS = /^(object|building|facility|day|date|type|category|name|total|sum|объект|здание|корпус|день|дата|тип|категория|итого|всего)/i;

function isHeaderRow(row) {
  if (!row || !row[0]) return true;
  if (HEADER_WORDS.test(row[0].trim())) return true;
  return false;
}

function classifyTable(rows) {
  const dataRows = rows.filter(r => r && r.length > 0 && !isHeaderRow(r));
  if (dataRows.length === 0) return null;
  const s = dataRows[0];

  // Days: first cell looks like a date — DD.MM, DD-MM, DD/MM, MM/DD, YYYY-MM-DD, or "Mar 1" / "March 1"
  const dateLike = /^\d{1,2}[.\-/]\d{1,2}([.\-/]\d{2,4})?$|^\d{4}-\d{2}-\d{2}$|^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{1,2}/i;
  if (dateLike.test(s[0].trim())) {
    return { type: "days", rows: dataRows };
  }
  // Loads: 3 cols, last col contains % (share)
  if (s.length >= 3 && /%$/.test(s[2].trim()) && !isNaN(parseNum(s[1]))) {
    return { type: "loads", rows: dataRows };
  }
  // Objects: 3+ cols, cols 1 and 2 are both numeric
  if (s.length >= 3 && !isNaN(parseNum(s[1])) && !isNaN(parseNum(s[2]))) {
    return { type: "objects", rows: dataRows };
  }
  // 2-col numeric → assume days (alternative format)
  if (s.length === 2 && !isNaN(parseNum(s[1]))) {
    return { type: "days", rows: dataRows };
  }
  return null;
}

async function parseDocx(file, fileName) {
  let zip;
  try {
    zip = await JSZip.loadAsync(file);
  } catch (e) {
    throw new Error("File can't be opened as an archive. Maybe it's a PDF, image, or a corrupted file rather than .docx.");
  }
  const docFile = zip.file("word/document.xml");
  if (!docFile) {
    if (zip.file(/^Index\/Document\.iwa$/) || zip.file(/^Index\/Tables\//)) {
      throw new Error("This is an Apple Pages (iWork) file, not Word. In Pages: File → Export To → Word (.docx), then upload the resulting file.");
    }
    if (zip.file("xl/workbook.xml")) {
      throw new Error("This is an Excel file (.xlsx). Only Word (.docx) is supported. Copy the tables into Word or save as .docx.");
    }
    if (zip.file("ppt/presentation.xml")) {
      throw new Error("This is a PowerPoint presentation. Only Word (.docx) is supported.");
    }
    if (zip.file("META-INF/manifest.xml")) {
      throw new Error("This is an OpenDocument file (.odt). In LibreOffice: File → Save As → Word 2007–365 (.docx).");
    }
    throw new Error("No word/document.xml inside the archive — this isn't a Word document. Only .docx saved from Microsoft Word is supported.");
  }
  const xml = await docFile.async("text");
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const body = dom.getElementsByTagNameNS(NS, "body")[0];
  if (!body) throw new Error("Document body is missing.");

  function paraText(p) {
    const ts = p.getElementsByTagNameNS(NS, "t");
    let s = "";
    for (let i = 0; i < ts.length; i++) s += ts[i].textContent || "";
    return s.trim();
  }

  const items = [];
  for (let n = body.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    if (n.localName === "p") items.push({ kind: "p", text: paraText(n) });
    else if (n.localName === "tbl") {
      const rows = [];
      const trs = n.getElementsByTagNameNS(NS, "tr");
      for (let i = 0; i < trs.length; i++) {
        const tcs = trs[i].getElementsByTagNameNS(NS, "tc");
        const cells = [];
        for (let j = 0; j < tcs.length; j++) {
          const ps = tcs[j].getElementsByTagNameNS(NS, "p");
          let txt = "";
          for (let k = 0; k < ps.length; k++) txt += paraText(ps[k]) + " ";
          cells.push(txt.trim());
        }
        rows.push(cells);
      }
      const counts = rows.map(r => r.length);
      const norm = counts.sort((a,b)=>a-b)[Math.floor(counts.length/2)] || 0;
      const cleanedRows = rows.filter(r => r.length <= norm * 1.5 || r.length <= 6);
      items.push({ kind: "tbl", rows: cleanedRows.length ? cleanedRows : rows });
    }
  }

  // Detect month labels (EN + RU) from any paragraph or first row
  const allText = items.map(it => it.kind === "p" ? it.text : it.rows.map(r => r.join(" ")).join(" ")).join("\n");
  let monthA = "Period A", monthB = "Period B";
  let monthALabel = "before", monthBLabel = "after";
  const monthRe = /(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь|Декабрь)\s+\d{4}/gi;
  const months = [...new Set(allText.match(monthRe) || [])];
  if (months[0]) monthA = months[0];
  if (months[1]) monthB = months[1];
  // Period labels: EN
  if (/before\s+(install|upgrad|retrofit|smart)/i.test(allText)) monthALabel = "before upgrade";
  if (/after\s+(install|upgrad|retrofit|smart)/i.test(allText))  monthBLabel = "after upgrade";
  if (/no\s+(system|sensors)/i.test(allText))   monthALabel = "no system";
  if (/with\s+(system|sensors)/i.test(allText)) monthBLabel = "with system";
  // Period labels: RU
  if (/без\s+систем|без\s+датчик/i.test(allText)) monthALabel = "no system";
  if (/с\s+систем|с\s+датчик/i.test(allText))   monthBLabel = "with system";
  if (/до\s+модерниз/i.test(allText))           monthALabel = "before upgrade";
  if (/после\s+модерниз/i.test(allText))        monthBLabel = "after upgrade";

  // Extract place name from heading or filename.
  let place = null;
  const sectionWord = /^(monthly|summary|breakdown|daily|daily\s+detail|details|loads|by\s+load|structure|total|analysis|table|method|сводные|детализац|разбивк|структура|итог|анализ|таблиц|метод)/i;
  const firstParas = items.filter(it => it.kind === "p" && it.text).slice(0, 8);
  // Pattern 1: explicit prefix "Site:", "Building:", "Client:", etc.
  for (const it of firstParas) {
    const m = it.text.match(/^(?:site|facility|building|company|organization|client|объект|компания|организац[ия]|здание|клиент)[\s:—–-]+(.+)$/i);
    if (m && m[1].trim()) {
      let cand = m[1].trim().replace(/^[«"](.+)[»"]$/, "$1");
      place = cand;
      break;
    }
  }
  // Pattern 2: first paragraph as title (split on em dash)
  if (!place && firstParas[0]) {
    const t = firstParas[0].text;
    if (!sectionWord.test(t) && !/^\d+[.)]/.test(t) && t.length < 140) {
      let m = t.match(/[—–]\s*(.+)$/);
      if (m && m[1].trim()) {
        let cand = m[1].trim();
        const opens = (cand.match(/\(/g) || []).length;
        const closes = (cand.match(/\)/g) || []).length;
        if (closes <= opens) place = cand.replace(/^[«"](.+)[»"]$/, "$1");
      }
      if (!place && t.length < 80) place = t.replace(/^[«"](.+)[»"]$/, "$1");
    }
  }
  // Fallback: filename
  if (!place && fileName) {
    place = fileName.replace(/\.[^.]+$/, "").replace(/[_]+/g, " ").trim();
    place = place.replace(/^(template|sample|шаблон)[\s\-—–]*\d*[\s\-—–]*/i, "").trim() || place;
  }
  if (!place) place = "Uploaded Site";

  const tables = items.filter(it => it.kind === "tbl").map(it => classifyTable(it.rows));
  const objectsTable = tables.find(t => t && t.type === "objects");
  const daysTables   = tables.filter(t => t && t.type === "days");
  const loadsTables  = tables.filter(t => t && t.type === "loads");

  if (!objectsTable) throw new Error("Couldn't find the objects table (3+ columns, columns 2 and 3 numeric).");

  const objects = objectsTable.rows
    .map(r => ({ name: r[0], a: parseNum(r[1]), b: parseNum(r[2]) }))
    .filter(o => o.name && !isNaN(o.a) && !isNaN(o.b) && !/итого|total/i.test(o.name));

  const toDays = t => t.rows
    .map(r => ({ day: r[0], value: parseNum(r[1]) }))
    .filter(d => d.day && !isNaN(d.value));
  const toLoads = t => t.rows
    .map((r, i) => ({ name: r[0], value: parseNum(r[1]), color: PALETTE[i % PALETTE.length] }))
    .filter(d => d.name && !isNaN(d.value));

  const daysA  = daysTables[0]  ? toDays(daysTables[0])  : [];
  const daysB  = daysTables[1]  ? toDays(daysTables[1])  : [];
  const loadsA = loadsTables[0] ? toLoads(loadsTables[0]) : [];
  const loadsB = loadsTables[1] ? toLoads(loadsTables[1]) : [];

  if (objects.length === 0) throw new Error("No numeric rows found in the objects table.");

  return {
    place,
    monthA, monthB, monthALabel, monthBLabel,
    objects,
    daysA: daysA.length ? daysA : DEFAULT_DATA.daysA,
    daysB: daysB.length ? daysB : DEFAULT_DATA.daysB,
    loadsA: loadsA.length ? loadsA : DEFAULT_DATA.loadsA,
    loadsB: loadsB.length ? loadsB : DEFAULT_DATA.loadsB,
  };
}

// =================== ICONS ===================
const Icon = {
  home: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 12l9-9 9 9"/><path d="M5 10v10h14V10"/></svg>,
  chart:<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 3v18h18"/><path d="M7 14l4-4 4 4 5-7"/></svg>,
  upload:<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
  meter:<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>,
  plus: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  cog:  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>,
  bolt: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>,
  leaf: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 20A7 7 0 0 1 4 13c0-5.5 6-9 14-11 0 4-1 13-7 18z"/><path d="M2 22c5-3 8-6 9-12"/></svg>,
  trend:<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="22 17 13 8 9 12 2 5"/><polyline points="16 17 22 17 22 11"/></svg>,
  download:<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
};

// =================== KPI ===================
function KPI({ label, value, unit, delta, sub, icon, iconColor, idx = 0, decimals = 0 }) {
  const [ref, inView] = useInView();
  const v = useCountUp(value, inView);
  return (
    <div ref={ref} className={"kpi " + (inView ? "in" : "")} style={{ animationDelay: (idx * 0.08) + "s" }}>
      <div className={"kpi-icon " + iconColor}>{icon}</div>
      <div className="label">{label}</div>
      <div className="value">{decimals ? fmt1(v) : fmt(v)}{unit && <span className="unit">{unit}</span>}</div>
      {delta && <div className="delta">{delta}</div>}
      {sub && <div style={{ marginTop: 4, fontSize: 12, color: "var(--text-dim)" }}>{sub}</div>}
    </div>
  );
}

// =================== BAR CHART ===================
function BarCompare({ data, labelA, labelB }) {
  const [ref, inView] = useInView();
  const [tip, setTip] = useState(null);
  const W = 760, H = 360, PAD = { l: 50, r: 20, t: 20, b: 60 };
  const innerW = W - PAD.l - PAD.r;
  const innerH = H - PAD.t - PAD.b;
  const maxV = Math.max(1, ...data.map(d => Math.max(d.a, d.b))) * 1.1;
  const groupW = innerW / Math.max(1, data.length);
  const barW = Math.min(40, groupW * 0.32);
  const ticks = 5;
  const yTicks = Array.from({ length: ticks + 1 }, (_, i) => Math.round((maxV / ticks) * i / 1000) * 1000);

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <div style={{ position: "relative" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", overflow: "visible" }}>
        <defs>
          <linearGradient id="gA" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#cbd5e1"/><stop offset="100%" stopColor="#94a3b8"/>
          </linearGradient>
          <linearGradient id="gB" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#34d399"/><stop offset="100%" stopColor="#059669"/>
          </linearGradient>
        </defs>
        {yTicks.map((t, i) => {
          const y = PAD.t + innerH - (t / maxV) * innerH;
          return (<g key={i}>
            <line x1={PAD.l} x2={W - PAD.r} y1={y} y2={y} stroke="rgba(0,0,0,0.06)"/>
            <text x={PAD.l - 8} y={y + 4} textAnchor="end" fontSize="10" fill="var(--text-mute)">{(t/1000)+"k"}</text>
          </g>);
        })}
        {data.map((d, i) => {
          const cx = PAD.l + groupW * i + groupW / 2;
          const hA = (d.a / maxV) * innerH;
          const hB = (d.b / maxV) * innerH;
          const delay = i * 0.08;
          return (<g key={i}>
            <rect className="bar" x={cx - barW - 3} y={PAD.t + innerH - hA} width={barW} height={hA}
                  fill="url(#gA)" rx="3"
                  onMouseEnter={() => setTip({ x: cx - barW/2 - 3, y: PAD.t + innerH - hA, text: `${labelA}: ${fmt(d.a)} kWh` })}
                  onMouseLeave={() => setTip(null)}>
              <animate attributeName="height" from="0" to={hA} dur="0.9s" begin={inView ? `${delay}s` : "indefinite"} fill="freeze"/>
              <animate attributeName="y" from={PAD.t + innerH} to={PAD.t + innerH - hA} dur="0.9s" begin={inView ? `${delay}s` : "indefinite"} fill="freeze"/>
            </rect>
            <rect className="bar" x={cx + 3} y={PAD.t + innerH - hB} width={barW} height={hB}
                  fill="url(#gB)" rx="3"
                  onMouseEnter={() => setTip({ x: cx + barW/2 + 3, y: PAD.t + innerH - hB, text: `${labelB}: ${fmt(d.b)} kWh` })}
                  onMouseLeave={() => setTip(null)}>
              <animate attributeName="height" from="0" to={hB} dur="0.9s" begin={inView ? `${delay+0.12}s` : "indefinite"} fill="freeze"/>
              <animate attributeName="y" from={PAD.t + innerH} to={PAD.t + innerH - hB} dur="0.9s" begin={inView ? `${delay+0.12}s` : "indefinite"} fill="freeze"/>
            </rect>
            <text x={cx} y={H - 30} textAnchor="middle" fontSize="11" fill="var(--text-dim)">
              {d.name.length > 15 ? d.name.slice(0,13) + "…" : d.name}
            </text>
            {d.a > 0 && <text x={cx} y={H - 14} textAnchor="middle" fontSize="11" fill="var(--accent)" fontWeight="600" opacity="0">
              −{fmt1(((d.a - d.b) / d.a) * 100)}%
              <animate attributeName="opacity" from="0" to="1" dur="0.5s" begin={inView ? `${delay + 0.7}s` : "indefinite"} fill="freeze"/>
            </text>}
          </g>);
        })}
      </svg>
      {tip && (
        <div className="tooltip" style={{ left: `${(tip.x / W) * 100}%`, top: `${(tip.y / H) * 100}%` }}>{tip.text}</div>
      )}
      </div>
      <div className="legend">
        <span><span className="swatch" style={{ background: "linear-gradient(180deg, #cbd5e1, #94a3b8)" }}></span>{labelA}</span>
        <span><span className="swatch" style={{ background: "linear-gradient(180deg, #34d399, #059669)" }}></span>{labelB}</span>
      </div>
    </div>
  );
}

// =================== LINE CHART ===================
function LineChart({ daysA, daysB, labelA, labelB }) {
  const [ref, inView] = useInView();
  const [hover, setHover] = useState(null);
  const W = 760, H = 320, PAD = { l: 50, r: 20, t: 20, b: 40 };
  const innerW = W - PAD.l - PAD.r;
  const innerH = H - PAD.t - PAD.b;
  const len = Math.max(daysA.length, daysB.length);
  const allVals = [...daysA.map(d => d.value), ...daysB.map(d => d.value)];
  if (allVals.length === 0) return <div className="empty-state">No daily data</div>;
  const minV = Math.min(...allVals) * 0.95;
  const maxV = Math.max(...allVals) * 1.05;
  const xAt = i => PAD.l + (len > 1 ? (i / (len - 1)) * innerW : innerW / 2);
  const yAt = v => PAD.t + innerH - ((v - minV) / Math.max(1, maxV - minV)) * innerH;
  const path = arr => arr.map((d, i) => `${i === 0 ? "M" : "L"} ${xAt(i)} ${yAt(d.value)}`).join(" ");
  const area = arr => path(arr) + ` L ${xAt(arr.length - 1)} ${PAD.t + innerH} L ${xAt(0)} ${PAD.t + innerH} Z`;
  const ticks = 5;
  const yT = Array.from({ length: ticks + 1 }, (_, i) => Math.round(minV + ((maxV - minV) / ticks) * i));

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", overflow: "visible" }}>
        <defs>
          <linearGradient id="aA" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="rgba(148,163,184,0.4)"/><stop offset="100%" stopColor="rgba(148,163,184,0)"/></linearGradient>
          <linearGradient id="aB" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="rgba(52,211,153,0.5)"/><stop offset="100%" stopColor="rgba(52,211,153,0)"/></linearGradient>
        </defs>
        {yT.map((t, i) => {
          const y = yAt(t);
          return (<g key={i}>
            <line x1={PAD.l} x2={W - PAD.r} y1={y} y2={y} stroke="rgba(0,0,0,0.05)"/>
            <text x={PAD.l - 8} y={y + 4} textAnchor="end" fontSize="10" fill="var(--text-mute)">{fmt(t)}</text>
          </g>);
        })}
        <path d={area(daysA)} fill="url(#aA)" opacity="0">
          <animate attributeName="opacity" from="0" to="1" dur="0.9s" begin={inView ? "0.4s" : "indefinite"} fill="freeze"/>
        </path>
        <path d={area(daysB)} fill="url(#aB)" opacity="0">
          <animate attributeName="opacity" from="0" to="1" dur="0.9s" begin={inView ? "0.6s" : "indefinite"} fill="freeze"/>
        </path>
        <path d={path(daysA)} fill="none" stroke="#94a3b8" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
              strokeDasharray="2000" strokeDashoffset={inView ? "0" : "2000"}
              style={{ transition: "stroke-dashoffset 1.4s ease-out" }} />
        <path d={path(daysB)} fill="none" stroke="#10b981" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
              strokeDasharray="2000" strokeDashoffset={inView ? "0" : "2000"}
              style={{ transition: "stroke-dashoffset 1.4s 0.2s ease-out" }} />
        {daysA.map((d, i) => (
          <circle key={"a"+i} cx={xAt(i)} cy={yAt(d.value)} r="4" fill="#94a3b8" opacity="0"
                  onMouseEnter={() => setHover({ i, type: "A", x: xAt(i), y: yAt(d.value), val: d.value, label: d.day })}
                  onMouseLeave={() => setHover(null)} style={{ cursor: "pointer" }}>
            <animate attributeName="opacity" from="0" to="1" dur="0.3s" begin={inView ? `${1.2 + i*0.05}s` : "indefinite"} fill="freeze"/>
          </circle>
        ))}
        {daysB.map((d, i) => (
          <circle key={"b"+i} cx={xAt(i)} cy={yAt(d.value)} r="5" fill="#10b981" stroke="#fff" strokeWidth="2" opacity="0"
                  onMouseEnter={() => setHover({ i, type: "B", x: xAt(i), y: yAt(d.value), val: d.value, label: d.day })}
                  onMouseLeave={() => setHover(null)} style={{ cursor: "pointer" }}>
            <animate attributeName="opacity" from="0" to="1" dur="0.3s" begin={inView ? `${1.4 + i*0.05}s` : "indefinite"} fill="freeze"/>
          </circle>
        ))}
        {Array.from({ length: len }).map((_, i) => {
          const lbl = (daysB[i] && daysB[i].day) || (daysA[i] && daysA[i].day) || "";
          return <text key={i} x={xAt(i)} y={H - 12} textAnchor="middle" fontSize="11" fill="var(--text-dim)">{lbl}</text>;
        })}
      </svg>
      {hover && (
        <div className="tooltip" style={{ left: `${(hover.x / W) * 100}%`, top: `${(hover.y / H) * 100}%` }}>
          {hover.label}: <strong>{fmt(hover.val)} kWh</strong>
        </div>
      )}
      <div className="legend">
        <span><span className="swatch" style={{ background: "#94a3b8" }}></span>{labelA}</span>
        <span><span className="swatch" style={{ background: "#10b981" }}></span>{labelB}</span>
      </div>
    </div>
  );
}

// =================== DONUT ===================
function Donut({ data, title, sub, delay = 0 }) {
  const [ref, inView] = useInView();
  const [hover, setHover] = useState(null);
  const total = Math.max(1, data.reduce((s, d) => s + d.value, 0));
  const R = 84, CX = 120, CY = 120, C = 2 * Math.PI * R;
  let acc = 0;
  const arcs = data.map((d, i) => { const f = d.value / total; const start = acc; acc += f; return { ...d, f, start, idx: i }; });
  const animVal = useCountUp(total, inView);
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <div className="chart-title">{title}</div>
      <div className="chart-sub" style={{ marginBottom: 14 }}>{sub}</div>
      <div style={{ display: "flex", alignItems: "center", gap: 18, flexWrap: "wrap" }}>
        <svg viewBox="0 0 240 240" width="200" height="200">
          <g transform={`rotate(-90 ${CX} ${CY})`}>
            {arcs.map((a, i) => {
              const len = a.f * C; const offset = a.start * C;
              return (<circle key={i} cx={CX} cy={CY} r={R} fill="none"
                stroke={a.color} strokeWidth={hover === i ? 36 : 28}
                strokeDasharray={`${len} ${C}`}
                strokeDashoffset={inView ? -offset : -C}
                style={{ transition: `stroke-dashoffset 1.1s ${delay + i*0.12}s cubic-bezier(0.34, 1.56, 0.64, 1), stroke-width 0.2s ease`, cursor: "pointer" }}
                onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(null)}/>);
            })}
          </g>
          <text x={CX} y={CY - 4} textAnchor="middle" fontSize="12" fill="var(--text-dim)">Total</text>
          <text x={CX} y={CY + 18} textAnchor="middle" fontSize="20" fontWeight="700" fill="var(--text)">{fmt(animVal)}</text>
          <text x={CX} y={CY + 36} textAnchor="middle" fontSize="10" fill="var(--text-mute)">kWh</text>
        </svg>
        <div style={{ flex: 1, minWidth: 140 }}>
          {arcs.map((a, i) => (
            <div key={i} onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(null)}
                 style={{ display: "flex", justifyContent: "space-between", alignItems: "center",
                   padding: "7px 8px", borderRadius: 6, cursor: "pointer",
                   background: hover === i ? "var(--blue-soft)" : "transparent",
                   transition: "background 0.18s ease" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span className="swatch" style={{ background: a.color }}></span>
                <span style={{ fontSize: 13 }}>{a.name}</span>
              </div>
              <div style={{ fontSize: 13, color: "var(--text-dim)", fontVariantNumeric: "tabular-nums" }}>
                {fmt(a.value)} <span style={{ color: "var(--accent)", marginLeft: 4, fontWeight: 600 }}>{(a.f*100).toFixed(0)}%</span>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// =================== TABLE ===================
function SavingsTable({ data, labelA, labelB }) {
  const [ref, inView] = useInView();
  const totalA = data.reduce((s, o) => s + o.a, 0);
  const totalB = data.reduce((s, o) => s + o.b, 0);
  const totalSaved = totalA - totalB;
  return (
    <div ref={ref} className="table-wrap">
      <table>
        <thead>
          <tr><th>Site</th><th style={{ textAlign: "right" }}>{labelA}</th><th style={{ textAlign: "right" }}>{labelB}</th><th style={{ textAlign: "right" }}>Saved</th><th style={{ textAlign: "right" }}>%</th></tr>
        </thead>
        <tbody>
          {data.map((o, i) => {
            const saved = o.a - o.b; const pct = o.a > 0 ? (saved / o.a) * 100 : 0;
            return (
              <tr key={i} className={"body-row " + (inView ? "in" : "")} style={{ animationDelay: `${i*0.07}s` }}>
                <td>{o.name}</td><td className="num">{fmt(o.a)}</td><td className="num">{fmt(o.b)}</td>
                <td className="num">{fmt(saved)}</td><td className="savings">−{fmt1(pct)}%</td>
              </tr>
            );
          })}
          <tr className={"total body-row " + (inView ? "in" : "")} style={{ animationDelay: `${data.length*0.07}s` }}>
            <td>Total</td><td className="num">{fmt(totalA)}</td><td className="num">{fmt(totalB)}</td>
            <td className="num">{fmt(totalSaved)}</td>
            <td className="savings">−{fmt1(totalA > 0 ? (totalSaved/totalA)*100 : 0)}%</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// =================== UPLOADER ===================
function Uploader({ onParsed }) {
  const [drag, setDrag] = useState(false);
  const [status, setStatus] = useState(null);
  const [busy, setBusy] = useState(false);
  const [fileName, setFileName] = useState(null);
  const inputRef = useRef(null);

  const handle = async (file) => {
    if (!file) return;
    setFileName(file.name);
    setBusy(true); setStatus(null);
    try {
      const buf = await file.arrayBuffer();
      const data = await parseDocx(buf, file.name);
      onParsed(data);
      setStatus({ kind: "success", msg: `Done! Parsed ${data.objects.length} sites, ${Math.max(data.daysA.length, data.daysB.length)} days, ${Math.max(data.loadsA.length, data.loadsB.length)} load types.` });
    } catch (e) {
      setStatus({ kind: "error", msg: e.message });
    } finally { setBusy(false); }
  };

  const onDrop = (e) => {
    e.preventDefault(); setDrag(false);
    const f = e.dataTransfer.files[0]; handle(f);
  };

  return (
    <div>
      <div className={"uploader" + (drag ? " drag" : "") + (status?.kind === "success" ? " success" : "") + (status?.kind === "error" ? " error" : "")}
           onClick={() => inputRef.current?.click()}
           onDragEnter={e => { e.preventDefault(); setDrag(true); }}
           onDragOver={e => { e.preventDefault(); setDrag(true); }}
           onDragLeave={e => { e.preventDefault(); setDrag(false); }}
           onDrop={onDrop}>
        <div className="uploader-icon">
          {busy ? "⏳" : (status?.kind === "success" ? "✓" : (status?.kind === "error" ? "⚠" : "📄"))}
        </div>
        <h3>{busy ? "Processing…" : (drag ? "Release the file here" : "Drop a .docx file or click to choose")}</h3>
        <p>The file should contain tables matching the template format below</p>
        <input ref={inputRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" style={{ display: "none" }}
               onChange={e => handle(e.target.files[0])} />
        <button className="upload-btn" onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}>
          Choose file
        </button>
        <a className="upload-btn secondary" href="/templates/template-campus.docx" download="energy-template-campus.docx"
           onClick={(e) => e.stopPropagation()}>
          Sample: Campus
        </a>
        <a className="upload-btn secondary" href="/templates/template-mall.docx" download="energy-template-mall.docx"
           onClick={(e) => e.stopPropagation()}>
          Sample: Mall
        </a>
        {fileName && <div style={{ marginTop: 14, fontSize: 12, color: "var(--text-dim)" }}>File: {fileName}</div>}
      </div>
      {status && (
        <div className={"toast " + status.kind}>
          {status.kind === "success" ? "✓" : "⚠"} {status.msg}
        </div>
      )}
    </div>
  );
}

// =================== TAB CONTENT ===================
function OverviewTab({ data }) {
  const totalA = data.objects.reduce((s, o) => s + o.a, 0);
  const totalB = data.objects.reduce((s, o) => s + o.b, 0);
  const saved = totalA - totalB;
  const pct = totalA > 0 ? (saved / totalA) * 100 : 0;
  const co2 = Math.round(saved * 0.42);

  const labelA = `${data.monthA} (${data.monthALabel})`;
  const labelB = `${data.monthB} (${data.monthBLabel})`;

  return (<>
    <div className="kpis">
      <KPI idx={0} label="Energy saved" value={saved} unit="kWh" delta={`−${fmt1(pct)}% vs baseline`} icon={Icon.bolt} iconColor="green" />
      <KPI idx={1} label={labelA.length > 28 ? data.monthA : labelA} value={totalA} unit="kWh" sub={data.monthALabel} icon={Icon.chart} iconColor="gray" />
      <KPI idx={2} label={labelB.length > 28 ? data.monthB : labelB} value={totalB} unit="kWh" sub={data.monthBLabel} icon={Icon.trend} iconColor="blue" />
      <KPI idx={3} label="CO₂ reduced" value={co2} unit="kg" sub="≈ a month of driving for several cars" icon={Icon.leaf} iconColor="green" />
    </div>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Comparison by site</div>
          <div className="chart-sub">Electricity use before and after the rollout</div>
        </div>
        <div className="chip-row">
          <span className="chip active">kWh</span>
          <span className="chip">% saved</span>
        </div>
      </div>
      <BarCompare data={data.objects} labelA={data.monthA} labelB={data.monthB} />
    </ObservedCard>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Daily trend</div>
          <div className="chart-sub">Daily consumption across the observation window</div>
        </div>
      </div>
      <LineChart daysA={data.daysA} daysB={data.daysB} labelA={`${data.monthA} (${data.monthALabel})`} labelB={`${data.monthB} (${data.monthBLabel})`} />
    </ObservedCard>

    <div className="grid-2">
      <ObservedCard>
        <Donut data={data.loadsA} title={data.monthA} sub={`Mix · ${data.monthALabel}`} delay={0.1} />
      </ObservedCard>
      <ObservedCard>
        <Donut data={data.loadsB} title={data.monthB} sub={`Mix · ${data.monthBLabel}`} delay={0.1} />
      </ObservedCard>
    </div>
  </>);
}

function HistoryTab({ data }) {
  return (
    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Detailed savings</div>
          <div className="chart-sub">Per-site breakdown with cumulative totals</div>
        </div>
      </div>
      <SavingsTable data={data.objects} labelA={`${data.monthA}, kWh`} labelB={`${data.monthB}, kWh`} />
    </ObservedCard>
  );
}

function ImportTab({ onParsed, data }) {
  return (<>
    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Import data</div>
          <div className="chart-sub">Upload your .docx with calculations in the template format — charts update automatically</div>
        </div>
      </div>
      <Uploader onParsed={onParsed} />
    </ObservedCard>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Template structure</div>
          <div className="chart-sub">The document should contain three sections with tables</div>
        </div>
      </div>
      <ol style={{ paddingLeft: 22, lineHeight: 1.8, color: "var(--text)" }}>
        <li><strong>Monthly summary</strong> — table: Site | Period A | Period B | Saved %</li>
        <li><strong>Daily breakdown</strong> — two subsections (Period A / Period B), each with a table: Day | Consumption</li>
        <li><strong>Load mix</strong> — two subsections, each with a table: Type | Consumption | Share</li>
      </ol>
      <p style={{ marginTop: 14, color: "var(--text-dim)", fontSize: 13 }}>
        The parser classifies tables by shape (numeric columns, date format), not by section names — so headings can vary. Numbers may include spaces, commas or percent signs; everything is normalized. Both English and Russian documents are supported.
      </p>
    </ObservedCard>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Current data</div>
          <div className="chart-sub">What the dashboard is showing right now</div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 24, flexWrap: "wrap", color: "var(--text)" }}>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Site</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.place}</div></div>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Period A</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.monthA}</div></div>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Period B</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.monthB}</div></div>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Sites</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.objects.length}</div></div>
      </div>
    </ObservedCard>
  </>);
}

function SettingsTab() {
  return (
    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Display settings</div>
          <div className="chart-sub">Visualization and calculation parameters</div>
        </div>
      </div>
      <div style={{ display: "grid", gap: 14, color: "var(--text)" }}>
        <div className="live-row"><span>CO₂ factor (kg / kWh)</span><strong>0.42</strong></div>
        <div className="live-row"><span>Number format</span><strong>EN (1,234,567)</strong></div>
        <div className="live-row"><span>Animations</span><strong>Enabled</strong></div>
        <div className="live-row"><span>Analysis range</span><strong>Month over month</strong></div>
      </div>
      <p style={{ marginTop: 18, color: "var(--text-dim)", fontSize: 13 }}>
        These values are baked into the template. To change them, edit the source or upload your own .docx.
      </p>
    </ObservedCard>
  );
}

function ObservedCard({ children }) {
  const [ref, inView] = useInView(0.05);
  return <div ref={ref} className={"card " + (inView ? "in" : "")}>{children}</div>;
}

// =================== RIGHT PANEL ===================
function RightPanel({ data }) {
  const totalB = data.objects.reduce((s, o) => s + o.b, 0);
  const totalA = data.objects.reduce((s, o) => s + o.a, 0);
  const saved = totalA - totalB;
  const date = new Date().toLocaleDateString("en-US", { day: "2-digit", month: "long", year: "numeric" });
  return (
    <aside className="right-panel">
      <div className="status-card">
        <div className="status-place">Site</div>
        <div className="status-name">{data.place}</div>
        <div className="status-temp">{fmt1((saved/Math.max(1,totalA))*100)}<small>%</small></div>
        <div style={{ fontSize: 12, opacity: 0.85, marginTop: 4 }}>saved this period</div>
        <div className="status-meta">
          <div><span style={{ opacity: 0.8 }}>Sites</span><strong>{data.objects.length}</strong></div>
          <div><span style={{ opacity: 0.8 }}>Saved</span><strong>{fmt(saved)} kWh</strong></div>
        </div>
      </div>

      <div className="live-card">
        <h4>Current load</h4>
        {data.objects.slice(0, 5).map((o, i) => {
          const pct = o.a > 0 ? Math.round((o.b / o.a) * 100) : 0;
          return (
            <div key={i} style={{ marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                <span style={{ color: "var(--text)" }}>{o.name}</span>
                <span style={{ color: "var(--text-dim)" }}>{pct}%</span>
              </div>
              <div style={{ height: 6, background: "#f3f4f6", borderRadius: 3, overflow: "hidden" }}>
                <div style={{
                  width: pct + "%", height: "100%",
                  background: "linear-gradient(90deg, var(--accent), #059669)",
                  transition: "width 0.9s cubic-bezier(0.34, 1.56, 0.64, 1)",
                }}/>
              </div>
            </div>
          );
        })}
      </div>

      <div className="live-card">
        <h4>About this period</h4>
        <div className="live-row"><span>Report date</span><strong>{date}</strong></div>
        <div className="live-row"><span>Period A</span><strong>{data.monthA}</strong></div>
        <div className="live-row"><span>Period B</span><strong>{data.monthB}</strong></div>
        <div className="live-row"><span>Source</span><strong>{data._source || "Sample"}</strong></div>
      </div>
    </aside>
  );
}

// =================== APP ===================
function App() {
  const [tab, setTab] = useState("overview");
  const [data, setData] = useState(DEFAULT_DATA);

  const handleParsed = (newData) => {
    setData({ ...newData, _source: "User upload" });
    setTab("overview");
  };

  const tabs = [
    { id: "overview", label: "Overview" },
    { id: "history",  label: "Details" },
    { id: "import",   label: "Import data" },
    { id: "settings", label: "Settings" },
  ];

  const sideItems = [
    { id: "overview", label: "Overview",   icon: Icon.home },
    { id: "history",  label: "Details",    icon: Icon.chart },
    { id: "import",   label: "Import",     icon: Icon.upload, badge: "DOCX" },
    { id: "settings", label: "Settings",   icon: Icon.cog },
  ];

  return (
    <>
      <header className="app-header">
        <div className="app-header-top">
          <div className="brand">
            <div className="brand-logo">⚡</div>
            <div className="brand-text">
              <strong>ENERGY MONITOR</strong>
              <span>Energy consumption monitoring system</span>
            </div>
          </div>
          <div className="header-right">
            <div className="header-pill"><span className="pulse-dot"></span> Connected</div>
            <div className="header-pill">{data.place}</div>
          </div>
        </div>
        <div className="tabs-row">
          {tabs.map(t => (
            <button key={t.id} className={"top-tab " + (tab === t.id ? "active" : "")} onClick={() => setTab(t.id)}>
              {t.label}
            </button>
          ))}
        </div>
      </header>

      <div className="layout">
        <aside className="sidebar">
          <div className="side-section">
            <div className="side-title">Navigation</div>
            {sideItems.map(s => (
              <button key={s.id} className={"side-item " + (tab === s.id ? "active" : "")} onClick={() => setTab(s.id)}>
                {s.icon}<span>{s.label}</span>
                {s.badge && <span className="side-badge">{s.badge}</span>}
              </button>
            ))}
          </div>
          <div className="side-section">
            <div className="side-title">Sites</div>
            <button className="side-item active">{Icon.meter}<span>{data.place}</span></button>
            <button className="side-item">{Icon.plus}<span>Add</span></button>
          </div>
        </aside>

        <main className="main-content" key={tab}>
          <div className="crumbs">{data.place} · {tabs.find(t => t.id === tab)?.label}</div>
          <div className="page-title">
            {tabs.find(t => t.id === tab)?.label}
          </div>
          <div className="page-sub">
            {tab === "overview"  && "Snapshot for the observation period"}
            {tab === "history"   && "Full per-site savings breakdown"}
            {tab === "import"    && "Upload your own data in the template format"}
            {tab === "settings"  && "Calculation and display parameters"}
          </div>

          {tab === "overview"  && <OverviewTab data={data} />}
          {tab === "history"   && <HistoryTab  data={data} />}
          {tab === "import"    && <ImportTab   data={data} onParsed={handleParsed} />}
          {tab === "settings"  && <SettingsTab />}
        </main>

        <RightPanel data={data} />
      </div>
    </>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);

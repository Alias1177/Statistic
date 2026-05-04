const { useState, useEffect, useRef, useMemo, useCallback } = React;

// =================== DEFAULT DATA (from исходный документ) ===================
const DEFAULT_DATA = {
  place: "ПИТИ — кампус",
  monthA: "Март 2026",
  monthB: "Апрель 2026",
  monthALabel: "без системы",
  monthBLabel: "с системой",
  objects: [
    { name: "Учебные корпуса", a: 52000, b: 44500 },
    { name: "Общежития",       a: 38000, b: 34200 },
    { name: "Администрация",   a: 12500, b: 10300 },
    { name: "Лаборатории",     a: 28000, b: 25800 },
    { name: "Освещение терр.", a:  9500, b:  6900 },
  ],
  daysA: [
    { day: "01.03", value: 4600 }, { day: "02.03", value: 4550 },
    { day: "03.03", value: 4700 }, { day: "04.03", value: 4650 },
    { day: "05.03", value: 4800 }, { day: "06.03", value: 4900 },
    { day: "07.03", value: 4750 },
  ],
  daysB: [
    { day: "01.04", value: 4100 }, { day: "02.04", value: 4050 },
    { day: "03.04", value: 4200 }, { day: "04.04", value: 4150 },
    { day: "05.04", value: 4250 }, { day: "06.04", value: 4300 },
    { day: "07.04", value: 4120 },
  ],
  loadsA: [
    { name: "Освещение",   value: 35000, color: "#f59e0b" },
    { name: "HVAC",        value: 45000, color: "#3b82f6" },
    { name: "Оборудование",value: 38000, color: "#8b5cf6" },
    { name: "Прочее",      value: 22000, color: "#6b7280" },
  ],
  loadsB: [
    { name: "Освещение",   value: 25000, color: "#10b981" },
    { name: "HVAC",        value: 42000, color: "#3b82f6" },
    { name: "Оборудование",value: 36000, color: "#8b5cf6" },
    { name: "Прочее",      value: 18700, color: "#6b7280" },
  ],
};

// =================== UTILS ===================
const fmt  = n => Math.round(n).toLocaleString("ru-RU").replace(/,/g, " ");
const fmt1 = n => n.toFixed(1).replace(".", ",");
const parseNum = s => {
  if (typeof s !== "string") return NaN;
  const cleaned = s.replace(/[\s ]/g, "").replace(",", ".").replace(/[^0-9.\-]/g, "");
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

// =================== DOCX PARSER ===================
const PALETTE = ["#f59e0b", "#3b82f6", "#8b5cf6", "#10b981", "#ec4899", "#6b7280"];
const HEADER_WORDS = /^(объект|здание|корпус|день|дата|тип|категория|name|date|итого|total|всего)/i;

function isHeaderRow(row) {
  if (!row || !row[0]) return true;
  if (HEADER_WORDS.test(row[0].trim())) return true;
  return false;
}

function classifyTable(rows) {
  // Find a sample data row that's not header/total
  const dataRows = rows.filter(r => r && r.length > 0 && !isHeaderRow(r));
  if (dataRows.length === 0) return null;
  const s = dataRows[0];

  // Days: first cell looks like DD.MM or DD-MM or DD/MM or YYYY-MM-DD
  if (/^\d{1,2}[.\-/]\d{1,2}([.\-/]\d{2,4})?$|^\d{4}-\d{2}-\d{2}$/.test(s[0].trim())) {
    return { type: "days", rows: dataRows };
  }
  // Loads: 3 cols, last col contains % (доля)
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
    throw new Error("Файл не открывается как архив. Возможно, это не .docx, а PDF, изображение или повреждённый файл.");
  }
  const docFile = zip.file("word/document.xml");
  if (!docFile) {
    // Detect known alternative formats and give a clear, actionable message
    if (zip.file(/^Index\/Document\.iwa$/) || zip.file(/^Index\/Tables\//)) {
      throw new Error("Это файл Apple Pages (iWork), а не Word. В Pages: Файл → Экспортировать в → Word (.docx), и загрузите получившийся файл.");
    }
    if (zip.file("xl/workbook.xml")) {
      throw new Error("Это Excel-файл (.xlsx). Сейчас поддерживается только Word (.docx). Скопируйте таблицы в Word или сохраните как .docx.");
    }
    if (zip.file("ppt/presentation.xml")) {
      throw new Error("Это PowerPoint-презентация. Поддерживается только Word (.docx).");
    }
    if (zip.file("META-INF/manifest.xml")) {
      throw new Error("Это документ OpenDocument (.odt). В LibreOffice: Файл → Сохранить как → Word 2007–365 (.docx).");
    }
    throw new Error("В архиве нет word/document.xml — это не Word-документ. Поддерживается только .docx, сохранённый из Microsoft Word.");
  }
  const xml = await docFile.async("text");
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  const NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const body = dom.getElementsByTagNameNS(NS, "body")[0];
  if (!body) throw new Error("Нет тела документа.");

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
        // Skip phantom merged-row noise (e.g. rows where col count != most common)
        rows.push(cells);
      }
      // Drop rows that are obviously merged-noise (column count > 2x next row)
      const counts = rows.map(r => r.length);
      const norm = counts.sort((a,b)=>a-b)[Math.floor(counts.length/2)] || 0;
      const cleanedRows = rows.filter(r => r.length <= norm * 1.5 || r.length <= 6);
      items.push({ kind: "tbl", rows: cleanedRows.length ? cleanedRows : rows });
    }
  }

  // Detect month labels from any paragraph or first row
  const allText = items.map(it => it.kind === "p" ? it.text : it.rows.map(r => r.join(" ")).join(" ")).join("\n");
  let monthA = "Период A", monthB = "Период B";
  let monthALabel = "без системы", monthBLabel = "с системой";
  const monthRe = /(Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь|Декабрь)\s+\d{4}/gi;
  const months = [...new Set(allText.match(monthRe) || [])];
  if (months[0]) monthA = months[0];
  if (months[1]) monthB = months[1];
  if (/без\s+систем|без\s+датчик/i.test(allText)) monthALabel = "без системы";
  if (/с\s+систем|с\s+датчик/i.test(allText))   monthBLabel = "с системой";
  if (/до\s+модерниз/i.test(allText))           monthALabel = "до модернизации";
  if (/после\s+модерниз/i.test(allText))        monthBLabel = "после модернизации";

  // Extract place name. Strategy:
  // 1) First non-empty paragraph that looks like a title (no leading section number, < 120 chars,
  //    not a section heading like "Сводные данные…", "Детализация…", "Разбивка…")
  //    → take part after " — " or ": " or use whole line.
  // 2) Or any paragraph starting with "Объект:", "Здание:", "Клиент:", etc.
  // 3) Fallback: filename without extension.
  let place = null;
  const sectionWord = /^(сводные|детализац|разбивк|структура|итог|анализ|таблиц|метод)/i;
  const firstParas = items.filter(it => it.kind === "p" && it.text).slice(0, 8);
  // Pattern 2 first: explicit prefix wins regardless of position
  for (const it of firstParas) {
    const m = it.text.match(/^(?:объект|компания|организац[ия]|здание|клиент)[\s:—–-]+(.+)$/i);
    if (m && m[1].trim()) {
      let cand = m[1].trim().replace(/^[«"](.+)[»"]$/, "$1");
      place = cand;
      break;
    }
  }
  // Pattern 1: first paragraph as title
  if (!place && firstParas[0]) {
    const t = firstParas[0].text;
    if (!sectionWord.test(t) && !/^\d+[.)]/.test(t) && t.length < 140) {
      // Take part after " — " (em dash) or " - " separator
      let m = t.match(/[—–]\s*(.+)$/);
      if (m && m[1].trim()) {
        let cand = m[1].trim();
        // Reject if candidate contains unbalanced parens (we caught a dash inside parens)
        const opens = (cand.match(/\(/g) || []).length;
        const closes = (cand.match(/\)/g) || []).length;
        if (closes <= opens) place = cand.replace(/^[«"](.+)[»"]$/, "$1");
      }
      if (!place) {
        // Use the full line if short and looks like a name
        if (t.length < 80) place = t.replace(/^[«"](.+)[»"]$/, "$1");
      }
    }
  }
  // Fallback: filename without extension
  if (!place && fileName) {
    place = fileName.replace(/\.[^.]+$/, "").replace(/[_]+/g, " ").trim();
    // Strip leading "шаблон-N-" prefix if present
    place = place.replace(/^шаблон[\s\-—–]*\d*[\s\-—–]*/i, "").trim() || place;
  }
  if (!place) place = "Загруженный объект";

  // Classify all tables by shape
  const tables = items.filter(it => it.kind === "tbl").map(it => classifyTable(it.rows));
  const objectsTable = tables.find(t => t && t.type === "objects");
  const daysTables   = tables.filter(t => t && t.type === "days");
  const loadsTables  = tables.filter(t => t && t.type === "loads");

  if (!objectsTable) throw new Error("Не нашёл таблицу с объектами (3+ колонки, колонки 2 и 3 — числа).");

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

  if (objects.length === 0) throw new Error("В таблице объектов не нашлось числовых строк.");

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

// =================== ICONS (inline svg) ===================
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

// =================== BAR CHART (compare A vs B) ===================
function BarCompare({ data, labelA, labelB }) {
  const [ref, inView] = useInView();
  const [tip, setTip] = useState(null);
  const containerRef = useRef(null);
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
      <div ref={containerRef} style={{ position: "relative" }}>
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
                  onMouseEnter={() => setTip({ x: cx - barW/2 - 3, y: PAD.t + innerH - hA, text: `${labelA}: ${fmt(d.a)} кВт·ч` })}
                  onMouseLeave={() => setTip(null)}>
              <animate attributeName="height" from="0" to={hA} dur="0.9s" begin={inView ? `${delay}s` : "indefinite"} fill="freeze"/>
              <animate attributeName="y" from={PAD.t + innerH} to={PAD.t + innerH - hA} dur="0.9s" begin={inView ? `${delay}s` : "indefinite"} fill="freeze"/>
            </rect>
            <rect className="bar" x={cx + 3} y={PAD.t + innerH - hB} width={barW} height={hB}
                  fill="url(#gB)" rx="3"
                  onMouseEnter={() => setTip({ x: cx + barW/2 + 3, y: PAD.t + innerH - hB, text: `${labelB}: ${fmt(d.b)} кВт·ч` })}
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
  if (allVals.length === 0) return <div className="empty-state">Нет данных по дням</div>;
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
          {hover.label}: <strong>{fmt(hover.val)} кВт·ч</strong>
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
          <text x={CX} y={CY - 4} textAnchor="middle" fontSize="12" fill="var(--text-dim)">Всего</text>
          <text x={CX} y={CY + 18} textAnchor="middle" fontSize="20" fontWeight="700" fill="var(--text)">{fmt(animVal)}</text>
          <text x={CX} y={CY + 36} textAnchor="middle" fontSize="10" fill="var(--text-mute)">кВт·ч</text>
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
          <tr><th>Объект</th><th style={{ textAlign: "right" }}>{labelA}</th><th style={{ textAlign: "right" }}>{labelB}</th><th style={{ textAlign: "right" }}>Экономия</th><th style={{ textAlign: "right" }}>%</th></tr>
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
            <td>Итого</td><td className="num">{fmt(totalA)}</td><td className="num">{fmt(totalB)}</td>
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
  const [status, setStatus] = useState(null); // {kind, msg}
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
      setStatus({ kind: "success", msg: `Готово! Распознано объектов: ${data.objects.length}, дней: ${Math.max(data.daysA.length, data.daysB.length)}, типов нагрузки: ${Math.max(data.loadsA.length, data.loadsB.length)}.` });
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
        <h3>{busy ? "Обрабатываем…" : (drag ? "Отпустите файл здесь" : "Перетащите .docx или нажмите для выбора")}</h3>
        <p>Файл должен содержать таблицы в том же формате, что и шаблон ниже</p>
        <input ref={inputRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" style={{ display: "none" }}
               onChange={e => handle(e.target.files[0])} />
        <button className="upload-btn" onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}>
          Выбрать файл
        </button>
        <a className="upload-btn secondary" href="/templates/template-piti.docx" download="шаблон-кампус.docx"
           onClick={(e) => e.stopPropagation()}>
          Шаблон: кампус
        </a>
        <a className="upload-btn secondary" href="/templates/template-mall.docx" download="шаблон-ТЦ.docx"
           onClick={(e) => e.stopPropagation()}>
          Шаблон: ТЦ
        </a>
        {fileName && <div style={{ marginTop: 14, fontSize: 12, color: "var(--text-dim)" }}>Файл: {fileName}</div>}
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
      <KPI idx={0} label="Сэкономлено" value={saved} unit="кВт·ч" delta={`−${fmt1(pct)}% к базовому периоду`} icon={Icon.bolt} iconColor="green" />
      <KPI idx={1} label={labelA.length > 28 ? data.monthA : labelA} value={totalA} unit="кВт·ч" sub={data.monthALabel} icon={Icon.chart} iconColor="gray" />
      <KPI idx={2} label={labelB.length > 28 ? data.monthB : labelB} value={totalB} unit="кВт·ч" sub={data.monthBLabel} icon={Icon.trend} iconColor="blue" />
      <KPI idx={3} label="Снижение CO₂" value={co2} unit="кг" sub="≈ месяц езды нескольких авто" icon={Icon.leaf} iconColor="green" />
    </div>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Сравнение по объектам</div>
          <div className="chart-sub">Расход электроэнергии до и после внедрения</div>
        </div>
        <div className="chip-row">
          <span className="chip active">кВт·ч</span>
          <span className="chip">% экономии</span>
        </div>
      </div>
      <BarCompare data={data.objects} labelA={data.monthA} labelB={data.monthB} />
    </ObservedCard>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Динамика по дням</div>
          <div className="chart-sub">Суточное потребление за период наблюдения</div>
        </div>
      </div>
      <LineChart daysA={data.daysA} daysB={data.daysB} labelA={`${data.monthA} (${data.monthALabel})`} labelB={`${data.monthB} (${data.monthBLabel})`} />
    </ObservedCard>

    <div className="grid-2">
      <ObservedCard>
        <Donut data={data.loadsA} title={data.monthA} sub={`Структура · ${data.monthALabel}`} delay={0.1} />
      </ObservedCard>
      <ObservedCard>
        <Donut data={data.loadsB} title={data.monthB} sub={`Структура · ${data.monthBLabel}`} delay={0.1} />
      </ObservedCard>
    </div>
  </>);
}

function HistoryTab({ data }) {
  return (
    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Детализация экономии</div>
          <div className="chart-sub">Разбивка по объектам с накопленной экономией</div>
        </div>
      </div>
      <SavingsTable data={data.objects} labelA={`${data.monthA}, кВт·ч`} labelB={`${data.monthB}, кВт·ч`} />
    </ObservedCard>
  );
}

function ImportTab({ onParsed, data }) {
  return (<>
    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Импорт данных</div>
          <div className="chart-sub">Загрузите свой .docx с расчётами в формате шаблона — графики обновятся автоматически</div>
        </div>
      </div>
      <Uploader onParsed={onParsed} />
    </ObservedCard>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Структура шаблона</div>
          <div className="chart-sub">Документ должен содержать три раздела с таблицами</div>
        </div>
      </div>
      <ol style={{ paddingLeft: 22, lineHeight: 1.8, color: "var(--text)" }}>
        <li><strong>Сводные данные по месяцам</strong> — таблица: Объект | Месяц A | Месяц B | Экономия %</li>
        <li><strong>Детализация по дням</strong> — два подраздела «Март» / «Апрель», в каждом таблица: День | Потребление</li>
        <li><strong>Разбивка по типам нагрузки</strong> — два подраздела, в каждом таблица: Тип | Потребление | Доля</li>
      </ol>
      <p style={{ marginTop: 14, color: "var(--text-dim)", fontSize: 13 }}>
        Парсер ищет ключевые слова «сводные данные», «детализация», «разбивка» и автоматически распределяет таблицы по разделам. Числа могут быть с пробелами/запятыми/процентами — всё нормализуется.
      </p>
    </ObservedCard>

    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Текущие данные</div>
          <div className="chart-sub">Что сейчас отображается в дашборде</div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 24, flexWrap: "wrap", color: "var(--text)" }}>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Объект</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.place}</div></div>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Период A</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.monthA}</div></div>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Период B</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.monthB}</div></div>
        <div><div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Объектов</div><div style={{ fontSize: 16, fontWeight: 600 }}>{data.objects.length}</div></div>
      </div>
    </ObservedCard>
  </>);
}

function SettingsTab() {
  return (
    <ObservedCard>
      <div className="chart-head">
        <div>
          <div className="chart-title">Настройки отображения</div>
          <div className="chart-sub">Параметры визуализации и расчётов</div>
        </div>
      </div>
      <div style={{ display: "grid", gap: 14, color: "var(--text)" }}>
        <div className="live-row"><span>Коэффициент CO₂ (кг/кВт·ч)</span><strong>0,42</strong></div>
        <div className="live-row"><span>Формат чисел</span><strong>RU (1 234 567)</strong></div>
        <div className="live-row"><span>Анимации</span><strong>Включены</strong></div>
        <div className="live-row"><span>Период анализа</span><strong>Месяц-к-месяцу</strong></div>
      </div>
      <p style={{ marginTop: 18, color: "var(--text-dim)", fontSize: 13 }}>
        Эти значения зашиты в шаблон. Чтобы изменить — отредактируйте параметры в коде или загрузите свой .docx.
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
  const date = new Date().toLocaleDateString("ru-RU", { day: "2-digit", month: "long", year: "numeric" });
  return (
    <aside className="right-panel">
      <div className="status-card">
        <div className="status-place">Объект</div>
        <div className="status-name">{data.place}</div>
        <div className="status-temp">{fmt1((saved/Math.max(1,totalA))*100)}<small>%</small></div>
        <div style={{ fontSize: 12, opacity: 0.85, marginTop: 4 }}>экономия за период</div>
        <div className="status-meta">
          <div><span style={{ opacity: 0.8 }}>Объектов</span><strong>{data.objects.length}</strong></div>
          <div><span style={{ opacity: 0.8 }}>Сэкономлено</span><strong>{fmt(saved)} кВт·ч</strong></div>
        </div>
      </div>

      <div className="live-card">
        <h4>Текущая нагрузка</h4>
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
        <h4>О периоде</h4>
        <div className="live-row"><span>Дата отчёта</span><strong>{date}</strong></div>
        <div className="live-row"><span>Период A</span><strong>{data.monthA}</strong></div>
        <div className="live-row"><span>Период B</span><strong>{data.monthB}</strong></div>
        <div className="live-row"><span>Источник</span><strong>{data._source || "Шаблон"}</strong></div>
      </div>
    </aside>
  );
}

// =================== APP ===================
function App() {
  const [tab, setTab] = useState("overview");
  const [data, setData] = useState(DEFAULT_DATA);

  const handleParsed = (newData) => {
    setData({ ...newData, _source: "Загружено пользователем" });
    setTab("overview");
  };

  const tabs = [
    { id: "overview", label: "Обзор" },
    { id: "history", label: "Детализация" },
    { id: "import", label: "Импорт данных" },
    { id: "settings", label: "Настройки" },
  ];

  const sideItems = [
    { id: "overview", label: "Обзор", icon: Icon.home },
    { id: "history", label: "Детализация", icon: Icon.chart },
    { id: "import", label: "Импорт", icon: Icon.upload, badge: "DOCX" },
    { id: "settings", label: "Настройки", icon: Icon.cog },
  ];

  return (
    <>
      <header className="app-header">
        <div className="app-header-top">
          <div className="brand">
            <div className="brand-logo">⚡</div>
            <div className="brand-text">
              <strong>ЭНЕРГОМОНИТОР</strong>
              <span>Система мониторинга энергопотребления</span>
            </div>
          </div>
          <div className="header-right">
            <div className="header-pill"><span className="pulse-dot"></span> Подключено</div>
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
            <div className="side-title">Навигация</div>
            {sideItems.map(s => (
              <button key={s.id} className={"side-item " + (tab === s.id ? "active" : "")} onClick={() => setTab(s.id)}>
                {s.icon}<span>{s.label}</span>
                {s.badge && <span className="side-badge">{s.badge}</span>}
              </button>
            ))}
          </div>
          <div className="side-section">
            <div className="side-title">Объекты</div>
            <button className="side-item active">{Icon.meter}<span>{data.place}</span></button>
            <button className="side-item">{Icon.plus}<span>Добавить</span></button>
          </div>
        </aside>

        <main className="main-content" key={tab}>
          <div className="crumbs">{data.place} · {tabs.find(t => t.id === tab)?.label}</div>
          <div className="page-title">
            {tabs.find(t => t.id === tab)?.label}
          </div>
          <div className="page-sub">
            {tab === "overview"  && "Сводная картина за период наблюдения"}
            {tab === "history"   && "Полная разбивка экономии по объектам"}
            {tab === "import"    && "Загрузка собственных данных в формате шаблона"}
            {tab === "settings"  && "Параметры расчётов и отображения"}
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

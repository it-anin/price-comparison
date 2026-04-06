import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { useEffect } from "react";

// โ”€โ”€โ”€ SAP-Style Theme Constants โ”€โ”€โ”€
const SAP = {
  shell: "#354A5F",
  shellText: "#FFFFFF",
  brand: "#0A6ED1",
  brandHover: "#0854A0",
  positive: "#107E3E",
  negative: "#BB0000",
  warning: "#E9730C",
  neutral: "#6A6D70",
  bg: "#F7F7F7",
  surface: "#FFFFFF",
  border: "#D9D9D9",
  borderDark: "#89919A",
  headerBg: "#F2F2F2",
  rowHover: "#E8F0FE",
  selectedRow: "#D4E5F7",
  text: "#32363A",
  textLight: "#6A6D70",
  fontFamily: "'72', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
};

// โ”€โ”€โ”€ Icon Components โ”€โ”€โ”€
const Icon = ({ d, size = 16, color = "currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d={d} />
  </svg>
);

const Icons = {
  upload: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12",
  download: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3",
  compare: "M16 3h5v5M4 20L21 3M21 16v5h-5M15 15l6 6M4 4l5 5",
  filter: "M22 3H2l8 9.46V19l4 2v-8.54L22 3z",
  check: "M20 6L9 17l-5-5",
  alert: "M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0zM12 9v4M12 17h.01",
  file: "M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8",
  search: "M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z",
  x: "M18 6L6 18M6 6l12 12",
  chevDown: "M6 9l6 6 6-6",
  refresh: "M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15",
  table: "M9 3H5a2 2 0 00-2 2v4m6-6h10a2 2 0 012 2v4M9 3v18m0 0h10a2 2 0 002-2V9M9 21H5a2 2 0 01-2-2V9m0 0h18",
  settings: "M12 15a3 3 0 100-6 3 3 0 000 6zM19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 112.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 114 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 112.83 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z",
};

// โ”€โ”€โ”€ Utility Functions โ”€โ”€โ”€
const fmt = (n) => {
  if (n == null || isNaN(n)) return "-";
  return Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const pct = (n) => {
  if (n == null || isNaN(n)) return "-";
  return (Number(n) * 100).toFixed(2) + "%";
};

const parseNum = (v) => {
  if (v == null) return null;
  const n = typeof v === "string" ? parseFloat(v.replace(/,/g, "")) : Number(v);
  return isNaN(n) ? null : n;
};

const isSameValue = (a, b) => {
  if (a == null && b == null) return true;
  if (a == null || b == null) return false;
  return Math.abs(a - b) <= 0.001;
};

const calcPriceFromCostRatio = (newCost, oldCost, oldPrice) => {
  if (newCost == null || oldCost == null || oldPrice == null) return null;
  if (Math.abs(oldCost) <= 0.0000001) return null;
  return (newCost / oldCost) * oldPrice;
};

const roundPriceLevel = (value) => {
  if (value == null) return null;
  return Math.round(value);
};

const CHECKLIST_STORAGE_KEY = "price-compare-row-checklist";

const px = (v, fallback) => {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseInt(v, 10);
    if (!Number.isNaN(n)) return n;
  }
  return fallback;
};

// โ”€โ”€โ”€ Column letter to 0-based index โ”€โ”€โ”€
const colToIdx = (s) => {
  if (s == null) return -1;
  s = String(s).trim().toUpperCase();
  if (/^\d+$/.test(s)) return parseInt(s);
  let idx = 0;
  for (let i = 0; i < s.length; i++) idx = idx * 26 + s.charCodeAt(i) - 64;
  return idx - 1;
};

// โ”€โ”€โ”€ Hardcoded column definitions โ”€โ”€โ”€
// File 1: เธ—เธธเธเนเธซเธกเน (Product Master)
const PM = {
  SKU:  colToIdx("A"),  // 0
  NAME: colToIdx("B"),  // 1
  BASE_UNIT: colToIdx("C"),  // 2 ▹ BASE_UNIT (หน่วยฐาน)
  LEGACY_COST: colToIdx("C"),
  LEGACY_UNIT: colToIdx("D"),
  UNIT: colToIdx("H"),
  COST: colToIdx("J"),
};

// File 2: เธฃเธฒเธเธฒเธ—เธธเธเน€เธเนเธฒ (Promaxx)
// Row types in Promaxx:
//   Old-cost rows  : ColD = "1", ColF = "4", ColH = เธ—เธธเธเน€เธเนเธฒ value, ColB = SKU
//   Per-unit rows  : ColE = unit, ColF = level (0/1/4/5), ColG = lv0 price, ColH = lv1/lv4/lv5 price
const PR = {
  SKU:   colToIdx("B"),  // 1
  BASEMULTIPLE: colToIdx("D"),  // 3 โ’ "1" = เธซเธเนเธงเธขเน€เธฅเนเธเธชเธธเธ”
  UNIT: colToIdx("E"),  // 4 โ’ unit code for per-unit price rows
  FILTERORLEVEL: colToIdx("F"),  // 5 โ’ "4" = old-cost row; otherwise = level 0/1/5
  LV0_PRICE:  colToIdx("G"),  // 6
  COL_H: colToIdx("H"),  // 7 โ’ old-cost value, or lv1/lv5 price
};

// โ”€โ”€โ”€ SAP Button โ”€โ”€โ”€
const SAPButton = ({ children, variant = "default", icon, onClick, disabled, small, style = {} }) => {
  const base = {
    display: "inline-flex", alignItems: "center", gap: "6px",
    padding: small ? "4px 10px" : "6px 16px",
    fontSize: small ? "12px" : "13px", fontWeight: 600,
    fontFamily: SAP.fontFamily, border: "1px solid",
    borderRadius: "4px", cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.5 : 1, transition: "all 0.15s",
    whiteSpace: "nowrap",
  };
  const variants = {
    default: { background: SAP.surface, borderColor: SAP.border, color: SAP.text },
    primary: { background: SAP.brand, borderColor: SAP.brand, color: "#fff" },
    positive: { background: SAP.positive, borderColor: SAP.positive, color: "#fff" },
    negative: { background: SAP.negative, borderColor: SAP.negative, color: "#fff" },
    ghost: { background: "transparent", borderColor: "transparent", color: SAP.brand },
  };
  return (
    <button style={{ ...base, ...variants[variant], ...style }} onClick={onClick} disabled={disabled}>
      {icon && <Icon d={Icons[icon]} size={small ? 14 : 16} />}
      {children}
    </button>
  );
};

// โ”€โ”€โ”€ SAP Toolbar โ”€โ”€โ”€
const Toolbar = ({ children }) => (
  <div style={{
    display: "flex", alignItems: "center", gap: "8px", padding: "8px 16px",
    background: SAP.surface, borderBottom: `1px solid ${SAP.border}`,
    flexWrap: "wrap",
  }}>
    {children}
  </div>
);

const ToolbarSep = () => (
  <div style={{ width: "1px", height: "28px", background: SAP.border, margin: "0 4px" }} />
);

// โ”€โ”€โ”€ SAP Panel โ”€โ”€โ”€
const Panel = ({ title, children, toolbar, badge, icon, collapsible = false, defaultOpen = true }) => {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div style={{
      background: SAP.surface, border: `1px solid ${SAP.border}`,
      borderRadius: "4px", overflow: "hidden",
    }}>
      <div
        onClick={collapsible ? () => setOpen(!open) : undefined}
        style={{
          display: "flex", alignItems: "center", gap: "8px",
          padding: "10px 16px", background: SAP.headerBg,
          borderBottom: open ? `1px solid ${SAP.border}` : "none",
          cursor: collapsible ? "pointer" : "default",
          userSelect: "none",
        }}
      >
        {collapsible && (
          <span style={{ transform: open ? "rotate(0)" : "rotate(-90deg)", transition: "0.2s", display: "flex" }}>
            <Icon d={Icons.chevDown} size={14} color={SAP.textLight} />
          </span>
        )}
        {icon && <Icon d={Icons[icon]} size={16} color={SAP.brand} />}
        <span style={{ fontWeight: 700, fontSize: "13px", color: SAP.text, fontFamily: SAP.fontFamily }}>{title}</span>
        {badge != null && (
          <span style={{
            background: SAP.brand, color: "#fff", fontSize: "11px", fontWeight: 700,
            padding: "1px 8px", borderRadius: "10px", marginLeft: "4px",
          }}>{badge}</span>
        )}
        <div style={{ flex: 1 }} />
        {toolbar}
      </div>
      {open && <div>{children}</div>}
    </div>
  );
};

// โ”€โ”€โ”€ SAP Status Badge โ”€โ”€โ”€
const StatusBadge = ({ type, children }) => {
  const colors = {
    success: { bg: "#E6F4EA", text: SAP.positive, border: "#B7DFC3" },
    error: { bg: "#FFEAEA", text: SAP.negative, border: "#FFCACA" },
    warning: { bg: "#FFF3E0", text: SAP.warning, border: "#FFD4A8" },
    info: { bg: "#E8F0FE", text: SAP.brand, border: "#B3D4FC" },
    neutral: { bg: SAP.headerBg, text: SAP.textLight, border: SAP.border },
  };
  const c = colors[type] || colors.neutral;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: "4px",
      padding: "2px 8px", fontSize: "11px", fontWeight: 700,
      background: c.bg, color: c.text, border: `1px solid ${c.border}`,
      borderRadius: "4px", fontFamily: SAP.fontFamily,
    }}>
      {type === "success" && <Icon d={Icons.check} size={12} color={c.text} />}
      {type === "error" && <Icon d={Icons.alert} size={12} color={c.text} />}
      {type === "warning" && <Icon d={Icons.alert} size={12} color={c.text} />}
      {children}
    </span>
  );
};

// โ”€โ”€โ”€ File Upload Zone โ”€โ”€โ”€
const FileUploadZone = ({ label, description, onFile, fileName, fileInfo }) => {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);

  const handleFile = (file) => {
    if (file) onFile(file);
  };

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
      onClick={() => inputRef.current?.click()}
      style={{
        border: `2px dashed ${dragOver ? SAP.brand : SAP.border}`,
        borderRadius: "4px", padding: "24px", textAlign: "center",
        cursor: "pointer", transition: "all 0.2s",
        background: dragOver ? "#E8F0FE" : fileName ? "#F0FAF0" : SAP.surface,
        minWidth: "280px", flex: 1,
      }}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
        onChange={(e) => handleFile(e.target.files[0])} />
      <div style={{ marginBottom: "8px" }}>
        <Icon d={fileName ? Icons.check : Icons.upload} size={32} color={fileName ? SAP.positive : SAP.brand} />
      </div>
      <div style={{ fontWeight: 700, fontSize: "14px", color: SAP.text, fontFamily: SAP.fontFamily }}>{label}</div>
      <div style={{ fontSize: "12px", color: SAP.textLight, marginTop: "4px" }}>{description}</div>
      {fileName && (
        <div style={{ marginTop: "8px" }}>
          <StatusBadge type="success">{fileName}</StatusBadge>
          {fileInfo && <div style={{ fontSize: "11px", color: SAP.textLight, marginTop: "4px" }}>{fileInfo}</div>}
        </div>
      )}
    </div>
  );
};

// โ”€โ”€โ”€ Virtual Scrolling Table โ”€โ”€โ”€
const ROW_HEIGHT = 34;
const VirtualTable = ({ columns, data, maxHeight = 500, onRowClick, selectedIdx }) => {
  const containerRef = useRef(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const totalHeight = data.length * ROW_HEIGHT;
  const visibleCount = Math.ceil(maxHeight / ROW_HEIGHT) + 2;
  const startIdx = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - 1);
  const endIdx = Math.min(data.length, startIdx + visibleCount);
  const visibleData = data.slice(startIdx, endIdx);
  const totalMinWidth = columns.reduce((sum, col) => sum + px(col.minWidth, 100), 50);
  const HEADER_ROW_HEIGHT = 34;
  const gridTemplateColumns = ["50px", ...columns.map((col) => `${px(col.minWidth, 100)}px`)].join(" ");
  const headerCells = [{ key: "__rownum", label: "#", row: "1 / span 2", col: 1, align: "center" }];
  let colCursor = 2;
  for (let idx = 0; idx < columns.length; idx++) {
    const col = columns[idx];
    if (col.groupHeader) {
      const start = idx;
      let end = idx;
      while (end + 1 < columns.length && columns[end + 1].groupHeader === col.groupHeader) end += 1;
      if (start === idx) {
        headerCells.push({
          key: `group-${col.groupHeader}-${idx}`,
          label: col.groupHeader,
          row: "1 / 2",
          col: `${colCursor} / span ${end - start + 1}`,
          align: "center",
        });
        for (let sub = start; sub <= end; sub++) {
          headerCells.push({
            key: `sub-${columns[sub].key}`,
            label: columns[sub].header,
            row: "2 / 3",
            col: colCursor + (sub - start),
            align: columns[sub].align || "center",
          });
        }
      }
      colCursor += end - start + 1;
      idx = end;
    } else {
      headerCells.push({
        key: `plain-${col.key}`,
        label: col.header,
        row: "1 / span 2",
        col: colCursor,
        align: col.align || "center",
      });
      colCursor += 1;
    }
  }

  return (
    <div style={{ overflow: "hidden", border: `1px solid ${SAP.border}`, borderRadius: "4px" }}>
      {/* Header */}
      <div style={{ overflow: "hidden", borderBottom: `2px solid ${SAP.borderDark}`, background: SAP.headerBg }}>
        <div style={{
          minWidth: `${totalMinWidth}px`,
          transform: `translateX(-${scrollLeft}px)`,
          display: "grid",
          gridTemplateColumns,
          gridTemplateRows: `${HEADER_ROW_HEIGHT}px ${HEADER_ROW_HEIGHT}px`,
        }}>
          {headerCells.map((cell) => (
            <div key={cell.key} style={{
              gridColumn: cell.col,
              gridRow: cell.row,
              padding: "8px 10px",
              fontSize: "11px",
              fontWeight: 700,
              color: SAP.text,
              fontFamily: SAP.fontFamily,
              textAlign: cell.align === "right" ? "right" : "center",
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis",
              background: SAP.headerBg,
              borderRight: `1px solid ${SAP.border}`,
              borderBottom: `1px solid ${SAP.border}`,
              display: "flex",
              alignItems: "center",
              justifyContent: cell.align === "right" ? "flex-end" : "center",
            }}>
              {cell.label}
            </div>
          ))}
        </div>
      </div>
      {/* Body */}
      <div
        ref={containerRef}
        onScroll={(e) => {
          setScrollTop(e.target.scrollTop);
          setScrollLeft(e.target.scrollLeft);
        }}
        style={{ maxHeight, overflow: "auto", position: "relative" }}
      >
        <div style={{ height: totalHeight, position: "relative", minWidth: `${totalMinWidth}px` }}>
          {visibleData.map((row, vi) => {
            const actualIdx = startIdx + vi;
            const isSelected = actualIdx === selectedIdx;
            return (
              <div
                key={actualIdx}
                onClick={() => onRowClick?.(actualIdx)}
                style={{
                  display: "flex", position: "absolute", top: actualIdx * ROW_HEIGHT,
                  width: `${totalMinWidth}px`, height: ROW_HEIGHT,
                  background: isSelected ? SAP.selectedRow : actualIdx % 2 === 0 ? SAP.surface : "#FAFAFA",
                  borderBottom: `1px solid ${SAP.border}`,
                  cursor: "pointer", transition: "background 0.1s",
                }}
              >
                <div style={{ minWidth: "50px", maxWidth: "50px", padding: "0 6px", fontSize: "11px", color: SAP.textLight, fontFamily: SAP.fontFamily, display: "flex", alignItems: "center", justifyContent: "center", borderRight: `1px solid ${SAP.border}` }}>
                  {actualIdx + 1}
                </div>
                {columns.map((col, ci) => (
                  <div key={ci} style={{
                    minWidth: col.minWidth || "100px",
                    maxWidth: col.minWidth || "100px",
                    padding: "0 10px", fontSize: "12px",
                    fontFamily: SAP.fontFamily,
                    borderRight: ci < columns.length - 1 ? `1px solid ${SAP.border}` : "none",
                    textAlign: col.align || "left",
                    display: "flex", alignItems: "center",
                    overflow: "hidden",
                  }}>
                    <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", width: "100%" }}>
                      {col.render ? col.render(row) : String(row[col.key] ?? "")}
                    </div>
                  </div>
                ))}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

// โ”€โ”€โ”€ MAIN APP โ”€โ”€โ”€
export default function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [file1Data, setFile1Data] = useState(null);
  const [file2Data, setFile2Data] = useState(null);
  const [file1Info, setFile1Info] = useState("");
  const [file2Info, setFile2Info] = useState("");
  const [results, setResults] = useState(null);
  const [search, setSearch] = useState("");
  const [filterMode, setFilterMode] = useState("mismatch");
  const [processing, setProcessing] = useState(false);
  const [tab, setTab] = useState("upload");
  const [exportLog, setExportLog] = useState([]);
  const [selectedRowIds, setSelectedRowIds] = useState([]);
  const [selectedPriceLevel, setSelectedPriceLevel] = useState(0);
  const [checkedRows, setCheckedRows] = useState(() => {
    try {
      const raw = window.localStorage.getItem(CHECKLIST_STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });

  // Parse file helper
  const parseFile = useCallback(async (file, setData, setInfo) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const t0 = performance.now();
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const t1 = performance.now();
        setData(raw);
        setInfo(`${raw.length.toLocaleString()} rows ร— ${(raw[0]?.length || 0)} cols | ${(t1 - t0).toFixed(0)}ms`);
        resolve(raw);
      };
      reader.readAsArrayBuffer(file);
    });
  }, []);

  const handleFile1 = useCallback(async (file) => {
    setFile1(file.name);
    await parseFile(file, setFile1Data, setFile1Info);
  }, [parseFile]);

  const handleFile2 = useCallback(async (file) => {
    setFile2(file.name);
    await parseFile(file, setFile2Data, setFile2Info);
  }, [parseFile]);

  // โ”€โ”€โ”€ Compare Logic โ”€โ”€โ”€
  const runCompare = useCallback(() => {
    if (!file1Data || !file2Data) return;
    setProcessing(true);

    setTimeout(() => {
      try {
        const t0 = performance.now();

        // โ”€โ”€ Build Promaxx maps from File 2 โ”€โ”€
        // f2UnitMap : Map<SKU, Map<unit, {lv0,lv1,lv4,lv5}>>  (per-unit price rows)
        const f2UnitMap = new Map();

        for (let i = 1; i < file2Data.length; i++) {
          const row = file2Data[i];
          if (!row) continue;
          const sku = String(row[PR.SKU] ?? "").trim();
          if (!sku) continue;
          const unit = String(row[PR.UNIT] ?? "").trim();
          const filterOrLevel = String(row[PR.FILTERORLEVEL] ?? "").trim();

          if (unit) {
            const level = filterOrLevel;
            if (!f2UnitMap.has(sku)) f2UnitMap.set(sku, new Map());
            const uMap = f2UnitMap.get(sku);
            if (!uMap.has(unit)) uMap.set(unit, { lv0: null, lv1: null, lv4: null, lv5: null });
            const ud = uMap.get(unit);
            if (level === "0") ud.lv0 = parseNum(row[PR.LV0_PRICE]);
            else if (level === "1") ud.lv1 = parseNum(row[PR.COL_H]);
            else if (level === "4") ud.lv4 = parseNum(row[PR.COL_H]);
            else if (level === "5") ud.lv5 = parseNum(row[PR.COL_H]);
          }
        }

        const output = [];
        for (let i = 1; i < file1Data.length; i++) {
          const row = file1Data[i];
          if (!row) continue;
          const sku = String(row[PM.SKU] ?? "").trim();
          if (!sku) continue;
          const name = String(row[PM.NAME] ?? "");
          const file1Unit = String(row[PM.UNIT] ?? "").trim() || String(row[PM.LEGACY_UNIT] ?? "").trim();
          const baseUnit = String(row[PM.BASE_UNIT] ?? "").trim();
          const file1Cost = parseNum(row[PM.COST]) ?? parseNum(row[PM.LEGACY_COST]);
          const newCost = file1Cost;
          let oldCost = null;

          const matched = f2UnitMap.has(sku);
          const diff = null;
          const costMismatch = false;

          const skuUnitMap = f2UnitMap.get(sku);
          const basePromaxx = skuUnitMap?.get(file1Unit) ?? {};
          const baseOldCost = basePromaxx.lv4 ?? null;
          const unitEntries = skuUnitMap ? Array.from(skuUnitMap.entries()) : [];
          const units = unitEntries.map(([unitName, promaxx]) => {
            const oldCostUnit = promaxx.lv4 ?? null;
            const lv0_old = promaxx.lv0 ?? null;
            const lv1_old = promaxx.lv1 ?? null;
            const lv5_old = promaxx.lv5 ?? null;
            const newCostUnit = calcPriceFromCostRatio(file1Cost, baseOldCost, oldCostUnit);
            const lv0_new = roundPriceLevel(calcPriceFromCostRatio(newCostUnit, oldCostUnit, lv0_old));
            const lv1_new = roundPriceLevel(calcPriceFromCostRatio(newCostUnit, oldCostUnit, lv1_old));
            const lv5_new = roundPriceLevel(calcPriceFromCostRatio(newCostUnit, oldCostUnit, lv5_old));
            return {
              unitName,
              oldCostUnit,
              newCostUnit,
              lv0_new, lv1_new, lv5_new,
              lv0_old,
              lv1_old,
              lv5_old,
            };
          });
          if (!units.length && file1Unit) {
            units.push({
              unitName: file1Unit,
              oldCostUnit: null,
              newCostUnit: file1Cost,
              lv0_new: null,
              lv1_new: null,
              lv5_new: null,
              lv0_old: null,
              lv1_old: null,
              lv5_old: null,
            });
          }
          if (oldCost == null && baseOldCost != null) oldCost = baseOldCost;
          const unitMismatch = units.some((u) =>
            !isSameValue(u.lv0_new, u.lv0_old) ||
            !isSameValue(u.lv1_new, u.lv1_old) ||
            !isSameValue(u.lv5_new, u.lv5_old)
          );
          const isMismatch = costMismatch || unitMismatch;

          output.push({ sku, name, newCost, oldCost, diff, matched, isMismatch, units, baseUnit });
        }

        const t1 = performance.now();
        setResults({
          data: output,
          totalItems: output.length,
          matched: output.filter((r) => r.matched).length,
          mismatched: output.filter((r) => r.isMismatch).length,
          notFound: output.filter((r) => !r.matched).length,
          f2Count: f2UnitMap.size,
          processTime: (t1 - t0).toFixed(0),
        });

        setTab("compare");
      } catch (err) {
        console.error("Compare failed", err);
        window.alert(`Compare failed: ${err?.message || err}`);
      } finally {
        setProcessing(false);
      }
    }, 50);
  }, [file1Data, file2Data]);

  // Filtered SKU data
  const filteredSkus = useMemo(() => {
    if (!results) return [];
    let d = results.data;
    if (filterMode === "mismatch") d = d.filter((r) => r.isMismatch);
    else if (filterMode === "match") d = d.filter((r) => r.matched && !r.isMismatch);
    else if (filterMode === "notfound") d = d.filter((r) => !r.matched);
    return d;
  }, [results, filterMode]);

  const displayRows = useMemo(() => {
    let rows = filteredSkus.flatMap((r) => {
      if (!r.units.length) {
        return [{
          rowId: `${r.sku}__no_unit`,
          sku: r.sku,
          name: r.name,
          unitName: "",
          baseUnit: r.baseUnit,
          oldCost: r.oldCost,
          newCost: r.newCost,
          lv0_new: null,
          lv0_old: null,
          lv1_new: null,
          lv1_old: null,
          lv5_new: null,
          lv5_old: null,
        }];
      }
      return r.units.map((u, idx) => ({
        rowId: `${r.sku}__${u.unitName || "unit"}__${idx}`,
        sku: r.sku,
        name: r.name,
        unitName: u.unitName,
        baseUnit: r.baseUnit,
        oldCost: u.oldCostUnit,
        newCost: u.newCostUnit,
        lv0_new: u.lv0_new,
        lv0_old: u.lv0_old,
        lv1_new: u.lv1_new,
        lv1_old: u.lv1_old,
        lv5_new: u.lv5_new,
        lv5_old: u.lv5_old,
      }));
    });
    if (search) {
      const s = search.toLowerCase();
      rows = rows.filter((row) =>
        row.sku.toLowerCase().includes(s) ||
        row.name.toLowerCase().includes(s) ||
        String(row.unitName ?? "").toLowerCase().includes(s)
      );
    }
    return rows;
  }, [filteredSkus, search]);

  const toggleSelectedRow = useCallback((idx) => {
    const row = displayRows[idx];
    if (!row?.rowId) return;
    setSelectedRowIds((prev) => {
      if (prev.includes(row.rowId)) return prev.filter((id) => id !== row.rowId);
      if (prev.length >= 10) {
        window.alert("เลือกได้สูงสุด 10 แถว");
        return prev;
      }
      return [...prev, row.rowId];
    });
  }, [displayRows]);

  const selectedRows = useMemo(() => {
    const map = new Map(displayRows.map((row) => [row.rowId, row]));
    return selectedRowIds.map((id) => map.get(id)).filter(Boolean);
  }, [displayRows, selectedRowIds]);

  const selectedLevelRows = useMemo(() => {
    return selectedRows.map((row) => {
      const levelMap = {
        0: { oldValue: row.lv0_old, newValue: row.lv0_new, label: "ระดับ 0" },
        1: { oldValue: row.lv1_old, newValue: row.lv1_new, label: "ระดับ 1" },
        4: { oldValue: row.oldCost, newValue: row.newCost, label: "ระดับ 4" },
        5: { oldValue: row.lv5_old, newValue: row.lv5_new, label: "ระดับ 5" },
      };
      return {
        ...row,
        selectedLevelLabel: levelMap[selectedPriceLevel].label,
        selectedLevelOld: levelMap[selectedPriceLevel].oldValue,
        selectedLevelNew: levelMap[selectedPriceLevel].newValue,
      };
    });
  }, [selectedRows, selectedPriceLevel]);

  const exportUnits = useMemo(() => {
    return Array.from(new Set(displayRows.map((row) => row.unitName).filter(Boolean))).sort((a, b) => a.localeCompare(b, "th"));
  }, [displayRows]);

  useEffect(() => {
    try {
      window.localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(checkedRows));
    } catch {
      // Ignore localStorage write failures.
    }
  }, [checkedRows]);

  const toggleChecklistRow = useCallback((rowId) => {
    setCheckedRows((prev) => ({
      ...prev,
      [rowId]: !prev[rowId],
    }));
  }, []);

  // โ”€โ”€โ”€ Export Excel โ”€โ”€โ”€
  const exportExcel = useCallback(() => {
    if (!displayRows.length) return;
    const log = [];
    log.push(`[${new Date().toLocaleTimeString()}] Generating Excel workbook...`);
    setExportLog([...log]);

    const wb = XLSX.utils.book_new();

    // Detail sheet: one row per unit per SKU
    const headers = [
      "รหัสสินค้า", "ชื่อสินค้า", "หน่วย",
      "ทุนเก่า (Promaxx)", "ทุนใหม่ (Product Master)",
      "ราคาระดับ 0 ใหม่", "ราคาระดับ 0 เก่า (Promaxx)",
      "ราคาระดับ 1 ใหม่", "ราคาระดับ 1 เก่า (Promaxx)",
      "ราคาระดับ 5 ใหม่", "ราคาระดับ 5 เก่า (Promaxx)",
    ];
    const rows = [headers];
    displayRows.forEach((r) => {
      rows.push([
        r.sku,
        r.name,
        r.unitName,
        r.oldCost ?? "",
        r.newCost ?? "",
        r.lv0_new ?? "", r.lv0_old ?? "",
        r.lv1_new ?? "", r.lv1_old ?? "",
        r.lv5_new ?? "", r.lv5_old ?? "",
      ]);
    });

    const wsDetail = XLSX.utils.aoa_to_sheet(rows);
    wsDetail["!cols"] = [
      { wch: 16 }, { wch: 35 }, { wch: 16 }, { wch: 16 }, { wch: 16 },
      { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 18 },
    ];
    XLSX.utils.book_append_sheet(wb, wsDetail, "Price Comparison");

    log.push(`[${new Date().toLocaleTimeString()}] ${rows.length - 1} rows prepared`);
    log.push(`[${new Date().toLocaleTimeString()}] Writing XLSX...`);
    setExportLog([...log]);

    XLSX.writeFile(wb, "Price_Comparison_Report.xlsx");

    log.push(`[${new Date().toLocaleTimeString()}] โ… Excel exported successfully`);
    setExportLog([...log]);
  }, [displayRows, results]);

  // โ”€โ”€โ”€ Export Text for POS โ”€โ”€โ”€
  const exportTextByLevel = useCallback((level) => {
    if (!displayRows.length) return;
    const log = [...exportLog];
    log.push(`[${new Date().toLocaleTimeString()}] Generating level ${level} text file...`);
    setExportLog([...log]);

    const priceAccessor = {
      0: (r) => r.lv0_new,
      1: (r) => r.lv1_new,
      4: (r) => r.newCost,
      5: (r) => r.lv5_new,
    }[level];

    const lines = displayRows
      .filter((r) => r.newCost != null && r.oldCost != null && r.newCost > r.oldCost)
      .map((r) => {
        const price = priceAccessor?.(r);
        if (price == null) return null;
        return [r.sku, fmt(price), String(level)].join("\t");
      })
      .filter(Boolean);

    if (!lines.length) {
      log.push(`[${new Date().toLocaleTimeString()}] No rows found for level ${level} where new cost is higher than old cost`);
      setExportLog([...log]);
      window.alert(
        `ไม่พบข้อมูลสำหรับ export ราคาระดับ ${level}\n\n` +
        `เงื่อนไขที่ระบบใช้คือ:\n` +
        `1. ทุนใหม่ ต้องมากกว่า ทุนเก่า\n` +
        `2. ต้องมีราคาใหม่ของระดับ ${level}\n\n` +
        `ถ้าเป็นระดับ 5 มักเกิดจากยังไม่มีค่า "ราคาใหม่" ของระดับ 5 ในผลลัพธ์`
      );
      return;
    }

    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `ราคาระดับ ${level}.txt`;
    a.click();
    URL.revokeObjectURL(url);

    log.push(`[${new Date().toLocaleTimeString()}] Level ${level} text file exported (${lines.length} rows)`);
    setExportLog([...log]);
    window.alert(`Export ราคาระดับ ${level} สำเร็จ\nจำนวน ${lines.length} รายการ`);
  }, [displayRows, exportLog]);

  const exportTextByUnitAndLevel = useCallback((unitName, level) => {
    if (!displayRows.length) return;
    const log = [...exportLog];
    log.push(`[${new Date().toLocaleTimeString()}] Generating ${unitName} level ${level} text file...`);
    setExportLog([...log]);

    const priceAccessor = {
      0: (r) => r.lv0_new,
      1: (r) => r.lv1_new,
      4: (r) => r.newCost,
      5: (r) => r.lv5_new,
    }[level];

    const lines = displayRows
      .filter((r) => r.unitName === unitName && r.newCost != null && r.oldCost != null && r.newCost > r.oldCost)
      .map((r) => {
        const price = priceAccessor?.(r);
        if (price == null) return null;
        return [r.sku, fmt(price), String(level)].join("\t");
      })
      .filter(Boolean);

    if (!lines.length) {
      log.push(`[${new Date().toLocaleTimeString()}] No rows found for unit ${unitName} level ${level}`);
      setExportLog([...log]);
      window.alert(
        `ไม่พบข้อมูลสำหรับ export\nหน่วย: ${unitName}\nระดับราคา: ${level}\n\n` +
        `เงื่อนไขที่ระบบใช้คือ:\n` +
        `1. หน่วยต้องตรงกับ "${unitName}"\n` +
        `2. ทุนใหม่ ต้องมากกว่า ทุนเก่า\n` +
        `3. ต้องมีราคาใหม่ของระดับ ${level}`
      );
      return;
    }

    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${unitName}_ราคาระดับ ${level}.txt`;
    a.click();
    URL.revokeObjectURL(url);

    log.push(`[${new Date().toLocaleTimeString()}] ${unitName} level ${level} text file exported (${lines.length} rows)`);
    setExportLog([...log]);
    window.alert(`Export สำเร็จ\nหน่วย: ${unitName}\nระดับราคา: ${level}\nจำนวน ${lines.length} รายการ`);
  }, [displayRows, exportLog]);


  // --- Export BASE_UNIT ---
  const exportBaseUnit = useCallback((level) => {
    if (!displayRows.length) return;
    const log = [...exportLog];
    log.push(`[${new Date().toLocaleTimeString()}] Generating BASE_UNIT level ${level} text file...`);
    setExportLog([...log]);

    const priceAccessor = {
      0: (r) => r.lv0_new,
      1: (r) => r.lv1_new,
      4: (r) => r.newCost,
      5: (r) => r.lv5_new,
    }[level];

    const lines = displayRows
      .filter((r) => r.newCost != null && r.oldCost != null && r.newCost > r.oldCost && r.baseUnit && r.unitName === r.baseUnit)
      .map((r) => {
        const price = priceAccessor?.(r);
        if (price == null) return null;
        return [r.sku, fmt(price), String(level)].join("	");
      })
      .filter(Boolean);

    if (!lines.length) {
      log.push(`[${new Date().toLocaleTimeString()}] No BASE_UNIT rows found for level ${level}`);
      setExportLog([...log]);
      window.alert(
        `ไม่พบข้อมูล BASE_UNIT สำหรับ export ราคาระดับ ${level}

เงื่อนไข:
` +
        `1. หน่วย ColC ไฟล์ที่1 ต้องตรงกับหน่วย ColE ไฟล์ที่2
` +
        `2. ทุนใหม่ > ทุนเก่า
` +
        `3. ต้องมีราคาระดับ ${level}`
      );
      return;
    }

    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `BASE_UNIT_ราคาระดับ ${level}.txt`;
    a.click();
    URL.revokeObjectURL(url);

    log.push(`[${new Date().toLocaleTimeString()}] BASE_UNIT level ${level} exported (${lines.length} rows)`);
    setExportLog([...log]);
    window.alert(`Export BASE_UNIT ราคาระดับ ${level} สำเร็จ
จำนวน ${lines.length} รายการ`);
  }, [displayRows, exportLog]);
  // โ”€โ”€โ”€ Table columns (main list) โ”€โ”€โ”€
  const tableColumns = useMemo(() => {
    const renderCompareValue = (value, compareTo) => {
      const same = isSameValue(value, compareTo);
      return (
        <span style={{ fontFamily: "monospace", color: same ? SAP.text : SAP.negative, fontWeight: same ? 400 : 700 }}>
          {fmt(value)}
        </span>
      );
    };

    return [
      { key: "sku", header: "SKU / รหัสสินค้า", minWidth: "130px", flex: 1.1 },
      { key: "name", header: "ชื่อสินค้า", minWidth: "260px", flex: 2.3 },
      { key: "unitName", header: "หน่วย", minWidth: "110px", flex: 0.9 },
      {
        key: "oldCost", header: "ทุนเก่า (Promaxx)", groupHeader: "ราคาระดับ 4 (ราคาทุน)", minWidth: "130px", flex: 1, align: "right",
        render: (r) => <span style={{ fontFamily: "monospace" }}>{fmt(r.oldCost)}</span>,
      },
      {
        key: "newCost", header: "ทุนใหม่ (Product Master)", groupHeader: "ราคาระดับ 4 (ราคาทุน)", minWidth: "150px", flex: 1.1, align: "right",
        render: (r) => renderCompareValue(r.newCost, r.oldCost),
      },
      {
        key: "lv0_old", header: "ราคาเก่า", groupHeader: "ราคาระดับ 0", minWidth: "100px", flex: 0.8, align: "right",
        render: (r) => renderCompareValue(r.lv0_old, r.lv0_new),
      },
      {
        key: "lv0_new", header: "ราคาใหม่", groupHeader: "ราคาระดับ 0", minWidth: "100px", flex: 0.8, align: "right",
        render: (r) => renderCompareValue(r.lv0_new, r.lv0_old),
      },
      {
        key: "lv1_old", header: "ราคาเก่า", groupHeader: "ราคาระดับ 1", minWidth: "100px", flex: 0.8, align: "right",
        render: (r) => renderCompareValue(r.lv1_old, r.lv1_new),
      },
      {
        key: "lv1_new", header: "ราคาใหม่", groupHeader: "ราคาระดับ 1", minWidth: "100px", flex: 0.8, align: "right",
        render: (r) => renderCompareValue(r.lv1_new, r.lv1_old),
      },
      {
        key: "lv5_old", header: "ราคาเก่า", groupHeader: "ราคาระดับ 5", minWidth: "100px", flex: 0.8, align: "right",
        render: (r) => renderCompareValue(r.lv5_old, r.lv5_new),
      },
      {
        key: "lv5_new", header: "ราคาใหม่", groupHeader: "ราคาระดับ 5", minWidth: "100px", flex: 0.8, align: "right",
        render: (r) => renderCompareValue(r.lv5_new, r.lv5_old),
      },
      {
        key: "checklist", header: "Checklist", minWidth: "90px", flex: 0.7, align: "center",
        render: (r) => (
          <input
            type="checkbox"
            checked={!!checkedRows[r.rowId]}
            onChange={() => toggleChecklistRow(r.rowId)}
            onClick={(e) => e.stopPropagation()}
          />
        ),
      },
    ];
  }, [checkedRows, toggleChecklistRow]);

  // โ”€โ”€โ”€ RENDER โ”€โ”€โ”€
  return (
    <div style={{ fontFamily: SAP.fontFamily, background: SAP.bg, minHeight: "100vh", color: SAP.text }}>
      {/* Shell Bar */}
      <div style={{
        background: SAP.shell, color: SAP.shellText, padding: "0 20px",
        display: "flex", alignItems: "center", height: "44px",
        boxShadow: "0 2px 4px rgba(0,0,0,0.15)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <div style={{
            width: "28px", height: "28px", background: SAP.brand, borderRadius: "4px",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontWeight: 900, fontSize: "13px",
          }}>PC</div>
          <span style={{ fontWeight: 700, fontSize: "14px", letterSpacing: "0.5px" }}>
            Price Comparison Module
          </span>
        </div>
        <div style={{ flex: 1 }} />
        <div style={{ fontSize: "11px", opacity: 0.7 }}>
          Procurement โ€ข Price Management โ€ข v3.0
        </div>
      </div>

      {/* Tab Bar */}
      <div style={{
        display: "flex", background: SAP.surface, borderBottom: `1px solid ${SAP.border}`,
        padding: "0 16px",
      }}>
        {[
          { id: "upload", label: "Data Upload" },
          { id: "compare", label: "Compare & Analyze", disabled: !results },
          { id: "export", label: "Export", disabled: !results },
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => !t.disabled && setTab(t.id)}
            disabled={t.disabled}
            style={{
              padding: "10px 20px", fontSize: "13px", fontWeight: 600,
              fontFamily: SAP.fontFamily, background: "none",
              border: "none", borderBottom: tab === t.id ? `3px solid ${SAP.brand}` : "3px solid transparent",
              color: tab === t.id ? SAP.brand : t.disabled ? SAP.border : SAP.text,
              cursor: t.disabled ? "not-allowed" : "pointer",
              transition: "all 0.2s",
            }}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div style={{ padding: "16px", maxWidth: "1400px", margin: "0 auto" }}>

        {/* โ•โ•โ• UPLOAD TAB โ•โ•โ• */}
        {tab === "upload" && (
          <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
            {/* Info Bar */}
            <div style={{
              background: "#E8F0FE", border: `1px solid #B3D4FC`, borderRadius: "4px",
              padding: "12px 16px", display: "flex", alignItems: "flex-start", gap: "10px",
            }}>
              <Icon d={Icons.alert} size={18} color={SAP.brand} />
              <div style={{ fontSize: "12px", color: SAP.text, lineHeight: "1.8" }}>
                <b>วิธีใช้:</b> อัปโหลดไฟล์ทั้ง 2 ไฟล์ → กดเปรียบเทียบ → ดูผลลัพธ์ราคาทุกหน่วย → ดาวน์โหลด Excel/Text<br />
                <b>ทุนใหม่ (Product Master):</b> ใช้ SKU=ColA, ชื่อ=ColB, หน่วยอ้างอิง=ColD, หน่วยขาย=ColF/ColQ/ColAB/ColAM/ColAX/ColBI, ทุนต่อหน่วย=ColL/ColW/ColAH/ColAS/ColBD/ColBO และระดับราคา 0/1/5 ตามคอลัมน์ที่กำหนด<br />
                <b>ราคาทุนเก่า (Promaxx):</b> ใช้ SKU=ColB และ match หน่วยจาก Product Master กับ ColE; ถ้า ColF="4" จะดึงทุนเก่าจาก ColH และใช้ ColF=ระดับราคา (0/1/5), ColG/H สำหรับราคาเก่า
              </div>
            </div>

            <Panel title="File Upload" icon="upload">
              <div style={{ padding: "20px", display: "flex", gap: "20px", flexWrap: "wrap" }}>
                <FileUploadZone
                  label="ทุนใหม่ (Product Master)"
                  description="ไฟล์ข้อมูลราคาทุกหน่วย - ใช้ SKU(A), PURCHASE_UNIT(D), หน่วย(F/Q/AB/AM/AX/BI), ทุนต่อหน่วย(L/W/AH/AS/BD/BO) และราคาทุกระดับ"
                  onFile={handleFile1}
                  fileName={file1}
                  fileInfo={file1Info}
                />
                <FileUploadZone
                  label="ราคาทุนเก่า (Promaxx)"
                  description="ไฟล์อ้างอิง - ใช้ SKU(B), match หน่วยกับ ColE, ดึงทุนเก่าระดับ 4 จาก ColH และใช้ ColF เป็นระดับราคา"
                  onFile={handleFile2}
                  fileName={file2}
                  fileInfo={file2Info}
                />
              </div>
            </Panel>

            {/* Action */}
            <div style={{ display: "flex", justifyContent: "center", padding: "8px" }}>
              <SAPButton
                variant="primary"
                icon="compare"
                onClick={runCompare}
                disabled={!file1Data || !file2Data || processing}
                style={{ padding: "10px 32px", fontSize: "14px" }}
              >
                {processing ? "Processing..." : "เปรียบเทียบราคา (Compare Prices)"}
              </SAPButton>
            </div>
          </div>
        )}

        {/* โ•โ•โ• COMPARE TAB โ•โ•โ• */}
        {tab === "compare" && results && (
          <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
            {/* KPI Cards */}
            <div style={{ display: "flex", gap: "12px", flexWrap: "wrap" }}>
              {[
                { label: "Total Items", value: results.totalItems.toLocaleString(), color: SAP.text, bg: SAP.surface },
                { label: "Matched", value: (results.matched - results.mismatched).toLocaleString(), color: SAP.positive, bg: "#E6F4EA" },
                { label: "Mismatched", value: results.mismatched.toLocaleString(), color: SAP.negative, bg: "#FFEAEA" },
                { label: "Not Found", value: results.notFound.toLocaleString(), color: SAP.warning, bg: "#FFF3E0" },
                { label: "Promaxx Records", value: results.f2Count.toLocaleString(), color: SAP.brand, bg: "#E8F0FE" },
                { label: "Process Time", value: `${results.processTime}ms`, color: SAP.textLight, bg: SAP.headerBg },
              ].map((kpi, i) => (
                <div key={i} style={{
                  flex: 1, minWidth: "140px", padding: "12px 16px",
                  background: kpi.bg, borderRadius: "4px",
                  border: `1px solid ${SAP.border}`,
                }}>
                  <div style={{ fontSize: "11px", color: SAP.textLight, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px" }}>
                    {kpi.label}
                  </div>
                  <div style={{ fontSize: "24px", fontWeight: 700, color: kpi.color, fontFamily: "monospace", marginTop: "2px" }}>
                    {kpi.value}
                  </div>
                </div>
              ))}
            </div>

            {/* Toolbar */}
            <Toolbar>
              <div style={{ display: "flex", alignItems: "center", gap: "6px", background: SAP.headerBg, borderRadius: "4px", padding: "2px", border: `1px solid ${SAP.border}` }}>
                {[
                  { id: "mismatch", label: `ไม่ตรง (${results.mismatched})`, type: "error" },
                  { id: "all", label: `ทั้งหมด (${results.totalItems})`, type: "neutral" },
                  { id: "match", label: `ตรง (${results.matched - results.mismatched})`, type: "success" },
                  { id: "notfound", label: `ไม่พบ (${results.notFound})`, type: "warning" },
                ].map((f) => (
                  <button
                    key={f.id}
                    onClick={() => setFilterMode(f.id)}
                    style={{
                      padding: "4px 12px", fontSize: "12px", fontWeight: 600,
                      fontFamily: SAP.fontFamily, border: "none",
                      borderRadius: "3px", cursor: "pointer",
                      background: filterMode === f.id ? SAP.brand : "transparent",
                      color: filterMode === f.id ? "#fff" : SAP.text,
                      transition: "all 0.15s",
                    }}
                  >
                    {f.label}
                  </button>
                ))}
              </div>
              <ToolbarSep />
              <div style={{ display: "flex", alignItems: "center", gap: "6px", border: `1px solid ${SAP.border}`, borderRadius: "4px", padding: "4px 8px", background: SAP.surface }}>
                <Icon d={Icons.search} size={14} color={SAP.textLight} />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Search SKU or Product Name..."
                  style={{
                    border: "none", outline: "none", fontSize: "12px",
                    fontFamily: SAP.fontFamily, width: "200px", background: "transparent", color: SAP.text,
                  }}
                />
                {search && (
                  <button onClick={() => setSearch("")} style={{ background: "none", border: "none", cursor: "pointer", padding: "2px" }}>
                    <Icon d={Icons.x} size={12} color={SAP.textLight} />
                  </button>
                )}
              </div>
              <div style={{ flex: 1 }} />
              <span style={{ fontSize: "12px", color: SAP.textLight }}>
                Showing {displayRows.length.toLocaleString()} rows
              </span>
            </Toolbar>

            {/* Data Table */}
            <Panel title="Comparison Results" icon="table" badge={displayRows.length}>
              <VirtualTable
                columns={tableColumns}
                data={displayRows}
                maxHeight={420}
                selectedIdx={-1}
                onRowClick={toggleSelectedRow}
              />
            </Panel>

            <Panel title="Selected Results" icon="table" badge={selectedRows.length}>
              <div style={{ padding: "12px", display: "flex", flexDirection: "column", gap: "12px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                  <span style={{ fontSize: "12px", color: SAP.textLight }}>ระดับราคาที่แสดง:</span>
                  {[0, 1, 4, 5].map((level) => (
                    <button
                      key={level}
                      onClick={() => setSelectedPriceLevel(level)}
                      style={{
                        padding: "4px 10px",
                        fontSize: "12px",
                        fontWeight: 600,
                        fontFamily: SAP.fontFamily,
                        border: `1px solid ${selectedPriceLevel === level ? SAP.brand : SAP.border}`,
                        borderRadius: "4px",
                        background: selectedPriceLevel === level ? SAP.brand : SAP.surface,
                        color: selectedPriceLevel === level ? "#fff" : SAP.text,
                        cursor: "pointer",
                      }}
                    >
                      {level}
                    </button>
                  ))}
                  <span style={{ fontSize: "12px", color: SAP.textLight }}>
                    เลือกได้สูงสุด 10 แถว
                  </span>
                </div>

                <div style={{ overflowX: "auto", border: `1px solid ${SAP.border}`, borderRadius: "4px" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px", fontFamily: SAP.fontFamily }}>
                    <thead>
                      <tr style={{ background: SAP.headerBg }}>
                        {["SKU / รหัสสินค้า", "ชื่อสินค้า", "หน่วย", `${selectedLevelRows[0]?.selectedLevelLabel || `ระดับ ${selectedPriceLevel}`} (ทุนเก่า/ทุนใหม่)`].map((h) => (
                          <th key={h} style={{ padding: "8px 10px", textAlign: "left", borderBottom: `1px solid ${SAP.border}`, whiteSpace: "nowrap" }}>
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {selectedLevelRows.length ? selectedLevelRows.map((row, idx) => (
                        <tr key={row.rowId} style={{ background: idx % 2 === 0 ? SAP.surface : "#FAFAFA", borderBottom: `1px solid ${SAP.border}` }}>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{row.sku}</td>
                          <td style={{ padding: "8px 10px" }}>{row.name}</td>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{row.unitName || "-"}</td>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap", fontFamily: "monospace" }}>
                            {fmt(row.selectedLevelOld)} / {fmt(row.selectedLevelNew)}
                          </td>
                        </tr>
                      )) : (
                        <tr>
                          <td colSpan={4} style={{ padding: "12px", color: SAP.textLight }}>
                            คลิกเลือกแถวจากตารางด้านบนเพื่อเพิ่มลงในหน้าต่างนี้
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </Panel>
          </div>
        )}

        {/* โ•โ•โ• EXPORT TAB โ•โ•โ• */}
        {tab === "export" && results && (
          <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
            <Panel title="Export Options" icon="download">
              <div style={{ padding: "20px", display: "flex", gap: "20px", flexWrap: "wrap" }}>
                {/* Text Export */}
                <div style={{
                  width: "100%", padding: "24px",
                  border: `2px solid ${SAP.brand}`, borderRadius: "6px",
                  background: "#F8FAFF", textAlign: "center",
                }}>
                  <div style={{ fontWeight: 700,fontSize: "20px", marginBottom: "8px" }}>EXPORT</div>
                  <div style={{ fontWeight: 700, fontSize: "16px", marginBottom: "4px" }}>All_UNIT Export</div>
                  <div style={{ fontSize: "15px", color: SAP.textLight, marginBottom: "16px" }}>
                    • Export ทุกแถว ที่ราคาใหม่ &gt; ราคาเก่า (ไม่กรองหน่วย) •
                  </div>
                  <div style={{ fontSize: "12px", color: SAP.text, marginBottom: "12px" }}>
                    <b>{displayRows.length.toLocaleString()}</b> rows
                  </div>
                  <div style={{ display: "flex", gap: "8px", justifyContent: "center", flexWrap: "nowrap", marginBottom: "12px", overflowX: "auto" }}>
                    {[0, 1, 4, 5].map((level) => (
                      <SAPButton key={level} variant="primary" icon="download" onClick={() => exportTextByLevel(level)}>
                        {`ราคาระดับ ${level}`}
                      </SAPButton>
                    ))}
                  </div>
                  <div style={{ borderTop: `1px solid ${SAP.border}`, paddingTop: "12px", marginBottom: "12px", display: "flex", flexDirection: "column", gap: "6px" }}>
                    <div style={{ fontSize: "15px", fontWeight: 700, color: SAP.positive }}>BASE_UNIT Export</div>
                    <div style={{ fontSize: "15px", color: SAP.textLight, marginBottom: "4px" }}>
                      • Export แยกหน่วย ที่ราคาใหม่ &gt; ราคาเก่า •
                    </div>
                    <div style={{ display: "flex", gap: "8px", justifyContent: "center", flexWrap: "nowrap", overflowX: "auto" }}>
                      {[0, 1, 4, 5].map((level) => (
                        <SAPButton key={`base-${level}`} variant="positive" icon="download" onClick={() => exportBaseUnit(level)}>
                          {`BASE_UNIT ระดับ ${level}`}
                        </SAPButton>
                      ))}
                    </div>
                  </div>
                  <div style={{ borderTop: `1px solid ${SAP.border}`, paddingTop: "12px", marginBottom: "12px", display: "flex", flexDirection: "column", gap: "6px" }}>
                    <div style={{ fontSize: "15px", fontWeight: 700, color: SAP.warning }}>Export Comparison Results</div>
                    <div style={{ fontSize: "15px", color: SAP.textLight, marginBottom: "4px" }}>
                      • Export ตารางทั้งหมดทุกระดับราคาและหน่วย เป็นไฟล์ Excel •
                    </div>
                    <div style={{ display: "flex", gap: "8px", justifyContent: "center" }}>
                      <SAPButton variant="default" icon="download" onClick={exportExcel}>
                        Export All (.xlsx)
                      </SAPButton>
                    </div>
                  </div>
                  <div style={{ borderTop: `1px solid ${SAP.border}`, paddingTop: "12px", display: "flex", flexDirection: "column", gap: "10px" }}>
                    <div style={{ fontSize: "12px", color: SAP.textLight }}>Export แยกตามหน่วย + ระดับราคา</div>
                    {exportUnits.length ? exportUnits.map((unit) => (
                      <div key={unit} style={{ display: "flex", alignItems: "center", gap: "8px", justifyContent: "center", flexWrap: "nowrap", overflowX: "auto" }}>
                        <span style={{ minWidth: "70px", fontSize: "12px", fontWeight: 700, color: SAP.text }}>{unit}</span>
                        {[0, 1, 4, 5].map((level) => (
                          <SAPButton key={`${unit}-${level}`} variant="default" icon="download" onClick={() => exportTextByUnitAndLevel(unit, level)}>
                            {`ระดับ ${level}`}
                          </SAPButton>
                        ))}
                      </div>
                    )) : (
                      <div style={{ fontSize: "12px", color: SAP.textLight }}>ยังไม่พบข้อมูลหน่วยสำหรับ export</div>
                    )}
                  </div>
                </div>
              </div>
            </Panel>

            {/* Export Log */}
            {exportLog.length > 0 && (
              <Panel title="Export Log" icon="file" collapsible>
                <div style={{
                  padding: "12px 16px", background: "#1C2833", borderRadius: "0 0 4px 4px",
                  fontFamily: "'Courier New', monospace", fontSize: "12px", color: "#33FF33",
                  maxHeight: "200px", overflowY: "auto",
                }}>
                  {exportLog.map((line, i) => (
                    <div key={i} style={{ padding: "2px 0" }}>{line}</div>
                  ))}
                  <div style={{ animation: "blink 1s infinite" }}>|</div>
                </div>
              </Panel>
            )}

            {/* Preview */}
            <Panel title="Export Preview (First 10 Rows)" icon="table" collapsible defaultOpen={false}>
              <div style={{ padding: "12px", overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11px", fontFamily: "monospace" }}>
                  <thead>
                    <tr style={{ background: SAP.headerBg }}>
                      {["SKU", "ชื่อสินค้า", "หน่วย", "ทุนเก่า", "ทุนใหม่", "Lv0 ใหม่", "Lv0 เก่า", "Lv1 ใหม่", "Lv1 เก่า", "Lv5 ใหม่", "Lv5 เก่า"].map((h, i) => (
                        <th key={i} style={{ padding: "6px 8px", textAlign: i < 2 ? "left" : "right", borderBottom: `2px solid ${SAP.borderDark}`, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {displayRows.slice(0, 10).map((r, ri) => (
                      <tr key={`${r.sku}-${r.unitName}-${ri}`} style={{ borderBottom: `1px solid ${SAP.border}`, background: ri % 2 === 0 ? SAP.surface : "#FAFAFA" }}>
                        <td style={{ padding: "4px 8px" }}>{r.sku}</td>
                        <td style={{ padding: "4px 8px", maxWidth: "160px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.name}</td>
                        <td style={{ padding: "4px 8px", color: SAP.brand, fontWeight: 700 }}>{r.unitName || "-"}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right" }}>{fmt(r.oldCost)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right", color: isSameValue(r.newCost, r.oldCost) ? SAP.text : SAP.negative, fontWeight: isSameValue(r.newCost, r.oldCost) ? 400 : 700 }}>{fmt(r.newCost)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right" }}>{fmt(r.lv0_new)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right", color: isSameValue(r.lv0_new, r.lv0_old) ? SAP.text : SAP.negative }}>{fmt(r.lv0_old)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right" }}>{fmt(r.lv1_new)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right", color: isSameValue(r.lv1_new, r.lv1_old) ? SAP.text : SAP.negative }}>{fmt(r.lv1_old)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right" }}>{fmt(r.lv5_new)}</td>
                        <td style={{ padding: "4px 8px", textAlign: "right", color: isSameValue(r.lv5_new, r.lv5_old) ? SAP.text : SAP.negative }}>{fmt(r.lv5_old)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Panel>
          </div>
        )}
      </div>

      {/* Footer */}
      <div style={{
        padding: "8px 20px", borderTop: `1px solid ${SAP.border}`,
        background: SAP.surface, display: "flex", justifyContent: "space-between",
        alignItems: "center", fontSize: "11px", color: SAP.textLight,
        fontFamily: SAP.fontFamily, marginTop: "16px",
      }}>
        <span>Price Comparison Module | Procurement Department</span>
        <span>Powered by Client-Side Processing | No Server Required</span>
      </div>

      <style>{`
        @keyframes blink { 0%, 49% { opacity: 1; } 50%, 100% { opacity: 0; } }
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: ${SAP.headerBg}; }
        ::-webkit-scrollbar-thumb { background: ${SAP.borderDark}; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: ${SAP.neutral}; }
        input:focus { outline: 2px solid ${SAP.brand}; outline-offset: -1px; }
        button:hover:not(:disabled) { filter: brightness(0.95); }
      `}</style>
    </div>
  );
}

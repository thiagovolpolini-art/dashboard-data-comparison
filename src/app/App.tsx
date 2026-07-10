import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area,
  PieChart, Pie, Cell, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";
import {
  Upload, BarChart2, GitCompare, Cpu, Hash, AlignLeft,
  AlertCircle, Loader2, SendHorizonal, FileText, Trash2, Key,
  LayoutDashboard, ChevronDown, X, Plus, Layers, CheckCircle2,
  TableProperties
} from "lucide-react";

// ─── Types ───────────────────────────────────────────────────────────────────

type Row = Record<string, string | number>;

interface DataFile {
  id: string;
  name: string;
  sheet: string;
  rows: Row[];
  columns: string[];
  numericColumns: string[];
  categoricalColumns: string[];
}

interface PendingExcel {
  fileName: string;
  workbook: XLSX.WorkBook;
}

interface AiMessage {
  role: "user" | "assistant";
  content: string;
}

type ChartType = "bar" | "line" | "area" | "pie" | "scatter";
type Tab = "dashboard" | "compare" | "ai";

// ─── Constants ────────────────────────────────────────────────────────────────

const COLORS = ["#00d4ff", "#7b61ff", "#10b981", "#f59e0b", "#f43f5e", "#06b6d4", "#a78bfa", "#34d399"];

const CHART_TYPES: { id: ChartType; label: string }[] = [
  { id: "bar", label: "Bar" },
  { id: "line", label: "Line" },
  { id: "area", label: "Area" },
  { id: "pie", label: "Pie" },
  { id: "scatter", label: "Scatter" },
];

// ─── Helpers ──────────────────────────────────────────────────────────────────

function isNumeric(val: unknown): boolean {
  if (val === null || val === undefined || val === "") return false;
  return !isNaN(Number(val));
}

function parseRows(raw: Record<string, unknown>[]): Row[] {
  return raw.map((r) => {
    const out: Row = {};
    for (const k of Object.keys(r)) {
      const v = r[k];
      out[k] = isNumeric(v) ? Number(v) : String(v ?? "");
    }
    return out;
  });
}

function buildFile(name: string, sheet: string, raw: Record<string, unknown>[]): DataFile {
  const rows = parseRows(raw);
  const columns = rows.length > 0 ? Object.keys(rows[0]) : [];
  const numericColumns = columns.filter((c) => rows.some((r) => typeof r[c] === "number"));
  const categoricalColumns = columns.filter((c) => !numericColumns.includes(c));
  return { id: crypto.randomUUID(), name, sheet, rows, columns, numericColumns, categoricalColumns };
}

function sheetToFile(wb: XLSX.WorkBook, fileName: string, sheetName: string): DataFile {
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws) as Record<string, unknown>[];
  return buildFile(`${fileName} · ${sheetName}`, sheetName, data);
}

function summarize(file: DataFile): string {
  const preview = file.rows.slice(0, 8).map((r) => JSON.stringify(r)).join("\n");
  return `Dataset: "${file.name}" (aba: ${file.sheet})\nLinhas: ${file.rows.length}\nColunas: ${file.columns.join(", ")}\nAmostra:\n${preview}`;
}

function statLine(file: DataFile, col: string) {
  const vals = file.rows.map((r) => Number(r[col])).filter((v) => !isNaN(v));
  if (!vals.length) return { min: 0, max: 0, avg: 0, sum: 0 };
  const sum = vals.reduce((a, b) => a + b, 0);
  return { min: Math.min(...vals), max: Math.max(...vals), avg: sum / vals.length, sum };
}

function topN(file: DataFile, catCol: string, numCol: string, n = 12): Row[] {
  const agg: Record<string, number> = {};
  for (const r of file.rows) {
    const cat = String(r[catCol] ?? "");
    agg[cat] = (agg[cat] ?? 0) + Number(r[numCol] ?? 0);
  }
  return Object.entries(agg)
    .sort((a, b) => b[1] - a[1])
    .slice(0, n)
    .map(([k, v]) => ({ [catCol]: k, [numCol]: v }));
}

function fmt(n: number, dec = 2) {
  return n.toLocaleString("pt-BR", { maximumFractionDigits: dec });
}

// ─── Sheet Picker Modal ───────────────────────────────────────────────────────

function SheetPickerModal({
  pending, onConfirm, onCancel,
}: {
  pending: PendingExcel;
  onConfirm: (sheets: string[]) => void;
  onCancel: () => void;
}) {
  const sheets = pending.workbook.SheetNames;
  const [selected, setSelected] = useState<string[]>(sheets.length === 1 ? sheets : []);

  const toggle = (s: string) =>
    setSelected((prev) =>
      prev.includes(s) ? prev.filter((x) => x !== s) : [...prev, s]
    );

  const selectAll = () => setSelected(sheets);
  const clearAll = () => setSelected([]);

  return (
    <div className="absolute inset-0 z-50 bg-background/85 backdrop-blur-sm flex items-center justify-center p-6">
      <div className="bg-card border border-border rounded-2xl p-6 w-full max-w-md flex flex-col gap-5 shadow-2xl">
        {/* Header */}
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <Layers className="w-4 h-4 text-primary" />
              <h3 className="font-semibold text-foreground">Selecionar abas</h3>
            </div>
            <p className="text-xs text-muted-foreground mt-1 font-mono truncate">{pending.fileName}</p>
          </div>
          <button onClick={onCancel} className="text-muted-foreground hover:text-foreground transition-colors mt-0.5">
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Sheets grid */}
        <div className="flex flex-col gap-2">
          <div className="flex items-center justify-between mb-1">
            <span className="text-xs text-muted-foreground font-mono">{sheets.length} aba{sheets.length !== 1 ? "s" : ""} encontrada{sheets.length !== 1 ? "s" : ""}</span>
            <div className="flex gap-3 text-xs">
              <button onClick={selectAll} className="text-primary hover:opacity-80 transition-opacity">Todas</button>
              <button onClick={clearAll} className="text-muted-foreground hover:text-foreground transition-colors">Nenhuma</button>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-2 max-h-56 overflow-y-auto pr-1">
            {sheets.map((s) => {
              const ws = pending.workbook.Sheets[s];
              const rowCount = XLSX.utils.sheet_to_json(ws).length;
              const isOn = selected.includes(s);
              return (
                <button
                  key={s}
                  onClick={() => toggle(s)}
                  className={`flex items-start gap-2.5 p-3 rounded-lg border text-left transition-all ${
                    isOn
                      ? "bg-primary/10 border-primary/40 text-foreground"
                      : "bg-muted/40 border-border text-muted-foreground hover:border-border/80 hover:text-foreground"
                  }`}
                >
                  <div className={`mt-0.5 w-4 h-4 rounded-sm border flex items-center justify-center shrink-0 transition-colors ${isOn ? "bg-primary border-primary" : "border-muted-foreground/40"}`}>
                    {isOn && <CheckCircle2 className="w-3 h-3 text-primary-foreground" />}
                  </div>
                  <div className="min-w-0">
                    <div className="text-xs font-medium truncate">{s}</div>
                    <div className="text-[10px] font-mono text-muted-foreground">{rowCount} linhas</div>
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        {/* Actions */}
        <div className="flex items-center justify-between pt-1 border-t border-border">
          <span className="text-xs text-muted-foreground">{selected.length} selecionada{selected.length !== 1 ? "s" : ""}</span>
          <div className="flex gap-2">
            <button onClick={onCancel} className="px-3 py-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors">Cancelar</button>
            <button
              onClick={() => selected.length > 0 && onConfirm(selected)}
              disabled={selected.length === 0}
              className="px-4 py-1.5 text-sm bg-primary text-primary-foreground rounded-lg font-medium hover:opacity-90 disabled:opacity-40 transition-opacity"
            >
              Importar {selected.length > 0 ? `(${selected.length})` : ""}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Upload Zone ──────────────────────────────────────────────────────────────

function UploadZone({ onFile, onExcel }: { onFile: (f: DataFile) => void; onExcel: (p: PendingExcel) => void }) {
  const [drag, setDrag] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => onFile(buildFile(file.name, "CSV", res.data as Record<string, unknown>[])),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      if (wb.SheetNames.length === 1) {
        onFile(sheetToFile(wb, file.name, wb.SheetNames[0]));
      } else {
        onExcel({ fileName: file.name, workbook: wb });
      }
    }
  }, [onFile, onExcel]);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDrag(false);
    Array.from(e.dataTransfer.files).forEach(handleFile);
  }, [handleFile]);

  return (
    <div
      onDrop={onDrop}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onClick={() => inputRef.current?.click()}
      className={`group border-2 border-dashed rounded-xl p-10 flex flex-col items-center gap-4 cursor-pointer transition-all duration-200 ${drag ? "border-primary bg-primary/10" : "border-border hover:border-primary/50 hover:bg-primary/5"}`}
    >
      <div className={`w-14 h-14 rounded-full flex items-center justify-center transition-colors ${drag ? "bg-primary/20" : "bg-muted group-hover:bg-primary/10"}`}>
        <Upload className={`w-6 h-6 transition-colors ${drag ? "text-primary" : "text-muted-foreground group-hover:text-primary"}`} />
      </div>
      <div className="text-center">
        <p className="text-sm font-medium text-foreground">Solte arquivos aqui ou clique para importar</p>
        <p className="text-xs text-muted-foreground mt-1">CSV, XLSX, XLS · múltiplas abas suportadas</p>
      </div>
      <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" multiple className="hidden"
        onChange={(e) => Array.from(e.target.files ?? []).forEach(handleFile)} />
    </div>
  );
}

// ─── Shared: Select ───────────────────────────────────────────────────────────

function Select({ value, onChange, options, placeholder }: {
  value: string; onChange: (v: string) => void;
  options: { value: string; label: string }[]; placeholder?: string;
}) {
  return (
    <div className="relative">
      <select value={value} onChange={(e) => onChange(e.target.value)}
        className="w-full appearance-none bg-secondary border border-border text-foreground text-sm rounded-md px-3 py-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary cursor-pointer">
        {placeholder && <option value="">{placeholder}</option>}
        {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
      <ChevronDown className="absolute right-2 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground pointer-events-none" />
    </div>
  );
}

// ─── Shared: Chart ────────────────────────────────────────────────────────────

function DataChart({ data, chartType, xKey, yKey, color = "#00d4ff" }: {
  data: Row[]; chartType: ChartType; xKey: string; yKey: string; color?: string;
}) {
  const slice = data.slice(0, 30);

  if (chartType === "pie") {
    return (
      <ResponsiveContainer width="100%" height={260}>
        <PieChart>
          <Pie data={slice} dataKey={yKey} nameKey={xKey} cx="50%" cy="50%" outerRadius={95}
            label={({ name, percent }) => `${String(name).slice(0, 10)} ${(percent * 100).toFixed(0)}%`}>
            {slice.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
          </Pie>
          <Tooltip contentStyle={{ background: "#0d1526", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 8 }} />
        </PieChart>
      </ResponsiveContainer>
    );
  }

  if (chartType === "scatter") {
    return (
      <ResponsiveContainer width="100%" height={260}>
        <ScatterChart margin={{ top: 10, right: 10, bottom: 10, left: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
          <XAxis dataKey={xKey} stroke="#64748b" tick={{ fontSize: 11, fontFamily: "JetBrains Mono" }} />
          <YAxis dataKey={yKey} stroke="#64748b" tick={{ fontSize: 11, fontFamily: "JetBrains Mono" }} />
          <Tooltip contentStyle={{ background: "#0d1526", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 8 }} />
          <Scatter data={slice} fill={color} />
        </ScatterChart>
      </ResponsiveContainer>
    );
  }

  const sharedAxis = (
    <>
      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
      <XAxis dataKey={xKey} stroke="#64748b" tick={{ fontSize: 11, fontFamily: "JetBrains Mono" }} angle={-35} textAnchor="end" interval="preserveStartEnd" />
      <YAxis stroke="#64748b" tick={{ fontSize: 11, fontFamily: "JetBrains Mono" }} width={60} />
      <Tooltip contentStyle={{ background: "#0d1526", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 8 }} />
    </>
  );

  const margin = { top: 10, right: 10, bottom: 36, left: 0 };

  if (chartType === "bar") return (
    <ResponsiveContainer width="100%" height={260}>
      <BarChart data={slice} margin={margin}>
        {sharedAxis}
        <Bar dataKey={yKey} fill={color} radius={[3, 3, 0, 0]} />
      </BarChart>
    </ResponsiveContainer>
  );

  if (chartType === "area") return (
    <ResponsiveContainer width="100%" height={260}>
      <AreaChart data={slice} margin={margin}>
        {sharedAxis}
        <defs>
          <linearGradient id={`grad-${color.replace("#", "")}`} x1="0" y1="0" x2="0" y2="1">
            <stop offset="5%" stopColor={color} stopOpacity={0.3} />
            <stop offset="95%" stopColor={color} stopOpacity={0} />
          </linearGradient>
        </defs>
        <Area type="monotone" dataKey={yKey} stroke={color} fill={`url(#grad-${color.replace("#", "")})`} strokeWidth={2} dot={false} />
      </AreaChart>
    </ResponsiveContainer>
  );

  return (
    <ResponsiveContainer width="100%" height={260}>
      <LineChart data={slice} margin={margin}>
        {sharedAxis}
        <Line type="monotone" dataKey={yKey} stroke={color} strokeWidth={2} dot={false} />
      </LineChart>
    </ResponsiveContainer>
  );
}

// ─── Shared: Table Preview ────────────────────────────────────────────────────

function TablePreview({ file }: { file: DataFile }) {
  const cols = file.columns.slice(0, 8);
  const rows = file.rows.slice(0, 12);
  return (
    <div className="overflow-auto max-h-60 rounded-lg border border-border">
      <table className="w-full text-xs font-mono">
        <thead className="sticky top-0 bg-secondary">
          <tr>{cols.map((c) => <th key={c} className="px-3 py-2 text-left text-muted-foreground font-medium whitespace-nowrap border-b border-border">{c}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} className={i % 2 === 0 ? "bg-card" : "bg-secondary/40"}>
              {cols.map((c) => <td key={c} className="px-3 py-1.5 whitespace-nowrap text-foreground/80 border-b border-border/50">{String(r[c] ?? "")}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── KPI Card ─────────────────────────────────────────────────────────────────

function KpiCard({ label, value, sub }: { label: string; value: string; sub?: string }) {
  return (
    <div className="bg-card border border-border rounded-lg p-4 flex flex-col gap-1">
      <span className="text-[10px] font-mono uppercase tracking-widest text-muted-foreground">{label}</span>
      <span className="text-2xl font-bold text-foreground leading-none">{value}</span>
      {sub && <span className="text-xs text-muted-foreground truncate">{sub}</span>}
    </div>
  );
}

// ─── Dashboard Tab ────────────────────────────────────────────────────────────

function DashboardTab({ files }: { files: DataFile[] }) {
  const [activeId, setActiveId] = useState("");
  const [chartType, setChartType] = useState<ChartType>("bar");
  const [xCol, setXCol] = useState("");
  const [yCol, setYCol] = useState("");

  const file = files.find((f) => f.id === activeId) ?? files[0];

  useEffect(() => {
    if (file) {
      setXCol(file.categoricalColumns[0] ?? file.columns[0] ?? "");
      setYCol(file.numericColumns[0] ?? file.columns[1] ?? "");
    }
  }, [file?.id]);

  if (!files.length) return (
    <div className="flex flex-col items-center justify-center h-full gap-3 text-muted-foreground">
      <BarChart2 className="w-10 h-10 opacity-20" />
      <p className="text-sm">Importe um arquivo para ver o dashboard</p>
    </div>
  );

  const stats = file && yCol ? statLine(file, yCol) : null;
  const chartData = file && xCol && yCol ? topN(file, xCol, yCol) : [];

  return (
    <div className="flex flex-col gap-5 h-full overflow-y-auto pr-1">
      <div className="flex flex-wrap gap-3 items-center justify-between">
        <div className="w-64">
          <Select value={activeId || file?.id || ""} onChange={setActiveId}
            options={files.map((f) => ({ value: f.id, label: f.name }))} />
        </div>
        <div className="flex gap-1">
          {CHART_TYPES.map((ct) => (
            <button key={ct.id} onClick={() => setChartType(ct.id)}
              className={`px-3 py-1.5 text-xs font-mono rounded-md transition-colors ${chartType === ct.id ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"}`}>
              {ct.label}
            </button>
          ))}
        </div>
      </div>

      {stats && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <KpiCard label="Linhas" value={file.rows.length.toLocaleString("pt-BR")} sub="registros" />
          <KpiCard label="Soma" value={fmt(stats.sum)} sub={yCol} />
          <KpiCard label="Média" value={fmt(stats.avg)} sub={yCol} />
          <KpiCard label="Máximo" value={fmt(stats.max)} sub={yCol} />
        </div>
      )}

      <div className="bg-card border border-border rounded-xl p-5">
        <div className="flex flex-wrap gap-3 mb-4">
          <div className="flex-1 min-w-36">
            <label className="text-xs text-muted-foreground mb-1 block font-mono">Eixo X / Categoria</label>
            <Select value={xCol} onChange={setXCol} options={file.columns.map((c) => ({ value: c, label: c }))} />
          </div>
          <div className="flex-1 min-w-36">
            <label className="text-xs text-muted-foreground mb-1 block font-mono">Eixo Y / Valor</label>
            <Select value={yCol} onChange={setYCol} options={file.columns.map((c) => ({ value: c, label: c }))} />
          </div>
        </div>
        {chartData.length > 0
          ? <DataChart data={chartData} chartType={chartType} xKey={xCol} yKey={yCol} />
          : <div className="h-52 flex items-center justify-center text-muted-foreground text-sm">Selecione colunas para visualizar</div>}
      </div>

      <div className="bg-card border border-border rounded-xl p-5">
        <h3 className="text-sm font-medium text-foreground mb-3 flex items-center gap-2">
          <AlignLeft className="w-4 h-4 text-muted-foreground" /> Prévia dos dados
        </h3>
        <TablePreview file={file} />
      </div>
    </div>
  );
}

// ─── Compare Tab ──────────────────────────────────────────────────────────────

interface PanelState {
  fileId: string;
  xCol: string;
  yCol: string;
  chartType: ChartType;
}

function ComparePanel({
  files, state, onChange, color, label,
}: {
  files: DataFile[];
  state: PanelState;
  onChange: (s: Partial<PanelState>) => void;
  color: string;
  label: string;
}) {
  const file = files.find((f) => f.id === state.fileId);

  useEffect(() => {
    if (file) {
      onChange({
        xCol: file.categoricalColumns[0] ?? file.columns[0] ?? "",
        yCol: file.numericColumns[0] ?? file.columns[1] ?? "",
      });
    }
  }, [state.fileId]);

  const chartData = file && state.xCol && state.yCol ? topN(file, state.xCol, state.yCol) : [];
  const stats = file && state.yCol ? statLine(file, state.yCol) : null;

  return (
    <div className="flex-1 min-w-0 flex flex-col gap-3">
      {/* Panel label */}
      <div className="flex items-center gap-2">
        <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: color }} />
        <span className="text-xs font-mono text-muted-foreground uppercase tracking-widest">{label}</span>
      </div>

      <div className="bg-card border border-border rounded-xl p-4 flex flex-col gap-3 flex-1">
        {/* File select */}
        <Select
          value={state.fileId}
          onChange={(v) => onChange({ fileId: v })}
          options={files.map((f) => ({ value: f.id, label: f.name }))}
          placeholder="Selecione um dataset..."
        />

        {file && (
          <>
            {/* Column selects */}
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-mono text-muted-foreground uppercase block mb-1">X / Categoria</label>
                <Select value={state.xCol} onChange={(v) => onChange({ xCol: v })}
                  options={file.columns.map((c) => ({ value: c, label: c }))} />
              </div>
              <div>
                <label className="text-[10px] font-mono text-muted-foreground uppercase block mb-1">Y / Valor</label>
                <Select value={state.yCol} onChange={(v) => onChange({ yCol: v })}
                  options={file.columns.map((c) => ({ value: c, label: c }))} />
              </div>
            </div>

            {/* Chart type */}
            <div className="flex gap-1">
              {CHART_TYPES.map((ct) => (
                <button key={ct.id} onClick={() => onChange({ chartType: ct.id })}
                  className={`flex-1 py-1 text-[10px] font-mono rounded transition-colors ${state.chartType === ct.id ? "text-foreground border border-border/80" : "text-muted-foreground hover:text-foreground"}`}
                  style={state.chartType === ct.id ? { background: `${color}18`, borderColor: `${color}50`, color } : {}}>
                  {ct.label}
                </button>
              ))}
            </div>

            {/* Chart */}
            {chartData.length > 0
              ? <DataChart data={chartData} chartType={state.chartType} xKey={state.xCol} yKey={state.yCol} color={color} />
              : <div className="h-52 flex items-center justify-center text-muted-foreground text-sm">Selecione colunas</div>}

            {/* Mini stats */}
            {stats && (
              <div className="grid grid-cols-3 gap-2 pt-2 border-t border-border">
                {[
                  { label: "Soma", value: fmt(stats.sum) },
                  { label: "Média", value: fmt(stats.avg) },
                  { label: "Máx", value: fmt(stats.max) },
                ].map(({ label: l, value: v }) => (
                  <div key={l} className="text-center">
                    <div className="text-[10px] font-mono text-muted-foreground uppercase">{l}</div>
                    <div className="text-sm font-bold text-foreground">{v}</div>
                  </div>
                ))}
              </div>
            )}

            {/* Table toggle */}
            <details className="group">
              <summary className="flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground cursor-pointer list-none transition-colors">
                <TableProperties className="w-3.5 h-3.5" />
                Ver tabela
              </summary>
              <div className="mt-2">
                <TablePreview file={file} />
              </div>
            </details>
          </>
        )}
      </div>
    </div>
  );
}

function CompareTab({ files }: { files: DataFile[] }) {
  const [panels, setPanels] = useState<PanelState[]>([
    { fileId: files[0]?.id ?? "", xCol: "", yCol: "", chartType: "bar" },
    { fileId: files[1]?.id ?? files[0]?.id ?? "", xCol: "", yCol: "", chartType: "line" },
  ]);

  const update = (i: number, patch: Partial<PanelState>) =>
    setPanels((prev) => prev.map((p, idx) => idx === i ? { ...p, ...patch } : p));

  if (!files.length) return (
    <div className="flex flex-col items-center justify-center h-full gap-3 text-muted-foreground">
      <GitCompare className="w-10 h-10 opacity-20" />
      <p className="text-sm">Importe arquivos para comparar</p>
    </div>
  );

  const PANEL_COLORS = ["#00d4ff", "#7b61ff", "#10b981", "#f59e0b"];
  const PANEL_LABELS = ["Painel A", "Painel B", "Painel C", "Painel D"];

  const addPanel = () => {
    if (panels.length >= 4) return;
    setPanels((prev) => [...prev, { fileId: files[0]?.id ?? "", xCol: "", yCol: "", chartType: "bar" }]);
  };

  const removePanel = (i: number) => {
    if (panels.length <= 1) return;
    setPanels((prev) => prev.filter((_, idx) => idx !== i));
  };

  return (
    <div className="flex flex-col gap-4 h-full overflow-y-auto pr-1">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2 text-sm text-muted-foreground">
          <GitCompare className="w-4 h-4" />
          <span>Compare datasets lado a lado — abas do mesmo arquivo ou arquivos diferentes</span>
        </div>
        <button onClick={addPanel} disabled={panels.length >= 4}
          className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg border border-border text-muted-foreground hover:text-foreground hover:border-primary/40 disabled:opacity-30 transition-colors">
          <Plus className="w-3.5 h-3.5" />
          Painel
        </button>
      </div>

      <div className={`grid gap-4 ${panels.length === 1 ? "grid-cols-1" : panels.length <= 2 ? "grid-cols-1 md:grid-cols-2" : "grid-cols-1 md:grid-cols-2 xl:grid-cols-2"}`}>
        {panels.map((p, i) => (
          <div key={i} className="relative">
            {panels.length > 1 && (
              <button onClick={() => removePanel(i)}
                className="absolute top-0 right-0 z-10 text-muted-foreground hover:text-destructive transition-colors p-1">
                <X className="w-3.5 h-3.5" />
              </button>
            )}
            <ComparePanel files={files} state={p} onChange={(patch) => update(i, patch)}
              color={PANEL_COLORS[i % PANEL_COLORS.length]} label={PANEL_LABELS[i]} />
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── AI Tab ───────────────────────────────────────────────────────────────────

function AiTab({ files }: { files: DataFile[] }) {
  const [messages, setMessages] = useState<AiMessage[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [apiKey, setApiKey] = useState(() => localStorage.getItem("anthropic_key") ?? "");
  const [showKeyModal, setShowKeyModal] = useState(false);
  const [tempKey, setTempKey] = useState("");
  const endRef = useRef<HTMLDivElement>(null);

  useEffect(() => { endRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages]);

  const saveKey = () => {
    localStorage.setItem("anthropic_key", tempKey);
    setApiKey(tempKey);
    setShowKeyModal(false);
  };

  const send = async () => {
    if (!input.trim() || loading) return;
    if (!apiKey) { setTempKey(""); setShowKeyModal(true); return; }

    const userMsg: AiMessage = { role: "user", content: input };
    setMessages((m) => [...m, userMsg]);
    setInput("");
    setLoading(true);

    const context = files.map(summarize).join("\n\n---\n\n");
    const system = `Você é um analista de dados especialista. Você tem acesso aos seguintes datasets:\n\n${context}\n\nResponda em português com análises precisas baseadas nos dados reais fornecidos. Cite números específicos.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "content-type": "application/json",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1024,
          system,
          messages: [...messages, userMsg].map((m) => ({ role: m.role, content: m.content })),
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error((err as { error?: { message?: string } }).error?.message ?? `HTTP ${res.status}`);
      }

      const data = await res.json() as { content: Array<{ text: string }> };
      setMessages((m) => [...m, { role: "assistant", content: data.content?.[0]?.text ?? "Sem resposta." }]);
    } catch (e) {
      setMessages((m) => [...m, { role: "assistant", content: `Erro: ${(e as Error).message}` }]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex flex-col h-full">
      {showKeyModal && (
        <div className="absolute inset-0 z-50 bg-background/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-card border border-border rounded-xl p-6 w-full max-w-md flex flex-col gap-4 shadow-2xl">
            <div className="flex items-center gap-2">
              <Key className="w-5 h-5 text-primary" />
              <h3 className="font-semibold text-foreground">Chave da API Anthropic</h3>
            </div>
            <p className="text-sm text-muted-foreground">Insira sua chave de API. Ela é salva somente no navegador, nunca enviada para servidores externos.</p>
            <input type="password" placeholder="sk-ant-..." value={tempKey}
              onChange={(e) => setTempKey(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && saveKey()}
              className="bg-muted border border-border rounded-md px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-1 focus:ring-primary font-mono" />
            <div className="flex gap-2 justify-end">
              <button onClick={() => setShowKeyModal(false)} className="px-4 py-2 text-sm text-muted-foreground hover:text-foreground">Cancelar</button>
              <button onClick={saveKey} className="px-4 py-2 text-sm bg-primary text-primary-foreground rounded-md font-medium hover:opacity-90">Salvar</button>
            </div>
          </div>
        </div>
      )}

      <div className="flex items-center justify-between pb-3 border-b border-border shrink-0">
        <div className="flex items-center gap-2">
          <Cpu className="w-4 h-4 text-primary" />
          <span className="text-sm font-medium">Assistente IA</span>
          <span className="text-[10px] font-mono text-muted-foreground bg-muted px-2 py-0.5 rounded-full">claude-sonnet-4-6</span>
        </div>
        <button onClick={() => { setTempKey(apiKey); setShowKeyModal(true); }}
          className="flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground transition-colors">
          <Key className="w-3.5 h-3.5" />
          {apiKey ? "Trocar chave" : "Configurar API"}
        </button>
      </div>

      {!files.length ? (
        <div className="flex-1 flex flex-col items-center justify-center gap-3 text-muted-foreground">
          <AlertCircle className="w-8 h-8 opacity-30" />
          <p className="text-sm">Importe dados antes de usar o assistente</p>
        </div>
      ) : (
        <div className="flex-1 overflow-y-auto py-4 flex flex-col gap-4">
          {messages.length === 0 && (
            <div className="flex flex-col gap-3 text-center mt-6">
              <div className="w-12 h-12 rounded-full bg-primary/10 border border-primary/20 flex items-center justify-center mx-auto">
                <Cpu className="w-5 h-5 text-primary" />
              </div>
              <p className="text-sm text-muted-foreground">
                {files.length} dataset{files.length > 1 ? "s" : ""} carregado{files.length > 1 ? "s" : ""}. Faça uma pergunta.
              </p>
              <div className="flex flex-wrap gap-2 justify-center mt-1">
                {["Qual categoria tem mais registros?", "Compare os totais entre os datasets", "Qual o maior valor e em qual linha?", "Faça um resumo estatístico dos dados"].map((s) => (
                  <button key={s} onClick={() => setInput(s)}
                    className="text-xs px-3 py-1.5 rounded-full bg-muted border border-border text-muted-foreground hover:text-foreground hover:border-primary/40 transition-colors">
                    {s}
                  </button>
                ))}
              </div>
            </div>
          )}
          {messages.map((m, i) => (
            <div key={i} className={`flex gap-3 ${m.role === "user" ? "flex-row-reverse" : "flex-row"}`}>
              <div className={`w-7 h-7 rounded-full flex items-center justify-center shrink-0 text-[10px] font-bold ${m.role === "user" ? "bg-accent text-accent-foreground" : "bg-primary/20 text-primary border border-primary/30"}`}>
                {m.role === "user" ? "U" : "IA"}
              </div>
              <div className={`max-w-[80%] rounded-xl px-4 py-3 text-sm leading-relaxed whitespace-pre-wrap ${m.role === "user" ? "bg-accent/20 text-foreground border border-accent/20" : "bg-card text-foreground border border-border"}`}>
                {m.content}
              </div>
            </div>
          ))}
          {loading && (
            <div className="flex gap-3 items-center">
              <div className="w-7 h-7 rounded-full bg-primary/20 border border-primary/30 flex items-center justify-center">
                <Loader2 className="w-3.5 h-3.5 text-primary animate-spin" />
              </div>
              <span className="text-sm text-muted-foreground">Analisando dados...</span>
            </div>
          )}
          <div ref={endRef} />
        </div>
      )}

      {files.length > 0 && (
        <div className="flex gap-2 pt-3 border-t border-border shrink-0">
          <input type="text" placeholder="Pergunte sobre seus dados..." value={input}
            onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => e.key === "Enter" && send()}
            className="flex-1 bg-muted border border-border rounded-lg px-4 py-2.5 text-sm text-foreground placeholder:text-muted-foreground focus:outline-none focus:ring-1 focus:ring-primary" />
          <button onClick={send} disabled={loading || !input.trim()}
            className="px-4 py-2.5 bg-primary text-primary-foreground rounded-lg hover:opacity-90 disabled:opacity-40 transition-opacity">
            <SendHorizonal className="w-4 h-4" />
          </button>
        </div>
      )}
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [files, setFiles] = useState<DataFile[]>([]);
  const [tab, setTab] = useState<Tab>("dashboard");
  const [showUpload, setShowUpload] = useState(false);
  const [pendingExcel, setPendingExcel] = useState<PendingExcel | null>(null);

  const addFile = (f: DataFile) => { setFiles((p) => [...p, f]); setShowUpload(false); };
  const removeFile = (id: string) => setFiles((p) => p.filter((f) => f.id !== id));

  const handleExcelConfirm = (sheets: string[]) => {
    if (!pendingExcel) return;
    sheets.forEach((s) => addFile(sheetToFile(pendingExcel.workbook, pendingExcel.fileName, s)));
    setPendingExcel(null);
    setShowUpload(false);
  };

  const NAV = [
    { id: "dashboard" as Tab, icon: LayoutDashboard, label: "Dashboard" },
    { id: "compare" as Tab, icon: GitCompare, label: "Comparar" },
    { id: "ai" as Tab, icon: Cpu, label: "IA Analista" },
  ];

  return (
    <div className="flex h-screen w-full bg-background text-foreground overflow-hidden" style={{ fontFamily: "Inter, sans-serif" }}>

      {/* Sidebar */}
      <aside className="w-56 shrink-0 flex flex-col bg-sidebar border-r border-sidebar-border">
        <div className="px-5 py-5 border-b border-sidebar-border">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded-md bg-primary flex items-center justify-center">
              <BarChart2 className="w-4 h-4 text-primary-foreground" />
            </div>
            <div>
              <div className="text-sm font-bold text-foreground leading-none">DataLens</div>
              <div className="text-[10px] text-muted-foreground font-mono mt-0.5">Analytics Platform</div>
            </div>
          </div>
        </div>

        <nav className="flex-1 px-3 py-4 flex flex-col gap-1">
          {NAV.map(({ id, icon: Icon, label }) => (
            <button key={id} onClick={() => setTab(id)}
              className={`flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-left transition-colors w-full ${tab === id ? "bg-primary/15 text-primary" : "text-sidebar-foreground hover:text-foreground hover:bg-sidebar-accent"}`}>
              <Icon className="w-4 h-4 shrink-0" />
              {label}
            </button>
          ))}
        </nav>

        {/* Files */}
        <div className="px-3 pb-4 flex flex-col gap-2 border-t border-sidebar-border pt-3">
          <div className="flex items-center justify-between px-1 mb-1">
            <span className="text-[10px] font-mono uppercase tracking-widest text-muted-foreground">Datasets</span>
            <button onClick={() => setShowUpload(true)} className="text-muted-foreground hover:text-primary transition-colors">
              <Plus className="w-3.5 h-3.5" />
            </button>
          </div>
          {files.length === 0 && <p className="text-xs text-muted-foreground px-1 italic">Nenhum arquivo</p>}
          <div className="flex flex-col gap-1 max-h-52 overflow-y-auto">
            {files.map((f) => (
              <div key={f.id} className="flex items-center gap-2 px-2 py-1.5 rounded-md bg-sidebar-accent/60 group">
                <FileText className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                <div className="flex-1 min-w-0">
                  <div className="text-xs text-foreground/80 truncate">{f.name}</div>
                  <div className="text-[10px] font-mono text-muted-foreground">{f.rows.length} linhas</div>
                </div>
                <button onClick={() => removeFile(f.id)} className="opacity-0 group-hover:opacity-100 text-muted-foreground hover:text-destructive transition-opacity shrink-0">
                  <Trash2 className="w-3 h-3" />
                </button>
              </div>
            ))}
          </div>
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 flex flex-col min-w-0 relative">
        <header className="flex items-center justify-between px-6 py-3.5 border-b border-border bg-card/40 backdrop-blur-sm shrink-0">
          <div>
            <h1 className="text-base font-semibold text-foreground">
              {tab === "dashboard" ? "Dashboard" : tab === "compare" ? "Comparação" : "IA Analista"}
            </h1>
            <p className="text-xs text-muted-foreground">
              {files.length ? `${files.length} dataset${files.length > 1 ? "s" : ""} carregado${files.length > 1 ? "s" : ""}` : "Nenhum dado carregado"}
            </p>
          </div>
          <button onClick={() => setShowUpload(true)}
            className="flex items-center gap-2 px-4 py-2 bg-primary text-primary-foreground text-sm font-medium rounded-lg hover:opacity-90 transition-opacity">
            <Upload className="w-4 h-4" />
            Importar dados
          </button>
        </header>

        {/* Upload modal */}
        {showUpload && !pendingExcel && (
          <div className="absolute inset-0 z-40 bg-background/80 backdrop-blur-sm flex items-center justify-center p-8">
            <div className="bg-card border border-border rounded-2xl p-6 w-full max-w-lg flex flex-col gap-4 shadow-2xl">
              <div className="flex items-center justify-between">
                <h2 className="font-semibold text-foreground">Importar dados</h2>
                <button onClick={() => setShowUpload(false)} className="text-muted-foreground hover:text-foreground transition-colors">
                  <X className="w-5 h-5" />
                </button>
              </div>
              <UploadZone onFile={addFile} onExcel={setPendingExcel} />
              <div className="flex gap-3 text-xs text-muted-foreground">
                {["CSV", "XLSX", "XLS"].map((t) => <span key={t} className="flex items-center gap-1"><Hash className="w-3 h-3" />{t}</span>)}
                <span className="ml-auto text-muted-foreground/60">Excel com múltiplas abas → seletor de sheets</span>
              </div>
            </div>
          </div>
        )}

        {/* Sheet picker */}
        {pendingExcel && (
          <SheetPickerModal
            pending={pendingExcel}
            onConfirm={handleExcelConfirm}
            onCancel={() => { setPendingExcel(null); setShowUpload(false); }}
          />
        )}

        {/* Ambient glow */}
        <div className="absolute inset-0 pointer-events-none overflow-hidden">
          <div className="absolute -top-32 -left-32 w-96 h-96 rounded-full bg-primary/5 blur-3xl" />
          <div className="absolute -bottom-32 -right-32 w-96 h-96 rounded-full bg-accent/5 blur-3xl" />
        </div>

        <div className="relative flex-1 overflow-hidden p-6">
          {tab === "dashboard" && <DashboardTab files={files} />}
          {tab === "compare" && <CompareTab files={files} />}
          {tab === "ai" && <AiTab files={files} />}
        </div>
      </main>
    </div>
  );
}

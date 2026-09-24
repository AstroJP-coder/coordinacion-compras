import React, { useState, useEffect, useMemo, useRef } from "react";
import { db } from "./firebase";
import {
  collection, doc, getDoc, getDocs, setDoc, deleteDoc, updateDoc, onSnapshot, serverTimestamp, arrayUnion,
} from "firebase/firestore";
import * as XLSX from "xlsx";
import Tesseract from "tesseract.js";
import {
  KeyRound, AlertTriangle, LogOut, Users, Trash2, ShoppingCart, ClipboardList,
  Truck, FileCheck2, Shield, PackageCheck, Upload, Download, Send, Plus, X, Save,
  FileSpreadsheet, FileText, Check, ChevronLeft, Paperclip, Search, Image as ImageIcon, ScanLine,
  CalendarClock, Hash, Clock, Building2, Inbox, History, User,
} from "lucide-react";

/* ============ Firestore refs (colecciones aisladas compras_) ============ */
const colUsr = collection(db, "compras_usuarios");
const colSol = collection(db, "compras_solicitudes");

/* ============ paleta (coherente con Salón VIP) ============ */
const C = {
  bg: "#16130E", surface: "#211C15", surface2: "#2B241A", line: "#3A3122",
  text: "#F5EEDD", muted: "#A89A7C", faint: "#6E6450",
  gold: "#D9B25F", amber: "#E0A73B", clay: "#C6603F", teal: "#4FB286", info: "#6FA8C7",
};

/* ============ roles ============ */
const ROLES = [
  { k: "analista_compras", l: "Analista", i: ClipboardList },
  { k: "administrativo_compras", l: "Administrativo", i: FileCheck2 },
  { k: "bodeguero", l: "Bodega", i: Truck },
  { k: "admin", l: "Admin", i: Shield },
];
const rolMeta = (k) => ROLES.find((r) => r.k === k) || { k, l: k, i: Shield };

/* ============ estados de la solicitud ============ */
const ESTADOS = {
  BORRADOR: { l: "Borrador", c: C.faint },
  ENVIADA: { l: "Enviada", c: C.info },
  OC_GENERADA: { l: "OC generada", c: C.amber },
  PENDIENTE_RECEPCION: { l: "Pend. recepción", c: C.gold },
  RECEPCION_PARCIAL: { l: "Recep. parcial", c: C.clay },
  RECEPCIONADA: { l: "Recepcionada", c: C.teal },
  CERRADA: { l: "Cerrada", c: C.muted },
};

/* ============ helpers ============ */
const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
const ms = (t) => (t == null ? 0 : typeof t === "number" ? t : typeof t.toMillis === "function" ? t.toMillis() : t.seconds ? t.seconds * 1000 : 0);
const fecha = (ts) => { const n = ms(ts); return n ? new Date(n).toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit", year: "2-digit" }) : ""; };
const fechaHora = (ts) => { const n = ms(ts); return n ? new Date(n).toLocaleString("es-CL", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : ""; };
const cellToStr = (v) => { if (v == null) return ""; if (v instanceof Date) return v.toISOString().slice(0, 10); return String(v).trim(); };

// Firestore topa en ~1 MB por documento; el original se guarda embebido en base64
// (infla ~33%). Limitamos el archivo original a 700 KB para dejar margen a los ítems.
const MAX_ARCHIVO = 700 * 1024;
const kb = (n) => `${Math.round(n / 1024)} KB`;
const fileToB64 = (file) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = () => res(String(r.result).split(",")[1] || "");
  r.onerror = () => rej(new Error("No se pudo leer el archivo"));
  r.readAsDataURL(file);
});
const descargarArchivo = (a) => {
  if (!a?.data) return;
  const link = document.createElement("a");
  link.href = `data:${a.tipo || "application/octet-stream"};base64,${a.data}`;
  link.download = a.nombre || "original";
  document.body.appendChild(link); link.click(); link.remove();
};
// Comprime/redimensiona una imagen en el navegador (canvas) para que quepa embebida.
const comprimirImagen = (file, maxBytes = MAX_ARCHIVO) => new Promise((resolve) => {
  const img = new window.Image();
  const url = URL.createObjectURL(file);
  img.onload = () => {
    let { width, height } = img;
    const MAXDIM = 1600;
    if (Math.max(width, height) > MAXDIM) { const r = MAXDIM / Math.max(width, height); width = Math.round(width * r); height = Math.round(height * r); }
    const canvas = document.createElement("canvas"); canvas.width = width; canvas.height = height;
    canvas.getContext("2d").drawImage(img, 0, 0, width, height);
    let q = 0.85, dataUrl = canvas.toDataURL("image/jpeg", q);
    while (dataUrl.length * 0.75 > maxBytes && q > 0.3) { q -= 0.1; dataUrl = canvas.toDataURL("image/jpeg", q); }
    URL.revokeObjectURL(url);
    const b64 = dataUrl.split(",")[1] || "";
    const size = Math.round(b64.length * 0.75);
    resolve({ nombre: file.name.replace(/\.\w+$/, "") + ".jpg", tipo: "image/jpeg", size, data: b64, tooBig: size > maxBytes });
  };
  img.onerror = () => resolve(null);
  img.src = url;
});

const ITEM_VACIO = { codigo: "", producto: "", cantidad: "", unidad: "", costo: "", proveedor: "", fechaRequerida: "", observaciones: "" };
const FIELD_DEFS = [
  { k: "codigo", l: "Código ERP", syn: ["codigo erp", "cod erp", "codigo producto", "codigo articulo", "codigo", "cod", "sku", "id producto", "code"] },
  { k: "producto", l: "Producto / Descripción", syn: ["producto", "descripcion", "detalle", "item", "articulo", "glosa", "nombre"] },
  { k: "cantidad", l: "Cantidad", syn: ["cantidad", "cant", "qty", "unidades", "pedido"] },
  { k: "unidad", l: "Unidad", syn: ["unidad", "um", "unidad de medida", "medida", "unit"] },
  { k: "costo", l: "Costo unitario", syn: ["costo unitario", "costo uni", "costo unit", "costo", "precio unitario", "precio uni", "precio", "valor unitario", "valor", "costo neto", "p unit", "precio unit"] },
  { k: "proveedor", l: "Proveedor", syn: ["proveedor", "prov", "vendor"] },
  { k: "fechaRequerida", l: "Fecha requerida", syn: ["fecha requerida", "fecha_requerida", "fecha", "requerida", "fecha entrega", "plazo"] },
  { k: "observaciones", l: "Observaciones", syn: ["observaciones", "observacion", "obs", "nota", "notas", "comentario", "comentarios"] },
];
// columnas de la tabla de ítems (orden + ancho + etiqueta corta)
const COLMETA = [
  { k: "codigo", l: "Cód. ERP", w: 100 },
  { k: "producto", l: "Producto / Descripción", w: null },
  { k: "cantidad", l: "Cantidad", w: 80 },
  { k: "unidad", l: "Unidad", w: 90 },
  { k: "costo", l: "Costo unit.", w: 100 },
  { k: "proveedor", l: "Proveedor", w: 130 },
  { k: "fechaRequerida", l: "Fecha req.", w: 110 },
  { k: "observaciones", l: "Observaciones", w: null },
];
// Parseo tolerante de montos/cantidades: "$1,716", "1.716", "1.716,50", "0,6" → número
const parseMonto = (v) => {
  let s = String(v ?? "").replace(/[^\d.,-]/g, "").trim();
  if (!s) return 0;
  const tienePunto = s.includes("."), tieneComa = s.includes(",");
  if (tienePunto && tieneComa) {
    // el último separador es el decimal
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (tieneComa) {
    const dec = s.split(",").pop();
    s = dec.length === 2 ? s.replace(",", ".") : s.replace(/,/g, ""); // ",XX" = decimal; si no, miles
  } else if (tienePunto) {
    const dec = s.split(".").pop();
    if (dec.length === 3 && s.split(".").length === 2 && s.replace(".", "").length > 3) s = s.replace(/\./g, ""); // "1.716" = miles
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};
const totalItems = (items) => (items || []).reduce((a, it) => a + parseMonto(it.cantidad) * parseMonto(it.costo), 0);
const montoCLP = (n) => new Intl.NumberFormat("es-CL", { style: "currency", currency: "CLP", maximumFractionDigits: 0 }).format(n || 0);
const guessField = (header) => {
  const h = norm(header);
  if (!h) return "ignorar";
  for (const f of FIELD_DEFS) if (f.syn.some((s) => h === norm(s))) return f.k;
  for (const f of FIELD_DEFS) if (f.syn.some((s) => h.includes(norm(s)))) return f.k;
  return "ignorar";
};
// Interpreta una línea de texto (OCR) en un ítem, detectando cantidad y unidad si están.
const UNID = "x|und|un|u|uds?|cajas?|caja|kg|kgs|grs?|gr|lt|lts?|l|bidon(?:es)?|pack|docenas?|doc|bolsas?|sacos?|rollos?|pares?|latas?|botellas?";
const parseLineaOCR = (linea) => {
  const s = linea.trim().replace(/\s+/g, " ");
  if (!s) return null;
  let m = s.match(new RegExp(`^(\\d+(?:[.,]\\d+)?)\\s*(${UNID})?\\.?\\s*[x*·-]?\\s*(.+)$`, "i"));
  if (m && m[3] && /[a-záéíóúñ]/i.test(m[3])) return { ...ITEM_VACIO, producto: m[3].trim(), cantidad: m[1], unidad: (m[2] || "").toLowerCase() };
  m = s.match(new RegExp(`^(.+?)\\s+(\\d+(?:[.,]\\d+)?)\\s*(${UNID})?\\.?$`, "i"));
  if (m && m[1] && /[a-záéíóúñ]/i.test(m[1])) return { ...ITEM_VACIO, producto: m[1].trim(), cantidad: m[2], unidad: (m[3] || "").toLowerCase() };
  return { ...ITEM_VACIO, producto: s };
};

const parseXlsx = async (file) => {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: false });
  const clean = aoa.filter((r) => r.some((c) => cellToStr(c) !== ""));
  if (!clean.length) return { headers: [], rows: [] };
  const headers = clean[0].map(cellToStr);
  const w = headers.length;
  const rows = clean.slice(1).map((r) => Array.from({ length: w }, (_, i) => cellToStr(r[i])));
  return { headers, rows };
};
const detectDelim = (line) => {
  const c = { "\t": (line.match(/\t/g) || []).length, ";": (line.match(/;/g) || []).length, ",": (line.match(/,/g) || []).length };
  const best = Object.entries(c).sort((a, b) => b[1] - a[1])[0];
  return best[1] > 0 ? best[0] : null;
};
const parseTxt = (text) => {
  const lines = text.split(/\r?\n/).filter((l) => l.trim() !== "");
  if (!lines.length) return { headers: [], rows: [] };
  const delim = detectDelim(lines[0]);
  const split = (l) => (delim ? l.split(delim) : l.split(/\t|\s{2,}/)).map((c) => c.trim());
  const first = split(lines[0]);
  const looksHeader = first.some((c) => guessField(c) !== "ignorar");
  const headers = looksHeader ? first : first.map((_, i) => `Columna ${i + 1}`);
  const dataLines = looksHeader ? lines.slice(1) : lines;
  const w = headers.length;
  const rows = dataLines.map((l) => { const r = split(l); return Array.from({ length: w }, (_, i) => r[i] ?? ""); });
  return { headers, rows };
};

/* ============ UI primitives ============ */
const Card = ({ children, style }) => (
  <div className="rounded-2xl p-4" style={{ background: C.surface, border: `1px solid ${C.line}`, ...style }}>{children}</div>
);
const Btn = ({ children, onClick, disabled, bg = C.gold, fg = C.bg, full, style }) => (
  <button onClick={onClick} disabled={disabled} className={`inline-flex items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-sm font-semibold ${full ? "w-full" : ""}`}
    style={{ background: disabled ? C.line : bg, color: disabled ? C.faint : fg, cursor: disabled ? "not-allowed" : "pointer", transition: "background .2s", ...style }}>{children}</button>
);
const Ghost = ({ children, onClick, color = C.muted }) => (
  <button onClick={onClick} className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium" style={{ color, background: color + "14", cursor: "pointer" }}>{children}</button>
);
const Field = ({ label, ...p }) => (
  <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>{label}
    <input {...p} className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.text }} /></label>
);
const Select = ({ label, children, ...p }) => (
  <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>{label}
    <select {...p} className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.text }}>{children}</select></label>
);
const Badge = ({ estado }) => {
  const e = ESTADOS[estado] || { l: estado, c: C.muted };
  return <span className="rounded-md px-2 py-0.5 text-xs font-semibold" style={{ background: e.c + "22", color: e.c }}>{e.l}</span>;
};
const Tabs = ({ value, onChange, items }) => (
  <div className="mb-4 flex gap-2 rounded-xl p-1" style={{ background: C.surface }}>
    {items.map(([k, l, I]) => {
      const on = value === k;
      return (
        <button key={k} onClick={() => onChange(k)} className="flex flex-1 items-center justify-center gap-1.5 rounded-lg px-3 py-2 text-sm font-semibold"
          style={{ background: on ? C.gold : "transparent", color: on ? C.bg : C.muted, cursor: "pointer" }}>{I && <I size={15} />} {l}</button>
      );
    })}
  </div>
);

/* ============ LOGIN / SETUP ============ */
function LoginPin({ onLogin }) {
  const [pin, setPin] = useState(""); const [err, setErr] = useState(""); const [cargando, setCargando] = useState(false);
  const entrar = async () => {
    if (!pin) return; setCargando(true); setErr("");
    const ok = await onLogin(pin); setCargando(false);
    if (!ok) { setErr("PIN incorrecto."); setPin(""); }
  };
  return (
    <div className="mx-auto max-w-sm pt-10">
      <div className="mb-4 text-center">
        <div className="flex items-center justify-center gap-2 text-xl font-bold" style={{ color: C.gold, letterSpacing: "0.02em" }}><ShoppingCart size={20} /> Coordinación de Compras</div>
        <div className="text-xs" style={{ color: C.faint }}>Solicitud · OC · Recepción</div>
      </div>
      <Card style={{ borderColor: C.gold + "55" }}>
        <div className="mb-3 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><KeyRound size={18} color={C.gold} /> Ingresa tu PIN</div>
        <Field label="PIN" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="••••" inputMode="numeric" type="password" onKeyDown={(e) => e.key === "Enter" && entrar()} />
        {err && <div className="mt-2 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
        <div className="mt-4"><Btn full onClick={entrar} disabled={cargando}>{cargando ? "Entrando…" : "Ingresar"}</Btn></div>
      </Card>
    </div>
  );
}
function SetupAdmin({ onDone }) {
  const [pin, setPin] = useState(""); const [nombre, setNombre] = useState(""); const [err, setErr] = useState("");
  const crear = async () => {
    if (!/^\d{4,6}$/.test(pin)) return setErr("PIN de 4 a 6 dígitos.");
    if (!nombre.trim()) return setErr("Falta tu nombre.");
    await setDoc(doc(colUsr, pin), { nombre: nombre.trim(), rol: "admin" });
    onDone({ pin, nombre: nombre.trim(), rol: "admin" });
  };
  return (
    <div className="mx-auto max-w-sm pt-10">
      <div className="mb-4 text-center">
        <div className="flex items-center justify-center gap-2 text-xl font-bold" style={{ color: C.gold }}><ShoppingCart size={20} /> Coordinación de Compras</div>
        <div className="text-xs" style={{ color: C.faint }}>Primera configuración</div>
      </div>
      <Card style={{ borderColor: C.gold + "55" }}>
        <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><KeyRound size={18} color={C.gold} /> Crea el usuario Admin</div>
        <div className="mb-4 text-xs" style={{ color: C.faint }}>Este será el primer acceso. Desde Admin crearás los PIN de analista, administrativo y bodega.</div>
        <div className="flex flex-col gap-3">
          <Field label="Tu nombre" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: Juan Pablo" />
          <Field label="PIN admin (4-6 díg.)" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="1234" inputMode="numeric" />
        </div>
        {err && <div className="mt-2 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
        <div className="mt-4"><Btn full onClick={crear}>Crear y entrar</Btn></div>
      </Card>
    </div>
  );
}

/* ============ ADMIN · usuarios ============ */
function AdminUsuarios({ miPin }) {
  const [usuarios, setUsuarios] = useState([]);
  const [pin, setPin] = useState(""); const [nombre, setNombre] = useState(""); const [rol, setRol] = useState("analista_compras"); const [err, setErr] = useState("");
  useEffect(() => onSnapshot(colUsr, (s) => setUsuarios(s.docs.map((d) => ({ pin: d.id, ...d.data() })))), []);
  const crear = async () => {
    if (!/^\d{4,6}$/.test(pin)) return setErr("PIN de 4 a 6 dígitos.");
    if (!nombre.trim()) return setErr("Falta el nombre.");
    const ex = await getDoc(doc(colUsr, pin)); if (ex.exists()) return setErr("Ese PIN ya existe.");
    await setDoc(doc(colUsr, pin), { nombre: nombre.trim(), rol }); setPin(""); setNombre(""); setErr("");
  };
  const del = (p) => { if (p === miPin) return; if (window.confirm("¿Eliminar este usuario?")) deleteDoc(doc(colUsr, p)); };
  return (
    <div className="flex flex-col gap-4">
      <Card>
        <div className="mb-3 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><KeyRound size={16} color={C.gold} /> Nuevo usuario (PIN)</div>
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
          <Field label="Nombre" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Nombre" />
          <Field label="PIN (4-6 díg.)" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))} placeholder="1234" inputMode="numeric" />
          <Select label="Rol" value={rol} onChange={(e) => setRol(e.target.value)}>{ROLES.map((r) => <option key={r.k} value={r.k}>{r.l}</option>)}</Select>
        </div>
        {err && <div className="mt-2 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
        <div className="mt-3"><Btn onClick={crear}>Crear usuario</Btn></div>
      </Card>
      <Card>
        <div className="mb-3 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><Users size={16} color={C.gold} /> Usuarios ({usuarios.length})</div>
        <div className="flex flex-col gap-2">
          {usuarios.length === 0 && <div className="text-xs" style={{ color: C.faint }}>Aún no hay usuarios.</div>}
          {usuarios.sort((a, b) => (a.rol > b.rol ? 1 : -1)).map((u) => {
            const m = rolMeta(u.rol); const I = m.i;
            return (
              <div key={u.pin} className="flex items-center justify-between rounded-xl px-3 py-2" style={{ background: C.surface2 }}>
                <div className="flex items-center gap-2"><I size={15} color={C.gold} />
                  <div><div className="text-sm" style={{ color: C.text }}>{u.nombre} {u.pin === miPin && <span style={{ color: C.gold }}>(tú)</span>}</div>
                    <div className="text-xs" style={{ color: C.faint }}>PIN {u.pin} · {m.l}</div></div></div>
                {u.pin !== miPin && <button onClick={() => del(u.pin)} className="rounded-lg p-1.5" style={{ color: C.clay, cursor: "pointer" }}><Trash2 size={14} /></button>}
              </div>
            );
          })}
        </div>
      </Card>
    </div>
  );
}

/* ============ Editor de ítems (previsualización + corrección) ============ */
function ItemsEditor({ items, setItems }) {
  const upd = (i, k, v) => setItems(items.map((it, idx) => (idx === i ? { ...it, [k]: v } : it)));
  const del = (i) => setItems(items.filter((_, idx) => idx !== i));
  const add = () => setItems([...items, { ...ITEM_VACIO }]);
  const th = { color: C.faint, fontWeight: 600 };
  const inp = { background: C.surface2, border: `1px solid ${C.line}`, color: C.text };
  const total = totalItems(items);
  return (
    <div>
      <div className="overflow-x-auto rounded-xl" style={{ border: `1px solid ${C.line}` }}>
        <table className="w-full text-sm" style={{ minWidth: 920, borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.surface2 }}>
            <th className="px-2 py-2 text-left text-xs" style={{ ...th, width: 34 }}>#</th>
            {COLMETA.map((c) => <th key={c.k} className="px-2 py-2 text-left text-xs" style={{ ...th, width: c.w || undefined }}>{c.l}</th>)}
            <th style={{ width: 34 }}></th>
          </tr></thead>
          <tbody>
            {items.map((it, i) => {
              const vacio = !it.producto.trim();
              return (
                <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
                  <td className="px-2 py-1 text-xs" style={{ color: vacio ? C.clay : C.faint }}>{i + 1}</td>
                  {COLMETA.map((c) => (
                    <td key={c.k} className="px-1 py-1">
                      <input value={it[c.k] || ""} onChange={(e) => upd(i, c.k, e.target.value)}
                        className="w-full rounded-md px-2 py-1 text-sm outline-none" style={{ ...inp, borderColor: c.k === "producto" && vacio ? C.clay : C.line }}
                        inputMode={c.k === "cantidad" || c.k === "costo" ? "decimal" : undefined} />
                    </td>
                  ))}
                  <td className="px-1"><button onClick={() => del(i)} className="rounded-md p-1" style={{ color: C.clay, cursor: "pointer" }}><X size={14} /></button></td>
                </tr>
              );
            })}
            {items.length === 0 && <tr><td colSpan={COLMETA.length + 2} className="px-3 py-4 text-center text-xs" style={{ color: C.faint }}>Sin ítems.</td></tr>}
          </tbody>
        </table>
      </div>
      <div className="mt-2 flex items-center justify-between">
        <Ghost onClick={add} color={C.gold}><Plus size={13} /> Agregar fila</Ghost>
        <div className="flex items-center gap-3 text-xs">
          <span style={{ color: C.faint }}>{items.filter((it) => it.producto.trim()).length} ítem(s)</span>
          {total > 0 && <span className="font-semibold" style={{ color: C.gold }}>Total estimado: {montoCLP(total)}</span>}
        </div>
      </div>
    </div>
  );
}

/* ============ Nueva solicitud (Excel/TXT + Imagen/OCR) ============ */
function plantilla() {
  const ws = XLSX.utils.json_to_sheet([
    { producto: "Detergente industrial 5L", cantidad: 10, unidad: "bidón", proveedor: "Distribuidora XYZ", fecha_requerida: "2026-09-30", observaciones: "" },
    { producto: "Guantes nitrilo talla M", cantidad: 20, unidad: "caja", proveedor: "", fecha_requerida: "", observaciones: "Urgente" },
  ]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Solicitud");
  XLSX.writeFile(wb, "plantilla_solicitud_compras.xlsx");
}

function NuevaSolicitud({ session, onListo }) {
  const [step, setStep] = useState("upload"); // upload | map | ocr | review
  const [origen, setOrigen] = useState("archivo"); // archivo | imagen
  const [file, setFile] = useState(null);
  const [parsed, setParsed] = useState({ headers: [], rows: [] });
  const [mapping, setMapping] = useState([]);
  const [items, setItems] = useState([]);
  const [hdr, setHdr] = useState({ proveedor: "", fechaRequerida: "", observaciones: "" });
  const [err, setErr] = useState(""); const [aviso, setAviso] = useState(""); const [guardando, setGuardando] = useState(false);
  // imagen / OCR
  const [imgPrev, setImgPrev] = useState("");       // objectURL para vista previa
  const [imgEmbed, setImgEmbed] = useState(null);   // {nombre,tipo,size,data,tooBig} comprimido para guardar
  const [ocrText, setOcrText] = useState("");
  const [ocrProg, setOcrProg] = useState(0);
  const [ocrRun, setOcrRun] = useState(false);
  const fileRef = useRef(null); const imgRef = useRef(null);
  const archivoGrande = origen === "archivo" && file && file.size > MAX_ARCHIVO;

  const reset = () => { setFile(null); setParsed({ headers: [], rows: [] }); setMapping([]); setItems([]); setImgPrev(""); setImgEmbed(null); setOcrText(""); setOcrProg(0); setErr(""); setAviso(""); };

  const cargarArchivo = async (f) => {
    setErr(""); setAviso(""); if (!f) return;
    try {
      const ext = f.name.split(".").pop().toLowerCase();
      let p;
      if (["xlsx", "xls", "xlsm", "csv"].includes(ext)) p = await parseXlsx(f);
      else if (["txt", "tsv"].includes(ext)) p = parseTxt(await f.text());
      else return setErr("Formato no soportado. Usa Excel/TXT, o el botón de Imagen para OCR.");
      if (!p.headers.length || !p.rows.length) return setErr("No pude leer filas del archivo. Revisa que tenga encabezados y datos.");
      setOrigen("archivo"); setFile(f); setParsed(p); setMapping(p.headers.map((h) => guessField(h))); setStep("map");
    } catch (e) { setErr("Error al leer el archivo: " + (e?.message || e)); }
  };

  const cargarImagen = async (f) => {
    setErr(""); setAviso(""); if (!f) return;
    if (!f.type.startsWith("image/")) return setErr("Selecciona una imagen (foto o recorte de pantalla).");
    setOrigen("imagen"); setFile(f); setImgPrev(URL.createObjectURL(f)); setStep("ocr"); setOcrRun(true); setOcrProg(0); setOcrText("");
    try {
      const emb = await comprimirImagen(f); setImgEmbed(emb);
      const { data } = await Tesseract.recognize(f, "spa", { logger: (m) => { if (m.status === "recognizing text") setOcrProg(Math.round(m.progress * 100)); } });
      setOcrText((data?.text || "").trim());
    } catch (e) { setErr("No se pudo procesar el OCR: " + (e?.message || e)); }
    finally { setOcrRun(false); }
  };

  const aplicarMapeo = () => {
    const its = parsed.rows.map((r) => {
      const it = { ...ITEM_VACIO };
      parsed.headers.forEach((_, i) => { const f = mapping[i]; if (f && f !== "ignorar" && !it[f]) it[f] = r[i] || ""; });
      return it;
    }).filter((it) => it.producto.trim() || it.cantidad.trim());
    if (!its.length) return setErr("Ninguna fila quedó con producto. Ajusta el mapeo de columnas.");
    setItems(its); setErr(""); setStep("review");
  };

  const interpretarOCR = () => {
    const its = ocrText.split(/\r?\n/).map(parseLineaOCR).filter((it) => it && it.producto.trim());
    if (!its.length) return setErr("No se detectaron líneas con productos. Edita el texto o agrega filas manualmente.");
    setItems(its); setErr(""); setStep("review");
  };

  const guardar = async (enviar) => {
    const validos = items.filter((it) => it.producto.trim());
    if (!validos.length) return setErr("Necesitas al menos un ítem con producto.");
    setGuardando(true); setErr(""); setAviso("");
    try {
      let archivo = null, archivoOmitido = false;
      if (origen === "imagen") {
        if (imgEmbed && !imgEmbed.tooBig) archivo = { nombre: imgEmbed.nombre, tipo: imgEmbed.tipo, size: imgEmbed.size, data: imgEmbed.data };
        else archivoOmitido = true;
      } else if (file && !archivoGrande) {
        try { archivo = { nombre: file.name, tipo: file.type || "", size: file.size, data: await fileToB64(file) }; }
        catch (e) { archivoOmitido = true; }
      } else if (archivoGrande) archivoOmitido = true;

      const snap = await getDocs(colSol);
      const folio = snap.docs.reduce((mx, d) => Math.max(mx, Number(d.data().folio) || 0), 0) + 1;
      const por = { pin: session.pin, nombre: session.nombre };
      const estado = enviar ? "ENVIADA" : "BORRADOR";
      const historial = [{ accion: "creada", por, at: Date.now(), detalle: origen === "imagen" ? "OCR imagen" : "archivo" }];
      if (enviar) historial.push({ accion: "enviada", por, at: Date.now() });
      const base = {
        folio, estado,
        items: validos.map((it) => ({ ...it, cantidad: String(it.cantidad).trim() })),
        proveedor: hdr.proveedor.trim(), fechaRequerida: hdr.fechaRequerida.trim(), observaciones: hdr.observaciones.trim(),
        origen, ocNumero: "", recepcionProgramada: null, recepcion: null,
        creadoPor: por, creadoAt: serverTimestamp(), historial,
      };
      const nueva = doc(colSol);
      try { await setDoc(nueva, { ...base, archivo }); }
      catch (e2) { if (archivo) { await setDoc(nueva, { ...base, archivo: null }); archivoOmitido = true; } else throw e2; }

      if (archivoOmitido) { setAviso("Solicitud guardada. El original no se adjuntó por tamaño (máx. 700 KB); los ítems quedaron guardados."); setGuardando(false); setTimeout(() => onListo(), 2200); }
      else onListo();
    } catch (e) { setGuardando(false); setErr("No se pudo guardar: " + (e?.message || e)); }
  };

  return (
    <Card>
      {step === "upload" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><Upload size={16} color={C.gold} /> Nueva solicitud</div>
          <div className="mb-4 text-xs" style={{ color: C.faint }}>Carga un Excel/TXT, o una imagen (foto o recorte) para leerla con OCR. En todos los casos revisas y corriges antes de guardar.</div>
          <div className="flex flex-wrap items-center gap-2">
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.xlsm,.csv,.txt,.tsv" className="hidden" onChange={(e) => cargarArchivo(e.target.files?.[0])} />
            <input ref={imgRef} type="file" accept="image/*" className="hidden" onChange={(e) => cargarImagen(e.target.files?.[0])} />
            <Btn onClick={() => fileRef.current?.click()}><FileSpreadsheet size={15} /> Excel / TXT</Btn>
            <Btn onClick={() => imgRef.current?.click()} bg={C.info} fg={C.bg}><ImageIcon size={15} /> Imagen (OCR)</Btn>
            <Ghost onClick={plantilla} color={C.info}><Download size={13} /> Descargar plantilla</Ghost>
          </div>
          {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
        </>
      )}

      {step === "map" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><FileText size={16} color={C.gold} /> Interpretar columnas</div>
          <div className="mb-3 text-xs" style={{ color: C.faint }}>{file?.name} · {parsed.rows.length} fila(s). Asigna cada columna a un campo. Ajusté un mapeo automático; corrígelo si hace falta.</div>
          <div className="flex flex-col gap-2">
            {parsed.headers.map((h, i) => (
              <div key={i} className="grid grid-cols-1 items-center gap-2 rounded-xl px-3 py-2 sm:grid-cols-2" style={{ background: C.surface2 }}>
                <div className="truncate text-sm" style={{ color: C.text }}>
                  <span style={{ color: C.faint }}>Columna {i + 1}:</span> {h || <span style={{ color: C.faint }}>(sin nombre)</span>}
                  <div className="truncate text-xs" style={{ color: C.faint }}>ej: {parsed.rows.slice(0, 2).map((r) => r[i]).filter(Boolean).join(" · ") || "—"}</div>
                </div>
                <select value={mapping[i]} onChange={(e) => setMapping(mapping.map((m, idx) => (idx === i ? e.target.value : m)))}
                  className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.bg, border: `1px solid ${C.line}`, color: C.text }}>
                  <option value="ignorar">— Ignorar —</option>
                  {FIELD_DEFS.map((f) => <option key={f.k} value={f.k}>{f.l}</option>)}
                </select>
              </div>
            ))}
          </div>
          {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
          <div className="mt-4 flex items-center justify-between">
            <Ghost onClick={() => { reset(); setStep("upload"); }}><ChevronLeft size={13} /> Volver</Ghost>
            <Btn onClick={aplicarMapeo}>Previsualizar <Check size={15} /></Btn>
          </div>
        </>
      )}

      {step === "ocr" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><ScanLine size={16} color={C.gold} /> Texto detectado (OCR)</div>
          <div className="mb-3 text-xs" style={{ color: C.faint }}>El OCR sobre fotos/recortes es aproximado. Revisa y corrige el texto; luego lo interpreto en filas para que ajustes cantidades y unidades.</div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            {imgPrev && <div className="overflow-hidden rounded-xl" style={{ border: `1px solid ${C.line}`, maxHeight: 260 }}><img src={imgPrev} alt="original" style={{ width: "100%", objectFit: "contain", maxHeight: 260 }} /></div>}
            <div>
              {ocrRun ? (
                <div className="flex h-full flex-col items-center justify-center gap-2 rounded-xl p-4 text-center" style={{ background: C.surface2 }}>
                  <ScanLine size={20} color={C.gold} />
                  <div className="text-sm" style={{ color: C.text }}>Procesando OCR… {ocrProg}%</div>
                  <div className="h-1.5 w-full overflow-hidden rounded-full" style={{ background: C.line }}><div style={{ width: `${ocrProg}%`, height: "100%", background: C.gold, transition: "width .2s" }} /></div>
                  <div className="text-xs" style={{ color: C.faint }}>La primera vez descarga el motor (~unos MB).</div>
                </div>
              ) : (
                <textarea value={ocrText} onChange={(e) => setOcrText(e.target.value)} rows={10} placeholder="Aquí aparece el texto detectado; edítalo si hace falta." className="w-full rounded-xl px-3 py-2 text-sm outline-none" style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.text, minHeight: 200 }} />
              )}
            </div>
          </div>
          {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
          <div className="mt-4 flex items-center justify-between">
            <Ghost onClick={() => { reset(); setStep("upload"); }}><ChevronLeft size={13} /> Volver</Ghost>
            <Btn onClick={interpretarOCR} disabled={ocrRun || !ocrText.trim()}>Interpretar líneas <Check size={15} /></Btn>
          </div>
        </>
      )}

      {step === "review" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><ClipboardList size={16} color={C.gold} /> Revisar y confirmar</div>
          <div className="mb-3 flex items-center gap-1.5 text-xs" style={{ color: C.faint }}>
            {origen === "imagen"
              ? <><ImageIcon size={12} /> {file?.name} {imgEmbed && !imgEmbed.tooBig ? `(imagen ${kb(imgEmbed.size)}, se guarda)` : "(imagen muy grande, no se adjuntará)"}</>
              : file && <><Paperclip size={12} /> {file.name} ({kb(file.size)}){archivoGrande ? " — supera 700 KB, no se adjuntará" : " · se guarda como original"}</>}
          </div>
          <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-3">
            <Field label="Proveedor (opcional, cabecera)" value={hdr.proveedor} onChange={(e) => setHdr({ ...hdr, proveedor: e.target.value })} placeholder="Si aplica a toda la solicitud" />
            <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>Fecha requerida (opcional)
              <input type="date" value={hdr.fechaRequerida} onChange={(e) => setHdr({ ...hdr, fechaRequerida: e.target.value })} className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.text, colorScheme: "dark" }} />
            </label>
            <Field label="Observaciones (opcional)" value={hdr.observaciones} onChange={(e) => setHdr({ ...hdr, observaciones: e.target.value })} placeholder="Nota general" />
          </div>
          <ItemsEditor items={items} setItems={setItems} />
          {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
          {aviso && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.amber }}><AlertTriangle size={13} /> {aviso}</div>}
          <div className="mt-4 flex flex-wrap items-center justify-between gap-2">
            <Ghost onClick={() => setStep(origen === "imagen" ? "ocr" : "map")}><ChevronLeft size={13} /> Volver</Ghost>
            <div className="flex gap-2">
              <Btn onClick={() => guardar(false)} disabled={guardando} bg={C.surface2} fg={C.text} style={{ border: `1px solid ${C.line}` }}><Save size={15} /> {guardando ? "Guardando…" : "Guardar borrador"}</Btn>
              <Btn onClick={() => guardar(true)} disabled={guardando}><Send size={15} /> {guardando ? "Guardando…" : "Guardar y enviar"}</Btn>
            </div>
          </div>
        </>
      )}
    </Card>
  );
}

/* ============ Detalle / edición de una solicitud ============ */
function SolicitudDetalle({ sol, session, onClose }) {
  const esBorrador = sol.estado === "BORRADOR";
  const [items, setItems] = useState(sol.items || []);
  const [hdr, setHdr] = useState({ proveedor: sol.proveedor || "", fechaRequerida: sol.fechaRequerida || "", observaciones: sol.observaciones || "" });
  const [msg, setMsg] = useState("");
  const por = { pin: session.pin, nombre: session.nombre };
  const esImg = (sol.archivo?.tipo || "").startsWith("image/");

  const guardarCambios = async () => {
    const validos = items.filter((it) => it.producto.trim());
    if (!validos.length) return setMsg("Debe quedar al menos un ítem con producto.");
    await updateDoc(doc(colSol, sol.id), {
      items: validos.map((it) => ({ ...it, cantidad: String(it.cantidad).trim() })),
      proveedor: hdr.proveedor.trim(), fechaRequerida: hdr.fechaRequerida.trim(), observaciones: hdr.observaciones.trim(),
      historial: arrayUnion({ accion: "editada", por, at: Date.now() }),
    });
    setMsg("Cambios guardados.");
  };
  const enviar = async () => {
    if (!window.confirm("¿Enviar la solicitud al administrativo?")) return;
    await updateDoc(doc(colSol, sol.id), { estado: "ENVIADA", historial: arrayUnion({ accion: "enviada", por, at: Date.now() }) });
    onClose();
  };
  const eliminar = async () => {
    if (!window.confirm("¿Eliminar este borrador? No se puede deshacer.")) return;
    await deleteDoc(doc(colSol, sol.id)); onClose();
  };

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto p-3" style={{ background: "#000000aa" }}>
      <div className="mt-6 w-full max-w-3xl rounded-2xl p-4" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
        <div className="mb-3 flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>Solicitud #{sol.folio} <Badge estado={sol.estado} /></div>
          <button onClick={onClose} className="rounded-lg p-1.5" style={{ color: C.muted, cursor: "pointer" }}><X size={16} /></button>
        </div>
        <div className="mb-3 text-xs" style={{ color: C.faint }}>Creada por {sol.creadoPor?.nombre} · {fechaHora(sol.creadoAt)}{sol.origen === "imagen" ? " · desde imagen (OCR)" : ""}</div>

        {sol.archivo?.data ? (
          <div className="mb-3 flex flex-col gap-2">
            {esImg && <div className="overflow-hidden rounded-xl" style={{ border: `1px solid ${C.line}`, maxHeight: 220 }}><img src={`data:${sol.archivo.tipo};base64,${sol.archivo.data}`} alt="original" style={{ width: "100%", objectFit: "contain", maxHeight: 220 }} /></div>}
            <button onClick={() => descargarArchivo(sol.archivo)} className="inline-flex w-fit items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium" style={{ background: C.info + "1c", color: C.info, cursor: "pointer" }}>
              <Download size={13} /> Descargar original ({sol.archivo.nombre})
            </button>
          </div>
        ) : <div className="mb-3 text-xs" style={{ color: C.faint }}>Sin archivo original adjunto.</div>}

        {esBorrador ? (
          <>
            <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-3">
              <Field label="Proveedor (cabecera)" value={hdr.proveedor} onChange={(e) => setHdr({ ...hdr, proveedor: e.target.value })} />
              <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>Fecha requerida
                <input type="date" value={hdr.fechaRequerida} onChange={(e) => setHdr({ ...hdr, fechaRequerida: e.target.value })} className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.text, colorScheme: "dark" }} />
              </label>
              <Field label="Observaciones" value={hdr.observaciones} onChange={(e) => setHdr({ ...hdr, observaciones: e.target.value })} />
            </div>
            <ItemsEditor items={items} setItems={setItems} />
            {msg && <div className="mt-3 text-xs" style={{ color: C.teal }}>{msg}</div>}
            <div className="mt-4 flex flex-wrap items-center justify-between gap-2">
              <Ghost onClick={eliminar} color={C.clay}><Trash2 size={13} /> Eliminar borrador</Ghost>
              <div className="flex gap-2">
                <Btn onClick={guardarCambios} bg={C.surface2} fg={C.text} style={{ border: `1px solid ${C.line}` }}><Save size={15} /> Guardar cambios</Btn>
                <Btn onClick={enviar}><Send size={15} /> Enviar</Btn>
              </div>
            </div>
          </>
        ) : (
          <ItemsTabla items={sol.items} />
        )}
      </div>
    </div>
  );
}

/* ============ ANALISTA ============ */
function AnalistaView({ session }) {
  const [tab, setTab] = useState("lista");
  const [sols, setSols] = useState([]);
  const [q, setQ] = useState("");
  const [abierta, setAbierta] = useState(null);
  useEffect(() => onSnapshot(colSol, (s) => setSols(s.docs.map((d) => ({ id: d.id, ...d.data() })))), []);

  const mias = useMemo(() => sols
    .filter((s) => s.creadoPor?.pin === session.pin)
    .filter((s) => !q.trim() || String(s.folio).includes(q) || (s.items || []).some((it) => norm(it.producto).includes(norm(q)) || norm(it.codigo).includes(norm(q))) || norm(s.proveedor).includes(norm(q)))
    .sort((a, b) => ms(b.creadoAt) - ms(a.creadoAt)), [sols, session.pin, q]);
  const stats = useMemo(() => {
    const m = sols.filter((s) => s.creadoPor?.pin === session.pin);
    return { borr: m.filter((s) => s.estado === "BORRADOR").length, env: m.filter((s) => s.estado === "ENVIADA").length, tot: m.length };
  }, [sols, session.pin]);

  return (
    <div className="flex flex-col gap-1">
      <div className="mb-3 flex items-center gap-3 text-xs">
        <span style={{ color: C.faint }}>{stats.borr} borrador(es)</span>
        <span style={{ color: C.info }}>{stats.env} enviada(s)</span>
        <span style={{ color: C.muted }}>{stats.tot} total</span>
      </div>
      <Tabs value={tab} onChange={setTab} items={[["lista", "Mis solicitudes", ClipboardList], ["nueva", "Nueva", Plus]]} />

      {tab === "nueva" && <NuevaSolicitud session={session} onListo={() => setTab("lista")} />}

      {tab === "lista" && (
        <div className="flex flex-col gap-3">
          <div className="flex items-center gap-2 rounded-xl px-3 py-2" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
            <Search size={15} color={C.faint} />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por folio, producto o proveedor" className="w-full bg-transparent text-sm outline-none" style={{ color: C.text }} />
          </div>
          {mias.length === 0 && <Card><div className="text-center text-sm" style={{ color: C.faint }}>No tienes solicitudes todavía. Crea una en “Nueva”.</div></Card>}
          {mias.map((s) => {
            const nItems = (s.items || []).length;
            return (
              <button key={s.id} onClick={() => setAbierta(s)} className="text-left" style={{ cursor: "pointer" }}>
                <Card>
                  <div className="flex items-center justify-between gap-3">
                    <div>
                      <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>#{s.folio} <Badge estado={s.estado} /></div>
                      <div className="mt-0.5 text-xs" style={{ color: C.faint }}>{nItems} ítem(s){s.proveedor ? ` · ${s.proveedor}` : ""} · {fecha(s.creadoAt)}{s.archivo ? " · con original" : ""}{s.ocNumero ? ` · OC ${s.ocNumero}` : ""}</div>
                    </div>
                    <div className="text-xs" style={{ color: C.faint }}>{s.estado === "BORRADOR" ? "Editar →" : "Ver →"}</div>
                  </div>
                </Card>
              </button>
            );
          })}
        </div>
      )}
      {abierta && <SolicitudDetalle sol={sols.find((s) => s.id === abierta.id) || abierta} session={session} onClose={() => setAbierta(null)} />}
    </div>
  );
}

/* ============ tabla de ítems (solo lectura) ============ */
function ItemsTabla({ items }) {
  const total = totalItems(items);
  return (
    <div>
      <div className="overflow-x-auto rounded-xl" style={{ border: `1px solid ${C.line}` }}>
        <table className="w-full text-sm" style={{ minWidth: 720 }}>
          <thead><tr style={{ background: C.surface2, color: C.faint }}>
            {COLMETA.map((c) => <th key={c.k} className="px-2 py-2 text-left text-xs">{c.l}</th>)}
          </tr></thead>
          <tbody>{(items || []).map((it, i) => (
            <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
              {COLMETA.map((c) => (
                <td key={c.k} className="px-2 py-1.5" style={{ color: c.k === "producto" ? C.text : C.muted }}>{it[c.k] || ""}</td>
              ))}
            </tr>
          ))}</tbody>
        </table>
      </div>
      {total > 0 && <div className="mt-2 text-right text-xs font-semibold" style={{ color: C.gold }}>Total estimado: {montoCLP(total)}</div>}
    </div>
  );
}

/* ============ ADMINISTRATIVO · detalle (registrar OC + programar recepción) ============ */
function AdminDetalle({ sol, session, onClose }) {
  const [proveedor, setProveedor] = useState(sol.proveedor || (sol.items || []).map((i) => i.proveedor).find(Boolean) || "");
  const [oc, setOc] = useState(sol.ocNumero || "");
  const [fecha, setFecha] = useState(sol.recepcionProgramada?.fecha || "");
  const [hora, setHora] = useState(sol.recepcionProgramada?.hora || "");
  const [err, setErr] = useState(""); const [guardando, setGuardando] = useState(false);
  const por = { pin: session.pin, nombre: session.nombre };
  const yaProgramada = sol.estado === "PENDIENTE_RECEPCION" || sol.estado === "OC_GENERADA";
  const esImg = (sol.archivo?.tipo || "").startsWith("image/");

  const registrar = async () => {
    if (!oc.trim()) return setErr("Ingresa el N° de OC generado en el ERP.");
    if (!proveedor.trim()) return setErr("Ingresa el proveedor (aparece en la agenda de bodega).");
    if (!fecha) return setErr("Selecciona la fecha de recepción.");
    if (!hora) return setErr("Selecciona la hora de recepción.");
    const at = new Date(`${fecha}T${hora}`).getTime();
    if (!at || isNaN(at)) return setErr("Fecha u hora inválida.");
    setGuardando(true); setErr("");
    try {
      await updateDoc(doc(colSol, sol.id), {
        estado: "PENDIENTE_RECEPCION",
        ocNumero: oc.trim(),
        proveedor: proveedor.trim(),
        recepcionProgramada: { fecha, hora, at },
        historial: arrayUnion({ accion: yaProgramada ? "reprogramada" : "oc_registrada", por, at: Date.now(), detalle: `OC ${oc.trim()} · recepción ${fecha} ${hora}` }),
      });
      onClose();
    } catch (e) { setGuardando(false); setErr("No se pudo guardar: " + (e?.message || e)); }
  };

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto p-3" style={{ background: "#000000aa" }}>
      <div className="mt-6 w-full max-w-3xl rounded-2xl p-4" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
        <div className="mb-3 flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>Solicitud #{sol.folio} <Badge estado={sol.estado} /></div>
          <button onClick={onClose} className="rounded-lg p-1.5" style={{ color: C.muted, cursor: "pointer" }}><X size={16} /></button>
        </div>
        <div className="mb-3 text-xs" style={{ color: C.faint }}>Enviada por {sol.creadoPor?.nombre} · {fechaHora(sol.creadoAt)}{sol.origen === "imagen" ? " · desde imagen (OCR)" : ""}</div>

        {sol.archivo?.data ? (
          <div className="mb-3 flex flex-col gap-2">
            {esImg && <div className="overflow-hidden rounded-xl" style={{ border: `1px solid ${C.line}`, maxHeight: 200 }}><img src={`data:${sol.archivo.tipo};base64,${sol.archivo.data}`} alt="original" style={{ width: "100%", objectFit: "contain", maxHeight: 200 }} /></div>}
            <button onClick={() => descargarArchivo(sol.archivo)} className="inline-flex w-fit items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium" style={{ background: C.info + "1c", color: C.info, cursor: "pointer" }}>
              <Download size={13} /> Descargar original ({sol.archivo.nombre})
            </button>
          </div>
        ) : <div className="mb-3 text-xs" style={{ color: C.faint }}>Sin archivo original adjunto.</div>}

        <div className="mb-3"><ItemsTabla items={sol.items} /></div>
        {sol.observaciones && <div className="mb-3 text-xs" style={{ color: C.muted }}>Obs.: {sol.observaciones}</div>}

        <div className="rounded-xl p-3" style={{ background: C.surface2, border: `1px solid ${C.gold}44` }}>
          <div className="mb-3 flex items-center gap-2 text-sm font-semibold" style={{ color: C.gold }}><FileCheck2 size={15} /> Registrar OC y programar recepción</div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>N° de OC (del ERP)
              <div className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background: C.bg, border: `1px solid ${C.line}` }}>
                <Hash size={14} color={C.faint} /><input value={oc} onChange={(e) => setOc(e.target.value)} placeholder="Ej: 4587" className="w-full bg-transparent text-sm outline-none" style={{ color: C.text }} /></div>
            </label>
            <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>Proveedor
              <div className="flex items-center gap-2 rounded-lg px-3 py-2" style={{ background: C.bg, border: `1px solid ${C.line}` }}>
                <Building2 size={14} color={C.faint} /><input value={proveedor} onChange={(e) => setProveedor(e.target.value)} placeholder="Proveedor de la OC" className="w-full bg-transparent text-sm outline-none" style={{ color: C.text }} /></div>
            </label>
            <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>Fecha de recepción
              <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.bg, border: `1px solid ${C.line}`, color: C.text, colorScheme: "dark" }} />
            </label>
            <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted }}>Hora de recepción
              <input type="time" value={hora} onChange={(e) => setHora(e.target.value)} className="rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.bg, border: `1px solid ${C.line}`, color: C.text, colorScheme: "dark" }} />
            </label>
          </div>
          {err && <div className="mt-2 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
          <div className="mt-3 flex justify-end">
            <Btn onClick={registrar} disabled={guardando}><Send size={15} /> {guardando ? "Guardando…" : yaProgramada ? "Actualizar y reenviar a bodega" : "Registrar OC y enviar a bodega"}</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============ ADMINISTRATIVO ============ */
function AdministrativoView({ session }) {
  const [tab, setTab] = useState("pendientes");
  const [sols, setSols] = useState([]);
  const [q, setQ] = useState("");
  const [abierta, setAbierta] = useState(null);
  useEffect(() => onSnapshot(colSol, (s) => setSols(s.docs.map((d) => ({ id: d.id, ...d.data() })))), []);

  const pendientes = useMemo(() => sols.filter((s) => s.estado === "ENVIADA").sort((a, b) => ms(a.creadoAt) - ms(b.creadoAt)), [sols]);
  const programadas = useMemo(() => sols.filter((s) => s.estado === "PENDIENTE_RECEPCION" || s.estado === "OC_GENERADA" || s.estado === "RECEPCION_PARCIAL").sort((a, b) => ms(a.recepcionProgramada?.at) - ms(b.recepcionProgramada?.at)), [sols]);
  const filtrar = (arr) => arr.filter((s) => !q.trim() || String(s.folio).includes(q) || norm(s.proveedor).includes(norm(q)) || String(s.ocNumero || "").includes(q) || (s.items || []).some((it) => norm(it.producto).includes(norm(q)) || norm(it.codigo).includes(norm(q))));
  const lista = tab === "pendientes" ? filtrar(pendientes) : filtrar(programadas);

  return (
    <div className="flex flex-col gap-1">
      <div className="mb-3 flex items-center gap-3 text-xs">
        <span style={{ color: C.info }}>{pendientes.length} pendiente(s)</span>
        <span style={{ color: C.gold }}>{programadas.length} programada(s)</span>
      </div>
      <Tabs value={tab} onChange={setTab} items={[["pendientes", "Pendientes", Inbox], ["programadas", "Programadas", CalendarClock]]} />

      <div className="mb-3 flex items-center gap-2 rounded-xl px-3 py-2" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
        <Search size={15} color={C.faint} />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por folio, producto, proveedor u OC" className="w-full bg-transparent text-sm outline-none" style={{ color: C.text }} />
      </div>

      <div className="flex flex-col gap-3">
        {lista.length === 0 && <Card><div className="text-center text-sm" style={{ color: C.faint }}>{tab === "pendientes" ? "No hay solicitudes pendientes por procesar." : "No hay recepciones programadas."}</div></Card>}
        {lista.map((s) => {
          const nItems = (s.items || []).length;
          return (
            <button key={s.id} onClick={() => setAbierta(s)} className="text-left" style={{ cursor: "pointer" }}>
              <Card>
                <div className="flex items-center justify-between gap-3">
                  <div>
                    <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>#{s.folio} <Badge estado={s.estado} />{s.origen === "imagen" && <ImageIcon size={12} color={C.faint} />}</div>
                    <div className="mt-0.5 text-xs" style={{ color: C.faint }}>
                      {nItems} ítem(s){s.proveedor ? ` · ${s.proveedor}` : ""} · de {s.creadoPor?.nombre}
                      {tab === "programadas" && s.recepcionProgramada?.at ? ` · recep. ${fechaHora(s.recepcionProgramada.at)} · OC ${s.ocNumero}` : ""}
                    </div>
                  </div>
                  <div className="text-xs" style={{ color: C.faint }}>{tab === "pendientes" ? "Procesar →" : "Editar →"}</div>
                </div>
              </Card>
            </button>
          );
        })}
      </div>
      {abierta && <AdminDetalle sol={sols.find((s) => s.id === abierta.id) || abierta} session={session} onClose={() => setAbierta(null)} />}
    </div>
  );
}

/* ============ BODEGA · confirmar recepción ============ */
const fechaAgenda = (at) => {
  const d = new Date(at);
  return d.toLocaleDateString("es-CL", { weekday: "short", day: "2-digit", month: "2-digit" }) + " · " + d.toLocaleTimeString("es-CL", { hour: "2-digit", minute: "2-digit" });
};

function RecepcionModal({ sol, session, onClose }) {
  const [obs, setObs] = useState("");
  const [accion, setAccion] = useState(null); // null | "parcial" | "completa"
  const [guardando, setGuardando] = useState(false);
  const [err, setErr] = useState("");
  const at = sol.recepcionProgramada?.at;
  const parcialPrevia = sol.estado === "RECEPCION_PARCIAL" ? sol.recepcion : null;

  const ejecutar = async () => {
    setGuardando(true); setErr("");
    try {
      const por = { pin: session.pin, nombre: session.nombre };
      const tipo = accion; // "parcial" | "completa"
      const estado = tipo === "completa" ? "RECEPCIONADA" : "RECEPCION_PARCIAL";
      const nota = obs.trim();
      const evento = { accion: tipo === "completa" ? "recepcionada" : "recepcion_parcial", por, at: Date.now(), ...(nota ? { detalle: nota } : {}) };
      await updateDoc(doc(colSol, sol.id), {
        estado,
        recepcion: { at: Date.now(), por, tipo, observaciones: nota },
        historial: arrayUnion(evento),
      });
      onClose();
    } catch (e) { setGuardando(false); setErr("No se pudo confirmar: " + (e?.message || e)); }
  };

  const labelAccion = accion === "completa" ? "completa" : "parcial";
  const colorAccion = accion === "completa" ? C.teal : C.clay;

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto p-3" style={{ background: "#000000aa" }}>
      <div className="mt-10 w-full max-w-md rounded-2xl p-5" style={{ background: C.surface, border: `1px solid ${C.gold}55` }}>
        <div className="mb-4 flex items-center justify-between">
          <div className="flex items-center gap-2 text-base font-bold" style={{ color: C.gold }}><Hash size={18} /> OC {sol.ocNumero || "—"}</div>
          <button onClick={onClose} className="rounded-lg p-1.5" style={{ color: C.muted, cursor: "pointer" }}><X size={18} /></button>
        </div>

        <div className="flex flex-col gap-3">
          <div className="flex items-center gap-2 rounded-xl px-3 py-2.5" style={{ background: C.surface2 }}>
            <Building2 size={16} color={C.faint} />
            <div><div className="text-xs" style={{ color: C.faint }}>Proveedor</div><div className="text-sm font-semibold" style={{ color: C.text }}>{sol.proveedor || "—"}</div></div>
          </div>
          <div className="flex items-center gap-2 rounded-xl px-3 py-2.5" style={{ background: C.surface2 }}>
            <CalendarClock size={16} color={C.faint} />
            <div><div className="text-xs" style={{ color: C.faint }}>Fecha / hora programada</div><div className="text-sm font-semibold" style={{ color: C.text }}>{at ? fechaAgenda(at) : "—"}</div></div>
          </div>

          {parcialPrevia && (
            <div className="rounded-xl px-3 py-2.5" style={{ background: C.clay + "14", border: `1px solid ${C.clay}44` }}>
              <div className="flex items-center gap-1.5 text-xs font-semibold" style={{ color: C.clay }}><PackageCheck size={13} /> Recepción parcial previa</div>
              <div className="mt-1 text-xs" style={{ color: C.muted }}>{parcialPrevia.por?.nombre} · {fechaHora(parcialPrevia.at)}{parcialPrevia.observaciones ? ` · ${parcialPrevia.observaciones}` : ""}</div>
            </div>
          )}

          {/* Tarjeta de observaciones */}
          <div className="rounded-xl p-3" style={{ background: C.surface2, border: `1px solid ${C.line}` }}>
            <div className="mb-2 flex items-center gap-1.5 text-xs font-semibold" style={{ color: C.text }}><FileText size={13} color={C.gold} /> Observaciones (opcional)</div>
            <textarea value={obs} onChange={(e) => setObs(e.target.value)} rows={3} placeholder="Anota lo que corresponda: faltantes, diferencias, estado de la mercadería, etc." className="w-full rounded-lg px-3 py-2 text-sm outline-none" style={{ background: C.bg, border: `1px solid ${C.line}`, color: C.text }} />
          </div>

          <div className="text-xs" style={{ color: C.faint }}>La OC se verifica en el ERP. Aquí solo confirmas la recepción.</div>
        </div>

        {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}

        {!accion ? (
          <div className="mt-5 flex gap-2">
            <Btn full onClick={() => setAccion("parcial")} bg={C.clay} fg="#1a0f0b"><PackageCheck size={16} /> Parcial</Btn>
            <Btn full onClick={() => setAccion("completa")} bg={C.teal} fg="#0d1a12"><PackageCheck size={16} /> Completa</Btn>
          </div>
        ) : (
          <div className="mt-5 rounded-xl p-3" style={{ background: colorAccion + "14", border: `1px solid ${colorAccion}55` }}>
            <div className="mb-3 text-center text-sm font-semibold" style={{ color: C.text }}>¿Confirmar recepción {labelAccion} de la OC {sol.ocNumero}?</div>
            <div className="flex gap-2">
              <Btn full onClick={() => setAccion(null)} bg={C.surface2} fg={C.text} style={{ border: `1px solid ${C.line}` }}>Cancelar</Btn>
              <Btn full onClick={ejecutar} disabled={guardando} bg={colorAccion} fg={accion === "completa" ? "#0d1a12" : "#1a0f0b"}><Check size={16} /> {guardando ? "Guardando…" : "Confirmar"}</Btn>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============ BODEGA · agenda de recepciones ============ */
function BodegaView({ session }) {
  const [sols, setSols] = useState([]);
  const [abierta, setAbierta] = useState(null);
  useEffect(() => onSnapshot(colSol, (s) => setSols(s.docs.map((d) => ({ id: d.id, ...d.data() })))), []);

  const hoy0 = new Date(); hoy0.setHours(0, 0, 0, 0);
  const ini = hoy0.getTime(); const fin = ini + 86400000;

  const grupos = useMemo(() => {
    const pend = sols.filter((s) => (s.estado === "PENDIENTE_RECEPCION" || s.estado === "RECEPCION_PARCIAL") && s.recepcionProgramada?.at)
      .sort((a, b) => a.recepcionProgramada.at - b.recepcionProgramada.at);
    return {
      atrasadas: pend.filter((s) => s.recepcionProgramada.at < ini),
      hoy: pend.filter((s) => s.recepcionProgramada.at >= ini && s.recepcionProgramada.at < fin),
      proximas: pend.filter((s) => s.recepcionProgramada.at >= fin),
    };
  }, [sols, ini, fin]);
  const recepHoy = useMemo(() => sols.filter((s) => s.estado === "RECEPCIONADA" && ms(s.recepcion?.at) >= ini && ms(s.recepcion?.at) < fin).length, [sols, ini, fin]);

  const Registro = ({ s, color }) => (
    <button onClick={() => setAbierta(s)} className="w-full text-left" style={{ cursor: "pointer" }}>
      <div className="flex items-center justify-between gap-3 rounded-xl px-4 py-3" style={{ background: C.surface, border: `1px solid ${C.line}`, borderLeft: `4px solid ${color}` }}>
        <div>
          <div className="flex items-center gap-2 text-sm font-bold" style={{ color: C.text }}><Hash size={14} color={color} />OC {s.ocNumero || "—"}{s.estado === "RECEPCION_PARCIAL" && <span className="rounded px-1.5 py-0.5 text-xs" style={{ background: C.clay + "22", color: C.clay }}>Parcial</span>}</div>
          <div className="mt-0.5 text-xs" style={{ color: C.muted }}>{s.proveedor || "—"} · {fechaAgenda(s.recepcionProgramada.at)}</div>
        </div>
        <div className="flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-xs font-semibold" style={{ background: color + "1c", color }}><PackageCheck size={13} /> Recepcionar</div>
      </div>
    </button>
  );

  const Seccion = ({ titulo, Icono, color, items }) => (
    <div>
      <div className="mb-2 flex items-center gap-2 text-sm font-bold" style={{ color }}>
        <Icono size={16} /> {titulo} <span className="rounded-full px-2 text-xs" style={{ background: color + "22" }}>{items.length}</span>
      </div>
      {items.length === 0 ? (
        <div className="rounded-xl px-4 py-3 text-xs" style={{ background: C.surface, border: `1px solid ${C.line}`, color: C.faint }}>Sin recepciones.</div>
      ) : (
        <div className="flex flex-col gap-2">{items.map((s) => <Registro key={s.id} s={s} color={color} />)}</div>
      )}
    </div>
  );

  return (
    <div className="flex flex-col gap-5">
      <div className="flex items-center gap-2 text-base font-bold" style={{ color: C.gold }}><CalendarClock size={18} /> Agenda de recepciones</div>
      <Seccion titulo="ATRASADAS" Icono={AlertTriangle} color={C.clay} items={grupos.atrasadas} />
      <Seccion titulo="HOY" Icono={Clock} color={C.gold} items={grupos.hoy} />
      <Seccion titulo="PRÓXIMAS" Icono={CalendarClock} color={C.info} items={grupos.proximas} />
      {recepHoy > 0 && <div className="flex items-center gap-1.5 text-xs" style={{ color: C.teal }}><PackageCheck size={13} /> {recepHoy} recepcionada(s) hoy</div>}
      {abierta && <RecepcionModal sol={sols.find((s) => s.id === abierta.id) || abierta} session={session} onClose={() => setAbierta(null)} />}
    </div>
  );
}

/* ============ TRAZABILIDAD · línea de tiempo de una solicitud ============ */
const ACCIONES = {
  creada: { l: "Solicitud creada", i: ClipboardList, c: C.info },
  editada: { l: "Solicitud editada", i: FileText, c: C.muted },
  enviada: { l: "Enviada al administrativo", i: Send, c: C.info },
  oc_registrada: { l: "OC registrada · recepción programada", i: FileCheck2, c: C.amber },
  reprogramada: { l: "Recepción reprogramada", i: CalendarClock, c: C.amber },
  recepcion_parcial: { l: "Recepción parcial", i: PackageCheck, c: C.clay },
  recepcionada: { l: "Recepción confirmada", i: PackageCheck, c: C.teal },
};

function TrazaDetalle({ sol, onClose }) {
  const eventos = [...(sol.historial || [])].sort((a, b) => ms(a.at) - ms(b.at));
  const esImg = (sol.archivo?.tipo || "").startsWith("image/");
  const Dato = ({ icon: I, label, valor }) => (
    <div className="flex items-start gap-2">
      <I size={14} color={C.faint} style={{ marginTop: 2 }} />
      <div><div className="text-xs" style={{ color: C.faint }}>{label}</div><div className="text-sm" style={{ color: C.text }}>{valor || "—"}</div></div>
    </div>
  );

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto p-3" style={{ background: "#000000aa" }}>
      <div className="mt-6 w-full max-w-3xl rounded-2xl p-4" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
        <div className="mb-3 flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>Trazabilidad · Solicitud #{sol.folio} <Badge estado={sol.estado} /></div>
          <button onClick={onClose} className="rounded-lg p-1.5" style={{ color: C.muted, cursor: "pointer" }}><X size={16} /></button>
        </div>

        {/* Resumen reconstruido */}
        <div className="mb-4 grid grid-cols-1 gap-3 rounded-xl p-3 sm:grid-cols-2" style={{ background: C.surface2 }}>
          <Dato icon={User} label="Analista (creó)" valor={sol.creadoPor?.nombre} />
          <Dato icon={Paperclip} label="Archivo original" valor={sol.archivo ? sol.archivo.nombre : (sol.origen === "imagen" ? "imagen (no adjuntada)" : "sin archivo")} />
          <Dato icon={Building2} label="Proveedor" valor={sol.proveedor} />
          <Dato icon={Hash} label="N° OC (ERP)" valor={sol.ocNumero} />
          <Dato icon={ShoppingCart} label="Total estimado" valor={totalItems(sol.items) > 0 ? montoCLP(totalItems(sol.items)) : ""} />
          <Dato icon={CalendarClock} label="Recepción programada" valor={sol.recepcionProgramada?.at ? fechaAgenda(sol.recepcionProgramada.at) : ""} />
          <Dato icon={PackageCheck} label="Recepción real" valor={sol.recepcion?.at ? `${fechaHora(sol.recepcion.at)} · ${sol.recepcion.por?.nombre || ""}${sol.recepcion.tipo === "parcial" ? " · parcial" : sol.recepcion.tipo === "completa" ? " · completa" : ""}${sol.recepcion.observaciones ? ` · ${sol.recepcion.observaciones}` : ""}` : ""} />
        </div>

        {sol.archivo?.data && (
          <div className="mb-4 flex flex-col gap-2">
            {esImg && <div className="overflow-hidden rounded-xl" style={{ border: `1px solid ${C.line}`, maxHeight: 180 }}><img src={`data:${sol.archivo.tipo};base64,${sol.archivo.data}`} alt="original" style={{ width: "100%", objectFit: "contain", maxHeight: 180 }} /></div>}
            <button onClick={() => descargarArchivo(sol.archivo)} className="inline-flex w-fit items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium" style={{ background: C.info + "1c", color: C.info, cursor: "pointer" }}>
              <Download size={13} /> Descargar original
            </button>
          </div>
        )}

        {/* Línea de tiempo */}
        <div className="mb-2 flex items-center gap-2 text-sm font-semibold" style={{ color: C.gold }}><History size={15} /> Línea de tiempo</div>
        <div className="flex flex-col">
          {eventos.length === 0 && <div className="text-xs" style={{ color: C.faint }}>Sin eventos registrados.</div>}
          {eventos.map((ev, i) => {
            const a = ACCIONES[ev.accion] || { l: ev.accion, i: Clock, c: C.muted };
            const I = a.i; const ultimo = i === eventos.length - 1;
            return (
              <div key={i} className="flex gap-3">
                <div className="flex flex-col items-center">
                  <div className="flex h-7 w-7 items-center justify-center rounded-full" style={{ background: a.c + "22", border: `1px solid ${a.c}66` }}><I size={14} color={a.c} /></div>
                  {!ultimo && <div style={{ width: 2, flex: 1, background: C.line, minHeight: 14 }} />}
                </div>
                <div className="pb-4">
                  <div className="text-sm font-medium" style={{ color: C.text }}>{a.l}</div>
                  <div className="text-xs" style={{ color: C.faint }}>{ev.por?.nombre || "—"} · {fechaHora(ev.at)}{ev.detalle ? ` · ${ev.detalle}` : ""}</div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function TrazabilidadView() {
  const [sols, setSols] = useState([]);
  const [q, setQ] = useState("");
  const [filtro, setFiltro] = useState("TODAS");
  const [abierta, setAbierta] = useState(null);
  useEffect(() => onSnapshot(colSol, (s) => setSols(s.docs.map((d) => ({ id: d.id, ...d.data() })))), []);

  const lista = useMemo(() => sols
    .filter((s) => filtro === "TODAS" || s.estado === filtro)
    .filter((s) => !q.trim() || String(s.folio).includes(q) || norm(s.proveedor).includes(norm(q)) || String(s.ocNumero || "").includes(q) || norm(s.creadoPor?.nombre).includes(norm(q)) || (s.items || []).some((it) => norm(it.codigo).includes(norm(q))))
    .sort((a, b) => ms(b.creadoAt) - ms(a.creadoAt)), [sols, filtro, q]);

  return (
    <div className="flex flex-col gap-3">
      <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
        <div className="flex items-center gap-2 rounded-xl px-3 py-2 sm:col-span-2" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
          <Search size={15} color={C.faint} />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por folio, proveedor, OC o analista" className="w-full bg-transparent text-sm outline-none" style={{ color: C.text }} />
        </div>
        <select value={filtro} onChange={(e) => setFiltro(e.target.value)} className="rounded-xl px-3 py-2 text-sm outline-none" style={{ background: C.surface, border: `1px solid ${C.line}`, color: C.text }}>
          <option value="TODAS">Todos los estados</option>
          {Object.keys(ESTADOS).map((k) => <option key={k} value={k}>{ESTADOS[k].l}</option>)}
        </select>
      </div>

      {lista.length === 0 && <Card><div className="text-center text-sm" style={{ color: C.faint }}>No hay solicitudes que coincidan.</div></Card>}
      {lista.map((s) => (
        <button key={s.id} onClick={() => setAbierta(s)} className="text-left" style={{ cursor: "pointer" }}>
          <Card>
            <div className="flex items-center justify-between gap-3">
              <div>
                <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>#{s.folio} <Badge estado={s.estado} /></div>
                <div className="mt-0.5 text-xs" style={{ color: C.faint }}>
                  {s.creadoPor?.nombre}{s.proveedor ? ` · ${s.proveedor}` : ""}{s.ocNumero ? ` · OC ${s.ocNumero}` : ""} · {(s.historial || []).length} evento(s)
                </div>
              </div>
              <History size={16} color={C.faint} />
            </div>
          </Card>
        </button>
      ))}
      {abierta && <TrazaDetalle sol={sols.find((s) => s.id === abierta.id) || abierta} onClose={() => setAbierta(null)} />}
    </div>
  );
}

/* ============ ADMIN (pestañas: usuarios + trazabilidad) ============ */
function AdminView({ miPin }) {
  const [tab, setTab] = useState("usuarios");
  return (
    <div>
      <Tabs value={tab} onChange={setTab} items={[["usuarios", "Usuarios", Users], ["traza", "Trazabilidad", History]]} />
      {tab === "usuarios" && <AdminUsuarios miPin={miPin} />}
      {tab === "traza" && <TrazabilidadView />}
    </div>
  );
}

/* ============ Placeholder por rol ============ */
function EnConstruccion({ titulo, etapa, desc, Icono }) {
  return (
    <Card>
      <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><Icono size={18} color={C.gold} /> {titulo}</div>
      <div className="mt-2 text-xs" style={{ color: C.muted }}>{desc}</div>
      <div className="mt-3 inline-flex items-center gap-1.5 rounded-lg px-2.5 py-1 text-xs font-medium" style={{ background: C.amber + "1c", color: C.amber }}>
        <PackageCheck size={13} /> Etapa {etapa} · en construcción
      </div>
    </Card>
  );
}

/* ============ ROOT ============ */
export default function App() {
  const [session, setSession] = useState(() => { try { return JSON.parse(localStorage.getItem("compras_session")) || null; } catch { return null; } });
  const [boot, setBoot] = useState("loading");
  useEffect(() => { getDocs(colUsr).then((s) => setBoot(s.empty ? "setup" : "ready")).catch(() => setBoot("ready")); }, []);
  useEffect(() => { if (session) localStorage.setItem("compras_session", JSON.stringify(session)); else localStorage.removeItem("compras_session"); }, [session]);
  const login = async (pin) => { try { const snap = await getDoc(doc(colUsr, pin)); if (snap.exists()) { setSession({ pin, ...snap.data() }); return true; } } catch (e) {} return false; };

  const wrap = (children) => (
    <div style={{ background: C.bg, color: C.text, minHeight: "100vh", fontFamily: "ui-sans-serif, system-ui, sans-serif" }}>
      <div className="mx-auto max-w-4xl px-4 py-5">{children}</div>
    </div>
  );

  if (boot === "loading") return wrap(<div className="pt-20 text-center text-sm" style={{ color: C.faint }}>Cargando…</div>);
  if (boot === "setup" && !session) return wrap(<SetupAdmin onDone={(u) => { setSession(u); setBoot("ready"); }} />);
  if (!session) return wrap(<LoginPin onLogin={login} />);

  const m = rolMeta(session.rol);
  return wrap(
    <>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2 text-lg font-bold" style={{ color: C.gold, letterSpacing: "0.02em" }}><ShoppingCart size={18} /> Compras · {m.l}</div>
          <div className="text-xs" style={{ color: C.faint }}>{session.nombre} · esta app coordina; el ERP genera la OC y hace la recepción formal</div>
        </div>
        <Ghost onClick={() => setSession(null)}><LogOut size={12} /> Salir</Ghost>
      </div>

      {session.rol === "admin" && <AdminView miPin={session.pin} />}
      {session.rol === "analista_compras" && <AnalistaView session={session} />}
      {session.rol === "administrativo_compras" && <AdministrativoView session={session} />}
      {session.rol === "bodeguero" && <BodegaView session={session} />}
    </>
  );
}

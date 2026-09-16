import React, { useState, useEffect, useMemo, useRef } from "react";
import { db, storage } from "./firebase";
import {
  collection, doc, getDoc, getDocs, setDoc, deleteDoc, updateDoc, onSnapshot, serverTimestamp, arrayUnion,
} from "firebase/firestore";
import { ref as sref, uploadBytes, getDownloadURL, deleteObject } from "firebase/storage";
import * as XLSX from "xlsx";
import {
  KeyRound, AlertTriangle, LogOut, Users, Trash2, ShoppingCart, ClipboardList,
  Truck, FileCheck2, Shield, PackageCheck, Upload, Download, Send, Plus, X, Save,
  FileSpreadsheet, FileText, Check, ChevronLeft, Paperclip, Search,
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
  RECEPCIONADA: { l: "Recepcionada", c: C.teal },
  CERRADA: { l: "Cerrada", c: C.muted },
};

/* ============ helpers ============ */
const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
const ms = (t) => (t == null ? 0 : typeof t === "number" ? t : typeof t.toMillis === "function" ? t.toMillis() : t.seconds ? t.seconds * 1000 : 0);
const fecha = (ts) => { const n = ms(ts); return n ? new Date(n).toLocaleDateString("es-CL", { day: "2-digit", month: "2-digit", year: "2-digit" }) : ""; };
const fechaHora = (ts) => { const n = ms(ts); return n ? new Date(n).toLocaleString("es-CL", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : ""; };
const safeName = (n) => String(n || "archivo").replace(/[^\w.\-]+/g, "_").slice(-80);
const cellToStr = (v) => { if (v == null) return ""; if (v instanceof Date) return v.toISOString().slice(0, 10); return String(v).trim(); };

const ITEM_VACIO = { producto: "", cantidad: "", unidad: "", proveedor: "", fechaRequerida: "", observaciones: "" };
const FIELD_DEFS = [
  { k: "producto", l: "Producto / Descripción", syn: ["producto", "descripcion", "detalle", "item", "articulo", "glosa", "nombre"] },
  { k: "cantidad", l: "Cantidad", syn: ["cantidad", "cant", "qty", "unidades"] },
  { k: "unidad", l: "Unidad", syn: ["unidad", "um", "unidad de medida", "medida", "unit"] },
  { k: "proveedor", l: "Proveedor", syn: ["proveedor", "prov", "vendor"] },
  { k: "fechaRequerida", l: "Fecha requerida", syn: ["fecha requerida", "fecha_requerida", "fecha", "requerida", "fecha entrega", "plazo"] },
  { k: "observaciones", l: "Observaciones", syn: ["observaciones", "observacion", "obs", "nota", "notas", "comentario", "comentarios"] },
];
const COLS = FIELD_DEFS.map((f) => f.k);
const guessField = (header) => {
  const h = norm(header);
  if (!h) return "ignorar";
  for (const f of FIELD_DEFS) if (f.syn.some((s) => h === norm(s))) return f.k;
  for (const f of FIELD_DEFS) if (f.syn.some((s) => h.includes(norm(s)))) return f.k;
  return "ignorar";
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

/* ============ ADMIN ============ */
function AdminView({ miPin }) {
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
  return (
    <div>
      <div className="overflow-x-auto rounded-xl" style={{ border: `1px solid ${C.line}` }}>
        <table className="w-full text-sm" style={{ minWidth: 720, borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.surface2 }}>
            <th className="px-2 py-2 text-left text-xs" style={{ ...th, width: 34 }}>#</th>
            <th className="px-2 py-2 text-left text-xs" style={th}>Producto / Descripción</th>
            <th className="px-2 py-2 text-left text-xs" style={{ ...th, width: 80 }}>Cantidad</th>
            <th className="px-2 py-2 text-left text-xs" style={{ ...th, width: 90 }}>Unidad</th>
            <th className="px-2 py-2 text-left text-xs" style={{ ...th, width: 140 }}>Proveedor</th>
            <th className="px-2 py-2 text-left text-xs" style={{ ...th, width: 120 }}>Fecha req.</th>
            <th className="px-2 py-2 text-left text-xs" style={th}>Observaciones</th>
            <th style={{ width: 34 }}></th>
          </tr></thead>
          <tbody>
            {items.map((it, i) => {
              const vacio = !it.producto.trim();
              return (
                <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
                  <td className="px-2 py-1 text-xs" style={{ color: vacio ? C.clay : C.faint }}>{i + 1}</td>
                  {COLS.map((k) => (
                    <td key={k} className="px-1 py-1">
                      <input value={it[k]} onChange={(e) => upd(i, k, e.target.value)}
                        className="w-full rounded-md px-2 py-1 text-sm outline-none" style={{ ...inp, borderColor: k === "producto" && vacio ? C.clay : C.line }}
                        inputMode={k === "cantidad" ? "decimal" : undefined} />
                    </td>
                  ))}
                  <td className="px-1"><button onClick={() => del(i)} className="rounded-md p-1" style={{ color: C.clay, cursor: "pointer" }}><X size={14} /></button></td>
                </tr>
              );
            })}
            {items.length === 0 && <tr><td colSpan={8} className="px-3 py-4 text-center text-xs" style={{ color: C.faint }}>Sin ítems.</td></tr>}
          </tbody>
        </table>
      </div>
      <div className="mt-2 flex items-center justify-between">
        <Ghost onClick={add} color={C.gold}><Plus size={13} /> Agregar fila</Ghost>
        <span className="text-xs" style={{ color: C.faint }}>{items.filter((it) => it.producto.trim()).length} ítem(s) válido(s)</span>
      </div>
    </div>
  );
}

/* ============ Nueva solicitud (Excel/TXT) ============ */
function plantilla() {
  const ws = XLSX.utils.json_to_sheet([
    { producto: "Detergente industrial 5L", cantidad: 10, unidad: "bidón", proveedor: "Distribuidora XYZ", fecha_requerida: "2026-09-30", observaciones: "" },
    { producto: "Guantes nitrilo talla M", cantidad: 20, unidad: "caja", proveedor: "", fecha_requerida: "", observaciones: "Urgente" },
  ]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Solicitud");
  XLSX.writeFile(wb, "plantilla_solicitud_compras.xlsx");
}

function NuevaSolicitud({ session, onListo }) {
  const [step, setStep] = useState("upload"); // upload | map | review
  const [file, setFile] = useState(null);
  const [parsed, setParsed] = useState({ headers: [], rows: [] });
  const [mapping, setMapping] = useState([]);
  const [items, setItems] = useState([]);
  const [hdr, setHdr] = useState({ proveedor: "", fechaRequerida: "", observaciones: "" });
  const [err, setErr] = useState(""); const [aviso, setAviso] = useState(""); const [guardando, setGuardando] = useState(false);
  const fileRef = useRef(null);

  const cargarArchivo = async (f) => {
    setErr(""); setAviso(""); if (!f) return;
    try {
      const ext = f.name.split(".").pop().toLowerCase();
      let p;
      if (["xlsx", "xls", "xlsm", "csv"].includes(ext)) p = await parseXlsx(f);
      else if (["txt", "tsv"].includes(ext)) p = parseTxt(await f.text());
      else return setErr("Formato no soportado en esta etapa. Usa Excel o TXT (las imágenes llegan en la Etapa 3).");
      if (!p.headers.length || !p.rows.length) return setErr("No pude leer filas del archivo. Revisa que tenga encabezados y datos.");
      setFile(f); setParsed(p); setMapping(p.headers.map((h) => guessField(h))); setStep("map");
    } catch (e) { setErr("Error al leer el archivo: " + (e?.message || e)); }
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

  const guardar = async (enviar) => {
    const validos = items.filter((it) => it.producto.trim());
    if (!validos.length) return setErr("Necesitas al menos un ítem con producto.");
    setGuardando(true); setErr(""); setAviso("");
    try {
      const nueva = doc(colSol);
      let archivo = null, archivoFallo = false;
      if (file) {
        try {
          const path = `compras_archivos/${nueva.id}/${Date.now()}_${safeName(file.name)}`;
          const r = sref(storage, path);
          await uploadBytes(r, file);
          const url = await getDownloadURL(r);
          archivo = { nombre: file.name, path, url, tipo: file.type || "", size: file.size };
        } catch (e) { archivoFallo = true; }
      }
      const snap = await getDocs(colSol);
      const folio = snap.docs.reduce((mx, d) => Math.max(mx, Number(d.data().folio) || 0), 0) + 1;
      const por = { pin: session.pin, nombre: session.nombre };
      const estado = enviar ? "ENVIADA" : "BORRADOR";
      const historial = [{ accion: "creada", por, at: Date.now() }];
      if (enviar) historial.push({ accion: "enviada", por, at: Date.now() });
      await setDoc(nueva, {
        folio, estado,
        items: validos.map((it) => ({ ...it, cantidad: String(it.cantidad).trim() })),
        proveedor: hdr.proveedor.trim(), fechaRequerida: hdr.fechaRequerida.trim(), observaciones: hdr.observaciones.trim(),
        archivo,
        ocNumero: "", recepcionProgramada: null, recepcion: null,
        creadoPor: por, creadoAt: serverTimestamp(), historial,
      });
      if (archivoFallo) { setAviso("Solicitud guardada, pero el archivo original no se pudo subir (revisa las reglas de Storage). Los ítems quedaron guardados."); setGuardando(false); setTimeout(() => onListo(), 1800); }
      else onListo();
    } catch (e) { setGuardando(false); setErr("No se pudo guardar: " + (e?.message || e)); }
  };

  return (
    <Card>
      {step === "upload" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><Upload size={16} color={C.gold} /> Nueva solicitud</div>
          <div className="mb-4 text-xs" style={{ color: C.faint }}>Carga un Excel (.xlsx/.csv) o TXT. Interpreto las columnas, muestro previsualización y podrás corregir antes de guardar.</div>
          <div className="flex flex-wrap items-center gap-2">
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.xlsm,.csv,.txt,.tsv" className="hidden" onChange={(e) => cargarArchivo(e.target.files?.[0])} />
            <Btn onClick={() => fileRef.current?.click()}><FileSpreadsheet size={15} /> Elegir archivo</Btn>
            <Ghost onClick={plantilla} color={C.info}><Download size={13} /> Descargar plantilla</Ghost>
          </div>
          {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
        </>
      )}

      {step === "map" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><FileText size={16} color={C.gold} /> Interpretar columnas</div>
          <div className="mb-3 text-xs" style={{ color: C.faint }}>{file?.name} · {parsed.rows.length} fila(s). Asigna cada columna del archivo a un campo. Ajusté un mapeo automático; corrígelo si hace falta.</div>
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
            <Ghost onClick={() => { setStep("upload"); setErr(""); }}><ChevronLeft size={13} /> Volver</Ghost>
            <Btn onClick={aplicarMapeo}>Previsualizar <Check size={15} /></Btn>
          </div>
        </>
      )}

      {step === "review" && (
        <>
          <div className="mb-1 flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}><ClipboardList size={16} color={C.gold} /> Revisar y confirmar</div>
          <div className="mb-3 flex items-center gap-1.5 text-xs" style={{ color: C.faint }}>
            {file && <><Paperclip size={12} /> {file.name} (se guarda como original)</>}
          </div>
          <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-3">
            <Field label="Proveedor (opcional, cabecera)" value={hdr.proveedor} onChange={(e) => setHdr({ ...hdr, proveedor: e.target.value })} placeholder="Si aplica a toda la solicitud" />
            <Field label="Fecha requerida (opcional)" value={hdr.fechaRequerida} onChange={(e) => setHdr({ ...hdr, fechaRequerida: e.target.value })} placeholder="AAAA-MM-DD" />
            <Field label="Observaciones (opcional)" value={hdr.observaciones} onChange={(e) => setHdr({ ...hdr, observaciones: e.target.value })} placeholder="Nota general" />
          </div>
          <ItemsEditor items={items} setItems={setItems} />
          {err && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.clay }}><AlertTriangle size={13} /> {err}</div>}
          {aviso && <div className="mt-3 flex items-center gap-1.5 text-xs" style={{ color: C.amber }}><AlertTriangle size={13} /> {aviso}</div>}
          <div className="mt-4 flex flex-wrap items-center justify-between gap-2">
            <Ghost onClick={() => { setStep("map"); setErr(""); }}><ChevronLeft size={13} /> Volver al mapeo</Ghost>
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
    if (sol.archivo?.path) { try { await deleteObject(sref(storage, sol.archivo.path)); } catch (e) {} }
    await deleteDoc(doc(colSol, sol.id)); onClose();
  };

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto p-3" style={{ background: "#000000aa" }}>
      <div className="mt-6 w-full max-w-3xl rounded-2xl p-4" style={{ background: C.surface, border: `1px solid ${C.line}` }}>
        <div className="mb-3 flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.text }}>Solicitud #{sol.folio} <Badge estado={sol.estado} /></div>
          <button onClick={onClose} className="rounded-lg p-1.5" style={{ color: C.muted, cursor: "pointer" }}><X size={16} /></button>
        </div>
        <div className="mb-3 text-xs" style={{ color: C.faint }}>Creada por {sol.creadoPor?.nombre} · {fechaHora(sol.creadoAt)}</div>

        {sol.archivo ? (
          <a href={sol.archivo.url} target="_blank" rel="noreferrer" className="mb-3 inline-flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium" style={{ background: C.info + "1c", color: C.info }}>
            <Download size={13} /> Descargar original ({sol.archivo.nombre})
          </a>
        ) : <div className="mb-3 text-xs" style={{ color: C.faint }}>Sin archivo original adjunto.</div>}

        {esBorrador ? (
          <>
            <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-3">
              <Field label="Proveedor (cabecera)" value={hdr.proveedor} onChange={(e) => setHdr({ ...hdr, proveedor: e.target.value })} />
              <Field label="Fecha requerida" value={hdr.fechaRequerida} onChange={(e) => setHdr({ ...hdr, fechaRequerida: e.target.value })} />
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
          <div className="overflow-x-auto rounded-xl" style={{ border: `1px solid ${C.line}` }}>
            <table className="w-full text-sm" style={{ minWidth: 640 }}>
              <thead><tr style={{ background: C.surface2, color: C.faint }}>
                <th className="px-2 py-2 text-left text-xs">Producto</th><th className="px-2 py-2 text-left text-xs">Cant.</th>
                <th className="px-2 py-2 text-left text-xs">Unidad</th><th className="px-2 py-2 text-left text-xs">Proveedor</th>
                <th className="px-2 py-2 text-left text-xs">Fecha req.</th><th className="px-2 py-2 text-left text-xs">Obs.</th>
              </tr></thead>
              <tbody>{(sol.items || []).map((it, i) => (
                <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
                  <td className="px-2 py-1.5" style={{ color: C.text }}>{it.producto}</td><td className="px-2 py-1.5" style={{ color: C.muted }}>{it.cantidad}</td>
                  <td className="px-2 py-1.5" style={{ color: C.muted }}>{it.unidad}</td><td className="px-2 py-1.5" style={{ color: C.muted }}>{it.proveedor}</td>
                  <td className="px-2 py-1.5" style={{ color: C.muted }}>{it.fechaRequerida}</td><td className="px-2 py-1.5" style={{ color: C.muted }}>{it.observaciones}</td>
                </tr>
              ))}</tbody>
            </table>
          </div>
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
    .filter((s) => !q.trim() || String(s.folio).includes(q) || (s.items || []).some((it) => norm(it.producto).includes(norm(q))) || norm(s.proveedor).includes(norm(q)))
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
      {session.rol === "administrativo_compras" && <EnConstruccion Icono={FileCheck2} titulo="Solicitudes pendientes" etapa="4" desc="Ver solicitudes enviadas por el analista, descargar el original, registrar el N° de OC del ERP y programar fecha/hora de recepción hacia la agenda de bodega." />}
      {session.rol === "bodeguero" && <EnConstruccion Icono={Truck} titulo="Agenda de recepciones" etapa="5" desc="Atrasadas · Hoy · Próximas. Abrir una recepción y confirmar (estado RECEPCIONADA + fecha/hora real + usuario). Sin pedir productos ni cantidades." />}
    </>
  );
}

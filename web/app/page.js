"use client";
import { useState, useRef, useEffect } from "react";

const fmt = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(Number(n) || 0);

function chipFor(o) {
  if (o.hoja === "Cobros")
    return { cls: "cobro", badge: "Cobro", desc: `Alquiler · inquilino #${o.id_inquilino ?? "?"}`, amt: o.moneda === "USD" ? `US$ ${o.monto}` : fmt(o.monto) };
  if (o.hoja === "Movimientos")
    return { cls: "mov", badge: o.tipo === "Ingreso" ? "Ingreso" : "Egreso", desc: o.categoria || "Movimiento", amt: fmt(o.monto) };
  if (o.hoja === "Inflacion INDEC")
    return { cls: "ipc", badge: "IPC", desc: `Índice ${o.mes}`, amt: String(o.indice) };
  return { cls: "mov", badge: o.hoja, desc: "", amt: "" };
}

export default function Home() {
  const [texto, setTexto] = useState("");
  const [estado, setEstado] = useState("");
  const [salida, setSalida] = useState(null);
  const [grabando, setGrabando] = useState(false);
  const [enviando, setEnviando] = useState(false);
  const [fileName, setFileName] = useState("");
  const [stats, setStats] = useState(null);
  const fileRef = useRef(null);
  const recRef = useRef(null);
  const chunksRef = useRef([]);
  const audioRef = useRef(null);

  async function cargarStats() {
    try {
      const r = await fetch("/api/summary");
      const d = await r.json();
      if (!d.error) setStats(d);
    } catch (e) { /* sin datos aún */ }
  }
  useEffect(() => { cargarStats(); }, []);

  async function toggleRec() {
    if (grabando) { recRef.current && recRef.current.stop(); return; }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const rec = new MediaRecorder(stream);
      chunksRef.current = [];
      rec.ondataavailable = (e) => chunksRef.current.push(e.data);
      rec.onstop = () => {
        audioRef.current = new Blob(chunksRef.current, { type: "audio/webm" });
        setFileName("audio grabado 🎧");
        setGrabando(false);
        stream.getTracks().forEach((t) => t.stop());
      };
      rec.start(); recRef.current = rec; setGrabando(true); setEstado("Grabando… tocá de nuevo para frenar.");
    } catch (e) { setEstado("No pude acceder al micrófono: " + e.message); }
  }

  async function enviar() {
    const fd = new FormData();
    fd.append("texto", texto || "");
    const f = fileRef.current && fileRef.current.files[0];
    if (f) fd.append("archivo", f);
    else if (audioRef.current) fd.append("archivo", audioRef.current, "audio.webm");
    setEnviando(true); setEstado("Procesando con IA…"); setSalida(null);
    try {
      const r = await fetch("/api/process", { method: "POST", body: fd });
      const data = await r.json();
      setSalida(data);
      setEstado("");
      audioRef.current = null; setTexto(""); setFileName("");
      if (fileRef.current) fileRef.current.value = "";
      if (!data.error) cargarStats();
    } catch (e) { setEstado("Error: " + e.message); }
    finally { setEnviando(false); }
  }

  async function subirExcel(e) {
    const f = e.target.files[0];
    if (!f) return;
    setEstado("Subiendo tu planilla…"); setSalida(null);
    const fd = new FormData();
    fd.append("archivo", f);
    try {
      const r = await fetch("/api/upload", { method: "POST", body: fd });
      const d = await r.json();
      if (d.error) setEstado("⚠️ " + d.error);
      else { setEstado("✅ " + d.mensaje); cargarStats(); }
    } catch (err) { setEstado("Error: " + err.message); }
    e.target.value = "";
  }

  return (
    <>
      <header className="topbar">
        <div className="topbar-inner">
          <a className="brand" href="/" aria-label="Finanzas Familia — inicio">
            <span className="brand-mark" aria-hidden="true">
              <svg viewBox="0 0 24 24" width="17" height="17" fill="none">
                <rect x="3" y="13" width="4" height="8" rx="1.4" fill="currentColor" />
                <rect x="10" y="8" width="4" height="13" rx="1.4" fill="currentColor" />
                <rect x="17" y="3" width="4" height="18" rx="1.4" fill="currentColor" />
              </svg>
            </span>
            <span className="brand-name">Finanzas<span className="brand-sep">·</span><span className="brand-soft">Familia</span></span>
          </a>
          <div className="topbar-actions">
            <span className="year-pill">2026</span>
            <a href="/api/excel" className="topbar-excel">
              <button type="button" className="btn btn-ghost">
                <svg viewBox="0 0 24 24" width="15" height="15" fill="none" aria-hidden="true">
                  <path d="M12 3v11m0 0l-4-4m4 4l4-4M5 19h14" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                <span className="excel-label">Excel</span>
              </button>
            </a>
          </div>
        </div>
      </header>

      <div className="wrap">
      <header className="masthead">
        <p className="eyebrow">Economía doméstica</p>
        <h1 className="title">El dinero de la familia, <em>en orden.</em></h1>
        <p className="subtitle">Contale lo que pasó —por voz, texto o una foto del comprobante— y se anota solo en la planilla.</p>
      </header>

      <section className="stats">
        <div className="stat ingresos">
          <div className="lbl"><span className="dot" />Ingresos del año</div>
          <div className="val">{stats ? fmt(stats.total_ingresos) : "—"}</div>
          <div className="sub">alquileres + otros</div>
        </div>
        <div className="stat egresos">
          <div className="lbl"><span className="dot" />Egresos del año</div>
          <div className="val">{stats ? fmt(stats.egresos) : "—"}</div>
          <div className="sub">gastos cargados</div>
        </div>
        <div className="stat neto">
          <div className="lbl"><span className="dot" />Resultado neto</div>
          <div className="val">{stats ? fmt(stats.resultado_neto) : "—"}</div>
          <div className="sub">lo ahorrado</div>
        </div>
      </section>

      <section className="card">
        <div className="card-head">
          <h2>Anotar un movimiento</h2>
          <span className="hint">audio · texto · foto · PDF</span>
        </div>
        <textarea
          className="field"
          value={texto}
          onChange={(e) => setTexto(e.target.value)}
          placeholder="Ej: «Cobré el alquiler del Local 12, 350 mil por transferencia.»  ·  «Pagué la luz, 48.500 con débito.»  ·  «El IPC de abril dio 7600.»"
        />
        <div className="controls">
          <button className={"mic" + (grabando ? " on" : "")} onClick={toggleRec} title="Grabar audio" aria-label="Grabar audio">
            {grabando ? "■" : "●"}
          </button>
          <label className="file-pill">
            📎 {fileName || "Adjuntar foto / PDF / audio"}
            <input type="file" ref={fileRef} accept="audio/*,image/*,.pdf"
              onChange={(e) => setFileName(e.target.files[0]?.name || "")} />
          </label>
        </div>
        <div className="actions">
          <button className="btn btn-primary" onClick={enviar} disabled={enviando}>
            {enviando ? "Procesando…" : "Anotar en la planilla →"}
          </button>
          <a href="/api/excel"><button type="button" className="btn btn-ghost">Descargar Excel</button></a>
          <label className="btn btn-ghost" style={{ cursor: "pointer" }}>
            Subir Excel editado
            <input type="file" accept=".xlsx" style={{ display: "none" }} onChange={subirExcel} />
          </label>
        </div>
        {estado && <p className="status">{estado}</p>}
      </section>

      {salida && (
        <section className="card result">
          <div className="card-head"><h2>Resultado</h2></div>
          {salida.error ? (
            <div className="alert">⚠️ {salida.error}</div>
          ) : (
            <>
              {salida.resumen && <p className="summary-line">{salida.resumen}</p>}
              {(salida.operaciones || []).map((o, i) => {
                const c = chipFor(o);
                return (
                  <div className="op" key={i} style={{ animationDelay: `${i * 60}ms` }}>
                    <span className={"badge " + c.cls}>{c.badge}</span>
                    <span className="desc">{c.desc} <span className="meta">· fila {o.fila}</span></span>
                    <span className="amt">{c.amt}</span>
                  </div>
                );
              })}
              {!(salida.operaciones || []).length && <div className="empty">No se cargó ninguna fila.</div>}
            </>
          )}
        </section>
      )}

      <p className="footnote">Los números se actualizan solos en el Excel · podés editar todo a mano cuando quieras</p>
      </div>
    </>
  );
}

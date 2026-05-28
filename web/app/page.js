"use client";
import { useState, useRef, useEffect } from "react";
import { useSession, signIn, signOut } from "next-auth/react";

function UserBadge() {
  const { data: session, status } = useSession();
  if (status === "loading") return null;
  if (status !== "authenticated") return null;
  const nombre = session.user?.nombreFamilia || session.user?.name || "";
  const inicial = String(nombre || "?").trim().charAt(0).toUpperCase();
  return (
    <div className="user-badge" title={session.user?.email || ""}>
      <span className="user-dot">{inicial}</span>
      <span className="user-name">{nombre}</span>
      <button type="button" className="user-logout" onClick={() => signOut({ callbackUrl: "/login" })}>
        Salir
      </button>
    </div>
  );
}

function ThemeToggle() {
  const [tema, setTema] = useState("light");
  useEffect(() => {
    const inicial = document.documentElement.getAttribute("data-theme") || "light";
    setTema(inicial);
  }, []);
  function toggle() {
    const next = tema === "dark" ? "light" : "dark";
    setTema(next);
    document.documentElement.setAttribute("data-theme", next);
    try { localStorage.setItem("tema", next); } catch (e) {}
  }
  const esDark = tema === "dark";
  return (
    <button type="button" className="theme-toggle" onClick={toggle}
            aria-label={esDark ? "Cambiar a modo claro" : "Cambiar a modo oscuro"}
            title={esDark ? "Modo claro" : "Modo oscuro"}>
      {esDark ? (
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <circle cx="12" cy="12" r="4" />
          <path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41" />
        </svg>
      ) : (
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
        </svg>
      )}
    </button>
  );
}

function renderNegritas(linea) {
  const partes = String(linea).split(/(\*\*[^*]+\*\*)/g);
  return partes.map((p, i) => p.startsWith("**") && p.endsWith("**")
    ? <strong key={i}>{p.slice(2, -2)}</strong>
    : <span key={i}>{p}</span>);
}

function AnalisisTexto({ texto }) {
  const lineas = String(texto || "").split(/\r?\n/);
  const bloques = [];
  let bulletAcc = [];
  const flush = () => {
    if (bulletAcc.length) {
      bloques.push(<ul key={"u" + bloques.length}>{bulletAcc.map((b, i) => <li key={i}>{renderNegritas(b)}</li>)}</ul>);
      bulletAcc = [];
    }
  };
  for (const raw of lineas) {
    const l = raw.trim();
    if (!l) { flush(); continue; }
    const esBullet = /^[-·*]\s+/.test(l);
    if (esBullet) { bulletAcc.push(l.replace(/^[-·*]\s+/, "")); continue; }
    flush();
    const titulo = /^\*\*[^*]+\*\*$/.test(l);
    if (titulo) bloques.push(<h3 key={"t" + bloques.length}>{l.replace(/^\*\*|\*\*$/g, "")}</h3>);
    else bloques.push(<p key={"p" + bloques.length}>{renderNegritas(l)}</p>);
  }
  flush();
  return <div className="analisis">{bloques}</div>;
}

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
  const [historial, setHistorial] = useState([]); // [{rol:'user'|'model', texto}]
  const [pregunta, setPregunta] = useState(false); // la IA espera una respuesta
  const [grabando, setGrabando] = useState(false);
  const [enviando, setEnviando] = useState(false);
  const [fileName, setFileName] = useState("");
  const [stats, setStats] = useState(null);
  const [analisis, setAnalisis] = useState(null);
  const [analizando, setAnalizando] = useState(false);
  const [errAnalisis, setErrAnalisis] = useState("");
  const fileRef = useRef(null);
  const recRef = useRef(null);
  const chunksRef = useRef([]);
  const audioRef = useRef(null);
  const fieldRef = useRef(null);

  async function cargarStats() {
    try {
      const r = await fetch("/api/summary");
      const d = await r.json();
      if (!d.error) setStats(d);
    } catch (e) { /* sin datos aún */ }
  }
  useEffect(() => { cargarStats(); }, []);
  // Cuando la IA pregunta, dejamos el cursor listo para responder.
  useEffect(() => { if (pregunta && fieldRef.current) fieldRef.current.focus(); }, [pregunta]);

  async function toggleRec() {
    if (grabando) { recRef.current && recRef.current.stop(); return; }
    if (!navigator.mediaDevices?.getUserMedia) {
      setEstado("El micrófono requiere abrir la app desde http://localhost:3000 (o configurar HTTPS). El hostname personalizado no tiene acceso al micrófono por restricciones del navegador.");
      return;
    }
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
    const f = fileRef.current && fileRef.current.files[0];
    const tieneTexto = (texto || "").trim().length > 0;
    if (!tieneTexto && !f && !audioRef.current) {
      setEstado("Escribí algo, grabá un audio o adjuntá un comprobante.");
      return;
    }
    const fd = new FormData();
    fd.append("texto", texto || "");
    fd.append("historial", JSON.stringify(historial.slice(-12)));
    if (f) fd.append("archivo", f);
    else if (audioRef.current) fd.append("archivo", audioRef.current, "audio.webm");
    const textoEnviado = (texto || "").trim() || (f ? "📎 " + f.name : "🎧 audio");
    setEnviando(true); setEstado(pregunta ? "Pensando tu respuesta…" : "Procesando con IA…");
    try {
      const r = await fetch("/api/process", { method: "POST", body: fd });
      const data = await r.json();
      setSalida(data);
      setEstado("");
      if (data.error) { setEnviando(false); return; }
      // Acumulamos la conversación para mantener el contexto en la repregunta.
      setHistorial((h) => [...h, { rol: "user", texto: textoEnviado }, { rol: "model", texto: data.resumen || "Listo." }]);
      setPregunta(Boolean(data.pregunta));
      audioRef.current = null; setTexto(""); setFileName("");
      if (fileRef.current) fileRef.current.value = "";
      cargarStats();
    } catch (e) { setEstado("Error: " + e.message); }
    finally { setEnviando(false); }
  }

  async function analizarMes() {
    setAnalizando(true); setErrAnalisis(""); setAnalisis(null);
    try {
      const r = await fetch("/api/analizar");
      const d = await r.json();
      if (d.error) setErrAnalisis(d.error);
      else setAnalisis(d);
    } catch (e) { setErrAnalisis(e.message); }
    finally { setAnalizando(false); }
  }

  function nuevaConsulta() {
    setHistorial([]); setSalida(null); setPregunta(false); setTexto(""); setFileName("");
    if (fileRef.current) fileRef.current.value = "";
    audioRef.current = null; setEstado("");
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

  function onKeyDown(e) {
    if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) { e.preventDefault(); enviar(); }
  }

  const ops = (salida && !salida.error && salida.operaciones) || [];

  return (
    <>
      <header className="topbar">
        <div className="topbar-inner">
          <a className="brand" href="/" aria-label="Finanzas Familia — inicio">
            <span className="brand-name">Finanzas&nbsp;<span className="brand-soft">Familia</span><span className="brand-dot">.</span></span>
          </a>
          <div className="topbar-actions">
            <span className="year-pill">2026</span>
            <UserBadge />
            <ThemeToggle />
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
          <div className="lbl"><span className="dot" />Ingresos del mes</div>
          <div className="val">{stats ? fmt(stats.total_ingresos) : "—"}</div>
          <div className="sub">{stats?.mes_nombre ? `${stats.mes_nombre} · alquileres + otros` : "alquileres + otros"}</div>
        </div>
        <div className="stat egresos">
          <div className="lbl"><span className="dot" />Egresos del mes</div>
          <div className="val">{stats ? fmt(stats.egresos) : "—"}</div>
          <div className="sub">{stats?.mes_nombre ? `${stats.mes_nombre} · gastos cargados` : "gastos cargados"}</div>
        </div>
        <div className="stat neto">
          <div className="lbl"><span className="dot" />Resultado del mes</div>
          <div className="val">{stats ? fmt(stats.resultado_neto) : "—"}</div>
          <div className="sub">{stats?.mes_nombre ? `${stats.mes_nombre} · lo ahorrado` : "lo ahorrado"}</div>
        </div>
      </section>

      <section className="analisis-bar">
        <button type="button" className="btn btn-analisis" onClick={analizarMes} disabled={analizando}>
          {analizando ? "Pensando…" : (analisis ? "↻ Volver a analizar" : "🧠 Analizar el mes con IA")}
        </button>
        {analisis && <span className="analisis-hint">{analisis.datos?.mes?.nombre} {analisis.datos?.mes?.anio}</span>}
      </section>

      {(analisis || errAnalisis) && (
        <section className="card analisis-card">
          {errAnalisis && <div className="alert">⚠️ {errAnalisis}</div>}
          {analisis && <AnalisisTexto texto={analisis.texto} />}
        </section>
      )}

      <section className="card">
        <div className="card-head">
          <h2>{pregunta ? "Respondé para completar" : "Anotar un movimiento"}</h2>
          <span className="hint">audio · texto · foto · PDF</span>
        </div>
        <textarea
          ref={fieldRef}
          className="field"
          value={texto}
          onChange={(e) => setTexto(e.target.value)}
          onKeyDown={onKeyDown}
          placeholder={pregunta
            ? "Respondé acá lo que te preguntó (ej: «fueron 48.500 con débito»)…"
            : "Ej: «Cobré el alquiler del Local 12, 350 mil por transferencia.»  ·  «Pagué la luz, 48.500 con débito.»  ·  «El IPC de abril dio 7600.»"}
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
            {enviando ? "Procesando…" : pregunta ? "Responder →" : "Anotar en la planilla →"}
          </button>
          {historial.length > 0 && (
            <button type="button" className="btn btn-ghost" onClick={nuevaConsulta} disabled={enviando}>
              Empezar de nuevo
            </button>
          )}
          <label className="btn btn-ghost" style={{ cursor: "pointer" }}>
            Subir Excel editado
            <input type="file" accept=".xlsx" style={{ display: "none" }} onChange={subirExcel} />
          </label>
        </div>
        {estado && <p className="status">{estado}</p>}
      </section>

      {(historial.length > 0 || (salida && salida.error)) && (
        <section className="card result">
          <div className="card-head">
            <h2>Conversación</h2>
            {pregunta && <span className="hint hint-warn">esperando tu respuesta</span>}
          </div>

          {salida && salida.error && <div className="alert">⚠️ {salida.error}</div>}

          <div className="thread" style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {historial.map((m, i) => {
              const esIA = m.rol === "model";
              const ultimo = i === historial.length - 1;
              return (
                <div
                  key={i}
                  className={"bubble " + (esIA ? "ia" : "yo") + (ultimo && pregunta && esIA ? " preguntando" : "")}
                  style={{ alignSelf: esIA ? "flex-start" : "flex-end" }}
                >
                  {esIA ? (ultimo && pregunta ? "🤔 " : "") : "🗣️ "}{m.texto}
                </div>
              );
            })}
          </div>

          {ops.length > 0 && (
            <div style={{ marginTop: 12 }}>
              {ops.map((o, i) => {
                const c = chipFor(o);
                return (
                  <div className="op" key={i} style={{ animationDelay: `${i * 60}ms` }}>
                    <span className={"badge " + c.cls}>{c.badge}</span>
                    <span className="desc">{c.desc} <span className="meta">· fila {o.fila}</span></span>
                    <span className="amt">{c.amt}</span>
                  </div>
                );
              })}
            </div>
          )}
        </section>
      )}

      <p className="footnote">Los números se actualizan solos en el Excel · podés editar todo a mano cuando quieras · Ctrl/⌘ + Enter envía</p>
      </div>
    </>
  );
}

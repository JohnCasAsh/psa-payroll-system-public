import React, { useRef, useState, useEffect, useCallback } from "react";
import { createPortal } from "react-dom";
import CSForm6 from "./components/CSForm6/CSForm6.jsx";

/* ── Signature capture ───────────────────────────────────────────────────────
   The card beside the form only shows the saved signature. Signing happens in
   a large modal pad (same look as the 201 Files viewer) so there is room to
   write naturally with a mouse, finger, or stylus. */

const INK = "#111827";
const BASE_PEN_WIDTH = 2.8;  // mouse/finger stroke width on the pad, in CSS px
const INK_WEIGHT     = 0.021; // exported stroke width as a share of the ink height
const MAX_UPLOAD_BYTES = 1.5 * 1024 * 1024; // leave routes accept 2 MB bodies; base64 adds ~33%

const PenIcon = ({ size = 15 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M12 20h9" /><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z" />
  </svg>
);

const btnBase = {
  display: "inline-flex", alignItems: "center", gap: 6, border: "none", borderRadius: 8,
  padding: "8px 14px", fontSize: "0.82rem", fontWeight: 600, cursor: "pointer", boxShadow: "none", lineHeight: 1.2,
};
const ghostBtn   = { ...btnBase, background: "transparent", color: "var(--text-muted)", border: "1px solid var(--border-2)" };
const primaryBtn = { ...btnBase, background: "var(--psa-navy)", color: "#fff", fontWeight: 700 };
const smallBtn   = { padding: "6px 11px", fontSize: "0.76rem" };

function readImageFile(file, onDone, onError) {
  if (!/^image\/(png|jpeg|webp)$/.test(file.type)) { onError("Use a PNG, JPG, or WEBP image."); return; }
  if (file.size > MAX_UPLOAD_BYTES) { onError("That image is too large (max 1.5 MB)."); return; }
  const reader = new FileReader();
  reader.onload = () => onDone(reader.result);
  reader.onerror = () => onError("Could not read that file.");
  reader.readAsDataURL(file);
}

// Smooth stroke: quadratic curves through the midpoints between samples.
function drawStroke(ctx, pts, width) {
  ctx.lineWidth = width;
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
  ctx.strokeStyle = INK;
  ctx.fillStyle = INK;
  if (pts.length === 1) {
    ctx.beginPath();
    ctx.arc(pts[0].x, pts[0].y, width / 2, 0, Math.PI * 2);
    ctx.fill();
    return;
  }
  ctx.beginPath();
  ctx.moveTo(pts[0].x, pts[0].y);
  for (let i = 1; i < pts.length - 1; i++) {
    ctx.quadraticCurveTo(pts[i].x, pts[i].y, (pts[i].x + pts[i + 1].x) / 2, (pts[i].y + pts[i + 1].y) / 2);
  }
  const last = pts[pts.length - 1];
  ctx.lineTo(last.x, last.y);
  ctx.stroke();
}

function SignatureModal({ label, onSave, onClose }) {
  const wrapRef   = useRef(null);
  const canvasRef = useRef(null);
  const fileRef   = useRef(null);
  const strokes   = useRef([]);
  const active    = useRef(null);
  const closeRef  = useRef(onClose);
  closeRef.current = onClose;
  const [count, setCount]       = useState(0);
  const [pressing, setPressing] = useState(false);
  const [error, setError]       = useState("");

  const redraw = useCallback(() => {
    const c = canvasRef.current;
    if (!c) return;
    const ctx = c.getContext("2d");
    const dpr = window.devicePixelRatio || 1;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, c.width, c.height);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    for (const s of strokes.current) drawStroke(ctx, s.pts, s.w);
    if (active.current) drawStroke(ctx, active.current.pts, active.current.w);
  }, []);

  // Keep the canvas backing store matched to its on-screen size and pixel ratio.
  useEffect(() => {
    const wrap = wrapRef.current;
    const c = canvasRef.current;
    const fit = () => {
      const dpr = window.devicePixelRatio || 1;
      const { width, height } = wrap.getBoundingClientRect();
      c.width = Math.round(width * dpr);
      c.height = Math.round(height * dpr);
      redraw();
    };
    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(wrap);
    return () => ro.disconnect();
  }, [redraw]);

  useEffect(() => {
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    const onKey = (e) => { if (e.key === "Escape") closeRef.current(); };
    document.addEventListener("keydown", onKey);
    return () => {
      document.body.style.overflow = prevOverflow;
      document.removeEventListener("keydown", onKey);
    };
  }, []);

  const pointAt = (ev) => {
    const r = canvasRef.current.getBoundingClientRect();
    return { x: ev.clientX - r.left, y: ev.clientY - r.top };
  };

  function down(e) {
    e.preventDefault();
    canvasRef.current.setPointerCapture(e.pointerId);
    const w = e.pointerType === "pen" ? 1.8 + (e.pressure || 0.5) * 2 : BASE_PEN_WIDTH;
    active.current = { pts: [pointAt(e)], w };
    setPressing(true);
    setError("");
    redraw();
  }
  function move(e) {
    if (!active.current) return;
    e.preventDefault();
    const ne = e.nativeEvent;
    const batch = ne.getCoalescedEvents?.();
    for (const ev of batch?.length ? batch : [ne]) active.current.pts.push(pointAt(ev));
    redraw();
  }
  function up() {
    if (!active.current) return;
    strokes.current.push(active.current);
    active.current = null;
    setPressing(false);
    setCount(strokes.current.length);
    redraw();
  }
  function undo()  { strokes.current.pop(); setCount(strokes.current.length); redraw(); }
  function clear() { strokes.current = []; setCount(0); redraw(); }

  // Crop to the ink so the image sits neatly over the printed name on the form.
  // The form shrinks every signature to the same height, so a fixed pen width
  // would print thin for a signature drawn large and heavy for one drawn small.
  // Widths are rescaled to a share of the ink height (keeping pen-pressure
  // differences) so every signature prints with the same pen weight.
  function exportPng() {
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const s of strokes.current) for (const p of s.pts) {
      minX = Math.min(minX, p.x); minY = Math.min(minY, p.y);
      maxX = Math.max(maxX, p.x); maxY = Math.max(maxY, p.y);
    }
    const inkH = Math.max(maxY - minY, 1);
    const k = Math.min(6, Math.max(0.5, (INK_WEIGHT * inkH) / BASE_PEN_WIDTH));
    const widths = strokes.current.map(s => s.w * k);
    // Margin scales with the drawing too, so small and large signatures fill the line alike.
    const edge = Math.max(...widths) + Math.max(2, inkH * 0.03);
    minX -= edge; minY -= edge; maxX += edge; maxY += edge;
    const w = maxX - minX, h = maxY - minY;
    const scale = Math.min(2, 1000 / w);
    const out = document.createElement("canvas");
    out.width = Math.ceil(w * scale);
    out.height = Math.ceil(h * scale);
    const ctx = out.getContext("2d");
    ctx.setTransform(scale, 0, 0, scale, -minX * scale, -minY * scale);
    strokes.current.forEach((s, i) => drawStroke(ctx, s.pts, widths[i]));
    return out.toDataURL("image/png");
  }

  return createPortal(
    <div className="no-print" role="dialog" aria-modal="true" aria-label={label}
      style={{ position: "fixed", inset: 0, background: "rgba(13,35,70,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16, backdropFilter: "blur(2px)" }}>
      <div style={{ background: "var(--surface)", borderRadius: 16, width: "100%", maxWidth: 920, maxHeight: "100%", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>

        <div style={{ padding: "14px 14px 14px 20px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          <span style={{ color: "var(--psa-navy)", display: "inline-flex" }}><PenIcon /></span>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: "0.92rem", fontWeight: 800, color: "var(--psa-navy)", letterSpacing: "-0.01em" }}>{label}</div>
            <div style={{ fontSize: "0.74rem", color: "var(--text-muted)", marginTop: 2 }}>Sign inside the box with your mouse, finger, or stylus.</div>
          </div>
          <button type="button" onClick={onClose} title="Close (Esc)" aria-label="Close"
            style={{ padding: "4px 8px", background: "transparent", border: "none", color: "var(--text-muted)", cursor: "pointer", fontSize: "1.2rem", lineHeight: 1, boxShadow: "none", flexShrink: 0 }}
            onMouseEnter={e => { e.currentTarget.style.color = "#991b1b"; }}
            onMouseLeave={e => { e.currentTarget.style.color = "var(--text-muted)"; }}>✕</button>
        </div>

        <div style={{ padding: 18, background: "var(--surface-2)", overflow: "auto" }}>
          <div ref={wrapRef}
            style={{ position: "relative", width: "100%", height: "clamp(220px, 42vh, 360px)", background: "#fff", border: "1.5px dashed var(--border-2)", borderRadius: 12, overflow: "hidden", touchAction: "none", cursor: "crosshair" }}>
            <div aria-hidden="true" style={{ position: "absolute", left: "6%", right: "6%", bottom: "26%", borderBottom: "1.5px solid #cbd5e1", pointerEvents: "none" }} />
            <div aria-hidden="true" style={{ position: "absolute", left: "6%", bottom: "calc(26% + 6px)", fontSize: 22, lineHeight: 1, color: "#94a3b8", pointerEvents: "none" }}>✕</div>
            {count === 0 && !pressing && (
              <div aria-hidden="true" style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", color: "#94a3b8", fontSize: "1rem", fontWeight: 500, pointerEvents: "none" }}>
                Sign here
              </div>
            )}
            <canvas ref={canvasRef}
              onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={up}
              style={{ position: "absolute", inset: 0, width: "100%", height: "100%", touchAction: "none" }} />
          </div>
          <div style={{ fontSize: "0.74rem", color: "var(--text-muted)", marginTop: 8, textAlign: "center" }}>
            Sign above the line. It will be placed on the form over your printed name.
          </div>
        </div>

        <div style={{ padding: "12px 16px", borderTop: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", flexShrink: 0 }}>
          <button type="button" onClick={undo} disabled={!count} style={{ ...ghostBtn, opacity: count ? 1 : 0.5, cursor: count ? "pointer" : "not-allowed" }}>Undo</button>
          <button type="button" onClick={clear} disabled={!count} style={{ ...ghostBtn, opacity: count ? 1 : 0.5, cursor: count ? "pointer" : "not-allowed" }}>Clear</button>
          <button type="button" onClick={() => fileRef.current?.click()} style={ghostBtn}>Upload image</button>
          <input ref={fileRef} type="file" accept="image/png,image/jpeg,image/webp" style={{ display: "none" }}
            onChange={e => {
              const f = e.target.files?.[0];
              e.target.value = "";
              if (f) readImageFile(f, onSave, setError);
            }} />
          {error && <span style={{ color: "#b91c1c", fontSize: "0.78rem" }}>{error}</span>}
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <button type="button" onClick={onClose} style={ghostBtn}>Cancel</button>
            <button type="button" onClick={() => onSave(exportPng())} disabled={!count}
              style={{ ...primaryBtn, padding: "9px 18px", opacity: count ? 1 : 0.5, cursor: count ? "pointer" : "not-allowed" }}>
              Use this signature
            </button>
          </div>
        </div>
      </div>
    </div>,
    document.body
  );
}

// `signAs` = { options, value, onChange } — shown when a section has more than
// one name set up (e.g. Regional Director and OIC), so the signer picks theirs.
function SignaturePad({ value, onChange, label, signAs }) {
  const [open, setOpen]   = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef(null);

  return (
    <div className="no-print" style={{ border: "1px solid var(--border-2, #d8dee8)", borderRadius: 12, padding: 14, background: "var(--surface, #fff)", width: 340, maxWidth: "100%" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 10 }}>
        <div style={{ fontSize: "0.82rem", fontWeight: 700, color: "var(--text)" }}>{label}</div>
        <span style={{ fontSize: "0.68rem", fontWeight: 700, padding: "2px 9px", borderRadius: 99, whiteSpace: "nowrap",
          background: value ? "#f0fdf4" : "var(--surface-2)", color: value ? "#166534" : "var(--text-muted)" }}>
          {value ? "Signed" : "Not signed"}
        </span>
      </div>

      {signAs?.options?.length > 1 && (
        <label style={{ marginBottom: 10, gap: 4, fontSize: "0.7rem", fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--text-muted)" }}>
          Sign as
          <select value={signAs.value ?? ""} onChange={e => signAs.onChange(Number(e.target.value))}
            style={{ padding: "8px 32px 8px 10px", fontSize: "0.84rem", fontWeight: 600, textTransform: "none", letterSpacing: 0, color: "var(--text)", borderRadius: 8 }}>
            {signAs.options.map(o => <option key={o.id} value={o.id}>{o.name} — {o.designation}</option>)}
          </select>
        </label>
      )}

      <button type="button" onClick={() => setOpen(true)} title={value ? "Sign again" : "Click to sign"}
        style={{ width: "100%", height: 104, padding: 8, display: "flex", alignItems: "center", justifyContent: "center",
          background: "#fff", border: "1.5px dashed var(--border-2)", borderRadius: 10, cursor: "pointer", boxShadow: "none", color: "var(--text-muted)" }}>
        {value
          ? <img src={value} alt="Signature" style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain" }} />
          : <span style={{ display: "inline-flex", alignItems: "center", gap: 8, fontSize: "0.88rem", fontWeight: 600 }}><PenIcon size={16} /> Click to sign</span>}
      </button>

      <div style={{ display: "flex", gap: 6, marginTop: 10, flexWrap: "wrap" }}>
        <button type="button" onClick={() => setOpen(true)} style={{ ...primaryBtn, ...smallBtn }}>
          <PenIcon size={13} /> {value ? "Sign again" : "Sign"}
        </button>
        <button type="button" onClick={() => fileRef.current?.click()} style={{ ...ghostBtn, ...smallBtn }}>Upload image</button>
        {value && (
          <button type="button" onClick={() => onChange("")} style={{ ...ghostBtn, ...smallBtn, color: "#b91c1c", borderColor: "#fecaca" }}>Remove</button>
        )}
      </div>
      {error && <div style={{ color: "#b91c1c", fontSize: "0.74rem", marginTop: 6 }}>{error}</div>}
      <input ref={fileRef} type="file" accept="image/png,image/jpeg,image/webp" style={{ display: "none" }}
        onChange={e => {
          const f = e.target.files?.[0];
          e.target.value = "";
          if (f) readImageFile(f, (v) => { setError(""); onChange(v); }, setError);
        }} />

      {open && (
        <SignatureModal label={label} onClose={() => setOpen(false)} onSave={(v) => { setError(""); onChange(v); setOpen(false); }} />
      )}
    </div>
  );
}

// Who receives an application at a section: names set to the applicant's
// division, else "All divisions" names, else everyone in the section.
// Mirrors eligibleLeaveSignatories in services/api/src/db.mysql.js.
export function eligibleSignatories(all, role, divisionId) {
  const ofRole = all.filter(s => s.role === role);
  const exact = divisionId != null ? ofRole.filter(s => s.division_id != null && Number(s.division_id) === Number(divisionId)) : [];
  if (exact.length) return exact;
  const general = ofRole.filter(s => s.division_id == null);
  return general.length ? general : ofRole;
}

/* ── LeaveForm ───────────────────────────────────────────────────────────────
   Consolidated props vs. the reference leave app:
     - 22 individual applicant field props → `fields` object + `onFieldChange`
     - 7 individual section-7 props       → `s7Fields` object + `onS7Change`
     - 4 individual signature props       → `sigs` object + `onSigChange`
   This removes ~36 prop declarations and the same number on every call site.
──────────────────────────────────────────────────────────────────────────── */
export default function LeaveForm({
  leaveView,
  isAdmin,
  userRole,
  app,
  /* consolidated applicant fields */
  fields,        onFieldChange,
  /* consolidated section 7 fields */
  s7Fields,      onS7Change,
  /* consolidated signatures */
  sigs,          onSigChange,
  /* names printed under the Section 7 lines, and who the signer picked */
  signatories = [], signerIds = {}, onSignerChange,
  newDivisionId = null,
  /* status */
  submitting,    submitError,
  s7Saving,      s7Error,      s7Success,
  actionSaving,  actionError,  actionSuccess,
  /* callbacks */
  leaveHistory,
  onSubmit,
  onSaveSection7,
  onAction,
  onBack,
}) {
  const isNew  = leaveView === "new";
  const isView = leaveView === "view";

  const canEditHR       = isView && (isAdmin || userRole === "hr");
  const canEditCRASD    = isView && (isAdmin || userRole === "crasd");
  const canEditDirector = isView && (isAdmin || userRole === "director");

  const ws = app?.workflow_status ?? "pending";
  const hrCanAct       = userRole === "hr"       && ws === "pending";
  const crasdCanAct    = userRole === "crasd"    && ws === "hr_approved";
  const directorCanAct = userRole === "director" && ws === "crasd_approved";
  const anyCanAct = hrCanAct || crasdCanAct || directorCanAct;

  const APPLICANT_FIELDS = new Set([
    "office_department","last_name","first_name","middle_name","date_of_filing","position","salary",
    "leave_type","leave_type_other","details_vacation_local","details_vacation_abroad",
    "details_sick_in_hospital","details_sick_out_patient","details_special_women_illness",
    "details_study_masters","details_study_bar","details_study_other",
    "details_monetization","details_terminal","num_working_days","inclusive_dates","commutation",
  ]);
  const IDENTITY_FIELDS = new Set(["office_department","last_name","first_name","middle_name","position","salary"]);
  const HR_FIELDS    = new Set(["credits_as_of","vl_total_earned","vl_less_application","vl_balance","sl_total_earned","sl_less_application","sl_balance"]);
  const CRASD_FIELDS = new Set(["recommendation","disapproval_reason"]);
  const DIR_FIELDS   = new Set(["approved_days_with_pay","approved_days_without_pay","approved_others_specify","disapproved_due_to"]);

  function canEdit(field) {
    if (IDENTITY_FIELDS.has(field)) return false;
    if (APPLICANT_FIELDS.has(field)) return isNew;
    if (HR_FIELDS.has(field))    return canEditHR;
    if (CRASD_FIELDS.has(field)) return canEditCRASD;
    if (DIR_FIELDS.has(field))   return canEditDirector;
    return false;
  }

  const data = {};
  for (const f of APPLICANT_FIELDS) {
    data[f] = isNew ? (fields?.[f] ?? "") : (app?.[f] ?? "");
  }
  for (const f of [...HR_FIELDS, ...CRASD_FIELDS, ...DIR_FIELDS]) {
    data[f] = canEdit(f) ? (s7Fields?.[f] ?? "") : (app?.[f] ?? "");
  }

  function handleChange(field, value) {
    if (!canEdit(field)) return;
    if (APPLICANT_FIELDS.has(field)) onFieldChange?.(field, value);
    else onS7Change?.(field, value);
  }

  // Only the officer whose turn it is can sign; everyone else (admins included)
  // sees the saved signatures.
  const signatures = {
    applicant: isNew ? (sigs?.applicant ?? "") : (app?.applicant_signature ?? ""),
    hr:        hrCanAct       ? (sigs?.hr       ?? "") : (app?.hr_signature       ?? ""),
    crasd:     crasdCanAct    ? (sigs?.crasd     ?? "") : (app?.crasd_signature    ?? ""),
    director:  directorCanAct ? (sigs?.director  ?? "") : (app?.director_signature ?? ""),
  };

  // Name under each line: a signed section keeps the name it was signed with;
  // the section being signed shows the signer's own name (or their pick); the
  // rest show the first name routed to the applicant's division.
  const divisionId = isNew ? newDivisionId : app?.applicant_division_id;
  // `mine` comes from the server: the name is for the signer's own division or all divisions.
  const isMine = (s) => !!s.mine;
  const acting = { hr: hrCanAct, crasd: crasdCanAct, director: directorCanAct };
  const optionsFor = (role) => {
    const eligible = eligibleSignatories(signatories, role, divisionId);
    return acting[role] ? eligible.filter(isMine) : eligible;
  };
  const pickFor = (role) => {
    const opts = optionsFor(role);
    return opts.find(s => s.id === signerIds[role])
      ?? (acting[role] ? opts.find(s => s.division_id != null) : null)
      ?? opts[0] ?? null;
  };
  const officers = Object.fromEntries(["hr", "crasd", "director"].map(role => {
    if (!isNew && app?.[`${role}_name`]) return [role, { name: app[`${role}_name`], designation: app[`${role}_designation`] ?? "" }];
    const p = pickFor(role);
    return [role, { name: p?.name ?? "", designation: p?.designation ?? "" }];
  }));
  const signAsFor = (role) => acting[role]
    ? { options: optionsFor(role), value: pickFor(role)?.id, onChange: (id) => onSignerChange?.(role, id) }
    : undefined;

  const wsBadge = {
    pending: "Pending HR Review", hr_approved: "Pending CRASD Review",
    crasd_approved: "Pending Director Review", approved: "Approved", disapproved: "Disapproved",
  }[ws] ?? ws;
  const wsBadgeStyle = ws === "approved"
    ? { bg: "#f0fdf4", text: "#15803d", border: "#bbf7d0" }
    : ws === "disapproved"
    ? { bg: "#fff1f0", text: "#b91c1c", border: "#fecaca" }
    : { bg: "#fffbef", text: "#92400e", border: "#fde68a" };

  const btn = { border: "none", borderRadius: 6, padding: "8px 22px", cursor: "pointer", fontWeight: 700, fontSize: "0.88rem", color: "#fff" };

  return (
    <div>
      {isView && (
        <div className="no-print" style={{ maxWidth: "210mm", margin: "0 auto 10px", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <span style={{ fontSize: "0.78rem", fontWeight: 700, padding: "4px 14px", borderRadius: 99, background: wsBadgeStyle.bg, color: wsBadgeStyle.text, border: `1px solid ${wsBadgeStyle.border}` }}>
            Status: {wsBadge}
          </span>
          {anyCanAct && (
            <span style={{ fontSize: "0.75rem", color: "#92400e", background: "#fef9c3", border: "1px solid #fde68a", borderRadius: 6, padding: "3px 10px" }}>
              Your action is required. Fill in your section below, sign, then click the action button.
            </span>
          )}
        </div>
      )}

      <div className="cs-sheet-wrap" style={{ display: "flex", justifyContent: "center" }}>
        <CSForm6
          data={data}
          onChange={handleChange}
          canEdit={canEdit}
          signatures={signatures}
          officers={officers}
        />
      </div>

      {(isNew || anyCanAct) && (
        <div className="no-print" style={{ maxWidth: "210mm", margin: "16px auto 0", display: "flex", gap: 14, flexWrap: "wrap" }}>
          {isNew          && <SignaturePad value={sigs?.applicant ?? ""} onChange={v => onSigChange?.("applicant", v)} label="Signature of Applicant" />}
          {hrCanAct       && <SignaturePad value={sigs?.hr        ?? ""} onChange={v => onSigChange?.("hr",        v)} label="HR Officer Signature (7.A)"  signAs={signAsFor("hr")} />}
          {crasdCanAct    && <SignaturePad value={sigs?.crasd     ?? ""} onChange={v => onSigChange?.("crasd",     v)} label="CRASD Chief Signature (7.B)" signAs={signAsFor("crasd")} />}
          {directorCanAct && <SignaturePad value={sigs?.director ?? ""} onChange={v => onSigChange?.("director", v)} label="Regional Director Signature" signAs={signAsFor("director")} />}
        </div>
      )}

      <div className="no-print" style={{ maxWidth: "210mm", margin: "14px auto 0", display: "flex", gap: 10, justifyContent: "flex-end", flexWrap: "wrap", alignItems: "center" }}>
        {isNew && (
          <>
            {submitError && <p style={{ color: "#c0392b", fontSize: "0.85rem", margin: "0 auto 0 0" }}>{submitError}</p>}
            <button type="button" onClick={onBack} style={{ background: "transparent", border: "1.5px solid #bbb", borderRadius: 6, padding: "8px 20px", cursor: "pointer", fontSize: "0.88rem", fontWeight: 600, color: "#333" }}>
              Cancel
            </button>
            <button type="button" onClick={onSubmit} disabled={submitting} style={{ ...btn, background: "#0038a8", padding: "8px 26px" }}>
              {submitting ? "Submitting…" : "Submit Application"}
            </button>
          </>
        )}

        {isView && (
          <>
            {actionError   && <p style={{ color: "#c0392b", fontSize: "0.85rem", margin: "0 auto 0 0" }}>{actionError}</p>}
            {actionSuccess && <p style={{ color: "#15803d", fontSize: "0.85rem", margin: "0 auto 0 0" }}>{actionSuccess}</p>}
            {s7Error   && !actionError   && <p style={{ color: "#c0392b", fontSize: "0.85rem", margin: "0 auto 0 0" }}>{s7Error}</p>}
            {s7Success && !actionSuccess && <p style={{ color: "#15803d", fontSize: "0.85rem", margin: "0 auto 0 0" }}>{s7Success}</p>}

            {isAdmin && (
              <button type="button" onClick={onSaveSection7} disabled={s7Saving} style={{ ...btn, background: "#0038a8" }}>
                {s7Saving ? "Saving…" : "Save Section 7"}
              </button>
            )}
            {hrCanAct && (
              <button type="button" onClick={() => onAction("hr_approved")} disabled={actionSaving} style={{ ...btn, background: "#15803d" }}>
                {actionSaving ? "Processing…" : "Certify & Forward to CRASD"}
              </button>
            )}
            {crasdCanAct && (
              <button type="button" onClick={() => onAction("crasd_approved")} disabled={actionSaving} style={{ ...btn, background: "#15803d" }}>
                {actionSaving ? "Processing…" : "Recommend & Forward to Director"}
              </button>
            )}
            {directorCanAct && (
              <>
                <button type="button" onClick={() => onAction("director_disapproved")} disabled={actionSaving} style={{ ...btn, background: "#b91c1c" }}>
                  {actionSaving ? "Processing…" : "Disapprove"}
                </button>
                <button type="button" onClick={() => onAction("director_approved")} disabled={actionSaving} style={{ ...btn, background: "#15803d" }}>
                  {actionSaving ? "Processing…" : "Approve"}
                </button>
              </>
            )}
          </>
        )}

        <button type="button" onClick={() => window.print()}
          style={{ background: "transparent", border: "1.5px solid #bbb", borderRadius: 6, padding: "8px 20px", cursor: "pointer", fontSize: "0.88rem", fontWeight: 600, color: "#333" }}>
          Print
        </button>
      </div>

      <div className="no-print" style={{ maxWidth: "210mm", margin: "6px auto 0", textAlign: "right", fontSize: "0.72rem", color: "var(--text-muted, #667)" }}>
        For an exact copy, print with <strong>Margins: None</strong>, <strong>Scale: 100%</strong>, and <strong>Background graphics</strong> on.
      </div>

      {isView && leaveHistory && leaveHistory.length > 0 && (
        <div className="no-print" style={{ maxWidth: "210mm", margin: "20px auto 0" }}>
          <div style={{ fontWeight: 700, fontSize: "0.88rem", color: "var(--text-label)", textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 10 }}>
            Application History
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 0, borderLeft: "2px solid var(--border-2)", paddingLeft: 18 }}>
            {leaveHistory.map((h, i) => {
              const actionLabel = {
                submitted:            "Application Submitted",
                hr_approved:          "Certified by HR — Forwarded to CRASD",
                crasd_approved:       "Recommended by CRASD — Forwarded to Director",
                director_approved:    "APPROVED by Regional Director",
                director_disapproved: "DISAPPROVED by Regional Director",
              }[h.action] ?? h.action;
              const dot = {
                submitted: "#6b7280", hr_approved: "#2563eb", crasd_approved: "#16a34a",
                director_approved: "#15803d", director_disapproved: "#b91c1c",
              }[h.action] ?? "#6b7280";
              return (
                <div key={h.id} style={{ position: "relative", paddingBottom: i < leaveHistory.length - 1 ? 18 : 0, paddingLeft: 4 }}>
                  <div style={{ position: "absolute", left: -25, top: 3, width: 12, height: 12, borderRadius: "50%", background: dot, border: "2px solid #fff", boxShadow: `0 0 0 2px ${dot}` }} />
                  <div style={{ fontSize: "0.82rem", fontWeight: 700, color: "var(--text)" }}>{actionLabel}</div>
                  <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", marginTop: 2 }}>
                    {h.actor_full_name || h.actor_username}
                    {h.actor_full_name && ` (${h.actor_username})`}
                    {" · "}
                    <span style={{ textTransform: "capitalize" }}>{h.actor_role}</span>
                    {" · "}
                    {new Date(h.created_at).toLocaleString("en-PH", { year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

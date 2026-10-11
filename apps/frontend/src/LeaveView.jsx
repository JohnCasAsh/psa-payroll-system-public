import { useCallback, useEffect, useMemo, useState } from "react";
import LeaveForm, { eligibleSignatories } from "./LeaveForm.jsx";

const apiUrl = import.meta.env.VITE_API_URL ?? "/api";

const EMPTY_FIELDS = {
  office_department: "", last_name: "", first_name: "", middle_name: "",
  date_of_filing: "", position: "", salary: "",
  leave_type: "", leave_type_other: "",
  details_vacation_local: "", details_vacation_abroad: "",
  details_sick_in_hospital: "", details_sick_out_patient: "",
  details_special_women_illness: "",
  details_study_masters: false, details_study_bar: false, details_study_other: "",
  details_monetization: false, details_terminal: false,
  num_working_days: "", inclusive_dates: "", commutation: "",
};

const S7_KEYS = {
  hr:       ["credits_as_of","vl_total_earned","vl_less_application","vl_balance","sl_total_earned","sl_less_application","sl_balance"],
  crasd:    ["recommendation","disapproval_reason"],
  director: ["approved_days_with_pay","approved_days_without_pay","approved_others_specify","disapproved_due_to"],
};
const ALL_S7_KEYS = [...S7_KEYS.hr, ...S7_KEYS.crasd, ...S7_KEYS.director];
const EMPTY_SIGS = { applicant: "", hr: "", crasd: "", director: "" };
const EMPTY_MSG  = { submitError: "", s7Error: "", s7Success: "", actionError: "", actionSuccess: "" };

const STATUS = {
  pending:        { label: "Pending HR",       bg: "#fffbeb", color: "#92400e" },
  hr_approved:    { label: "Pending CRASD",    bg: "#eff6ff", color: "#1e40af" },
  crasd_approved: { label: "Pending Director", bg: "#f5f3ff", color: "#5b21b6" },
  approved:       { label: "Approved",         bg: "#f0fdf4", color: "#166534" },
  disapproved:    { label: "Disapproved",      bg: "#fef2f2", color: "#991b1b" },
};

// Shown to the officer after they act; the application then leaves their list.
const AFTER_ACTION = {
  hr_approved:    "certified and sent to the Chief for recommendation.",
  crasd_approved: "recommended and sent to the Regional Director.",
  approved:       "approved. The applicant will see it on their Leave page.",
  disapproved:    "disapproved. The applicant will see the reason on their Leave page.",
};

const isSick   = (a) => (a.leave_type ?? "").split(",").map(s => s.trim()).includes("Sick Leave");
const sickKind = (a) => a.details_sick_in_hospital?.trim() ? "In Hospital" : a.details_sick_out_patient?.trim() ? "Out Patient" : "";
const days     = (a) => Number.parseFloat(a.num_working_days) || 0;
const today    = () => new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });

async function api(path, opts = {}) {
  const res = await fetch(`${apiUrl}${path}`, {
    credentials: "include",
    headers: opts.body ? { "Content-Type": "application/json" } : undefined,
    ...opts,
    body: opts.body ? JSON.stringify(opts.body) : undefined,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error ?? `Request failed (${res.status}).`);
  return data;
}

// Only while the form is open: print the sheet alone, not the sidebar/topbar.
const PRINT_ONLY_SHEET = `
@media print {
  body * { visibility: hidden !important; }
  .cs-page, .cs-page * { visibility: visible !important; }
  .cs-page { position: absolute !important; left: 0; top: 0; }
}`;

// refreshSignal: bumps when a leave notification arrives (reload the list quietly).
// focusId: a leave to open right away (from a pop-up's "Open" / "Review").
export default function LeaveView({ currentUser, isSuperAdmin, permissions = [], refreshSignal = 0, focusId = null, onFocusDone }) {
  const can = useCallback((k) => isSuperAdmin || permissions.includes(k), [isSuperAdmin, permissions]);
  const isAdmin = can("leave.manage");
  const me = (currentUser ?? "").toLowerCase();

  const [apps, setApps]         = useState([]);
  const [loading, setLoading]   = useState(true);
  const [listError, setListError] = useState("");
  const [mode, setMode]         = useState("list");
  const [current, setCurrent]   = useState(null);
  const [history, setHistory]   = useState([]);
  const [fields, setFields]     = useState(EMPTY_FIELDS);
  const [s7, setS7]             = useState({});
  const [sigs, setSigs]         = useState(EMPTY_SIGS);
  const [signatories, setSignatories] = useState([]);
  const [signerIds, setSignerIds]     = useState({});
  const [newDivisionId, setNewDivisionId] = useState(null);
  const [notice, setNotice]           = useState("");
  const [busy, setBusy]         = useState({ submit: false, s7: false, action: false });
  const [msg, setMsg]           = useState(EMPTY_MSG);
  const [filter, setFilter]     = useState("all");
  const [query, setQuery]       = useState("");

  const setField = useCallback((k, v) => setFields(p => ({ ...p, [k]: v })), []);
  const setS7Field = useCallback((k, v) => setS7(p => ({ ...p, [k]: v })), []);
  const setSig = useCallback((k, v) => setSigs(p => ({ ...p, [k]: v })), []);
  const setSigner = useCallback((role, id) => setSignerIds(p => ({ ...p, [role]: id })), []);

  const load = useCallback(async (silent = false) => {
    if (!silent) setLoading(true);
    setListError("");
    try { setApps(await api("/leave-applications")); }
    catch (e) { setListError(e.message); }
    finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);
  // A new pop-up means something moved along the chain: refresh without a spinner.
  useEffect(() => { if (refreshSignal) load(true); }, [refreshSignal, load]);
  useEffect(() => {
    api("/leave-signatories").then(setSignatories).catch(() => setSignatories([]));
  }, []);

  // The one approver role this user can act as on this application right now.
  // Nobody reviews their own leave, and the superadmin is a system account,
  // not an officer, so it never signs. The server enforces the same rules.
  // An approver only gets applications routed to their division: an eligible
  // name set to their own division or to all divisions (`mine`, from the server).
  const roleFor = useCallback((app) => {
    if (!app || app.mine || app.submitted_by === me) return "employee";
    if (isSuperAdmin) return "viewer";
    const routedToMe = (role) => {
      if (!signatories.some(s => s.role === role)) return true;
      return eligibleSignatories(signatories, role, app.applicant_division_id).some(s => s.mine);
    };
    const ws = app.workflow_status;
    if (ws === "pending"        && can("leave.hr")       && routedToMe("hr"))       return "hr";
    if (ws === "hr_approved"    && can("leave.crasd")    && routedToMe("crasd"))    return "crasd";
    if (ws === "crasd_approved" && can("leave.director") && routedToMe("director")) return "director";
    return "viewer";
  }, [me, isSuperAdmin, can, signatories]);

  async function startNew() {
    setMsg(EMPTY_MSG); setListError("");
    try {
      const { division_id = null, employee_id: _emp, ...snap } = await api("/leave-applications/snapshot");
      setFields({ ...EMPTY_FIELDS, ...snap, date_of_filing: today() });
      setNewDivisionId(division_id);
      setSigs(EMPTY_SIGS);
      setSignerIds({});
      setMode("new");
    } catch (e) { setListError(e.message); }
  }

  async function open(app) {
    setMsg(EMPTY_MSG); setNotice("");
    try {
      const [rec, hist] = await Promise.all([
        api(`/leave-applications/${app.id}`),
        api(`/leave-applications/${app.id}/history`),
      ]);
      setCurrent(rec);
      setHistory(hist);
      setS7(Object.fromEntries(ALL_S7_KEYS.map(k => [k, rec[k] ?? ""])));
      setSigs({ applicant: "", hr: rec.hr_signature ?? "", crasd: rec.crasd_signature ?? "", director: rec.director_signature ?? "" });
      setSignerIds({});
      setMode("view");
    } catch (e) {
      // 403 here usually means the step was already handled and it moved on.
      setListError(e.message === "Access denied." ? "That application is no longer waiting for you. It may already have been handled." : e.message);
    }
  }

  // Opened from a pop-up: go straight to that application.
  useEffect(() => {
    if (!focusId) return;
    open({ id: focusId });
    onFocusDone?.();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focusId]);

  function back() { setMode("list"); setCurrent(null); setMsg(EMPTY_MSG); }

  async function submit() {
    const types = fields.leave_type.split(",").map(s => s.trim()).filter(Boolean);
    const err =
      types.length === 0 ? "Please tick at least one type of leave (6.A)." :
      !String(fields.num_working_days).trim() ? "Please enter the number of working days (6.C) — use the 📅 picker." :
      types.includes("Sick Leave") && !fields.details_sick_in_hospital.trim() && !fields.details_sick_out_patient.trim()
        ? "For Sick Leave, tick In Hospital or Out Patient and specify the illness (6.B)." :
      !sigs.applicant ? "Please sign the application below the form." : "";
    if (err) { setMsg(m => ({ ...m, submitError: err })); return; }

    setBusy(b => ({ ...b, submit: true })); setMsg(EMPTY_MSG);
    try {
      await api("/leave-applications", { method: "POST", body: { ...fields, applicant_signature: sigs.applicant } });
      await load();
      setMode("list");
    } catch (e) { setMsg(m => ({ ...m, submitError: e.message })); }
    finally { setBusy(b => ({ ...b, submit: false })); }
  }

  async function act(action) {
    const role = roleFor(current);
    const keys = S7_KEYS[role] ?? [];
    if (action === "director_disapproved" && !s7.disapproved_due_to?.trim()) {
      setMsg(m => ({ ...m, actionError: "State the reason under 7.D (Disapproved due to) before disapproving." })); return;
    }
    if (!sigs[role]) { setMsg(m => ({ ...m, actionError: "Please sign your section below the form first." })); return; }
    if (action.startsWith("director_") && !window.confirm(`${action === "director_approved" ? "Approve" : "Disapprove"} this leave application? This is final.`)) return;

    setBusy(b => ({ ...b, action: true })); setMsg(EMPTY_MSG);
    try {
      const body = { action, signature: sigs[role], ...Object.fromEntries(keys.map(k => [k, s7[k] ?? ""])) };
      if (signerIds[role]) body.signatory_id = signerIds[role];
      const rec = await api(`/leave-applications/${current.id}/action`, { method: "POST", body });
      // The application now moves to the next officer in the chain and leaves
      // this list, so go back with a confirmation instead of staying on it.
      const who = [rec.last_name, rec.first_name].filter(Boolean).join(", ");
      setNotice(`${who}'s leave: ${AFTER_ACTION[rec.workflow_status] ?? "Done."}`);
      setMode("list"); setCurrent(null);
      load();
    } catch (e) { setMsg(m => ({ ...m, actionError: e.message })); }
    finally { setBusy(b => ({ ...b, action: false })); }
  }

  async function saveS7() {
    setBusy(b => ({ ...b, s7: true })); setMsg(EMPTY_MSG);
    try {
      const rec = await api(`/leave-applications/${current.id}`, { method: "PATCH", body: s7 });
      setCurrent(rec);
      setMsg(m => ({ ...m, s7Success: "Section 7 saved." }));
      load();
    } catch (e) { setMsg(m => ({ ...m, s7Error: e.message })); }
    finally { setBusy(b => ({ ...b, s7: false })); }
  }

  // The applicant's own leave record (only their applications, never ones they
  // are reviewing): this year's totals plus the latest balance HR certified.
  const year = new Date().getFullYear();
  const myStats = useMemo(() => {
    const own = apps.filter(a => a.mine);
    const thisYear = own.filter(a => new Date(a.created_at).getFullYear() === year);
    const approved = thisYear.filter(a => a.workflow_status === "approved");
    const sick = approved.filter(isSick);
    const byType = {};
    for (const a of approved) {
      const t = (a.leave_type ?? "").trim() || "Other";
      byType[t] = (byType[t] ?? 0) + days(a);
    }
    const credited = [...own]
      .sort((a, b) => new Date(b.created_at) - new Date(a.created_at))
      .find(a => String(a.vl_balance ?? "").trim() || String(a.sl_balance ?? "").trim());
    return {
      hasOwn: own.length > 0,
      filed: thisYear.length,
      approvedDays: approved.reduce((n, a) => n + days(a), 0),
      inReview: own.filter(a => !["approved","disapproved"].includes(a.workflow_status)).length,
      sickDays: sick.reduce((n, a) => n + days(a), 0),
      hospital: sick.filter(a => sickKind(a) === "In Hospital").length,
      outPatient: sick.filter(a => sickKind(a) === "Out Patient").length,
      byType: Object.entries(byType),
      credits: credited ? { vl: credited.vl_balance, sl: credited.sl_balance, asOf: credited.credits_as_of } : null,
    };
  }, [apps, year]);

  const counts = useMemo(() => ({
    all: apps.length,
    sick: apps.filter(isSick).length,
    review: apps.filter(a => !["approved","disapproved"].includes(a.workflow_status)).length,
    approved: apps.filter(a => a.workflow_status === "approved").length,
    disapproved: apps.filter(a => a.workflow_status === "disapproved").length,
    mine: apps.filter(a => roleFor(a) !== "employee" && roleFor(a) !== "viewer").length,
  }), [apps, roleFor]);

  const shown = useMemo(() => {
    const q = query.trim().toLowerCase();
    return apps.filter(a => {
      if (filter === "sick" && !isSick(a)) return false;
      if (filter === "review" && ["approved","disapproved"].includes(a.workflow_status)) return false;
      if (filter === "approved" && a.workflow_status !== "approved") return false;
      if (filter === "disapproved" && a.workflow_status !== "disapproved") return false;
      if (filter === "mine" && ["employee","viewer"].includes(roleFor(a))) return false;
      if (!q) return true;
      return [a.last_name, a.first_name, a.office_department, a.leave_type, a.inclusive_dates, a.submitted_by]
        .some(v => (v ?? "").toLowerCase().includes(q));
    });
  }, [apps, filter, query, roleFor]);

  if (mode !== "list") {
    return (
      <div style={{ padding: "4px 0 40px" }}>
        <style>{PRINT_ONLY_SHEET}</style>
        <div className="no-print" style={{ marginBottom: 14 }}>
          <button type="button" className="ghost-button" onClick={back}>← Back to Leave Applications</button>
        </div>
        <div style={{ overflowX: "auto" }}>
          <LeaveForm
            leaveView={mode}
            isAdmin={isAdmin}
            userRole={roleFor(current)}
            app={current}
            fields={fields}   onFieldChange={setField}
            s7Fields={s7}     onS7Change={setS7Field}
            sigs={sigs}       onSigChange={setSig}
            signatories={signatories} signerIds={signerIds} onSignerChange={setSigner}
            newDivisionId={newDivisionId}
            submitting={busy.submit}   submitError={msg.submitError}
            s7Saving={busy.s7}         s7Error={msg.s7Error}         s7Success={msg.s7Success}
            actionSaving={busy.action} actionError={msg.actionError} actionSuccess={msg.actionSuccess}
            leaveHistory={history}
            onSubmit={submit}
            onSaveSection7={saveS7}
            onAction={act}
            onBack={back}
          />
        </div>
      </div>
    );
  }

  const tabs = [
    ["all", "All"],
    ...(counts.mine > 0 ? [["mine", "Needs My Action"]] : []),
    ["review", "In Review"],
    ["sick", "Sick Leave"],
    ["approved", "Approved"],
    ["disapproved", "Disapproved"],
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div>
          <h2 style={{ margin: 0, fontSize: "1.25rem", color: "var(--psa-navy)" }}>Leave Applications</h2>
          <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", marginTop: 4 }}>CS Form No. 6 — HR certifies, your division's Chief recommends, the Regional Director approves. Only you and the officer at the current step can see your application.</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button type="button" className="ghost-button" onClick={() => load()}>Refresh</button>
          {can("leave.file") && <button type="button" onClick={startNew}>+ File Leave</button>}
        </div>
      </div>

      {notice && (
        <div role="status" style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderRadius: 10, border: "1px solid #bbf7d0", background: "#f0fdf4", color: "#166534", fontSize: "0.86rem", fontWeight: 600 }}>
          <span style={{ flex: 1 }}>{notice}</span>
          <button type="button" className="ghost-button" onClick={() => setNotice("")} style={{ fontSize: "0.76rem", padding: "4px 10px" }}>Dismiss</button>
        </div>
      )}

      {(can("leave.file") || myStats.hasOwn) && (
        <div style={{ border: "1px solid var(--border)", borderRadius: 10, padding: "14px 16px", background: "var(--surface)" }}>
          <div style={{ fontSize: "0.72rem", fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: 10 }}>
            My Leave · {year}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12 }}>
            {[
              ["Days approved", myStats.approvedDays],
              ["Applications filed", myStats.filed],
              ["In review", myStats.inReview],
              ["Sick leave days", myStats.sickDays, (myStats.hospital || myStats.outPatient) ? `${myStats.hospital} in hospital · ${myStats.outPatient} out patient` : ""],
            ].map(([label, value, sub]) => (
              <div key={label}>
                <div style={{ fontSize: "1.35rem", fontWeight: 800, color: "var(--psa-navy)" }}>{value}</div>
                <div style={{ fontSize: "0.75rem", color: "var(--text-muted)" }}>{label}</div>
                {sub && <div style={{ fontSize: "0.7rem", color: "var(--text-muted)" }}>{sub}</div>}
              </div>
            ))}
          </div>
          {(myStats.byType.length > 0 || myStats.credits) && (
            <div style={{ display: "flex", flexDirection: "column", gap: 4, marginTop: 12, paddingTop: 10, borderTop: "1px solid var(--border)", fontSize: "0.78rem", color: "var(--text-muted)" }}>
              {myStats.byType.length > 0 && (
                <div>Approved this year: {myStats.byType.map(([t, d]) => `${t} — ${d} day${d === 1 ? "" : "s"}`).join(" · ")}</div>
              )}
              {myStats.credits && (
                <div>
                  Latest leave balance certified by HR{myStats.credits.asOf ? ` (as of ${myStats.credits.asOf})` : ""}:{" "}
                  <strong style={{ color: "var(--text)" }}>Vacation {String(myStats.credits.vl ?? "").trim() || "—"}</strong> ·{" "}
                  <strong style={{ color: "var(--text)" }}>Sick {String(myStats.credits.sl ?? "").trim() || "—"}</strong>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        {tabs.map(([id, label]) => (
          <button key={id} type="button" onClick={() => setFilter(id)}
            className={filter === id ? undefined : "ghost-button"}
            style={{ fontSize: "0.8rem", padding: "6px 12px" }}>
            {label} ({counts[id]})
          </button>
        ))}
        <input type="search" placeholder="Search name, office, type, dates…" value={query} onChange={e => setQuery(e.target.value)}
          style={{ marginLeft: "auto", minWidth: 220, flex: "0 1 300px" }} />
      </div>

      {listError && <div style={{ color: "#b91c1c", fontSize: "0.85rem" }}>{listError}</div>}

      <div style={{ border: "1px solid var(--border)", borderRadius: 10, overflow: "hidden", background: "var(--surface)" }}>
        {loading ? (
          <div style={{ padding: 24, textAlign: "center", color: "var(--text-muted)", fontSize: "0.88rem" }}>Loading…</div>
        ) : shown.length === 0 ? (
          <div style={{ padding: 24, textAlign: "center", color: "var(--text-muted)", fontSize: "0.88rem" }}>
            {apps.length === 0 ? "No leave applications yet." : "No applications match this filter."}
          </div>
        ) : shown.map((a, i) => {
          const st = STATUS[a.workflow_status] ?? STATUS.pending;
          const role = roleFor(a);
          const needsMe = role !== "employee" && role !== "viewer";
          const kind = isSick(a) ? sickKind(a) : "";
          const illness = (a.details_sick_in_hospital || a.details_sick_out_patient || "").trim();
          return (
            <div key={a.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 16px", borderTop: i ? "1px solid var(--border)" : "none", flexWrap: "wrap" }}>
              <div style={{ flex: "1 1 260px", minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontSize: "0.9rem", color: "var(--text)" }}>
                  {[a.last_name, a.first_name].filter(Boolean).join(", ") || a.submitted_by}
                  <span style={{ fontWeight: 400, color: "var(--text-muted)", fontSize: "0.78rem" }}> · {a.office_department || "—"}</span>
                </div>
                <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", marginTop: 2 }}>
                  {a.leave_type || "—"}
                  {kind && ` (${kind}${illness ? `: ${illness}` : ""})`}
                  {a.num_working_days && ` · ${a.num_working_days} day(s)`}
                  {a.inclusive_dates && ` · ${a.inclusive_dates}`}
                </div>
              </div>
              <span style={{ fontSize: "0.72rem", fontWeight: 700, padding: "3px 10px", borderRadius: 99, background: st.bg, color: st.color, whiteSpace: "nowrap" }}>{st.label}</span>
              {needsMe && <span style={{ fontSize: "0.72rem", fontWeight: 700, color: "#92400e" }}>Your action</span>}
              <button type="button" onClick={() => open(a)} style={{ fontSize: "0.78rem", padding: "5px 14px" }}>{needsMe ? "Review" : "View"}</button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ── Leave form signatories (Settings) ───────────────────────────────────────
   The names printed under the 7.A, 7.B and Regional Director lines. Each name
   has a division; an application goes to the name for the applicant's division
   (see eligibleSignatories in LeaveForm.jsx). Signed forms keep the name they
   were signed with, so editing here never rewrites history. */
const SIGNATORY_SECTIONS = [
  { role: "hr",       code: "7.A", label: "HR Officer",        sub: "Certification of leave credits" },
  { role: "crasd",    code: "7.B", label: "CRASD Chief",       sub: "Recommendation" },
  { role: "director", code: "RD",  label: "Regional Director", sub: "Final approval (7.C / 7.D)" },
];

export function LeaveSignatoriesPanel() {
  const [rows, setRows]       = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState("");
  const [editing, setEditing] = useState(null); // { id (null = new), role, name, designation, division_id }
  const [busy, setBusy]       = useState(false);
  const [choices, setChoices] = useState({ divisions: [] });

  const load = useCallback(async () => {
    try { setRows(await api("/leave-signatories")); setError(""); }
    catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);
  useEffect(() => { load(); }, [load]);
  useEffect(() => {
    api("/leave-signatories/options").then(setChoices).catch(() => {});
  }, []);

  async function run(fn) {
    setBusy(true); setError("");
    try { await fn(); await load(); return true; }
    catch (e) { setError(e.message); return false; }
    finally { setBusy(false); }
  }

  async function save() {
    const { id, role, name, designation } = editing;
    const body = { name, designation, division_id: editing.division_id || null };
    const ok = await run(() => id
      ? api(`/leave-signatories/${id}`, { method: "PUT", body })
      : api("/leave-signatories", { method: "POST", body: { role, ...body } }));
    if (ok) setEditing(null);
  }
  const remove = (r) => {
    if (!window.confirm(`Remove ${r.name}? Forms they already signed keep their name.`)) return;
    run(() => api(`/leave-signatories/${r.id}`, { method: "DELETE" }));
  };

  // A render helper, not a component, so the inputs keep focus while typing.
  const editor = () => (
    <form className="ls-edit" onSubmit={e => { e.preventDefault(); save(); }}>
      <div className="sig-fields">
        <label className="sig-field">
          Full name
          <input value={editing.name} onChange={e => setEditing(x => ({ ...x, name: e.target.value }))}
            placeholder="e.g. ENGR. ANTONIO B. CRUZ" maxLength={150} required autoFocus />
        </label>
        <label className="sig-field">
          Position
          <input value={editing.designation} onChange={e => setEditing(x => ({ ...x, designation: e.target.value }))}
            placeholder="e.g. REGIONAL DIRECTOR, RSSO II" maxLength={200} required />
        </label>
        <label className="sig-field">
          Division
          <select value={editing.division_id ?? ""} onChange={e => setEditing(x => ({ ...x, division_id: e.target.value ? Number(e.target.value) : null }))}>
            <option value="">All divisions</option>
            {choices.divisions.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
          </select>
          <span className="ls-hint">Applications from this division go only to this person.</span>
        </label>
      </div>
      <div className="ls-edit-actions">
        <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => setEditing(null)} disabled={busy}>Cancel</button>
        <button type="submit" className="ls-save" disabled={busy}>{busy ? "Saving…" : editing.id ? "Save" : "Add name"}</button>
      </div>
    </form>
  );

  return (
    <div className="panel">
      <div className="sig-head">
        <div>
          <h3 className="panel-title"><span className="panel-title-dot" />Leave Form Signatories</h3>
          <div className="um-sub">The names printed under the Section 7 signature lines of CS Form No. 6. Set a division so that division's applications go only to this person; a chief from another division won't receive them.</div>
        </div>
      </div>
      {error && <div className="ls-error">{error}</div>}
      {loading ? (
        <div className="um-empty">Loading…</div>
      ) : (
        <div className="ls-grid">
          {SIGNATORY_SECTIONS.map(sec => {
            const list = rows.filter(r => r.role === sec.role);
            const addingHere = editing && !editing.id && editing.role === sec.role;
            return (
              <div key={sec.role} className="sig-card">
                <div className="sig-card-head">
                  <span className="sig-letter">{sec.code}</span>
                  <div className="sig-card-title">
                    <div className="sig-role">{sec.label}</div>
                    <div className="um-sub">{sec.sub}</div>
                  </div>
                </div>
                <div className="ls-list">
                  {list.length === 0 && !addingHere && <div className="ls-empty">No name yet. This section can't be signed until one is added.</div>}
                  {list.map(r => editing?.id === r.id ? <div key={r.id}>{editor()}</div> : (
                    <div key={r.id} className="ls-row">
                      <div className="ls-who">
                        <div className="ls-name">{r.name}</div>
                        <div className="ls-desig">{r.designation}</div>
                        <div className="ls-tags">
                          <span className={`ls-tag${r.division_id ? " ls-tag--div" : ""}`}>{r.division_name ?? "All divisions"}</span>
                        </div>
                      </div>
                      <div className="ls-row-actions">
                        <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => setEditing({ ...r })} disabled={busy}>Edit</button>
                        {list.length > 1 && <button type="button" className="um-btn um-btn--danger um-btn--sm" onClick={() => remove(r)} disabled={busy}>Remove</button>}
                      </div>
                    </div>
                  ))}
                  {addingHere ? editor() : (
                    <button type="button" className="um-btn um-btn--neutral um-btn--sm ls-add" disabled={busy}
                      onClick={() => setEditing({ id: null, role: sec.role, name: "", designation: "", division_id: null })}>
                      + Add another name
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
      <div className="sig-foot">
        <span className="um-sub">Forms that are already signed keep the name they were signed with. Changes here only apply to sections that haven't been signed yet.</span>
      </div>
    </div>
  );
}

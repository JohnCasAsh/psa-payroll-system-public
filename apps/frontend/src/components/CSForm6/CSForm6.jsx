import React, { useRef, useState, useEffect, useLayoutEffect } from "react";
import { createPortal } from "react-dom";
import "./CSForm6.css";

/* ═══════════════════════════════════════════════════════════════════════
   CS Form No. 6 — direct port of cs-form-6-reference.html.

   The reference is the source of truth. Its inline styles were moved
   verbatim into CSForm6.css classes (safer than JSX style objects,
   which silently turn unitless numbers into px and destroy the pt-based
   geometry). There are NO inline styles in this file.

   Fill mode and print mode render the SAME elements with the same
   heights and borders — print mode only sets readOnly and the is-print
   class. Elements are never swapped for <span>, and empty fields still
   render at full height, so a blank and a filled form are dimensionally
   identical.
   ═══════════════════════════════════════════════════════════════════════ */

const LEAVE_TYPES = [
  ["Vacation Leave", "Vacation Leave", "(Sec. 51, Rule XVI, Omnibus Rules Implementing E.O. No. 292)"],
  ["Mandatory/Forced Leave", "Mandatory/Forced Leave", "(Sec. 25, Rule XVI, Omnibus Rules Implementing E.O. No. 292)"],
  ["Sick Leave", "Sick Leave", "(Sec. 43, Rule XVI, Omnibus Rules Implementing E.O. No. 292)"],
  ["Maternity Leave", "Maternity Leave", "(R.A. No. 11210 / IRR issued by CSC, DOLE and SSS)"],
  ["Paternity Leave", "Paternity Leave", "(R.A. No. 8187 / CSC MC No. 71, s. 1998, as amended)"],
  ["Special Privilege Leave", "Special Privilege Leave", "(Sec. 21, Rule XVI, Omnibus Rules Implementing E.O. No. 292)"],
  ["Solo Parent Leave", "Solo Parent Leave", "(RA No. 8972 / CSC MC No. 8, s. 2004)"],
  ["Study Leave", "Study Leave", "(Sec. 68, Rule XVI, Omnibus Rules Implementing E.O. No. 292)"],
  ["10-Day VAWC Leave", "10-Day VAWC Leave", "(RA No. 9262 / CSC MC No. 15, s. 2005)"],
  ["Rehabilitation Privilege", "Rehabilitation Privilege", "(Sec. 55, Rule XVI, Omnibus Rules Implementing E.O. No. 292)"],
  ["Special Leave Benefits for Women", "Special Leave Benefits for Women", "(RA No. 9710 / CSC MC No. 25, s. 2010)"],
  ["Special Emergency (Calamity) Leave", "Special Emergency (Calamity) Leave", "(CSC MC No. 2, s. 2012, as amended)"],
  ["Adoption Leave", "Adoption Leave", "(R.A. No. 8552)"],
];

const WEEKDAYS = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];
const sameDay = (a, b) => !!a && !!b && a.getTime() === b.getTime();
const isWeekend = (d) => d.getDay() === 0 || d.getDay() === 6;
const startOfToday = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };

function countWorkingDays(start, end) {
  let count = 0;
  const cur = new Date(start);
  while (cur <= end) {
    const day = cur.getDay();
    if (day !== 0 && day !== 6) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

function formatInclusiveDates(start, end) {
  const monthName = d => d.toLocaleDateString("en-US", { month: "long" });
  if (start.getTime() === end.getTime()) {
    return `${monthName(start)} ${start.getDate()}, ${start.getFullYear()}`;
  }
  if (start.getFullYear() === end.getFullYear() && start.getMonth() === end.getMonth()) {
    return `${monthName(start)} ${start.getDate()}-${end.getDate()}, ${start.getFullYear()}`;
  }
  return `${monthName(start)} ${start.getDate()}, ${start.getFullYear()} - ${monthName(end)} ${end.getDate()}, ${end.getFullYear()}`;
}

const ChevronIcon = ({ dir }) => (
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <polyline points={dir === "left" ? "15 18 9 12 15 6" : "9 18 15 12 9 6"} />
  </svg>
);

const CalendarIcon = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <rect x="3" y="4.5" width="18" height="16.5" rx="2.5" />
    <line x1="16" y1="2.5" x2="16" y2="6.5" />
    <line x1="8" y1="2.5" x2="8" y2="6.5" />
    <line x1="3" y1="10" x2="21" y2="10" />
    <circle cx="8.5" cy="15" r="0.6" fill="currentColor" />
    <circle cx="12" cy="15" r="0.6" fill="currentColor" />
    <circle cx="15.5" cy="15" r="0.6" fill="currentColor" />
  </svg>
);

// First click sets the start day, second click (or the same day again, for a
// one-day leave) sets the end and fills both 6.C fields. Weekends are muted
// because they are not counted as working days.
function DateRangePicker({ onApply }) {
  const today = startOfToday();
  const [month, setMonth] = useState(() => new Date(today.getFullYear(), today.getMonth(), 1));
  const [start, setStart] = useState(null);
  const [hover, setHover] = useState(null);

  const year = month.getFullYear();
  const m = month.getMonth();
  const lead = new Date(year, m, 1).getDay();
  const total = new Date(year, m + 1, 0).getDate();
  const cells = [...Array(lead).fill(null), ...Array.from({ length: total }, (_, i) => new Date(year, m, i + 1))];

  // While the end date is being chosen, preview the range up to the hovered day.
  const end = start && hover && hover >= start ? hover : null;
  const hi = end ?? start;

  function pick(date) {
    if (!start || date < start) { setStart(date); return; }
    onApply(start, date);
  }

  const shortDate = (d) => d.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" });
  const n = end ? countWorkingDays(start, end) : 0;
  const footer = !start
    ? <>Click the <strong>first day</strong> of leave.</>
    : end
      ? <><strong>{formatInclusiveDates(start, end)}</strong> · {n} working day{n === 1 ? "" : "s"}</>
      : <>From <strong>{shortDate(start)}</strong> — now click the <strong>last day</strong> (or the same day for one day).</>;

  return (
    <div className="cs-dp no-print" role="dialog" aria-label="Choose leave dates">
      <div className="cs-dp-head">
        <button type="button" className="cs-dp-nav" aria-label="Previous month"
          onClick={() => setMonth(new Date(year, m - 1, 1))}><ChevronIcon dir="left" /></button>
        <div className="cs-dp-title">{month.toLocaleDateString("en-US", { month: "long", year: "numeric" })}</div>
        <button type="button" className="cs-dp-nav" aria-label="Next month"
          onClick={() => setMonth(new Date(year, m + 1, 1))}><ChevronIcon dir="right" /></button>
      </div>

      <div className="cs-dp-grid" onMouseLeave={() => setHover(null)}>
        {WEEKDAYS.map((d, i) => (
          <div key={d} className={`cs-dp-wd${i === 0 || i === 6 ? " is-weekend" : ""}`}>{d}</div>
        ))}
        {cells.map((date, i) => {
          if (!date) return <div key={i} />;
          const inRange = start && date >= start && date <= hi;
          const isStart = sameDay(date, start);
          const isEnd = sameDay(date, hi);
          const cell = ["cs-dp-cell", inRange && "in-range", isStart && "is-start", isEnd && "is-end"].filter(Boolean).join(" ");
          const btn = ["cs-dp-day", isWeekend(date) && "is-weekend", sameDay(date, today) && "is-today",
                       isStart && "is-start", isEnd && "is-end"].filter(Boolean).join(" ");
          return (
            <div key={i} className={cell}>
              <button type="button" className={btn}
                onClick={() => pick(date)}
                onMouseEnter={() => setHover(date)}
                aria-label={date.toDateString()}
                aria-pressed={!!inRange}>
                {date.getDate()}
              </button>
            </div>
          );
        })}
      </div>

      <div className="cs-dp-foot">{footer}</div>
    </div>
  );
}

// Collapses DateRangePicker behind a small icon button, anchored to the
// 6.C "Inclusive Dates" field it fills. 6.C's column has overflow:hidden
// (needed so long print content clips the same way it does on paper), so
// the dropdown is portaled to <body> and positioned from the button's own
// screen coordinates — that escapes the clip instead of fighting it, and
// still never shifts the surrounding print-exact layout.
function InclusiveDatesPicker({ onApply }) {
  const [open, setOpen] = useState(false);
  const [rect, setRect] = useState(null);
  const btnRef = useRef(null);
  const popRef = useRef(null);

  useLayoutEffect(() => {
    if (!open) return;
    const updateRect = () => setRect(btnRef.current?.getBoundingClientRect() ?? null);
    updateRect();
    window.addEventListener("scroll", updateRect, true);
    window.addEventListener("resize", updateRect);
    return () => {
      window.removeEventListener("scroll", updateRect, true);
      window.removeEventListener("resize", updateRect);
    };
  }, [open]);

  useEffect(() => {
    if (!open) return;
    function onDown(e) {
      if (btnRef.current?.contains(e.target)) return;
      if (popRef.current && !popRef.current.contains(e.target)) setOpen(false);
    }
    function onKey(e) { if (e.key === "Escape") setOpen(false); }
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const POPUP_WIDTH = 272;
  const POPUP_HEIGHT = 340;
  // Open upward when there is not enough room below the button.
  const top = rect && rect.bottom + 6 + POPUP_HEIGHT > window.innerHeight && rect.top - 6 - POPUP_HEIGHT > 0
    ? rect.top - 6 - POPUP_HEIGHT
    : (rect?.bottom ?? 0) + 6;

  return (
    <span className="no-print cs-incl-pick">
      <button
        ref={btnRef}
        type="button"
        className={`cs-dp-trigger${open ? " is-open" : ""}`}
        onClick={() => setOpen(v => !v)}
        title="Pick leave dates"
        aria-label="Pick leave dates"
        aria-expanded={open}
      >
        <CalendarIcon />
      </button>
      {open && rect && createPortal(
        <div
          ref={popRef}
          className="no-print"
          style={{
            position: "fixed",
            top,
            left: Math.max(8, Math.min(rect.right - POPUP_WIDTH, window.innerWidth - POPUP_WIDTH - 8)),
            zIndex: 1000,
          }}
        >
          <DateRangePicker onApply={(s, e) => { onApply(s, e); setOpen(false); }} />
        </div>,
        document.body
      )}
    </span>
  );
}

const HISTORY_ACTION_LABELS = {
  submitted: "Application Submitted",
  hr_approved: "Certified by HR — Forwarded to CRASD",
  crasd_approved: "Recommended by CRASD — Forwarded to Director",
  director_approved: "APPROVED by Regional Director",
  director_disapproved: "DISAPPROVED by Regional Director",
};

export default function CSForm6({
  data = {},
  mode = "print",
  onChange,
  canEdit,
  signatures = {},
  // { hr, crasd, director } -> { name, designation } printed under each line
  officers = {},
  logoSrc = "/PSA.png",
  history = [],
}) {
  const editable = canEdit ?? (() => mode === "fill");
  const ro = (field) => !editable(field);

  // Controlled text input. `checked`/`value` always paired with onChange
  // so React never warns or silently makes the field read-only.
  const T = (field, className) => (
    <input
      type="text"
      className={className}
      value={data[field] ?? ""}
      readOnly={ro(field)}
      onChange={ro(field) ? undefined : (e) => onChange?.(field, e.target.value)}
      autoComplete="off"
      spellCheck={false}
    />
  );

  // A decorative continuation line that exists on the paper form but
  // holds no data — always read-only, never bound.
  const Blank = ({ className }) => (
    <input type="text" className={className} value="" readOnly />
  );

  const C = (checked, onToggle, field) => (
    <input
      type="checkbox"
      checked={!!checked}
      readOnly={ro(field)}
      onChange={ro(field) ? undefined : () => onToggle()}
    />
  );

  // Printed name under the applicant's signature, styled like the approvers' (e.g. "JUAN D. DELA CRUZ").
  const middle = (data.middle_name ?? "").trim();
  const applicantName = [data.first_name, /^[a-z]$/i.test(middle) ? `${middle}.` : middle, data.last_name]
    .map((s) => (s ?? "").trim()).filter(Boolean).join(" ").toUpperCase();

  const commutation   = data.commutation ?? "";
  const recommendation = data.recommendation ?? "";

  /* 6.A allows MORE THAN ONE type of leave to be ticked. They are stored
     in the single existing `leave_type` column as a comma-separated list,
     so no schema change is needed and older single-value records keep
     working (a lone value parses to a one-item list). */
  const leaveTypes = (data.leave_type ?? "").split(",").map((s) => s.trim()).filter(Boolean);
  const hasLeaveType = (k) => leaveTypes.includes(k);
  const writeLeaveTypes = (list) => onChange?.("leave_type", list.join(", "));
  const setLeaveType = (k) => () =>
    writeLeaveTypes(hasLeaveType(k) ? leaveTypes.filter((x) => x !== k) : [...leaveTypes, k]);
  const setBool         = (f) => () => onChange?.(f, !data[f]);
  // A single space marks a detail line checked with nothing typed yet.
  const setDetail       = (f) => () => onChange?.(f, data[f] ? "" : " ");
  const setCommutation  = (v) => () => onChange?.("commutation", commutation === v ? "" : v);
  const setRecommend    = (v) => () => onChange?.("recommendation", recommendation === v ? "" : v);

  return (
    <div className="cs-sheet">
      <div className={`cs-page${mode === "fill" ? "" : " is-print"}`}>

        <div className="cs-hdr">
          <div className="cs-formid">
            <div>Civil Service Form No. 6</div>
            <div>Revised 2020</div>
          </div>
          <img className="cs-logo" src={logoSrc} alt="" />
          <div className="cs-agency">
            <div className="cs-agency-1">Republic of the Philippines</div>
            <div className="cs-agency-2">PHILIPPINE STATISTICS AUTHORITY</div>
            <div className="cs-agency-3">PSA Complex, East Avenue, Diliman, Quezon City 1100</div>
          </div>
          <div className="cs-stamp">Stamp of Date of Receipt</div>
        </div>

        <div className="cs-title">APPLICATION FOR LEAVE</div>

        <div className="cs-box">

          {/* Row 1 — Office / Name. 36/64 split (not the section 55/45) so
              "2. NAME" sits directly above "4. POSITION"; no divider. */}
          <div className="cs-row">
            <div className="cs-col-l cs-r1 cs-p1">
              <div>1.&nbsp;&nbsp;&nbsp;OFFICE/DEPARTMENT</div>
              {T("office_department", "cs-in-office")}
            </div>
            <div className="cs-col-r cs-r1 cs-p1">
              <div className="cs-nm-hdr">
                <span className="cs-nm-lead">2.&nbsp;&nbsp;NAME :</span>
                <span className="cs-nm-sub">
                  <span>(Last)</span><span>(First)</span><span>(Middle)</span>
                </span>
              </div>
              <div className="cs-nm-inputs">
                {T("last_name")}
                {T("first_name")}
                {T("middle_name")}
              </div>
            </div>
          </div>

          {/* Row 2 — Filing / Position / Salary */}
          <div className="cs-row2">
            <div className="cs-fg cs-w36"><span>3.&nbsp;&nbsp;&nbsp;DATE OF FILING</span>{T("date_of_filing")}</div>
            <div className="cs-fg cs-w41"><span>4.&nbsp;&nbsp;&nbsp;POSITION</span>{T("position")}</div>
            <div className="cs-fg cs-w23"><span>5. SALARY</span>{T("salary")}</div>
          </div>

          <div className="cs-banner">6.&nbsp;&nbsp;DETAILS OF APPLICATION</div>

          {/* No rule under the 6.A / 6.B subheads — they read as one cell
              with the content below, as on the form. */}
          <div className="cs-row-nb">
            <div className="cs-col-l cs-sub">6.A&nbsp;&nbsp;TYPE OF LEAVE TO BE AVAILED OF</div>
            <div className="cs-col-r cs-sub">6.B&nbsp;&nbsp;DETAILS OF LEAVE</div>
          </div>

          {/* 6.A / 6.B body */}
          <div className="cs-row">
            <div className="cs-col-l cs-6a">
              {LEAVE_TYPES.map(([key, label, cite]) => (
                <label key={key}>
                  {C(hasLeaveType(key), setLeaveType(key), "leave_type")}
                  <span>{label} <span className="cs-cite">{cite}</span></span>
                </label>
              ))}
              <div className="cs-others">Others:</div>
              {/* Typing here adds "Others" to the list; clearing it removes it. */}
              <input
                type="text"
                value={data.leave_type_other ?? ""}
                readOnly={ro("leave_type_other")}
                onChange={ro("leave_type_other") ? undefined : (e) => {
                  const v = e.target.value;
                  onChange?.("leave_type_other", v);
                  const has = hasLeaveType("Others");
                  if (v.trim() && !has) writeLeaveTypes([...leaveTypes, "Others"]);
                  else if (!v.trim() && has) writeLeaveTypes(leaveTypes.filter((x) => x !== "Others"));
                }}
                autoComplete="off"
                spellCheck={false}
              />
            </div>

            <div className="cs-col-r cs-6b">
              <div className="cs-ital">In case of Vacation/Special Privilege Leave:</div>
              <label className="cs-drow">
                {C(!!data.details_vacation_local, setDetail("details_vacation_local"), "details_vacation_local")}
                <span>Within the Philippines</span>
                {T("details_vacation_local")}
              </label>
              <label className="cs-drow">
                {C(!!data.details_vacation_abroad, setDetail("details_vacation_abroad"), "details_vacation_abroad")}
                <span>Abroad (Specify)</span>
                {T("details_vacation_abroad")}
              </label>

              <div className="cs-ital">In case of Sick Leave:</div>
              <label className="cs-drow">
                {C(!!data.details_sick_in_hospital, setDetail("details_sick_in_hospital"), "details_sick_in_hospital")}
                <span>In Hospital (Specify Illness)</span>
                {T("details_sick_in_hospital")}
              </label>
              <label className="cs-drow">
                {C(!!data.details_sick_out_patient, setDetail("details_sick_out_patient"), "details_sick_out_patient")}
                <span>Out Patient (Specify Illness)</span>
                {T("details_sick_out_patient")}
              </label>
              <Blank />

              <div className="cs-ital-mt">In case of Special Leave Benefits for Women:</div>
              <div className="cs-drow">
                <span>(Specify Illness)</span>
                {T("details_special_women_illness")}
              </div>
              <Blank />

              <div className="cs-ital-mt">In case of Study Leave:</div>
              <label className="cs-drow">
                {C(!!data.details_study_masters, setBool("details_study_masters"), "details_study_masters")}
                <span>Completion of Master&apos;s Degree</span>
              </label>
              <label className="cs-drow">
                {C(!!data.details_study_bar, setBool("details_study_bar"), "details_study_bar")}
                <span>BAR/Board Examination Review</span>
              </label>

              <div className="cs-ital-mt">Other purpose:</div>
              <label className="cs-drow">
                {C(!!data.details_monetization, setBool("details_monetization"), "details_monetization")}
                <span>Monetization of Leave Credits</span>
              </label>
              <label className="cs-drow">
                {C(!!data.details_terminal, setBool("details_terminal"), "details_terminal")}
                <span>Terminal Leave</span>
              </label>
            </div>
          </div>

          <div className="cs-row-nb">
            <div className="cs-col-l cs-sub">6.C&nbsp;&nbsp;NUMBER OF WORKING DAYS APPLIED FOR</div>
            <div className="cs-col-r cs-sub">6.D&nbsp;&nbsp;COMMUTATION</div>
          </div>

          {/* 6.C / 6.D body — no bottom rule; the section 7 band supplies it */}
          <div className="cs-row-nb">
            <div className="cs-col-l cs-6c">
              {T("num_working_days")}
              <div className="cs-incl">INCLUSIVE DATES</div>
              <div className="cs-incl-row">
                {T("inclusive_dates")}
                {editable("inclusive_dates") && (
                  <InclusiveDatesPicker
                    onApply={(start, end) => {
                      onChange?.("num_working_days", String(countWorkingDays(start, end)));
                      onChange?.("inclusive_dates", formatInclusiveDates(start, end));
                    }}
                  />
                )}
              </div>
            </div>
            <div className="cs-col-r cs-6d">
              <label className="cs-drow">
                {C(commutation === "not_requested", setCommutation("not_requested"), "commutation")}
                <span>Not Requested</span>
              </label>
              <label className="cs-drow">
                {C(commutation === "requested", setCommutation("requested"), "commutation")}
                <span>Requested</span>
              </label>
              <div className="cs-sig6d">
                <div className="cs-sig-img">{signatures.applicant ? <img src={signatures.applicant} alt="" /> : null}</div>
                <input type="text" value={applicantName} readOnly tabIndex={-1} aria-label="Applicant name" />
                <div className="cs-rule" />
                <div className="cs-cap">(Signature of Applicant)</div>
              </div>
            </div>
          </div>

          <div className="cs-banner">7.&nbsp;&nbsp;DETAILS OF ACTION ON APPLICATION</div>

          <div className="cs-row-nb">
            <div className="cs-col-l cs-sub">7.A&nbsp;&nbsp;CERTIFICATION OF LEAVE CREDITS</div>
            <div className="cs-col-r cs-sub">7.B&nbsp;&nbsp;RECOMMENDATION</div>
          </div>

          {/* 7.A / 7.B body — closed by a heavy rule above 7.C / 7.D */}
          <div className="cs-row-thick">
            <div className="cs-col-l cs-7a">
              <div className="cs-asof">As of {T("credits_as_of")}</div>
              <table className="cs-ctab">
                <tbody>
                  <tr>
                    <td className="cs-c0"></td>
                    <td className="cs-th">Vacation Leave</td>
                    <td className="cs-th">Sick Leave</td>
                  </tr>
                  <tr>
                    <td className="cs-rl">Total Earned</td>
                    <td className="cs-vc">{T("vl_total_earned")}</td>
                    <td className="cs-vc">{T("sl_total_earned")}</td>
                  </tr>
                  <tr>
                    <td className="cs-rl">Less this application</td>
                    <td className="cs-vc">{T("vl_less_application")}</td>
                    <td className="cs-vc">{T("sl_less_application")}</td>
                  </tr>
                  <tr>
                    <td className="cs-rl">Balance</td>
                    <td className="cs-vc">{T("vl_balance")}</td>
                    <td className="cs-vc">{T("sl_balance")}</td>
                  </tr>
                </tbody>
              </table>
              <div className="cs-sig cs-sig-a">
                <div className="cs-sig-img">{signatures.hr ? <img src={signatures.hr} alt="" /> : null}</div>
                <input type="text" className="cs-l1" value={officers.hr?.name ?? ""} readOnly />
                <input type="text" className="cs-l2" value={officers.hr?.designation ?? ""} readOnly />
                <div className="cs-rule" />
                <div className="cs-cap">(Authorized Officer)</div>
              </div>
            </div>

            <div className="cs-col-r cs-7b">
              <label className="cs-drow">
                {C(recommendation === "for_approval", setRecommend("for_approval"), "recommendation")}
                <span>For approval</span>
              </label>
              <label className="cs-drow">
                {C(recommendation === "for_disapproval", setRecommend("for_disapproval"), "recommendation")}
                <span>For disapproval due to</span>
                {T("disapproval_reason")}
              </label>
              {/* three continuation lines — four writing lines in total,
                  counting the one beside "For disapproval due to" */}
              <div className="cs-pl8"><Blank /></div>
              <div className="cs-pl8"><Blank /></div>
              <div className="cs-pl8"><Blank /></div>
              <div className="cs-sig">
                <div className="cs-sig-img">{signatures.crasd ? <img src={signatures.crasd} alt="" /> : null}</div>
                <input type="text" className="cs-l1" value={officers.crasd?.name ?? ""} readOnly />
                <input type="text" className="cs-l2" value={officers.crasd?.designation ?? ""} readOnly />
                <div className="cs-rule" />
                <div className="cs-cap">(Authorized Officer)</div>
              </div>
            </div>
          </div>

          <div className="cs-row-nb">
            <div className="cs-col-l cs-nodiv cs-sub">7.C&nbsp;&nbsp;APPROVED FOR:</div>
            <div className="cs-col-r cs-sub">7.D&nbsp;&nbsp;DISAPPROVED DUE TO:</div>
          </div>

          {/* 7.C / 7.D body — no rule beneath: it runs straight into the
              Director's block, as on the reference. */}
          <div className="cs-row-nb">
            <div className="cs-col-l cs-nodiv cs-7c">
              <div className="cs-arow">{T("approved_days_with_pay")}<span>days with pay</span></div>
              <div className="cs-arow">{T("approved_days_without_pay")}<span>days without pay</span></div>
              <div className="cs-arow">{T("approved_others_specify")}<span>others (Specify)</span></div>
            </div>
            <div className="cs-col-r cs-7d">
              {T("disapproved_due_to")}
              <Blank />
              <Blank />
            </div>
          </div>

          {/* Director's block sits INSIDE the form box on the official form,
              spanning the full width below the 7.C / 7.D rule. */}
          <div className="cs-foot">
            <div className="cs-sig-img">{signatures.director ? <img src={signatures.director} alt="" /> : null}</div>
            <input type="text" className="cs-l1" value={officers.director?.name ?? ""} readOnly />
            <input type="text" className="cs-l2" value={officers.director?.designation ?? ""} readOnly />
            <div className="cs-rule" />
            <div className="cs-cap">(Authorized Official)</div>
          </div>
        </div>
      </div>

      {history && history.length > 0 && (
        <div className="cs-history-page">
          <div className="cs-history-bar">APPLICATION HISTORY</div>
          {history.map((h) => (
            <div className="cs-history-item" key={h.id}>
              <div className="cs-h-label">{HISTORY_ACTION_LABELS[h.action] ?? h.action}</div>
              <div className="cs-h-meta">
                {(h.actor_full_name ? `${h.actor_full_name} (${h.actor_username})` : h.actor_username)}
                {" · "}{h.actor_role}{" · "}
                {new Date(h.created_at).toLocaleString("en-PH")}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

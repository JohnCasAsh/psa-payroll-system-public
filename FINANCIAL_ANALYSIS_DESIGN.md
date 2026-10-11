# Personal Financial Database System — Design Document

## Overview

This system goes beyond payroll printing. It tracks each employee's complete financial picture — salary, loans, investments, insurance, and savings — and analyzes behavioral patterns over time to detect whether an employee is progressive, neutral, or in downturn.

---

## The Four Layers

```
Layer 1 — Personal Financial Records
        ↓ triggers
Layer 2 — Salary Assessment
        ↓ generates
Layer 3 — Timeline Report
        ↓ archived into
Layer 4 — History Archive
```

---

## Layer 1 — Personal Financial Record

Stores every financial component per employee. Each one has a lifecycle — a start date, an optional end date, and whether it is continuous (ongoing).

```sql
CREATE TABLE personal_financial_record (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_id   UUID NOT NULL REFERENCES employee(id) ON DELETE RESTRICT,
  type          TEXT NOT NULL CHECK (type IN ('salary', 'loan', 'investment', 'insurance', 'savings')),
  name          TEXT NOT NULL,
  -- examples: 'GSIS MPL', 'Pag-IBIG Savings', 'Sun Life Insurance', 'GSIS Emergency Loan'
  amount        NUMERIC(12,2) NOT NULL,
  started_date  DATE NOT NULL,
  end_date      DATE,            -- NULL means continuous / still active
  is_continuous BOOLEAN NOT NULL DEFAULT FALSE,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
```

### Types explained

| Type | Examples |
|---|---|
| `salary` | Basic Monthly Pay, Step Increase, Salary Grade Change |
| `loan` | GSIS MPL, GSIS Emergency Loan, Pag-IBIG Multi-Purpose Loan |
| `investment` | GSIS Policy, Pag-IBIG MP2, Personal Investment |
| `insurance` | GSIS Life Insurance, Optional Insurance, Private Insurance |
| `savings` | Pag-IBIG Regular Savings, Personal Savings |

---

## Layer 2 — Salary Assessment

Every time a financial component is added, changed, or ends — the system runs a salary assessment. It computes the employee's overall financial health at that moment and labels it: **progressive**, **neutral**, or **downturn**.

```sql
CREATE TABLE salary_assessment (
  id                   UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_id          UUID NOT NULL REFERENCES employee(id) ON DELETE RESTRICT,
  record_id            UUID REFERENCES personal_financial_record(id),
  -- the record that triggered this assessment
  assessed_at          DATE NOT NULL,
  gross_salary         NUMERIC(12,2) NOT NULL,
  total_loans          NUMERIC(12,2) NOT NULL DEFAULT 0,
  total_investments    NUMERIC(12,2) NOT NULL DEFAULT 0,
  total_insurance      NUMERIC(12,2) NOT NULL DEFAULT 0,
  total_savings        NUMERIC(12,2) NOT NULL DEFAULT 0,
  net_financial_health NUMERIC(12,2) NOT NULL,
  -- gross_salary - total_loans + total_investments + total_savings
  status               TEXT NOT NULL CHECK (status IN ('progressive', 'neutral', 'downturn')),
  notes                TEXT,
  created_at           TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
```

### Status logic

| Status | Condition |
|---|---|
| `progressive` | net_financial_health improving vs previous assessment, loans shrinking or savings growing |
| `neutral` | no significant change — stable within threshold |
| `downturn` | net_financial_health declining, loans increasing, savings shrinking |

---

## Layer 3 — Timeline Report

Tracks how long an employee has been in each status. When status changes, the previous period is closed with an end date and a new period opens.

```sql
CREATE TABLE financial_timeline (
  id             UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_id    UUID NOT NULL REFERENCES employee(id) ON DELETE RESTRICT,
  assessment_id  UUID NOT NULL REFERENCES salary_assessment(id),
  status         TEXT NOT NULL CHECK (status IN ('progressive', 'neutral', 'downturn')),
  started_date   DATE NOT NULL,
  end_date       DATE,            -- NULL = currently in this status
  is_continuous  BOOLEAN NOT NULL DEFAULT TRUE,
  duration_days  INTEGER,         -- computed: end_date - started_date, NULL if ongoing
  created_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
```

### Example timeline output

```
Juan Dela Cruz — Financial Timeline
────────────────────────────────────────────────────────
Status        Started       Ended         Duration
progressive   Jan 1, 2025   Mar 15, 2025  73 days
neutral       Mar 15, 2025  Apr 1, 2025   17 days
downturn      Apr 1, 2025   present       CONTINUOUS ← danger flag
```

A `downturn` with `is_continuous = true` and no `end_date` means the employee has not recovered. This is the primary alert signal for HR intervention.

---

## Layer 4 — History Archive

Simple, human-readable summary of each completed status period. No calculations — just what happened, when, and how long.

```sql
CREATE TABLE financial_history_archive (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_id   UUID NOT NULL REFERENCES employee(id) ON DELETE RESTRICT,
  timeline_id   UUID NOT NULL REFERENCES financial_timeline(id),
  status        TEXT NOT NULL CHECK (status IN ('progressive', 'neutral', 'downturn')),
  started_date  DATE NOT NULL,
  end_date      DATE,
  duration_days INTEGER,
  summary       TEXT,
  -- plain language: "Employee took 2 new loans. Net pay reduced by ₱2,300."
  archived_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
```

### Example history archive entry

```
Employee:   Juan Dela Cruz
Status:     downturn
Started:    April 1, 2025
Ended:      still ongoing
Duration:   continuous
Summary:    Employee added GSIS Emergency Loan (₱1,200/month) and
            Pag-IBIG MPL (₱800/month). Net financial health dropped
            by ₱2,300. No recovery observed in 3 quincenas.
```

---

## Full connection diagram

```
personal_financial_record  (new loan / investment / insurance / savings entered)
        ↓
salary_assessment          (system computes: progressive / neutral / downturn)
        ↓
financial_timeline         (tracks: how long in this status, start date, end date)
        ↓
financial_history_archive  (permanent record: what happened, when, how long)
```

---

## What this enables

| Query | Answer |
|---|---|
| Who is currently in downturn? | `SELECT * FROM financial_timeline WHERE status='downturn' AND end_date IS NULL` |
| How long has this person been struggling? | `duration_days` on their open timeline record |
| Has anyone recovered this year? | Timeline records where status changed from downturn → progressive |
| Who has the highest loan burden? | `total_loans / gross_salary` ratio from latest salary_assessment |
| Who is progressively investing? | Records where investment amounts increase across assessments |

---

## Notes

- Records in `personal_financial_record`, `financial_timeline`, and `financial_history_archive` use `ON DELETE RESTRICT` — they can never be deleted. Financial history is permanent.
- This connects to `payroll_history` (built by @mistomu) — payroll data feeds into salary assessments automatically each quincena.
- Leave Without Pay deductions (built by @ajay-dev-ctrl) and salary grade rates (built by @alyror-sd) both feed into the `gross_salary` and `net_financial_health` calculations.
- Security and audit logging (@TheAFDesu233) applies to all assessment writes — every status change is logged.

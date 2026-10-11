/**
 * Sync deductions from the most recent payroll archive → live deduction table.
 *
 * Reads payroll_archive snapshot (which contains per-employee deduction arrays)
 * and upserts each amount into the `deduction` table so that MY PAYSLIP OVERVIEW
 * shows real figures instead of ₱0.00.
 *
 * Usage:
 *   node scripts/sync-deductions-from-archive.mjs
 *   MYSQL_URL=mysql://root:@127.0.0.1:3306/psa_payroll node scripts/sync-deductions-from-archive.mjs
 *
 * Optional — target a specific period:
 *   YEAR=2026 MONTH=6 QUINCENA=2 node scripts/sync-deductions-from-archive.mjs
 */
import mysql from "mysql2/promise";
import { randomUUID } from "crypto";

const pool = mysql.createPool({
  uri: process.env.MYSQL_URL ?? "mysql://root:@127.0.0.1:3306/psa_payroll",
  waitForConnections: true,
  connectionLimit: 5,
});

async function q(sql, params = []) {
  const [rows] = await pool.execute(sql, params);
  return rows;
}

async function main() {
  // ── 1. Pick the archive to restore from ─────────────────────────────────────
  let archive;

  if (process.env.YEAR && process.env.MONTH && process.env.QUINCENA) {
    const rows = await q(
      "SELECT * FROM `payroll_archive` WHERE year=? AND month=? AND quincena=?",
      [Number(process.env.YEAR), Number(process.env.MONTH), Number(process.env.QUINCENA)]
    );
    archive = rows[0];
    if (!archive) {
      console.error(`No archive found for ${process.env.YEAR}-${process.env.MONTH} Q${process.env.QUINCENA}`);
      process.exit(1);
    }
  } else {
    // Most recent archive
    const rows = await q(
      "SELECT * FROM `payroll_archive` ORDER BY year DESC, month DESC, quincena DESC LIMIT 1"
    );
    archive = rows[0];
    if (!archive) {
      console.log("No payroll archives in database. Nothing to sync.");
      await pool.end();
      return;
    }
  }

  const snapshot = JSON.parse(archive.snapshot);
  const items = snapshot.items ?? [];
  const period = `${archive.month}/${archive.year} Q${archive.quincena}`;

  console.log(`\nSyncing from archive: ${period}`);
  console.log(`Archive saved by: ${archive.saved_by}`);
  console.log(`Items in snapshot: ${items.length}\n`);

  // ── 2. Check if payroll_history_deduction has richer data ───────────────────
  const histRows = await q(`
    SELECT ph.employee_id, phd.deductiontype_id, phd.amount
    FROM payroll_history ph
    JOIN payroll_history_deduction phd ON phd.payroll_history_id = ph.id
    WHERE ph.year=? AND ph.month=? AND ph.quincena=? AND phd.amount > 0
  `, [archive.year, archive.month, archive.quincena]);

  let source, deductionEntries;

  if (histRows.length > 0) {
    source = "payroll_history_deduction";
    deductionEntries = histRows.map(r => ({
      employee_id: r.employee_id,
      deductiontype_id: r.deductiontype_id,
      amount: Number(r.amount),
    }));
    console.log(`Source: payroll_history_deduction (${deductionEntries.length} rows)`);
  } else {
    // Fall back to parsing the archive snapshot JSON
    source = "archive snapshot JSON";
    deductionEntries = [];
    for (const item of items) {
      const empId = item.employee?.id;
      if (!empId) continue;
      for (const ded of (item.deductions ?? [])) {
        if (!ded.deductiontype_id || !ded.amount || Number(ded.amount) === 0) continue;
        deductionEntries.push({
          employee_id: empId,
          deductiontype_id: ded.deductiontype_id,
          amount: Number(ded.amount),
        });
      }
    }
    console.log(`Source: archive snapshot JSON (${deductionEntries.length} deduction entries)`);
  }

  if (deductionEntries.length === 0) {
    console.log("\nNo deduction amounts found in archive. The archive may not have itemized deductions.");
    console.log("Check that employees have deductions filled in before saving an archive.");
    await pool.end();
    return;
  }

  // ── 3. Upsert into deduction table ──────────────────────────────────────────
  let inserted = 0, updated = 0, skipped = 0;

  for (const { employee_id, deductiontype_id, amount } of deductionEntries) {
    // Verify employee exists
    const empCheck = await q("SELECT id FROM `employee` WHERE id=?", [employee_id]);
    if (!empCheck.length) { skipped++; continue; }

    // Verify deductiontype exists
    const dtCheck = await q("SELECT id FROM `deductiontype` WHERE id=?", [deductiontype_id]);
    if (!dtCheck.length) { skipped++; continue; }

    const existing = await q(
      "SELECT id FROM `deduction` WHERE employee_id=? AND deductiontype_id=?",
      [employee_id, deductiontype_id]
    );

    if (existing.length > 0) {
      await q(
        "UPDATE `deduction` SET amount=?, updated_at=NOW() WHERE employee_id=? AND deductiontype_id=?",
        [amount, employee_id, deductiontype_id]
      );
      process.stdout.write("U");
      updated++;
    } else {
      await q(
        "INSERT INTO `deduction` (id, employee_id, deductiontype_id, amount) VALUES (?,?,?,?)",
        [randomUUID(), employee_id, deductiontype_id, amount]
      );
      process.stdout.write("I");
      inserted++;
    }
  }

  console.log(`\n\n✓ Done (source: ${source})`);
  console.log(`  Inserted: ${inserted}`);
  console.log(`  Updated:  ${updated}`);
  console.log(`  Skipped:  ${skipped} (employee/type not found)`);

  // ── 4. Summary ───────────────────────────────────────────────────────────────
  console.log("\nTop employees by total deductions:");
  const summary = await q(`
    SELECT e.name, e.employee_number, COUNT(d.id) AS ded_count, SUM(d.amount) AS total_ded
    FROM employee e
    JOIN deduction d ON d.employee_id = e.id
    GROUP BY e.id, e.name, e.employee_number
    ORDER BY total_ded DESC
    LIMIT 15
  `);

  for (const r of summary) {
    const num = r.employee_number ? `[${r.employee_number}]` : "";
    console.log(`  ${r.name.padEnd(38)} ${num.padEnd(10)}  ${String(r.ded_count).padStart(2)} types  ₱${Number(r.total_ded).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`);
  }

  const [totalRow] = await q("SELECT COUNT(*) AS cnt, SUM(amount) AS total FROM `deduction`");
  console.log(`\nTotal deduction rows: ${totalRow.cnt}  |  Grand total: ₱${Number(totalRow.total ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`);

  await pool.end();
}

main().catch(e => { console.error(e); process.exit(1); });

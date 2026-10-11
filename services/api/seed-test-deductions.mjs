/**
 * Inserts realistic test deductions for all employees in local MySQL.
 *
 * Run this once to populate the deduction table so MY PAYSLIP OVERVIEW
 * shows real numbers instead of ₱0.00.
 *
 * Usage: node seed-test-deductions.mjs
 *
 * Deduction amounts are computed from each employee's basic_monthly_pay
 * using Philippine government payroll rules (GSIS, PhilHealth, Pag-IBIG).
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

function computeDeductions(basicMonthlyPay) {
  const basic = Number(basicMonthlyPay);
  const round2 = v => Math.round(v * 100) / 100;

  // GSIS Life/Retirement: 9% of basic
  const gsisLifeRet = round2(basic * 0.09);

  // PhilHealth: 5% of basic, split employee/employer, employee pays 2.5%, max ₱5,000/month total
  const philHealthTotal = Math.min(basic * 0.05, 5000);
  const medic = round2(philHealthTotal / 2); // employee share

  // Pag-IBIG: ₱200 standard max for employee
  const pagibigPremium = 200;

  // Union dues: ₱200 standard
  const unionDues = 200;

  // Withholding tax: rough estimate for typical gov salary range
  let withholdingTax = 0;
  if (basic > 20833) {
    // Simplified: use ₱0 for those below ₱250,000/year (~₱20,833/month)
    // Simple flat estimate: ~10% of amount over ₱20,833
    withholdingTax = round2((basic - 20833) * 0.10);
  }

  const deductions = {
    gsis_life_ret:    gsisLifeRet,
    medic:            medic,
    pagibig_premium:  pagibigPremium,
    union_dues:       unionDues,
  };
  if (withholdingTax > 0) deductions.withholding_tax = withholdingTax;

  return deductions;
}

async function main() {
  // Get all active employees
  const employees = await q(
    "SELECT id, name, employee_number, basic_monthly_pay FROM `employee` WHERE status='active' ORDER BY name"
  );

  if (!employees.length) {
    console.log("No active employees found.");
    await pool.end();
    return;
  }

  console.log(`Found ${employees.length} active employee(s)\n`);

  // Get deductiontype_id map: name → id
  const dedTypes = await q("SELECT id, name FROM `deductiontype`");
  const dedTypeMap = {};
  for (const dt of dedTypes) dedTypeMap[dt.name] = dt.id;

  let inserted = 0, updated = 0, skipped = 0;

  for (const emp of employees) {
    const basic = Number(emp.basic_monthly_pay);
    if (!basic) { skipped++; continue; }

    const deductions = computeDeductions(basic);
    process.stdout.write(`  ${emp.name.padEnd(38)} [${(emp.employee_number ?? "").padEnd(8)}]  basic=₱${basic.toLocaleString()}\n`);

    for (const [typeName, amount] of Object.entries(deductions)) {
      const dedTypeId = dedTypeMap[typeName];
      if (!dedTypeId) {
        console.log(`    ⚠ Unknown deduction type: ${typeName}`);
        continue;
      }

      const existing = await q(
        "SELECT id FROM `deduction` WHERE employee_id=? AND deductiontype_id=?",
        [emp.id, dedTypeId]
      );

      if (existing.length > 0) {
        // Only update if current amount is 0
        const [cur] = await q(
          "SELECT amount FROM `deduction` WHERE employee_id=? AND deductiontype_id=?",
          [emp.id, dedTypeId]
        );
        if (Number(cur.amount) === 0) {
          await q(
            "UPDATE `deduction` SET amount=?, updated_at=NOW() WHERE employee_id=? AND deductiontype_id=?",
            [amount, emp.id, dedTypeId]
          );
          process.stdout.write("    U ");
          updated++;
        } else {
          process.stdout.write("    . ");
          skipped++;
        }
      } else {
        await q(
          "INSERT INTO `deduction` (id, employee_id, deductiontype_id, amount) VALUES (?,?,?,?)",
          [randomUUID(), emp.id, dedTypeId, amount]
        );
        process.stdout.write("    I ");
        inserted++;
      }
      process.stdout.write(`${typeName}=₱${amount.toLocaleString()}\n`);
    }
  }

  console.log(`\n✓ Done`);
  console.log(`  Inserted: ${inserted}`);
  console.log(`  Updated:  ${updated}`);
  console.log(`  Skipped:  ${skipped}`);

  // Final check for psahr
  const psahrEmp = employees.find(e => e.name?.toLowerCase().includes("de los santos"));
  if (psahrEmp) {
    const deds = await q(
      `SELECT dt.name, d.amount FROM deduction d JOIN deductiontype dt ON dt.id=d.deductiontype_id WHERE d.employee_id=? AND d.amount>0`,
      [psahrEmp.id]
    );
    const total = deds.reduce((s, r) => s + Number(r.amount), 0);
    console.log(`\nMiguel De Los Santos deductions:`);
    for (const r of deds) console.log(`  ${r.name.padEnd(25)} ₱${Number(r.amount).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`);
    console.log(`  ${"TOTAL".padEnd(25)} ₱${total.toLocaleString("en-PH", { minimumFractionDigits: 2 })}`);
  }

  await pool.end();
}

main().catch(e => { console.error(e); process.exit(1); });

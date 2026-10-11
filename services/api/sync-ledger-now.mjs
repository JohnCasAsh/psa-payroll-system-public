import mysql from 'mysql2/promise';
import { randomUUID } from 'crypto';

const pool = await mysql.createPool({
  host: 'localhost', user: 'root', password: '', database: 'psa_payroll', port: 3306,
  waitForConnections: true,
});

const [employees] = await pool.execute('SELECT id, basic_monthly_pay FROM `employee`');
const [types] = await pool.execute(`
  SELECT cc.id AS cc_id, cc.mode, vc.value, pt.name
  FROM payroll_types pt
  JOIN computation_config cc ON cc.payroll_type_id = pt.id
  JOIN value_config vc ON vc.computation_config_id = cc.id AND vc.is_active = 1
  WHERE pt.is_active = 1 AND pt.name != 'basic_pay' AND cc.mode IN (2, 3)
`);

let inserted = 0, updated = 0;

for (const emp of employees) {
  const basicPay = Number(emp.basic_monthly_pay ?? 0);
  for (const t of types) {
    const amount = t.mode === 2 ? Math.round(basicPay * Number(t.value) * 100) / 100 : Number(t.value);
    const [existing] = await pool.execute(
      'SELECT id FROM employee_deduction_ledger WHERE employee_id=? AND computation_config_id=? AND status="active"',
      [emp.id, t.cc_id]
    );
    if (existing.length > 0) {
      await pool.execute(
        'UPDATE employee_deduction_ledger SET amount=? WHERE employee_id=? AND computation_config_id=? AND status="active"',
        [amount, emp.id, t.cc_id]
      );
      updated++;
    } else {
      await pool.execute(
        'INSERT INTO employee_deduction_ledger (id,employee_id,computation_config_id,amount,effective_from,status,created_by) VALUES (?,?,?,?,CURDATE(),"active","system")',
        [randomUUID(), emp.id, t.cc_id, amount]
      );
      inserted++;
    }
    console.log(`  [${t.mode===2?'RATE':'FIXED'}] ${t.name} → emp ${emp.id.slice(0,8)} = ${amount}`);
  }
}

console.log(`\nDone. Inserted: ${inserted}, Updated: ${updated}`);
await pool.end();

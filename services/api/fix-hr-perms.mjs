import mysql from "mysql2/promise";
const pool = await mysql.createPool({
  host: process.env.DB_HOST || "localhost",
  port: process.env.DB_PORT || 3306,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
});

const [roles] = await pool.execute("SELECT id FROM `role` WHERE name = 'HR' LIMIT 1");
if (!roles[0]) { console.error("HR role not found"); process.exit(1); }
const hrId = roles[0].id;

// Add account.avatar alongside account.change_password
const [p] = await pool.execute("SELECT id FROM `permission` WHERE `key` = 'account.avatar' LIMIT 1");
if (p[0]) {
  await pool.execute("INSERT IGNORE INTO `role_permission` (role_id, permission_id) VALUES (?, ?)", [hrId, p[0].id]);
  console.log("Added account.avatar to HR");
} else {
  console.log("account.avatar permission not found in DB");
}

const [check] = await pool.execute(
  "SELECT p.`key` FROM `permission` p JOIN role_permission rp ON rp.permission_id = p.id WHERE rp.role_id = ? ORDER BY p.sort_order",
  [hrId]
);
console.log("HR permissions now:", check.map(r => r.key).join(", "));
await pool.end();

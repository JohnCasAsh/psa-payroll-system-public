import mysql from "mysql2/promise";
const pool = await mysql.createPool({
  host: process.env.DB_HOST || "localhost",
  port: process.env.DB_PORT || 3306,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
});
const [roles] = await pool.execute("SELECT r.name, GROUP_CONCAT(p.`key` ORDER BY p.sort_order SEPARATOR ', ') AS perms FROM `role` r LEFT JOIN role_permission rp ON rp.role_id = r.id LEFT JOIN `permission` p ON p.id = rp.permission_id GROUP BY r.id, r.name");
console.log("ROLES:");
roles.forEach(r => console.log("  " + r.name + ": " + (r.perms || "(none)")));
const [u] = await pool.execute("SELECT u.username, r.name AS role_name, GROUP_CONCAT(p.`key` ORDER BY p.sort_order SEPARATOR ', ') AS perms FROM users u LEFT JOIN `role` r ON r.id = u.role_id LEFT JOIN role_permission rp ON rp.role_id = r.id LEFT JOIN `permission` p ON p.id = rp.permission_id WHERE u.username = 'employee1' GROUP BY u.id, u.username, r.name");
console.log("\nEMPLOYEE1:");
u.forEach(r => console.log("  role=" + r.role_name + ", perms=" + (r.perms || "(none)")));
await pool.end();

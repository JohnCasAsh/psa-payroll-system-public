// One-time script: creates the admin user in MySQL
// Usage: node seed-admin.mjs <password>
// Example: node --env-file=.env seed-admin.mjs "YourStrongPassword!1"  (pick your own)

import bcrypt from "bcryptjs";
import mysql from "mysql2/promise";
import { randomUUID } from "crypto";

const password = process.argv[2];
if (!password) {
  console.error("Usage: node seed-admin.mjs <password>");
  process.exit(1);
}

const MYSQL_URL   = process.env.MYSQL_URL   ?? "mysql://root:@127.0.0.1:3306/psa_payroll";
const ADMIN_USERNAME = (process.env.ADMIN_USERNAME ?? "admin").toLowerCase();

const pool = mysql.createPool({ uri: MYSQL_URL, connectionLimit: 1 });
const hash = await bcrypt.hash(password, 12);

try {
  await pool.execute(
    "INSERT INTO `users` (id, username, password_hash, status) VALUES (?, ?, ?, 'active')",
    [randomUUID(), ADMIN_USERNAME, hash]
  );
  console.log(`✅ Admin user '${ADMIN_USERNAME}' created. You can now log in.`);
} catch (err) {
  if (err.code === "ER_DUP_ENTRY") {
    await pool.execute(
      "UPDATE `users` SET password_hash = ?, status = 'active' WHERE LOWER(username) = ?",
      [hash, ADMIN_USERNAME]
    );
    console.log(`✅ Admin user '${ADMIN_USERNAME}' password reset and activated.`);
  } else {
    console.error("❌ Error:", err.message);
    process.exit(1);
  }
}

await pool.end();

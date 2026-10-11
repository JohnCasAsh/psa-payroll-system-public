/**
 * db.mysql.js — MySQL / MariaDB application database layer
 * Drop-in replacement for db.js when deploying to MySQL.
 *
 * NOTE (for director): All primary keys currently use UUID (VARCHAR 36) for consistency with
 * the PostgreSQL/Railway production database. If the director decides to simplify the MySQL
 * database for easier reading in phpMyAdmin, all IDs can be changed to INT AUTO_INCREMENT.
 * That change only affects this file (db.mysql.js) and server.mysql.js — PostgreSQL is not affected.
 *
 * Set MYSQL_URL in environment:
 *   mysql://root:@127.0.0.1:3306/psa_payroll
 *
 * NOTE: server.js also needs updating when switching to MySQL:
 *   - Change pool.query($1,$2) → pool.execute(?,?)
 *   - Change UUID_RE check → numeric ID check
 */

import mysql from "mysql2/promise";
import { randomUUID, randomInt } from "crypto";
import { cacheGet, cacheSet, incrementVersion, getVersion } from "./cache.js";

const ADMIN_USERNAME = (process.env.ADMIN_USERNAME ?? "admin123").toLowerCase();

export const pool = mysql.createPool({
  uri: process.env.MYSQL_URL ?? `mysql://root:@${process.env.DB_HOST ?? "127.0.0.1"}:${process.env.DB_PORT ?? "3306"}/${process.env.DB_NAME ?? "psa_payroll"}`,
  waitForConnections: true,
  connectionLimit: 10,
  connectTimeout: 30000,
  timezone: "local",
});
pool.on("connection", (conn) => { conn.query("SET foreign_key_checks=0"); });

async function q(sql, params = []) {
  const [rows] = await pool.execute(sql, params);
  return rows;
}

async function ddl(sql) {
  const conn = await pool.getConnection();
  try {
    await conn.query("SET foreign_key_checks=0");
    await conn.query(sql);
  } finally {
    conn.release();
  }
}

async function addColIfMissing(table, col, def) {
  const [rows] = await pool.execute(
    `SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ? AND COLUMN_NAME = ?`,
    [table, col]
  );
  if (rows.length === 0) await pool.query(`ALTER TABLE \`${table}\` ADD COLUMN \`${col}\` ${def}`);
}

// ── SCHEMA ────────────────────────────────────────────────────────────────────

export async function ensureSchema() {
  // Wait for MySQL to be fully ready (important after XAMPP restarts)
  for (let attempt = 1; attempt <= 10; attempt++) {
    try {
      await pool.query("SELECT 1");
      break;
    } catch {
      if (attempt === 10) throw new Error("MySQL not ready after 30s — is XAMPP running?");
      await new Promise(r => setTimeout(r, 3000));
    }
  }
  await pool.query("SET foreign_key_checks=0");
  // Detect old schema: earningtype table exists but lacks the `id` primary key column.
  // This happens when XAMPP MySQL has tables from a pre-id-column build.
  // Safe to drop and recreate because there is no real payroll data yet.
  const [[{ tblExists }]] = await pool.execute(
    `SELECT COUNT(*) AS tblExists FROM INFORMATION_SCHEMA.TABLES
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'earningtype'`
  );
  if (parseInt(tblExists, 10) > 0) {
    const [[{ hasId }]] = await pool.execute(
      `SELECT COUNT(*) AS hasId FROM INFORMATION_SCHEMA.COLUMNS
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'earningtype' AND COLUMN_NAME = 'id'`
    );
    if (parseInt(hasId, 10) === 0) {
      console.log("Old schema detected — dropping all tables for fresh recreation…");
      await ddl("SET FOREIGN_KEY_CHECKS = 0");
      for (const t of ["earning","deduction","employee","earningtype","deductiontype","employee_position","division","users"]) {
        await ddl(`DROP TABLE IF EXISTS \`${t}\``);
      }
      await ddl("SET FOREIGN_KEY_CHECKS = 1");
    }
  }

  // Drop legacy tables fully replaced by payroll_types + employee_deduction_ledger
  await pool.query("SET foreign_key_checks=0");
  await pool.query("DROP TABLE IF EXISTS `payroll_history_earning`");
  await pool.query("DROP TABLE IF EXISTS `payroll_history_deduction`");
  await pool.query("DROP TABLE IF EXISTS `earning`");
  await pool.query("DROP TABLE IF EXISTS `deduction`");
  await pool.query("DROP TABLE IF EXISTS `earningtype`");
  await pool.query("DROP TABLE IF EXISTS `deductiontype`");
  await pool.query("SET foreign_key_checks=1");

  // ── Migration: payslip → employee-direct schema ─────────────────────────────
  // Step 1: migrate earning.payslip_id → earning.employee_id
  {
    const [[{ hasEarnPsCol }]] = await pool.execute(
      `SELECT COUNT(*) AS hasEarnPsCol FROM INFORMATION_SCHEMA.COLUMNS
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'earning' AND COLUMN_NAME = 'payslip_id'`
    );
    if (parseInt(hasEarnPsCol, 10) > 0) {
      await addColIfMissing("earning", "employee_id", "VARCHAR(36) DEFAULT NULL");
      await pool.query(`UPDATE \`earning\` er JOIN \`payslip\` p ON p.id = er.payslip_id SET er.employee_id = p.employee_id WHERE er.employee_id IS NULL`);
      await ddl("SET FOREIGN_KEY_CHECKS = 0");
      await ddl("ALTER TABLE `earning` DROP FOREIGN KEY `fk_earning_payslip`");
      await ddl("ALTER TABLE `earning` DROP INDEX `uq_earning`");
      await ddl("ALTER TABLE `earning` DROP INDEX `idx_earning_payslip`");
      await ddl("ALTER TABLE `earning` DROP COLUMN `payslip_id`");
      await ddl("ALTER TABLE `earning` MODIFY `employee_id` VARCHAR(36) NOT NULL");
      await ddl("ALTER TABLE `earning` ADD CONSTRAINT `fk_earning_employee` FOREIGN KEY (`employee_id`) REFERENCES `employee` (`id`) ON DELETE CASCADE");
      await ddl("ALTER TABLE `earning` ADD UNIQUE KEY `uq_earning` (`employee_id`, `earningtype_id`)");
      await ddl("ALTER TABLE `earning` ADD KEY `idx_earning_employee` (`employee_id`)");
      await ddl("SET FOREIGN_KEY_CHECKS = 1");
    }
  }
  // Step 2: migrate deduction.payslip_id → deduction.employee_id
  {
    const [[{ hasDedPsCol }]] = await pool.execute(
      `SELECT COUNT(*) AS hasDedPsCol FROM INFORMATION_SCHEMA.COLUMNS
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'deduction' AND COLUMN_NAME = 'payslip_id'`
    );
    if (parseInt(hasDedPsCol, 10) > 0) {
      await addColIfMissing("deduction", "employee_id", "VARCHAR(36) DEFAULT NULL");
      await pool.query(`UPDATE \`deduction\` d JOIN \`payslip\` p ON p.id = d.payslip_id SET d.employee_id = p.employee_id WHERE d.employee_id IS NULL`);
      await ddl("SET FOREIGN_KEY_CHECKS = 0");
      await ddl("ALTER TABLE `deduction` DROP FOREIGN KEY `fk_ded_payslip`");
      await ddl("ALTER TABLE `deduction` DROP INDEX `uq_deduction`");
      await ddl("ALTER TABLE `deduction` DROP INDEX `idx_deduction_payslip`");
      await ddl("ALTER TABLE `deduction` DROP COLUMN `payslip_id`");
      await ddl("ALTER TABLE `deduction` MODIFY `employee_id` VARCHAR(36) NOT NULL");
      await ddl("ALTER TABLE `deduction` ADD CONSTRAINT `fk_ded_employee` FOREIGN KEY (`employee_id`) REFERENCES `employee` (`id`) ON DELETE CASCADE");
      await ddl("ALTER TABLE `deduction` ADD UNIQUE KEY `uq_deduction` (`employee_id`, `deductiontype_id`)");
      await ddl("ALTER TABLE `deduction` ADD KEY `idx_deduction_employee` (`employee_id`)");
      await ddl("SET FOREIGN_KEY_CHECKS = 1");
    }
  }
  // Step 3: copy first_quincena from payslip to employee, then drop payslip
  {
    const [[{ hasPsTable }]] = await pool.execute(
      `SELECT COUNT(*) AS hasPsTable FROM INFORMATION_SCHEMA.TABLES
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'payslip'`
    );
    if (parseInt(hasPsTable, 10) > 0) {
      await ddl("SET FOREIGN_KEY_CHECKS = 0");
      await ddl("DROP TABLE IF EXISTS `payslip`");
      await ddl("SET FOREIGN_KEY_CHECKS = 1");
    }
  }

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`users\` (
      \`id\`            VARCHAR(36)  NOT NULL,
      \`username\`      VARCHAR(100) NOT NULL,
      \`password_hash\` TEXT         NOT NULL,
      \`status\`        ENUM('pending','active','blocked','rejected') NOT NULL DEFAULT 'pending',
      \`avatar_key\`    MEDIUMTEXT   DEFAULT NULL,
      \`created_at\`    DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_username\` (\`username\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── salary_grade (normalized parent, INT ids) ─────────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`salary_grade\` (
      \`id\`         INT UNSIGNED NOT NULL AUTO_INCREMENT,
      \`grade_code\` VARCHAR(10)  NOT NULL,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_grade_code\` (\`grade_code\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── salary_grade_step (normalized child, INT ids) ─────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`salary_grade_step\` (
      \`id\`               INT UNSIGNED     NOT NULL AUTO_INCREMENT,
      \`salary_grade_id\`  INT UNSIGNED     NOT NULL,
      \`step\`             TINYINT UNSIGNED NOT NULL,
      \`amount\`           DECIMAL(12,2)    NOT NULL DEFAULT 0,
      \`effective_date\`   DATE             NOT NULL DEFAULT '2024-01-01',
      \`created_at\`       DATETIME         NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_sgs\` (\`salary_grade_id\`, \`step\`, \`effective_date\`),
      CONSTRAINT \`fk_sgs_grade\` FOREIGN KEY (\`salary_grade_id\`) REFERENCES \`salary_grade\`(\`id\`) ON DELETE RESTRICT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── Migrate: old flat salary_grade VARCHAR column → normalized salary_grade_id ─
  {
    const [[{ hasVarcharSG }]] = await pool.execute(
      "SELECT COUNT(*) AS hasVarcharSG FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='salary_grade_step' AND COLUMN_NAME='salary_grade'"
    );
    if (parseInt(hasVarcharSG, 10) > 0) {
      // Populate salary_grade parent from distinct values in the old column
      const [distinct] = await pool.execute("SELECT DISTINCT salary_grade FROM salary_grade_step WHERE salary_grade IS NOT NULL");
      for (const { salary_grade } of distinct) {
        await pool.execute("INSERT IGNORE INTO `salary_grade` (grade_code) VALUES (?)", [salary_grade]);
      }
      // Add salary_grade_id column
      await ddl("ALTER TABLE `salary_grade_step` ADD COLUMN `salary_grade_id` INT UNSIGNED DEFAULT NULL AFTER `id`");
      // Populate it
      await pool.execute("UPDATE `salary_grade_step` sgs JOIN `salary_grade` sg ON sg.grade_code = sgs.salary_grade SET sgs.salary_grade_id = sg.id");
      // Make NOT NULL
      await ddl("ALTER TABLE `salary_grade_step` MODIFY `salary_grade_id` INT UNSIGNED NOT NULL");
      // Drop old unique key and column
      try { await ddl("ALTER TABLE `salary_grade_step` DROP INDEX `uq_grade_step_effective`"); } catch {}
      await ddl("ALTER TABLE `salary_grade_step` DROP COLUMN `salary_grade`");
      // Add new unique key and FK
      try { await ddl("ALTER TABLE `salary_grade_step` ADD UNIQUE KEY `uq_sgs` (`salary_grade_id`, `step`, `effective_date`)"); } catch {}
      try { await ddl("ALTER TABLE `salary_grade_step` ADD CONSTRAINT `fk_sgs_grade` FOREIGN KEY (`salary_grade_id`) REFERENCES `salary_grade`(`id`) ON DELETE RESTRICT"); } catch {}
    }
  }

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`payroll_types\` (
      \`id\`         INT          NOT NULL AUTO_INCREMENT,
      \`name\`       VARCHAR(100) NOT NULL,
      \`category\`   VARCHAR(50)  DEFAULT NULL,
      \`side\`       ENUM('deduction','less_deduction','gross') NOT NULL DEFAULT 'deduction',
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_name\` (\`name\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);
  await addColIfMissing("payroll_types", "category",      "VARCHAR(50) DEFAULT NULL AFTER `name`");
  await addColIfMissing("payroll_types", "display_name",  "VARCHAR(200) DEFAULT NULL");
  await addColIfMissing("payroll_types", "is_active",     "TINYINT(1) NOT NULL DEFAULT 1");
  await addColIfMissing("payroll_types", "sort_order",    "INT NOT NULL DEFAULT 0");
  await addColIfMissing("payroll_types", "merge_in_excel","TINYINT(1) NOT NULL DEFAULT 0");
  await addColIfMissing("payroll_types", "mode",         "TINYINT NOT NULL DEFAULT 1 COMMENT '1=manual,2=rate,3=fixed'");
  await addColIfMissing("payroll_types", "is_loan",      "TINYINT(1) NOT NULL DEFAULT 0");
  try { await pool.query("ALTER TABLE `payroll_types` DROP COLUMN `created_at`"); } catch {}
  try { await ddl("ALTER TABLE `payroll_types` ADD UNIQUE KEY `uq_side_sort` (`side`, `sort_order`)"); } catch { /* already exists */ }
  await pool.query("UPDATE `payroll_types` SET side = 'deduction' WHERE side IS NULL");

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`value_config\` (
      \`id\`             INT           NOT NULL AUTO_INCREMENT,
      \`payroll_type_id\`INT           NOT NULL,
      \`value_type\`     ENUM('RATE','FIXED') NOT NULL DEFAULT 'RATE',
      \`value\`          DECIMAL(12,4) DEFAULT NULL,
      \`start_date\`     DATE          NOT NULL,
      \`end_date\`       DATE          DEFAULT NULL,
      \`is_active\`      TINYINT(1)    NOT NULL DEFAULT 1,
      PRIMARY KEY (\`id\`),
      CONSTRAINT \`fk_vc_payroll_type\` FOREIGN KEY (\`payroll_type_id\`) REFERENCES \`payroll_types\` (\`id\`) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`employee_deduction_ledger\` (
      \`id\`             BIGINT        NOT NULL,
      \`employee_id\`    INT           NOT NULL,
      \`payroll_type_id\`INT           NOT NULL,
      \`amount\`         DECIMAL(12,2) NOT NULL DEFAULT 0,
      \`created_date\`   DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`status\`         ENUM('active','closed','cancelled','inactive','invalidated','refunded') NOT NULL DEFAULT 'active',
      \`created_by\`     VARCHAR(100)  NOT NULL DEFAULT '',
      PRIMARY KEY (\`id\`),
      CONSTRAINT \`fk_edl_employee\`     FOREIGN KEY (\`employee_id\`)    REFERENCES \`employee\`      (\`id\`) ON DELETE RESTRICT,
      CONSTRAINT \`fk_edl_payroll_type\` FOREIGN KEY (\`payroll_type_id\`) REFERENCES \`payroll_types\` (\`id\`) ON DELETE RESTRICT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);
  // number_of_days removed — architecture change (2026-08-12)
  // Migrate: replace UUID id with short INT id (promote order_no → id as PK)
  {
    const [[idCol]] = await pool.execute(
      "SELECT COLUMN_TYPE FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'id'"
    );
    if (idCol && idCol.COLUMN_TYPE.toLowerCase().includes('char')) {
      await pool.query("ALTER TABLE `employee_deduction_ledger` DROP FOREIGN KEY `fk_edl_employee`");
      await pool.query("ALTER TABLE `employee_deduction_ledger` DROP FOREIGN KEY `fk_edl_comp_config`");
      await pool.query("ALTER TABLE `employee_deduction_ledger` DROP PRIMARY KEY");
      await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `id`");
      await pool.query("ALTER TABLE `employee_deduction_ledger` CHANGE `order_no` `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY FIRST");
      await pool.query("ALTER TABLE `employee_deduction_ledger` ADD CONSTRAINT `fk_edl_employee` FOREIGN KEY (`employee_id`) REFERENCES `employee` (`id`) ON DELETE RESTRICT");
      await pool.query("ALTER TABLE `employee_deduction_ledger` ADD CONSTRAINT `fk_edl_comp_config` FOREIGN KEY (`computation_config_id`) REFERENCES `computation_config` (`id`) ON DELETE RESTRICT");
    }
  }
  // Handle case where order_no exists but id doesn't (UUID id was already dropped, order_no not yet renamed)
  {
    const [[{ hasOrderNo }]] = await pool.execute(
      "SELECT COUNT(*) AS hasOrderNo FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'order_no'"
    );
    const [[{ hasId }]] = await pool.execute(
      "SELECT COUNT(*) AS hasId FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'id'"
    );
    if (hasOrderNo > 0 && hasId === 0) {
      // Drop AUTO_INCREMENT first (required to add PK separately), then rename, then set PK
      await pool.query("ALTER TABLE `employee_deduction_ledger` MODIFY `order_no` INT NOT NULL");
      await pool.query("ALTER TABLE `employee_deduction_ledger` CHANGE `order_no` `id` BIGINT NOT NULL FIRST");
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` ADD PRIMARY KEY (`id`)"); } catch { /* already PK */ }
    }
  }
  // Migrate id from INT AUTO_INCREMENT → BIGINT (random 10-digit IDs)
  {
    const [[idCol]] = await pool.execute(
      "SELECT COLUMN_TYPE FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'id'"
    );
    if (idCol && idCol.COLUMN_TYPE.toLowerCase() === 'int') {
      await pool.query("ALTER TABLE `employee_deduction_ledger` MODIFY COLUMN `id` BIGINT NOT NULL");
    }
  }
  // Drop legacy order_no column — id IS the order id now
  {
    const [[{ hasOrderNo }]] = await pool.execute(
      "SELECT COUNT(*) AS hasOrderNo FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'order_no'"
    );
    if (hasOrderNo > 0) {
      await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `order_no`");
    }
  }
  // number_of_days default fix removed — column dropped (2026-08-12)
  // Fix missing AUTO_INCREMENT on id — assign unique ids to any id=0 rows first, then add AUTO_INCREMENT
  {
    const [[idCol]] = await pool.execute(
      "SELECT EXTRA FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'id'"
    );
    if (idCol && !idCol.EXTRA.toLowerCase().includes('auto_increment')) {
      const [[{ zeroCount }]] = await pool.execute("SELECT COUNT(*) AS zeroCount FROM `employee_deduction_ledger` WHERE `id` = 0");
      if (parseInt(zeroCount) > 0) {
        await pool.query("SET @rn = 0");
        await pool.query("UPDATE `employee_deduction_ledger` SET `id` = (SELECT FLOOR(RAND()*9000000000)+1000000000) WHERE `id` = 0");
      }
      await pool.query("ALTER TABLE `employee_deduction_ledger` MODIFY COLUMN `id` BIGINT NOT NULL AUTO_INCREMENT");
    }
  }
  // Add 'cancelled' (reversed-out), 'inactive' (older close stepped past) and 'return'
  // (money credited back from a dead row) statuses. For pre-existing tables.
  await pool.query("ALTER TABLE `employee_deduction_ledger` MODIFY COLUMN `status` ENUM('active','closed','cancelled','inactive','invalidated','refunded') NOT NULL DEFAULT 'active'");
  // Drop unused tracking columns — replaced by audit_log
  try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `count`"); } catch {}
  try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `difference`"); } catch {}

  // Fix zero/invalid created_date values
  try { await pool.query("UPDATE `employee_deduction_ledger` SET `created_date` = NOW() WHERE `created_date` = '0000-00-00 00:00:00' OR `created_date` IS NULL"); } catch {}

  // Fix rows with id = 0 or low sequential IDs from the old AUTO_INCREMENT era that collide:
  // re-key any id < 1000000000 to a proper 10-digit random ID.
  {
    const [badRows] = await pool.execute(
      "SELECT id FROM `employee_deduction_ledger` WHERE id < 1000000000 ORDER BY id ASC"
    );
    for (const row of badRows) {
      let newId;
      while (true) {
        newId = Math.floor(1_000_000_000 + Math.random() * 9_000_000_000);
        const [[{ cnt }]] = await pool.execute(
          "SELECT COUNT(*) AS cnt FROM `employee_deduction_ledger` WHERE id = ?", [newId]
        );
        if (cnt === 0) break;
      }
      await pool.query("UPDATE `employee_deduction_ledger` SET id = ? WHERE id = ?", [newId, row.id]);
    }
  }

  // number_of_days backfill removed — column dropped (2026-08-12)
  // Schema migrations: rename effective_from → created_date, drop effective_to, drop number_of_days
  {
    const [[{ hasEf }]] = await pool.execute("SELECT COUNT(*) AS hasEf FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee_deduction_ledger' AND COLUMN_NAME='effective_from'");
    if (parseInt(hasEf) > 0) {
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` CHANGE COLUMN `effective_from` `created_date` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP"); } catch (e) { console.warn("rename effective_from:", e.message); }
    }
  }
  {
    const [[{ hasEt }]] = await pool.execute("SELECT COUNT(*) AS hasEt FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee_deduction_ledger' AND COLUMN_NAME='effective_to'");
    if (parseInt(hasEt) > 0) {
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `effective_to`"); } catch {}
    }
  }
  {
    const [[{ hasNod }]] = await pool.execute("SELECT COUNT(*) AS hasNod FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee_deduction_ledger' AND COLUMN_NAME='number_of_days'");
    if (parseInt(hasNod) > 0) {
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `number_of_days`"); } catch {}
    }
  }

  // employee_position table is created after employee — see below

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`division\` (
      \`id\`         INT UNSIGNED NOT NULL AUTO_INCREMENT,
      \`name\`       VARCHAR(100) NOT NULL,
      \`created_at\` DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_division_name\` (\`name\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`employee\` (
      \`id\`                    INT           NOT NULL,
      \`employee_number\`       VARCHAR(20)   NOT NULL,
      \`name\`                  VARCHAR(100)  NOT NULL,
      \`employment_type\`       TINYINT       NOT NULL DEFAULT 1 COMMENT '1=Regular, 2=Contractual',
      \`basic_monthly_pay\`     DECIMAL(12,2) NOT NULL DEFAULT 0,
      \`status\`                VARCHAR(20)   NOT NULL DEFAULT 'active',
      \`position_id\`           INT UNSIGNED  DEFAULT NULL,
      \`division_id\`           INT UNSIGNED  DEFAULT NULL,
      \`start_date\`            DATE          DEFAULT NULL,
      \`granted_step\`          TINYINT       DEFAULT NULL,
      \`step\`                  TINYINT       DEFAULT NULL,
      \`first_quincena_override\` DECIMAL(12,2) DEFAULT NULL,
      \`separation_reason\`     VARCHAR(50)   DEFAULT NULL,
      \`deactivated_at\`        DATETIME      DEFAULT NULL,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_employee_number\` (\`employee_number\`),
      KEY \`idx_name\` (\`name\`),
      KEY \`idx_status\` (\`status\`),
      CONSTRAINT \`fk_emp_division\` FOREIGN KEY (\`division_id\`) REFERENCES \`division\` (\`id\`) ON DELETE SET NULL
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── Position (plantilla) + Employee Position History tables ────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`position\` (
      \`id\`               INT UNSIGNED NOT NULL AUTO_INCREMENT,
      \`position_title\`   VARCHAR(150) NOT NULL,
      \`division_id\`      INT UNSIGNED DEFAULT NULL,
      \`salary_grade_id\`  INT UNSIGNED DEFAULT NULL,
      \`created_at\`       DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_position_title\` (\`position_title\`),
      CONSTRAINT \`fk_pos_division\`      FOREIGN KEY (\`division_id\`)     REFERENCES \`division\`      (\`id\`) ON DELETE SET NULL,
      CONSTRAINT \`fk_pos_salary_grade\`  FOREIGN KEY (\`salary_grade_id\`) REFERENCES \`salary_grade\` (\`id\`) ON DELETE SET NULL
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Migrate: position.default_salary_grade VARCHAR → salary_grade_id INT FK
  {
    const [[{ hasOldCol }]] = await pool.execute(
      "SELECT COUNT(*) AS hasOldCol FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='position' AND COLUMN_NAME='default_salary_grade'"
    );
    if (parseInt(hasOldCol) > 0) {
      await addColIfMissing("position", "salary_grade_id", "INT UNSIGNED DEFAULT NULL");
      await pool.execute(`
        UPDATE \`position\` p
        JOIN \`salary_grade\` sg ON sg.grade_code = p.default_salary_grade
        SET p.salary_grade_id = sg.id
        WHERE p.default_salary_grade IS NOT NULL
      `);
      await ddl("ALTER TABLE `position` DROP COLUMN `default_salary_grade`");
      try {
        await ddl("ALTER TABLE `position` ADD CONSTRAINT `fk_pos_salary_grade` FOREIGN KEY (`salary_grade_id`) REFERENCES `salary_grade`(`id`) ON DELETE SET NULL");
      } catch { /* already exists */ }
    }
  }

  // Migration: old employee_position (UUID lookup) → new position (INT) + new employee_position (history)
  {
    const [[{ hasEpName }]] = await pool.execute(
      "SELECT COUNT(*) AS hasEpName FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee_position' AND COLUMN_NAME='name'"
    );
    if (parseInt(hasEpName, 10) > 0) {
      await pool.execute("INSERT IGNORE INTO `position` (position_title) SELECT name FROM `employee_position`");
      await addColIfMissing("employee", "position_id_new", "INT UNSIGNED DEFAULT NULL");
      await pool.execute(`
        UPDATE employee e
        INNER JOIN employee_position ep ON e.position_id = ep.id
        INNER JOIN \`position\` p ON p.position_title = ep.name
        SET e.position_id_new = p.id
      `);
      try { await ddl("ALTER TABLE `employee` DROP FOREIGN KEY `fk_emp_position`"); } catch {}
      try { await ddl("ALTER TABLE `employee` DROP COLUMN `position_id`"); } catch {}
      await ddl("ALTER TABLE `employee` CHANGE `position_id_new` `position_id` INT UNSIGNED DEFAULT NULL");
      await ddl("DROP TABLE IF EXISTS `employee_position`");
    } else {
      // Fresh install: position_id may still be VARCHAR(36) from old CREATE TABLE
      const [[posIdRow]] = await pool.execute(
        "SELECT COLUMN_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee' AND COLUMN_NAME='position_id'"
      ).catch(() => [[null]]);
      if (posIdRow && posIdRow.COLUMN_TYPE?.toLowerCase().includes('varchar')) {
        try { await ddl("ALTER TABLE `employee` DROP FOREIGN KEY `fk_emp_position`"); } catch {}
        await ddl("ALTER TABLE `employee` MODIFY `position_id` INT UNSIGNED DEFAULT NULL");
      }
    }
  }

  // New employee_position — assignment history (APPOINTMENT / PROMOTION / TRANSFER)
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`employee_position\` (
      \`id\`            INT UNSIGNED NOT NULL AUTO_INCREMENT,
      \`employee_id\`   INT          NOT NULL,
      \`position_id\`   INT UNSIGNED NOT NULL,
      \`start_date\`    DATE         NOT NULL,
      \`end_date\`      DATE         DEFAULT NULL,
      \`status\`        ENUM('active','superseded') NOT NULL DEFAULT 'active',
      \`movement_type\` ENUM('APPOINTMENT','PROMOTION','TRANSFER') NOT NULL DEFAULT 'APPOINTMENT',
      PRIMARY KEY (\`id\`),
      CONSTRAINT \`fk_ep_employee\` FOREIGN KEY (\`employee_id\`) REFERENCES \`employee\`(\`id\`) ON DELETE CASCADE,
      CONSTRAINT \`fk_ep_position\` FOREIGN KEY (\`position_id\`) REFERENCES \`position\`(\`id\`) ON DELETE RESTRICT,
      INDEX \`idx_ep_emp_status\` (\`employee_id\`, \`status\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Non-creditable service — tolls the 3-year step clock
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`employee_suspension\` (
      \`id\`          INT UNSIGNED NOT NULL AUTO_INCREMENT,
      \`employee_id\` INT          NOT NULL,
      \`start_date\`  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`end_date\`    DATETIME     DEFAULT NULL,
      \`reason\`      TINYINT(1)   DEFAULT NULL COMMENT '1=On Leave, 2=Suspended',
      \`status\`      ENUM('served','lifted','appealed') NOT NULL DEFAULT 'served',
      \`created_at\`  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      KEY \`idx_es_employee\` (\`employee_id\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Ensure employee.position_id FK → position table
  {
    const [[{ hasFk }]] = await pool.execute(
      "SELECT COUNT(*) AS hasFk FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee' AND CONSTRAINT_NAME='fk_emp_position'"
    );
    if (parseInt(hasFk, 10) === 0) {
      try { await ddl("ALTER TABLE `employee` ADD CONSTRAINT `fk_emp_position` FOREIGN KEY (`position_id`) REFERENCES `position`(`id`) ON DELETE SET NULL"); } catch {}
    }
  }

  // ── Seed data ──────────────────────────────────────────────────────────────

  const positions = [
    "Director II", "Administrative Assistant I", "Chief Administrative Officer",
    "Registration Officer IV", "Accountant III", "Statistical Specialist II",
    "Administrative Officer IV", "Administrative Officer III", "Administrative Assistant III",
    "Administrative Aide VI", "Chief Statistical Specialist", "Supervising Statistical Specialist",
    "Senior Statistical Specialist", "Information Systems Analyst I", "Information Officer I",
    "Statistical Analyst", "Assistant Statistician",
  ];
  for (const name of positions) {
    const rows = await q("SELECT id FROM `position` WHERE position_title = ?", [name]);
    if (rows.length === 0) await q("INSERT INTO `position` (position_title) VALUES (?)", [name]);
  }

  // Recency-ranking columns for lookup tables (employment_types now has these in CREATE TABLE)

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`payroll_archive\` (
      \`id\`          INT           NOT NULL AUTO_INCREMENT,
      \`year\`        SMALLINT      NOT NULL,
      \`month\`       TINYINT       NOT NULL,
      \`quincena\`    TINYINT       NOT NULL,
      \`snapshot\`    LONGTEXT      NOT NULL,
      \`row_count\`   INT           NOT NULL DEFAULT 0,
      \`saved_by\`    VARCHAR(100)  NOT NULL DEFAULT '',
      \`notes\`       TEXT          DEFAULT NULL,
      \`created_at\`  DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`updated_at\`  DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_archive\` (\`year\`, \`month\`, \`quincena\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── Payroll history ───────────────────────────────────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`payroll_history\` (
      \`id\`                INT            NOT NULL AUTO_INCREMENT,
      \`employee_id\`       INT            NOT NULL,
      \`year\`              SMALLINT       NOT NULL,
      \`month\`             TINYINT        NOT NULL,
      \`quincena\`          TINYINT        NOT NULL,
      \`gross_pay\`         DECIMAL(12,2)  NOT NULL DEFAULT 0,
      \`total_earnings\`    DECIMAL(12,2)  NOT NULL DEFAULT 0,
      \`total_deductions\`  DECIMAL(12,2)  NOT NULL DEFAULT 0,
      \`net_pay\`           DECIMAL(12,2)  NOT NULL DEFAULT 0,
      \`snapshot_json\`     LONGTEXT       NULL,
      \`saved_at\`          DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_payroll_history\` (\`employee_id\`, \`year\`, \`month\`, \`quincena\`),
      CONSTRAINT \`fk_ph_employee\` FOREIGN KEY (\`employee_id\`) REFERENCES \`employee\` (\`id\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── RBAC ──────────────────────────────────────────────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`role\` (
      \`id\`          INT          NOT NULL AUTO_INCREMENT,
      \`name\`        VARCHAR(100) NOT NULL,
      \`description\` TEXT         DEFAULT NULL,
      \`created_at\`  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_role_name\` (\`name\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`permission\` (
      \`id\`         INT          NOT NULL AUTO_INCREMENT,
      \`key\`        VARCHAR(100) NOT NULL,
      \`label\`      VARCHAR(200) NOT NULL,
      \`category\`   VARCHAR(100) NOT NULL,
      \`sort_order\` INT          NOT NULL DEFAULT 0,
      PRIMARY KEY (\`id\`),
      UNIQUE KEY \`uq_permission_key\` (\`key\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Deduplicate permission table — keep only MIN(id) per key, remove orphan role_permission rows
  {
    const [[{ hasDup }]] = await pool.execute(
      "SELECT COUNT(*) - COUNT(DISTINCT `key`) AS hasDup FROM `permission`"
    );
    if (parseInt(hasDup, 10) > 0) {
      await pool.execute(`
        DELETE rp FROM \`role_permission\` rp
        WHERE rp.permission_id NOT IN (
          SELECT id FROM (SELECT MIN(id) AS id FROM \`permission\` GROUP BY \`key\`) k
        )
      `);
      await pool.execute(`
        DELETE p FROM \`permission\` p
        WHERE p.id NOT IN (
          SELECT id FROM (SELECT MIN(id) AS id FROM \`permission\` GROUP BY \`key\`) k
        )
      `);
    }
  }
  // Ensure unique key exists (may be missing if table was created before this constraint was added)
  try { await ddl("ALTER TABLE `permission` ADD UNIQUE KEY `uq_permission_key` (`key`)"); } catch {}

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`role_permission\` (
      \`role_id\`       INT         NOT NULL,
      \`permission_id\` INT NOT NULL,
      PRIMARY KEY (\`role_id\`, \`permission_id\`),
      CONSTRAINT \`fk_rp_role\`       FOREIGN KEY (\`role_id\`)       REFERENCES \`role\`(\`id\`)       ON DELETE CASCADE,
      CONSTRAINT \`fk_rp_permission\` FOREIGN KEY (\`permission_id\`) REFERENCES \`permission\`(\`id\`) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Add role_id to users (idempotent)
  await addColIfMissing("users", "role_id",        "INT DEFAULT NULL");
  await addColIfMissing("users", "employee_id",    "INT DEFAULT NULL");
  await addColIfMissing("users", "token_version",  "INT NOT NULL DEFAULT 0");
  await addColIfMissing("users", "failed_attempts","TINYINT NOT NULL DEFAULT 0");
  await addColIfMissing("users", "locked_until",   "DATETIME DEFAULT NULL");

  // FK users.employee_id → employee.id
  {
    const [[{ hasFk }]] = await pool.execute(
      "SELECT COUNT(*) AS hasFk FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='users' AND CONSTRAINT_NAME='fk_users_employee'"
    );
    if (parseInt(hasFk, 10) === 0) {
      try {
        await ddl("ALTER TABLE `users` ADD CONSTRAINT `fk_users_employee` FOREIGN KEY (`employee_id`) REFERENCES `employee`(`id`) ON DELETE SET NULL");
      } catch { /* already exists — ignore */ }
    }
  }

  // ── Migrate role.id UUID → INT AUTO_INCREMENT ─────────────────────────────
  // Recovery: if _new_id still exists (last step failed), finish the rename
  {
    const [[{ hasNewId }]] = await pool.execute(
      "SELECT COUNT(*) AS hasNewId FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='role' AND COLUMN_NAME='_new_id'"
    );
    if (parseInt(hasNewId) > 0) {
      // Can't rename directly because _new_id has AUTO_INCREMENT+UNIQUE — add id, copy, drop _new_id, then make it PK
      const [[{ hasIdCol }]] = await pool.execute(
        "SELECT COUNT(*) AS hasIdCol FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='role' AND COLUMN_NAME='id'"
      );
      if (!parseInt(hasIdCol)) {
        await pool.query("ALTER TABLE `role` ADD COLUMN `id` INT DEFAULT NULL");
      }
      await pool.query("UPDATE `role` SET `id` = `_new_id`");
      await pool.query("ALTER TABLE `role` DROP COLUMN `_new_id`"); // auto-drops UNIQUE KEY _new_id
      await pool.query("ALTER TABLE `role` MODIFY COLUMN `id` INT NOT NULL AUTO_INCREMENT, ADD PRIMARY KEY (`id`)");
      // Restore FKs now that role.id is INT
      try { await pool.query("ALTER TABLE `users` ADD CONSTRAINT `fk_users_role` FOREIGN KEY (`role_id`) REFERENCES `role`(`id`) ON DELETE SET NULL"); } catch {}
      try { await pool.query("ALTER TABLE `role_permission` ADD CONSTRAINT `fk_rp_role` FOREIGN KEY (`role_id`) REFERENCES `role`(`id`) ON DELETE CASCADE"); } catch {}
    }
  }
  {
    const [[{ isVarchar }]] = await pool.execute(
      "SELECT COUNT(*) AS isVarchar FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='role' AND COLUMN_NAME='id' AND DATA_TYPE='varchar'"
    );
    if (parseInt(isVarchar) > 0) {
      // 1. Add temp INT id to role
      await pool.query("ALTER TABLE `role` ADD COLUMN `_new_id` INT NOT NULL AUTO_INCREMENT UNIQUE");
      // 2. Add temp role_id columns to child tables
      await pool.query("ALTER TABLE `users` ADD COLUMN `_new_role_id` INT DEFAULT NULL");
      await pool.query("ALTER TABLE `role_permission` ADD COLUMN `_new_role_id` INT DEFAULT NULL");
      // 3. Populate temp columns via the mapping
      await pool.query("UPDATE `users` u JOIN `role` r ON r.id = u.role_id SET u._new_role_id = r._new_id");
      await pool.query("UPDATE `role_permission` rp JOIN `role` r ON r.id = rp.role_id SET rp._new_role_id = r._new_id");
      // 4. Drop FKs and composite PK that reference old UUID
      try { await pool.query("ALTER TABLE `users` DROP FOREIGN KEY `fk_users_role`"); } catch {}
      try { await pool.query("ALTER TABLE `role_permission` DROP FOREIGN KEY `fk_rp_role`"); } catch {}
      try { await pool.query("ALTER TABLE `role_permission` DROP PRIMARY KEY"); } catch {}
      // 5. Drop old UUID columns from child tables
      await pool.query("ALTER TABLE `users` DROP COLUMN `role_id`");
      await pool.query("ALTER TABLE `role_permission` DROP COLUMN `role_id`");
      // 6. Rename temp columns to role_id
      await pool.query("ALTER TABLE `users` CHANGE COLUMN `_new_role_id` `role_id` INT DEFAULT NULL");
      await pool.query("ALTER TABLE `role_permission` CHANGE COLUMN `_new_role_id` `role_id` INT NOT NULL");
      // 7. Swap role.id — drop old VARCHAR PK, rename _new_id to id
      await pool.query("ALTER TABLE `role` DROP PRIMARY KEY, DROP COLUMN `id`");
      await pool.query("ALTER TABLE `role` CHANGE COLUMN `_new_id` `id` INT NOT NULL AUTO_INCREMENT, ADD PRIMARY KEY (`id`)");
      // 8. Restore composite PK and FKs
      await pool.query("ALTER TABLE `role_permission` ADD PRIMARY KEY (`role_id`, `permission_id`)");
      try { await pool.query("ALTER TABLE `users` ADD CONSTRAINT `fk_users_role` FOREIGN KEY (`role_id`) REFERENCES `role`(`id`) ON DELETE SET NULL"); } catch {}
      try { await pool.query("ALTER TABLE `role_permission` ADD CONSTRAINT `fk_rp_role` FOREIGN KEY (`role_id`) REFERENCES `role`(`id`) ON DELETE CASCADE"); } catch {}
    }
  }

  // Add FK from users.role_id → role.id if not already present
  {
    const [[{ hasFk }]] = await pool.execute(
      "SELECT COUNT(*) AS hasFk FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='users' AND CONSTRAINT_NAME='fk_users_role'"
    );
    if (parseInt(hasFk, 10) === 0) {
      try {
        await ddl("ALTER TABLE `users` ADD CONSTRAINT `fk_users_role` FOREIGN KEY (`role_id`) REFERENCES `role`(`id`) ON DELETE SET NULL");
      } catch { /* already exists or table not ready — ignore */ }
    }
  }

  // Seed permission catalog (idempotent)
  const PERMISSION_CATALOG = [
    { key: "account.change_password",       label: "Change Password",                         category: "Account",                    sort_order: 1 },
    { key: "account.avatar",               label: "Change Profile Picture",                  category: "Account",                    sort_order: 2 },
    { key: "dashboard.view",                label: "View Dashboard",                          category: "Dashboard",                  sort_order: 10 },
    { key: "dashboard.financial_report",    label: "View Financial Report (Deductions & Burden)", category: "Dashboard",               sort_order: 11 },
    { key: "files201.view",                 label: "View 201 Files (Self-Service)",            category: "Main",                       sort_order: 12 },
    { key: "leave.file",                    label: "File Own Leave Application",               category: "Leave",                      sort_order: 13 },
    { key: "leave.hr",                      label: "Certify Leave Credits (HR — 7.A)",         category: "Leave",                      sort_order: 14 },
    { key: "leave.crasd",                   label: "Recommend Leave (CRASD — 7.B)",            category: "Leave",                      sort_order: 15 },
    { key: "leave.director",                label: "Approve / Disapprove Leave (Director)",    category: "Leave",                      sort_order: 16 },
    { key: "leave.manage",                  label: "Manage All Leave Applications",            category: "Leave",                      sort_order: 17 },
    { key: "employees.view",                label: "View Employees",                          category: "Employees",                  sort_order: 20 },
    { key: "employees.create",              label: "Add Employee",                            category: "Employees",                  sort_order: 21 },
    { key: "employees.edit",                label: "Edit Employee",                           category: "Employees",                  sort_order: 22 },
    { key: "employees.delete",              label: "Deactivate Employee",                     category: "Employees",                  sort_order: 23 },
    { key: "employees.restore",             label: "Restore Inactive Employee",               category: "Settings - Employees",       sort_order: 65 },
    { key: "deductions.view",               label: "View Employee Deductions",                category: "Employee Deductions",        sort_order: 30 },
    { key: "deductions.edit",               label: "Edit Deductions",                         category: "Employee Deductions",        sort_order: 31 },
    { key: "deductions.bulk_edit",          label: "Bulk Edit Deductions",                    category: "Employee Deductions",        sort_order: 32 },
    { key: "deductions.manage",             label: "Manage Loans (Create/Close/Consolidate)", category: "Employee Deductions",        sort_order: 33 },
    { key: "payslips.view",                 label: "View Pay Slips",                          category: "Pay Slips",                  sort_order: 40 },
    { key: "payslips.export_pdf",           label: "Download Pay Slips (DOCX/PDF)",            category: "Pay Slips",                  sort_order: 41 },
    { key: "payslips.export_excel",         label: "Download Payroll Register (Excel)",        category: "Pay Slips",                  sort_order: 42 },
    { key: "settings.reference.view",       label: "View Reference Data",                     category: "Settings - Reference Data",  sort_order: 50 },
    { key: "settings.reference.manage",     label: "Manage Reference Data (Add/Edit/Delete)", category: "Settings - Reference Data",  sort_order: 51 },
    { key: "settings.payroll_types.view",   label: "View Payroll Types",                      category: "Settings - Payroll Types",   sort_order: 60 },
    { key: "settings.payroll_types.manage", label: "Manage Payroll Types (Add/Edit/Delete)",  category: "Settings - Payroll Types",   sort_order: 61 },
    { key: "archive.view",                  label: "View Payroll Archive",                    category: "Archive",                    sort_order: 45 },
    { key: "archive.manage",               label: "Save / Delete Payroll Archive",           category: "Archive",                    sort_order: 46 },
    { key: "admin.users.view",              label: "View User Accounts",                      category: "User Management",            sort_order: 70 },
    { key: "admin.users.approve",           label: "Approve / Reject Users",                  category: "User Management",            sort_order: 71 },
    { key: "admin.users.block",             label: "Block / Unblock Users",                   category: "User Management",            sort_order: 72 },
    { key: "admin.users.delete",            label: "Delete Users",                            category: "User Management",            sort_order: 73 },
  ];
  for (const p of PERMISSION_CATALOG) {
    await pool.execute(
      "INSERT INTO `permission` (`key`,`label`,`category`,`sort_order`) VALUES (?,?,?,?) ON DUPLICATE KEY UPDATE `label`=VALUES(`label`),`category`=VALUES(`category`),`sort_order`=VALUES(`sort_order`)",
      [p.key, p.label, p.category, p.sort_order]
    );
  }

  await ddl(`
    CREATE TABLE IF NOT EXISTS \`app_settings\` (
      \`key\`   VARCHAR(100) NOT NULL,
      \`value\` LONGTEXT     NOT NULL DEFAULT ('{}'),
      PRIMARY KEY (\`key\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // payroll_history tables are kept — used by the Employee History feature in Archive

  // ── Migrate division.id from UUID VARCHAR(36) → INT UNSIGNED AUTO_INCREMENT ──
  {
    const [[{ idType }]] = await pool.execute(
      `SELECT DATA_TYPE AS idType FROM INFORMATION_SCHEMA.COLUMNS
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'division' AND COLUMN_NAME = 'id'`
    );
    if (idType === 'varchar') {
      await ddl("SET FOREIGN_KEY_CHECKS = 0");

      // Build new division table with INT id
      await ddl(`CREATE TABLE IF NOT EXISTS \`division_new\` (
        id INT UNSIGNED NOT NULL AUTO_INCREMENT,
        name VARCHAR(100) NOT NULL,
        created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (id),
        UNIQUE KEY uq_division_name (name)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci`);

      // Copy existing rows (ordered so IDs are stable)
      await pool.execute(
        "INSERT IGNORE INTO `division_new` (name, created_at) SELECT name, created_at FROM `division` ORDER BY created_at ASC"
      );

      // Add temp mapping columns to employee and position
      await ddl("ALTER TABLE `employee` ADD COLUMN IF NOT EXISTS `__div_new` INT UNSIGNED DEFAULT NULL");
      await ddl("ALTER TABLE `position` ADD COLUMN IF NOT EXISTS `__div_new` INT UNSIGNED DEFAULT NULL");

      // Fill mapping via name join
      await pool.execute(
        "UPDATE `employee` e JOIN `division` d ON d.id = e.division_id JOIN `division_new` dn ON dn.name = d.name SET e.__div_new = dn.id"
      );
      await pool.execute(
        "UPDATE `position` p JOIN `division` d ON d.id = p.division_id JOIN `division_new` dn ON dn.name = d.name SET p.__div_new = dn.id"
      );

      // Swap columns on employee
      try { await ddl("ALTER TABLE `employee` DROP FOREIGN KEY `fk_emp_division`"); } catch {}
      await ddl("ALTER TABLE `employee` DROP COLUMN `division_id`");
      await ddl("ALTER TABLE `employee` CHANGE `__div_new` `division_id` INT UNSIGNED DEFAULT NULL");

      // Swap columns on position
      try { await ddl("ALTER TABLE `position` DROP FOREIGN KEY `fk_pos_division`"); } catch {}
      await ddl("ALTER TABLE `position` DROP COLUMN `division_id`");
      await ddl("ALTER TABLE `position` CHANGE `__div_new` `division_id` INT UNSIGNED DEFAULT NULL");

      // Replace old division table
      await ddl("DROP TABLE `division`");
      await ddl("RENAME TABLE `division_new` TO `division`");

      await ddl("SET FOREIGN_KEY_CHECKS = 1");

      // Restore FKs
      try {
        await ddl("ALTER TABLE `employee` ADD CONSTRAINT `fk_emp_division` FOREIGN KEY (`division_id`) REFERENCES `division`(`id`) ON DELETE SET NULL");
      } catch {}
      try {
        await ddl("ALTER TABLE `position` ADD CONSTRAINT `fk_pos_division` FOREIGN KEY (`division_id`) REFERENCES `division`(`id`) ON DELETE SET NULL");
      } catch {}
    }
  }

  // ── Ensure employee.division_id and position.division_id are INT UNSIGNED ────
  // (Needed if migration ran but left column as wrong type due to partial run)
  {
    const [[{ colType }]] = await pool.execute(
      `SELECT DATA_TYPE AS colType FROM INFORMATION_SCHEMA.COLUMNS
       WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee' AND COLUMN_NAME = 'division_id'`
    );
    if (colType === 'varchar') {
      // employee still has varchar division_id but division table is already INT — just NULL it out
      await ddl("SET FOREIGN_KEY_CHECKS = 0");
      try { await ddl("ALTER TABLE `employee` DROP FOREIGN KEY `fk_emp_division`"); } catch {}
      try { await ddl("ALTER TABLE `position` DROP FOREIGN KEY `fk_pos_division`"); } catch {}
      await ddl("ALTER TABLE `employee` MODIFY `division_id` INT UNSIGNED DEFAULT NULL");
      await ddl("ALTER TABLE `position` MODIFY `division_id` INT UNSIGNED DEFAULT NULL");
      await ddl("SET FOREIGN_KEY_CHECKS = 1");
      try { await ddl("ALTER TABLE `employee` ADD CONSTRAINT `fk_emp_division` FOREIGN KEY (`division_id`) REFERENCES `division`(`id`) ON DELETE SET NULL"); } catch {}
      try { await ddl("ALTER TABLE `position` ADD CONSTRAINT `fk_pos_division` FOREIGN KEY (`division_id`) REFERENCES `division`(`id`) ON DELETE SET NULL"); } catch {}
    }
  }

  // Backfill employee_position for existing employees that have position_id but no history row
  {
    const orphans = await q(`
      SELECT e.id AS employee_id, e.position_id
      FROM \`employee\` e
      WHERE e.status = 'active'
        AND e.position_id IS NOT NULL
        AND NOT EXISTS (
          SELECT 1 FROM \`employee_position\` ep WHERE ep.employee_id = e.id
        )
    `);
    if (orphans.length > 0) {
      const today = new Date().toISOString().slice(0, 10);
      for (const { employee_id, position_id } of orphans) {
        await pool.execute(
          "INSERT IGNORE INTO `employee_position` (employee_id, position_id, start_date, movement_type, status) VALUES (?,?,?,'APPOINTMENT','active')",
          [employee_id, position_id, today]
        );
      }
    }
  }


  // Drop basic_salary table — replaced by generated columns on employee table
  try { await ddl("DROP TABLE IF EXISTS `basic_salary`"); } catch {}

  // Drop employee_position_step table — feature removed entirely
  try { await ddl("DROP TABLE IF EXISTS `employee_position_step`"); } catch {}

  // employee_promotion table removed entirely — step/position changes now write
  // directly to the employee row, no separate promotion timeline is tracked
  try { await ddl("DROP TABLE IF EXISTS `employee_promotion`"); } catch {}


  // Drop salary_grade_id/salary_rank if still present (now derived via JOIN from position table)
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `salary_grade_id`"); } catch {}
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `salary_rank`"); } catch {}
  // Quincena columns removed — now computed by employee_payslip VIEW
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `first_quincena`"); } catch {}
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `second_quincena`"); } catch {}
  // Manual quincena override — null = use net_pay/2 default
  await addColIfMissing("employee", "first_quincena_override", "DECIMAL(12,2) DEFAULT NULL");
  await addColIfMissing("employee", "separation_reason", "VARCHAR(50) DEFAULT NULL");
  await addColIfMissing("employee", "deactivated_at", "DATETIME DEFAULT NULL");
  await addColIfMissing("employee", "start_date", "DATE DEFAULT NULL");
  await addColIfMissing("employee", "granted_step", "TINYINT DEFAULT NULL");
  // Sole anchor for step-increment math — 3-year cycles computed live, no history table
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `last_accessed_at`"); } catch {}
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `created_at`"); } catch {}
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `updated_at`"); } catch {}
  try { await ddl("ALTER TABLE `employee_position` DROP COLUMN `created_at`"); } catch {}
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `payslip_visible_to_user`"); } catch {}
  await addColIfMissing("position", "last_accessed_at", "DATETIME DEFAULT NULL");
  // Recency ranking for the Salary Rank list — same rule as position.last_accessed_at.
  await addColIfMissing("salary_grade", "last_accessed_at", "DATETIME DEFAULT NULL");
  try { await pool.query("ALTER TABLE `salary_grade` DROP COLUMN `created_at`"); } catch {}
  try { await pool.query("DROP VIEW IF EXISTS `employee_payslip`"); } catch {}

  // Remove duplicate grade_code rows (keep lowest id per grade_code) then enforce UNIQUE KEY
  await pool.query(`
    DELETE sg FROM salary_grade sg
    JOIN (SELECT grade_code, MIN(id) AS keep_id FROM salary_grade GROUP BY grade_code) AS first
      ON first.grade_code = sg.grade_code AND sg.id != first.keep_id
    WHERE NOT EXISTS (SELECT 1 FROM salary_grade_step sgs WHERE sgs.salary_grade_id = sg.id)
  `).catch(() => {});
  try { await ddl("ALTER TABLE `salary_grade` ADD UNIQUE KEY `uq_grade_code` (`grade_code`)"); } catch {}

  // effective_from / effective_to modify removed — columns renamed/dropped (2026-08-12)
  // Drop the unused employee_position_id FK and column from employee_suspension
  try { await ddl("ALTER TABLE `employee_suspension` DROP FOREIGN KEY `fk_sus_emp_pos`"); } catch {}
  try { await ddl("ALTER TABLE `employee_suspension` DROP COLUMN `employee_position_id`"); } catch {}
  // Ensure employee_id column exists and is correct type
  await addColIfMissing("employee_suspension", "employee_id", "INT DEFAULT NULL AFTER `id`");
  try { await ddl("ALTER TABLE `employee_suspension` MODIFY COLUMN `employee_id` INT DEFAULT NULL"); } catch {}
  // Upgrade start_date to DATETIME so exact time is recorded
  try { await ddl("ALTER TABLE `employee_suspension` MODIFY COLUMN `start_date` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP"); } catch {}
  // end_date stores full datetime (hours matter for half-day leaves)
  try { await ddl("ALTER TABLE `employee_suspension` MODIFY COLUMN `end_date` DATETIME DEFAULT NULL"); } catch {}
  // ended_at is redundant — end_date now serves as both planned and actual close time
  try { await ddl("ALTER TABLE `employee_suspension` DROP COLUMN `ended_at`"); } catch {}
  // Convert reason from text to TINYINT: 1=On Leave, 2=Suspended
  try {
    await ddl("ALTER TABLE `employee_suspension` MODIFY COLUMN `reason` VARCHAR(255) DEFAULT NULL");
    await pool.execute("UPDATE `employee_suspension` SET reason = '1' WHERE reason = 'On Leave'");
    await pool.execute("UPDATE `employee_suspension` SET reason = '2' WHERE reason = 'Suspended'");
    await ddl("ALTER TABLE `employee_suspension` MODIFY COLUMN `reason` TINYINT(1) DEFAULT NULL COMMENT '1=On Leave, 2=Suspended'");
  } catch {}

  // ── REFACTOR: computation_config → 3-row enum; value_config & ledger → payroll_type_id ──
  {
    const [[{ hasLedgerCcId }]] = await pool.execute(
      "SELECT COUNT(*) AS hasLedgerCcId FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_deduction_ledger' AND COLUMN_NAME = 'computation_config_id'"
    );
    if (parseInt(hasLedgerCcId) > 0) {
      await addColIfMissing("payroll_types", "mode", "TINYINT NOT NULL DEFAULT 1 COMMENT '1=manual,2=rate,3=fixed'");
      await addColIfMissing("value_config", "payroll_type_id", "VARCHAR(36) DEFAULT NULL");
      await addColIfMissing("employee_deduction_ledger", "payroll_type_id", "VARCHAR(36) DEFAULT NULL");

      // Best-effort migrate: computation_config.payroll_type_id stored the INT-cast of pt.id UUID
      try {
        await pool.query(`UPDATE \`payroll_types\` pt JOIN \`computation_config\` cc ON cc.payroll_type_id = CAST(SUBSTRING_INDEX(pt.id,'-',1) AS UNSIGNED) SET pt.mode = cc.mode`);
      } catch {}
      try {
        await pool.query(`UPDATE \`value_config\` vc JOIN \`computation_config\` cc ON cc.id = vc.computation_config_id JOIN \`payroll_types\` pt ON cc.payroll_type_id = CAST(SUBSTRING_INDEX(pt.id,'-',1) AS UNSIGNED) SET vc.payroll_type_id = pt.id WHERE vc.payroll_type_id IS NULL`);
      } catch {}
      try {
        await pool.query(`UPDATE \`employee_deduction_ledger\` edl JOIN \`computation_config\` cc ON cc.id = edl.computation_config_id JOIN \`payroll_types\` pt ON cc.payroll_type_id = CAST(SUBSTRING_INDEX(pt.id,'-',1) AS UNSIGNED) SET edl.payroll_type_id = pt.id WHERE edl.payroll_type_id IS NULL`);
      } catch {}

      // Drop old FKs and computation_config_id columns
      try { await pool.query("ALTER TABLE `value_config` DROP FOREIGN KEY `fk_vc_computation`"); } catch {}
      try { await pool.query("ALTER TABLE `value_config` DROP COLUMN `computation_config_id`"); } catch {}
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP FOREIGN KEY `fk_edl_comp_config`"); } catch {}
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `computation_config_id`"); } catch {}

      // Add FK for new payroll_type_id columns
      try { await pool.query("ALTER TABLE `value_config` ADD CONSTRAINT `fk_vc_payroll_type` FOREIGN KEY (`payroll_type_id`) REFERENCES `payroll_types`(`id`) ON DELETE CASCADE"); } catch {}
      try { await pool.query("ALTER TABLE `employee_deduction_ledger` ADD CONSTRAINT `fk_edl_payroll_type` FOREIGN KEY (`payroll_type_id`) REFERENCES `payroll_types`(`id`) ON DELETE RESTRICT"); } catch {}

      // Restructure computation_config → 3-row enum (delete all old per-type rows)
      try { await pool.query("ALTER TABLE `computation_config` DROP FOREIGN KEY `fk_cc_payroll_type`"); } catch {}
      await pool.query("DELETE FROM `computation_config`");
      try { await pool.query("ALTER TABLE `computation_config` DROP COLUMN `payroll_type_id`"); } catch {}
      try { await pool.query("ALTER TABLE `computation_config` DROP COLUMN `created_at`"); } catch {}
      try { await pool.query("ALTER TABLE `computation_config` ADD COLUMN `label` VARCHAR(20) NOT NULL DEFAULT 'Manual'"); } catch {}
      await pool.query("INSERT IGNORE INTO `computation_config` (id, mode, label) VALUES (1,1,'Manual'),(2,2,'Rate'),(3,3,'Fixed')");
    }
  }

  // Ensure value_config.payroll_type_id migration ran independently of ledger migration
  {
    const [[{ hasVcCcId }]] = await pool.execute(
      "SELECT COUNT(*) AS hasVcCcId FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'value_config' AND COLUMN_NAME = 'computation_config_id'"
    );
    if (parseInt(hasVcCcId) > 0) {
      await addColIfMissing("value_config", "payroll_type_id", "VARCHAR(36) DEFAULT NULL");
      try { await pool.query("ALTER TABLE `value_config` DROP FOREIGN KEY `fk_vc_computation`"); } catch {}
      try { await pool.query("ALTER TABLE `value_config` DROP COLUMN `computation_config_id`"); } catch {}
      try { await pool.query("ALTER TABLE `value_config` ADD CONSTRAINT `fk_vc_payroll_type` FOREIGN KEY (`payroll_type_id`) REFERENCES `payroll_types`(`id`) ON DELETE CASCADE"); } catch {}
    }
  }


  // Drop redundant text columns from employee — now derived via JOIN to reference tables
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `position`"); } catch {}
  try { await ddl("ALTER TABLE `employee` DROP COLUMN `office_division`"); } catch {}

  // Migrate employment_type: FK → VARCHAR → TINYINT (1=Regular, 2=Contractual)
  await addColIfMissing("employee", "employment_type", "TINYINT NOT NULL DEFAULT 1 COMMENT '1=Regular, 2=Contractual'");
  {
    const [[{ hasEmpTypeId }]] = await pool.execute(
      "SELECT COUNT(*) AS hasEmpTypeId FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee' AND COLUMN_NAME = 'employment_type_id'"
    );
    if (parseInt(hasEmpTypeId) > 0) {
      // coming from old FK-based schema — migrate name directly to integer
      try { await pool.query("UPDATE `employee` e JOIN `employment_types` et ON et.id = e.employment_type_id SET e.employment_type = CASE WHEN et.name = 'Contractual' THEN 2 ELSE 1 END WHERE e.employment_type_id IS NOT NULL"); } catch {}
      try { await pool.query("ALTER TABLE `employee` DROP FOREIGN KEY `fk_emp_type`"); } catch {}
      try { await pool.query("ALTER TABLE `employee` DROP COLUMN `employment_type_id`"); } catch {}
    }
  }
  // Convert VARCHAR → TINYINT for databases on the intermediate VARCHAR state
  {
    const [[col]] = await pool.execute(
      "SELECT COLUMN_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee' AND COLUMN_NAME = 'employment_type'"
    );
    if (col && col.COLUMN_TYPE.toLowerCase().includes('varchar')) {
      await addColIfMissing("employee", "employment_type_int", "TINYINT NOT NULL DEFAULT 1");
      await pool.query("UPDATE `employee` SET employment_type_int = CASE WHEN employment_type = 'Contractual' THEN 2 ELSE 1 END");
      try { await pool.query("ALTER TABLE `employee` DROP COLUMN `employment_type`"); } catch {}
      try { await pool.query("ALTER TABLE `employee` CHANGE `employment_type_int` `employment_type` TINYINT NOT NULL DEFAULT 1 COMMENT '1=Regular, 2=Contractual'"); } catch {}
    }
  }
  try { await pool.query("DROP TABLE IF EXISTS `employment_types`"); } catch {}

  // Seed PERA into payroll_types (side='gross') + value_config if not present
  await pool.query(`
    INSERT IGNORE INTO \`payroll_types\` (name, display_name, side, sort_order, mode, is_active)
    VALUES ('pera', 'PERA', 'gross', 2, 3, 1)
  `).catch(() => {});
  await pool.query(`
    INSERT IGNORE INTO \`value_config\` (payroll_type_id, value_type, value, start_date, is_active)
    SELECT pt.id, 'FIXED', 2000, CURDATE(), 1
    FROM \`payroll_types\` pt
    WHERE pt.name = 'pera'
      AND NOT EXISTS (SELECT 1 FROM \`value_config\` vc WHERE vc.payroll_type_id = pt.id AND vc.is_active = 1)
  `).catch(() => {});

  // Auto-seed default deductions for active employees missing ledger entries
  const empNoDefaults = await q(`
    SELECT e.id, e.basic_monthly_pay FROM \`employee\` e
    WHERE e.status = 'active'
      AND NOT EXISTS (SELECT 1 FROM \`employee_deduction_ledger\` WHERE employee_id = e.id)
  `);
  for (const e of empNoDefaults) {
    await _seedEmployeeDefaults(e.id, Number(e.basic_monthly_pay));
  }


  // ── Employee Loans ─────────────────────────────────────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`employee_loans\` (
      \`id\`               INT          NOT NULL AUTO_INCREMENT,
      \`employee_id\`      INT          NOT NULL,
      \`payroll_type_id\`  INT          NOT NULL,
      \`total_amount\`     DECIMAL(12,2) NOT NULL DEFAULT 0,
      \`remaining_balance\`DECIMAL(12,2) NOT NULL DEFAULT 0,
      \`monthly_amount\`   DECIMAL(12,2) NOT NULL DEFAULT 0,
      \`status\`           ENUM('active','closed','consolidated') NOT NULL DEFAULT 'active',
      \`start_date\`       DATETIME     DEFAULT NULL,
      \`closed_date\`      DATETIME     DEFAULT NULL,
      \`consolidated_into\`INT          DEFAULT NULL,
      \`created_by\`       VARCHAR(100) NOT NULL DEFAULT '',
      PRIMARY KEY (\`id\`),
      KEY \`idx_el_employee\` (\`employee_id\`),
      CONSTRAINT \`fk_el_employee\` FOREIGN KEY (\`employee_id\`) REFERENCES \`employee\`(\`id\`) ON DELETE CASCADE,
      CONSTRAINT \`fk_el_payroll_type\` FOREIGN KEY (\`payroll_type_id\`) REFERENCES \`payroll_types\`(\`id\`) ON DELETE RESTRICT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);
  // ── Loan Payment History ───────────────────────────────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`loan_payment_history\` (
      \`id\`              BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
      \`loan_id\`         INT              NOT NULL,
      \`employee_id\`     INT              NOT NULL,
      \`amount_deducted\` DECIMAL(12,2)    NOT NULL,
      \`balance_before\`  DECIMAL(12,2)    NOT NULL,
      \`balance_after\`   DECIMAL(12,2)    NOT NULL,
      \`deduction_month\` VARCHAR(7)       NOT NULL,
      \`processed_at\`    DATETIME         NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`processed_by\`    VARCHAR(100)     NOT NULL DEFAULT 'system',
      PRIMARY KEY (\`id\`),
      KEY \`idx_lph_loan\`     (\`loan_id\`),
      KEY \`idx_lph_employee\` (\`employee_id\`),
      CONSTRAINT \`fk_lph_loan\`     FOREIGN KEY (\`loan_id\`)     REFERENCES \`employee_loans\`(\`id\`) ON DELETE RESTRICT,
      CONSTRAINT \`fk_lph_employee\` FOREIGN KEY (\`employee_id\`) REFERENCES \`employee\`(\`id\`)     ON DELETE RESTRICT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // ── Loan architecture migration ────────────────────────────────────────────
  // Add monthly_amount to employee_loans (source of truth, replacing ledger loan rows)
  try { await pool.query("ALTER TABLE `employee_loans` ADD COLUMN `monthly_amount` DECIMAL(12,2) NOT NULL DEFAULT 0"); } catch {}
  // Fix any NULL monthly_amounts left by earlier incomplete migrations
  try { await pool.query("UPDATE `employee_loans` SET `monthly_amount` = 0 WHERE `monthly_amount` IS NULL"); } catch {}
  // Populate monthly_amount from existing active ledger rows before removing them
  try {
    await pool.query(`
      UPDATE employee_loans el
      JOIN employee_deduction_ledger edl ON edl.loan_id = el.id AND edl.status = 'active'
      SET el.monthly_amount = edl.amount
      WHERE el.monthly_amount = 0 AND edl.amount > 0
    `);
  } catch {}
  // Remove loan rows from ledger — loan_payment_history is now the transaction record
  try { await pool.query("DELETE FROM `employee_deduction_ledger` WHERE `loan_id` IS NOT NULL"); } catch {}
  try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP FOREIGN KEY `fk_edl_loan`"); } catch {}
  try { await pool.query("ALTER TABLE `employee_deduction_ledger` DROP COLUMN `loan_id`"); } catch {}

  // Drop amount_per_deduction — monthly_amount on employee_loans is now the source of truth
  {
    const [[{ hasApd }]] = await pool.execute(
      "SELECT COUNT(*) AS hasApd FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_loans' AND COLUMN_NAME = 'amount_per_deduction'"
    );
    if (parseInt(hasApd) > 0) {
      await pool.query("ALTER TABLE `employee_loans` DROP COLUMN `amount_per_deduction`");
    }
  }
  // Drop notes column if still present (was removed in earlier migration)
  {
    const [[{ hasNotes }]] = await pool.execute(
      "SELECT COUNT(*) AS hasNotes FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee_loans' AND COLUMN_NAME = 'notes'"
    );
    if (parseInt(hasNotes) > 0) {
      await pool.query("ALTER TABLE `employee_loans` DROP COLUMN `notes`");
    }
  }
  // Architecture change 2026-08-12: start_date/closed_date → DATETIME, remove created_at
  try { await pool.query("ALTER TABLE `employee_loans` MODIFY COLUMN `start_date` DATETIME DEFAULT NULL"); } catch {}
  try { await pool.query("ALTER TABLE `employee_loans` MODIFY COLUMN `closed_date` DATETIME DEFAULT NULL"); } catch {}
  {
    const [[{ hasCa }]] = await pool.execute(
      "SELECT COUNT(*) AS hasCa FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='employee_loans' AND COLUMN_NAME='created_at'"
    );
    if (parseInt(hasCa) > 0) {
      await pool.query("ALTER TABLE `employee_loans` DROP COLUMN `created_at`");
    }
  }
  // Dedup: if multiple active ledger rows exist for same employee+payroll_type, keep the highest id (most recent)
  try { await pool.query(`
    UPDATE employee_deduction_ledger edl
    JOIN (
      SELECT employee_id, payroll_type_id, MAX(id) AS keep_id
      FROM employee_deduction_ledger
      WHERE status = 'active'
      GROUP BY employee_id, payroll_type_id
      HAVING COUNT(*) > 1
    ) dupes ON edl.employee_id = dupes.employee_id AND edl.payroll_type_id = dupes.payroll_type_id
    SET edl.status = 'inactive'
    WHERE edl.status = 'active' AND edl.id != dupes.keep_id
  `); } catch {}

  // `subject` is a readable label ("DELA CRUZ, JUAN D. - Premium"), not an id — the
  // ids it used to hold referenced schemas that have since changed. There is no
  // `entity` column: it was fully determined by `action` (see
  // AUDIT_ENTITY_BY_ACTION) and duplicated on every row. old_amount/new_amount
  // are generated from the logged JSON so the change is readable at a glance.
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`audit_log\` (
      \`id\`         BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
      \`action\`     VARCHAR(50)  NOT NULL,
      \`subject\`    VARCHAR(150) DEFAULT NULL,
      \`changed_by\` VARCHAR(100) NOT NULL,
      \`changed_at\` DATETIME(6)  NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
      \`old_value\`  JSON DEFAULT NULL,
      \`new_value\`  JSON DEFAULT NULL,
      \`old_amount\` DECIMAL(12,2) AS (JSON_UNQUOTE(JSON_EXTRACT(\`old_value\`, '$.amount'))) VIRTUAL,
      \`new_amount\` DECIMAL(12,2) AS (JSON_UNQUOTE(JSON_EXTRACT(\`new_value\`, '$.amount'))) VIRTUAL,
      PRIMARY KEY (\`id\`),
      KEY \`idx_audit_changed_at\` (\`changed_at\`),
      KEY \`idx_audit_action\` (\`action\`),
      KEY \`idx_audit_subject\` (\`subject\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Migrate an existing table to the shape above. Each step is skipped when the
  // column is already in its target state, so this is safe on every boot.
  try {
    const [cols] = await pool.query("SHOW COLUMNS FROM `audit_log`");
    const names = new Set(cols.map(c => c.Field));
    if (names.has("entity_id") && !names.has("subject"))
      await pool.query("ALTER TABLE `audit_log` CHANGE `entity_id` `subject` VARCHAR(150) NULL");
    if (!names.has("old_amount"))
      await pool.query("ALTER TABLE `audit_log` ADD COLUMN `old_amount` DECIMAL(12,2) AS (JSON_UNQUOTE(JSON_EXTRACT(`old_value`, '$.amount'))) VIRTUAL");
    if (!names.has("new_amount"))
      await pool.query("ALTER TABLE `audit_log` ADD COLUMN `new_amount` DECIMAL(12,2) AS (JSON_UNQUOTE(JSON_EXTRACT(`new_value`, '$.amount'))) VIRTUAL");
    if (names.has("entity")) {
      try { await pool.query("ALTER TABLE `audit_log` DROP INDEX `idx_audit_entity`"); } catch {}
      await pool.query("ALTER TABLE `audit_log` DROP COLUMN `entity`");
    }
  } catch (e) { console.warn("audit_log migration skipped:", e.message); }

  // A browsable view of audit_log. For deduction rows the JSON payload holds
  // nothing that subject/old_amount/new_amount do not already show, so it is
  // left out; for loan and employee rows its unique fields are flattened into
  // one `details` string. The underlying table keeps the raw payloads.
  try {
    await pool.query(`
      CREATE OR REPLACE VIEW \`audit_log_readable\` AS
      SELECT
        id,
        changed_at,
        changed_by,
        action,
        subject,
        old_amount,
        new_amount,
        CASE WHEN old_amount IS NOT NULL AND new_amount IS NOT NULL
             THEN ROUND(new_amount - old_amount, 2) END AS amount_change,
        CASE action
          WHEN 'create_loan' THEN CONCAT(
            'total ',   JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.total_amount')),
            ', loan #', JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.loan_id')))
          WHEN 'loan_deduction' THEN CONCAT(
            JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.loan_type')),
            ': deducted ', JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.deducted')),
            ', balance ',  JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.remaining_balance')))
          WHEN 'update_employee' THEN CONCAT(
            'status ', JSON_UNQUOTE(JSON_EXTRACT(old_value, '$.status')),
            ' -> ',    JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.status')),
            ', pay ',  JSON_UNQUOTE(JSON_EXTRACT(old_value, '$.basic_monthly_pay')),
            ' -> ',    JSON_UNQUOTE(JSON_EXTRACT(new_value, '$.basic_monthly_pay')))
          ELSE NULL
        END AS details
      FROM \`audit_log\`
    `);
  } catch (e) { console.warn("audit_log_readable view skipped:", e.message); }

  // Backfill: insert audit entries for loan ledger rows that were created before audit logging was added
  try { await pool.query(`
    INSERT INTO audit_log (action, subject, changed_by, changed_at, new_value)
    SELECT
      'create_loan', NULL, 'system', NULL, NULL
    FROM employee_loans WHERE FALSE
  `); } catch (e) { /* backfill no longer needed — loan rows removed from ledger */ }

  // Drop legacy computation_config table — mode is now stored as a string column on payroll_types
  try { await pool.query("DROP TABLE IF EXISTS `computation_config`"); } catch {}

  // Drop buggy increment_day and initial_step columns — salary is set explicitly on save
  {
    const [[{ hasIncrementDay }]] = await pool.execute(
      "SELECT COUNT(*) AS hasIncrementDay FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee' AND COLUMN_NAME = 'increment_day'"
    );
    if (parseInt(hasIncrementDay) > 0) {
      try { await pool.query("ALTER TABLE `employee` DROP COLUMN `increment_day`"); } catch {}
    }
  }
  {
    const [[{ hasInitialStep }]] = await pool.execute(
      "SELECT COUNT(*) AS hasInitialStep FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'employee' AND COLUMN_NAME = 'initial_step'"
    );
    if (parseInt(hasInitialStep) > 0) {
      try { await pool.query("ALTER TABLE `employee` DROP COLUMN `initial_step`"); } catch {}
    }
  }

  // Auto-update basic_monthly_pay for every active employee based on start_date.
  // Step 1 + floor(years_since_start_date / 3), capped at 8.
  // If user picked a custom starting step via the form, that's already saved as basic_monthly_pay.
  {
    const [employees] = await pool.execute(
      "SELECT id, position_id, start_date FROM `employee` WHERE status = 'active' AND position_id IS NOT NULL AND start_date IS NOT NULL"
    );
    for (const emp of employees) {
      const { currentStep } = computeCurrentStep(emp.start_date, 1);
      const stepInfo = await getStepAmount(emp.position_id, currentStep);
      if (stepInfo?.amount != null) {
        const newAmt = Number(stepInfo.amount);
        const [[cur]] = await pool.execute("SELECT basic_monthly_pay FROM `employee` WHERE id = ?", [emp.id]);
        if (cur && Number(cur.basic_monthly_pay) !== newAmt) {
          await pool.execute("UPDATE `employee` SET basic_monthly_pay = ? WHERE id = ?", [newAmt, emp.id]);
          try { await syncComputedDeductionsToLedger(emp.id, newAmt); } catch {}
        }
      }
    }
  }

  // ── 201 Files ──────────────────────────────────────────────────────────────
  await ddl(`
    CREATE TABLE IF NOT EXISTS \`employee_201_files\` (
      \`id\`          INT UNSIGNED   NOT NULL AUTO_INCREMENT,
      \`employee_id\` INT            NOT NULL,
      \`doc_item\`    CHAR(1)        NOT NULL COMMENT 'A–Q',
      \`filename\`    VARCHAR(255)   NOT NULL,
      \`mime_type\`   VARCHAR(100)   NOT NULL DEFAULT 'application/pdf',
      \`file_data\`   LONGBLOB       NOT NULL,
      \`uploaded_at\` DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (\`id\`),
      KEY \`idx_201_employee\` (\`employee_id\`),
      CONSTRAINT \`fk_201_employee\` FOREIGN KEY (\`employee_id\`) REFERENCES \`employee\`(\`id\`) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

}

// ── 201 File DB helpers ────────────────────────────────────────────────────────

export async function list201Files(employeeId) {
  const [rows] = await pool.execute(
    "SELECT `id`, `doc_item`, `filename`, `mime_type`, `uploaded_at` FROM `employee_201_files` WHERE `employee_id` = ? ORDER BY `doc_item`, `uploaded_at` DESC",
    [employeeId]
  );
  return rows;
}

export async function get201FileData(fileId) {
  const [[row]] = await pool.execute(
    "SELECT `file_data`, `mime_type`, `filename` FROM `employee_201_files` WHERE `id` = ?",
    [fileId]
  );
  return row ?? null;
}

export async function insert201File(employeeId, docItem, filename, mimeType, fileData) {
  const [result] = await pool.execute(
    "INSERT INTO `employee_201_files` (`employee_id`, `doc_item`, `filename`, `mime_type`, `file_data`) VALUES (?, ?, ?, ?, ?)",
    [employeeId, docItem, filename, mimeType, fileData]
  );
  return result.insertId;
}

export async function delete201File(fileId) {
  const [result] = await pool.execute(
    "DELETE FROM `employee_201_files` WHERE `id` = ?",
    [fileId]
  );
  return result.affectedRows > 0;
}

// ──────────────────────────────────────────────────────────────────────────────

// Which entity each action concerns. The old `entity` column stored this on every
// row even though it is fully determined by the action, so it lives here instead.
export const AUDIT_ENTITY_BY_ACTION = {
  create_deduction: "employee_deduction",
  update_deduction: "employee_deduction",
  remove_deduction: "employee_deduction",
  create_loan:      "employee_deduction",
  loan_deduction:   "employee_loan",
  update_employee:  "employee",
  export_view:      "archive",
  upsert_archive:   "archive",
  delete_archive:   "archive",
  leave_submit:     "leave_application",
  leave_action:     "leave_application",
  leave_section7:   "leave_application",
  leave_delete:     "leave_application",
  leave_signatory_create:  "leave_signatory",
  leave_signatory_update:  "leave_signatory",
  leave_signatory_delete:  "leave_signatory",
};

// `subject` is a human-readable label ("DELA CRUZ, JUAN D. - Premium"), not an id.
export async function insertAuditLog(action, subject, changedBy, oldValue, newValue) {
  await pool.execute(
    `INSERT INTO \`audit_log\` (action, subject, changed_by, old_value, new_value)
     VALUES (?, ?, ?, ?, ?)`,
    [action, subject != null ? String(subject) : null, changedBy,
     oldValue != null ? JSON.stringify(oldValue) : null,
     newValue != null ? JSON.stringify(newValue) : null]
  );
}

export async function getAuditLog({ q: search = "", entity = "", limit = 100 } = {}) {
  const safeLimit = Math.min(Math.max(parseInt(limit, 10) || 100, 1), 500);
  const clauses = [];
  const params = [];

  // Callers still filter by entity; resolve it to the actions it covers.
  const wanted = entity.trim();
  if (wanted) {
    const actions = Object.entries(AUDIT_ENTITY_BY_ACTION)
      .filter(([, e]) => e === wanted)
      .map(([a]) => a);
    if (actions.length === 0) return [];
    clauses.push(`action IN (${actions.map(() => "?").join(",")})`);
    params.push(...actions);
  }

  if (search.trim()) {
    clauses.push("(changed_by LIKE ? OR subject LIKE ? OR action LIKE ? OR CAST(old_value AS CHAR) LIKE ? OR CAST(new_value AS CHAR) LIKE ?)");
    const like = `%${search.trim()}%`;
    params.push(like, like, like, like, like);
  }
  const where = clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
  return q(`SELECT * FROM \`audit_log\` ${where} ORDER BY changed_at DESC LIMIT ${safeLimit}`, params);
}

async function _seedEmployeeDefaults(employeeId, basic) {
  const round = v => Math.round(Number(v) * 100) / 100;

  const computed = [
    ["gsis_life_ret",   round(basic * 0.09)],
    ["pagibig_premium", round(basic * 0.02)],
    ["medic",           round(basic * 0.025)],
    ["union_dues",      round(basic * 0.005)],
  ];
  for (const [name, amount] of computed) {
    const [pt] = await q("SELECT id FROM `payroll_types` WHERE name = ?", [name]);
    if (!pt) continue;
    const already = await q(
      "SELECT id FROM `employee_deduction_ledger` WHERE employee_id = ? AND payroll_type_id = ? AND status = 'active' LIMIT 1",
      [employeeId, pt.id]
    );
    if (already[0]) continue; // already has an active entry — don't overwrite
    await q(
      `INSERT INTO \`employee_deduction_ledger\`
         (id, employee_id, payroll_type_id, amount, created_date, status, created_by)
       VALUES (?, ?, ?, ?, NOW(), 'active', 'system')`,
      [await randomLedgerId(), employeeId, pt.id, amount]
    );
  }
}

// ── AUTH / USER MANAGEMENT ────────────────────────────────────────────────────

export async function getPendingUsers() {
  return q("SELECT id, username, status, created_at FROM `users` WHERE status = 'pending' ORDER BY created_at ASC");
}

export async function approveUser(userId) {
  await q("UPDATE `users` SET status = 'active' WHERE id = ? AND status = 'pending'", [userId]);
  const rows = await q("SELECT id, username, status FROM `users` WHERE id = ?", [userId]);
  return rows[0] ?? null;
}

export async function rejectUser(userId) {
  await q("UPDATE `users` SET status = 'rejected' WHERE id = ? AND status = 'pending'", [userId]);
  const rows = await q("SELECT id, username, status FROM `users` WHERE id = ?", [userId]);
  return rows[0] ?? null;
}

export async function getAllManagedUsers() {
  return q(
    `SELECT u.id, u.username, u.status, u.created_at, u.role_id, u.employee_id,
            e.name AS employee_name, e.employee_number
     FROM \`users\` u
     LEFT JOIN \`employee\` e ON e.id = u.employee_id
     WHERE LOWER(u.username) <> ?
     ORDER BY u.created_at DESC`,
    [ADMIN_USERNAME]
  );
}

export async function blockUser(userId) {
  await q(
    "UPDATE `users` SET status = 'blocked' WHERE id = ? AND LOWER(username) <> ? AND status = 'active'",
    [userId, ADMIN_USERNAME]
  );
  const rows = await q("SELECT id, username, status FROM `users` WHERE id = ?", [userId]);
  return rows[0] ?? null;
}

export async function unblockUser(userId) {
  await q("UPDATE `users` SET status = 'active' WHERE id = ? AND status = 'blocked'", [userId]);
  const rows = await q("SELECT id, username, status FROM `users` WHERE id = ?", [userId]);
  return rows[0] ?? null;
}

export async function deleteSystemUser(userId) {
  const rows = await q("SELECT id, username FROM `users` WHERE id = ?", [userId]);
  const user = rows[0];
  if (!user || user.username.toLowerCase() === ADMIN_USERNAME) return null;
  await q("DELETE FROM `users` WHERE id = ?", [userId]);
  return user;
}

// ── EMPLOYEES ─────────────────────────────────────────────────────────────────

const EMP_COLS = `
  e.id, e.employee_number, e.name, e.status, e.basic_monthly_pay,
  e.first_quincena_override,
  e.position_id, e.division_id,
  e.start_date, e.granted_step,
  ep_cur.start_date                   AS appointed_date,
  COALESCE(p.position_title, '')      AS position,
  COALESCE(d.name, pd.name, '')       AS office_division,
  CASE e.employment_type WHEN 2 THEN 'Contractual' ELSE 'Regular' END AS employment_type,
  sg.grade_code                       AS salary_rank,
  sg.grade_code                       AS salary_grade_code,
  e.separation_reason,
  (SELECT u.username FROM \`users\` u WHERE u.employee_id = e.id LIMIT 1) AS linked_username
`;
const EMP_JOINS = `
  LEFT JOIN \`position\`         p      ON p.id      = e.position_id
  LEFT JOIN \`salary_grade\`     sg     ON sg.id     = p.salary_grade_id
  LEFT JOIN \`division\`         d      ON d.id      = e.division_id
  LEFT JOIN \`division\`         pd     ON pd.id     = p.division_id
  LEFT JOIN \`employee_position\` ep_cur ON ep_cur.employee_id = e.id
                                        AND ep_cur.status = 'active'
                                        AND ep_cur.position_id = e.position_id
`;

export async function searchEmployees(q_str, limit = 100) {
  const version = await getVersion();
  const cacheKey = `emp:search:v${version}:${q_str}:${limit}`;
  const cached = await cacheGet(cacheKey);
  if (cached) return cached;

  // mysql2 prepared statements (pool.execute) reject a bound LIMIT parameter
  // ("Incorrect arguments to mysqld_stmt_execute") — inline a sanitized integer instead.
  const safeLimit = Math.max(0, parseInt(limit, 10) || 100);
  let rows;
  if (!q_str || !q_str.trim()) {
    rows = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.status = 'active' ORDER BY e.name ASC LIMIT ${safeLimit}`
    );
  } else {
    rows = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.status = 'active' AND (
         LOWER(e.employee_number) LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(e.name)            LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(p.position_title)  LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(d.name)            LIKE LOWER(CONCAT('%', ?, '%'))
       ) ORDER BY e.name ASC LIMIT ${safeLimit}`,
      [q_str.trim(), q_str.trim(), q_str.trim(), q_str.trim()]
    );
  }

  const result = addEmpStep(rows);
  await cacheSet(cacheKey, result, 30);
  return result;
}

export async function searchEmployeesPage(q_str, limit, offset) {
  const version = await getVersion();
  const cacheKey = `emp:page:v${version}:${q_str}:${limit}:${offset}`;
  const cached = await cacheGet(cacheKey);
  if (cached) return cached;

  // mysql2 prepared statements (pool.execute) don't support COUNT(*) OVER() window functions.
  // Use a separate COUNT query instead.
  // mysql2 prepared statements (pool.execute) reject bound LIMIT/OFFSET parameters
  // ("Incorrect arguments to mysqld_stmt_execute") — inline sanitized integers instead.
  const safeLimit = Math.max(0, parseInt(limit, 10) || 0);
  const safeOffset = Math.max(0, parseInt(offset, 10) || 0);

  let rows, total;
  if (!q_str || !q_str.trim()) {
    const [countRes] = await pool.execute("SELECT COUNT(*) AS cnt FROM `employee` e WHERE e.status = 'active'");
    total = parseInt(countRes[0].cnt, 10);
    rows = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.status = 'active' ORDER BY e.name ASC LIMIT ${safeLimit} OFFSET ${safeOffset}`
    );
  } else {
    const str = q_str.trim();
    const [countRes] = await pool.execute(
      `SELECT COUNT(*) AS cnt FROM \`employee\` e
       LEFT JOIN \`position\` p ON p.id = e.position_id
       LEFT JOIN \`division\` d ON d.id = e.division_id
       WHERE e.status = 'active' AND (
         LOWER(e.employee_number)  LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(e.name)             LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(p.position_title)   LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(d.name)             LIKE LOWER(CONCAT('%', ?, '%'))
       )`,
      [str, str, str, str]
    );
    total = parseInt(countRes[0].cnt, 10);
    rows = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.status = 'active' AND (
         LOWER(e.employee_number)  LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(e.name)             LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(p.position_title)   LIKE LOWER(CONCAT('%', ?, '%')) OR
         LOWER(d.name)             LIKE LOWER(CONCAT('%', ?, '%'))
       ) ORDER BY e.name ASC LIMIT ${safeLimit} OFFSET ${safeOffset}`,
      [str, str, str, str]
    );
  }

  const out = { rows: addEmpStep(rows), total };
  await cacheSet(cacheKey, out, 30);
  return out;
}

export async function getEmployeeById(employeeId) {
  const rows = await q(
    `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.id = ? AND e.status != 'deleted' LIMIT 1`,
    [employeeId]
  );
  return addEmpStep(rows)[0] ?? null;
}

// Finds or creates a division by name and returns its INT id.
async function resolveOrCreateDivisionId(name) {
  if (!name?.trim()) return null;
  const n = name.trim();
  const rows = await q("SELECT id FROM `division` WHERE LOWER(name) = LOWER(?)", [n]);
  if (rows[0]) return rows[0].id;
  const [result] = await pool.execute("INSERT IGNORE INTO `division` (name) VALUES (?)", [n]);
  if (result.insertId) return result.insertId;
  const rows2 = await q("SELECT id FROM `division` WHERE LOWER(name) = LOWER(?)", [n]);
  return rows2[0]?.id ?? null;
}

// Finds or creates a position by title and returns its INT id.
async function resolveOrCreatePositionId(title) {
  if (!title?.trim()) return null;
  const t = title.trim();
  const rows = await q("SELECT id FROM `position` WHERE LOWER(position_title) = LOWER(?)", [t]);
  if (rows[0]) return rows[0].id;
  const [result] = await pool.execute("INSERT IGNORE INTO `position` (position_title) VALUES (?)", [t]);
  if (result.insertId) return result.insertId;
  const rows2 = await q("SELECT id FROM `position` WHERE LOWER(position_title) = LOWER(?)", [t]);
  return rows2[0]?.id ?? null;
}

// Computes the current salary step.
// initialStep = the step the employee held as of startDate (user-set, default 1).
// +1 every completed 3-year cycle from startDate, capped at Step 8.
// Example: initialStep=2, startDate=Aug 2020, today=Aug 2026 → 6 yrs / 3 = 2 increments → Step 4.
function computeCurrentStep(startDate, initialStep = 1) {
  const base = Math.min(Math.max(1, parseInt(initialStep, 10) || 1), 8);
  if (!startDate) return { currentStep: base, nextIncrementDate: null, currentCycleStart: null };
  const start = new Date(startDate);
  const today = new Date();
  if (isNaN(start.getTime())) return { currentStep: base, nextIncrementDate: null, currentCycleStart: null };
  let completeYears = today.getFullYear() - start.getFullYear();
  const hadAnniversary = today.getMonth() > start.getMonth() ||
    (today.getMonth() === start.getMonth() && today.getDate() >= start.getDate());
  if (!hadAnniversary) completeYears--;
  const completedIncrements = Math.max(0, Math.floor(completeYears / 3));
  const currentStep = Math.min(base + completedIncrements, 8);
  const cycleStart = new Date(start);
  cycleStart.setFullYear(start.getFullYear() + completedIncrements * 3);
  const currentCycleStart = cycleStart.toISOString().slice(0, 10);
  let nextIncrementDate = null;
  if (currentStep < 8) {
    const nxt = new Date(start);
    nxt.setFullYear(start.getFullYear() + (completedIncrements + 1) * 3);
    nextIncrementDate = nxt.toISOString().slice(0, 10);
  }
  return { currentStep, nextIncrementDate, currentCycleStart };
}

function addEmpStep(rows) {
  if (!Array.isArray(rows)) return rows;
  return rows.map(r => ({
    ...r,
    // granted_step = boss directly assigned this exact step; null = auto-compute from date
    salary_grade_step: r.granted_step != null
      ? r.granted_step
      : computeCurrentStep(r.appointed_date ?? r.start_date).currentStep,
  }));
}

async function getStepAmount(positionId, step) {
  if (!positionId) return null;
  const [rows] = await pool.execute(`
    SELECT sgs.id AS step_id, sgs.amount
    FROM \`position\` p
    JOIN salary_grade sg ON sg.id = p.salary_grade_id
    JOIN salary_grade_step sgs ON sgs.salary_grade_id = sg.id AND sgs.step = ?
    WHERE p.id = ?
    ORDER BY sgs.effective_date DESC LIMIT 1
  `, [step, positionId]);
  return rows[0] ?? null;
}

export async function addEmployee(data) {
  const empNum = data.employeeNumber?.trim();
  if (!empNum) throw Object.assign(new Error("Employee number is required"), { status: 400 });
  // Generate unique 7-digit surrogate PK
  let id;
  for (let attempt = 0; attempt < 10; attempt++) {
    const candidate = String(randomInt(1000000, 9999999));
    const existing = await q("SELECT id FROM `employee` WHERE id = ? LIMIT 1", [candidate]);
    if (!existing[0]) { id = candidate; break; }
  }
  if (!id) throw new Error("Could not generate unique employee ID");

  let basic = Number(data.basicMonthlyPay) || 0;
  const round = v => Math.round(v * 100) / 100;

  const positionId       = data.positionId       != null ? data.positionId       : await resolveOrCreatePositionId(data.position);
  const divisionId       = data.divisionId        != null ? data.divisionId       : await resolveOrCreateDivisionId(data.officeDivision);

  await q(
    `INSERT INTO \`employee\`
       (id, employee_number, name, basic_monthly_pay, position_id, employment_type, division_id, start_date, granted_step)
     VALUES (?,?,?,?,?,?,?,?,?)`,
    [
      id,
      empNum,
      data.name.trim(),
      basic,
      positionId,
      data.employmentType === 'Contractual' ? 2 : 1,
      divisionId,
      data.startDate || null,
      data.initialStep ? parseInt(data.initialStep, 10) : null,
    ]
  );

  // Record initial appointment in employee_position history
  if (positionId) {
    const today = new Date().toISOString().slice(0, 10);
    const stepStart = data.startDate ? String(data.startDate).slice(0, 10) : today;
    const [epResult] = await pool.execute(
      "INSERT INTO `employee_position` (employee_id, position_id, start_date, movement_type, status) VALUES (?,?,?,'APPOINTMENT','active')",
      [id, positionId, stepStart]
    );
    if (epResult.insertId) {
      // Boss picks exact step → use it directly. No step picked → compute from appointed date.
      const currentStep = data.initialStep
        ? parseInt(data.initialStep, 10)
        : computeCurrentStep(data.startDate || null).currentStep;
      const stepInfo = await getStepAmount(positionId, currentStep);
      if (stepInfo?.amount != null) {
        const stepAmt = Number(stepInfo.amount);
        await q("UPDATE `employee` SET basic_monthly_pay=? WHERE id=?", [stepAmt, id]);
        basic = stepAmt;
      }
    }
  }

  await _seedEmployeeDefaults(id, basic);

  await incrementVersion();
  const rows = await q(`SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.id = ?`, [id]);
  return addEmpStep(rows)[0];
}

export async function updateEmployee(id, data) {
  // Fetch current position_id before updating so we can track changes
  const [currentRows] = await pool.execute(
    "SELECT position_id, start_date FROM `employee` WHERE id = ? AND status = 'active'", [id]
  );
  const currentPositionId = currentRows[0]?.position_id ?? null;
  const currentStartDate  = currentRows[0]?.start_date ?? null;

  const fields = [];
  const values = [];
  const add = (col, val) => { fields.push(`\`${col}\` = ?`); values.push(val); };

  if (data.employeeNumber         !== undefined) add("employee_number",          data.employeeNumber?.trim() || null);
  if (data.name                   !== undefined) add("name",                     data.name.trim());
  if (data.basicMonthlyPay        !== undefined) add("basic_monthly_pay",        Number(data.basicMonthlyPay));
  if ("firstQuincenaOverride" in data)           add("first_quincena_override",  data.firstQuincenaOverride != null ? Number(data.firstQuincenaOverride) : null);
  if ("startDate" in data)                       add("start_date",               data.startDate || null);
  if ("initialStep" in data)                     add("granted_step",             data.initialStep != null ? (parseInt(data.initialStep, 10) || null) : null);

  const resolvedPositionId  = data.positionId != null ? data.positionId      : (data.position       ? await resolveOrCreatePositionId(data.position)                                  : undefined);
  const resolvedDivisionId  = data.divisionId != null ? data.divisionId : (data.officeDivision ? await resolveOrCreateDivisionId(data.officeDivision) : undefined);
  if (resolvedPositionId  !== undefined) add("position_id",  resolvedPositionId  || null);
  if (data.employmentType !== undefined) add("employment_type", data.employmentType === 'Contractual' ? 2 : 1);
  if (resolvedDivisionId  !== undefined) add("division_id",  resolvedDivisionId  || null);

  if (fields.length === 0) {
    const rows = await q(`SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.id = ?`, [id]);
    return addEmpStep(rows)[0] ?? null;
  }
  values.push(id);

  await q(`UPDATE \`employee\` SET ${fields.join(", ")} WHERE id = ? AND status = 'active'`, values);


  // Sync employee_position history if the position changed
  if (resolvedPositionId != null) {
    const newPosId = parseInt(resolvedPositionId, 10);
    const oldPosId = currentPositionId ? parseInt(currentPositionId, 10) : null;
    const today = new Date().toISOString().slice(0, 10);
    let epInsertId = null;

    if (newPosId !== oldPosId) {
      if (oldPosId) {
        // Supersede steps under all currently active positions first
        const [oldEpRows] = await pool.execute(
          "SELECT id FROM `employee_position` WHERE employee_id=? AND status='active'",
          [id]
        );
        await pool.execute(
          "UPDATE `employee_position` SET status='superseded', end_date=? WHERE employee_id=? AND status='active'",
          [today, id]
        );
      }
      const appointedDate = ("startDate" in data && data.startDate) ? String(data.startDate).slice(0, 10) : today;
      const validMovements = ["APPOINTMENT", "PROMOTION", "TRANSFER"];
      const movementType = validMovements.includes(data.movementType) ? data.movementType : "APPOINTMENT";
      const [epResult] = await pool.execute(
        "INSERT INTO `employee_position` (employee_id, position_id, start_date, movement_type, status) VALUES (?,?,?,?,'active')",
        [id, newPosId, appointedDate, movementType]
      );
      epInsertId = epResult.insertId;

      // Boss picks exact step → use it directly. No step → compute from appointed date.
      const grantedStep = data.initialStep ? parseInt(data.initialStep, 10) : null;
      await pool.execute("UPDATE `employee` SET granted_step=? WHERE id=?", [grantedStep, id]);
      const currentStep = grantedStep ?? computeCurrentStep(appointedDate).currentStep;
      const stepInfo = await getStepAmount(newPosId, currentStep);
      if (stepInfo?.amount != null) {
        await pool.execute("UPDATE `employee` SET basic_monthly_pay=? WHERE id=?", [stepInfo.amount, id]);
        await syncComputedDeductionsToLedger(id, stepInfo.amount);
      } else {
        // Fallback: use Step 1 salary if the computed step doesn't exist in this salary grade
        const [step1Rows] = await pool.execute(`
          SELECT sgs.amount FROM \`position\` p
          JOIN salary_grade sg  ON sg.id = p.salary_grade_id
          JOIN salary_grade_step sgs ON sgs.salary_grade_id = sg.id AND sgs.step = 1
          WHERE p.id = ? ORDER BY sgs.effective_date DESC LIMIT 1
        `, [newPosId]);
        if (step1Rows[0]?.amount != null) {
          await pool.execute("UPDATE `employee` SET basic_monthly_pay=? WHERE id=?", [step1Rows[0].amount, id]);
          await syncComputedDeductionsToLedger(id, step1Rows[0].amount);
        }
      }
    } // end if (newPosId !== oldPosId)
  } // end if (resolvedPositionId != null)

  const positionChangedThisSave = resolvedPositionId != null && parseInt(resolvedPositionId, 10) !== parseInt(currentPositionId, 10);

  // Same-position: update appointment date and/or step
  const shouldRecomputeStep = !positionChangedThisSave && ("startDate" in data || "initialStep" in data);
  if (shouldRecomputeStep) {
    const finalPositionId = resolvedPositionId != null ? parseInt(resolvedPositionId, 10) : currentPositionId;

    // Update employee_position.start_date when appointment date changes
    if ("startDate" in data && data.startDate) {
      await pool.execute(
        "UPDATE `employee_position` SET start_date=? WHERE employee_id=? AND status='active'",
        [String(data.startDate).slice(0, 10), id]
      );
    }

    const appointedDate = "startDate" in data ? (data.startDate || null) : currentStartDate;

    // Boss picks exact step → use it directly. No step picked → compute from date.
    const targetStep = data.initialStep != null
      ? parseInt(data.initialStep, 10)
      : computeCurrentStep(appointedDate).currentStep;

    if (finalPositionId && targetStep) {
      const stepInfo = await getStepAmount(finalPositionId, targetStep);
      if (stepInfo?.amount != null) {
        const amt = Number(stepInfo.amount);
        await pool.execute("UPDATE `employee` SET basic_monthly_pay=? WHERE id=?", [amt, id]);
        try { await syncComputedDeductionsToLedger(id, amt); } catch {}
      }
    }
  } else if (data.basicMonthlyPay !== undefined) {
    // Direct pay edit — sync deductions to new amount
    try { await syncComputedDeductionsToLedger(id, Number(data.basicMonthlyPay)); } catch {}
  }

  if (fields.length > 0) await incrementVersion();

  const rows = await q(`SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.id = ? AND e.status = 'active'`, [id]);
  return addEmpStep(rows)[0] ?? null;
}

const PERMANENT_SEPARATIONS  = ["Resigned", "Retired", "Terminated"];
const TEMPORARY_SEPARATIONS   = ["On Leave", "Suspended"];

export async function softDeleteEmployee(id, reason = null, startDate = null, endDate = null) {
  const rows = await q("SELECT id, name FROM `employee` WHERE id = ? AND status = 'active'", [id]);
  const emp = rows[0];
  if (!emp) return null;
  await q("UPDATE `employee` SET status = 'inactive', separation_reason = ?, deactivated_at = NOW() WHERE id = ?", [reason, id]);
  if (!reason || PERMANENT_SEPARATIONS.includes(reason)) {
    await q("UPDATE `users` SET employee_id = NULL WHERE employee_id = ?", [id]);
  }
  if (reason && PERMANENT_SEPARATIONS.includes(reason)) {
    await q("UPDATE `employee_deduction_ledger` SET status = 'inactive' WHERE employee_id = ?", [id]);
  }
  if (reason && TEMPORARY_SEPARATIONS.includes(reason)) {
    // Close any existing open suspension records before inserting a new one
    await q(
      "UPDATE `employee_suspension` SET end_date = NOW() WHERE employee_id = ? AND (end_date IS NULL OR end_date > NOW())",
      [id]
    );
    const start = startDate ? new Date(startDate).toISOString().slice(0, 19).replace("T", " ") : new Date().toISOString().slice(0, 19).replace("T", " ");
    const end = endDate ? new Date(endDate).toISOString().slice(0, 19).replace("T", " ") : null;
    const reasonCode = reason === "On Leave" ? 1 : reason === "Suspended" ? 2 : null;
    await q(
      "INSERT INTO `employee_suspension` (employee_id, start_date, end_date, reason, status) VALUES (?, ?, ?, ?, 'served')",
      [id, start, end, reasonCode]
    );
  }
  await incrementVersion();
  return emp;
}

export async function getInactiveEmployees(q_str = "", limit = 5) {
  // mysql2 prepared statements (pool.execute) reject a bound LIMIT parameter
  // ("Incorrect arguments to mysqld_stmt_execute") — inline a sanitized integer instead.
  const safeLimit = Math.max(0, parseInt(limit, 10) || 5);
  let emps;
  if (!q_str || !q_str.trim()) {
    emps = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.status = 'inactive' ORDER BY e.name ASC LIMIT ${safeLimit}`
    );
  } else {
    const like = `%${q_str.trim()}%`;
    emps = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS}
       WHERE e.status = 'inactive' AND (
         e.name LIKE ? OR e.employee_number LIKE ? OR e.separation_reason LIKE ? OR
         COALESCE(p.position_title, '') LIKE ? OR COALESCE(d.name, pd.name, '') LIKE ?
       ) ORDER BY e.name ASC LIMIT ${safeLimit}`,
      [like, like, like, like, like]
    );
  }
  if (emps.length === 0) return emps;
  // Attach suspension/leave history for each employee
  const ids = emps.map(e => e.id);
  const placeholders = ids.map(() => "?").join(",");
  const suspRows = await q(
    `SELECT es.employee_id, es.reason, es.start_date, es.end_date, es.status, es.created_at,
            DATEDIFF(COALESCE(es.end_date, CURDATE()), es.start_date) AS days
     FROM \`employee_suspension\` es
     WHERE es.employee_id IN (${placeholders})
     ORDER BY es.start_date DESC`,
    ids
  );
  const suspMap = {};
  for (const r of suspRows) {
    if (!suspMap[r.employee_id]) suspMap[r.employee_id] = [];
    suspMap[r.employee_id].push({
      reason:     r.reason,
      start_date: r.start_date,
      end_date:   r.end_date,
      days:       r.days,
      status:     r.status,
      created_at: r.created_at,
    });
  }
  return addEmpStep(emps).map(e => ({ ...e, suspension_history: suspMap[e.id] ?? [] }));
}

export async function getEmployeeProfile(id) {
  const rows = addEmpStep(await q(
    `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS} WHERE e.id = ?`, [id]
  ));
  if (!rows[0]) return null;
  const [epRow] = await q(
    `SELECT start_date FROM \`employee_position\` WHERE employee_id = ? ORDER BY start_date ASC LIMIT 1`, [id]
  );
  const [userRow] = await q(
    `SELECT avatar_key FROM \`users\` WHERE employee_id = ? LIMIT 1`, [id]
  );
  const suspRows = await q(`
    SELECT
      CASE reason WHEN 1 THEN 'On Leave' WHEN 2 THEN 'Suspended' ELSE 'Unknown' END AS reason,
      start_date, end_date,
      TIMESTAMPDIFF(HOUR,   start_date, COALESCE(end_date, NOW())) AS total_hours,
      TIMESTAMPDIFF(MINUTE, start_date, COALESCE(end_date, NOW())) % 60 AS extra_minutes
    FROM \`employee_suspension\` WHERE employee_id = ? ORDER BY start_date DESC
  `, [id]);
  return {
    employee: {
      ...rows[0],
      position_start_date: epRow?.start_date ?? null,
      avatarDataUrl: userRow?.avatar_key?.startsWith("data:") ? userRow.avatar_key : null,
    },
    suspensions: suspRows,
  };
}

export async function activateEmployee(id) {
  const rows = await q("SELECT id, name FROM `employee` WHERE id = ? AND status = 'inactive'", [id]);
  const emp = rows[0];
  if (!emp) return null;
  await q("UPDATE `employee` SET status = 'active', separation_reason = NULL, deactivated_at = NULL WHERE id = ?", [id]);
  // Close any open suspension/leave records for this employee
  await q(
    "UPDATE `employee_suspension` SET end_date = NOW() WHERE employee_id = ? AND (end_date IS NULL OR end_date > NOW())",
    [id]
  );
  await incrementVersion();
  return emp;
}

// ── PAYROLL ───────────────────────────────────────────────────────────────────

export async function getLivePayslips(q_str = "", limit = 5) {
  const version = await getVersion();
  const cacheKey = `payslips:v${version}:${q_str.trim()}:${limit}`;
  const cached = await cacheGet(cacheKey);
  if (cached) return cached;

  // mysql2 prepared statements (pool.execute) reject a bound LIMIT parameter
  // ("Incorrect arguments to mysqld_stmt_execute") — inline a sanitized integer instead.
  const safeLimit = Math.max(0, parseInt(limit, 10) || 5);
  let employees;
  if (q_str.trim()) {
    employees = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS}
       WHERE e.status = 'active' AND LOWER(e.name) LIKE LOWER(CONCAT('%',?,'%'))
       ORDER BY e.name ASC LIMIT ${safeLimit}`,
      [q_str.trim()]
    );
  } else {
    employees = await q(
      `SELECT ${EMP_COLS} FROM \`employee\` e ${EMP_JOINS}
       WHERE e.status = 'active'
       ORDER BY e.name ASC LIMIT ${safeLimit}`
    );
  }

  if (employees.length === 0) return [];

  const empIds = employees.map(e => e.id);

  // Fetch gross earning amounts from payroll_types (side='gross') + value_config — same for all employees
  const grossRows = await q(
    `SELECT pt.name, pt.display_name, COALESCE(vc.value, 0) AS amount
     FROM \`payroll_types\` pt
     LEFT JOIN \`value_config\` vc ON vc.payroll_type_id = pt.id AND vc.is_active = 1
     WHERE pt.is_active = 1 AND pt.side = 'gross' AND pt.name != 'basic_pay'`
  );
  const grossEarnings = {};
  const typeLabels = {};
  for (const r of grossRows) {
    grossEarnings[r.name] = Number(r.amount ?? 0);
    typeLabels[r.name] = r.display_name || r.name;
  }
  const peraAmount      = grossEarnings.pera ?? 0;
  const customEarnTotal = Object.entries(grossEarnings)
    .filter(([n]) => n !== 'pera')
    .reduce((s, [, v]) => s + v, 0);

  // Fetch computed (rate/fixed) deduction defaults — used as fallback when no ledger entry exists
  const computedDedRows = await q(
    `SELECT pt.name, pt.display_name, pt.side, pt.mode, vc.value
     FROM \`payroll_types\` pt
     LEFT JOIN \`value_config\` vc ON vc.payroll_type_id = pt.id AND vc.is_active = 1
     WHERE pt.is_active = 1 AND pt.side != 'gross' AND pt.mode IN (2, 3) AND vc.value IS NOT NULL`
  );

  // Fetch active deductions from ledger (source of truth) — keyed by payroll_types.name
  // Exclude gross-side types (e.g. PERA) — they belong in allowances, not deductions
  const dedRows = await q(
    `SELECT edl.employee_id, pt.name, pt.display_name, pt.side, edl.amount
     FROM \`employee_deduction_ledger\` edl
     JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id
     WHERE edl.status = 'active'
       AND pt.is_active = 1
       AND pt.side != 'gross'
       AND edl.employee_id IN (${empIds.map(() => "?").join(",")})`,
    empIds
  );
  // { empId: { ptName: amount } }, dedSide: { ptName: side }
  const dedMap = {};
  const dedSide = {};
  for (const r of dedRows) {
    if (!dedMap[r.employee_id]) dedMap[r.employee_id] = {};
    dedMap[r.employee_id][r.name] = Number(r.amount);
    dedSide[r.name] = r.side;
    typeLabels[r.name] = r.display_name || r.name;
  }
  // Pre-populate dedSide/typeLabels for computed types (needed even without ledger entries)
  for (const r of computedDedRows) {
    dedSide[r.name] = r.side;
    typeLabels[r.name] = r.display_name || r.name;
  }

  const round = v => Math.round(Number(v) * 100) / 100;

  const result = employees.map(emp => {
    const basicPay  = Number(emp.basic_monthly_pay) || 0;
    const grossPay  = round(basicPay + peraAmount + customEarnTotal);
    const ded       = { ...dedMap[emp.id] };
    // Fill in computed (rate/fixed) deduction types that have no ledger entry yet
    for (const r of computedDedRows) {
      if (ded[r.name] != null) continue;
      ded[r.name] = r.mode === 2 ? round(basicPay * Number(r.value)) : round(Number(r.value));
    }

    // Sum only what's actually in the employee's ledger
    const lessDedTotal = round(
      Object.entries(ded)
        .filter(([name]) => dedSide[name] === 'less_deduction')
        .reduce((s, [, amt]) => s + amt, 0)
    );
    const deductionColTotal = round(
      Object.entries(ded)
        .filter(([name]) => dedSide[name] === 'deduction')
        .reduce((s, [, amt]) => s + amt, 0)
    );
    const netPay = round(grossPay - lessDedTotal - deductionColTotal);

    // Honor manual quincena override; fall back to equal split
    const firstQ  = emp.first_quincena_override != null
      ? round(Number(emp.first_quincena_override))
      : round(netPay / 2);
    const secondQ = round(netPay - firstQ);

    return {
      id: emp.id,
      employee_id: emp.id,
      employee_name: emp.name,
      employee_number: emp.employee_number,
      position: emp.position,
      office_division: emp.office_division,
      employment_type: emp.employment_type,
      basic_pay: basicPay,
      allowances: grossEarnings,
      deductions: ded,
      type_labels: typeLabels,
      gross_pay: grossPay,
      total_deductions: round(lessDedTotal + deductionColTotal),
      net_pay: netPay,
      first_quincena: firstQ,
      second_quincena: secondQ,
    };
  });

  await cacheSet(cacheKey, result, 30);
  return result;
}

export async function getPayrollStats() {
  const cached = await cacheGet("payroll:stats");
  if (cached) return cached;

  const [empRow] = await q("SELECT COUNT(*) AS cnt FROM `employee` WHERE status = 'active'");
  const [netRow] = await q("SELECT COALESCE(SUM(basic_monthly_pay), 0) AS total FROM `employee` WHERE status = 'active'");
  const breakdownRows = await q(`
    SELECT CASE e.employment_type WHEN 2 THEN 'contractual' ELSE 'regular' END AS type,
           COUNT(*) AS emp_count,
           COALESCE(SUM(e.basic_monthly_pay), 0) AS salary_total
    FROM \`employee\` e
    WHERE e.status = 'active'
    GROUP BY e.employment_type
  `);

  const byType = {};
  for (const row of breakdownRows) {
    byType[row.type] = { count: parseInt(row.emp_count, 10), salary: parseFloat(row.salary_total) };
  }

  // Employees per division, broken down by employment type (Regular / Contractual / etc.)
  const divisionRows = await q(`
    SELECT COALESCE(d.name, 'Unassigned') AS division,
           CASE e.employment_type WHEN 2 THEN 'Contractual' ELSE 'Regular' END AS type,
           COUNT(*) AS emp_count
    FROM \`employee\` e
    LEFT JOIN \`division\` d ON d.id = e.division_id
    WHERE e.status = 'active'
    GROUP BY division, type
    ORDER BY division ASC
  `);
  const divisionMap = new Map();
  for (const row of divisionRows) {
    if (!divisionMap.has(row.division)) divisionMap.set(row.division, { division: row.division, total: 0 });
    const entry = divisionMap.get(row.division);
    const count = parseInt(row.emp_count, 10);
    entry[row.type] = count;
    entry.total += count;
  }
  const employeesByDivision = [...divisionMap.values()].sort((a, b) => b.total - a.total);

  // Total deductions per category (from ledger + payroll_types) — exclude earnings (gross)
  const deductionRows = await q(`
    SELECT pt.category, COALESCE(SUM(edl.amount), 0) AS total
    FROM \`employee_deduction_ledger\` edl
    JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id AND pt.is_active = 1 AND pt.side != 'gross'
    JOIN \`employee\` e ON e.id = edl.employee_id AND e.status = 'active'
    WHERE edl.status = 'active'
    GROUP BY pt.category
  `);
  const deductionsByCategory = {};
  for (const r of deductionRows) {
    if (r.category) deductionsByCategory[r.category] = parseFloat(r.total);
  }

  // Salary rank — all active employees sorted by gross total (basic + allowances)
  const salaryRankRows = await q(`
    SELECT e.id, e.name, e.basic_monthly_pay,
      COALESCE(earn.allowances, 0) AS allowances,
      COALESCE(earn.allowance_detail, '[]') AS allowance_detail,
      e.basic_monthly_pay + COALESCE(earn.allowances, 0) AS gross_total
    FROM \`employee\` e
    LEFT JOIN (
      SELECT edl.employee_id,
        SUM(edl.amount) AS allowances,
        CONCAT('[', GROUP_CONCAT(
          JSON_OBJECT('name', pt.display_name, 'amount', edl.amount)
          ORDER BY pt.display_name
        ), ']') AS allowance_detail
      FROM \`employee_deduction_ledger\` edl
      JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id
        AND pt.is_active = 1 AND pt.side = 'gross' AND pt.name != 'basic_pay'
      WHERE edl.status = 'active'
      GROUP BY edl.employee_id
    ) earn ON earn.employee_id = e.id
    WHERE e.status = 'active'
    ORDER BY gross_total DESC
  `);

  // Earnings totals: basic salary from employee table + gross ledger items (e.g. PERA)
  const [peraRow] = await q(`
    SELECT COALESCE(SUM(edl.amount), 0) AS total_pera
    FROM \`employee_deduction_ledger\` edl
    JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id AND pt.is_active = 1 AND pt.side = 'gross' AND pt.name != 'basic_pay'
    JOIN \`employee\` e ON e.id = edl.employee_id AND e.status = 'active'
    WHERE edl.status = 'active'
  `);

  // Total loan burden per employee (top 5 highest) — only active deduction types, exclude earnings
  const loanRows = await q(`
    SELECT e.name, e.id,
      COALESCE(SUM(d.amount), 0) AS total_deductions,
      e.basic_monthly_pay,
      ROUND((COALESCE(SUM(d.amount), 0) / NULLIF(e.basic_monthly_pay, 0)) * 100, 1) AS burden_pct
    FROM \`employee\` e
    LEFT JOIN (
      SELECT edl.employee_id, edl.amount
      FROM \`employee_deduction_ledger\` edl
      JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id AND pt.is_active = 1 AND pt.side = 'deduction'
      WHERE edl.status = 'active'
    ) d ON d.employee_id = e.id
    WHERE e.status = 'active'
    GROUP BY e.id, e.name, e.basic_monthly_pay
    ORDER BY burden_pct DESC
    LIMIT 5
  `);

  // All employees with their current status for Employee Status tab
  const allEmpStatusRows = await q(`
    SELECT e.id, e.name, e.status,
      CASE WHEN e.status = 'active' THEN 'Active' ELSE COALESCE(e.separation_reason, 'Inactive') END AS current_status,
      es.start_date, es.end_date
    FROM \`employee\` e
    LEFT JOIN \`employee_suspension\` es ON es.employee_id = e.id AND (es.end_date IS NULL OR es.end_date > NOW())
    ORDER BY
      CASE WHEN e.status = 'active' THEN 0 ELSE 1 END ASC,
      CASE WHEN e.separation_reason IN ('On Leave','Suspended') THEN 0 ELSE 1 END ASC,
      e.name ASC
  `);

  // Count by status
  const statusRows = await q(`
    SELECT
      SUM(status = 'active') AS active_cnt,
      SUM(status = 'inactive' AND separation_reason = 'On Leave') AS on_leave_cnt,
      SUM(status = 'inactive' AND separation_reason = 'Suspended') AS suspended_cnt,
      SUM(status = 'inactive' AND (separation_reason NOT IN ('On Leave','Suspended') OR separation_reason IS NULL)) AS other_inactive_cnt
    FROM \`employee\`
  `);

  // Top employees by quincena take-home — respects first_quincena_override, date-aware
  const topQuincenaRows = await q(`
    SELECT e.id, e.name, e.first_quincena_override,
      e.basic_monthly_pay + COALESCE(earn.allowances, 0) AS gross_total,
      COALESCE(ded.total_deductions, 0) AS total_deductions
    FROM \`employee\` e
    LEFT JOIN (
      SELECT edl.employee_id, SUM(edl.amount) AS allowances
      FROM \`employee_deduction_ledger\` edl
      JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id AND pt.is_active = 1 AND pt.side = 'gross' AND pt.name != 'basic_pay'
      WHERE edl.status = 'active'
      GROUP BY edl.employee_id
    ) earn ON earn.employee_id = e.id
    LEFT JOIN (
      SELECT edl.employee_id, SUM(edl.amount) AS total_deductions
      FROM \`employee_deduction_ledger\` edl
      JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id AND pt.is_active = 1 AND pt.side != 'gross'
      WHERE edl.status = 'active'
      GROUP BY edl.employee_id
    ) ded ON ded.employee_id = e.id
    WHERE e.status = 'active'
  `);

  // Archive count
  const [archiveRow] = await q("SELECT COUNT(*) AS cnt FROM `payroll_archive`");

  // Average net pay (basic_monthly_pay - total deductions, excluding earnings)
  const [avgRow] = await q(`
    SELECT AVG(e.basic_monthly_pay - COALESCE(d.total, 0)) AS avg_net
    FROM \`employee\` e
    LEFT JOIN (
      SELECT edl.employee_id, SUM(edl.amount) AS total
      FROM \`employee_deduction_ledger\` edl
      JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id AND pt.side != 'gross'
      WHERE edl.status = 'active'
      GROUP BY edl.employee_id
    ) d ON d.employee_id = e.id
    WHERE e.status = 'active'
  `);

  const totalBasicSalary  = parseFloat(netRow.total);
  const totalPera         = parseFloat(peraRow.total_pera ?? 0);
  const totalGrossSalary  = totalBasicSalary + totalPera;

  const stats = {
    totalEmployees:       parseInt(empRow.cnt, 10),
    totalNetPaid:         totalBasicSalary,
    totalBasicSalary,
    totalPera,
    totalGrossSalary,
    regularEmployees:     byType['regular']?.count     ?? 0,
    contractualEmployees: byType['contractual']?.count ?? 0,
    regularSalary:        byType['regular']?.salary    ?? 0,
    contractualSalary:    byType['contractual']?.salary ?? 0,
    deductionsByCategory,
    employeesByDivision,
    topLoanBurden:        loanRows.map(r => ({
      name:             r.name,
      id:               r.id,
      totalDeductions:  parseFloat(r.total_deductions),
      basicPay:         parseFloat(r.basic_monthly_pay),
      burdenPct:        parseFloat(r.burden_pct ?? 0),
    })),
    archiveCount:         parseInt(archiveRow.cnt, 10),
    avgNetPay:            parseFloat(avgRow.avg_net ?? 0),
    salaryRank:           salaryRankRows.map((r, i) => ({
      rank:            i + 1,
      id:              r.id,
      name:            r.name,
      basicPay:        parseFloat(r.basic_monthly_pay),
      allowances:      parseFloat(r.allowances),
      grossTotal:      parseFloat(r.gross_total),
      allowanceDetail: (() => { try { return JSON.parse(r.allowance_detail ?? '[]'); } catch { return []; } })(),
    })),
    employeeStatus: {
      active:        parseInt(statusRows[0]?.active_cnt         ?? 0, 10),
      onLeave:       parseInt(statusRows[0]?.on_leave_cnt       ?? 0, 10),
      suspended:     parseInt(statusRows[0]?.suspended_cnt      ?? 0, 10),
      otherInactive: parseInt(statusRows[0]?.other_inactive_cnt ?? 0, 10),
      all: allEmpStatusRows.map(r => ({
        id:     r.id,
        name:   r.name,
        status: r.current_status,
        since:  r.start_date ?? null,
        until:  r.end_date   ?? null,
      })),
    },
    quincenaLabel:        new Date().getDate() <= 15 ? "1st Quincena" : "2nd Quincena",
    topQuincena:          (() => {
      const isFirstQ = new Date().getDate() <= 15;
      return topQuincenaRows.map(r => {
        const grossTotal       = parseFloat(r.gross_total);
        const totalDeductions  = parseFloat(r.total_deductions);
        const netPay           = grossTotal - totalDeductions;
        const q1Override       = r.first_quincena_override != null ? parseFloat(r.first_quincena_override) : null;
        const q1               = q1Override ?? Math.round(netPay / 2 * 100) / 100;
        const q2               = Math.round((netPay - q1) * 100) / 100;
        return { id: r.id, name: r.name, grossTotal, totalDeductions, quincena: isFirstQ ? q1 : q2 };
      })
      .sort((a, b) => b.quincena - a.quincena)
      .map((r, i) => ({ ...r, rank: i + 1 }));
    })(),
  };

  await cacheSet("payroll:stats", stats, 30);
  return stats;
}

// ── EMPLOYEE DEDUCTIONS ───────────────────────────────────────────────────────

// Syncs computed (rate/fixed) deduction amounts into employee_deduction_ledger so
// every deduction type has a ledger row — needed for history and archive snapshots.
/**
 * Keep only the single most-recent 'closed' count reversible for a deduction; any other
 * 'closed' row becomes permanently 'inactive'. Applied after every edit that pushes a new
 * count (Employee Deductions, Payroll Types rate/fixed changes, auto-sync), so Reversal
 * always has exactly one unambiguous predecessor to roll back to.
 */
async function randomLedgerId() {
  while (true) {
    const id = Math.floor(1_000_000_000 + Math.random() * 9_000_000_000);
    const [[{ cnt }]] = await pool.execute(
      "SELECT COUNT(*) AS cnt FROM `employee_deduction_ledger` WHERE id = ?", [id]
    );
    if (cnt === 0) return id;
  }
}

// Audit keys read as "09008:GSIS Salary Loan" rather than a pair of UUIDs. The exact
// ids stay in the logged JSON, so nothing is lost. Falls back to the ids if a lookup fails.
const auditKey = (meta, employeeId, ptId) =>
  meta?.employee_name && meta?.display_name
    ? `${meta.employee_name} - ${meta.display_name}`
    : `${employeeId}:${ptId}`;

// Writes (creates or updates in place) the active ledger row for one employee/payroll-type
// pair, and records the change in audit_log. Replaces the old close-row-and-insert-new-row
// versioning — the ledger now holds exactly one row per active deduction, and history lives
// in audit_log instead.
async function writeLedgerAmount(employeeId, ptId, newAmount, createdBy, actorForLog) {
  const existing = await q(
    "SELECT id, amount FROM `employee_deduction_ledger` WHERE employee_id = ? AND payroll_type_id = ? AND status = 'active' LIMIT 1",
    [employeeId, ptId]
  );
  const [[meta]] = await pool.execute(
    `SELECT e.name AS employee_name, e.employee_number, pt.display_name
     FROM \`employee\` e, \`payroll_types\` pt
     WHERE e.id = ? AND pt.id = ? LIMIT 1`,
    [employeeId, ptId]
  );
  if (existing[0]) {
    const oldAmount = Number(existing[0].amount);
    if (oldAmount === newAmount) return;
    await q(
      "UPDATE `employee_deduction_ledger` SET amount = ?, created_date = NOW(), created_by = ? WHERE id = ?",
      [newAmount, createdBy, existing[0].id]
    );
    await insertAuditLog(
      "update_deduction", auditKey(meta, employeeId, ptId), actorForLog,
      { employee_name: meta?.employee_name, payroll_type: meta?.display_name, amount: oldAmount, employee_id: employeeId, payroll_type_id: ptId },
      { employee_name: meta?.employee_name, payroll_type: meta?.display_name, amount: newAmount, employee_id: employeeId, payroll_type_id: ptId }
    );
  } else {
    await q(
      `INSERT INTO \`employee_deduction_ledger\`
         (id, employee_id, payroll_type_id, amount, created_date, status, created_by)
       VALUES (?, ?, ?, ?, NOW(), 'active', ?)`,
      [await randomLedgerId(), employeeId, ptId, newAmount, createdBy]
    );
    await insertAuditLog(
      "create_deduction", auditKey(meta, employeeId, ptId), actorForLog,
      null,
      { employee_name: meta?.employee_name, payroll_type: meta?.display_name, amount: newAmount, employee_id: employeeId, payroll_type_id: ptId }
    );
  }
}

// Deletes the active ledger row for one employee/payroll-type pair and logs the removal.
async function removeLedgerAmount(employeeId, ptId, actorForLog) {
  const existing = await q(
    "SELECT id, amount FROM `employee_deduction_ledger` WHERE employee_id = ? AND payroll_type_id = ? AND status = 'active' LIMIT 1",
    [employeeId, ptId]
  );
  if (!existing[0]) return;
  const [[meta]] = await pool.execute(
    `SELECT e.name AS employee_name, e.employee_number, pt.display_name
     FROM \`employee\` e, \`payroll_types\` pt
     WHERE e.id = ? AND pt.id = ? LIMIT 1`,
    [employeeId, ptId]
  );
  await q("DELETE FROM `employee_deduction_ledger` WHERE id = ?", [existing[0].id]);
  await insertAuditLog(
    "remove_deduction", auditKey(meta, employeeId, ptId), actorForLog,
    { employee_name: meta?.employee_name, payroll_type: meta?.display_name, amount: Number(existing[0].amount), employee_id: employeeId, payroll_type_id: ptId },
    null
  );
}

export async function syncComputedDeductionsToLedger(employeeId, basicPay) {
  const types = await q(
    `SELECT pt.id AS pt_id, pt.mode, vc.value
     FROM \`payroll_types\` pt
     LEFT JOIN \`value_config\` vc ON vc.payroll_type_id = pt.id AND vc.is_active = 1
     WHERE pt.is_active = 1 AND pt.name != 'basic_pay' AND pt.mode IN (2, 3)`
  );

  for (const t of types) {
    if (t.value === null) continue;
    const newAmount = Math.round((t.mode === 2 ? Number(basicPay) * Number(t.value) : Number(t.value)) * 100) / 100;

    const existing = await q(
      "SELECT amount, created_by FROM `employee_deduction_ledger` WHERE employee_id = ? AND payroll_type_id = ? AND status = 'active' LIMIT 1",
      [employeeId, t.pt_id]
    );
    if (existing[0] && existing[0].created_by === 'user') continue;
    if (existing[0] && Number(existing[0].amount) === newAmount) continue;
    await writeLedgerAmount(employeeId, t.pt_id, newAmount, 'system', 'system');
  }
}

export async function getEmployeeDeductions(employeeId, { sync = true } = {}) {
  const empRows = await q("SELECT basic_monthly_pay FROM `employee` WHERE id = ?", [employeeId]);
  if (!empRows[0]) return {};
  const basicPay = Number(empRows[0].basic_monthly_pay ?? 0);

  if (sync) await syncComputedDeductionsToLedger(employeeId, basicPay);

  const rows = await q(
    `SELECT pt.name, pt.side, edl.amount
     FROM \`payroll_types\` pt
     LEFT JOIN \`employee_deduction_ledger\` edl
       ON edl.payroll_type_id = pt.id
       AND edl.employee_id = ?
       AND edl.status = 'active'
     WHERE pt.is_active = 1 AND pt.name != 'basic_pay'
     ORDER BY pt.sort_order ASC`,
    [employeeId]
  );

  const result = {};
  const sides = {};
  for (const r of rows) {
    result[r.name] = r.amount !== null ? Number(r.amount) : null;
    sides[r.name] = r.side ?? 'deduction';
  }
  result._sides = sides;

  return result;
}


export async function upsertEmployeeDeduction(employeeId, column, value, actorUsername = 'user') {
  const ptRows = await q(
    "SELECT id AS pt_id FROM `payroll_types` WHERE name = ? AND is_active = 1 LIMIT 1",
    [column]
  );
  if (!ptRows[0]) {
    const err = new Error(`Invalid deduction column: ${column}`);
    err.status = 400;
    throw err;
  }
  const ptId = ptRows[0].pt_id;

  if (value === null) {
    await removeLedgerAmount(employeeId, ptId, actorUsername);
  } else {
    await writeLedgerAmount(employeeId, ptId, Number(value), 'user', actorUsername);
  }

  await incrementVersion();
  return getEmployeeDeductions(employeeId);
}

export async function getEmployeeLedger(employeeId) {
  return q(
    `SELECT edl.id, edl.payroll_type_id AS pt_id,
            pt.name, pt.display_name, pt.category, pt.side, pt.is_active AS pt_is_active,
            edl.amount,
            edl.created_date,
            edl.status, edl.created_by
     FROM \`employee_deduction_ledger\` edl
     JOIN \`payroll_types\` pt ON pt.id = edl.payroll_type_id
     WHERE edl.employee_id = ?
     ORDER BY pt.is_active DESC, pt.sort_order ASC, edl.created_date ASC`,
    [employeeId]
  );
}

// ── LOOKUP TABLE CRUD ─────────────────────────────────────────────────────────

export async function getPositions(q_str = "", limit = 5) {
  const HAS_STEPS = `(p.salary_grade_id IS NOT NULL AND EXISTS (
    SELECT 1 FROM salary_grade_step sgs WHERE sgs.salary_grade_id = p.salary_grade_id
  )) AS has_steps`;
  const EMP_COUNT = `(SELECT COUNT(*) FROM \`employee\` e WHERE e.position_id = p.id AND e.status = 'active') AS employee_count`;
  const safeLimit = Math.max(0, parseInt(limit, 10) || 5);
  if (!q_str || !q_str.trim()) {
    return q(`SELECT p.id, p.position_title AS name, p.division_id, p.salary_grade_id,
        sg.grade_code AS default_salary_grade, ${HAS_STEPS}, ${EMP_COUNT}
      FROM \`position\` p
      LEFT JOIN salary_grade sg ON sg.id = p.salary_grade_id
      ORDER BY p.last_accessed_at DESC, p.position_title ASC LIMIT ${safeLimit}`);
  }
  const like = `%${q_str.trim()}%`;
  return q(`SELECT p.id, p.position_title AS name, p.division_id, p.salary_grade_id,
      sg.grade_code AS default_salary_grade, ${HAS_STEPS}, ${EMP_COUNT}
    FROM \`position\` p
    LEFT JOIN salary_grade sg ON sg.id = p.salary_grade_id
    WHERE p.position_title LIKE ? OR COALESCE(sg.grade_code, '') LIKE ?
    ORDER BY p.last_accessed_at DESC, p.position_title ASC LIMIT ${safeLimit}`, [like, like]);
}
export async function hitPosition(id) {
  await pool.execute("UPDATE `position` SET last_accessed_at = NOW() WHERE id = ?", [id]);
}
async function resolveSalaryGradeId(gradeCode) {
  if (!gradeCode) return null;
  const rows = await q("SELECT id FROM `salary_grade` WHERE grade_code = ? LIMIT 1", [gradeCode]);
  return rows[0]?.id ?? null;
}
export async function createPosition(name, divisionId, defaultSalaryGrade) {
  const sgId = await resolveSalaryGradeId(defaultSalaryGrade);
  const [result] = await pool.execute(
    "INSERT INTO `position` (position_title, division_id, salary_grade_id) VALUES (?,?,?)",
    [name.trim(), divisionId || null, sgId]
  );
  const rows = await q(`SELECT p.id, p.position_title AS name, p.division_id, p.salary_grade_id,
    sg.grade_code AS default_salary_grade FROM \`position\` p
    LEFT JOIN salary_grade sg ON sg.id = p.salary_grade_id WHERE p.id = ?`, [result.insertId]);
  return rows[0];
}
export async function updatePosition(id, name, divisionId, defaultSalaryGrade) {
  const sgId = await resolveSalaryGradeId(defaultSalaryGrade);
  await q(
    "UPDATE `position` SET position_title = ?, division_id = ?, salary_grade_id = ? WHERE id = ?",
    [name.trim(), divisionId || null, sgId, id]
  );
  const rows = await q(`SELECT p.id, p.position_title AS name, p.division_id, p.salary_grade_id,
    sg.grade_code AS default_salary_grade FROM \`position\` p
    LEFT JOIN salary_grade sg ON sg.id = p.salary_grade_id WHERE p.id = ?`, [id]);
  return rows[0] ?? null;
}
export async function deletePosition(id) {
  await q("UPDATE `employee` SET position_id = NULL WHERE position_id = ?", [id]);
  const rows = await q("SELECT id FROM `position` WHERE id = ?", [id]);
  if (!rows[0]) return null;
  await q("DELETE FROM `position` WHERE id = ?", [id]);
  return rows[0];
}

export async function getDivisions() {
  return q("SELECT id, name FROM `division` ORDER BY name ASC");
}
export async function createDivision(name) {
  const [result] = await pool.execute("INSERT INTO `division` (name) VALUES (?)", [name.trim()]);
  const rows = await q("SELECT id, name FROM `division` WHERE id = ?", [result.insertId]);
  return rows[0];
}
export async function updateDivision(id, name) {
  await q("UPDATE `division` SET name = ? WHERE id = ?", [name.trim(), id]);
  const rows = await q("SELECT id, name FROM `division` WHERE id = ?", [id]);
  return rows[0] ?? null;
}
export async function deleteDivision(id) {
  await q("UPDATE `employee` SET division_id = NULL WHERE division_id = ?", [id]);
  await q("UPDATE `position` SET division_id = NULL WHERE division_id = ?", [id]);
  const rows = await q("SELECT id FROM `division` WHERE id = ?", [id]);
  if (!rows[0]) return null;
  await q("DELETE FROM `division` WHERE id = ?", [id]);
  return rows[0];
}

// ── PAYROLL TYPE HELPERS ──────────────────────────────────────────────────────

const PAYROLL_TYPE_SELECT = `
  SELECT pt.id, pt.name, pt.display_name, pt.category, pt.side, pt.is_active, pt.sort_order, pt.merge_in_excel, pt.is_loan,
    CASE pt.mode WHEN 1 THEN 'manual' WHEN 2 THEN 'rate' WHEN 3 THEN 'fixed' ELSE 'manual' END AS computation_mode,
    CASE WHEN vc.value_type = 'RATE'  THEN vc.value ELSE NULL END AS rate,
    CASE WHEN vc.value_type = 'FIXED' THEN vc.value ELSE NULL END AS amount_fixed
  FROM \`payroll_types\` pt
  LEFT JOIN \`value_config\` vc ON vc.payroll_type_id = pt.id AND vc.is_active = 1`;

function modeInt(cm) {
  return cm === 'rate' ? 2 : cm === 'fixed' ? 3 : 1;
}

async function upsertComputationConfig(ptId, computation_mode, rate, amount_fixed) {
  const mode = modeInt(computation_mode);

  // mode lives directly on payroll_types — computation_config is a 3-row enum, not touched here
  await q("UPDATE `payroll_types` SET mode = ? WHERE id = ?", [mode, ptId]);

  if (mode === 2 || mode === 3) {
    const vtype = mode === 2 ? 'RATE' : 'FIXED';
    const vval  = mode === 2 ? (rate != null ? Number(rate) : null) : (amount_fixed != null ? Number(amount_fixed) : null);

    const currentVC = await q(
      "SELECT value_type, value FROM `value_config` WHERE payroll_type_id = ? AND is_active = 1 LIMIT 1",
      [ptId]
    );
    const sameVConfig = currentVC[0]
      && currentVC[0].value_type === vtype
      && Number(currentVC[0].value) === vval;

    if (!sameVConfig) {
      await q("UPDATE `value_config` SET is_active = 0, end_date = CURDATE() WHERE payroll_type_id = ? AND is_active = 1", [ptId]);
      await q(
        "INSERT INTO `value_config` (payroll_type_id, value_type, value, start_date, is_active) VALUES (?,?,?,CURDATE(),1)",
        [ptId, vtype, vval]
      );
    }

    const employees = await q("SELECT id, basic_monthly_pay FROM `employee`");
    for (const emp of employees) {
      const basicPay  = Number(emp.basic_monthly_pay ?? 0);
      const newAmount = mode === 2 ? Math.round(basicPay * Number(vval) * 100) / 100 : Number(vval);
      await writeLedgerAmount(emp.id, ptId, newAmount, 'system', 'system');
    }
  } else {
    await q("UPDATE `value_config` SET is_active = 0, end_date = CURDATE() WHERE payroll_type_id = ? AND is_active = 1", [ptId]);
  }
}

// ── DEDUCTION TYPE CRUD ───────────────────────────────────────────────────────

export async function getDeductionTypes({ includeInactive = false } = {}) {
  const where = includeInactive
    ? "WHERE pt.side IN ('deduction','less_deduction')"
    : "WHERE pt.is_active = 1 AND pt.side IN ('deduction','less_deduction')";
  return q(`${PAYROLL_TYPE_SELECT} ${where} ORDER BY pt.sort_order ASC, pt.name ASC`);
}

async function getNextSortOrder(side) {
  const rows = await q(
    "SELECT COALESCE(MAX(sort_order), -10) AS max_so FROM `payroll_types` WHERE side = ?",
    [side]
  );
  return (rows[0]?.max_so ?? -10) + 10;
}

async function shiftSortOrdersUp(side, fromValue, excludeId) {
  const params = excludeId != null ? [side, fromValue, excludeId] : [side, fromValue];
  const excludeClause = excludeId != null ? " AND id != ?" : "";
  await q(
    `UPDATE \`payroll_types\` SET sort_order = sort_order + 10 WHERE side = ? AND sort_order >= ?${excludeClause} ORDER BY sort_order DESC`,
    params
  );
}

export async function createDeductionType({ name, display_name, category, side, computation_mode, rate, amount_fixed, sort_order, merge_in_excel, is_loan }) {
  const effectiveSide = side || 'deduction';
  const targetSort = sort_order != null ? parseInt(sort_order, 10) : await getNextSortOrder(effectiveSide);
  const conflict = await q(
    "SELECT id FROM `payroll_types` WHERE side = ? AND sort_order = ? LIMIT 1",
    [effectiveSide, targetSort]
  );
  if (conflict.length > 0) await shiftSortOrdersUp(effectiveSide, targetSort, null);
  const result = await q(
    "INSERT INTO `payroll_types` (name, display_name, category, side, sort_order, merge_in_excel, is_loan) VALUES (?,?,?,?,?,?,?)",
    [name, display_name || name, category || null, effectiveSide, targetSort, merge_in_excel ? 1 : 0, is_loan ? 1 : 0]
  );
  const id = result.insertId;
  await upsertComputationConfig(id, computation_mode, rate, amount_fixed);
  await incrementVersion();
  const rows = await q(`${PAYROLL_TYPE_SELECT} WHERE pt.id = ?`, [id]);
  return rows[0];
}

export async function updateDeductionType(id, { display_name, category, side, computation_mode, rate, amount_fixed, sort_order, is_active, merge_in_excel, is_loan }) {
  const effectiveSide = side || 'deduction';
  const targetSort = sort_order != null ? parseInt(sort_order, 10) : 0;
  const conflict = await q(
    "SELECT id FROM `payroll_types` WHERE side = ? AND sort_order = ? AND id != ? LIMIT 1",
    [effectiveSide, targetSort, id]
  );
  if (conflict.length > 0) await shiftSortOrdersUp(effectiveSide, targetSort, id);
  await q(
    "UPDATE `payroll_types` SET display_name=?, category=?, side=?, sort_order=?, is_active=?, merge_in_excel=?, is_loan=? WHERE id=?",
    [display_name, category || null, effectiveSide, targetSort, is_active !== false ? 1 : 0, merge_in_excel ? 1 : 0, is_loan ? 1 : 0, id]
  );
  await upsertComputationConfig(id, computation_mode, rate, amount_fixed);
  await incrementVersion();
  const rows = await q(`${PAYROLL_TYPE_SELECT} WHERE pt.id = ?`, [id]);
  return rows[0] ?? null;
}

export async function softDeleteDeductionType(id) {
  const check = await q("SELECT COUNT(*) AS cnt FROM `employee_deduction_ledger` WHERE payroll_type_id = ?", [id]);
  if (parseInt(check[0].cnt) > 0) {
    await q("UPDATE `payroll_types` SET is_active = 0 WHERE id = ?", [id]);
    return { deleted: false, deactivated: true };
  }
  await q("DELETE FROM `value_config` WHERE payroll_type_id = ?", [id]);
  await q("DELETE FROM `payroll_types` WHERE id = ?", [id]);
  await incrementVersion();
  return { deleted: true, deactivated: false };
}

// ── EARNING TYPE CRUD ─────────────────────────────────────────────────────────

export async function getEarningTypes({ includeInactive = false } = {}) {
  const where = includeInactive
    ? "WHERE pt.side = 'gross'"
    : "WHERE pt.is_active = 1 AND pt.side = 'gross'";
  return q(`${PAYROLL_TYPE_SELECT} ${where} ORDER BY pt.sort_order ASC, pt.name ASC`);
}

export async function createEarningType({ name, display_name, computation_mode, rate, amount_fixed, sort_order }) {
  const targetSort = sort_order != null ? parseInt(sort_order, 10) : await getNextSortOrder('gross');
  const conflict = await q(
    "SELECT id FROM `payroll_types` WHERE side = 'gross' AND sort_order = ? LIMIT 1",
    [targetSort]
  );
  if (conflict.length > 0) await shiftSortOrdersUp('gross', targetSort, null);
  const result = await q(
    "INSERT INTO `payroll_types` (name, display_name, category, side, sort_order) VALUES (?,?,NULL,'gross',?)",
    [name, display_name || name, targetSort]
  );
  const id = result.insertId;
  await upsertComputationConfig(id, computation_mode, rate, amount_fixed);
  await incrementVersion();
  const rows = await q(`${PAYROLL_TYPE_SELECT} WHERE pt.id = ?`, [id]);
  return rows[0];
}

export async function updateEarningType(id, { display_name, computation_mode, rate, amount_fixed, sort_order, is_active }) {
  const targetSort = sort_order != null ? parseInt(sort_order, 10) : 0;
  const conflict = await q(
    "SELECT id FROM `payroll_types` WHERE side = 'gross' AND sort_order = ? AND id != ? LIMIT 1",
    [targetSort, id]
  );
  if (conflict.length > 0) await shiftSortOrdersUp('gross', targetSort, id);
  await q(
    "UPDATE `payroll_types` SET display_name=?, sort_order=?, is_active=? WHERE id=?",
    [display_name, targetSort, is_active !== false ? 1 : 0, id]
  );
  await upsertComputationConfig(id, computation_mode, rate, amount_fixed);
  await incrementVersion();
  const rows = await q(`${PAYROLL_TYPE_SELECT} WHERE pt.id = ?`, [id]);
  return rows[0] ?? null;
}

export async function softDeleteEarningType(id) {
  const typeRow = await q("SELECT name FROM `payroll_types` WHERE id = ?", [id]);
  if (typeRow[0]?.name === 'basic_pay') throw Object.assign(new Error("Basic Pay cannot be deleted."), { status: 400 });
  const check = await q("SELECT COUNT(*) AS cnt FROM `employee_deduction_ledger` WHERE payroll_type_id = ?", [id]);
  if (parseInt(check[0].cnt) > 0) {
    await q("UPDATE `payroll_types` SET is_active = 0 WHERE id = ?", [id]);
    return { deleted: false, deactivated: true };
  }
  await q("DELETE FROM `value_config` WHERE payroll_type_id = ?", [id]);
  await q("DELETE FROM `payroll_types` WHERE id = ?", [id]);
  await incrementVersion();
  return { deleted: true, deactivated: false };
}

// ── EMPLOYEE LOANS ────────────────────────────────────────────────────────────

export async function getEmployeeLoans(employeeId) {
  return q(`
    SELECT el.*, pt.display_name AS loan_type_name, pt.category AS loan_category, pt.name AS loan_type_col
    FROM employee_loans el
    LEFT JOIN payroll_types pt ON pt.id = el.payroll_type_id
    WHERE el.employee_id = ?
    ORDER BY el.id DESC
  `, [employeeId]);
}

export async function getLoanPaymentHistory(employeeId) {
  return q(`
    SELECT lph.*, pt.display_name AS loan_type_name
    FROM loan_payment_history lph
    LEFT JOIN employee_loans el ON el.id = lph.loan_id
    LEFT JOIN payroll_types pt ON pt.id = el.payroll_type_id
    WHERE lph.employee_id = ?
    ORDER BY lph.processed_at DESC
  `, [employeeId]);
}

export async function createLoan({ employeeId, payrollTypeId, totalAmount, amountPerDeduction, startDate, createdBy }) {
  const total   = Number(totalAmount);
  const monthly = Number(amountPerDeduction);
  const conn    = await pool.getConnection();
  try {
    await conn.beginTransaction();

    const [result] = await conn.execute(
      `INSERT INTO employee_loans (employee_id, payroll_type_id, total_amount, remaining_balance, monthly_amount, start_date, created_by)
       VALUES (?,?,?,?,?,?,?)`,
      [employeeId, payrollTypeId, total, total, monthly, startDate, createdBy]
    );
    const loanId = result.insertId;

    await conn.commit();

    // Log the loan creation so it appears in Change History
    const [[loanMeta]] = await pool.execute(
      `SELECT e.name AS employee_name, e.employee_number, pt.display_name FROM \`employee\` e, \`payroll_types\` pt WHERE e.id = ? AND pt.id = ? LIMIT 1`,
      [employeeId, payrollTypeId]
    );
    await insertAuditLog(
      "create_loan", auditKey(loanMeta, employeeId, payrollTypeId), createdBy,
      null,
      { employee_name: loanMeta?.employee_name, payroll_type: loanMeta?.display_name, amount: monthly, total_amount: total, loan_id: loanId, employee_id: employeeId, payroll_type_id: payrollTypeId }
    );

    const rows = await q(`
      SELECT el.*, pt.display_name AS loan_type_name, pt.category AS loan_category
      FROM employee_loans el JOIN payroll_types pt ON pt.id = el.payroll_type_id
      WHERE el.id = ?`, [loanId]);
    return rows[0];
  } catch (e) {
    await conn.rollback();
    throw e;
  } finally {
    conn.release();
  }
}

export async function consolidateLoan({ oldLoanId, newLoanAmount, amountPerDeduction, startDate, createdBy }) {
  const oldLoan = (await q("SELECT * FROM employee_loans WHERE id = ?", [oldLoanId]))[0];
  if (!oldLoan) throw Object.assign(new Error("Loan not found."), { status: 404 });
  if (oldLoan.status !== 'active') throw Object.assign(new Error("Loan is not active."), { status: 400 });

  const consolidatedTotal = Number(oldLoan.remaining_balance) + Number(newLoanAmount);
  const monthly           = Number(amountPerDeduction);
  const conn              = await pool.getConnection();
  try {
    await conn.beginTransaction();

    // Create new consolidated loan with the new monthly amount
    const [result] = await conn.execute(
      `INSERT INTO employee_loans (employee_id, payroll_type_id, total_amount, remaining_balance, monthly_amount, start_date, created_by)
       VALUES (?,?,?,?,?,?,?)`,
      [oldLoan.employee_id, oldLoan.payroll_type_id, consolidatedTotal, consolidatedTotal, monthly, startDate, createdBy]
    );
    const newLoanId = result.insertId;

    // Close old loan
    await conn.execute(
      "UPDATE employee_loans SET status='consolidated', consolidated_into=?, closed_date=NOW() WHERE id=?",
      [newLoanId, oldLoanId]
    );

    await conn.commit();
    const rows = await q(`
      SELECT el.*, pt.display_name AS loan_type_name, pt.category AS loan_category
      FROM employee_loans el JOIN payroll_types pt ON pt.id = el.payroll_type_id
      WHERE el.id = ?`, [newLoanId]);
    return { newLoan: rows[0], consolidatedTotal };
  } catch (e) {
    await conn.rollback();
    throw e;
  } finally {
    conn.release();
  }
}

export async function closeLoan(loanId, closedBy) {
  const loan = (await q("SELECT * FROM employee_loans WHERE id = ?", [loanId]))[0];
  if (!loan) throw Object.assign(new Error("Loan not found."), { status: 404 });
  await q("UPDATE employee_loans SET status='closed', closed_date=NOW() WHERE id=?", [loanId]);
  return { closed: true };
}

export async function extendLoan(loanId, addAmount, createdBy) {
  const old = (await q("SELECT * FROM employee_loans WHERE id = ?", [loanId]))[0];
  if (!old) throw Object.assign(new Error("Loan not found."), { status: 404 });
  if (old.status !== "active") throw Object.assign(new Error("Loan is not active."), { status: 400 });

  const newTotal   = Number(old.remaining_balance) + Number(addAmount);
  const conn       = await pool.getConnection();
  try {
    await conn.beginTransaction();

    const monthly = Number(old.monthly_amount);

    const [result] = await conn.execute(
      `INSERT INTO employee_loans (employee_id, payroll_type_id, total_amount, remaining_balance, monthly_amount, start_date, created_by)
       VALUES (?,?,?,?,?,NOW(),?)`,
      [old.employee_id, old.payroll_type_id, newTotal, newTotal, monthly, createdBy]
    );
    const newLoanId = result.insertId;

    await conn.execute(
      "UPDATE employee_loans SET status='consolidated', consolidated_into=?, closed_date=NOW() WHERE id=?",
      [newLoanId, loanId]
    );

    await conn.commit();
    return { newLoanId, newTotal };
  } catch (e) {
    await conn.rollback();
    throw e;
  } finally {
    conn.release();
  }
}

export async function processLoanDeductionsForEmployee(employeeId) {
  const loans = await q(`
    SELECT el.id AS loan_id, el.remaining_balance, el.employee_id, el.payroll_type_id,
           el.monthly_amount, pt.display_name AS loan_type_name
    FROM employee_loans el
    LEFT JOIN payroll_types pt ON pt.id = el.payroll_type_id
    WHERE el.employee_id = ? AND el.status = 'active'
  `, [employeeId]);

  const halted   = [];
  const deducted = [];
  const now      = new Date();
  const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

  for (const loan of loans) {
    const remaining = Number(loan.remaining_balance);
    const monthly   = Number(loan.monthly_amount);

    if (monthly <= 0) continue; // no monthly amount set yet — skip

    if (remaining < monthly) {
      // Remaining too low — close loan without deducting
      await q("UPDATE `employee_loans` SET status='closed', closed_date=NOW() WHERE id=?", [loan.loan_id]);
      halted.push(loan);
    } else {
      const newBalance = Math.round((remaining - monthly) * 100) / 100;
      await q("UPDATE `employee_loans` SET remaining_balance=? WHERE id=?", [newBalance, loan.loan_id]);
      await q(
        `INSERT INTO \`loan_payment_history\`
           (loan_id, employee_id, amount_deducted, balance_before, balance_after, deduction_month, processed_by)
         VALUES (?, ?, ?, ?, ?, ?, 'system')`,
        [loan.loan_id, employeeId, monthly, remaining, newBalance, monthKey]
      );
      if (newBalance === 0) {
        await q("UPDATE `employee_loans` SET status='closed', closed_date=NOW() WHERE id=?", [loan.loan_id]);
      }
      deducted.push({ ...loan, newBalance });
    }
  }

  return { halted, deducted };
}

export async function processAllLoanDeductions() {
  const employees = await q(
    `SELECT DISTINCT employee_id FROM employee_loans WHERE status = 'active'`
  );
  const results = [];
  for (const { employee_id } of employees) {
    try {
      const r = await processLoanDeductionsForEmployee(employee_id);
      results.push({ employee_id, ...r });
    } catch (e) {
      results.push({ employee_id, error: e.message });
    }
  }
  return results;
}

// ── PAYROLL ARCHIVE ───────────────────────────────────────────────────────────

export async function getArchives() {
  return q(
    `SELECT id, year, month, quincena, row_count, saved_by, notes, created_at, updated_at
     FROM \`payroll_archive\`
     ORDER BY year DESC, month DESC, quincena DESC`
  );
}

export async function getArchive(year, month, quincena) {
  const rows = await q(
    "SELECT * FROM `payroll_archive` WHERE year=? AND month=? AND quincena=?",
    [Number(year), Number(month), Number(quincena)]
  );
  if (!rows[0]) return null;
  return { ...rows[0], snapshot: JSON.parse(rows[0].snapshot) };
}

export async function upsertArchive(year, month, quincena, snapshot, rowCount, savedBy, notes) {
  const existing = await q(
    "SELECT id FROM `payroll_archive` WHERE year=? AND month=? AND quincena=?",
    [Number(year), Number(month), Number(quincena)]
  );
  const snapshotStr = JSON.stringify(snapshot);
  if (existing.length > 0) {
    await q(
      "UPDATE `payroll_archive` SET snapshot=?, row_count=?, saved_by=?, notes=?, updated_at=NOW() WHERE id=?",
      [snapshotStr, rowCount, savedBy, notes ?? null, existing[0].id]
    );
    return existing[0].id;
  }
  const [result] = await q(
    "INSERT INTO `payroll_archive` (year, month, quincena, snapshot, row_count, saved_by, notes) VALUES (?,?,?,?,?,?,?)",
    [Number(year), Number(month), Number(quincena), snapshotStr, rowCount, savedBy, notes ?? null]
  );
  return result.insertId;
}

export async function savePayrollHistory(items, year, month, quincena) {
  for (const item of items) {
    const empId = item.employee_id ?? item.id;
    if (!empId) continue;
    const totalEarnings = Object.values(item.allowances ?? {}).reduce((s, v) => s + Number(v ?? 0), 0);
    await q(
      `INSERT INTO \`payroll_history\`
         (employee_id, year, month, quincena, gross_pay, total_earnings, total_deductions, net_pay, snapshot_json)
       VALUES (?,?,?,?,?,?,?,?,?)
       ON DUPLICATE KEY UPDATE
         gross_pay      = VALUES(gross_pay),
         total_earnings = VALUES(total_earnings),
         total_deductions = VALUES(total_deductions),
         net_pay        = VALUES(net_pay),
         snapshot_json  = VALUES(snapshot_json),
         saved_at       = NOW()`,
      [
        empId,
        Number(year), Number(month), Number(quincena),
        Number(item.gross_pay ?? 0),
        totalEarnings,
        Number(item.total_deductions ?? 0),
        Number(item.net_pay ?? 0),
        JSON.stringify(item),
      ]
    );
  }
}

export async function deleteArchive(id) {
  await q("DELETE FROM `payroll_archive` WHERE id=?", [id]);
}

export async function getSignatories() {
  const rows = await q("SELECT `value` FROM `app_settings` WHERE `key`='signatories' LIMIT 1");
  return rows[0] ? JSON.parse(rows[0].value) : null;
}

export async function setSignatories(data) {
  const val = JSON.stringify(data);
  await q(
    "INSERT INTO `app_settings` (`key`, `value`) VALUES ('signatories', ?) ON DUPLICATE KEY UPDATE `value`=?",
    [val, val]
  );
  return data;
}

// ── Leave Applications (CS Form No. 6) ────────────────────────────────────────

export async function ensureLeaveSchema() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS \`leave_applications\` (
      \`id\`                          CHAR(36)     PRIMARY KEY,
      \`submitted_by\`                TEXT         NOT NULL,
      \`office_department\`           TEXT         NOT NULL DEFAULT '',
      \`last_name\`                   TEXT         NOT NULL DEFAULT '',
      \`first_name\`                  TEXT         NOT NULL DEFAULT '',
      \`middle_name\`                 TEXT         NOT NULL DEFAULT '',
      \`date_of_filing\`              TEXT         NOT NULL DEFAULT '',
      \`position\`                    TEXT         NOT NULL DEFAULT '',
      \`salary\`                      TEXT         NOT NULL DEFAULT '',
      \`leave_type\`                  TEXT         NOT NULL DEFAULT '',
      \`leave_type_other\`            TEXT         NOT NULL DEFAULT '',
      \`details_vacation_local\`      TEXT         NOT NULL DEFAULT '',
      \`details_vacation_abroad\`     TEXT         NOT NULL DEFAULT '',
      \`details_sick_in_hospital\`    TEXT         NOT NULL DEFAULT '',
      \`details_sick_out_patient\`    TEXT         NOT NULL DEFAULT '',
      \`details_special_women_illness\` TEXT       NOT NULL DEFAULT '',
      \`details_study_masters\`       TINYINT(1)   NOT NULL DEFAULT 0,
      \`details_study_bar\`           TINYINT(1)   NOT NULL DEFAULT 0,
      \`details_study_other\`         TEXT         NOT NULL DEFAULT '',
      \`details_monetization\`        TINYINT(1)   NOT NULL DEFAULT 0,
      \`details_terminal\`            TINYINT(1)   NOT NULL DEFAULT 0,
      \`num_working_days\`            TEXT         NOT NULL DEFAULT '',
      \`inclusive_dates\`             TEXT         NOT NULL DEFAULT '',
      \`commutation\`                 TEXT         NOT NULL DEFAULT '',
      \`workflow_status\`             VARCHAR(20)  NOT NULL DEFAULT 'pending',
      \`applicant_signature\`         LONGTEXT,
      \`created_at\`                  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`updated_at\`                  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
      KEY \`idx_leave_submitted_by\` (\`submitted_by\`(100)),
      KEY \`idx_leave_created_at\`   (\`created_at\`),
      KEY \`idx_leave_workflow\`      (\`workflow_status\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS \`leave_action_history\` (
      \`id\`              CHAR(36)     PRIMARY KEY,
      \`leave_id\`        CHAR(36)     NOT NULL,
      \`actor_username\`  TEXT         NOT NULL,
      \`actor_role\`      TEXT         NOT NULL,
      \`actor_full_name\` TEXT         NOT NULL,
      \`action\`          TEXT         NOT NULL,
      \`details\`         JSON,
      \`created_at\`      DATETIME(6)  NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
      KEY \`idx_lah_leave_id\`   (\`leave_id\`),
      KEY \`idx_lah_created_at\` (\`created_at\`),
      CONSTRAINT \`fk_lah_leave\` FOREIGN KEY (\`leave_id\`) REFERENCES \`leave_applications\`(\`id\`) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // Pop-up notifications for the leave chain: "waiting for you" for the next
  // officer (kind 'action'), progress / final result for the applicant
  // ('info' | 'approved' | 'disapproved'). Kept until the user dismisses them,
  // so people who were offline still see them.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS \`leave_notification\` (
      \`id\`         INT          AUTO_INCREMENT PRIMARY KEY,
      \`user_id\`    VARCHAR(36)  NOT NULL,
      \`leave_id\`   CHAR(36)     NOT NULL,
      \`kind\`       VARCHAR(16)  NOT NULL,
      \`message\`    VARCHAR(400) NOT NULL,
      \`created_at\` DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`read_at\`    DATETIME     DEFAULT NULL,
      KEY \`idx_ln_user\` (\`user_id\`, \`read_at\`),
      KEY \`idx_ln_leave\` (\`leave_id\`),
      CONSTRAINT \`fk_ln_leave\` FOREIGN KEY (\`leave_id\`) REFERENCES \`leave_applications\`(\`id\`) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);

  // The applicant's division when they filed. It decides which chief / director
  // receives the application (see eligibleLeaveSignatories).
  await addColIfMissing("leave_applications", "applicant_division_id", "INT UNSIGNED DEFAULT NULL");
  // The applicant's employee record, so their leave history follows the person
  // (not just the login) and nobody can review a leave that is their own.
  await addColIfMissing("leave_applications", "employee_id", "INT DEFAULT NULL");
  await pool.query("CREATE INDEX IF NOT EXISTS `idx_leave_employee` ON `leave_applications` (`employee_id`)");
  const [unlinked] = await pool.query(
    "SELECT id, submitted_by, employee_id, applicant_division_id FROM `leave_applications` WHERE applicant_division_id IS NULL OR employee_id IS NULL"
  );
  for (const r of unlinked) {
    const [[d]] = await pool.execute(
      `SELECT e.id AS employee_id, COALESCE(e.division_id, p.division_id) AS division_id
       FROM \`users\` u JOIN \`employee\` e ON e.id = u.employee_id
       LEFT JOIN \`position\` p ON p.id = e.position_id
       WHERE LOWER(u.username) = LOWER(?) LIMIT 1`, [r.submitted_by]
    );
    if (!d) continue;
    await pool.execute(
      "UPDATE `leave_applications` SET employee_id = COALESCE(employee_id, ?), applicant_division_id = COALESCE(applicant_division_id, ?) WHERE id=?",
      [d.employee_id, d.division_id, r.id]
    );
  }

  // Names printed under the 7.A / 7.B / Director signature lines. A role may
  // have several (e.g. one chief per division); the division decides who gets
  // each application. They list in the order added (id).
  await pool.query(`
    CREATE TABLE IF NOT EXISTS \`leave_signatory\` (
      \`id\`          INT          AUTO_INCREMENT PRIMARY KEY,
      \`role\`        VARCHAR(16)  NOT NULL,
      \`name\`        VARCHAR(150) NOT NULL,
      \`designation\` VARCHAR(200) NOT NULL,
      \`created_at\`  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
      \`updated_at\`  DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
      KEY \`idx_ls_role\` (\`role\`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  `);
  // Columns that turned out redundant (2026-10-12): leave_signatory.sort_order
  // (order added = id), leave_applications.status (duplicated workflow_status),
  // leave_action_history.notes (never filled). Dropped where an older database has them.
  await pool.query("ALTER TABLE `leave_signatory` DROP COLUMN IF EXISTS `sort_order`");
  await pool.query("ALTER TABLE `leave_applications` DROP COLUMN IF EXISTS `status`");
  await pool.query("ALTER TABLE `leave_action_history` DROP COLUMN IF EXISTS `notes`");
  // division_id NULL = all divisions. Applications from a division go to the
  // approvers whose own division matches (see leaveSignatoriesFor).
  await addColIfMissing("leave_signatory", "division_id", "INT UNSIGNED DEFAULT NULL");
  const [[{ n }]] = await pool.query("SELECT COUNT(*) AS n FROM `leave_signatory`");
  if (Number(n) === 0) {
    for (const s of LEAVE_SIGNATORY_SEED) {
      await pool.execute("INSERT INTO `leave_signatory` (role, name, designation) VALUES (?,?,?)", [s.role, s.name, s.designation]);
    }
  }
  // Forms signed before the names were editable printed the seed names; record
  // them on those actions so later edits don't rewrite a signed form.
  for (const s of LEAVE_SIGNATORY_SEED) {
    const inList = LEAVE_SIGN_ACTIONS[s.role].map(a => `'${a}'`).join(",");
    await pool.execute(
      `UPDATE \`leave_action_history\`
         SET details = JSON_SET(COALESCE(details,'{}'), '$.${s.role}_name', ?, '$.${s.role}_designation', ?)
       WHERE action IN (${inList}) AND JSON_EXTRACT(COALESCE(details,'{}'), '$.${s.role}_name') IS NULL`,
      [s.name, s.designation]
    );
  }
}

const LEAVE_SIGNATORY_SEED = [
  { role: "hr",       name: "MARIA L. SANTOS",       designation: "ADMINISTRATIVE OFFICER IV /HR" },
  { role: "crasd",    name: "ENGR. JOSE R. REYES",    designation: "CHIEF ADMINISTRATIVE OFFICER/CRASD" },
  { role: "director", name: "ENGR. ANTONIO B. CRUZ", designation: "REGIONAL DIRECTOR, RSSO II" },
];
// The actions that put a signature on each section.
const LEAVE_SIGN_ACTIONS = {
  hr:       ["hr_approved"],
  crasd:    ["crasd_approved"],
  director: ["director_approved","director_disapproved"],
};
export const LEAVE_SIGNATORY_ROLES = Object.keys(LEAVE_SIGN_ACTIONS);

const SIGNATORY_SELECT = `
  SELECT s.id, s.role, s.name, s.designation,
         s.division_id, d.name AS division_name
  FROM \`leave_signatory\` s
  LEFT JOIN \`division\` d ON d.id = s.division_id`;

export async function listLeaveSignatories() {
  const [rows] = await pool.query(`${SIGNATORY_SELECT} ORDER BY s.role, s.id`);
  return rows;
}

export async function getLeaveSignatory(id) {
  const [[row]] = await pool.execute(`${SIGNATORY_SELECT} WHERE s.id=? LIMIT 1`, [id]);
  return row ?? null;
}

export async function createLeaveSignatory(role, { name, designation, division_id }) {
  const [r] = await pool.execute(
    "INSERT INTO `leave_signatory` (role, name, designation, division_id) VALUES (?,?,?,?)",
    [role, name, designation, division_id]
  );
  return getLeaveSignatory(r.insertId);
}

export async function updateLeaveSignatory(id, { name, designation, division_id }) {
  await pool.execute(
    "UPDATE `leave_signatory` SET name=?, designation=?, division_id=? WHERE id=?",
    [name, designation, division_id, id]
  );
  return getLeaveSignatory(id);
}

// Who receives an application at a section: the names set to the applicant's
// division; if none, the "All divisions" names; if none of those either,
// everyone in the section, so a form is never left with no one to sign it.
// The frontend mirrors this in LeaveForm.jsx.
export function eligibleLeaveSignatories(all, role, divisionId) {
  const ofRole = all.filter(s => s.role === role);
  const exact = divisionId != null ? ofRole.filter(s => s.division_id != null && Number(s.division_id) === Number(divisionId)) : [];
  if (exact.length) return exact;
  const general = ofRole.filter(s => s.division_id == null);
  return general.length ? general : ofRole;
}

// A name belongs to an approver when it is set to their own division (from
// their employee record) or to "All divisions".
export const isOwnLeaveSignatory = (s, actorDivisionId) =>
  s.division_id == null || (actorDivisionId != null && Number(s.division_id) === Number(actorDivisionId));

// The names this approver may sign as on this application: the eligible ones
// that belong to them.
export function leaveSignatoriesFor(actor, record, role, all) {
  return eligibleLeaveSignatories(all, role, record.applicant_division_id)
    .filter(s => isOwnLeaveSignatory(s, actor.divisionId));
}

// The applicant: the account that filed it, or any account linked to the same
// employee record (so a second login can't approve one's own leave).
export function isLeaveOwner(actor, record) {
  if (record.submitted_by === actor.username) return true;
  return record.employee_id != null && actor.employeeId != null
    && Number(record.employee_id) === Number(actor.employeeId);
}

// An approver receives an application for a section when they hold that
// section's permission, it isn't their own leave, and at least one eligible
// name belongs to them, so a chief never gets another division's applications.
// With no names set up at all, the permission alone decides.
export function isLeaveRecipient(actor, record, role, all) {
  if (!actor.perms.includes(`leave.${role}`)) return false;
  if (isLeaveOwner(actor, record)) return false;
  if (!all.some(s => s.role === role)) return true;
  return leaveSignatoriesFor(actor, record, role, all).length > 0;
}

// The chain: each approver sees an application only at their own step — HR
// while pending, the Chief once HR certifies, the Director once the Chief
// recommends. After they act it moves on and leaves their list, so it stays
// private to the applicant (and leave.manage oversight).
const LEAVE_STAGE = { hr: "pending", crasd: "hr_approved", director: "crasd_approved" };

export function canSeeLeaveRecord(actor, record, all) {
  if (actor.isAdmin) return true;
  if (isLeaveOwner(actor, record)) return true;
  return LEAVE_SIGNATORY_ROLES.some(role =>
    record.workflow_status === LEAVE_STAGE[role] && isLeaveRecipient(actor, record, role, all));
}

// Refuses to remove the last signatory of a role, since that section could no
// longer be signed. Returns "last" in that case.
export async function deleteLeaveSignatory(id) {
  const row = await getLeaveSignatory(id);
  if (!row) return null;
  const [[{ n }]] = await pool.execute("SELECT COUNT(*) AS n FROM `leave_signatory` WHERE role=?", [row.role]);
  if (Number(n) <= 1) return "last";
  await pool.execute("DELETE FROM `leave_signatory` WHERE id=?", [id]);
  return row;
}

// Text fields that used to be flat columns — coerce NULL (no action yet) back to '' to match API contract.
const LEAVE_REVIEW_TEXT_FIELDS = [
  "credits_as_of","vl_total_earned","vl_less_application","vl_balance",
  "sl_total_earned","sl_less_application","sl_balance",
  "recommendation","disapproval_reason",
  "approved_days_with_pay","approved_days_without_pay","approved_others_specify","disapproved_due_to",
  "hr_name","hr_designation","crasd_name","crasd_designation","director_name","director_designation",
];

function reconstructLeave(row) {
  if (!row) return row;
  for (const f of LEAVE_REVIEW_TEXT_FIELDS) if (row[f] == null) row[f] = "";
  return row;
}

// Action families: each approver role owns one family of history rows.
const LEAVE_ACTION_FAMILIES = {
  hr:       ["hr_approved","hr_correction"],
  crasd:    ["crasd_approved","crasd_correction"],
  director: ["director_approved","director_disapproved","director_correction"],
};

function latestFamilyJoin(alias, actions) {
  const inList = actions.map(a => `'${a}'`).join(",");
  return `(
    SELECT leave_id, details FROM (
      SELECT leave_id, details,
             ROW_NUMBER() OVER (PARTITION BY leave_id ORDER BY created_at DESC) AS rn
      FROM leave_action_history WHERE action IN (${inList})
    ) t WHERE rn = 1
  ) ${alias}`;
}

const REVIEW_SELECT = `
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.credits_as_of'))           AS credits_as_of,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.vl_total_earned'))         AS vl_total_earned,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.vl_less_application'))     AS vl_less_application,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.vl_balance'))              AS vl_balance,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.sl_total_earned'))         AS sl_total_earned,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.sl_less_application'))     AS sl_less_application,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.sl_balance'))              AS sl_balance,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.hr_signature'))            AS hr_signature,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.hr_name'))                 AS hr_name,
  JSON_UNQUOTE(JSON_EXTRACT(hr.details, '$.hr_designation'))          AS hr_designation,
  JSON_UNQUOTE(JSON_EXTRACT(crasd.details, '$.recommendation'))       AS recommendation,
  JSON_UNQUOTE(JSON_EXTRACT(crasd.details, '$.disapproval_reason'))   AS disapproval_reason,
  JSON_UNQUOTE(JSON_EXTRACT(crasd.details, '$.crasd_signature'))      AS crasd_signature,
  JSON_UNQUOTE(JSON_EXTRACT(crasd.details, '$.crasd_name'))           AS crasd_name,
  JSON_UNQUOTE(JSON_EXTRACT(crasd.details, '$.crasd_designation'))    AS crasd_designation,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.approved_days_with_pay'))    AS approved_days_with_pay,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.approved_days_without_pay')) AS approved_days_without_pay,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.approved_others_specify'))   AS approved_others_specify,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.disapproved_due_to'))        AS disapproved_due_to,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.director_signature'))        AS director_signature,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.director_name'))             AS director_name,
  JSON_UNQUOTE(JSON_EXTRACT(director.details, '$.director_designation'))      AS director_designation
`;
const REVIEW_JOIN = `
  LEFT JOIN ${latestFamilyJoin("hr",       LEAVE_ACTION_FAMILIES.hr)}       ON hr.leave_id       = la.id
  LEFT JOIN ${latestFamilyJoin("crasd",    LEAVE_ACTION_FAMILIES.crasd)}    ON crasd.leave_id    = la.id
  LEFT JOIN ${latestFamilyJoin("director", LEAVE_ACTION_FAMILIES.director)} ON director.leave_id = la.id
`;

export async function createLeaveApplication(data, actorUsername, actorRole, actorFullName) {
  const id = randomUUID();
  await pool.execute(
    `INSERT INTO \`leave_applications\` (
      id, submitted_by, office_department, last_name, first_name, middle_name,
      date_of_filing, position, salary, leave_type, leave_type_other,
      details_vacation_local, details_vacation_abroad, details_sick_in_hospital,
      details_sick_out_patient, details_special_women_illness,
      details_study_masters, details_study_bar, details_study_other,
      details_monetization, details_terminal, num_working_days, inclusive_dates, commutation,
      applicant_signature, applicant_division_id, employee_id
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [
      id, data.submitted_by, data.office_department ?? "", data.last_name ?? "",
      data.first_name ?? "", data.middle_name ?? "", data.date_of_filing ?? "",
      data.position ?? "", data.salary ?? "", data.leave_type ?? "",
      data.leave_type_other ?? "", data.details_vacation_local ?? "",
      data.details_vacation_abroad ?? "", data.details_sick_in_hospital ?? "",
      data.details_sick_out_patient ?? "", data.details_special_women_illness ?? "",
      data.details_study_masters ? 1 : 0, data.details_study_bar ? 1 : 0,
      data.details_study_other ?? "", data.details_monetization ? 1 : 0,
      data.details_terminal ? 1 : 0, data.num_working_days ?? "",
      data.inclusive_dates ?? "", data.commutation ?? "",
      data.applicant_signature ?? null,
      data.applicant_division_id ?? null,
      data.employee_id ?? null,
    ]
  );
  await pool.execute(
    `INSERT INTO \`leave_action_history\` (id, leave_id, actor_username, actor_role, actor_full_name, action)
     VALUES (?,?,?,?,?,'submitted')`,
    [randomUUID(), id, actorUsername, actorRole, actorFullName ?? ""]
  );
  const [[row]] = await pool.execute("SELECT * FROM `leave_applications` WHERE id=?", [id]);
  return row;
}

export async function getLeaveApplication(id) {
  const [[row]] = await pool.execute(
    `SELECT la.*, ${REVIEW_SELECT} FROM \`leave_applications\` la ${REVIEW_JOIN} WHERE la.id=? LIMIT 1`,
    [id]
  );
  return reconstructLeave(row) ?? null;
}

// ── Leave notifications ──────────────────────────────────────────────────────

// Active accounts whose role holds a leave permission, shaped like a leave actor
// so isLeaveRecipient can pick the ones an application is routed to.
export async function listLeaveApprovers(permKey) {
  const [rows] = await pool.execute(
    `SELECT u.id, LOWER(u.username) AS username, e.id AS employee_id,
            COALESCE(e.division_id, p.division_id) AS division_id
     FROM \`users\` u
     JOIN \`role_permission\` rp ON rp.role_id = u.role_id
     JOIN \`permission\` pm      ON pm.id = rp.permission_id AND pm.\`key\` = ?
     LEFT JOIN \`employee\` e    ON e.id = u.employee_id
     LEFT JOIN \`position\` p    ON p.id = e.position_id
     WHERE u.status = 'active'`, [permKey]
  );
  return rows.map(r => ({ id: r.id, username: r.username, employeeId: r.employee_id, divisionId: r.division_id, perms: [permKey] }));
}

// The applicant's accounts: the login that filed it and any linked to the same employee.
export async function listLeaveOwnerAccounts(record) {
  const [rows] = await pool.execute(
    `SELECT id FROM \`users\`
     WHERE status = 'active' AND (LOWER(username) = LOWER(?) OR (employee_id IS NOT NULL AND employee_id = ?))`,
    [record.submitted_by, record.employee_id ?? -1]
  );
  return rows.map(r => r.id);
}

export async function addLeaveNotifications(userIds, leaveId, kind, message) {
  for (const uid of new Set(userIds)) {
    await pool.execute(
      "INSERT INTO `leave_notification` (user_id, leave_id, kind, message) VALUES (?,?,?,?)",
      [uid, leaveId, kind, message.slice(0, 400)]
    );
  }
}

// Someone acted on this application, so earlier "waiting for you" pop-ups are stale.
export async function clearLeaveActionNotifications(leaveId) {
  await pool.execute(
    "UPDATE `leave_notification` SET read_at = NOW() WHERE leave_id = ? AND kind = 'action' AND read_at IS NULL",
    [leaveId]
  );
}

export async function listLeaveNotifications(userId) {
  const [rows] = await pool.execute(
    `SELECT id, leave_id, kind, message, created_at FROM \`leave_notification\`
     WHERE user_id = ? AND read_at IS NULL ORDER BY created_at DESC, id DESC LIMIT 20`, [userId]
  );
  return rows;
}

export async function markLeaveNotificationsRead(userId, id = null) {
  if (id == null) {
    await pool.execute("UPDATE `leave_notification` SET read_at = NOW() WHERE user_id = ? AND read_at IS NULL", [userId]);
  } else {
    await pool.execute("UPDATE `leave_notification` SET read_at = NOW() WHERE user_id = ? AND id = ? AND read_at IS NULL", [userId, id]);
  }
}

// `actor` = { username, employeeId, divisionId, isAdmin, perms } (see leaveActor in server).
// Each row gets `mine` = it is the caller's own leave.
export async function listLeaveApplicationsByUser(actor) {
  const base = `SELECT la.*, ${REVIEW_SELECT} FROM \`leave_applications\` la ${REVIEW_JOIN}`;
  const tag = (rows) => rows.map(reconstructLeave).map(r => ({ ...r, mine: isLeaveOwner(actor, r) }));

  if (actor.isAdmin) {
    const [rows] = await pool.execute(`${base} ORDER BY la.created_at DESC LIMIT 500`);
    return tag(rows);
  }
  const isApprover = LEAVE_SIGNATORY_ROLES.some(role => actor.perms.includes(`leave.${role}`));
  if (!isApprover) {
    // employee — own applications only (by login or by employee record)
    const [rows] = await pool.execute(
      `${base} WHERE LOWER(la.submitted_by)=LOWER(?) OR (la.employee_id IS NOT NULL AND la.employee_id = ?)
       ORDER BY la.created_at DESC LIMIT 200`,
      [actor.username, actor.employeeId ?? -1]
    );
    return tag(rows);
  }
  // Approvers: their own, plus what is waiting at their step of the chain.
  const [rows] = await pool.execute(`${base} ORDER BY la.created_at DESC LIMIT 500`);
  const sigs = await listLeaveSignatories();
  return tag(rows).filter(r => canSeeLeaveRecord(actor, r, sigs));
}

const LEAVE_FIELD_FAMILY = {
  credits_as_of:"hr", vl_total_earned:"hr", vl_less_application:"hr", vl_balance:"hr",
  sl_total_earned:"hr", sl_less_application:"hr", sl_balance:"hr",
  recommendation:"crasd", disapproval_reason:"crasd",
  approved_days_with_pay:"director", approved_days_without_pay:"director",
  approved_others_specify:"director", disapproved_due_to:"director",
};

export async function updateLeaveSection7(id, updates, actor) {
  const byFamily = { hr:{}, crasd:{}, director:{} };
  for (const [k, v] of Object.entries(updates)) {
    const fam = LEAVE_FIELD_FAMILY[k];
    if (fam) byFamily[fam][k] = v;
  }
  const familyEntries = Object.entries(byFamily).filter(([, fields]) => Object.keys(fields).length > 0);
  if (familyEntries.length === 0) return null;

  const conn = await pool.getConnection();
  try {
    await conn.beginTransaction();
    const [[exists]] = await conn.execute("SELECT id FROM `leave_applications` WHERE id=? LIMIT 1 FOR UPDATE", [id]);
    if (!exists) { await conn.rollback(); return null; }

    for (const [family, fields] of familyEntries) {
      const actions = LEAVE_ACTION_FAMILIES[family];
      const inList  = actions.map(a => `'${a}'`).join(",");
      const [[latest]] = await conn.execute(
        `SELECT id FROM \`leave_action_history\` WHERE leave_id=? AND action IN (${inList}) ORDER BY created_at DESC LIMIT 1 FOR UPDATE`,
        [id]
      );
      if (latest) {
        await conn.execute(
          "UPDATE `leave_action_history` SET details=JSON_MERGE_PATCH(COALESCE(details,'{}'),?) WHERE id=?",
          [JSON.stringify(fields), latest.id]
        );
      } else if (Object.values(fields).some(v => String(v ?? "").trim() !== "")) {
        // Only record a correction for a section that hasn't been reached yet
        // when something was actually entered; blank sections stay untouched.
        await conn.execute(
          `INSERT INTO \`leave_action_history\` (id, leave_id, actor_username, actor_role, actor_full_name, action, details)
           VALUES (?,?,?,?,?,?,?)`,
          [randomUUID(), id, actor.username, actor.role, actor.full_name ?? "", `${family}_correction`, JSON.stringify(fields)]
        );
      }
    }
    const [[row]] = await conn.execute(
      `SELECT la.*, ${REVIEW_SELECT} FROM \`leave_applications\` la ${REVIEW_JOIN} WHERE la.id=?`, [id]
    );
    await conn.commit();
    return reconstructLeave(row) ?? null;
  } catch (e) { await conn.rollback(); throw e; }
  finally { conn.release(); }
}

// `signatory` is the { name, designation } printed under the signature; it is
// copied into the action so later edits to the signatory list don't change it.
export async function processLeaveAction(leaveId, actor, action, updates, signature, signatory) {
  const conn = await pool.getConnection();
  try {
    await conn.beginTransaction();
    const workflowMap = {
      hr_approved:          { workflow_status: "hr_approved" },
      crasd_approved:       { workflow_status: "crasd_approved" },
      director_approved:    { workflow_status: "approved" },
      director_disapproved: { workflow_status: "disapproved" },
    };
    const wf = workflowMap[action];
    if (!wf) throw Object.assign(new Error("Invalid action."), { status: 400 });

    const [result] = await conn.execute(
      "UPDATE `leave_applications` SET workflow_status = ? WHERE id=?", [wf.workflow_status, leaveId]
    );
    if (result.affectedRows === 0) throw Object.assign(new Error("Not found."), { status: 404 });

    const details = { ...updates };
    const role = action.startsWith("hr") ? "hr" : action.startsWith("crasd") ? "crasd" : "director";
    if (signature !== undefined) details[`${role}_signature`] = signature;
    if (signatory) {
      details[`${role}_name`] = signatory.name;
      details[`${role}_designation`] = signatory.designation;
    }
    await conn.execute(
      `INSERT INTO \`leave_action_history\` (id, leave_id, actor_username, actor_role, actor_full_name, action, details)
       VALUES (?,?,?,?,?,?,?)`,
      [randomUUID(), leaveId, actor.username, actor.role, actor.full_name ?? "", action, JSON.stringify(details)]
    );
    const [[row]] = await conn.execute(
      `SELECT la.*, ${REVIEW_SELECT} FROM \`leave_applications\` la ${REVIEW_JOIN} WHERE la.id=?`, [leaveId]
    );
    await conn.commit();
    return reconstructLeave(row);
  } catch (e) { await conn.rollback(); throw e; }
  finally { conn.release(); }
}

export async function getLeaveHistory(leaveId) {
  const [rows] = await pool.execute(
    "SELECT * FROM `leave_action_history` WHERE leave_id=? ORDER BY created_at ASC",
    [leaveId]
  );
  return rows.map(r => ({ ...r, details: r.details ? (typeof r.details === "string" ? JSON.parse(r.details) : r.details) : null }));
}

// Identity block of CS Form 6, snapshotted from the user's linked employee at filing
// time so a printed form keeps what was true then. employee.name is "LAST, FIRST M.".
export async function getLeaveSnapshot(userId) {
  const [[row]] = await pool.execute(
    `SELECT e.id AS employee_id, e.name, e.basic_monthly_pay, p.position_title, d.id AS division_id, d.name AS division_name
     FROM \`users\` u
     JOIN \`employee\` e      ON e.id = u.employee_id
     LEFT JOIN \`position\` p ON p.id = e.position_id
     LEFT JOIN \`division\` d ON d.id = COALESCE(e.division_id, p.division_id)
     WHERE u.id = ? LIMIT 1`,
    [userId]
  );
  if (!row) return null;
  const [last = "", rest = ""] = String(row.name).split(",").map(s => s.trim());
  const parts  = rest.split(/\s+/).filter(Boolean);
  const middle = parts.length > 1 && /^[A-Z]\.?$/i.test(parts.at(-1)) ? parts.pop() : "";
  return {
    office_department: row.division_name ? `RSSO II - ${row.division_name}` : "RSSO II",
    division_id: row.division_id ?? null,
    employee_id: row.employee_id,
    last_name:   last,
    first_name:  parts.join(" "),
    middle_name: middle,
    position:    row.position_title ?? "",
    salary:      `₱${Number(row.basic_monthly_pay ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
  };
}


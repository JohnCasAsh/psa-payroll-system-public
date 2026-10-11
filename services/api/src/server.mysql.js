import cors from "cors";
import compression from "compression";
import express from "express";
import cron from "node-cron";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import multer from "multer";
import { randomUUID } from "crypto";
import pinoHttp from "pino-http";
import cookieParser from "cookie-parser";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import logger from "./logger.js";
import {
  ensureSchema,
  pool,
  getPendingUsers,
  approveUser,
  rejectUser,
  getAllManagedUsers,
  blockUser,
  unblockUser,
  deleteSystemUser,
  searchEmployeesPage,
  addEmployee,
  updateEmployee,
  softDeleteEmployee,
  getPayrollStats,
  getEmployeeDeductions,
  upsertEmployeeDeduction,
  getEmployeeLedger,
  getAuditLog,
  getInactiveEmployees, getEmployeeProfile,
  activateEmployee,
  getLivePayslips,
  getPositions, createPosition, updatePosition, deletePosition, hitPosition,
  getDivisions, createDivision, updateDivision, deleteDivision,
  getDeductionTypes, createDeductionType, updateDeductionType, softDeleteDeductionType,
  getEarningTypes, createEarningType, updateEarningType, softDeleteEarningType,
  getArchives, getArchive, upsertArchive, deleteArchive, savePayrollHistory,
  syncComputedDeductionsToLedger,
  getEmployeeLoans, getLoanPaymentHistory, createLoan, consolidateLoan, closeLoan, extendLoan, processLoanDeductionsForEmployee, processAllLoanDeductions,
  getSignatories, setSignatories,
  getEmployeeById,
  insertAuditLog,
  list201Files, get201FileData, insert201File, delete201File,
  ensureLeaveSchema, createLeaveApplication, getLeaveApplication, listLeaveApplicationsByUser,
  updateLeaveSection7, processLeaveAction, getLeaveHistory, getLeaveSnapshot,
  LEAVE_SIGNATORY_ROLES, listLeaveSignatories, getLeaveSignatory,
  createLeaveSignatory, updateLeaveSignatory, deleteLeaveSignatory,
  leaveSignatoriesFor, isLeaveRecipient, canSeeLeaveRecord, isOwnLeaveSignatory, isLeaveOwner,
  listLeaveApprovers, listLeaveOwnerAccounts, addLeaveNotifications, clearLeaveActionNotifications,
  listLeaveNotifications, markLeaveNotificationsRead,
} from "./db.mysql.js";
import { enqueueJob, redis } from "./queue.js";
import { ensureBucket, uploadObject, deleteObject } from "./storage.js";

const app  = express();
const port = Number(process.env.PORT ?? 3000);
const IS_PROD = process.env.NODE_ENV === "production";
// The Firebase-hosted site calls the API on another domain (through the HTTPS
// tunnel), which needs a cross-site cookie: SameSite=None, and that requires
// Secure. Plain-HTTP local / LAN use keeps a Lax cookie, which Secure would break.
// req.secure reads X-Forwarded-Proto from the tunnel ("trust proxy" below).
const cookieOpts = (req) => (IS_PROD || req.secure)
  ? { httpOnly: true, sameSite: "none", secure: true }
  : { httpOnly: true, sameSite: "lax",  secure: false };
app.set("trust proxy", 1);
app.disable("x-powered-by");

app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'none'"],
      frameAncestors: ["'none'"],
    },
  },
  hsts: { maxAge: 15552000, includeSubDomains: true },
  crossOriginResourcePolicy: { policy: "cross-origin" },
}));

const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET) throw new Error("JWT_SECRET must be set in environment.");

const _sentinelHash = await bcrypt.hash("__sentinel__", 12);

if (!process.env.ADMIN_USERNAME) throw new Error("ADMIN_USERNAME must be set in environment.");
const ADMIN_USERNAME = process.env.ADMIN_USERNAME.toLowerCase();

app.use(pinoHttp({
  logger,
  genReqId: () => randomUUID(),
  customLogLevel: (_req, res) => res.statusCode >= 500 ? "error" : res.statusCode >= 400 ? "warn" : "info",
  serializers: {
    req: (req) => ({ method: req.method, url: req.url, id: req.id }),
    res: (res) => ({ statusCode: res.statusCode })
  }
}));

app.use(compression());

// CORS_ORIGIN may list several origins, comma-separated (e.g. the Firebase site).
const CORS_ORIGINS = (process.env.CORS_ORIGIN ?? "http://localhost:5173").split(",").map(s => s.trim()).filter(Boolean);
if (process.env.NODE_ENV === "production" && CORS_ORIGINS.every(o => o.startsWith("http://localhost"))) {
  throw new Error("CORS_ORIGIN must not be localhost in production.");
}
const ALLOWED_ORIGINS = new Set([...CORS_ORIGINS, "http://localhost:5173", "http://localhost:5174", "http://localhost:3000"]);
app.use(cors({
  origin: (origin, cb) => {
    // Allow requests with no origin (Electron file://, curl, Postman)
    if (!origin) return cb(null, true);
    if (ALLOWED_ORIGINS.has(origin)) return cb(null, true);
    // Allow any device on the local network (192.168.x.x, 10.x.x.x, 172.16-31.x.x)
    try {
      const host = new URL(origin).hostname;
      if (/^192\.168\.\d+\.\d+$/.test(host) || /^10\.\d+\.\d+\.\d+$/.test(host) || /^172\.(1[6-9]|2\d|3[01])\.\d+\.\d+$/.test(host)) return cb(null, true);
    } catch {}
    cb(new Error("Not allowed by CORS"));
  },
  methods: ["GET", "POST", "PUT", "PATCH", "DELETE"],
  allowedHeaders: ["Content-Type", "Authorization"],
  credentials: true
}));

// CSRF defense: reject state-changing requests from unknown origins.
// Requests with no origin (Electron desktop app, curl) are allowed.
app.use((req, res, next) => {
  if (["GET", "HEAD", "OPTIONS"].includes(req.method)) return next();
  let origin = req.headers.origin;
  if (!origin && req.headers.referer) {
    try { origin = new URL(req.headers.referer).origin; } catch { /* leave undefined */ }
  }
  if (origin && !ALLOWED_ORIGINS.has(origin)) {
    try {
      const host = new URL(origin).hostname;
      if (/^192\.168\.\d+\.\d+$/.test(host) || /^10\.\d+\.\d+\.\d+$/.test(host) || /^172\.(1[6-9]|2\d|3[01])\.\d+\.\d+$/.test(host)) { next(); return; }
    } catch {}
    return res.status(403).json({ error: "Cross-origin request blocked." });
  }
  next();
});

// Leave applications carry base64 signature images, so they get a larger body limit.
const jsonSmall = express.json({ limit: "50kb" });
const jsonLeave = express.json({ limit: "2mb" });
app.use((req, res, next) => (req.path.startsWith("/leave-applications") ? jsonLeave : jsonSmall)(req, res, next));
app.use(cookieParser());

function stripTags(str) {
  let prev, cur = str;
  do { prev = cur; cur = prev.replace(/<[^>]*>?/g, ""); } while (cur !== prev);
  return cur.trim();
}
function sanitizePayload(obj, depth = 0) {
  if (depth > 10) throw new Error("Payload nesting too deep");
  if (typeof obj !== "object" || obj === null) return obj;
  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      if (typeof obj[i] === "string") obj[i] = stripTags(obj[i]);
      else if (typeof obj[i] === "object") sanitizePayload(obj[i], depth + 1);
    }
    return obj;
  }
  for (const key of Object.keys(obj)) {
    if (typeof obj[key] === "string") obj[key] = stripTags(obj[key]);
    else if (typeof obj[key] === "object") sanitizePayload(obj[key], depth + 1);
  }
  return obj;
}
app.use((req, res, next) => {
  if (req.body && !req.path.startsWith("/auth")) req.body = sanitizePayload(req.body);
  next();
});

function redisRateLimitStore(prefix, windowMs) {
  return {
    init() {},
    async increment(key) {
      try {
        const k = `${prefix}${key}`;
        const [count] = await redis.multi().incr(k).pExpire(k, windowMs, "NX").exec();
        return { totalHits: count, resetTime: new Date(Date.now() + windowMs) };
      } catch { return { totalHits: 1, resetTime: new Date(Date.now() + windowMs) }; }
    },
    async decrement(key) { await redis.decr(`${prefix}${key}`).catch(() => {}); },
    async resetKey(key) { await redis.del(`${prefix}${key}`).catch(() => {}); },
  };
}

const globalLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 500,
  standardHeaders: true,
  legacyHeaders: false,
  skip: (req) => req.path === "/events",
  keyGenerator: (req) => req.user?.username ?? req.ip,
  message: { error: "Too many requests. Please wait 15 minutes." },
  store: redisRateLimitStore("rl:global:", 15 * 60 * 1000),
});
app.use(globalLimiter);

const authLimiter = rateLimit({
  windowMs: 60 * 60 * 1000,
  max: 5,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many login attempts. Please wait 1 hour." },
  store: redisRateLimitStore("rl:auth:", 60 * 60 * 1000),
});

const avatarUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 2 * 1024 * 1024 },
  fileFilter(_req, file, cb) {
    if (["image/jpeg", "image/png", "image/webp", "image/gif"].includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(Object.assign(new Error("Only JPEG, PNG, WebP, or GIF images are accepted."), { status: 400 }));
    }
  }
});

function detectImageMime(buf) {
  if (buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF) return "image/jpeg";
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) return "image/png";
  if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x38) return "image/gif";
  if (buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46 &&
      buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50) return "image/webp";
  return null;
}

const UUID_RE  = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const EMP_ID_RE = /^\d{7}$/;

// basic_monthly_pay is NUMERIC(12,2) — 10 digits before the decimal point.
const MAX_BASIC_PAY = 9999999999.99;
function parseBasicPay(value) {
  if (value === undefined) return { ok: true, value: undefined };
  const n = Number(value);
  if (!Number.isFinite(n)) return { ok: false, error: "basicMonthlyPay must be a valid number." };
  if (n < 0)              return { ok: false, error: "basicMonthlyPay cannot be negative." };
  if (n > MAX_BASIC_PAY)  return { ok: false, error: "basicMonthlyPay exceeds the maximum allowed value." };
  return { ok: true, value: n };
}

async function resolvePosition(positionId) {
  if (!positionId) return { ok: true, name: undefined };
  const [rows] = await pool.execute("SELECT position_title AS name FROM `position` WHERE id = ? LIMIT 1", [positionId]);
  if (!rows[0]) return { ok: false, error: "Invalid position ID." };
  return { ok: true, name: rows[0].name };
}

// ── SSE ───────────────────────────────────────────────────────────────────────
const sseClients = new Set();
const sseUserCounts = new Map();
const SSE_MAX_PER_USER = 5;

function broadcast(event, data = {}) {
  if (sseClients.size === 0) return;
  const payload = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  for (const client of sseClients) {
    try { client.write(payload); client.flush?.(); } catch { sseClients.delete(client); }
  }
}

// ── AUTH MIDDLEWARE ───────────────────────────────────────────────────────────
function isLocalNetwork(ip) {
  if (!ip) return false;
  const addr = ip.replace(/^::ffff:/, "");
  return addr === "127.0.0.1" || addr === "::1" ||
    /^192\.168\.\d+\.\d+$/.test(addr) ||
    /^10\.\d+\.\d+\.\d+$/.test(addr) ||
    /^172\.(1[6-9]|2\d|3[01])\.\d+\.\d+$/.test(addr);
}

async function requireAuth(req, res, next) {
  const token = req.cookies?.auth_token;
  if (!token) return res.status(401).json({ error: "Not authenticated." });
  let payload;
  try {
    payload = jwt.verify(token, JWT_SECRET, { algorithms: ["HS256"] });
  } catch {
    res.clearCookie("auth_token", cookieOpts(req));
    return res.status(401).json({ error: "Session expired. Please sign in again." });
  }
  // DB token-version check — instant revocation without Redis
  try {
    const [[userRow]] = await pool.execute(
      "SELECT token_version, status FROM `users` WHERE id = ? LIMIT 1", [payload.sub]
    );
    if (!userRow || userRow.status === "blocked") {
      res.clearCookie("auth_token", cookieOpts(req));
      return res.status(401).json({ error: "Account has been restricted." });
    }
    if ((payload.tv ?? 0) !== userRow.token_version) {
      res.clearCookie("auth_token", cookieOpts(req));
      return res.status(401).json({ error: "Session expired. Please sign in again." });
    }
  } catch (e) {
    logger.warn({ err: e }, "requireAuth DB check failed — falling back to JWT-only");
  }
  // Redis blacklist check — skip silently if Redis is unavailable
  try {
    if (payload.jti && await redis.get(`bl:${payload.jti}`)) {
      res.clearCookie("auth_token", cookieOpts(req));
      return res.status(401).json({ error: "Session has been revoked. Please sign in again." });
    }
  } catch { /* Redis unavailable — skip blacklist check */ }
  req.user = payload;
  next();
}

function requireAdmin(req, res, next) {
  if (req.user?.username?.toLowerCase() !== ADMIN_USERNAME) {
    return res.status(403).json({ error: "Admin access required." });
  }
  next();
}

// ── RBAC HELPERS ──────────────────────────────────────────────────────────────
async function loadUserPermissions(userId) {
  const [rows] = await pool.execute(`
    SELECT p.\`key\` FROM \`permission\` p
    JOIN \`role_permission\` rp ON rp.permission_id = p.id
    JOIN \`users\` u ON u.role_id = rp.role_id
    WHERE u.id = ?
  `, [userId]);
  return rows.map(r => r.key);
}

function requirePermission(key) {
  return [requireAuth, async (req, res, next) => {
    try {
      if (req.user.username.toLowerCase() === ADMIN_USERNAME) return next();
      if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
      if (!req.userPermissions.includes(key)) return res.status(403).json({ error: "Permission denied." });
      next();
    } catch (err) { next(err); }
  }];
}

function requireAnyPermission(...keys) {
  return [requireAuth, async (req, res, next) => {
    try {
      if (req.user.username.toLowerCase() === ADMIN_USERNAME) return next();
      if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
      if (!keys.some(k => req.userPermissions.includes(k))) return res.status(403).json({ error: "Permission denied." });
      next();
    } catch (err) { next(err); }
  }];
}

// ── AUTH ROUTES ───────────────────────────────────────────────────────────────
// Self-registration creates a 'pending' account; an admin approves it in User Management.
app.post("/auth/register", authLimiter, async (req, res, next) => {
  try {
    const { username, password } = req.body ?? {};
    if (!username || typeof username !== "string" || username.trim().length < 3 || username.trim().length > 50)
      return res.status(400).json({ error: "Username must be 3–50 characters." });
    if (!/^[A-Za-z0-9._-]+$/.test(username.trim()))
      return res.status(400).json({ error: "Username may only contain letters, numbers, dots, hyphens, and underscores." });
    if (!password || typeof password !== "string" || password.length < 8)
      return res.status(400).json({ error: "Password must be at least 8 characters." });
    if (!/[A-Z]/.test(password) || !/[0-9]/.test(password) || !/[^A-Za-z0-9]/.test(password))
      return res.status(400).json({ error: "Password must include uppercase, number, and symbol." });

    // users.username has no unique index, so the duplicate check must be explicit.
    const [taken] = await pool.execute("SELECT 1 FROM `users` WHERE LOWER(username) = LOWER(?) LIMIT 1", [username.trim()]);
    if (taken.length > 0) return res.status(409).json({ error: "Username already taken." });

    const passwordHash = await bcrypt.hash(password, 12);
    await pool.execute(
      "INSERT INTO `users` (id, username, password_hash, status) VALUES (?, LOWER(?), ?, 'pending')",
      [randomUUID(), username.trim(), passwordHash]
    );
    res.status(201).json({ success: true, pending: true });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.post("/auth/login", authLimiter, async (req, res, next) => {
  try {

    const { username, password } = req.body ?? {};
    if (!username || typeof username !== "string" || !password || typeof password !== "string")
      return res.status(400).json({ error: "Username and password required." });

    const [rows] = await pool.execute(
      "SELECT id, username, password_hash, status, token_version, failed_attempts, locked_until FROM `users` WHERE LOWER(username) = LOWER(?) LIMIT 1",
      [username.trim()]
    );
    const user = rows[0];

    // Account lockout check
    if (user?.locked_until && new Date(user.locked_until) > new Date()) {
      const mins = Math.ceil((new Date(user.locked_until) - new Date()) / 60000);
      return res.status(423).json({ error: `Account locked due to too many failed attempts. Try again in ${mins} minute(s).` });
    }

    const match = await bcrypt.compare(password, user?.password_hash ?? _sentinelHash);

    if (!user || !match) {
      // Track failed attempt per account
      if (user) {
        const newCount = (user.failed_attempts ?? 0) + 1;
        const lockUntil = newCount >= 5 ? new Date(Date.now() + 30 * 60 * 1000) : null;
        await pool.execute(
          "UPDATE `users` SET failed_attempts = ?, locked_until = ? WHERE id = ?",
          [newCount, lockUntil, user.id]
        );
      }
      return res.status(401).json({ error: "Invalid username or password." });
    }
    if (user.status === "pending")  return res.status(403).json({ error: "Account pending system validation.", pending: true });
    if (user.status === "blocked")  return res.status(403).json({ error: "Account access has been restricted." });
    if (user.status === "rejected") return res.status(403).json({ error: "Account validation unsuccessful." });

    // Reset lockout on successful login
    await pool.execute("UPDATE `users` SET failed_attempts = 0, locked_until = NULL WHERE id = ?", [user.id]);

    const token = jwt.sign(
      { sub: user.id, username: user.username, jti: randomUUID(), tv: user.token_version },
      JWT_SECRET,
      { expiresIn: "8h" }
    );
    res.cookie("auth_token", token, { ...cookieOpts(req), maxAge: 8 * 60 * 60 * 1000 });
    res.json({ success: true, username: user.username });
  } catch (err) { next(err); }
});

app.get("/auth/me", requireAuth, async (req, res, next) => {
  try {
    const [rows] = await pool.execute("SELECT avatar_key, role_id FROM `users` WHERE id = ? LIMIT 1", [req.user.sub]);
    const row = rows[0];
    const hasAvatar    = !!row?.avatar_key;
    const avatarDataUrl = row?.avatar_key?.startsWith("data:") ? row.avatar_key : null;
    const isSuperAdmin = req.user.username?.toLowerCase() === ADMIN_USERNAME;
    const isAdmin      = isSuperAdmin;

    let permissions = [];
    let role = null;

    if (isSuperAdmin) {
      // Super admin gets all permissions
      const [allPerms] = await pool.execute("SELECT `key` FROM `permission` ORDER BY sort_order");
      permissions = allPerms.map(r => r.key);
    } else if (row?.role_id) {
      const [roleRows] = await pool.execute(
        "SELECT id, name FROM `role` WHERE id = ?",
        [row.role_id]
      );
      if (roleRows[0]) {
        role = { id: roleRows[0].id, name: roleRows[0].name };
        permissions = await loadUserPermissions(req.user.sub);
      }
    }

    res.json({ username: req.user.username, hasAvatar, avatarDataUrl, isAdmin, isSuperAdmin, role, permissions });
  } catch (err) { next(err); }
});


const uploadLimiter = rateLimit({ windowMs: 60_000, max: 5, standardHeaders: true, legacyHeaders: false, message: { error: "Too many uploads. Please wait 60 seconds." } });

app.post("/auth/avatar", ...requirePermission("account.avatar"), uploadLimiter, avatarUpload.single("avatar"), async (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No image uploaded." });
    const mime = detectImageMime(req.file.buffer.subarray(0, 12));
    if (!mime) return res.status(400).json({ error: "File content is not a valid image." });

    const { minio } = await import("./storage.js");
    let avatarValue;

    if (minio) {
      try {
        const ext = { "image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif" }[mime];
        const objectKey = `avatars/${randomUUID()}.${ext}`;
        const [existing] = await pool.execute("SELECT avatar_key FROM `users` WHERE id = ? LIMIT 1", [req.user.sub]);
        const oldKey = existing[0]?.avatar_key ?? null;
        await uploadObject(objectKey, req.file.buffer, req.file.size, mime);
        if (oldKey && !oldKey.startsWith("data:") && oldKey !== objectKey) await deleteObject(oldKey);
        avatarValue = objectKey;
      } catch {
        avatarValue = `data:${mime};base64,${req.file.buffer.toString("base64")}`;
      }
    } else {
      avatarValue = `data:${mime};base64,${req.file.buffer.toString("base64")}`;
    }

    await pool.execute("UPDATE `users` SET avatar_key = ? WHERE id = ?", [avatarValue, req.user.sub]);
    res.json({ success: true });
  } catch (err) { next(err); }
});

app.post("/auth/logout", async (req, res) => {
  const token = req.cookies?.auth_token;
  if (token) {
    try {
      const payload = jwt.verify(token, JWT_SECRET);
      // Increment token_version — immediately invalidates all sessions for this user
      if (payload?.sub) {
        await pool.execute("UPDATE `users` SET token_version = token_version + 1 WHERE id = ?", [payload.sub]);
      }
    } catch { /* invalid/expired token — still clear the cookie */ }
  }
  res.clearCookie("auth_token", cookieOpts(req));
  res.json({ success: true });
});

app.patch("/auth/password", ...requirePermission("account.change_password"), async (req, res, next) => {
  try {
    const { currentPassword, newPassword } = req.body ?? {};
    if (!currentPassword || !newPassword) return res.status(400).json({ error: "Both passwords are required." });
    if (newPassword.length < 8 || !/[A-Z]/.test(newPassword) || !/[0-9]/.test(newPassword) || !/[^A-Za-z0-9]/.test(newPassword))
      return res.status(400).json({ error: "New password must be 8+ characters with uppercase, number, and symbol." });
    if (currentPassword === newPassword) return res.status(400).json({ error: "New password must differ from current." });

    const [rows] = await pool.execute("SELECT password_hash FROM `users` WHERE id = ? LIMIT 1", [req.user.sub]);
    if (!rows[0]) return res.status(404).json({ error: "User not found." });
    if (!await bcrypt.compare(currentPassword, rows[0].password_hash))
      return res.status(401).json({ error: "Current password is incorrect." });

    const newHash = await bcrypt.hash(newPassword, 12);
    await pool.execute("UPDATE `users` SET password_hash = ? WHERE id = ?", [newHash, req.user.sub]);
    if (req.cookies?.auth_token) {
      try {
        const payload = jwt.verify(req.cookies.auth_token, JWT_SECRET);
        if (payload?.jti && payload?.exp) {
          const ttl = payload.exp - Math.floor(Date.now() / 1000);
          if (ttl > 0) await redis.set(`bl:${payload.jti}`, "1", { EX: ttl });
        }
      } catch { /* ignore */ }
    }
    res.clearCookie("auth_token", cookieOpts(req));
    res.json({ success: true });
  } catch (err) { next(err); }
});

// ── SSE ───────────────────────────────────────────────────────────────────────
app.get("/events", requireAuth, (req, res) => {
  const uid = req.user.sub;
  if ((sseUserCounts.get(uid) ?? 0) >= SSE_MAX_PER_USER) {
    return res.status(429).json({ error: "Too many active connections." });
  }
  sseUserCounts.set(uid, (sseUserCounts.get(uid) ?? 0) + 1);

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  res.setHeader("Transfer-Encoding", "chunked");
  res.socket?.setTimeout(0);
  res.socket?.setNoDelay(true);
  res.socket?.setKeepAlive(true, 0);
  res.flushHeaders();
  res.write("event: connected\ndata: {}\n\n");
  res.flush?.();
  sseClients.add(res);
  const hb = setInterval(() => {
    try { res.write(": heartbeat\n\n"); res.flush?.(); }
    catch { clearInterval(hb); sseClients.delete(res); }
  }, 15_000);
  req.on("close", () => {
    clearInterval(hb);
    sseClients.delete(res);
    sseUserCounts.set(uid, Math.max(0, (sseUserCounts.get(uid) ?? 1) - 1));
  });
});

// ── HEALTH ────────────────────────────────────────────────────────────────────
app.get("/health", async (_req, res) => {
  const checks = {};
  try { await pool.execute("SELECT 1"); checks.mysql = "ok"; } catch { checks.mysql = "error"; }
  try { await redis.ping(); checks.redis = "ok"; } catch { checks.redis = "error"; }
  const allOk = Object.values(checks).every(v => v === "ok");
  res.status(allOk ? 200 : 503).json({
    status: allOk ? "ok" : "degraded",
    uptime: process.uptime(),
    timestamp: Date.now(),
    checks,
    service: "psa-payroll-api",
  });
});

// ── EMPLOYEE ROUTES ───────────────────────────────────────────────────────────
const searchLimiter = rateLimit({ windowMs: 60_000, max: 60, standardHeaders: true, legacyHeaders: false, message: { error: "Too many searches." } });
const employeeWriteLimiter = rateLimit({ windowMs: 15 * 60 * 1000, max: 50, standardHeaders: true, legacyHeaders: false, message: { error: "Too many employee write requests. Please wait 15 minutes." } });

app.get("/employees", ...requirePermission("employees.view"), searchLimiter, async (req, res, next) => {
  try {
    const q      = String(req.query.q ?? "").slice(0, 100);
    const page   = Math.max(1, parseInt(req.query.page  ?? "1",  10) || 1);
    const limit  = Math.min(100, Math.max(1, parseInt(req.query.limit ?? "20", 10) || 20));
    const offset = (page - 1) * limit;
    const { rows, total } = await searchEmployeesPage(q, limit, offset);
    res.json({ data: rows, page, limit, total });
  } catch (err) { next(err); }
});

app.post("/employees", ...requirePermission("employees.create"), employeeWriteLimiter, async (req, res, next) => {
  try {
    const { employeeNumber, name, position, officeDivision, employmentType, positionId, divisionId, salaryGradeStepId, startDate } = req.body ?? {};
    let { basicMonthlyPay } = req.body ?? {};
    if (!employeeNumber?.trim()) return res.status(400).json({ error: "Employee number is required." });
    if (!name?.trim()) return res.status(400).json({ error: "Name is required." });

    // Resolve step → amount if a step was explicitly selected (same as PATCH handler)
    if (salaryGradeStepId) {
      const [stepRows] = await pool.execute("SELECT amount FROM salary_grade_step WHERE id = ?", [salaryGradeStepId]);
      if (stepRows[0]?.amount != null) basicMonthlyPay = Number(stepRows[0].amount);
    }

    const pay = parseBasicPay(basicMonthlyPay);
    if (!pay.ok) return res.status(400).json({ error: pay.error });

    const pos = await resolvePosition(positionId);
    if (!pos.ok) return res.status(400).json({ error: pos.error });

    try {
      const emp = await addEmployee({
        employeeNumber, officeDivision, employmentType,
        name, positionId, divisionId,
        position: pos.name ?? position,
        basicMonthlyPay: pay.value,
        startDate: startDate || null,
      });

      const finalPay = Number(emp.basic_monthly_pay) || pay.value;
      await syncComputedDeductionsToLedger(emp.id, finalPay);
      res.status(201).json(emp);
      broadcast("refresh");
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: `Employee number "${employeeNumber}" is already in use.` });
      throw err;
    }
  } catch (err) { next(err); }
});

app.patch("/employees/:id", ...requirePermission("employees.edit"), employeeWriteLimiter, async (req, res, next) => {
  try {
    const { id } = req.params;
    if (!EMP_ID_RE.test(id)) return res.status(400).json({ error: "Invalid ID format." });
    const body = { ...(req.body ?? {}) };

    // If a specific salary grade step was selected, resolve its amount as the new pay
    if (body.salaryGradeStepId) {
      const [stepRows] = await pool.execute(
        "SELECT amount FROM salary_grade_step WHERE id = ?",
        [body.salaryGradeStepId]
      );
      if (stepRows[0]?.amount != null) {
        body.basicMonthlyPay = Number(stepRows[0].amount);
      }
    }

    // If initialStep (1–8) provided without a salaryGradeStepId, look up the step amount
    // from the employee's current (or new) position's salary grade
    if (body.initialStep && !body.salaryGradeStepId) {
      const targetPositionId = body.positionId ?? (
        await pool.execute("SELECT position_id FROM `employee` WHERE id=?", [id])
          .then(([r]) => r[0]?.position_id ?? null)
      );
      if (targetPositionId) {
        const [stepRows] = await pool.execute(`
          SELECT sgs.amount FROM \`position\` p
          JOIN salary_grade sg ON sg.id = p.salary_grade_id
          JOIN salary_grade_step sgs ON sgs.salary_grade_id = sg.id AND sgs.step = ?
          WHERE p.id = ? ORDER BY sgs.effective_date DESC LIMIT 1
        `, [parseInt(body.initialStep, 10), targetPositionId]);
        if (stepRows[0]?.amount != null) {
          body.basicMonthlyPay = Number(stepRows[0].amount);
        }
      }
    }

    const pay = parseBasicPay(body.basicMonthlyPay);
    if (!pay.ok) return res.status(400).json({ error: pay.error });
    if (pay.value !== undefined) body.basicMonthlyPay = pay.value;

    if (body.positionId !== undefined) {
      const pos = await resolvePosition(body.positionId);
      if (!pos.ok) return res.status(400).json({ error: pos.error });
      if (pos.name !== undefined) body.position = pos.name;
    }

    const oldEmp = await getEmployeeById(id);
    const emp = await updateEmployee(id, body);
    if (!emp) return res.status(404).json({ error: "Employee not found." });

    insertAuditLog("update_employee", emp.name ?? emp.employee_number ?? id, req.user.username,
      { name: oldEmp?.name, status: oldEmp?.status, basic_monthly_pay: oldEmp?.basic_monthly_pay, employee_id: id },
      { name: emp.name, status: emp.status, basic_monthly_pay: emp.basic_monthly_pay, employee_id: id }
    ).catch(() => {});

    res.json(emp);
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.delete("/employees/:id", ...requirePermission("employees.delete"), employeeWriteLimiter, async (req, res, next) => {
  try {
    const { id } = req.params;
    if (!EMP_ID_RE.test(id)) return res.status(400).json({ error: "Invalid ID format." });
    const reason = req.body?.reason ?? null;
    const suspensionStart = req.body?.suspensionStart ?? null;
    const suspensionEnd   = req.body?.suspensionEnd   ?? null;
    const emp = await softDeleteEmployee(id, reason, suspensionStart, suspensionEnd);
    if (!emp) return res.status(404).json({ error: "Employee not found." });
    res.json({ success: true, employee: emp });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.get("/employees/live-payslips", ...requireAnyPermission("payslips.view", "deductions.view", "dashboard.view"), async (req, res, next) => {
  try {
    const q = req.query.q ?? "";
    const limit = Math.min(parseInt(req.query.limit ?? (q.trim() ? "50" : "5"), 10), 500);
    res.json(await getLivePayslips(q, limit));
  } catch (err) { next(err); }
});

app.get("/employees/inactive", ...requirePermission("employees.view"), async (req, res, next) => {
  try {
    const q = req.query.q ?? "";
    const limit = Math.min(parseInt(req.query.limit ?? (q.trim() ? "50" : "5"), 10), 500);
    res.json(await getInactiveEmployees(q, limit));
  } catch (err) { next(err); }
});

// GET /employees/suspension-log — all suspension/leave records with employee info and exact durations
app.get("/employees/suspension-log", ...requirePermission("employees.view"), async (req, res, next) => {
  try {
    const [rows] = await pool.execute(`
      SELECT
        es.id,
        es.employee_id,
        e.name        AS employee_name,
        e.employee_number,
        CASE es.reason WHEN 1 THEN 'On Leave' WHEN 2 THEN 'Suspended' ELSE 'Unknown' END AS reason,
        es.start_date,
        es.end_date,
        es.status,
        es.created_at,
        TIMESTAMPDIFF(HOUR,   es.start_date, COALESCE(es.end_date, NOW())) AS total_hours,
        TIMESTAMPDIFF(MINUTE, es.start_date, COALESCE(es.end_date, NOW())) % 60 AS extra_minutes
      FROM \`employee_suspension\` es
      JOIN \`employee\` e ON e.id = es.employee_id
      ORDER BY es.start_date DESC
    `);
    res.json(rows);
  } catch (err) { next(err); }
});

app.put("/employees/:id/activate", ...requirePermission("employees.restore"), employeeWriteLimiter, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const emp = await activateEmployee(req.params.id);
    if (!emp) return res.status(404).json({ error: "Employee not found or already active." });
    res.json({ success: true, employee: emp });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.get("/employees/:id/avatar", requireAuth, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(404).end();
    const [rows] = await pool.execute(
      "SELECT u.avatar_key FROM `users` u WHERE u.employee_id = ? LIMIT 1",
      [req.params.id]
    );
    const avatarKey = rows[0]?.avatar_key;
    if (!avatarKey?.startsWith("data:")) return res.status(404).end();
    const commaIdx = avatarKey.indexOf(",");
    const buf = Buffer.from(avatarKey.slice(commaIdx + 1), "base64");
    const mime = detectImageMime(buf.subarray(0, 12));
    if (!mime) return res.status(400).end();
    res.setHeader("Cache-Control", "private, max-age=3600");
    res.setHeader("Content-Type", mime);
    return res.send(buf);
  } catch (err) { next(err); }
});

app.get("/employees/:id/profile", ...requirePermission("employees.view"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID." });
    const profile = await getEmployeeProfile(req.params.id);
    if (!profile) return res.status(404).json({ error: "Employee not found." });
    res.json(profile);
  } catch (err) { next(err); }
});

app.post("/employees/:id/search-hit", requireAuth, (req, res) => res.json({ ok: true }));

app.get("/employees/:id/position-history", requireAuth, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    if (req.user.username.toLowerCase() !== ADMIN_USERNAME) {
      // Allow viewing own position history (Employee Dashboard); otherwise require employees.view
      const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [req.user.sub]);
      const isSelf = userRow?.employee_id != null && String(userRow.employee_id) === req.params.id;
      if (!isSelf) {
        if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
        if (!req.userPermissions.includes("employees.view")) return res.status(403).json({ error: "Permission denied." });
      }
    }
    const [rows] = await pool.execute(
      `SELECT ep.id, p.position_title, ep.movement_type, ep.start_date, ep.end_date, ep.status
       FROM employee_position ep
       JOIN \`position\` p ON p.id = ep.position_id
       WHERE ep.employee_id = ?
       ORDER BY ep.start_date DESC, ep.id DESC`,
      [req.params.id]
    );
    res.json(rows);
  } catch (err) { next(err); }
});

app.get("/employees/salary-standing", ...requirePermission("employees.view"), async (req, res, next) => {
  try {
    const search = String(req.query.q ?? "").trim().slice(0, 100);
    const whereSearch = search
      ? "AND (LOWER(e.name) LIKE LOWER(CONCAT('%',?,'%')) OR e.employee_number LIKE CONCAT('%',?,'%'))"
      : "";
    const params = search ? [search, search] : [];
    const [rows] = await pool.query(
      `SELECT e.id, e.employee_number, e.name, e.step, e.basic_monthly_pay,
              COALESCE(p.position_title, '') AS position,
              sg.grade_code AS salary_grade,
              sgs.amount AS rate_amount
       FROM \`employee\` e
       LEFT JOIN \`position\` p ON p.id = e.position_id
       LEFT JOIN \`salary_grade\` sg ON sg.id = p.salary_grade_id
       LEFT JOIN \`salary_grade_step\` sgs ON sgs.salary_grade_id = sg.id AND sgs.step = COALESCE(e.step, 1)
       WHERE e.status = 'active' ${whereSearch}
       ORDER BY e.name ASC LIMIT 200`,
      params
    );
    res.json(rows);
  } catch (err) { next(err); }
});

app.patch("/employees/:id/step", ...requirePermission("employees.manage"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const step = parseInt(req.body?.step, 10);
    if (!step || step < 1 || step > 8) return res.status(400).json({ error: "Step must be between 1 and 8." });

    const [[emp]] = await pool.execute(
      `SELECT e.id, p.salary_grade_id
       FROM \`employee\` e
       LEFT JOIN \`position\` p ON p.id = e.position_id
       WHERE e.id = ? AND e.status != 'deleted' LIMIT 1`,
      [req.params.id]
    );
    if (!emp) return res.status(404).json({ error: "Employee not found." });

    let newPay = null;
    if (emp.salary_grade_id) {
      const [[stepRow]] = await pool.execute(
        "SELECT amount FROM `salary_grade_step` WHERE salary_grade_id = ? AND step = ? LIMIT 1",
        [emp.salary_grade_id, step]
      );
      if (stepRow) newPay = stepRow.amount;
    }

    if (newPay !== null) {
      await pool.execute(
        "UPDATE `employee` SET step = ?, basic_monthly_pay = ? WHERE id = ?",
        [step, newPay, req.params.id]
      );
    } else {
      await pool.execute("UPDATE `employee` SET step = ? WHERE id = ?", [step, req.params.id]);
    }

    res.json({ ok: true, step, basic_monthly_pay: newPay });
  } catch (err) { next(err); }
});

app.get("/employees/search-all", ...requireAnyPermission("archive.view", "archive.manage"), searchLimiter, async (req, res, next) => {
  try {
    const q_str = String(req.query.q ?? "").trim().slice(0, 100);
    const limit = Math.min(parseInt(req.query.limit ?? (q_str ? "50" : "5"), 10), 500);
    const orderBy = "FIELD(e.status,'active','inactive'), e.name ASC";
    // mysql2 prepared statements (pool.execute) reject a bound LIMIT parameter
    // ("Incorrect arguments to mysqld_stmt_execute") — inline the sanitized integer instead.
    const [rows] = await pool.execute(
      `SELECT e.id, e.name, e.employee_number, e.status,
              p.position_title AS position,
              COALESCE(d.name, pd.name, '') AS office_division
       FROM \`employee\` e
       LEFT JOIN \`position\`  p  ON p.id  = e.position_id
       LEFT JOIN \`division\`  d  ON d.id  = e.division_id
       LEFT JOIN \`division\`  pd ON pd.id = p.division_id
       WHERE (? = '' OR LOWER(e.name) LIKE LOWER(CONCAT('%',?,'%'))
              OR e.employee_number LIKE CONCAT('%',?,'%'))
       ORDER BY ${orderBy}
       LIMIT ${limit}`,
      [q_str, q_str, q_str]
    );
    res.json(rows);
  } catch (e) { next(e); }
});

app.get("/employees/:id/payroll-history", requireAuth, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID." });
    // Allow viewing own payroll history without archive.view permission
    const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id = ?", [req.user.sub]);
    const isSelf = userRow?.employee_id != null && String(userRow.employee_id) === req.params.id;
    if (!isSelf) {
      if (req.user.username?.toLowerCase() !== ADMIN_USERNAME) {
        if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
        if (!req.userPermissions.some(p => ["archive.view","archive.manage"].includes(p)))
          return res.status(403).json({ error: "Forbidden." });
      }
    }
    const [rows] = await pool.execute(
      `SELECT year, month, quincena, gross_pay, total_earnings, total_deductions, net_pay, snapshot_json, saved_at
       FROM \`payroll_history\`
       WHERE employee_id = ?
       ORDER BY year DESC, month DESC, quincena DESC`,
      [req.params.id]
    );
    res.json(rows.map(r => ({
      year: r.year, month: r.month, quincena: r.quincena,
      gross_pay: Number(r.gross_pay),
      total_earnings: Number(r.total_earnings),
      total_deductions: Number(r.total_deductions),
      net_pay: Number(r.net_pay),
      saved_at: r.saved_at,
      snapshot: r.snapshot_json ? JSON.parse(r.snapshot_json) : null,
    })));
  } catch (e) { next(e); }
});

app.get("/employees/:id/deductions", requireAuth, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    if (req.user.username.toLowerCase() !== ADMIN_USERNAME) {
      // Allow viewing own deductions without deductions.view; otherwise require it
      const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [req.user.sub]);
      const isSelf = userRow?.employee_id != null && String(userRow.employee_id) === req.params.id;
      if (!isSelf) {
        if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
        if (!req.userPermissions.includes("deductions.view")) return res.status(403).json({ error: "Permission denied." });
      }
    }
    const sync = req.query.readonly !== "1";
    const rows = await getEmployeeDeductions(req.params.id, { sync });
    res.json(rows);
  } catch (err) { next(err); }
});

app.put("/employees/:id/deductions", ...requirePermission("deductions.edit"), employeeWriteLimiter, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const { column, value } = req.body ?? {};
    if (!column) return res.status(400).json({ error: "column is required." });
    const row = await upsertEmployeeDeduction(req.params.id, column, value ?? null, req.user.username);
    res.json(row);
  } catch (err) { next(err); }
});

app.put("/employees/:id/deductions/bulk", ...requirePermission("deductions.bulk_edit"), employeeWriteLimiter, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const items = req.body;
    if (!Array.isArray(items)) return res.status(400).json({ error: "Body must be an array of {column, value}." });
    if (items.length > 30) return res.status(400).json({ error: "Too many items in bulk request (max 30)." });
    const rows = await Promise.all(
      items.filter(i => i.column).map(i => upsertEmployeeDeduction(req.params.id, i.column, i.value ?? null, req.user.username))
    );
    res.json(rows);
  } catch (err) { next(err); }
});

// ── Deduction Ledger + Audit Log ───────────────────────────────────────────────

app.get("/employees/:id/ledger", ...requirePermission("deductions.view"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const rows = await getEmployeeLedger(req.params.id);
    res.json(rows);
  } catch (err) { next(err); }
});

app.get("/audit-log", ...requirePermission("deductions.view"), async (req, res, next) => {
  try {
    const rows = await getAuditLog({ q: req.query.q ?? "", entity: req.query.entity ?? "", limit: req.query.limit });
    res.json(rows);
  } catch (err) { next(err); }
});


// ── Employee Loans ─────────────────────────────────────────────────────────────

app.get("/employees/:id/loans", requireAuth, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    if (req.user.username.toLowerCase() !== ADMIN_USERNAME) {
      const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [req.user.sub]);
      const isSelf = userRow?.employee_id != null && String(userRow.employee_id) === req.params.id;
      if (!isSelf) {
        if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
        if (!req.userPermissions.includes("deductions.view")) return res.status(403).json({ error: "Permission denied." });
      }
    }
    const loans = await getEmployeeLoans(req.params.id);
    res.json(loans);
  } catch (err) { next(err); }
});

app.get("/employees/:id/loan-payments", requireAuth, async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    if (req.user.username.toLowerCase() !== ADMIN_USERNAME) {
      const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [req.user.sub]);
      const isSelf = userRow?.employee_id != null && String(userRow.employee_id) === req.params.id;
      if (!isSelf) {
        if (!req.userPermissions) req.userPermissions = await loadUserPermissions(req.user.sub);
        if (!req.userPermissions.includes("deductions.view")) return res.status(403).json({ error: "Permission denied." });
      }
    }
    res.json(await getLoanPaymentHistory(req.params.id));
  } catch (err) { next(err); }
});

app.post("/employees/:id/loans", ...requirePermission("deductions.manage"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const { payrollTypeId, totalAmount, amountPerDeduction, startDate } = req.body ?? {};
    if (!payrollTypeId || !totalAmount || !amountPerDeduction || !startDate)
      return res.status(400).json({ error: "payrollTypeId, totalAmount, amountPerDeduction, and startDate are required." });
    const loan = await createLoan({
      employeeId: req.params.id,
      payrollTypeId, totalAmount, amountPerDeduction, startDate,
      createdBy: req.user.username,
    });
    res.status(201).json(loan);
  } catch (err) { next(err); }
});

app.post("/employees/:id/loans/:loanId/consolidate", ...requirePermission("deductions.manage"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const { newLoanAmount, amountPerDeduction, startDate } = req.body ?? {};
    if (!newLoanAmount || !amountPerDeduction || !startDate)
      return res.status(400).json({ error: "newLoanAmount, amountPerDeduction, and startDate are required." });
    const result = await consolidateLoan({
      oldLoanId: Number(req.params.loanId),
      newLoanAmount, amountPerDeduction, startDate,
      createdBy: req.user.username,
    });
    res.json(result);
  } catch (err) { next(err); }
});

app.patch("/employees/:id/loans/:loanId/close", ...requirePermission("deductions.manage"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const result = await closeLoan(Number(req.params.loanId), req.user.username);
    res.json(result);
  } catch (err) { next(err); }
});

app.patch("/employees/:id/loans/:loanId", ...requirePermission("deductions.manage"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const { remainingBalance, addAmount, monthlyAmount } = req.body ?? {};
    if (addAmount != null && Number(addAmount) > 0) {
      await extendLoan(Number(req.params.loanId), Number(addAmount), req.user?.username ?? "system");
    } else {
      if (remainingBalance != null) {
        await pool.execute(
          `UPDATE employee_loans SET remaining_balance = ? WHERE id = ?`,
          [Number(remainingBalance), Number(req.params.loanId)]
        );
      }
      if (monthlyAmount != null && Number(monthlyAmount) > 0) {
        await pool.execute(
          `UPDATE employee_loans SET monthly_amount = ? WHERE id = ?`,
          [Number(monthlyAmount), Number(req.params.loanId)]
        );
      }
    }
    const loans = await getEmployeeLoans(req.params.id);
    res.json(loans);
  } catch (err) { next(err); }
});

// ⚠️ PROTOTYPE TEST ONLY — remove before production
app.post("/employees/:id/loans/simulate-period", ...requirePermission("deductions.manage"), async (req, res, next) => {
  try {
    if (!EMP_ID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID format." });
    const result = await processLoanDeductionsForEmployee(req.params.id);
    const loans  = await getEmployeeLoans(req.params.id);
    res.json({ ...result, loans });
  } catch (err) { next(err); }
});


app.get("/stats", ...requirePermission("dashboard.view"), async (req, res, next) => {
  try { res.json(await getPayrollStats()); } catch (err) { next(err); }
});

// ── LOOKUP TABLES ─────────────────────────────────────────────────────────────
app.get("/positions", ...requirePermission("settings.reference.view"), async (req, res, next) => {
  try {
    const q = req.query.q ?? "";
    const limit = Math.min(parseInt(req.query.limit ?? (q.trim() ? "50" : "5"), 10), 500);
    res.json(await getPositions(q, limit));
  } catch (err) { next(err); }
});
app.post("/positions", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const { name, divisionId, defaultSalaryGrade } = req.body ?? {};
    if (!name?.trim()) return res.status(400).json({ error: "Position name is required." });
    try {
      res.status(201).json(await createPosition(name, divisionId, defaultSalaryGrade));
      broadcast("refresh");
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: "Position already exists." });
      throw err;
    }
  } catch (err) { next(err); }
});
app.patch("/positions/:id", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const id = parseInt(req.params.id, 10);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const { name, divisionId, defaultSalaryGrade } = req.body ?? {};
    if (!name?.trim()) return res.status(400).json({ error: "Position name is required." });

    // Capture old salary grade before update
    const [[oldPos]] = await pool.execute(
      "SELECT sg.grade_code AS default_salary_grade FROM `position` p LEFT JOIN salary_grade sg ON sg.id = p.salary_grade_id WHERE p.id = ?", [id]
    );
    const oldGrade = oldPos?.default_salary_grade ?? null;

    const row = await updatePosition(id, name, divisionId, defaultSalaryGrade);
    if (!row) return res.status(404).json({ error: "Position not found." });

    // If salary grade changed, update basic_monthly_pay for all active employees on this position
    // and resync their computed deductions (GSIS, PhilHealth, etc.)
    if (defaultSalaryGrade && defaultSalaryGrade !== oldGrade) {
      const [[step1]] = await pool.execute(`
        SELECT sgs.amount
        FROM salary_grade sg
        JOIN salary_grade_step sgs ON sgs.salary_grade_id = sg.id AND sgs.step = 1
        WHERE sg.grade_code = ?
        ORDER BY sgs.effective_date DESC LIMIT 1
      `, [defaultSalaryGrade]);
      if (step1) {
        const newAmount = Number(step1.amount);
        const [affected] = await pool.execute(
          "SELECT id FROM `employee` WHERE position_id = ? AND status = 'active'",
          [id]
        );
        await pool.execute(
          "UPDATE `employee` SET basic_monthly_pay = ? WHERE position_id = ? AND status = 'active'",
          [newAmount, id]
        );
        for (const emp of affected) {
          await syncComputedDeductionsToLedger(emp.id, newAmount);
        }
      }
    }

    res.json(row);
    broadcast("refresh");
  } catch (err) { next(err); }
});
app.delete("/positions/:id", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const id = parseInt(req.params.id, 10);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const row = await deletePosition(id);
    if (!row) return res.status(404).json({ error: "Position not found." });
    res.json({ success: true });
    broadcast("refresh");
  } catch (err) { next(err); }
});

// ── Salary Grades ─────────────────────────────────────────────────────────────
app.get("/salary-grades", requireAuth, async (req, res, next) => {
  try {
    const { effective_date } = req.query;
    const [rows] = await pool.execute(
      `SELECT sgs.id, sg.id AS salary_grade_id, sg.grade_code AS salary_grade,
              sgs.step, sgs.amount, sgs.effective_date, sgs.created_at
       FROM salary_grade_step sgs
       JOIN salary_grade sg ON sg.id = sgs.salary_grade_id
       ${effective_date ? "WHERE sgs.effective_date = ?" : ""}
       ORDER BY sg.grade_code + 0, sgs.step, sgs.effective_date`,
      effective_date ? [effective_date] : []
    );
    res.json(rows);
  } catch (err) { next(err); }
});

app.get("/salary-grades/list", requireAuth, async (req, res, next) => {
  try {
    const [rows] = await pool.execute("SELECT id, grade_code FROM salary_grade ORDER BY grade_code + 0");
    res.json(rows);
  } catch (err) { next(err); }
});

app.get("/salary-grades/ranked-list", requireAuth, async (req, res, next) => {
  try {
    const q = (req.query.q ?? "").trim();
    const limit = q ? Math.min(parseInt(req.query.limit) || 50, 200) : 5;
    // mysql2 prepared statements (pool.execute) reject a bound LIMIT parameter
    // ("Incorrect arguments to mysqld_stmt_execute") — inline the sanitized integer instead.
    let gradeRows;
    if (q) {
      [gradeRows] = await pool.execute(
        `SELECT id, grade_code FROM salary_grade WHERE grade_code LIKE ? ORDER BY CAST(REGEXP_REPLACE(grade_code, '[^0-9]', '') AS UNSIGNED) LIMIT ${limit}`,
        [`%${q}%`]
      );
    } else {
      // Most recently opened first (NULLs sort last under DESC), then by grade number.
      // grade_code mixes "SG 9" and "SG8", so sort on its digits — `grade_code + 0` casts all of them to 0.
      [gradeRows] = await pool.execute(
        `SELECT id, grade_code FROM salary_grade ORDER BY last_accessed_at DESC, CAST(REGEXP_REPLACE(grade_code, '[^0-9]', '') AS UNSIGNED) LIMIT ${limit}`
      );
    }
    if (gradeRows.length === 0) return res.json([]);
    const gradeIds = gradeRows.map(r => r.id);
    const [stepRows] = await pool.execute(
      `SELECT id, salary_grade_id, step, amount, effective_date FROM salary_grade_step WHERE salary_grade_id IN (${gradeIds.map(() => "?").join(",")}) ORDER BY salary_grade_id, step`,
      gradeIds
    );
    const stepsByGrade = {};
    for (const s of stepRows) {
      if (!stepsByGrade[s.salary_grade_id]) stepsByGrade[s.salary_grade_id] = [];
      stepsByGrade[s.salary_grade_id].push(s);
    }
    res.json(gradeRows.map(g => ({ rank: g.grade_code, steps: stepsByGrade[g.id] ?? [] })));
  } catch (err) { next(err); }
});


app.post("/salary-grades", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const { salary_grade, step, amount, effective_date } = req.body ?? {};
    if (!salary_grade?.trim()) return res.status(400).json({ error: "salary_grade is required." });
    const gradeCode = salary_grade.trim();
    if (gradeCode.length > 20) return res.status(400).json({ error: "salary_grade too long." });
    if (!/^SG\s?\d{1,2}$/i.test(gradeCode)) return res.status(400).json({ error: "Invalid salary grade format. Use 'SG 1' through 'SG 33'." });
    const stepInt = parseInt(step, 10);
    if (step == null || isNaN(stepInt) || stepInt < 1 || stepInt > 20) return res.status(400).json({ error: "step must be a number between 1 and 20." });
    const amountVal = parseFloat(amount);
    if (isNaN(amountVal) || amountVal < 0) return res.status(400).json({ error: "amount must be a non-negative number." });
    if (effective_date && !/^\d{4}-\d{2}-\d{2}$/.test(effective_date)) return res.status(400).json({ error: "effective_date must be YYYY-MM-DD." });
    const effDate = effective_date || "2024-01-01";
    const [[existing]] = await pool.execute("SELECT id FROM `salary_grade` WHERE grade_code = ?", [gradeCode]);
    if (!existing) {
      try {
        await pool.execute("INSERT INTO `salary_grade` (grade_code) VALUES (?)", [gradeCode]);
      } catch (e) {
        if (e.code === "ER_DUP_ENTRY") return res.status(409).json({ error: `${gradeCode} already exists.` });
        throw e;
      }
    }
    const [[sg]] = await pool.execute("SELECT id FROM `salary_grade` WHERE grade_code = ?", [gradeCode]);
    await pool.execute(
      `INSERT INTO salary_grade_step (salary_grade_id, step, amount, effective_date)
       VALUES (?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE amount = VALUES(amount)`,
      [sg.id, parseInt(step, 10), parseFloat(amount) || 0, effDate]
    );
    const [[row]] = await pool.execute(
      `SELECT sgs.id, sg.id AS salary_grade_id, sg.grade_code AS salary_grade,
              sgs.step, sgs.amount, sgs.effective_date
       FROM salary_grade_step sgs JOIN salary_grade sg ON sg.id = sgs.salary_grade_id
       WHERE sgs.salary_grade_id = ? AND sgs.step = ? AND sgs.effective_date = ?`,
      [sg.id, parseInt(step, 10), effDate]
    );
    res.status(201).json(row);
  } catch (err) { next(err); }
});

app.patch("/salary-grades/:id", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const id = parseInt(req.params.id, 10);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const { amount, effective_date } = req.body ?? {};
    const fields = [];
    const vals = [];
    if (amount != null) {
      const amountVal = parseFloat(amount);
      if (isNaN(amountVal) || amountVal < 0) return res.status(400).json({ error: "amount must be a non-negative number." });
      fields.push("amount = ?"); vals.push(amountVal);
    }
    if (effective_date != null) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(effective_date)) return res.status(400).json({ error: "effective_date must be YYYY-MM-DD." });
      fields.push("effective_date = ?"); vals.push(effective_date);
    }
    if (fields.length) {
      vals.push(id);
      await pool.execute(`UPDATE salary_grade_step SET ${fields.join(", ")} WHERE id = ?`, vals);
    }
    const [[row]] = await pool.execute(
      `SELECT sgs.id, sg.id AS salary_grade_id, sg.grade_code AS salary_grade,
              sgs.step, sgs.amount, sgs.effective_date
       FROM salary_grade_step sgs JOIN salary_grade sg ON sg.id = sgs.salary_grade_id
       WHERE sgs.id = ?`,
      [id]
    );
    if (!row) return res.status(404).json({ error: "Not found." });
    res.json(row);
  } catch (err) { next(err); }
});

app.delete("/salary-grades/:id", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const id = parseInt(req.params.id, 10);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const [[inUse]] = await pool.execute(
      "SELECT COUNT(*) AS cnt FROM `position` WHERE salary_grade_id = (SELECT salary_grade_id FROM salary_grade_step WHERE id = ?)", [id]
    );
    if (inUse.cnt > 0) return res.status(409).json({ error: "This salary grade is assigned to a position and cannot be deleted." });
    await pool.execute("DELETE FROM salary_grade_step WHERE id = ?", [id]);
    res.status(204).end();
  } catch (err) { next(err); }
});

// Delete all steps for a salary grade + the parent grade row
app.delete("/salary-grades/rank/:rank", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const [[sg]] = await pool.execute("SELECT id FROM `salary_grade` WHERE grade_code = ?", [req.params.rank]);
    if (!sg) return res.status(404).json({ error: "Salary grade not found." });
    const [[inUse]] = await pool.execute(
      "SELECT COUNT(*) AS cnt FROM `position` WHERE salary_grade_id = ?", [sg.id]
    );
    if (inUse.cnt > 0) return res.status(409).json({ error: "This salary grade is assigned to one or more positions and cannot be deleted." });
    await pool.execute("DELETE FROM salary_grade_step WHERE salary_grade_id = ?", [sg.id]);
    await pool.execute("DELETE FROM `salary_grade` WHERE id = ?", [sg.id]);
    res.status(204).end();
  } catch (err) { next(err); }
});

app.get("/divisions", ...requirePermission("settings.reference.view"), async (req, res, next) => {
  try { res.json(await getDivisions()); } catch (err) { next(err); }
});
app.post("/divisions", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const { name } = req.body ?? {};
    if (!name?.trim()) return res.status(400).json({ error: "Division name is required." });
    try {
      res.status(201).json(await createDivision(name));
      broadcast("refresh");
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: "Division already exists." });
      throw err;
    }
  } catch (err) { next(err); }
});
app.patch("/divisions/:id", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const id = parseInt(req.params.id, 10);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const { name } = req.body ?? {};
    if (!name?.trim()) return res.status(400).json({ error: "Division name is required." });
    const row = await updateDivision(id, name);
    if (!row) return res.status(404).json({ error: "Division not found." });
    res.json(row);
    broadcast("refresh");
  } catch (err) { next(err); }
});
app.delete("/divisions/:id", ...requirePermission("settings.reference.manage"), async (req, res, next) => {
  try {
    const id = parseInt(req.params.id, 10);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const row = await deleteDivision(id);
    if (!row) return res.status(404).json({ error: "Division not found." });
    res.json({ success: true });
    broadcast("refresh");
  } catch (err) { next(err); }
});

// ── REFERENCE DATA SEARCH-HITS ───────────────────────────────────────────────
app.post("/positions/:id/search-hit", requireAuth, async (req, res, next) => {
  try { await hitPosition(req.params.id); res.json({ success: true }); }
  catch (err) { next(err); }
});
app.post("/divisions/:id/search-hit", requireAuth, (_req, res) => res.json({ success: true }));
app.post("/salary-grades/rank/:rank/search-hit", requireAuth, async (req, res, next) => {
  try {
    await pool.execute("UPDATE `salary_grade` SET last_accessed_at = NOW() WHERE grade_code = ?", [req.params.rank]);
    res.json({ success: true });
  } catch (err) { next(err); }
});

// ── DEDUCTION TYPES ──────────────────────────────────────────────────────────
app.get("/settings/deduction-types", ...requireAnyPermission("settings.payroll_types.view", "settings.payroll_types.manage"), async (req, res, next) => {
  try { res.json(await getDeductionTypes({ includeInactive: req.query.all === '1' })); }
  catch (e) { next(e); }
});
app.post("/settings/deduction-types", ...requirePermission("settings.payroll_types.manage"), async (req, res, next) => {
  try {
    const { name, display_name } = req.body ?? {};
    if (!name?.trim()) return res.status(400).json({ error: "name (slug) is required." });
    if (!display_name?.trim()) return res.status(400).json({ error: "display_name is required." });
    try {
      res.status(201).json(await createDeductionType(req.body));
      broadcast("refresh");
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: "A deduction type with that name already exists." });
      throw err;
    }
  } catch (e) { next(e); }
});
app.patch("/settings/deduction-types/:id", ...requirePermission("settings.payroll_types.manage"), async (req, res, next) => {
  try {
    if (!req.params.id) return res.status(400).json({ error: "Invalid ID." });
    const t = await updateDeductionType(req.params.id, req.body);
    if (!t) return res.status(404).json({ error: "Not found." });
    if (req.body.is_active === 0 || req.body.is_active === false) {
      await pool.execute(
        `UPDATE \`employee_deduction_ledger\` SET status = 'inactive' WHERE payroll_type_id = ? AND status = 'active'`,
        [req.params.id]
      );
    }
    res.json(t);
    broadcast("refresh");
  } catch (e) { next(e); }
});
app.delete("/settings/deduction-types/:id", ...requirePermission("settings.payroll_types.manage"), async (req, res, next) => {
  try {
    if (!req.params.id) return res.status(400).json({ error: "Invalid ID." });
    await pool.execute(
      `UPDATE \`employee_deduction_ledger\` SET status = 'inactive' WHERE payroll_type_id = ? AND status = 'active'`,
      [req.params.id]
    );
    res.json(await softDeleteDeductionType(req.params.id));
    broadcast("refresh");
  } catch (e) { next(e); }
});

// ── EARNING TYPES ─────────────────────────────────────────────────────────────
app.get("/settings/earning-types", ...requireAnyPermission("settings.payroll_types.view", "settings.payroll_types.manage"), async (req, res, next) => {
  try { res.json(await getEarningTypes({ includeInactive: req.query.all === '1' })); }
  catch (e) { next(e); }
});
app.post("/settings/earning-types", ...requirePermission("settings.payroll_types.manage"), async (req, res, next) => {
  try {
    const { name, display_name } = req.body ?? {};
    if (!name?.trim()) return res.status(400).json({ error: "name (slug) is required." });
    if (!display_name?.trim()) return res.status(400).json({ error: "display_name is required." });
    try {
      res.status(201).json(await createEarningType(req.body));
      broadcast("refresh");
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: "An earning type with that name already exists." });
      throw err;
    }
  } catch (e) { next(e); }
});
app.patch("/settings/earning-types/:id", ...requirePermission("settings.payroll_types.manage"), async (req, res, next) => {
  try {
    if (!req.params.id) return res.status(400).json({ error: "Invalid ID." });
    const t = await updateEarningType(req.params.id, req.body);
    if (!t) return res.status(404).json({ error: "Not found." });
    res.json(t);
    broadcast("refresh");
  } catch (e) { next(e); }
});
app.delete("/settings/earning-types/:id", ...requirePermission("settings.payroll_types.manage"), async (req, res, next) => {
  try {
    if (!req.params.id) return res.status(400).json({ error: "Invalid ID." });
    res.json(await softDeleteEarningType(req.params.id));
    broadcast("refresh");
  } catch (e) { next(e); }
});

// ── SIGNATORIES ───────────────────────────────────────────────────────────────
app.get("/settings/signatories", requireAuth, async (req, res, next) => {
  try { res.json(await getSignatories()); }
  catch (e) { next(e); }
});

app.put("/settings/signatories", requireAuth, async (req, res, next) => {
  try {
    if (req.user.username.toLowerCase() !== ADMIN_USERNAME) return res.status(403).json({ error: "Superadmin only." });
    const { preparedByName, preparedByTitle, certifiedByName, certifiedByTitle,
            approvingName, approvingTitle, disbursingName, disbursingTitle } = req.body ?? {};
    const data = {
      prepared_by_name:  preparedByName?.trim()  ?? "",
      prepared_by_title: preparedByTitle?.trim() ?? "",
      certified_by_name:  certifiedByName?.trim()  ?? "",
      certified_by_title: certifiedByTitle?.trim() ?? "",
      approving_name:  approvingName?.trim()  ?? "",
      approving_title: approvingTitle?.trim() ?? "",
      disbursing_name:  disbursingName?.trim()  ?? "",
      disbursing_title: disbursingTitle?.trim() ?? "",
    };
    res.json(await setSignatories(data));
  } catch (e) { next(e); }
});

// ── PAYROLL ARCHIVE ───────────────────────────────────────────────────────────
app.get("/archives", ...requireAnyPermission("archive.view", "archive.manage"), async (req, res, next) => {
  try { res.json(await getArchives()); } catch (e) { next(e); }
});
app.get("/archives/:year/:month/:quincena", ...requireAnyPermission("archive.view", "archive.manage"), async (req, res, next) => {
  try {
    const arc = await getArchive(req.params.year, req.params.month, req.params.quincena);
    if (!arc) return res.status(404).json({ error: "Archive not found." });

    insertAuditLog("export_view", `${req.params.year}-${req.params.month}-${req.params.quincena}`, req.user.username, null, null).catch(() => {});

    const canManage = req.user.username?.toLowerCase() === ADMIN_USERNAME ||
      (await loadUserPermissions(req.user.sub)).includes("archive.manage");
    if (!canManage && req.user.division_id && arc.snapshot?.items) {
      arc.snapshot.items = arc.snapshot.items.filter(i => i.employee?.division_id === req.user.division_id);
    }

    res.json(arc);
  } catch (e) { next(e); }
});
app.put("/archives/:year/:month/:quincena", ...requirePermission("archive.manage"), async (req, res, next) => {
  try {
    const { year, month, quincena } = req.params;
    const { items, period, signatories, dedTypes, earnTypes, notes } = req.body ?? {};
    if (!Array.isArray(items)) return res.status(400).json({ error: "items must be an array." });
    const snapshot = { items, period, signatories, dedTypes, earnTypes };
    const id = await upsertArchive(year, month, quincena, snapshot, items.length, req.user.username, notes);
    insertAuditLog("upsert_archive", `${year}-${month}-${quincena}`, req.user.username, null, { employee_count: items.length }).catch(() => {});
    if (items.length > 0) savePayrollHistory(items, year, month, quincena).catch(e => logger.error({ err: e }, "savePayrollHistory failed"));

    // Process loan deductions for all employees in this archive
    const uniqueEmpIds = [...new Set(items.map(item => item.employee?.id).filter(Boolean))];
    const loanResults = { halted: [], deducted: [] };
    for (const empId of uniqueEmpIds) {
      try {
        const r = await processLoanDeductionsForEmployee(empId);
        loanResults.halted.push(...r.halted);
        loanResults.deducted.push(...r.deducted);
      } catch (e) { logger.error({ err: e }, `processLoanDeductions failed for ${empId}`); }
    }

    res.json({ id, saved: true, loanResults });
  } catch (e) { next(e); }
});
app.delete("/archives/:id", ...requirePermission("archive.manage"), async (req, res, next) => {
  try {
    const archiveId = parseInt(req.params.id, 10);
    if (!archiveId) return res.status(400).json({ error: "Invalid ID." });
    await deleteArchive(archiveId);
    insertAuditLog("delete_archive", String(archiveId), req.user.username, null, null).catch(() => {});
    res.json({ deleted: true });
  } catch (e) { next(e); }
});

// ── ADMIN ROUTES ──────────────────────────────────────────────────────────────
app.get("/admin/pending-users", ...requirePermission("admin.users.view"), async (req, res, next) => {
  try { res.json(await getPendingUsers()); } catch (err) { next(err); }
});

app.post("/admin/approve/:userId", ...requirePermission("admin.users.approve"), async (req, res, next) => {
  try {
    if (!UUID_RE.test(req.params.userId)) return res.status(400).json({ error: "Invalid ID format." });
    const user = await approveUser(req.params.userId);
    if (!user) return res.status(404).json({ error: "User not found or already active." });
    res.json({ success: true, user });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.post("/admin/reject/:userId", ...requirePermission("admin.users.approve"), async (req, res, next) => {
  try {
    if (!UUID_RE.test(req.params.userId)) return res.status(400).json({ error: "Invalid ID format." });
    const user = await rejectUser(req.params.userId);
    if (!user) return res.status(404).json({ error: "User not found or not pending." });
    res.json({ success: true, user });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.get("/admin/users", ...requirePermission("admin.users.view"), async (req, res, next) => {
  try { res.json(await getAllManagedUsers()); } catch (err) { next(err); }
});

app.post("/admin/block/:userId", ...requirePermission("admin.users.block"), async (req, res, next) => {
  try {
    if (!UUID_RE.test(req.params.userId)) return res.status(400).json({ error: "Invalid ID format." });
    const user = await blockUser(req.params.userId);
    if (!user) return res.status(404).json({ error: "User not found." });
    // Invalidate all active sessions immediately
    await pool.execute("UPDATE `users` SET token_version = token_version + 1 WHERE id = ?", [req.params.userId]);
    res.json({ success: true, user });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.post("/admin/unblock/:userId", ...requirePermission("admin.users.block"), async (req, res, next) => {
  try {
    if (!UUID_RE.test(req.params.userId)) return res.status(400).json({ error: "Invalid ID format." });
    const user = await unblockUser(req.params.userId);
    if (!user) return res.status(404).json({ error: "User not found or not blocked." });
    res.json({ success: true, user });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.post("/admin/unlock/:userId", ...requirePermission("admin.users.block"), async (req, res, next) => {
  try {
    if (!UUID_RE.test(req.params.userId)) return res.status(400).json({ error: "Invalid ID format." });
    const [[row]] = await pool.execute("SELECT id FROM `users` WHERE id = ? LIMIT 1", [req.params.userId]);
    if (!row) return res.status(404).json({ error: "User not found." });
    await pool.execute("UPDATE `users` SET failed_attempts = 0, locked_until = NULL WHERE id = ?", [req.params.userId]);
    res.json({ success: true });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.get("/admin/audit-log", ...requirePermission("admin.users.view"), async (req, res, next) => {
  try {
    const logs = await getAuditLog({ q: req.query.q, entity: req.query.entity, limit: req.query.limit });
    res.json(logs);
  } catch (err) { next(err); }
});

app.delete("/admin/users/:userId", ...requirePermission("admin.users.delete"), async (req, res, next) => {
  try {
    const { userId } = req.params;
    if (!UUID_RE.test(userId)) return res.status(400).json({ error: "Invalid ID format." });
    const user = await deleteSystemUser(userId);
    if (!user) return res.status(404).json({ error: "User not found or cannot be deleted." });
    res.json({ success: true, user });
    broadcast("refresh");
  } catch (err) { next(err); }
});

app.patch("/admin/users/:userId/password", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const { userId } = req.params;
    if (!UUID_RE.test(userId)) return res.status(400).json({ error: "Invalid ID format." });
    const { newPassword } = req.body ?? {};
    if (!newPassword || typeof newPassword !== "string" || newPassword.length < 8)
      return res.status(400).json({ error: "Password must be at least 8 characters." });
    if (!/[A-Z]/.test(newPassword) || !/[0-9]/.test(newPassword) || !/[^A-Za-z0-9]/.test(newPassword))
      return res.status(400).json({ error: "Password must include uppercase, number, and symbol." });
    const [rows] = await pool.execute("SELECT id, username FROM users WHERE id = ?", [userId]);
    if (!rows.length) return res.status(404).json({ error: "User not found." });
    if (rows[0].username === ADMIN_USERNAME) return res.status(403).json({ error: "Cannot reset the super admin password." });
    const hash = await bcrypt.hash(newPassword, 12);
    await pool.execute("UPDATE users SET password_hash = ? WHERE id = ?", [hash, userId]);
    res.json({ success: true });
  } catch (err) { next(err); }
});

// ── ROLE MANAGEMENT (super admin only) ────────────────────────────────────────

// GET /admin/permissions — return permission catalog grouped by category
app.get("/admin/permissions", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const [rows] = await pool.execute("SELECT `key`, label, category, sort_order FROM `permission` ORDER BY sort_order");
    const grouped = {};
    for (const p of rows) {
      if (!grouped[p.category]) grouped[p.category] = [];
      grouped[p.category].push({ key: p.key, label: p.label });
    }
    res.json(grouped);
  } catch (err) { next(err); }
});

// GET /admin/roles
app.get("/admin/roles", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const [roles] = await pool.execute(
      "SELECT id, name, description, created_at FROM `role` ORDER BY created_at ASC"
    );
    if (roles.length === 0) return res.json([]);

    // Fetch all role_permission + permission keys
    const [rpRows] = await pool.execute(
      "SELECT rp.role_id, p.`key` FROM `role_permission` rp JOIN `permission` p ON p.id = rp.permission_id"
    );
    const permMap = {};
    for (const rp of rpRows) {
      if (!permMap[rp.role_id]) permMap[rp.role_id] = [];
      permMap[rp.role_id].push(rp.key);
    }

    // Fetch user counts per role
    const [ucRows] = await pool.execute(
      "SELECT role_id, COUNT(*) AS cnt FROM `users` WHERE role_id IS NOT NULL GROUP BY role_id"
    );
    const ucMap = {};
    for (const r of ucRows) ucMap[r.role_id] = parseInt(r.cnt, 10);

    res.json(roles.map(r => ({
      ...r,
      permission_keys: permMap[r.id] ?? [],
      user_count: ucMap[r.id] ?? 0,
    })));
  } catch (err) { next(err); }
});

// POST /admin/roles
app.post("/admin/roles", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const name = stripTags(req.body?.name ?? "").trim();
    const description = stripTags(req.body?.description ?? "").trim();
    const permissionKeys = Array.isArray(req.body?.permissionKeys) ? req.body.permissionKeys : [];
    if (!name || name.length > 100) return res.status(400).json({ error: "Role name is required (max 100 chars)." });

    let roleId;
    try {
      const [result] = await pool.execute(
        "INSERT INTO `role` (name, description) VALUES (?, ?)",
        [name, description || null]
      );
      roleId = result.insertId;
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: "A role with that name already exists." });
      throw err;
    }

    if (permissionKeys.length > 0) {
      const placeholders = permissionKeys.map(() => "?").join(",");
      const [perms] = await pool.execute(
        `SELECT id FROM \`permission\` WHERE \`key\` IN (${placeholders})`,
        permissionKeys
      );
      for (const p of perms) {
        await pool.execute(
          "INSERT IGNORE INTO `role_permission` (role_id, permission_id) VALUES (?, ?)",
          [roleId, p.id]
        );
      }
    }

    const [rows] = await pool.execute("SELECT id, name, description, created_at FROM `role` WHERE id = ?", [roleId]);
    res.status(201).json({ ...rows[0], permission_keys: permissionKeys, user_count: 0 });
  } catch (err) { next(err); }
});

// PATCH /admin/roles/:id
app.patch("/admin/roles/:id", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const { id } = req.params;
    const name = stripTags(req.body?.name ?? "").trim();
    const description = stripTags(req.body?.description ?? "").trim();
    if (!name || name.length > 100) return res.status(400).json({ error: "Role name is required (max 100 chars)." });
    try {
      await pool.execute(
        "UPDATE `role` SET name=?, description=? WHERE id=?",
        [name, description || null, id]
      );
    } catch (err) {
      if (err.code === "ER_DUP_ENTRY") return res.status(409).json({ error: "A role with that name already exists." });
      throw err;
    }
    const [rows] = await pool.execute("SELECT id, name, description FROM `role` WHERE id = ?", [id]);
    if (!rows[0]) return res.status(404).json({ error: "Role not found." });
    res.json(rows[0]);
  } catch (err) { next(err); }
});

// DELETE /admin/roles/:id
app.delete("/admin/roles/:id", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const { id } = req.params;
    const [users] = await pool.execute("SELECT id FROM `users` WHERE role_id=? LIMIT 1", [id]);
    if (users.length > 0) return res.status(409).json({ error: "Cannot delete role — users are still assigned to it." });
    const [rows] = await pool.execute("SELECT id FROM `role` WHERE id=?", [id]);
    if (!rows[0]) return res.status(404).json({ error: "Role not found." });
    await pool.execute("DELETE FROM `role` WHERE id=?", [id]);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// PUT /admin/roles/:id/permissions — replace all permissions
app.put("/admin/roles/:id/permissions", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const { id } = req.params;
    const permissionKeys = Array.isArray(req.body?.permissionKeys) ? req.body.permissionKeys : [];
    const [role] = await pool.execute("SELECT id FROM `role` WHERE id=?", [id]);
    if (!role[0]) return res.status(404).json({ error: "Role not found." });

    await pool.execute("DELETE FROM `role_permission` WHERE role_id=?", [id]);
    if (permissionKeys.length > 0) {
      const placeholders = permissionKeys.map(() => "?").join(",");
      const [perms] = await pool.execute(
        `SELECT id FROM \`permission\` WHERE \`key\` IN (${placeholders})`,
        permissionKeys
      );
      for (const p of perms) {
        await pool.execute(
          "INSERT IGNORE INTO `role_permission` (role_id, permission_id) VALUES (?, ?)",
          [id, p.id]
        );
      }
    }
    res.json({ ok: true, permissionKeys });
  } catch (err) { next(err); }
});

// PATCH /admin/users/:userId/role — assign role to user
app.patch("/admin/users/:userId/role", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const { userId } = req.params;
    if (!UUID_RE.test(userId)) return res.status(400).json({ error: "Invalid ID format." });
    const roleId = req.body?.roleId ?? null;
    if (roleId) {
      if (isNaN(parseInt(roleId, 10))) return res.status(400).json({ error: "Invalid role ID." });
      const [rows] = await pool.execute("SELECT id FROM `role` WHERE id=?", [parseInt(roleId, 10)]);
      if (!rows[0]) return res.status(404).json({ error: "Role not found." });
    }
    await pool.execute("UPDATE `users` SET role_id=? WHERE id=?", [roleId, userId]);
    const [rows] = await pool.execute("SELECT id, username, role_id FROM `users` WHERE id=?", [userId]);
    if (!rows[0]) return res.status(404).json({ error: "User not found." });
    res.json(rows[0]);
    broadcast("refresh");
  } catch (err) { next(err); }
});

// PATCH /admin/users/:userId/link-employee — link user to employee
app.patch("/admin/users/:userId/link-employee", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const { userId } = req.params;
    if (!UUID_RE.test(userId)) return res.status(400).json({ error: "Invalid ID format." });
    const raw = req.body?.employeeId ?? null;
    let employeeId = null;
    if (raw) {
      // Accept either internal 7-digit id OR employee_number
      const [byId]  = await pool.execute("SELECT id FROM `employee` WHERE id=? LIMIT 1", [raw]);
      const [byNum] = await pool.execute("SELECT id FROM `employee` WHERE employee_number=? LIMIT 1", [raw]);
      const found   = byId[0] ?? byNum[0];
      if (!found) return res.status(404).json({ error: `Employee "${raw}" not found.` });
      employeeId = found.id;
    }
    await pool.execute("UPDATE `users` SET employee_id=? WHERE id=?", [employeeId, userId]);
    const [rows] = await pool.execute("SELECT id, username, employee_id FROM `users` WHERE id=?", [userId]);
    res.json(rows[0]);
    broadcast("refresh");
  } catch (err) { next(err); }
});

// GET /admin/user-employee-links — list all users with their linked employee
app.get("/admin/user-employee-links", requireAuth, requireAdmin, async (req, res, next) => {
  try {
    const q     = (req.query.q ?? "").trim();
    const limit = Math.min(parseInt(req.query.limit ?? (q ? "50" : "5"), 10), 500);
    const like  = `%${q}%`;
    // mysql2 prepared statements (pool.execute) reject a bound LIMIT parameter
    // ("Incorrect arguments to mysqld_stmt_execute") — inline the sanitized integer instead.
    const [rows] = await pool.execute(`
      SELECT u.id, u.username, u.status, u.employee_id,
             e.name AS employee_name,
             COALESCE(p.position_title, '') AS position,
             COALESCE(d.name, '')           AS office_division
      FROM \`users\` u
      LEFT JOIN \`employee\`         e  ON e.id  = u.employee_id
      LEFT JOIN \`position\`         p  ON p.id  = e.position_id
      LEFT JOIN \`division\`         d  ON d.id  = e.division_id
      ${q ? "WHERE u.username LIKE ? OR e.name LIKE ?" : ""}
      ORDER BY (u.employee_id IS NOT NULL) DESC, u.username ASC
      LIMIT ${limit}
    `, q ? [like, like] : []);
    res.json(rows);
  } catch (err) { next(err); }
});

// GET /me/dashboard — full employee record for the Employee Dashboard (only own data)
app.get("/me/dashboard", requireAuth, async (req, res, next) => {
  try {
    const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [req.user.sub]);
    if (!userRow?.employee_id) return res.json(null);
    const emp = await getEmployeeById(userRow.employee_id);
    res.json(emp);
  } catch (err) { next(err); }
});

// GET /me/suspension — leave/suspension history for the linked employee
app.get("/me/suspension", requireAuth, async (req, res, next) => {
  try {
    const [[userRow]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [req.user.sub]);
    if (!userRow?.employee_id) return res.json([]);
    const [rows] = await pool.execute(
      `SELECT
              CASE reason WHEN 1 THEN 'On Leave' WHEN 2 THEN 'Suspended' ELSE 'Unknown' END AS reason,
              start_date, end_date,
              TIMESTAMPDIFF(HOUR,   start_date, COALESCE(end_date, NOW())) AS total_hours,
              TIMESTAMPDIFF(MINUTE, start_date, COALESCE(end_date, NOW())) % 60 AS extra_minutes
       FROM \`employee_suspension\`
       WHERE employee_id = ?
       ORDER BY start_date DESC`,
      [userRow.employee_id]
    );
    res.json(rows);
  } catch (err) { next(err); }
});

// ── 201 File Routes ─────────────────────────────────────────────────────────

const upload201 = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter(_req, file, cb) {
    const allowed = ["application/pdf", "image/jpeg", "image/png", "image/webp"];
    if (allowed.includes(file.mimetype)) cb(null, true);
    else cb(Object.assign(new Error("Only PDF, JPG, PNG, or WebP files are accepted."), { status: 400 }));
  },
});

const VALID_201_ITEMS = new Set("ABCDEFGHIJKLMNOPQ".split(""));

// Files are stored as LONGBLOB, so a single upload must fit inside one MySQL
// packet. max_allowed_packet is the real ceiling — multer's 10 MB is not.
let MAX_201_BYTES = 768 * 1024;
try {
  const [[pv]] = await pool.query("SELECT @@max_allowed_packet AS pkt");
  MAX_201_BYTES = Math.max(256 * 1024, Math.min(10 * 1024 * 1024, Number(pv.pkt) - 256 * 1024));
  logger.info({ maxAllowedPacket: Number(pv.pkt), max201Bytes: MAX_201_BYTES }, "201 file upload ceiling");
  if (Number(pv.pkt) < 8 * 1024 * 1024)
    logger.warn(`max_allowed_packet is ${(Number(pv.pkt) / 1048576).toFixed(1)} MB — scanned 201 documents larger than ${(MAX_201_BYTES / 1048576).toFixed(1)} MB will be rejected.`);
} catch { /* keep the conservative default */ }

const MB = b => (b / 1048576).toFixed(1);

// Re-read the ceiling periodically so raising max_allowed_packet on the server
// takes effect without restarting the API.
let maxPacketCheckedAt = Date.now();
async function current201Limit() {
  if (Date.now() - maxPacketCheckedAt > 60_000) {
    maxPacketCheckedAt = Date.now();
    try {
      const [[pv]] = await pool.query("SELECT @@max_allowed_packet AS pkt");
      MAX_201_BYTES = Math.max(256 * 1024, Math.min(10 * 1024 * 1024, Number(pv.pkt) - 256 * 1024));
    } catch { /* keep the last known value */ }
  }
  return MAX_201_BYTES;
}

async function reject201IfTooLarge(req, res) {
  const limit = await current201Limit();
  if (req.file.size > limit) {
    res.status(413).json({ error: `File is ${MB(req.file.size)} MB. The database accepts up to ${MB(limit)} MB per file — please compress it or scan at a lower resolution.` });
    return true;
  }
  return false;
}

function is201SizeError(err) {
  return err?.code === "ER_NET_PACKET_TOO_LARGE" || err?.code === "ER_TOO_BIG_ROWSIZE" || /max_allowed_packet/i.test(err?.message ?? "");
}

async function getMyEmpId(userId) {
  const [[row]] = await pool.execute("SELECT employee_id FROM `users` WHERE id=?", [userId]);
  return row?.employee_id ?? null;
}

app.get("/me/201-files", requireAuth, async (req, res, next) => {
  try {
    const empId = await getMyEmpId(req.user.sub);
    if (!empId) return res.json([]);
    const rows = await list201Files(empId);
    res.json(rows.map(r => ({ id: r.id, docItem: r.doc_item, filename: r.filename, mimeType: r.mime_type, uploadedAt: r.uploaded_at })));
  } catch (err) { next(err); }
});

app.post("/me/201-files", requireAuth, upload201.single("file"), async (req, res, next) => {
  try {
    const empId = await getMyEmpId(req.user.sub);
    if (!empId) return res.status(403).json({ error: "Your account is not linked to an employee record." });
    const { docItem } = req.body ?? {};
    if (!docItem || !VALID_201_ITEMS.has(docItem.toUpperCase()))
      return res.status(400).json({ error: "Invalid document item (A–Q)." });
    if (!req.file) return res.status(400).json({ error: "No file uploaded." });
    if (await reject201IfTooLarge(req, res)) return;
    const safe = req.file.originalname.replace(/[^a-zA-Z0-9._\- ]/g, "_").slice(0, 200);
    const insertedId = await insert201File(empId, docItem.toUpperCase(), safe, req.file.mimetype, req.file.buffer);
    res.status(201).json({ id: insertedId, docItem: docItem.toUpperCase(), filename: safe, mimeType: req.file.mimetype });
  } catch (err) {
    if (is201SizeError(err)) return res.status(413).json({ error: `File is too large for the database (limit ${MB(MAX_201_BYTES)} MB). Please upload a smaller file.` });
    next(err);
  }
});

app.get("/me/201-files/:fileId/view", requireAuth, async (req, res, next) => {
  try {
    const empId = await getMyEmpId(req.user.sub);
    if (!empId) return res.status(403).json({ error: "Access denied." });
    const fileId = parseInt(req.params.fileId, 10);
    if (!Number.isInteger(fileId)) return res.status(400).json({ error: "Invalid file ID." });
    const [[fileRow]] = await pool.execute("SELECT employee_id FROM `employee_201_files` WHERE id=?", [fileId]);
    if (!fileRow || fileRow.employee_id !== empId) return res.status(403).json({ error: "Access denied." });
    const row = await get201FileData(fileId);
    if (!row) return res.status(404).json({ error: "File not found." });
    res.setHeader("Content-Type", row.mime_type ?? "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${row.filename}"`);
    // The global helmet CSP sets frame-ancestors 'none', which stops even our own
    // app from showing a PDF in the in-page viewer. Allow same-origin framing here
    // only; everything else stays locked down.
    res.setHeader("Content-Security-Policy", "default-src 'none'; img-src 'self' data: blob:; object-src 'none'; frame-ancestors 'self'");
    res.setHeader("X-Frame-Options", "SAMEORIGIN");
    res.setHeader("Cache-Control", "private, max-age=60");
    res.send(row.file_data);
  } catch (err) { next(err); }
});

app.delete("/me/201-files/:fileId", requireAuth, async (req, res, next) => {
  try {
    const empId = await getMyEmpId(req.user.sub);
    if (!empId) return res.status(403).json({ error: "Access denied." });
    const fileId = parseInt(req.params.fileId, 10);
    if (!Number.isInteger(fileId)) return res.status(400).json({ error: "Invalid file ID." });
    const [[fileRow]] = await pool.execute("SELECT employee_id FROM `employee_201_files` WHERE id=?", [fileId]);
    if (!fileRow || fileRow.employee_id !== empId) return res.status(403).json({ error: "Access denied." });
    const deleted = await delete201File(fileId);
    if (!deleted) return res.status(404).json({ error: "File not found." });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.get("/employees/:id/201-files", requireAuth, async (req, res, next) => {
  try {
    const empId = parseInt(req.params.id, 10);
    if (!Number.isInteger(empId)) return res.status(400).json({ error: "Invalid employee ID." });
    const rows = await list201Files(empId);
    res.json(rows.map(r => ({ id: r.id, docItem: r.doc_item, filename: r.filename, mimeType: r.mime_type, uploadedAt: r.uploaded_at })));
  } catch (err) { next(err); }
});

app.post("/employees/:id/201-files", requireAuth, upload201.single("file"), async (req, res, next) => {
  try {
    const empId = parseInt(req.params.id, 10);
    if (!Number.isInteger(empId)) return res.status(400).json({ error: "Invalid employee ID." });
    const { docItem } = req.body ?? {};
    if (!docItem || !VALID_201_ITEMS.has(docItem.toUpperCase()))
      return res.status(400).json({ error: "Invalid document item (A–Q)." });
    if (!req.file) return res.status(400).json({ error: "No file uploaded." });
    if (await reject201IfTooLarge(req, res)) return;
    const safe = req.file.originalname.replace(/[^a-zA-Z0-9._\- ]/g, "_").slice(0, 200);
    const insertedId = await insert201File(empId, docItem.toUpperCase(), safe, req.file.mimetype, req.file.buffer);
    res.status(201).json({ id: insertedId, docItem: docItem.toUpperCase(), filename: safe, mimeType: req.file.mimetype });
  } catch (err) {
    if (is201SizeError(err)) return res.status(413).json({ error: `File is too large for the database (limit ${MB(MAX_201_BYTES)} MB). Please upload a smaller file.` });
    next(err);
  }
});

app.get("/employees/:id/201-files/:fileId/view", requireAuth, async (req, res, next) => {
  try {
    const fileId = parseInt(req.params.fileId, 10);
    if (!Number.isInteger(fileId)) return res.status(400).json({ error: "Invalid file ID." });
    const row = await get201FileData(fileId);
    if (!row) return res.status(404).json({ error: "File not found." });
    res.setHeader("Content-Type", row.mime_type ?? "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${row.filename}"`);
    // The global helmet CSP sets frame-ancestors 'none', which stops even our own
    // app from showing a PDF in the in-page viewer. Allow same-origin framing here
    // only; everything else stays locked down.
    res.setHeader("Content-Security-Policy", "default-src 'none'; img-src 'self' data: blob:; object-src 'none'; frame-ancestors 'self'");
    res.setHeader("X-Frame-Options", "SAMEORIGIN");
    res.setHeader("Cache-Control", "private, max-age=60");
    res.send(row.file_data);
  } catch (err) { next(err); }
});

app.delete("/employees/:id/201-files/:fileId", requireAuth, async (req, res, next) => {
  try {
    const fileId = parseInt(req.params.fileId, 10);
    if (!Number.isInteger(fileId)) return res.status(400).json({ error: "Invalid file ID." });
    const deleted = await delete201File(fileId);
    if (!deleted) return res.status(404).json({ error: "File not found." });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// ── LEAVE APPLICATIONS (CS Form No. 6) ────────────────────────────────────────
const LEAVE_UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Section 7 fields each approver may write — single source for both PATCH and action.
const LEAVE_S7_KEYS = {
  hr:       ["credits_as_of","vl_total_earned","vl_less_application","vl_balance","sl_total_earned","sl_less_application","sl_balance"],
  crasd:    ["recommendation","disapproval_reason"],
  director: ["approved_days_with_pay","approved_days_without_pay","approved_others_specify","disapproved_due_to"],
};
LEAVE_S7_KEYS.admin = [...LEAVE_S7_KEYS.hr, ...LEAVE_S7_KEYS.crasd, ...LEAVE_S7_KEYS.director];

// Each workflow action: the permission it needs and the workflow_status it must follow.
const LEAVE_ACTIONS = {
  hr_approved:          { role: "hr",       perm: "leave.hr",       from: "pending" },
  crasd_approved:       { role: "crasd",    perm: "leave.crasd",    from: "hr_approved" },
  director_approved:    { role: "director", perm: "leave.director", from: "crasd_approved" },
  director_disapproved: { role: "director", perm: "leave.director", from: "crasd_approved" },
};

async function leaveActor(req) {
  const isSuper = req.user.username.toLowerCase() === ADMIN_USERNAME;
  const perms = isSuper ? [] : (req.userPermissions ??= await loadUserPermissions(req.user.sub));
  const has = (p) => isSuper || perms.includes(p);
  const [[emp]] = await pool.execute(
    `SELECT e.id AS employee_id, e.name, COALESCE(e.division_id, p.division_id) AS division_id
     FROM \`users\` u
     LEFT JOIN \`employee\` e ON e.id = u.employee_id
     LEFT JOIN \`position\` p ON p.id = e.position_id
     WHERE u.id = ?`,
    [req.user.sub]
  );
  return {
    username: req.user.username.toLowerCase(),
    employeeId: emp?.employee_id ?? null,
    divisionId: emp?.division_id ?? null,
    full_name: emp?.name ?? "",
    isAdmin: isSuper || perms.includes("leave.manage"),
    perms, has,
  };
}

// Applicants see their own; approvers see what is routed to their division
// (see canSeeLeaveRecord); leave.manage sees everything.
const canSeeLeave = async (actor, record) => canSeeLeaveRecord(actor, record, await listLeaveSignatories());

// Pop-ups along the chain: the next officer gets "waiting for you"; the
// applicant gets progress and the final result. Never fails the request.
const LEAVE_NEXT_ROLE = { pending: "hr", hr_approved: "crasd", crasd_approved: "director" };
const LEAVE_NEXT_TASK = { hr: "HR certification", crasd: "recommendation", director: "approval" };
async function notifyLeave(record, applicant = null) {
  try {
    const who  = [record.last_name, record.first_name].filter(Boolean).join(", ") || record.submitted_by;
    const what = `${record.leave_type || "Leave"}${record.num_working_days ? `, ${record.num_working_days} day(s)` : ""}`;
    await clearLeaveActionNotifications(record.id);
    const role = LEAVE_NEXT_ROLE[record.workflow_status];
    if (role) {
      const sigs = await listLeaveSignatories();
      const next = (await listLeaveApprovers(`leave.${role}`))
        .filter(u => u.username !== ADMIN_USERNAME && isLeaveRecipient(u, record, role, sigs));
      await addLeaveNotifications(next.map(u => u.id), record.id, "action",
        `${who}'s leave (${what}) is waiting for your ${LEAVE_NEXT_TASK[role]}.`);
    }
    if (applicant) {
      await addLeaveNotifications(await listLeaveOwnerAccounts(record), record.id, applicant.kind, applicant.message);
    }
    broadcast("leave-notify");
  } catch (err) { logger.warn({ err }, "leave notification failed"); }
}

// What the applicant is told after each step; `by` is the name printed on the form.
function applicantUpdate(action, record, by) {
  const what = `${record.leave_type || "leave"}${record.num_working_days ? `, ${record.num_working_days} day(s)` : ""}`;
  switch (action) {
    case "hr_approved":    return { kind: "info", message: `${by} (HR) certified your leave application (${what}). It is now with the Chief for recommendation.` };
    case "crasd_approved": return { kind: "info", message: `${by} recommended your leave application (${what}). It is now with the Regional Director for approval.` };
    case "director_approved":    return { kind: "approved", message: `Your leave application (${what}) was APPROVED by ${by}.` };
    case "director_disapproved": return { kind: "disapproved", message: `Your leave application (${what}) was DISAPPROVED by ${by}.${record.disapproved_due_to ? ` Reason: ${record.disapproved_due_to}` : ""}` };
    default: return null;
  }
}

app.get("/leave-notifications", requireAuth, async (req, res, next) => {
  try { res.json(await listLeaveNotifications(req.user.sub)); } catch (err) { next(err); }
});
app.post("/leave-notifications/read-all", requireAuth, async (req, res, next) => {
  try { await markLeaveNotificationsRead(req.user.sub); res.json({ ok: true }); } catch (err) { next(err); }
});
app.post("/leave-notifications/:id/read", requireAuth, async (req, res, next) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ error: "Invalid ID." });
    await markLeaveNotificationsRead(req.user.sub, id);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.get("/leave-applications", requireAuth, async (req, res, next) => {
  try {
    const actor = await leaveActor(req);
    res.json(await listLeaveApplicationsByUser(actor));
  } catch (err) { next(err); }
});

app.get("/leave-applications/snapshot", requireAuth, async (req, res, next) => {
  try {
    const snap = await getLeaveSnapshot(req.user.sub);
    if (!snap) return res.status(404).json({ error: "Your account is not linked to an employee record." });
    res.json(snap);
  } catch (err) { next(err); }
});

app.post("/leave-applications", ...requirePermission("leave.file"), async (req, res, next) => {
  try {
    const b = req.body ?? {};
    const snap = await getLeaveSnapshot(req.user.sub);
    if (!snap) return res.status(400).json({ error: "Your account is not linked to an employee record, so a leave application cannot be filed." });
    const str = (k) => String(b[k] ?? "");
    const data = {
      ...snap,
      submitted_by: req.user.username.toLowerCase(),
      date_of_filing: str("date_of_filing"),
      leave_type: str("leave_type"),
      leave_type_other: str("leave_type_other"),
      details_vacation_local: str("details_vacation_local"),
      details_vacation_abroad: str("details_vacation_abroad"),
      details_sick_in_hospital: str("details_sick_in_hospital"),
      details_sick_out_patient: str("details_sick_out_patient"),
      details_special_women_illness: str("details_special_women_illness"),
      details_study_masters: !!b.details_study_masters,
      details_study_bar: !!b.details_study_bar,
      details_study_other: str("details_study_other"),
      details_monetization: !!b.details_monetization,
      details_terminal: !!b.details_terminal,
      num_working_days: str("num_working_days"),
      inclusive_dates: str("inclusive_dates"),
      commutation: str("commutation"),
      applicant_signature: b.applicant_signature ? String(b.applicant_signature) : null,
    };
    if (!data.leave_type.trim() || !data.num_working_days.trim())
      return res.status(400).json({ error: "Type of leave and number of working days are required." });
    const types = data.leave_type.split(",").map(s => s.trim());
    if (types.includes("Sick Leave") && !data.details_sick_in_hospital.trim() && !data.details_sick_out_patient.trim())
      return res.status(400).json({ error: "For Sick Leave, tick In Hospital or Out Patient and specify the illness (6.B)." });

    // Fixed at filing: the division decides which chief / director receives it;
    // the employee record keeps the leave in that person's history.
    data.applicant_division_id = snap.division_id ?? null;
    data.employee_id = snap.employee_id ?? null;

    const actor = await leaveActor(req);
    const created = await createLeaveApplication(data, actor.username, "employee", actor.full_name);
    insertAuditLog("leave_submit", `${snap.last_name}, ${snap.first_name} - ${data.leave_type}`, actor.username, null,
      { leave_id: created.id, leave_type: data.leave_type, num_working_days: data.num_working_days, inclusive_dates: data.inclusive_dates })
      .catch(e => logger.warn({ err: e }, "audit log insert failed"));
    notifyLeave(created); // HR: "waiting for your HR certification"
    res.status(201).json(created);
  } catch (err) { next(err); }
});

app.get("/leave-applications/:id", requireAuth, async (req, res, next) => {
  try {
    if (!LEAVE_UUID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID." });
    const record = await getLeaveApplication(req.params.id);
    if (!record) return res.status(404).json({ error: "Not found." });
    const actor = await leaveActor(req);
    if (!(await canSeeLeave(actor, record))) return res.status(403).json({ error: "Access denied." });
    res.json({ ...record, mine: isLeaveOwner(actor, record) });
  } catch (err) { next(err); }
});

app.get("/leave-applications/:id/history", requireAuth, async (req, res, next) => {
  try {
    if (!LEAVE_UUID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID." });
    const record = await getLeaveApplication(req.params.id);
    if (!record) return res.status(404).json({ error: "Not found." });
    if (!(await canSeeLeave(await leaveActor(req), record))) return res.status(403).json({ error: "Access denied." });
    const rows = await getLeaveHistory(req.params.id);
    // Signatures live in details; the timeline only needs who/what/when.
    res.json(rows.map(({ details, ...r }) => r));
  } catch (err) { next(err); }
});

// Out-of-band Section 7 correction — admins only (approvers write via /action).
app.patch("/leave-applications/:id", requireAuth, async (req, res, next) => {
  try {
    if (!LEAVE_UUID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID." });
    const actor = await leaveActor(req);
    if (!actor.isAdmin) return res.status(403).json({ error: "Access denied." });
    const updates = {};
    for (const k of LEAVE_S7_KEYS.admin) if (k in (req.body ?? {})) updates[k] = String(req.body[k] ?? "");
    const before = await getLeaveApplication(req.params.id);
    if (!before) return res.status(404).json({ error: "Not found." });
    const record = await updateLeaveSection7(req.params.id, updates, { username: actor.username, role: "admin", full_name: actor.full_name });
    if (!record) return res.status(400).json({ error: "No Section 7 fields to update." });
    insertAuditLog("leave_section7", `${before.last_name}, ${before.first_name} - ${before.leave_type}`, actor.username,
      Object.fromEntries(Object.keys(updates).map(k => [k, before[k]])), updates)
      .catch(e => logger.warn({ err: e }, "audit log insert failed"));
    res.json(record);
  } catch (err) { next(err); }
});

app.post("/leave-applications/:id/action", requireAuth, async (req, res, next) => {
  try {
    if (!LEAVE_UUID_RE.test(req.params.id)) return res.status(400).json({ error: "Invalid ID." });
    const { action, signature, signatory_id, ...sectionData } = req.body ?? {};
    const spec = LEAVE_ACTIONS[action];
    if (!spec) return res.status(400).json({ error: "Invalid action." });

    // Signing is personal: only the officer holding this stage's permission may
    // sign, at that stage, and never on their own leave. The superadmin is a
    // system account and leave.manage is not a signing role, so neither bypasses this.
    if (req.user.username.toLowerCase() === ADMIN_USERNAME)
      return res.status(403).json({ error: "The superadmin account cannot sign leave applications. The HR officer, CRASD chief or Regional Director must sign with their own account." });
    const actor = await leaveActor(req);
    if (!actor.perms.includes(spec.perm))
      return res.status(403).json({ error: "You do not have permission to perform this action." });

    const current = await getLeaveApplication(req.params.id);
    if (!current) return res.status(404).json({ error: "Not found." });
    if (current.workflow_status !== spec.from)
      return res.status(400).json({ error: `Application is not at the stage for this action (currently: ${current.workflow_status}).` });
    // Own leave = same login or same employee record (a second account can't approve it).
    if (isLeaveOwner(actor, current))
      return res.status(403).json({ error: "You cannot review your own leave application." });
    if (typeof signature !== "string" || !/^data:image\/(png|jpeg|webp);base64,/.test(signature))
      return res.status(400).json({ error: "Please sign your section before submitting." });

    // Only an officer of the division this application is routed to may sign,
    // and only as one of their own names.
    const sigs = await listLeaveSignatories();
    if (!isLeaveRecipient(actor, current, spec.role, sigs))
      return res.status(403).json({ error: "This application is assigned to another division's officer for this section." });
    const mine = leaveSignatoriesFor(actor, current, spec.role, sigs);
    let signatory;
    if (signatory_id != null && signatory_id !== "") {
      signatory = mine.find(s => s.id === Number(signatory_id));
      if (!signatory) return res.status(400).json({ error: "Pick a valid name to sign as." });
    } else {
      // Prefer the name set to their division over an "All divisions" one.
      signatory = mine.find(s => s.division_id != null) ?? mine[0];
    }
    if (!signatory) return res.status(400).json({ error: "No name is set up for this section yet. Ask the administrator to add one in Settings → Leave Signatories." });

    const updates = {};
    for (const k of LEAVE_S7_KEYS[spec.role]) if (k in sectionData) updates[k] = String(sectionData[k] ?? "");

    const record = await processLeaveAction(req.params.id,
      { username: actor.username, role: spec.role, full_name: actor.full_name },
      action, updates, signature, { name: signatory.name, designation: signatory.designation });
    insertAuditLog("leave_action", `${current.last_name}, ${current.first_name} - ${current.leave_type}`, actor.username,
      { workflow_status: current.workflow_status },
      { action, workflow_status: record.workflow_status, signed_as: `${signatory.name} (${signatory.designation})`, ...updates })
      .catch(e => logger.warn({ err: e }, "audit log insert failed"));
    notifyLeave(record, applicantUpdate(action, record, signatory.name)); // next officer + the applicant
    res.json(record);
  } catch (err) { next(err); }
});

// No delete route: a filed CS Form No. 6 and its signed history are kept for the audit trail.

// ── LEAVE SIGNATORIES — names printed under the Section 7 signature lines ────
// Anyone signed in may read them (they print on every form); only leave.manage edits.
// division_id (null = all divisions) routes each division's applications to
// the officer of that division.
async function readLeaveSignatory(body) {
  const name = String(body?.name ?? "").trim();
  const designation = String(body?.designation ?? "").trim();
  if (!name || !designation) return { error: "Enter both the name and the position." };
  if (name.length > 150 || designation.length > 200) return { error: "The name or position is too long." };

  let division_id = null;
  if (body?.division_id != null && body.division_id !== "") {
    division_id = Number(body.division_id);
    const [[d]] = Number.isInteger(division_id) && division_id > 0
      ? await pool.execute("SELECT id FROM `division` WHERE id=?", [division_id]) : [[]];
    if (!d) return { error: "That division doesn't exist." };
  }
  return { name, designation, division_id };
}
const signatoryAuditView = (r) => ({ name: r.name, designation: r.designation, division: r.division_name ?? "All divisions" });
const leaveSignatoryId = (req) => {
  const id = Number(req.params.id);
  return Number.isInteger(id) && id > 0 ? id : null;
};
const auditSignatory = (action, row, user, oldV, newV) =>
  insertAuditLog(action, `${row.role}: ${row.name}`, user, oldV, newV)
    .catch(e => logger.warn({ err: e }, "audit log insert failed"));

// `mine` marks the names that belong to the caller (their division or "All
// divisions"), so the leave form can offer only those when they sign.
app.get("/leave-signatories", requireAuth, async (req, res, next) => {
  try {
    const actor = await leaveActor(req);
    const rows = await listLeaveSignatories();
    res.json(rows.map(s => ({ ...s, mine: isOwnLeaveSignatory(s, actor.divisionId) })));
  } catch (err) { next(err); }
});

// Division choices for the Settings editor.
app.get("/leave-signatories/options", ...requirePermission("leave.manage"), async (req, res, next) => {
  try { res.json({ divisions: await getDivisions() }); } catch (err) { next(err); }
});

app.post("/leave-signatories", ...requirePermission("leave.manage"), async (req, res, next) => {
  try {
    const role = String(req.body?.role ?? "");
    if (!LEAVE_SIGNATORY_ROLES.includes(role)) return res.status(400).json({ error: "Invalid section." });
    const v = await readLeaveSignatory(req.body);
    if (v.error) return res.status(400).json({ error: v.error });
    const row = await createLeaveSignatory(role, v);
    auditSignatory("leave_signatory_create", row, req.user.username, null, signatoryAuditView(row));
    res.status(201).json(row);
  } catch (err) { next(err); }
});

app.put("/leave-signatories/:id", ...requirePermission("leave.manage"), async (req, res, next) => {
  try {
    const id = leaveSignatoryId(req);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const before = await getLeaveSignatory(id);
    if (!before) return res.status(404).json({ error: "Not found." });
    const v = await readLeaveSignatory(req.body);
    if (v.error) return res.status(400).json({ error: v.error });
    const row = await updateLeaveSignatory(id, v);
    auditSignatory("leave_signatory_update", row, req.user.username, signatoryAuditView(before), signatoryAuditView(row));
    res.json(row);
  } catch (err) { next(err); }
});

app.delete("/leave-signatories/:id", ...requirePermission("leave.manage"), async (req, res, next) => {
  try {
    const id = leaveSignatoryId(req);
    if (!id) return res.status(400).json({ error: "Invalid ID." });
    const row = await deleteLeaveSignatory(id);
    if (!row) return res.status(404).json({ error: "Not found." });
    if (row === "last") return res.status(400).json({ error: "Each section needs at least one name. Add the new person first, then remove this one." });
    auditSignatory("leave_signatory_delete", row, req.user.username, { name: row.name, designation: row.designation }, null);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.use((req, res) => res.status(404).json({ error: "Not found." }));

// ── ERROR HANDLER ─────────────────────────────────────────────────────────────
app.use((err, req, res, _next) => {
  const log = req.log ?? logger;
  log.error({ err, reqId: req.id }, "Unhandled request error");
  res.status(err.status ?? 500).json({
    error: err.status ? err.message : "Internal server error."
  });
});

// ── STARTUP ───────────────────────────────────────────────────────────────────
await ensureSchema();
await ensureLeaveSchema();
if (process.env.REDIS_URL) {
  try { await redis.connect(); logger.info("Redis connected."); }
  catch { logger.warn("Redis unavailable — caching, rate-limiting, and JWT blacklist disabled."); }
} else {
  logger.warn("No REDIS_URL set — Redis features disabled (caching, rate-limiting, JWT blacklist).");
}
if (process.env.MINIO_ENDPOINT) {
  try { await ensureBucket(); }
  catch { logger.warn("MinIO unavailable — avatar uploads disabled."); }
} else {
  logger.warn("No MINIO_ENDPOINT set — avatar uploads disabled.");
}

// Auto-activate employees whose suspension/leave end_date has passed
async function autoActivateOverdueSuspensions() {
  try {
    const [rows] = await pool.execute(
      `SELECT DISTINCT es.employee_id
       FROM \`employee_suspension\` es
       JOIN \`employee\` e ON e.id = es.employee_id
       WHERE es.end_date IS NOT NULL
         AND es.end_date < NOW()
         AND e.status = 'inactive'`
    );
    if (rows.length === 0) return;
    for (const row of rows) {
      await activateEmployee(row.employee_id);
      logger.info({ employee_id: row.employee_id }, "auto-activated: suspension/leave end_date passed");
    }
    broadcast("refresh");
  } catch (err) {
    logger.error(err, "auto-activate overdue suspensions failed");
  }
}
setInterval(autoActivateOverdueSuspensions, 60 * 1000);
autoActivateOverdueSuspensions();

// ── MONTHLY LOAN DEDUCTION CRON ──────────────────────────────────────────────
function monthKey(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}
function prevMonthKey() {
  const d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() - 1);
  return monthKey(d);
}

async function doLoanDeductions(key) {
  const [rows] = await pool.execute(
    "SELECT `value` FROM `app_settings` WHERE `key` = 'loan_deduction_last_run' LIMIT 1"
  );
  if (rows[0]?.value === key) return; // already ran for this month
  logger.info({ month: key }, "Running monthly loan deductions");
  try {
    const results  = await processAllLoanDeductions();
    const deducted = results.reduce((s, r) => s + (r.deducted?.length ?? 0), 0);
    const halted   = results.reduce((s, r) => s + (r.halted?.length   ?? 0), 0);
    await pool.execute(
      "INSERT INTO `app_settings` (`key`, `value`) VALUES ('loan_deduction_last_run', ?) ON DUPLICATE KEY UPDATE `value` = ?",
      [key, key]
    );
    logger.info({ month: key, deducted, halted }, "Monthly loan deductions complete");
  } catch (e) {
    logger.error({ err: e }, "Monthly loan deduction error");
  }
}

// Runs at 11:59 PM on days 28–31, fires only on the actual last day of the month
cron.schedule("59 23 28-31 * *", async () => {
  const now = new Date(), tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);
  if (tomorrow.getDate() !== 1) return; // not the last day of month
  await doLoanDeductions(monthKey(now));
});

// On startup: catch up ONLY if last month's deduction was missed
// (never runs current month early — only catches a missed previous month)
doLoanDeductions(prevMonthKey());

// ────────────────────────────────────────────────────────────────────────────

app.listen(port, () => {
  logger.info({ port }, "PSA Payroll API listening (MySQL)");
});


// Thin Redis cache helper
import { redis } from "./queue.js";

export async function cacheGet(key) {
  try {
    const v = await redis.get(key);
    return v ? JSON.parse(v) : null;
  } catch { return null; }
}

export async function cacheSet(key, value, ttlSeconds = 30) {
  try {
    await redis.set(key, JSON.stringify(value), { EX: ttlSeconds });
  } catch { /* ignore cache errors */ }
}

export async function cacheDel(...keys) {
  try {
    if (keys.length) await redis.del(keys);
  } catch { /* ignore */ }
}

// Cache version counter — incrementing this busts all versioned employee cache keys
export async function getVersion() {
  try {
    const v = await redis.get("emp:version");
    return v ?? "0";
  } catch { return "0"; }
}

export async function incrementVersion() {
  try {
    await redis.incr("emp:version");
  } catch { /* ignore */ }
}

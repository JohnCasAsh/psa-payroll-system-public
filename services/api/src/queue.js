import { createClient } from "redis";

const queueName = process.env.JOB_QUEUE_NAME ?? "payroll-jobs";

// When no REDIS_URL is configured (e.g. XAMPP-only dev), use a no-op client
// so caching/rate-limiting silently degrade rather than spamming reconnect errors.
const _noopRedis = new Proxy({}, {
  get: (_, prop) => {
    if (prop === "on")      return () => _noopRedis;
    if (prop === "connect") return async () => {};
    if (prop === "multi")   return () => _noopRedis;
    return async () => null;
  },
});

export const redis = process.env.REDIS_URL
  ? (() => {
      const c = createClient({ url: process.env.REDIS_URL });
      c.on("error", err => console.error("Redis error", err));
      return c;
    })()
  : _noopRedis;

export async function enqueueJob(job) {
  if (process.env.REDIS_URL) {
    await redis.rPush(queueName, JSON.stringify({ id: job.id }));
  }
}

import pg from "pg";
import { createClient } from "redis";
import { processPayload } from "./processor.js";
import logger from "./logger.js";

const { Pool } = pg;
const queueName = process.env.JOB_QUEUE_NAME ?? "payroll-jobs";
const pollMs    = Number(process.env.WORKER_POLL_MS ?? 1000);

const pool = new Pool({ connectionString: process.env.DATABASE_URL });

const redis = createClient({ url: process.env.REDIS_URL });
redis.on("error", (err) => logger.error({ err }, "Redis client error"));

async function handleQueueItem(rawItem) {
  const item = JSON.parse(rawItem);
  logger.info({ jobId: item.id }, "Processing job");
  // Placeholder: extend for async payroll jobs
}

async function run() {
  await redis.connect();
  logger.info({ queue: queueName }, "Worker listening");

  while (true) {
    try {
      const item = await redis.blPop(queueName, pollMs / 1000);
      if (item) await handleQueueItem(item.element);
    } catch (err) {
      logger.error({ err }, "Worker poll error — retrying in 2s");
      await new Promise((r) => setTimeout(r, 2000));
    }
  }
}

async function shutdown(signal) {
  logger.info({ signal }, "Graceful shutdown started");
  await redis.quit().catch(() => {});
  await pool.end().catch(() => {});
  process.exit(0);
}
process.on("SIGTERM", () => shutdown("SIGTERM"));
process.on("SIGINT",  () => shutdown("SIGINT"));

run().catch((err) => {
  logger.fatal({ err }, "Worker fatal error");
  process.exit(1);
});

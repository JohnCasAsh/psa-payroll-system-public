import pino from "pino";

const logger = pino({
  level: process.env.LOG_LEVEL ?? "info",
  formatters: {
    level: (label) => ({ level: label })
  },
  base: { service: "psa-payroll-worker" }
});

export default logger;

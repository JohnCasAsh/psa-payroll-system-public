export function processPayload(payload) {
  const serialized = JSON.stringify(payload);
  const keys = Object.keys(payload);

  return {
    received: payload,
    summary: {
      keyCount: keys.length,
      keys,
      byteSize: Buffer.byteLength(serialized, "utf8"),
      processedAt: new Date().toISOString()
    }
  };
}

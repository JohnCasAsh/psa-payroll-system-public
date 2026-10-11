import * as Minio from "minio";

const BUCKET = process.env.MINIO_BUCKET ?? "psa-payroll-documents";

const MINIO_AVAILABLE = !!(process.env.MINIO_ACCESS_KEY && process.env.MINIO_SECRET_KEY);

export const minio = MINIO_AVAILABLE ? new Minio.Client({
  endPoint:  process.env.MINIO_ENDPOINT ?? "minio",
  port:      Number(process.env.MINIO_PORT ?? 9000),
  useSSL:    process.env.MINIO_USE_SSL === "true",
  accessKey: process.env.MINIO_ACCESS_KEY,
  secretKey: process.env.MINIO_SECRET_KEY
}) : null;

export async function ensureBucket() {
  if (!minio) return;
  try {
    const exists = await minio.bucketExists(BUCKET);
    if (!exists) {
      await minio.makeBucket(BUCKET, "us-east-1");
      const policy = JSON.stringify({
        Version: "2012-10-17",
        Statement: [{
          Effect: "Deny",
          Principal: "*",
          Action: ["s3:GetObject"],
          Resource: [`arn:aws:s3:::${BUCKET}/*`],
          Condition: {
            StringNotEquals: {
              "s3:signatureversion": ["AWS4-HMAC-SHA256"]
            }
          }
        }]
      });
      await minio.setBucketPolicy(BUCKET, policy);
    }
  } catch {
    console.warn("MinIO not available — file storage disabled.");
  }
}

export async function uploadObject(objectKey, buffer, sizeBytes, contentType = "application/pdf") {
  if (!minio) throw new Error("File storage is not configured.");
  await minio.putObject(BUCKET, objectKey, buffer, sizeBytes, {
    "Content-Type": contentType
  });
}

export async function getPresignedUrl(objectKey, expirySeconds = 60) {
  if (!minio) throw new Error("File storage is not configured.");
  const url = await minio.presignedGetObject(BUCKET, objectKey, expirySeconds);
  const internalHost   = `${process.env.MINIO_ENDPOINT ?? "minio"}:${process.env.MINIO_PORT ?? 9000}`;
  const publicEndpoint = process.env.MINIO_PUBLIC_ENDPOINT ?? "localhost:9002";
  return url.replace(internalHost, publicEndpoint);
}

export async function deleteObject(objectKey) {
  try {
    await minio.removeObject(BUCKET, objectKey);
  } catch {
    // Ignore if already deleted — idempotent
  }
}

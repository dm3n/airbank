import { Storage } from '@google-cloud/storage'

function getStorage(): Storage {
  const credentialsJson = process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON
  if (!credentialsJson) throw new Error('GOOGLE_APPLICATION_CREDENTIALS_JSON not set')
  const credentials = JSON.parse(credentialsJson)
  return new Storage({ credentials, projectId: process.env.GOOGLE_CLOUD_PROJECT })
}

/**
 * Upload a buffer to GCS and return the gs:// URI.
 */
export async function uploadToGCS(
  buffer: Buffer,
  gcsPath: string,
  contentType: string
): Promise<string> {
  const bucket = process.env.GCS_BUCKET_NAME!
  if (!bucket) throw new Error('GCS_BUCKET_NAME not set')

  const storage = getStorage()
  const file = storage.bucket(bucket).file(gcsPath)

  await file.save(buffer, {
    contentType,
    resumable: false,
    metadata: { cacheControl: 'no-cache' },
  })

  return `gs://${bucket}/${gcsPath}`
}

/**
 * Delete a file from GCS.
 */
export async function deleteFromGCS(gcsUri: string): Promise<void> {
  const storage = getStorage()
  const withoutScheme = gcsUri.replace('gs://', '')
  const slashIdx = withoutScheme.indexOf('/')
  const bucketName = withoutScheme.slice(0, slashIdx)
  const filePath = withoutScheme.slice(slashIdx + 1)
  await storage.bucket(bucketName).file(filePath).delete({ ignoreNotFound: true })
}

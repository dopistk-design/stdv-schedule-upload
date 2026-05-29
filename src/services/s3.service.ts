import { config } from '../config'

const SERVICE_NAME = 'stdv-schedule'
const ENV = config.NODE_ENV === 'production' ? 'prd' : 'dev'

function buildKey(path: string): string {
  return `${SERVICE_NAME}/${ENV}/${path}`
}

async function getPresignedUrl(key: string, action: 'PUT_OBJECT' | 'GET_OBJECT' | 'DELETE_OBJECT'): Promise<string> {
  const response = await fetch(`${config.S3_API_BASE_URL}/sign`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': config.S3_API_KEY,
    },
    body: JSON.stringify({ bucket: config.S3_BUCKET, key, action }),
  })
  if (!response.ok) throw new Error(`Presigned URL 발급 실패: ${response.status}`)
  const data = await response.json() as { url: string }
  return data.url
}

export async function uploadJson(path: string, data: unknown): Promise<void> {
  const key = buildKey(path)
  const presignedUrl = await getPresignedUrl(key, 'PUT_OBJECT')
  const body = JSON.stringify(data)
  const response = await fetch(presignedUrl, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body,
  })
  if (!response.ok) throw new Error(`S3 업로드 실패: ${response.status}`)
}

export async function downloadJson(path: string): Promise<unknown> {
  const key = buildKey(path)
  const presignedUrl = await getPresignedUrl(key, 'GET_OBJECT')
  const response = await fetch(presignedUrl)
  if (!response.ok) {
    if (response.status === 404) return null
    throw new Error(`S3 다운로드 실패: ${response.status}`)
  }
  return response.json()
}

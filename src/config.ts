import { z } from 'zod'
import dotenv from 'dotenv'
dotenv.config()

const envSchema = z.object({
  PORT: z.string().default('3000'),
  NODE_ENV: z.string().default('development'),
  S3_API_BASE_URL: z.string().default('https://aviyup1kyk.execute-api.ap-northeast-2.amazonaws.com/prod'),
  S3_API_KEY: z.string(),
  S3_BUCKET: z.string().default('svc-fnf-ax-platform-pub-s3'),
})

const parsed = envSchema.safeParse(process.env)
if (!parsed.success) {
  console.error('환경변수 오류:', parsed.error.format())
  process.exit(1)
}

export const config = parsed.data

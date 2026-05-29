import { Router, Request, Response } from 'express'
import multer from 'multer'
import { parseStExcel, parseDvExcel, extractImagesFromXlsx } from '../services/excel.service'
import { uploadJson, downloadJson } from '../services/s3.service'

const router = Router()
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } })

const ST_KEY = 'data/st-schedule.json'
const DV_KEY = 'data/dv-schedule.json'

router.post('/upload/st', upload.single('file'), async (req: Request, res: Response) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: '파일이 없습니다.' })
    const rows = parseStExcel(req.file.buffer)
    const images = extractImagesFromXlsx(req.file.buffer)
    const payload = { updatedAt: new Date().toISOString(), rows, images }
    await uploadJson(ST_KEY, payload)
    res.json({ success: true, count: rows.length, updatedAt: payload.updatedAt })
  } catch (e) {
    res.status(500).json({ success: false, error: e instanceof Error ? e.message : String(e) })
  }
})

router.post('/upload/dv', upload.single('file'), async (req: Request, res: Response) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: '파일이 없습니다.' })
    const rows = parseDvExcel(req.file.buffer)
    const images = extractImagesFromXlsx(req.file.buffer)
    const payload = { updatedAt: new Date().toISOString(), rows, images }
    await uploadJson(DV_KEY, payload)
    res.json({ success: true, count: rows.length, updatedAt: payload.updatedAt })
  } catch (e) {
    res.status(500).json({ success: false, error: e instanceof Error ? e.message : String(e) })
  }
})

router.get('/data/st', async (_req: Request, res: Response) => {
  try {
    const data = await downloadJson(ST_KEY)
    res.json(data ?? { rows: [], images: {}, updatedAt: null })
  } catch (e) {
    res.status(500).json({ success: false, error: e instanceof Error ? e.message : String(e) })
  }
})

router.get('/data/dv', async (_req: Request, res: Response) => {
  try {
    const data = await downloadJson(DV_KEY)
    res.json(data ?? { rows: [], images: {}, updatedAt: null })
  } catch (e) {
    res.status(500).json({ success: false, error: e instanceof Error ? e.message : String(e) })
  }
})

export { router as scheduleRoutes }

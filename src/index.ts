import express from 'express'
import cors from 'cors'
import path from 'path'
import { config } from './config'
import { scheduleRoutes } from './routes/schedule'

const app = express()

app.use(cors())
app.use(express.json())
app.use(express.static(path.join(__dirname, '../public')))
app.use('/api', scheduleRoutes)

// SPA fallback
app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, '../public/index.html'))
})

app.listen(Number(config.PORT), () => {
  console.log(`Server running on port ${config.PORT}`)
})

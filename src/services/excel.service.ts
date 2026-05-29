import * as XLSX from 'xlsx'
import AdmZip from 'adm-zip'

// ─── 날짜 파싱 (timezone-safe) ───────────────────────────────────────────────
function fmtDate(val: unknown): string {
  if (!val) return ''
  if (typeof val === 'number' && val > 1) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(val) * 86400000)
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`
  }
  if (val instanceof Date) {
    return `${val.getUTCFullYear()}-${String(val.getUTCMonth()+1).padStart(2,'0')}-${String(val.getUTCDate()).padStart(2,'0')}`
  }
  return String(val).trim()
}

function fmtStr(val: unknown): string {
  return val == null ? '' : String(val).trim()
}

// ─── 헤더 정규화 (공백·개행·특수문자 제거, 소문자) ─────────────────────────
function norm(s: unknown): string {
  return String(s ?? '').replace(/[\s\n\r\t※()\[\]·\-]+/g, '').toLowerCase()
}

// ─── ST 컬럼 규칙 ────────────────────────────────────────────────────────────
// 정규화된 헤더 → 필드명. 두 번째 'mold'는 mold_code로 처리
const ST_COL_RULES: [string[], string][] = [
  [['no', 'no.'], 'no'],
  [['image', '이미지'], '_img'],
  [['drop', '드랍'], 'drop'],
  [['season', '시즌'], 'season'],
  [['stylename', 'style', '스타일명', '제품명'], 'style'],
  [['mold', '금형'], '_mold'],          // 첫 번째 mold
  [['last', '라스트'], 'last'],
  [['launch', '출시', '런치'], 'launch'],
  [['price', '가격', '소비자가', 'krw'], 'price'],
  [['cost', '원가', '코스트'], 'cost'],
  [['designer', '디자이너'], 'designer'],
  [['md'], 'md'],
  [['소싱', 'sourcing'], 'sourcing'],
  [['factory', '공장', '팩토리'], 'factory'],
  [['1차작지', '1stspec', '1차작지완료'], 'spec_1'],
  [['1차샘플도착', '1차샘플', '1stsample'], 'arrive_1'],
  [['2차작지', '2ndspec', '2차작지완료'], 'spec_2'],
  [['2차샘플도착', '2차샘플', '2ndsample'], 'arrive_2'],
  [['3차작지', '3rdspec', '3차작지완료'], 'spec_3'],
  [['3차샘플도착', '3차샘플', '3rdsample'], 'arrive_3'],
  [['cfm작지', 'cfm작지완료', 'cfmspec'], 'spec_cfm'],
  [['cfm샘플도착', 'cfm샘플', 'cfmsample'], 'arrive_cfm'],
  [['po발주', 'po', 'purchaseorder'], 'po'],
  [['생산시화', '생산proto', 'proto'], 'proto_prod'],
  [['remark', '비고', '개발'], 'remark_dev'],
]

// ─── DV 컬럼 규칙 ────────────────────────────────────────────────────────────
const DV_COL_RULES: [string[], string][] = [
  [['no', 'no.'], 'no'],
  [['image', '이미지'], '_img'],
  [['drop', '드랍'], 'drop'],
  [['season', '시즌'], 'season'],
  [['domain', '도메인'], 'domain'],
  [['gender', '성별'], 'gender'],
  [['color', '컬러', '칼라'], 'color'],
  [['stylename', 'style', '스타일명', '제품명'], 'style'],
  [['mold', '금형'], '_mold'],
  [['last', '라스트'], 'last'],
  [['size', '사이즈'], 'size'],
  [['launch', '출시', '런치'], 'launch'],
  [['price', '가격', '소비자가', 'krw'], 'price'],
  [['cost', '원가', '코스트'], 'cost'],
  [['designer', '디자이너'], 'designer'],
  [['md'], 'md'],
  [['소싱', 'sourcing'], 'sourcing'],
  [['factory', '공장', '팩토리'], 'factory'],
  [['디자인보고', 'designreview', '디자인리뷰'], 'design_review'],
  [['1차작지', '1stspec', '1차작지완료'], 'spec_1'],
  [['1차샘플의뢰', '1차의뢰', 'samplereq'], 'sample_req_1'],
  [['시화', 'proto', 'prototype'], 'proto_1'],
  [['일정target', 'target', '일정타겟'], '_target'],     // 1st target
  [['1차샘플도착', '1차샘플', '1stsample'], 'arrive_1'],
  [['샘플보고', 'samplereport', '샘플리포트'], 'report_sample'],
  [['2차작지', '2ndspec', '2차작지완료'], 'spec_2'],
  [['2차샘플도착', '2차샘플', '2ndsample'], 'arrive_2'],
  [['3차작지', '3rdspec', '3차작지완료'], 'spec_3'],
  [['3차샘플도착', '3차샘플', '3rdsample'], 'arrive_3'],
  [['cfm작지', 'cfm작지완료', 'cfmspec'], 'spec_cfm'],
  [['cfm샘플도착', 'cfm샘플', 'cfmsample'], 'arrive_cfm'],
  [['po발주', 'po발주중국제외', 'po'], 'po'],
  [['생산시화', '생산proto'], 'proto_prod'],
  [['gbtest', 'gb테스트', 'gbtest'], 'gb_test'],
  [['po발주중국', 'pocn', 'po중국'], 'po_cn'],
  [['입고', 'inbound'], 'inbound'],
  [['remark개발', '개발비고', '개발remark'], 'remark_dev'],
  [['remark생산', '생산비고', '생산remark'], 'remark_prod'],
]

type ColRule = [string[], string]

function buildIndexMap(headers: unknown[], rules: ColRule[]): Record<number, string> {
  const map: Record<number, string> = {}
  const usedFields = new Set<string>()

  headers.forEach((h, idx) => {
    const n = norm(h)
    if (!n) return
    for (const [keywords, field] of rules) {
      if (keywords.some(k => n.includes(k))) {
        // _mold: 첫 번째 → 'mold', 두 번째 → 'mold_code'
        if (field === '_mold') {
          if (!usedFields.has('mold')) { map[idx] = 'mold'; usedFields.add('mold') }
          else { map[idx] = 'mold_code' }
          return
        }
        // _target: 순서대로 target_1, target_2, target_3
        if (field === '_target') {
          const n = ['target_1','target_2','target_3'].find(t => !usedFields.has(t))
          if (n) { map[idx] = n; usedFields.add(n) }
          return
        }
        if (!usedFields.has(field)) {
          map[idx] = field
          usedFields.add(field)
          return
        }
      }
    }
  })
  return map
}

// ─── xlsx에서 이미지 추출 (adm-zip으로 xlsx zip 파싱) ────────────────────────
// 반환: { [excelRowNum(1-based)]: [{mime, data}] }
export function extractImagesFromXlsx(buffer: Buffer): Record<string, {mime: string, data: string}[]> {
  const result: Record<string, {mime: string, data: string}[]> = {}
  try {
    const zip = new AdmZip(buffer)

    // 1) rels: rId → 파일경로
    const relsEntry = zip.getEntry('xl/drawings/_rels/drawing1.xml.rels')
    if (!relsEntry) return result
    const relsXml = relsEntry.getData().toString('utf-8')
    const ridToFile: Record<string, string> = {}
    const relRe = /Id="([^"]+)"[^>]+Target="([^"]+)"/g
    let m: RegExpExecArray | null
    while ((m = relRe.exec(relsXml)) !== null) {
      ridToFile[m[1]] = m[2].replace('../', 'xl/')
    }

    // 2) drawing XML 파싱: <xdr:twoCellAnchor> 블록별로 row + rId 추출
    const drawEntry = zip.getEntry('xl/drawings/drawing1.xml')
    if (!drawEntry) return result
    const drawXml = drawEntry.getData().toString('utf-8')

    // twoCellAnchor 블록 분리
    const anchorRe = /<xdr:twoCellAnchor[\s\S]*?<\/xdr:twoCellAnchor>/g
    while ((m = anchorRe.exec(drawXml)) !== null) {
      const block = m[0]
      // from row (0-based)
      const rowMatch = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(block)
      if (!rowMatch) continue
      const excelRow = parseInt(rowMatch[1]) + 1  // 1-based

      // 이 블록 내 모든 r:embed
      const embedRe = /r:embed="([^"]+)"/g
      let em: RegExpExecArray | null
      while ((em = embedRe.exec(block)) !== null) {
        const rid = em[1]
        const filePath = ridToFile[rid]
        if (!filePath) continue
        const mediaEntry = zip.getEntry(filePath)
        if (!mediaEntry) continue
        const ext = filePath.split('.').pop()?.toLowerCase() ?? 'png'
        const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/png'
        const data = mediaEntry.getData().toString('base64')
        if (!result[excelRow]) result[excelRow] = []
        result[excelRow].push({ mime, data })
      }
    }
  } catch (e) {
    console.warn('이미지 추출 실패:', e)
  }
  return result
}

// ─── ST 파싱 ─────────────────────────────────────────────────────────────────
export function parseStExcel(buffer: Buffer): object[] {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false })
  const ws = wb.Sheets['제품일정']
  if (!ws) throw new Error('시트 "제품일정"을 찾을 수 없습니다.')

  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
  const headers = rows[2] as unknown[]  // row 3 (0-based index 2) = 헤더
  const colMap = buildIndexMap(headers, ST_COL_RULES)
  const result: object[] = []

  for (let i = 3; i < rows.length; i++) {
    const r = rows[i] as unknown[]
    const noIdx = Object.entries(colMap).find(([,f]) => f === 'no')?.[0]
    const styleIdx = Object.entries(colMap).find(([,f]) => f === 'style')?.[0]
    const no = noIdx != null ? r[Number(noIdx)] : null
    const style = styleIdx != null ? fmtStr(r[Number(styleIdx)]) : ''
    if (no == null || no === '' || style === '') continue

    const row: Record<string, string> = { _sheetRow: String(i + 1) }
    for (const [idx, field] of Object.entries(colMap)) {
      if (field === '_img') continue
      const val = r[Number(idx)]
      const isDateField = ['launch','spec_1','arrive_1','spec_2','arrive_2','spec_3','arrive_3','spec_cfm','arrive_cfm','po','proto_prod'].includes(field)
      row[field] = isDateField ? fmtDate(val) : fmtStr(val)
    }
    result.push(row)
  }
  return result
}

// ─── DV 파싱 ─────────────────────────────────────────────────────────────────
export function parseDvExcel(buffer: Buffer): object[] {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false })
  const ws = wb.Sheets['제품일정']
  if (!ws) throw new Error('시트 "제품일정"을 찾을 수 없습니다.')

  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
  const headers = rows[2] as unknown[]
  const colMap = buildIndexMap(headers, DV_COL_RULES)
  const result: object[] = []

  const DATE_FIELDS = new Set(['launch','design_review','spec_1','sample_req_1','target_1','arrive_1','spec_2','target_2','arrive_2','spec_3','target_3','arrive_3','spec_cfm','arrive_cfm','po','proto_prod','po_cn','inbound'])

  for (let i = 3; i < rows.length; i++) {
    const r = rows[i] as unknown[]
    const noIdx = Object.entries(colMap).find(([,f]) => f === 'no')?.[0]
    const styleIdx = Object.entries(colMap).find(([,f]) => f === 'style')?.[0]
    const no = noIdx != null ? r[Number(noIdx)] : null
    const style = styleIdx != null ? fmtStr(r[Number(styleIdx)]) : ''
    if (no == null || no === '' || style === '') continue

    const row: Record<string, string> = { _sheetRow: String(i + 1) }
    for (const [idx, field] of Object.entries(colMap)) {
      if (field === '_img') continue
      const val = r[Number(idx)]
      row[field] = DATE_FIELDS.has(field) ? fmtDate(val) : fmtStr(val)
    }
    result.push(row)
  }
  return result
}

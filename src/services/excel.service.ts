import * as XLSX from 'xlsx'

// Excel serial date → YYYY-MM-DD (timezone-safe)
// cellDates: false로 읽어 serial number 그대로 처리
function fmtDate(val: unknown): string {
  if (!val) return ''
  if (typeof val === 'number' && val > 1) {
    // Excel epoch: Dec 30, 1899 = 0
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(val) * 86400000)
    const y = d.getUTCFullYear()
    const m = String(d.getUTCMonth() + 1).padStart(2, '0')
    const day = String(d.getUTCDate()).padStart(2, '0')
    return `${y}-${m}-${day}`
  }
  if (val instanceof Date) {
    // fallback: cellDates: true 모드 대비 (UTC 기준)
    const y = val.getUTCFullYear()
    const m = String(val.getUTCMonth() + 1).padStart(2, '0')
    const day = String(val.getUTCDate()).padStart(2, '0')
    return `${y}-${m}-${day}`
  }
  return String(val).trim()
}

function fmtStr(val: unknown): string {
  if (val == null) return ''
  return String(val).trim()
}

// ST (세르지오타키니) 제품일정 파싱
// Sheet range: B1:AB38  → 배열 인덱스 0 = 컬럼 B
// B=NO, C=IMAGE, D=DROP, E=season, F=STYLE NAME,
// G=MOLD, H=LAST, I=LAUNCH, J=PRICE, K=COST, L=MOLD_CODE,
// M=Designer, N=MD, O=소싱, P=Factory,
// Q=1차작지(15), R=1차샘플(16), S=2차작지(17), T=2차샘플(18),
// U=3차작지(19), V=3차샘플(20), W=CFM작지(21), X=CFM샘플(22),
// Y=PO(23), Z=생산시화(24), AA=REMARK(25)
export function parseStExcel(buffer: Buffer): object[] {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false })
  const ws = wb.Sheets['제품일정']
  if (!ws) throw new Error('시트 "제품일정"을 찾을 수 없습니다.')

  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
  const result: object[] = []

  for (let i = 3; i < rows.length; i++) {
    const r = rows[i] as unknown[]
    const no = r[0]  // B
    const style = fmtStr(r[4])  // F: STYLE NAME
    if (no == null || no === '' || style === '') continue  // 빈 placeholder 행 제외

    result.push({
      _sheetRow: String(i + 1),
      no: fmtStr(no),
      drop: fmtStr(r[2]),        // D
      season: fmtStr(r[3]),      // E
      style,                     // F
      mold: fmtStr(r[5]),        // G
      last: fmtStr(r[6]),        // H
      launch: fmtDate(r[7]),     // I
      price: fmtStr(r[8]),       // J
      cost: fmtStr(r[9]),        // K
      mold_code: fmtStr(r[10]),  // L
      designer: fmtStr(r[11]),   // M
      md: fmtStr(r[12]),         // N
      sourcing: fmtStr(r[13]),   // O
      factory: fmtStr(r[14]),    // P
      spec_1: fmtDate(r[15]),    // Q 1차 작지
      arrive_1: fmtDate(r[16]),  // R 1차 샘플
      spec_2: fmtDate(r[17]),    // S 2차 작지
      arrive_2: fmtDate(r[18]),  // T 2차 샘플
      spec_3: fmtDate(r[19]),    // U 3차 작지
      arrive_3: fmtDate(r[20]),  // V 3차 샘플
      spec_cfm: fmtDate(r[21]),  // W CFM 작지
      arrive_cfm: fmtDate(r[22]), // X CFM 샘플
      po: fmtDate(r[23]),        // Y PO
      proto_prod: fmtDate(r[24]), // Z 생산시화
      remark_dev: fmtStr(r[25]), // AA REMARK
    })
  }

  return result
}

// DV (듀베티카) 제품일정 파싱
// Sheet range: B1:AQ54  → 배열 인덱스 0 = 컬럼 B
// B=NO(0), C=IMAGE(1), D=DROP(2), E=SEASON(3), F=DOMAIN(4), G=Gender(5),
// H=COLOR(6), I=STYLE NAME(7), J=MOLD(8), K=LAST(9), L=SIZE(10),
// M=LAUNCH(11), N=PRICE(12), O=COST(13), P=MOLD_CODE(14),
// Q=Designer(15), R=MD(16), S=소싱(17), T=Factory(18),
// U=디자인보고(19), V=1차작지(20), W=1차샘플의뢰(21), X=시화(22),
// Y=TARGET1(23), Z=1차샘플도착(24), AA=샘플보고(25),
// AB=2차작지(26), AC=TARGET2(27), AD=2차샘플(28),
// AE=3차작지(29), AF=TARGET3(30), AG=3차샘플(31),
// AH=CFM작지(32), AI=CFM샘플(33), AJ=PO(34), AK=생산시화(35),
// AL=GBTEST(36), AM=PO중국(37), AN=입고(38),
// AO=REMARK개발(39), AP=REMARK생산(40)
export function parseDvExcel(buffer: Buffer): object[] {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false })
  const ws = wb.Sheets['제품일정']
  if (!ws) throw new Error('시트 "제품일정"을 찾을 수 없습니다.')

  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
  const result: object[] = []

  for (let i = 3; i < rows.length; i++) {
    const r = rows[i] as unknown[]
    const no = r[0]  // B
    if (no == null || no === '') continue

    result.push({
      _sheetRow: String(i + 1),
      no: fmtStr(no),
      drop: fmtStr(r[2]),            // D
      season: fmtStr(r[3]),          // E
      domain: fmtStr(r[4]),          // F
      gender: fmtStr(r[5]),          // G
      color: fmtStr(r[6]),           // H
      style: fmtStr(r[7]),           // I
      mold: fmtStr(r[8]),            // J
      last: fmtStr(r[9]),            // K
      size: fmtStr(r[10]),           // L
      launch: fmtDate(r[11]),        // M
      price: fmtStr(r[12]),          // N
      cost: fmtStr(r[13]),           // O
      mold_code: fmtStr(r[14]),      // P
      designer: fmtStr(r[15]),       // Q
      md: fmtStr(r[16]),             // R
      sourcing: fmtStr(r[17]),       // S
      factory: fmtStr(r[18]),        // T
      design_review: fmtDate(r[19]), // U 디자인보고
      spec_1: fmtDate(r[20]),        // V 1차작지
      sample_req_1: fmtDate(r[21]),  // W 1차샘플의뢰
      proto_1: fmtStr(r[22]),        // X 시화
      target_1: fmtDate(r[23]),      // Y TARGET
      arrive_1: fmtDate(r[24]),      // Z 1차샘플도착
      report_sample: fmtStr(r[25]),  // AA 샘플보고
      spec_2: fmtDate(r[26]),        // AB 2차작지
      target_2: fmtDate(r[27]),      // AC TARGET
      arrive_2: fmtDate(r[28]),      // AD 2차샘플
      spec_3: fmtDate(r[29]),        // AE 3차작지
      target_3: fmtDate(r[30]),      // AF TARGET
      arrive_3: fmtDate(r[31]),      // AG 3차샘플
      spec_cfm: fmtDate(r[32]),      // AH CFM작지
      arrive_cfm: fmtDate(r[33]),    // AI CFM샘플
      po: fmtDate(r[34]),            // AJ PO
      proto_prod: fmtDate(r[35]),    // AK 생산시화
      gb_test: fmtStr(r[36]),        // AL GB TEST
      po_cn: fmtDate(r[37]),         // AM PO중국
      inbound: fmtDate(r[38]),       // AN 입고
      remark_dev: fmtStr(r[39]),     // AO REMARK개발
      remark_prod: fmtStr(r[40]),    // AP REMARK생산
    })
  }

  return result
}

import { NextRequest, NextResponse } from 'next/server'
import { createSupabaseServerClient, createSupabaseServiceClient } from '@/lib/supabase-server'
import * as XLSX from 'xlsx'

interface WorkbookCell {
  section: string
  row_key: string
  period: string
  raw_value: number | null
  display_value: string | null
  is_overridden: boolean
  confidence: number | null
  source_excerpt: string | null
}

interface AuditEntry {
  edited_at: string
  old_value: number | null
  new_value: number | null
  note: string | null
  cell: {
    section: string
    row_key: string
    period: string
  } | null
}

// ─── Excel ────────────────────────────────────────────────────────────────────

function buildExcel(
  companyName: string,
  cells: WorkbookCell[],
  auditEntries: AuditEntry[],
  periods: string[]
): Buffer {
  const wb = XLSX.utils.book_new()

  const sections = [...new Set(cells.map((c) => c.section))]

  for (const section of sections) {
    const sectionCells = cells.filter((c) => c.section === section)
    const rowKeys = [...new Set(sectionCells.map((c) => c.row_key))]

    const headerRow = ['Line Item', ...periods, 'Notes']
    const dataRows = rowKeys.map((rowKey) => {
      const row: (string | number | null)[] = [rowKey]
      for (const period of periods) {
        const cell = sectionCells.find((c) => c.row_key === rowKey && c.period === period)
        row.push(cell?.raw_value ?? null)
      }
      row.push('')
      return row
    })

    const ws = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows])

    const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1')
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C })
      if (!ws[addr]) continue
      ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: 'D9E1F2' } } }
    }

    ws['!cols'] = [{ wch: 40 }, ...periods.map(() => ({ wch: 18 })), { wch: 30 }]

    const sheetName = section.replace(/-/g, ' ').slice(0, 31)
    XLSX.utils.book_append_sheet(wb, ws, sheetName)
  }

  const auditHeader = ['Date/Time', 'Section', 'Row Key', 'Period', 'Old Value', 'New Value', 'Note']
  const auditRows = auditEntries.map((e) => [
    new Date(e.edited_at).toLocaleString(),
    e.cell?.section ?? '',
    e.cell?.row_key ?? '',
    e.cell?.period ?? '',
    e.old_value ?? '',
    e.new_value ?? '',
    e.note ?? '',
  ])
  const auditWs = XLSX.utils.aoa_to_sheet([auditHeader, ...auditRows])
  auditWs['!cols'] = [{ wch: 22 }, { wch: 20 }, { wch: 35 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 40 }]
  XLSX.utils.book_append_sheet(wb, auditWs, 'Audit Trail')

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer
}

// ─── PDF ──────────────────────────────────────────────────────────────────────

async function buildPdf(
  companyName: string,
  cells: WorkbookCell[],
  auditEntries: AuditEntry[],
  periods: string[]
): Promise<Buffer> {
  // Dynamic import keeps PDFKit out of the edge bundle
  const PDFDocument = (await import('pdfkit')).default

  const PAGE_W = 792  // letter landscape
  const PAGE_H = 612
  const MARGIN = 40
  const USABLE_W = PAGE_W - MARGIN * 2

  const LABEL_W = 230
  const VAL_W = (USABLE_W - LABEL_W) / Math.max(periods.length, 1)
  const ROW_H = 18
  const HDR_H = 22

  const BRAND = '#1E3A5F'
  const HEADER_BG = '#D9E1F2'
  const ALT_BG = '#F7F9FC'
  const DARK = '#1A1A1A'
  const GRAY = '#666666'
  const RULE = '#E0E0E0'

  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: [PAGE_W, PAGE_H],
      margins: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      autoFirstPage: false,
      bufferPages: true,  // needed for footer pass
    })

    const chunks: Buffer[] = []
    doc.on('data', (chunk: Buffer) => chunks.push(chunk))
    doc.on('end', () => resolve(Buffer.concat(chunks)))
    doc.on('error', reject)

    // ── Cover ────────────────────────────────────────────────────
    doc.addPage()

    doc.rect(0, 0, PAGE_W, 130).fill(BRAND)

    doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(26)
      .text(companyName, MARGIN, 38, { width: USABLE_W })

    doc.fillColor('#B0C4DE').font('Helvetica').fontSize(14)
      .text('Quality of Earnings Report', MARGIN, 84, { width: USABLE_W })

    const sections = [...new Set(cells.map((c) => c.section))]
    const genDate = new Date().toLocaleDateString('en-US', {
      year: 'numeric', month: 'long', day: 'numeric',
    })

    doc.fillColor(DARK).font('Helvetica').fontSize(11)
      .text(`Generated: ${genDate}`, MARGIN, 155)
      .text(`Periods: ${periods.join('  |  ')}`, MARGIN, 175)
      .text(
        `Sections: ${sections.length}  |  Data points: ${cells.length}`,
        MARGIN, 195,
      )

    // ── Section pages ────────────────────────────────────────────
    for (const section of sections) {
      doc.addPage()
      const sectionCells = cells.filter((c) => c.section === section)
      const rowKeys = [...new Set(sectionCells.map((c) => c.row_key))]
      const displayName = section
        .replace(/-/g, ' ')
        .replace(/\b\w/g, (l) => l.toUpperCase())

      // Section title bar
      doc.rect(MARGIN, MARGIN, USABLE_W, 28).fill(BRAND)
      doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(13)
        .text(displayName, MARGIN + 8, MARGIN + 8, { width: USABLE_W - 16 })

      let y = MARGIN + 38

      // Column header row
      doc.rect(MARGIN, y, USABLE_W, HDR_H).fill(HEADER_BG)
      doc.fillColor(DARK).font('Helvetica-Bold').fontSize(9)
        .text('Line Item', MARGIN + 4, y + 6, { width: LABEL_W - 8 })
      for (let i = 0; i < periods.length; i++) {
        const x = MARGIN + LABEL_W + VAL_W * i
        doc.text(periods[i], x, y + 6, { width: VAL_W - 4, align: 'right' })
      }
      y += HDR_H

      // Data rows
      for (let ri = 0; ri < rowKeys.length; ri++) {
        const rowKey = rowKeys[ri]

        if (ri % 2 === 0) doc.rect(MARGIN, y, USABLE_W, ROW_H).fill(ALT_BG)

        doc.fillColor(DARK).font('Helvetica').fontSize(8.5)
          .text(
            rowKey.replace(/_/g, ' '),
            MARGIN + 4, y + 4,
            { width: LABEL_W - 8 },
          )

        for (let i = 0; i < periods.length; i++) {
          const cell = sectionCells.find(
            (c) => c.row_key === rowKey && c.period === periods[i],
          )
          const val = cell?.display_value
            ?? (cell?.raw_value != null ? String(cell.raw_value) : '—')
          const x = MARGIN + LABEL_W + VAL_W * i
          doc.text(val, x, y + 4, { width: VAL_W - 4, align: 'right' })
        }

        doc.moveTo(MARGIN, y + ROW_H)
          .lineTo(MARGIN + USABLE_W, y + ROW_H)
          .strokeColor(RULE).lineWidth(0.5).stroke()

        y += ROW_H

        if (y + ROW_H > PAGE_H - MARGIN - 24) {
          doc.addPage()
          y = MARGIN
        }
      }
    }

    // ── Audit trail ──────────────────────────────────────────────
    if (auditEntries.length > 0) {
      doc.addPage()

      doc.rect(MARGIN, MARGIN, USABLE_W, 28).fill(BRAND)
      doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(13)
        .text('Audit Trail', MARGIN + 8, MARGIN + 8)

      const AUDIT_W = [120, 100, 160, 55, 75, 75]
      const AUDIT_HDR = ['Date/Time', 'Section', 'Row Key', 'Period', 'Old Value', 'New Value']

      let y = MARGIN + 38

      doc.rect(MARGIN, y, USABLE_W, HDR_H).fill(HEADER_BG)
      doc.fillColor(DARK).font('Helvetica-Bold').fontSize(8.5)
      let cx = MARGIN + 4
      for (let i = 0; i < AUDIT_HDR.length; i++) {
        doc.text(AUDIT_HDR[i], cx, y + 6, { width: AUDIT_W[i] - 4 })
        cx += AUDIT_W[i]
      }
      y += HDR_H

      for (let ei = 0; ei < auditEntries.length; ei++) {
        const entry = auditEntries[ei]
        if (ei % 2 === 0) doc.rect(MARGIN, y, USABLE_W, ROW_H).fill(ALT_BG)

        doc.fillColor(DARK).font('Helvetica').fontSize(8)
        const row = [
          new Date(entry.edited_at).toLocaleString('en-US', {
            dateStyle: 'short', timeStyle: 'short',
          }),
          entry.cell?.section ?? '',
          (entry.cell?.row_key ?? '').replace(/_/g, ' '),
          entry.cell?.period ?? '',
          entry.old_value != null ? String(entry.old_value) : '—',
          entry.new_value != null ? String(entry.new_value) : '—',
        ]

        cx = MARGIN + 4
        for (let i = 0; i < row.length; i++) {
          doc.text(row[i], cx, y + 4, { width: AUDIT_W[i] - 4 })
          cx += AUDIT_W[i]
        }

        doc.moveTo(MARGIN, y + ROW_H)
          .lineTo(MARGIN + USABLE_W, y + ROW_H)
          .strokeColor(RULE).lineWidth(0.5).stroke()

        y += ROW_H

        if (y + ROW_H > PAGE_H - MARGIN - 24) {
          doc.addPage()
          y = MARGIN
        }
      }
    }

    // ── Footer on every page ─────────────────────────────────────
    const range = doc.bufferedPageRange()
    for (let i = 0; i < range.count; i++) {
      doc.switchToPage(range.start + i)
      doc.fillColor(GRAY).font('Helvetica').fontSize(7.5)
        .text(
          `${companyName} — Quality of Earnings Workbook  |  Page ${i + 1} of ${range.count}`,
          MARGIN,
          PAGE_H - MARGIN + 12,
          { width: USABLE_W, align: 'center' },
        )
    }

    doc.end()
  })
}

// ─── Google Sheets ────────────────────────────────────────────────────────────

async function exportToGoogleSheets(
  companyName: string,
  cells: WorkbookCell[],
  periods: string[]
): Promise<string> {
  const { google } = await import('googleapis')
  const credentialsJson = process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON
  if (!credentialsJson) throw new Error('Google credentials not configured')

  const credentials = JSON.parse(credentialsJson)
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  })

  const sheets = google.sheets({ version: 'v4', auth })

  const spreadsheet = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: `QoE Workbook - ${companyName}` },
    },
  })

  const spreadsheetId = spreadsheet.data.spreadsheetId!

  const sections = [...new Set(cells.map((c) => c.section))]
  const requests: object[] = []

  for (let i = 1; i < sections.length; i++) {
    requests.push({
      addSheet: {
        properties: {
          title: sections[i].replace(/-/g, ' '),
          index: i,
        },
      },
    })
  }

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    })
  }

  const valueRanges = sections.map((section) => {
    const sectionCells = cells.filter((c) => c.section === section)
    const rowKeys = [...new Set(sectionCells.map((c) => c.row_key))]
    const sheetTitle = section.replace(/-/g, ' ')

    const values = [
      ['Line Item', ...periods],
      ...rowKeys.map((rowKey) => [
        rowKey,
        ...periods.map((period) => {
          const cell = sectionCells.find((c) => c.row_key === rowKey && c.period === period)
          return cell?.raw_value ?? ''
        }),
      ]),
    ]

    return { range: `${sheetTitle}!A1`, values }
  })

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: 'RAW',
      data: valueRanges,
    },
  })

  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}`
}

// ─── Route handler ────────────────────────────────────────────────────────────

export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id: workbookId } = await params
  const supabase = await createSupabaseServerClient()
  const { data: { user }, error: authError } = await supabase.auth.getUser()
  if (authError || !user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const body = await req.json()
  const format: 'excel' | 'pdf' | 'sheets' | 'airtable' = body.format ?? 'excel'

  const serviceClient = createSupabaseServiceClient()

  const { data: workbook } = await serviceClient
    .from('workbooks')
    .select('id, company_name, periods, created_by')
    .eq('id', workbookId)
    .eq('created_by', user.id)
    .single()

  if (!workbook) return NextResponse.json({ error: 'Workbook not found' }, { status: 404 })

  const { data: cells } = await serviceClient
    .from('workbook_cells')
    .select('section, row_key, period, raw_value, display_value, is_overridden, confidence, source_excerpt')
    .eq('workbook_id', workbookId)
    .order('section')
    .order('row_key')

  const { data: auditEntries } = await serviceClient
    .from('audit_entries')
    .select('edited_at, old_value, new_value, note, cell:workbook_cells(section, row_key, period)')
    .eq('workbook_id', workbookId)
    .order('edited_at', { ascending: false })

  const periods: string[] = workbook.periods ?? ['FY20', 'FY21', 'FY22', 'TTM']

  if (format === 'excel') {
    const buf = buildExcel(
      workbook.company_name,
      (cells ?? []) as WorkbookCell[],
      ((auditEntries ?? []) as unknown) as AuditEntry[],
      periods,
    )
    const safe = workbook.company_name.replace(/[^a-z0-9]/gi, '_')
    return new NextResponse(buf as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${safe}_QoE.xlsx"`,
      },
    })
  }

  if (format === 'pdf') {
    try {
      const buf = await buildPdf(
        workbook.company_name,
        (cells ?? []) as WorkbookCell[],
        ((auditEntries ?? []) as unknown) as AuditEntry[],
        periods,
      )
      const safe = workbook.company_name.replace(/[^a-z0-9]/gi, '_')
      return new NextResponse(buf as unknown as BodyInit, {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${safe}_QoE.pdf"`,
        },
      })
    } catch (err) {
      return NextResponse.json({ error: String(err) }, { status: 500 })
    }
  }

  if (format === 'sheets') {
    try {
      const url = await exportToGoogleSheets(
        workbook.company_name,
        (cells ?? []) as WorkbookCell[],
        periods,
      )
      return NextResponse.json({ url })
    } catch (err) {
      return NextResponse.json({ error: String(err) }, { status: 500 })
    }
  }

  return NextResponse.json({ error: `Format "${format}" not yet implemented` }, { status: 400 })
}

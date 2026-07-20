import * as XLSX from 'xlsx'

export interface ParsedSheet {
  name: string
  headers: string[]
  rows: (string | number | null)[][]
  totalRows: number
  totalCols: number
}

export interface FileParseResult {
  fileName: string
  fileSize: number
  sheets: ParsedSheet[]
  textRepresentation: string // pour le prompt IA
}

async function readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as ArrayBuffer)
    reader.onerror = () => reject(new Error('Échec de la lecture du fichier'))
    reader.readAsArrayBuffer(file)
  })
}

function sheetToParsedSheet(
  sheet: XLSX.WorkSheet,
  name: string,
  maxRows: number = 50
): ParsedSheet {
  const data = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1,
    defval: null,
  })
  
  const headers =
    data.length > 0
      ? (data[0] as (string | number | null)[]).map((h) =>
          h !== null ? String(h) : ''
        )
      : []
      
  const rows = data.slice(1, maxRows).map(
    (row) =>
      (row as (string | number | null)[]).slice(0, headers.length) as (
        | string
        | number
        | null
      )[]
  )
  
  return {
    name,
    headers,
    rows,
    totalRows: data.length > 0 ? data.length - 1 : 0,
    totalCols: headers.length,
  }
}

function detectFormulas(sheet: XLSX.WorkSheet): Record<string, string> {
  const formulas: Record<string, string> = {};
  for (const cellAddr of Object.keys(sheet)) {
    if (cellAddr.startsWith("!")) continue;
    const cell = sheet[cellAddr];
    if (cell?.f) {
      formulas[cellAddr] = cell.f.startsWith("=") ? cell.f : `=${cell.f}`;
    }
  }
  return formulas;
}

function buildTextRepresentation(sheet: ParsedSheet, formulas: Record<string, string>): string {
  const lines: string[] = [];
  lines.push(`[Feuille : ${sheet.name}]`);
  lines.push(`Colonnes : ${sheet.totalCols} · Lignes : ${sheet.totalRows}`);
  
  if (sheet.totalCols > 0) {
    lines.push(`En-têtes : ${sheet.headers.join(", ")}`);
  }
  
  const formulaEntries = Object.entries(formulas);
  if (formulaEntries.length > 0) {
    lines.push(`⚠️ Formules détectées dans le fichier d'origine (${formulaEntries.length}) :`);
    for (const [addr, f] of formulaEntries.slice(0, 10)) {
      lines.push(`  ${addr} : ${f}`);
    }
  }
  
  lines.push("");
  lines.push("| " + sheet.headers.join(" | ") + " |");
  lines.push("|" + sheet.headers.map(() => "---|").join(""));
  for (const row of sheet.rows.slice(0, 20)) {
    const vals = row.map((c) => (c === null ? "" : String(c))).join(" | ");
    lines.push("| " + vals + " |");
  }
  if (sheet.totalRows > 20) {
    lines.push(`*... ${sheet.totalRows - 20} lignes supplémentaires masquées*`);
  }
  return lines.join("\n");
}

export async function parseFile(file: File): Promise<FileParseResult> {
  const buffer = await readFileAsArrayBuffer(file)

  if (file.name.endsWith('.csv')) {
    const text = new TextDecoder('utf-8').decode(buffer)
    const workbook = XLSX.read(text, { type: 'string', raw: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const parsed = sheetToParsedSheet(sheet, 'CSV')
    const formulas = detectFormulas(sheet);
    return {
      fileName: file.name,
      fileSize: file.size,
      sheets: [parsed],
      textRepresentation: buildTextRepresentation(parsed, formulas),
    }
  }

  const workbook = XLSX.read(buffer, { type: 'array', raw: true })
  const sheets: ParsedSheet[] = workbook.SheetNames.slice(0, 3).map((name) => {
    const sheet = workbook.Sheets[name]
    return sheetToParsedSheet(sheet, name)
  })

  const textRepresentations = sheets.map(s => {
    const sheet = workbook.Sheets[s.name];
    const formulas = detectFormulas(sheet);
    return buildTextRepresentation(s, formulas);
  });

  return {
    fileName: file.name,
    fileSize: file.size,
    sheets,
    textRepresentation: textRepresentations.join('\n\n'),
  }
}

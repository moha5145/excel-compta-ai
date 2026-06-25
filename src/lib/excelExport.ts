import ExcelJS from "exceljs";

// Dictionnaire complet de traduction Français -> Anglais des fonctions Excel courantes
const FRENCH_TO_ENGLISH_FUNCTIONS: Record<string, string> = {
  // Logique
  "SI.CONDITIONS": "IFS",
  "SI.NON.DISP": "IFNA",
  "SIERREUR": "IFERROR",
  "SI": "IF",
  "ESTERREUR": "ISERROR",
  "ESTERR": "ISERR",
  "ESTNA": "ISNA",
  "ESTVIDE": "ISBLANK",
  "ESTNUM": "ISNUMBER",
  "ESTTEXTE": "ISTEXT",
  "ET": "AND",
  "OU": "OR",
  "NON": "NOT",

  // Math & Statistiques
  "SOMME.SI.ENS": "SUMIFS",
  "SOMME.SI": "SUMIF",
  "SOMME": "SUM",
  "NB.SI.ENS": "COUNTIFS",
  "NB.SI": "COUNTIF",
  "NB.VIDE": "COUNTBLANK",
  "NBVAL": "COUNTA",
  "NB": "COUNT",
  "MOYENNE.SI.ENS": "AVERAGEIFS",
  "MOYENNE.SI": "AVERAGEIF",
  "MOYENNE": "AVERAGE",
  "GRANDE.VALEUR": "LARGE",
  "PETITE.VALEUR": "SMALL",
  "ARRONDI.SUP": "ROUNDUP",
  "ARRONDI.INF": "ROUNDDOWN",
  "ARRONDI": "ROUND",
  "SOMMEPROD": "SUMPRODUCT",
  "PRODUIT": "PRODUCT",
  "MAX": "MAX",
  "MIN": "MIN",
  "RANG": "RANK",
  "ENT": "INT",
  "ABS": "ABS",
  "MOD": "MOD",

  // Recherche & Tableaux
  "RECHERCHEV": "VLOOKUP",
  "RECHERCHEH": "HLOOKUP",
  "RECHERCHEX": "XLOOKUP",
  "EQUIVX": "XMATCH",
  "EQUIV": "MATCH",
  "INDEX": "INDEX",
  "CHOISIR": "CHOOSE",
  "DECALER": "OFFSET",
  "INDIRECT": "INDIRECT",
  "COLONNE": "COLUMN",
  "LIGNE": "ROW",
  "TRANSPOSE": "TRANSPOSE",
  "TRIERPAR": "SORTBY",
  "TRIER": "SORT",
  "UNIQUE": "UNIQUE",
  "FILTRE": "FILTER",

  // Finance
  "VPM": "PMT",
  "AMORLIN": "SLN",
  "AMORDEG": "DB",
  "DDB": "DDB",
  "VAN": "NPV",
  "TRI": "IRR",
  "TAUX": "RATE",
  "NPM": "NPER",
  "VC": "FV",
  "VA": "PV",

  // Texte
  "CONCATENER": "CONCATENATE",
  "CONCAT": "CONCAT",
  "TEXTE": "TEXT",
  "GAUCHE": "LEFT",
  "DROITE": "RIGHT",
  "STXT": "MID",
  "NBCAR": "LEN",
  "TROUVER": "FIND",
  "CHERCHER": "SEARCH",
  "SUBSTITUER": "SUBSTITUTE",
  "REMPLACER": "REPLACE",
  "EPURAGE": "CLEAN",
  "SUPPRESPACE": "TRIM",
  "MINUSCULE": "LOWER",
  "MAJUSCULE": "UPPER",
  "NOMPROPRE": "PROPER",

  // Date & Heure
  "AUJOURDHUI": "TODAY",
  "MAINTENANT": "NOW",
  "FIN.MOIS": "EOMONTH",
  "NB.JOURS.OUVRES": "NETWORKDAYS",
  "NO.SEMAINE": "WEEKNUM",
  "JOURSEM": "WEEKDAY",
  "DATEDIF": "DATEDIF",
  "ANNEE": "YEAR",
  "MOIS": "MONTH",
  "JOUR": "DAY",
  "DATE": "DATE",
  "JOURS": "DAYS",
};

// Nettoie la syntaxe Markdown pour l'affichage brut en cellule Excel
function cleanMarkdownFormatting(text: string): string {
  let cleaned = text
    .replace(/^\s*#{1,6}\s+/, "") // Titres #, ##, ###
    .replace(/^\s*---+\s*$/, "") // Séparateurs horizontaux ---
    .replace(/^\s*\*\*\*+\s*$/, "") // Séparateurs ***
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1") // Liens Markdown [texte](url) → texte
    .replace(/\*\*([^*]+)\*\*/g, "$1") // Gras **texte**
    .replace(/\*([^*]+)\*/g, "$1") // Italique *texte*
    .replace(/`([^`]+)`/g, "$1") // Code inline `code`
    .replace(/^\s*[-*]\s+/, "• ") // Listes puces
    .replace(/^\s*\d+\.\s+/, (m) => m.trim() + " ") // Listes numérotées
    .trim();
  return cleaned;
}

// Traduit les formules Excel du Français vers l'Anglais pour affichage (texte uniquement)
export function translateFrenchFormulaToEnglish(formula: string): string {
  let translated = formula.trim();

  // 1. Remplacer les noms de fonctions (les clés les plus longues d'abord pour éviter les conflits)
  const keys = Object.keys(FRENCH_TO_ENGLISH_FUNCTIONS).sort((a, b) => b.length - a.length);

  for (const key of keys) {
    // Échapper les points dans la clé pour la regex
    const escapedKey = key.replace(/\./g, "\\.");
    // Matcher le nom de fonction uniquement suivi d'une parenthèse ouvrante
    const regex = new RegExp(`\\b${escapedKey}(?=\\s*\\()`, "gi");
    translated = translated.replace(regex, FRENCH_TO_ENGLISH_FUNCTIONS[key]);
  }

  // 2. Remplacer les séparateurs d'arguments : ";" → ","
  // et les séparateurs décimaux : "," → "." uniquement hors chaînes de caractères
  let inString = false;
  let finalFormula = "";
  let i = 0;
  while (i < translated.length) {
    const char = translated[i];

    if (char === '"') {
      inString = !inString;
      finalFormula += char;
      i++;
      continue;
    }

    if (!inString) {
      if (char === ";") {
        finalFormula += ",";
      } else if (char === ",") {
        // Vérifier si c'est un séparateur décimal (chiffre avant ET chiffre/% après)
        const prev = translated[i - 1];
        const next = translated[i + 1];
        const isDecimalSeparator = prev && /\d/.test(prev) && next && /[\d%]/.test(next);
        if (isDecimalSeparator) {
          finalFormula += ".";
        } else {
          // C'est une virgule ordinaire (déjà un séparateur en anglais), conserver
          finalFormula += ",";
        }
      } else {
        finalFormula += char;
      }
    } else {
      finalFormula += char;
    }
    i++;
  }

  return finalFormula;
}

// Extrait la formule depuis le code Markdown
export function extractFormula(markdown: string): string {
  // Recherche d'un bloc de code avec une formule
  const codeBlockRegex = /```(?:excel|sheets|xlsx|txt|vba)?\s*\n([\s\S]*?)\n\s*```/gi;
  let match;
  while ((match = codeBlockRegex.exec(markdown)) !== null) {
    const content = match[1].trim();
    if (content.startsWith("=")) {
      return content;
    }
  }

  // Fallback : ligne qui commence par = (avec une longueur minimum pour éviter les faux positifs)
  const lines = markdown.split("\n");
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.startsWith("=") && trimmed.length > 4 && /[A-Z(]/.test(trimmed.substring(1, 5))) {
      return trimmed;
    }
  }

  return "";
}

interface ExtractedTable {
  headers: string[];
  rows: string[][];
}

// Extrait les tableaux Markdown pour en faire des données structurées
export function extractTables(markdown: string): ExtractedTable[] {
  const tables: ExtractedTable[] = [];
  const lines = markdown.split("\n");
  let currentTable: ExtractedTable | null = null;
  let inTable = false;
  let inCodeBlock = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line.startsWith("```")) { inCodeBlock = !inCodeBlock; continue; }
    if (inCodeBlock) continue;

    const isTableRow = line.startsWith("|") && line.endsWith("|");

    if (isTableRow) {
      const cells = line
        .split("|")
        .map((cell) => {
          let c = cell.trim();
          c = c.replace(/\*\*([^*]+)\*\*/g, "$1"); // Gras **texte**
          c = c.replace(/\*([^*]+)\*/g, "$1"); // Italique *texte*
          c = c.replace(/`([^`]+)`/g, "$1"); // Code inline `code`
          c = c.replace(/^'([^']+)'$/, "$1"); // Quotes simples 'texte'
          return c.trim();
        })
        .filter((_, idx, arr) => idx > 0 && idx < arr.length - 1);

      const isSeparator = cells.every((cell) => /^\s*:?-+:?\s*$/.test(cell));
      if (isSeparator) { continue; }

      if (!inTable) {
        inTable = true;
        currentTable = { headers: cells, rows: [] };
      } else if (currentTable) {
        currentTable.rows.push(cells);
      }
    } else {
      if (inTable && currentTable) {
        tables.push(currentTable);
        currentTable = null;
        inTable = false;
      }
    }
  }
  if (inTable && currentTable) { tables.push(currentTable); }
  return tables;
}

// Nettoie le Markdown et retourne un tableau de { text, isHeader, isSub }
type LineType = { text: string; isHeader: boolean; isBullet: boolean; isEmpty: boolean };

function parseExplanationLines(markdown: string): LineType[] {
  const lines = markdown.split("\n");
  const result: LineType[] = [];
  let inCodeBlock = false;

  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.startsWith("```")) { inCodeBlock = !inCodeBlock; continue; }
    if (inCodeBlock) continue;

    // Ignorer les lignes de tableau
    if (trimmed.startsWith("|") && trimmed.endsWith("|")) continue;
    // Ignorer les séparateurs horizontaux purs
    if (/^-{3,}$/.test(trimmed) || /^\*{3,}$/.test(trimmed)) continue;

    if (!trimmed) {
      result.push({ text: "", isHeader: false, isBullet: false, isEmpty: true });
      continue;
    }

    const isHeader = /^#{1,6}\s+/.test(trimmed);
    const isBullet = /^[-*]\s+/.test(trimmed) || /^\d+\.\s+/.test(trimmed);
    const cleaned = cleanMarkdownFormatting(trimmed);

    if (cleaned) {
      result.push({ text: cleaned, isHeader, isBullet, isEmpty: false });
    }
  }
  return result;
}

export async function downloadFormulaAsExcel(
  response: string,
  prompt: string
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Excel-Compta AI";
  workbook.created = new Date();
  workbook.modified = new Date();

  const formulaRaw = extractFormula(response);
  const tables = extractTables(response);
  const explanationLines = parseExplanationLines(response);
  // Version anglaise pour affichage (texte uniquement, pas de formule active)
  const formulaEnglish = formulaRaw ? translateFrenchFormulaToEnglish(formulaRaw) : "";

  // ─────────────────────────────────────────────────────────────
  // CRÉATION DES ONGLETS (l'ordre détermine l'ordre d'affichage)
  // ─────────────────────────────────────────────────────────────
  const sheet2 = workbook.addWorksheet("Exemple Pratique");
  const sheet1 = workbook.addWorksheet("Formule & Guide");

  // ─────────────────────────────────────────────────────────────
  // ONGLET : FORMULE & GUIDE
  // ─────────────────────────────────────────────────────────────
  sheet1.views = [{ showGridLines: true }];

  // Helpers styles
  const SLATE_900 = "FF0F172A";
  const SLATE_700 = "FF334155";
  const SLATE_500 = "FF64748B";
  const YELLOW    = "FFEAB308";
  const WHITE     = "FFFFFFFF";
  const AMBER     = "FFCA8A04";
  const LIGHT_BG  = "FFF8FAFC";

  // ── Titre
  sheet1.mergeCells("A1:H1");
  const t1 = sheet1.getCell("A1");
  t1.value = "EXCEL-COMPTA AI — FORMULE COMPTABLE";
  t1.font = { name: "Segoe UI", size: 16, bold: true, color: { argb: WHITE } };
  t1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: SLATE_900 } };
  t1.alignment = { vertical: "middle", horizontal: "center" };
  sheet1.getRow(1).height = 42;

  // ── Barre jaune déco
  sheet1.mergeCells("A2:H2");
  sheet1.getCell("A2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW } };
  sheet1.getRow(2).height = 5;

  // ── Demande utilisateur
  const setLabel = (cell: string, text: string) => {
    sheet1.getCell(cell).value = text;
    sheet1.getCell(cell).font = { name: "Segoe UI", size: 10, bold: true, color: { argb: SLATE_700 } };
    sheet1.getCell(cell).alignment = { vertical: "top" };
  };

  setLabel("A4", "Demande :");
  sheet1.mergeCells("B4:H4");
  const demCell = sheet1.getCell("B4");
  demCell.value = prompt;
  demCell.font = { name: "Segoe UI", size: 10, italic: true, color: { argb: SLATE_700 } };
  demCell.alignment = { wrapText: true, vertical: "top" };
  sheet1.getRow(4).height = 24;

  // ── Formule version française (telle que générée)
  setLabel("A6", "Formule\n(version FR) :");
  sheet1.getCell("A6").alignment = { wrapText: true, vertical: "top" };
  sheet1.mergeCells("B6:H6");
  const frCell = sheet1.getCell("B6");
  frCell.value = formulaRaw || "Aucune formule détectée";
  frCell.font = { name: "Consolas", size: 11, bold: true, color: { argb: AMBER } };
  frCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT_BG } };
  frCell.alignment = { wrapText: true, vertical: "middle" };
  frCell.border = {
    top: { style: "thin", color: { argb: "FFE2E8F0" } },
    bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
    left: { style: "thin", color: { argb: "FFE2E8F0" } },
    right: { style: "thin", color: { argb: "FFE2E8F0" } },
  };
  sheet1.getRow(6).height = 30;

  // ── Formule version anglaise (à copier pour Excel anglais / Google Sheets)
  if (formulaEnglish && formulaEnglish !== formulaRaw) {
    setLabel("A7", "Formule\n(version EN) :");
    sheet1.getCell("A7").alignment = { wrapText: true, vertical: "top" };
    sheet1.mergeCells("B7:H7");
    const enCell = sheet1.getCell("B7");
    enCell.value = formulaEnglish;
    enCell.font = { name: "Consolas", size: 11, bold: true, color: { argb: "FF16803C" } }; // Vert
    enCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0FDF4" } }; // Green 50
    enCell.alignment = { wrapText: true, vertical: "middle" };
    enCell.border = {
      top: { style: "thin", color: { argb: "FFD1FAE5" } },
      bottom: { style: "thin", color: { argb: "FFD1FAE5" } },
      left: { style: "thin", color: { argb: "FFD1FAE5" } },
      right: { style: "thin", color: { argb: "FFD1FAE5" } },
    };
    sheet1.getRow(7).height = 30;

    // Note sous la formule anglaise
    sheet1.mergeCells("B8:H8");
    const noteCell = sheet1.getCell("B8");
    noteCell.value = "ℹ️  Version anglaise : à utiliser si votre Excel est en anglais ou pour Google Sheets. Collez-la dans une cellule vide.";
    noteCell.font = { name: "Segoe UI", size: 9, italic: true, color: { argb: SLATE_500 } };
    noteCell.alignment = { wrapText: true };
    sheet1.getRow(8).height = 18;
  }

  // ── Séparateur avant les explications
  const sepRow = formulaEnglish && formulaEnglish !== formulaRaw ? 10 : 9;
  setLabel(`A${sepRow}`, "Explication :");

  let currentLineIdx = sepRow;
  let prevEmpty = false;
  for (const entry of explanationLines) {
    if (entry.isEmpty) {
      if (!prevEmpty) { currentLineIdx++; }
      prevEmpty = true;
      continue;
    }
    prevEmpty = false;

    sheet1.mergeCells(`B${currentLineIdx}:H${currentLineIdx}`);
    const lineCell = sheet1.getCell(`B${currentLineIdx}`);
    lineCell.value = entry.text;
    lineCell.alignment = { wrapText: true, vertical: "top" };

    if (entry.isHeader) {
      lineCell.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: SLATE_900 } };
      sheet1.getRow(currentLineIdx).height = 22;
    } else if (entry.isBullet) {
      lineCell.font = { name: "Segoe UI", size: 10, color: { argb: SLATE_700 } };
      sheet1.getRow(currentLineIdx).height = 18;
    } else {
      lineCell.font = { name: "Segoe UI", size: 10, color: { argb: SLATE_700 } };
      sheet1.getRow(currentLineIdx).height = 18;
    }
    currentLineIdx++;
  }

  // Largeurs colonnes onglet 1
  sheet1.getColumn(1).width = 18; // Labels
  for (let c = 2; c <= 8; c++) {
    sheet1.getColumn(c).width = 18;
  }

  // ─────────────────────────────────────────────────────────────
  // ONGLET : EXEMPLE PRATIQUE
  // ─────────────────────────────────────────────────────────────
  sheet2.views = [{ showGridLines: true }];

  // ── Titre
  sheet2.mergeCells("A1:H1");
  const t2 = sheet2.getCell("A1");
  t2.value = "EXEMPLE PRATIQUE DE SIMULATION";
  t2.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: WHITE } };
  t2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
  t2.alignment = { vertical: "middle", horizontal: "center" };
  sheet2.getRow(1).height = 36;

  // ── Barre jaune déco
  sheet2.mergeCells("A2:H2");
  sheet2.getCell("A2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW } };
  sheet2.getRow(2).height = 4;

  let nextRow = 4;

  // ── Section formules à copier-coller (point principal de valeur)
  const writeSection = (row: number, label: string, value: string, color: string, bgColor: string) => {
    const labelCell = sheet2.getCell(`B${row}`);
    labelCell.value = label;
    labelCell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: SLATE_500 } };

    sheet2.mergeCells(`C${row}:H${row}`);
    const valCell = sheet2.getCell(`C${row}`);
    valCell.value = value;
    valCell.font = { name: "Consolas", size: 11, bold: true, color: { argb: color } };
    valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
    valCell.alignment = { wrapText: true, vertical: "middle" };
    valCell.border = {
      top: { style: "thin", color: { argb: "FFE2E8F0" } },
      bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
      left: { style: "thin", color: { argb: "FFE2E8F0" } },
      right: { style: "thin", color: { argb: "FFE2E8F0" } },
    };
    sheet2.getRow(row).height = 28;
  };

  // Titre de la section formule
  sheet2.mergeCells(`B${nextRow}:H${nextRow}`);
  const secTitle = sheet2.getCell(`B${nextRow}`);
  secTitle.value = "📋  Formule prête à coller";
  secTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: SLATE_900 } };
  sheet2.getRow(nextRow).height = 22;
  nextRow++;

  if (formulaRaw) {
    writeSection(nextRow, "Version FR :", formulaRaw, AMBER, LIGHT_BG);
    nextRow++;
  }

  if (formulaEnglish && formulaEnglish !== formulaRaw) {
    writeSection(nextRow, "Version EN :", formulaEnglish, "FF16803C", "FFF0FDF4");
    nextRow++;
  }

  // Note instructions
  sheet2.mergeCells(`B${nextRow}:H${nextRow}`);
  const instrCell = sheet2.getCell(`B${nextRow}`);
  instrCell.value = "👆  Copiez la formule ci-dessus et collez-la dans une cellule vide de votre classeur Excel.";
  instrCell.font = { name: "Segoe UI", size: 9, italic: true, color: { argb: SLATE_500 } };
  sheet2.getRow(nextRow).height = 18;
  nextRow += 2;

  // ── Tableaux d'exemples (si détectés dans la réponse IA)
  if (tables.length > 0) {
    sheet2.mergeCells(`B${nextRow}:H${nextRow}`);
    const tableTitle = sheet2.getCell(`B${nextRow}`);
    tableTitle.value = "📊  Données d'exemple (extraites de la réponse IA)";
    tableTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: SLATE_900 } };
    sheet2.getRow(nextRow).height = 22;
    nextRow++;

    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];

      if (tables.length > 1) {
        sheet2.getCell(`B${nextRow}`).value = `Tableau n°${t + 1}`;
        sheet2.getCell(`B${nextRow}`).font = { name: "Segoe UI", size: 10, bold: true, color: { argb: SLATE_500 } };
        sheet2.getRow(nextRow).height = 18;
        nextRow++;
      }

      // En-têtes
      const headerRow = sheet2.getRow(nextRow);
      headerRow.height = 24;
      for (let h = 0; h < table.headers.length; h++) {
        const cell = headerRow.getCell(h + 2);
        cell.value = table.headers[h];
        cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: WHITE } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin", color: { argb: "FF475569" } },
          bottom: { style: "medium", color: { argb: "FF1E293B" } },
          left: { style: "thin", color: { argb: "FF475569" } },
          right: { style: "thin", color: { argb: "FF475569" } },
        };
      }
      nextRow++;

      // Données
      for (let r = 0; r < table.rows.length; r++) {
        const dataRow = sheet2.getRow(nextRow);
        dataRow.height = 20;
        const rowData = table.rows[r];

        for (let c = 0; c < rowData.length; c++) {
          const cell = dataRow.getCell(c + 2);
          const rawVal = rowData[c];
          const numVal = Number(rawVal.replace(/\s/g, "").replace(",", "."));
          if (!isNaN(numVal) && rawVal.trim() !== "") {
            cell.value = numVal;
            cell.numFmt = "#,##0.##";
            cell.alignment = { horizontal: "right" };
          } else {
            cell.value = rawVal;
            cell.alignment = { horizontal: "left" };
          }
          cell.font = { name: "Segoe UI", size: 10 };
          if (r % 2 === 1) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
          }
          cell.border = {
            top: { style: "thin", color: { argb: "FFE2E8F0" } },
            bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
            left: { style: "thin", color: { argb: "FFE2E8F0" } },
            right: { style: "thin", color: { argb: "FFE2E8F0" } },
          };
        }
        nextRow++;
      }
      nextRow += 2;
    }
  } else {
    // Pas de tableau : juste un message clair et de l'espace vide
    sheet2.mergeCells(`B${nextRow}:H${nextRow}`);
    const noTableMsg = sheet2.getCell(`B${nextRow}`);
    noTableMsg.value = "ℹ️  Aucun tableau de données n'a été détecté dans la réponse IA.";
    noTableMsg.font = { name: "Segoe UI", size: 10, italic: true, color: { argb: SLATE_500 } };
    sheet2.getRow(nextRow).height = 18;
    nextRow++;

    sheet2.mergeCells(`B${nextRow}:H${nextRow}`);
    const noTableHint = sheet2.getCell(`B${nextRow}`);
    noTableHint.value = "   Vous pouvez créer votre propre tableau ici et utiliser la formule ci-dessus.";
    noTableHint.font = { name: "Segoe UI", size: 9, color: { argb: SLATE_500 } };
    sheet2.getRow(nextRow).height = 18;
    nextRow += 2;
  }

  // Largeurs colonnes onglet 2
  sheet2.getColumn(1).width = 4;
  sheet2.getColumn(2).width = 22;
  for (let c = 3; c <= 8; c++) {
    sheet2.getColumn(c).width = 18;
  }

  // ─────────────────────────────────────────────────────────────
  // GÉNÉRATION ET TÉLÉCHARGEMENT
  // ─────────────────────────────────────────────────────────────
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;

  const sanitizedPrompt = prompt
    .toLowerCase()
    .replace(/[éèêë]/g, "e").replace(/[àâä]/g, "a").replace(/[îï]/g, "i").replace(/[ôö]/g, "o").replace(/[ùûü]/g, "u").replace(/ç/g, "c")
    .replace(/[^a-z0-9]+/g, "-")
    .slice(0, 35);
  a.download = `formule-${sanitizedPrompt || "excel"}-${Date.now()}.xlsx`;

  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

# Plan d'implémentation : Upload de fichier Excel

## Architecture

Le fichier est parsé **côté client** (navigateur), converti en représentation texte structurée, puis envoyé comme contexte supplémentaire dans la conversation existante. Aucune nouvelle route API n'est nécessaire.

### Flux
```
Fichier .xlsx/.csv
    ↓
[FileUpload] → lit avec SheetJS → extrait 1ère feuille
    ↓
Conversion en tableau texte (JSON + Markdown)
    ↓
Affichage preview + bouton "Analyser avec l'IA"
    ↓
Concatène le prompt utilisateur + données du fichier
    ↓
POST /api/gemini (existant, inchangé)
```

### Fichiers modifiés/créés

| Fichier | Action | Rôle |
|---|---|---|
| `package.json` | Modifier | Ajouter `xlsx` (SheetJS) |
| `src/components/FileUpload.tsx` | Créer | Zone de drop + sélecteur fichier + preview |
| `src/lib/fileParser.ts` | Créer | Parse .xlsx/.csv en texte structuré |
| `src/app/page.tsx` | Modifier | Ajouter l'état du fichier uploadé + intégrer FileUpload |
| `src/components/FormulaAssistant.tsx` | Modifier | Passer le contexte fichier dans l'envoi |
| `src/app/api/gemini/route.ts` | Modifier | Adapter le system prompt pour contexte fichier |

---

### Tâche 1 : Ajouter la dépendance SheetJS

- [ ] **Étape 1 : Installer xlsx**

```bash
npm install xlsx
```

- [ ] **Étape 2 : Vérifier l'installation**

```bash
node -e "require('xlsx'); console.log('ok')"
```

Attendu : `ok`

---

### Tâche 2 : Créer le parseur de fichiers (`src/lib/fileParser.ts`)

**Fichier :** Créer `src/lib/fileParser.ts`

Responsabilité : Lire un fichier `.xlsx` ou `.csv` côté client, extraire les données de la 1ère feuille, retourner un objet structuré.

- [ ] **Étape 1 : Implémenter `parseExcelFile` et `parseCSVFile`**

```typescript
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
    totalRows: data.length - 1,
    totalCols: headers.length,
  }
}

function buildTextRepresentation(sheet: ParsedSheet): string {
  const lines: string[] = []
  lines.push(`[${sheet.name}]`)
  lines.push('| ' + sheet.headers.join(' | ') + ' |')
  lines.push('|' + sheet.headers.map(() => '---|').join(''))
  for (const row of sheet.rows.slice(0, 20)) {
    const vals = row.map((c) => (c === null ? '' : String(c))).join(' | ')
    lines.push('| ' + vals + ' |')
  }
  if (sheet.rows.length > 20) {
    lines.push(`*... ${sheet.rows.length - 20} lignes supplémentaires*`)
  }
  return lines.join('\n')
}

export async function parseFile(file: File): Promise<FileParseResult> {
  const buffer = await readFileAsArrayBuffer(file)

  if (file.name.endsWith('.csv')) {
    const text = new TextDecoder('utf-8').decode(buffer)
    const workbook = XLSX.read(text, { type: 'string', raw: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const parsed = sheetToParsedSheet(sheet, 'CSV')
    return {
      fileName: file.name,
      fileSize: file.size,
      sheets: [parsed],
      textRepresentation: buildTextRepresentation(parsed),
    }
  }

  const workbook = XLSX.read(buffer, { type: 'array', raw: true })
  const sheets: ParsedSheet[] = workbook.SheetNames.slice(0, 1).map((name) => {
    const sheet = workbook.Sheets[name]
    return sheetToParsedSheet(sheet, name)
  })

  return {
    fileName: file.name,
    fileSize: file.size,
    sheets,
    textRepresentation: sheets.map(buildTextRepresentation).join('\n\n'),
  }
}
```

- [ ] **Étape 2 : Vérifier que le module s'importe sans erreur**

Ajouter temporairement dans `src/app/page.tsx` :
```typescript
import { parseFile } from '@/lib/fileParser'
console.log('fileParser loaded')
```
Puis `npm run build` — doit passer. Supprimer après vérification.

---

### Tâche 3 : Créer le composant FileUpload (`src/components/FileUpload.tsx`)

**Fichier :** Créer `src/components/FileUpload.tsx`

Responsabilité : Zone de drop, sélecteur de fichier, validation du type, affichage de la preview avec bouton "retirer".

```typescript
'use client'

import { useCallback, useRef, useState } from 'react'
import { Upload, FileSpreadsheet, X, AlertCircle } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { cn } from '@/lib/utils'

interface FileUploadProps {
  onFileParsed: (result: { fileName: string; textRepresentation: string } | null) => void
  disabled?: boolean
}

const ACCEPTED_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'text/csv',
  '.xlsx',
  '.xls',
  '.csv',
]

export function FileUpload({ onFileParsed, disabled }: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [file, setFile] = useState<File | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)

  const handleFile = useCallback(async (f: File) => {
    setError(null)
    const ext = f.name.split('.').pop()?.toLowerCase()
    if (!ext || !['xlsx', 'xls', 'csv'].includes(ext)) {
      setError('Format accepté : .xlsx, .xls, .csv')
      return
    }
    if (f.size > 10 * 1024 * 1024) {
      setError('Fichier trop volumineux (max 10 Mo)')
      return
    }
    setLoading(true)
    try {
      const { parseFile } = await import('@/lib/fileParser')
      const result = await parseFile(f)
      setFile(f)
      onFileParsed({
        fileName: f.name,
        textRepresentation: result.textRepresentation,
      })
    } catch {
      setError('Erreur lors de la lecture du fichier')
      onFileParsed(null)
    } finally {
      setLoading(false)
    }
  }, [onFileParsed])

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      const f = e.dataTransfer.files[0]
      if (f) handleFile(f)
    },
    [handleFile]
  )

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0]
      if (f) handleFile(f)
    },
    [handleFile]
  )

  const remove = useCallback(() => {
    setFile(null)
    setError(null)
    onFileParsed(null)
    if (inputRef.current) inputRef.current.value = ''
  }, [onFileParsed])

  if (file) {
    return (
      <div className="flex items-center gap-2 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-2 text-sm dark:border-emerald-800 dark:bg-emerald-950/30">
        <FileSpreadsheet className="h-4 w-4 text-emerald-600 dark:text-emerald-400" />
        <span className="flex-1 truncate text-emerald-700 dark:text-emerald-300">
          {file.name}
        </span>
        <button
          onClick={remove}
          className="rounded p-0.5 text-emerald-500 hover:bg-emerald-200 dark:hover:bg-emerald-800"
        >
          <X className="h-3.5 w-3.5" />
        </button>
      </div>
    )
  }

  return (
    <div>
      <div
        onDrop={handleDrop}
        onDragOver={(e) => e.preventDefault()}
        onClick={() => inputRef.current?.click()}
        className={cn(
          'flex cursor-pointer items-center gap-2 rounded-lg border-2 border-dashed px-3 py-2 text-sm transition-colors',
          error
            ? 'border-red-300 bg-red-50 text-red-600 dark:border-red-800 dark:bg-red-950/30 dark:text-red-400'
            : 'border-muted-foreground/25 text-muted-foreground hover:border-muted-foreground/50 hover:bg-muted/50',
          disabled && 'pointer-events-none opacity-50'
        )}
      >
        {loading ? (
          <span className="text-xs">Lecture en cours...</span>
        ) : (
          <>
            <Upload className="h-4 w-4 shrink-0" />
            <span className="truncate">{error || 'Ajouter un fichier Excel'}</span>
          </>
        )}
      </div>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleChange}
        disabled={disabled}
      />
    </div>
  )
}
```

---

### Tâche 4 : Intégrer FileUpload dans la page principale

**Fichier :** Modifier `src/app/page.tsx`

- [ ] **Étape 1 : Ajouter l'état du fichier dans le composant Home**

Après la ligne `const [isEnhancing, setIsEnhancing] = useState(false)`, ajouter :

```typescript
const [fileContext, setFileContext] = useState<{
  fileName: string
  textRepresentation: string
} | null>(null)
```

- [ ] **Étape 2 : Importer et intégrer FileUpload dans le rendu**

Ajouter l'import en haut :
```typescript
import { FileUpload } from '@/components/FileUpload'
```

Ajouter le FileUpload dans le rendu, juste après la `FormulaInputBar` et avant la `FormulaResultArea`. Chercher le bloc qui contient la `FormulaInputBar` et `FormulaResultArea` (vers la fin du render), et ajouter :

```tsx
<FileUpload
  onFileParsed={(result) => setFileContext(result)}
  disabled={isGenerating}
/>
```

- [ ] **Étape 3 : Passer le fileContext dans le prompt envoyé à l'API**

Dans la fonction `handleGenerate`, avant l'appel à `fetch('/api/gemini', {...})`, modifier le `body` :

```typescript
let fileContextSection = ''
if (fileContext) {
  fileContextSection = `\n\nFichier chargé : "${fileContext.fileName}"\nDonnées du fichier :\n\`\`\`\n${fileContext.textRepresentation}\n\`\`\``
}

const body = JSON.stringify({
  messages: [
    ...(messages.length > 0
      ? messages.map((m) => ({ role: m.role, content: m.content }))
      : []),
    { role: 'user', content: userPrompt + fileContextSection },
  ],
  ...
})
```

- [ ] **Étape 4 : Vider le context fichier au "Nouvelle conversation"**

Dans le gestionnaire de "Nouvelle conversation", ajouter :
```typescript
setFileContext(null)
```

---

### Tâche 5 : Adapter le system prompt Gemini

**Fichier :** Modifier `src/app/api/gemini/route.ts`

- [ ] **Étape 1 : Ajouter des instructions pour le contexte fichier**

Dans le system instruction (lignes 81-136), ajouter vers la fin, avant la ligne de vérification :

```
SPREADSHEET FILE CONTEXT:
Si l'utilisateur fournit des données de fichier (tableau markdown avec en-têtes et valeurs) :
1. Analyse la structure : colonnes, types de données, relations
2. Utilise ces données pour formuler des formules pertinentes
3. Les colonnes commencent à la colonne A pour les données fournies
4. Si une simulation est demandée sur les données du fichier, utilise les valeurs fournies comme données d'entrée
5. La simulation table doit refléter la structure du fichier importé
```

---

### Tâche 6 : Build & vérification

- [ ] **Étape 1 : Build**

```bash
npm run build
```

Attendu : succès, pas d'erreurs TypeScript.

- [ ] **Étape 2 : Test manuel**

1. `npm run dev`
2. Ouvrir http://localhost:3000
3. Créer un petit fichier .xlsx avec 2-3 colonnes (Nom, Âge, Ville) et 3-4 lignes
4. Glisser-déposer sur la zone d'upload → voir le nom du fichier apparaître
5. Taper "Calcule la moyenne d'âge" → voir la réponse utiliser les données
6. Cliquer "Nouvelle conversation" → vérifier que le fichier est retiré

---

### Tâche 7 : Nettoyage

- [ ] **Étape 1 : Supprimer le code temporaire** (console.log de test dans page.tsx si ajouté)

- [ ] **Étape 2 : Commit**

```bash
git add -A
git commit -m "feat: ajout import fichier Excel pour analyse par l'IA"
```

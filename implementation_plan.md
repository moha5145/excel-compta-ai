# Implementation Plan - Mode Selector for Formula & Table Generation

Allow users to explicitly select between 3 generation modes:
1. **Formule seule** (`formula_only`): Quick copy-paste formula and explanation, no tables.
2. **Formule + Tableau simple** (`simple_table`): Formula, explanation, and a clean Markdown demonstration table.
3. **Formule + Tableau complexe** (`complex_table`): Full interactive simulation table with multi-parameter schema and formatted Excel workbook export.

---

## User Review Required

> [!IMPORTANT]
> - **Default Mode**: Set to `formula_only` upon opening the app to provide immediate, lightweight formula generation.
> - **Quick Examples**: Selecting an example from the dropdown will automatically switch to the most appropriate mode for that specific scenario (e.g. VAT calculation $\rightarrow$ `formula_only`, VNC Depreciation $\rightarrow$ `complex_table`).
> - **Export Button**: The `.xlsx` download button will be hidden in `formula_only` mode in favor of a prominent **"Copier la formule"** action button.

---

## Proposed Changes

### UI & Component Layer

#### [MODIFY] [FormulaAssistant.tsx](file:///home/mo-pro/Bureau/dev/excel-formule-ai/src/components/FormulaAssistant.tsx)
- Add state and types for `GenerationMode = 'formula_only' | 'simple_table' | 'complex_table'`.
- Add a 3-button segmented selector in `FormulaInputBar` above the textarea with intuitive icons (`Code2`, `Table`, `Calculator`).
- Attach suggested `defaultMode` properties to items in `FORMULA_EXAMPLES`.
- Conditionally render the Excel `.xlsx` Download button in `FormulaResultArea` based on `generationMode` (hidden for `formula_only`, shown for `simple_table` and `complex_table`).

---

### Backend API Layer

#### [MODIFY] [route.ts](file:///home/mo-pro/Bureau/dev/excel-formule-ai/src/app/api/gemini/route.ts)
- Accept `generationMode` parameter from the client payload.
- Adjust `systemInstruction` dynamically based on `generationMode`:
  - `formula_only`: Instruct Gemini to output ONLY the formula code block and bulleted explanation. Prohibit Markdown tables and JSON schemas.
  - `simple_table`: Instruct Gemini to include a 3-5 row Markdown table with realistic values, but skip `<!-- TABLE_SCHEMA -->`.
  - `complex_table`: Include the full `<!-- TABLE_SCHEMA -->` instructions for interactive multi-parameter simulation tables.

---

### App Page & State Layer

#### [MODIFY] [page.tsx](file:///home/mo-pro/Bureau/dev/excel-formule-ai/src/app/page.tsx)
- Manage `generationMode` state and pass it down to `FormulaInputBar` and `FormulaResultArea`.
- Include `generationMode` in the POST request body sent to `/api/gemini`.

---

## Verification Plan

### Automated Tests
- Run `npm run lint` and `npm run build` to verify TypeScript types and Next.js compilation.

### Manual Verification
1. Open the web app (`npm run dev`).
2. Test **Formule seule**: Verify quick response without tables, check that the "Copier la formule" button is primary and `.xlsx` download is hidden.
3. Test **Tableau simple**: Verify a Markdown table is returned and `.xlsx` export works.
4. Test **Tableau complexe**: Verify interactive parameters & simulation zone render, and complex `.xlsx` download works.
5. Test **Quick Examples**: Select various examples and verify mode auto-switching.

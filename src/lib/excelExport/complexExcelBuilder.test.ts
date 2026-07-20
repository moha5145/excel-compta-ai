import { describe, it, expect } from "vitest";
import ExcelJS from "exceljs";
import { buildComplexWorkbook } from "./complexExcelBuilder";
import { type TableSchema } from "./schemaParser";

describe("complexExcelBuilder", () => {
  it("devrait générer un workbook complet sans crasher", () => {
    const workbook = new ExcelJS.Workbook();
    const schema: TableSchema = {
      type: "complex_table",
      title: "Tableau d'amortissement",
      parameters: [
        { name: "Montant du prêt", ref: "C5", value: 500000, type: "currency" },
        { name: "Taux annuel", ref: "C6", value: 0.012, type: "percentage" },
      ],
      columns: [
        { header: "Mois", type: "integer", formula: null, formula_en: null },
        { header: "Intérêts", type: "currency", formula: "=C{row}*$C$6/12", formula_en: "=C{row}*$C$6/12" },
      ],
      data_start_row: 10,
      sample_rows: 3,
    };

    const response = `
    Voici les explications.
    | Mois | Intérêts |
    |---|---|
    | 1 | 500 |
    | 2 | 499 |
    <!-- TABLE_SCHEMA: {"type": "complex_table"} -->
    `;

    const { workbook: result, warnings } = buildComplexWorkbook(
      workbook,
      schema,
      response,
      "Simuler prêt de 500k€",
      "excel-fr"
    );

    expect(result.worksheets.length).toBe(2);
    expect(result.getWorksheet("Tableau Interactif")).toBeDefined();
    expect(result.getWorksheet("Formule & Guide")).toBeDefined();
    expect(warnings.length).toBe(0);

    const sheet = result.getWorksheet("Tableau Interactif")!;
    // Vérifier les paramètres
    expect(sheet.getCell("C5").value).toBe(500000);
    // Vérifier les en-têtes
    expect(sheet.getCell("B10").value).toBe("Ligne");
    expect(sheet.getCell("C10").value).toBe("Mois");
    expect(sheet.getCell("D10").value).toBe("Intérêts");

    // Vérifier les données
    expect(sheet.getCell("B11").value).toBe("Ligne 1");
    // Mois (input)
    expect(sheet.getCell("C11").value).toBe(1);
    // Intérêts (formule)
    expect(sheet.getCell("D11").value).toEqual({ formula: "C11*$C$6/12" });
  });
});

import { describe, it, expect } from "vitest";
import { resolveFormulaTemplate, convertToUsInvariant } from "./postProcessFormula";

describe("postProcessFormula", () => {
  describe("convertToUsInvariant", () => {
    it("devrait traduire les fonctions françaises en anglais et changer les séparateurs", () => {
      const frFormula = "=SI(A1; SOMME(B1:B10); 0)";
      expect(convertToUsInvariant(frFormula, "excel-fr")).toBe("=IF(A1, SUM(B1:B10), 0)");
    });

    it("devrait remplacer VRAI/FAUX par TRUE/FALSE", () => {
      expect(convertToUsInvariant("=SI(A1; VRAI; FAUX)", "excel-fr")).toBe("=IF(A1, TRUE, FALSE)");
    });
  });

  describe("resolveFormulaTemplate", () => {
    it("devrait remplacer {row} par le numéro de ligne", () => {
      const template = "=A{row}+B{row}";
      expect(resolveFormulaTemplate(template, 10, "excel-en", true)).toBe("=A10+B10");
    });

    it("devrait convertir la formule en anglais si isAlreadyEnglish est false", () => {
      const template = "=SI(A{row}; SOMME(B{row}:C{row}); 0)";
      expect(resolveFormulaTemplate(template, 10, "excel-fr", false)).toBe("=IF(A10, SUM(B10:C10), 0)");
    });

    it("devrait appliquer les prefixes _xlfn si format n'est pas sheets", () => {
      const template = "=IFS(A{row}=1, 1, A{row}=2, 2)";
      expect(resolveFormulaTemplate(template, 10, "excel-en", true)).toBe("=_xlfn.IFS(A10=1, 1, A10=2, 2)");
    });

    it("devrait convertir XLOOKUP en INDEX/MATCH pour LibreOffice", () => {
      const template = "=XLOOKUP(A{row}, B1:B10, C1:C10)";
      // Pour libreoffice-en
      expect(resolveFormulaTemplate(template, 10, "libreoffice-en", true)).toBe(
        "=INDEX(C1:C10, MATCH(A10, B1:B10, 0))"
      );
      // Pour libreoffice-fr
      expect(resolveFormulaTemplate(template, 10, "libreoffice-fr", true)).toBe(
        "=INDEX(C1:C10, MATCH(A10, B1:B10, 0))"
      );
    });
  });
});

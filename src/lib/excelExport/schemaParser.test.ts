import { describe, it, expect } from "vitest";
import {
  extractTableSchema,
  validateSchema,
  isComplexResponse,
  SchemaValidationError,
} from "./schemaParser";

describe("schemaParser", () => {
  describe("isComplexResponse", () => {
    it("devrait retourner true si le commentaire TABLE_SCHEMA est présent", () => {
      const response = `
      Voici le tableau :
      <!-- TABLE_SCHEMA: {"type": "complex_table", "title": "Test"} -->
      `;
      expect(isComplexResponse(response)).toBe(true);
    });

    it("devrait retourner false si le commentaire TABLE_SCHEMA est absent", () => {
      const response = `
      Voici le tableau :
      <!-- FORMULA_EN: =SUM(A1:A2) -->
      `;
      expect(isComplexResponse(response)).toBe(false);
    });
  });

  describe("extractTableSchema", () => {
    it("devrait extraire le JSON valide", () => {
      const response = `
      <!-- TABLE_SCHEMA: {"type": "complex_table", "title": "Test Table"} -->
      `;
      expect(extractTableSchema(response)).toEqual({
        type: "complex_table",
        title: "Test Table",
      });
    });

    it("devrait retourner null si le JSON est mal formé", () => {
      const response = `
      <!-- TABLE_SCHEMA: {"type": "complex_table", "title": "Test Table" -->
      `;
      expect(extractTableSchema(response)).toBeNull();
    });

    it("devrait retourner null si le commentaire est absent", () => {
      expect(extractTableSchema("Pas de schéma")).toBeNull();
    });
  });

  describe("validateSchema", () => {
    it("devrait valider un schéma minimal correct", () => {
      const raw = {
        type: "complex_table",
        title: "Simulation",
        columns: [
          { header: "Mois", type: "integer", formula: null, formula_en: null },
          { header: "Calcul", type: "number", formula: "=1+1", formula_en: "=1+1" },
        ],
        parameters: [
          { name: "Montant", ref: "C3", value: 1000, type: "number" }
        ],
        data_start_row: 10,
        sample_rows: 5,
      };
      expect(validateSchema(raw)).toEqual(raw);
    });

    it("devrait lever une erreur si formula ou formula_en n'ont pas la même présence de null", () => {
      const raw = {
        type: "complex_table",
        title: "Simulation",
        columns: [
          { header: "Mois", type: "integer", formula: "=1+1", formula_en: null },
        ],
      };
      expect(() => validateSchema(raw)).toThrow(SchemaValidationError);
    });

    it("devrait lever une erreur si la formule ne commence pas par =", () => {
      const raw = {
        type: "complex_table",
        title: "Simulation",
        columns: [
          { header: "Mois", type: "integer", formula: "1+1", formula_en: "1+1" },
        ],
      };
      expect(() => validateSchema(raw)).toThrow(SchemaValidationError);
    });

    it("devrait lever une erreur si des colonnes sont d'input (formula: null) mais parameters est vide", () => {
      const raw = {
        type: "complex_table",
        title: "Simulation",
        columns: [
          { header: "Mois", type: "integer", formula: null, formula_en: null },
          { header: "Calcul", type: "number", formula: "=1+1", formula_en: "=1+1" },
        ],
        parameters: [],
      };
      expect(() => validateSchema(raw)).toThrow(SchemaValidationError);
    });

    it("devrait lever une erreur si ref n'est pas valide", () => {
      const raw = {
        type: "complex_table",
        title: "Simulation",
        columns: [
          { header: "Calcul", type: "number", formula: "=1+1", formula_en: "=1+1" },
        ],
        parameters: [
          { name: "Montant", ref: "invalid_ref", value: 1000, type: "number" }
        ],
      };
      expect(() => validateSchema(raw)).toThrow(SchemaValidationError);
    });
  });
});

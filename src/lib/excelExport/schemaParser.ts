import { z } from "zod";

export type ColumnType =
  | "currency"
  | "percentage"
  | "integer"
  | "number"
  | "date"
  | "text";

export interface TableParameter {
  name: string;          // "Montant du prêt"
  ref: string;           // "C3" — cellule d'input existante
  value?: number | string;   // Optionnel si formula est fourni (cellule calculée)
  type: ColumnType;
  unit?: string;         // "€", "%", "ans"
  formula?: string | null;       // Formule EN ENGLAIS INVARIANT pour cellule paramètre calculée (ex: "=DATE(YEAR(C10),12,31)-DATE(YEAR(C10),1,1)+1")
  formula_label?: string | null; // Formule adaptée au format utilisateur (affichage guide, optionnel)
}

export interface TableColumn {
  header: string;
  type: ColumnType;
  formula: string | null;       // null = colonne input (statique)
  formula_en: string | null;     // version anglaise invariante
  description?: string;
}

export interface TableSchema {
  type: "complex_table";
  title: string;
  parameters: TableParameter[];        // OBLIGATOIRE si ≥1 colonne formula === null
  columns: TableColumn[];               // ≥ 2 colonnes
  data_start_row: number;               // défaut 10, bornes [2, 100]
  sample_rows: number;                  // défaut 3, bornes [1, 100]
}

export class SchemaValidationError extends Error {
  constructor(
    public readonly issues: z.ZodIssue[],
    public readonly raw: unknown
  ) {
    super(`Schéma table invalide: ${issues.map((i) => i.message).join("; ")}`);
    this.name = "SchemaValidationError";
  }
}

const CELL_REF_REGEX = /^[A-Z]{1,3}[0-9]{1,3}$/; // "A1", "AB12", "C3"

const ColumnTypeSchema = z.enum([
  "currency",
  "percentage",
  "integer",
  "number",
  "date",
  "text",
]);

const TableParameterSchema = z.object({
  name: z.string().min(1).max(80),
  ref: z.string().regex(CELL_REF_REGEX, "ref doit être une référence de cellule type 'C3'"),
  value: z.union([z.number(), z.string().min(1)]).optional(),
  type: ColumnTypeSchema,
  unit: z.string().max(10).optional(),
  formula: z.union([z.string().min(1).max(500), z.null()]).optional(),
  formula_label: z.union([z.string().min(1).max(500), z.null()]).optional(),
}).superRefine((param, ctx) => {
  // Un paramètre doit avoir soit une valeur (constante) soit une formule (cellule calculée)
  if (param.value === undefined && (param.formula === undefined || param.formula === null)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `Paramètre "${param.name}": doit avoir soit "value" (constante) soit "formula" (cellule calculée)`,
      path: ["value"],
    });
  }
  // Une formule doit commencer par "="
  if (param.formula !== undefined && param.formula !== null && !param.formula.startsWith("=")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `Paramètre "${param.name}": formula doit commencer par "="`,
      path: ["formula"],
    });
  }
});

const TableColumnSchema = z.object({
  header: z.string().min(1).max(60),
  type: ColumnTypeSchema,
  formula: z.union([z.string().min(1).max(500), z.null()]),
  formula_en: z.union([z.string().min(1).max(500), z.null()]),
  description: z.string().max(200).optional(),
}).superRefine((col, ctx) => {
  // Cohérence : si formula existe, formula_en doit exister aussi (et inversement)
  if ((col.formula === null) !== (col.formula_en === null)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `Colonne "${col.header}": formula et formula_en doivent être null ou non-null ensemble`,
      path: ["formula"],
    });
  }
  // Une formule doit commencer par "="
  if (col.formula !== null && !col.formula.startsWith("=")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `Colonne "${col.header}": formula doit commencer par "="`,
      path: ["formula"],
    });
  }
});

const TableSchemaZod = z.object({
  type: z.literal("complex_table"),
  title: z.string().min(1).max(120),
  parameters: z.array(TableParameterSchema).max(20).default([]),
  columns: z.array(TableColumnSchema).min(2).max(50),
  data_start_row: z.number().int().min(2).max(100).default(10),
  sample_rows: z.number().int().min(1).max(100).default(3),
}).superRefine((schema, ctx) => {
  // Règle F4: si au moins 1 colonne est input (formula === null),
  // parameters doit être non-vide
  const inputColumns = schema.columns.filter((c) => c.formula === null);
  if (inputColumns.length > 0 && (!schema.parameters || schema.parameters.length === 0)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Le schéma contient des colonnes input (formula: null) mais parameters est vide",
      path: ["parameters"],
    });
  }
});

const SCHEMA_REGEX = /<!--\s*TABLE_SCHEMA:\s*(\{[\s\S]*?\})\s*-->/;

export function extractTableSchema(markdown: string): unknown | null {
  const match = markdown.match(SCHEMA_REGEX);
  if (!match) return null;
  try {
    return JSON.parse(match[1]);
  } catch {
    return null; // JSON mal formé
  }
}

export function validateSchema(raw: unknown): TableSchema {
  const parsed = TableSchemaZod.safeParse(raw);
  if (!parsed.success) {
    throw new SchemaValidationError(parsed.error.issues, raw);
  }
  return parsed.data as TableSchema;
}

export function isComplexResponse(markdown: string): boolean {
  return SCHEMA_REGEX.test(markdown);
}

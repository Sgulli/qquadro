import type { CellStyle, ColumnDef, NumberFormat } from "./types.js";

export interface SchemaFieldBase {
  width?: number;
  header?: string;
  style?: CellStyle;
  headerStyle?: CellStyle;
  hidden?: boolean;
}

export interface TextField extends SchemaFieldBase {
  type: "text";
  maxLength?: number;
}

export interface NumberField extends SchemaFieldBase {
  type: "number";
  min?: number;
  max?: number;
  decimals?: number;
  format?: NumberFormat;
}

export interface DateField extends SchemaFieldBase {
  type: "date";
  format?: NumberFormat;
}

export interface BooleanField extends SchemaFieldBase {
  type: "boolean";
}

export interface EnumField extends SchemaFieldBase {
  type: "enum";
  values: readonly string[];
}

export interface CurrencyField extends SchemaFieldBase {
  type: "currency";
  symbol?: string;
  format?: NumberFormat;
}

export interface PercentField extends SchemaFieldBase {
  type: "percent";
  format?: NumberFormat;
}

export type SchemaField =
  | TextField
  | NumberField
  | DateField
  | BooleanField
  | EnumField
  | CurrencyField
  | PercentField;

export type SchemaDefinition = Record<string, SchemaField>;

export type InferRowType<T extends SchemaDefinition> = {
  [K in keyof T]: T[K] extends TextField
    ? string
    : T[K] extends NumberField
      ? number
      : T[K] extends DateField
        ? Date | string
        : T[K] extends BooleanField
          ? boolean
          : T[K] extends EnumField
            ? T[K]["values"][number]
            : T[K] extends CurrencyField
              ? number
              : T[K] extends PercentField
                ? number
                : never;
};

export function text(opts?: Omit<TextField, "type">): TextField {
  return { type: "text", ...opts };
}

export function numberField(opts?: Omit<NumberField, "type">): NumberField {
  return { type: "number", ...opts };
}

export function dateField(opts?: Omit<DateField, "type">): DateField {
  return { type: "date", ...opts };
}

export function booleanField(opts?: Omit<BooleanField, "type">): BooleanField {
  return { type: "boolean", ...opts };
}

export function enumType<T extends readonly string[]>(
  values: T,
  opts?: Omit<EnumField, "type" | "values">,
): EnumField {
  return { type: "enum", values, ...opts };
}

export function currencyField(opts?: Omit<CurrencyField, "type">): CurrencyField {
  return { type: "currency", ...opts };
}

export function percentField(opts?: Omit<PercentField, "type">): PercentField {
  return { type: "percent", ...opts };
}

export const Schema = {
  text,
  number: numberField,
  date: dateField,
  boolean: booleanField,
  enum: enumType,
  currency: currencyField,
  percent: percentField,
} as const;

function fieldToColumnDef(key: string, field: SchemaField): ColumnDef {
  const def: ColumnDef = {
    key,
    header: field.header ?? key,
    width: field.width ?? defaultWidth(field),
    style: field.style ?? defaultStyle(field),
    headerStyle: field.headerStyle,
    hidden: field.hidden,
  };
  if ("format" in field && field.format) {
    def.format = field.format as ColumnDef["format"];
  }
  return def;
}

function defaultWidth(field: SchemaField): number {
  switch (field.type) {
    case "text":
      return 20;
    case "number":
      return 15;
    case "date":
      return 14;
    case "boolean":
      return 10;
    case "enum":
      return Math.max(...field.values.map((v) => v.length)) + 4;
    case "currency":
      return 18;
    case "percent":
      return 12;
  }
}

function defaultStyle(field: SchemaField): CellStyle | undefined {
  switch (field.type) {
    case "currency":
      return { numberFormat: field.format ?? "currency" };
    case "percent":
      return { numberFormat: field.format ?? "percent" };
    case "date":
      return { numberFormat: field.format ?? "date" };
    case "number":
      if (field.decimals !== undefined) {
        return { numberFormat: field.decimals > 0 ? "float" : "integer" };
      }
      return undefined;
    default:
      return undefined;
  }
}

export function schemaToColumnDefs(schema: SchemaDefinition): ColumnDef[] {
  return Object.entries(schema).map(([key, field]) => fieldToColumnDef(key, field));
}

export function validateRow(
  row: Record<string, unknown>,
  schema: SchemaDefinition,
  rowIndex: number,
): string[] {
  const errors: string[] = [];
  for (const [key, field] of Object.entries(schema)) {
    const value = row[key];
    if (value === undefined || value === null) continue;

    switch (field.type) {
      case "text":
        if (typeof value !== "string") {
          errors.push(`Row ${rowIndex}: "${key}" expected string, got ${typeof value}`);
        } else if (field.maxLength && value.length > field.maxLength) {
          errors.push(`Row ${rowIndex}: "${key}" exceeds maxLength ${field.maxLength}`);
        }
        break;
      case "number":
      case "currency":
      case "percent":
        if (typeof value !== "number") {
          errors.push(`Row ${rowIndex}: "${key}" expected number, got ${typeof value}`);
        } else {
          if (field.type === "number") {
            if (field.min !== undefined && value < field.min) {
              errors.push(`Row ${rowIndex}: "${key}" below min ${field.min}`);
            }
            if (field.max !== undefined && value > field.max) {
              errors.push(`Row ${rowIndex}: "${key}" above max ${field.max}`);
            }
          }
        }
        break;
      case "date":
        if (!(value instanceof Date) && typeof value !== "string") {
          errors.push(`Row ${rowIndex}: "${key}" expected Date or string, got ${typeof value}`);
        }
        break;
      case "boolean":
        if (typeof value !== "boolean") {
          errors.push(`Row ${rowIndex}: "${key}" expected boolean, got ${typeof value}`);
        }
        break;
      case "enum":
        if (typeof value !== "string" || !field.values.includes(value)) {
          errors.push(`Row ${rowIndex}: "${key}" must be one of [${field.values.join(", ")}]`);
        }
        break;
    }
  }
  return errors;
}

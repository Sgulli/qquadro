import { type ColumnMap, createColumnMap } from "./column-map.js";
import type { InferRowType, SchemaDefinition } from "./schema.js";
import { schemaToColumnDefs, validateRow } from "./schema.js";
import type { SheetBuilder } from "./sheet-builder.js";
import type { CellStyle, RowData } from "./types.js";
import type { WorkbookBuilder } from "./workbook-builder.js";

export interface DefinedSheet<T extends SchemaDefinition> {
  sheet: SheetBuilder;
  columns: ColumnMap<Record<keyof T & string, { header?: string; width?: number }>>;
  addRows(rows: InferRowType<T>[]): SheetBuilder;
  addRow(row: InferRowType<T>): SheetBuilder;
}

export function defineSheet<T extends SchemaDefinition>(
  workbook: WorkbookBuilder,
  name: string,
  schema: T,
  options?: { writeHeaders?: boolean; headerStyle?: CellStyle },
): DefinedSheet<T> {
  const columnDefs = schemaToColumnDefs(schema);
  const writeHeaders = options?.writeHeaders ?? true;

  let sheetRef: SheetBuilder | undefined;
  workbook.addSheet({ name }, (sheet) => {
    if (writeHeaders) {
      sheet.headers(columnDefs, options?.headerStyle);
    } else {
      sheet.columns(columnDefs);
    }
    sheetRef = sheet;
  });

  if (!sheetRef) {
    throw new Error(`[defineSheet] Failed to create sheet "${name}".`);
  }
  const sheet = sheetRef;

  const columnSchema: Record<string, { header?: string; width?: number }> = {};
  for (const def of columnDefs) {
    columnSchema[def.key] = { header: def.header, width: def.width };
  }

  const columns = createColumnMap(columnSchema, () => ({
    rowCount: sheet.rowCount,
    headerWritten: writeHeaders,
  }));

  return {
    sheet,
    columns: columns as DefinedSheet<T>["columns"],
    addRows(rows: InferRowType<T>[]): SheetBuilder {
      const allErrors: string[] = [];
      for (let i = 0; i < rows.length; i++) {
        const errors = validateRow(rows[i] as Record<string, unknown>, schema, i + 1);
        allErrors.push(...errors);
      }
      if (allErrors.length > 0) {
        throw new Error(`[defineSheet] Validation failed:\n${allErrors.join("\n")}`);
      }
      return sheet.addRows(rows as unknown as RowData[]);
    },
    addRow(row: InferRowType<T>): SheetBuilder {
      const errors = validateRow(row as Record<string, unknown>, schema, 1);
      if (errors.length > 0) {
        throw new Error(`[defineSheet] Validation failed:\n${errors.join("\n")}`);
      }
      return sheet.addRow(row as unknown as RowData);
    },
  };
}

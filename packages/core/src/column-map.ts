import { cellRef, colLetter } from "./coords.js";
import type { CellStyle, ColumnDef } from "./types.js";

export interface ColumnSchema {
  header?: string;
  width?: number;
  style?: CellStyle;
  headerStyle?: CellStyle;
  hidden?: boolean;
  format?: string;
}

export type ColumnSchemaMap = Record<string, ColumnSchema>;

export interface SheetState {
  rowCount: number;
  headerWritten: boolean;
}

export class ColumnRef {
  constructor(
    private readonly _key: string,
    private readonly _index: number,
    private readonly _schema: ColumnSchema,
    private readonly _state: () => SheetState,
  ) {}

  get key(): string {
    return this._key;
  }

  index(): number {
    return this._index;
  }

  letter(): string {
    return colLetter(this._index);
  }

  range(fromRow?: number, toRow?: number): string {
    const state = this._state();
    const start = fromRow ?? (state.headerWritten ? 2 : 1);
    const end = toRow ?? start + state.rowCount - 1;
    const c = this.letter();
    return `${c}${start}:${c}${end}`;
  }

  cell(rowNum: number): string {
    return cellRef(this._index, rowNum);
  }

  toColumnDef(): ColumnDef {
    return {
      key: this._key,
      header: this._schema.header ?? this._key,
      width: this._schema.width,
      style: this._schema.style,
      headerStyle: this._schema.headerStyle,
      hidden: this._schema.hidden,
      format: this._schema.format as ColumnDef["format"],
    };
  }
}

export type ColumnMap<T extends ColumnSchemaMap> = {
  [K in keyof T]: ColumnRef;
};

export function createColumnMap<T extends ColumnSchemaMap>(
  schema: T,
  state: () => SheetState,
): ColumnMap<T> {
  const map = {} as Record<string, ColumnRef>;
  let idx = 1;
  for (const [key, def] of Object.entries(schema)) {
    map[key] = new ColumnRef(key, idx, def, state);
    idx++;
  }
  return map as ColumnMap<T>;
}

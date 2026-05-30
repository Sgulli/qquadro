import type {
  CellIsOperators,
  ConditionalFormattingOptions,
  ContainsTextOperators,
  DataValidation,
  DataValidationWithFormulae,
  Style as ExcelStyle,
  IconSetTypes,
  TimePeriodTypes,
} from "@cj-tech-master/excelts";
import { resolveRange } from "./coords.js";
import type { CellRange, CellStyle, CellValue, RangeValidationDef } from "./types.js";

export interface SheetLike {
  styleRange(range: CellRange, style: CellStyle): unknown;
  addDataValidation(addr: string, validation: DataValidation): unknown;
  addListValidation(
    range: CellRange,
    list: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ): unknown;
  addRangeValidation(range: CellRange, validation: RangeValidationDef): unknown;
  addConditionalFormatting(cf: ConditionalFormattingOptions): unknown;
  addCellIsRule(
    range: CellRange,
    operator: CellIsOperators,
    formulae: (string | number)[],
    style?: Partial<ExcelStyle>,
  ): unknown;
  addExpressionRule(range: CellRange, formula: string, style?: Partial<ExcelStyle>): unknown;
  addDataBar(range: CellRange, color?: { argb?: string; theme?: number }): unknown;
  addColorScale(
    range: CellRange,
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): unknown;
  addIconSet(
    range: CellRange,
    iconSet?: IconSetTypes,
    cfvo?: {
      type: "percent" | "num" | "percentile" | "formula";
      value?: number | string;
    }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): unknown;
  addTop10Rule(
    range: CellRange,
    rank: number,
    options?: { percent?: boolean; bottom?: boolean; style?: Partial<ExcelStyle> },
  ): unknown;
  addAboveAverageRule(
    range: CellRange,
    options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> },
  ): unknown;
  addContainsTextRule(
    range: CellRange,
    text: string,
    operator?: ContainsTextOperators,
    style?: Partial<ExcelStyle>,
  ): unknown;
  addTimePeriodRule(
    range: CellRange,
    timePeriod: TimePeriodTypes,
    style?: Partial<ExcelStyle>,
  ): unknown;
  merge(
    range: CellRange,
    options?: { value?: CellValue; style?: CellStyle; height?: number },
  ): unknown;
  fillFormula(refStr: string, formula: string): unknown;
}

export class RangeBuilder {
  private readonly _ref: string;

  constructor(
    private readonly _sheet: SheetLike,
    range: CellRange,
  ) {
    this._ref = resolveRange(range);
  }

  get ref(): string {
    return this._ref;
  }

  style(s: CellStyle): this {
    this._sheet.styleRange(this._ref, s);
    return this;
  }

  validation(v: DataValidation): this {
    this._sheet.addDataValidation(this._ref, v);
    return this;
  }

  listValidation(
    list: (string | number | Date)[],
    options?: Omit<DataValidationWithFormulae, "type" | "formulae" | "operator">,
  ): this {
    this._sheet.addListValidation(this._ref, list, options);
    return this;
  }

  rangeValidation(v: RangeValidationDef): this {
    this._sheet.addRangeValidation(this._ref, v);
    return this;
  }

  conditionalFormatting(cf: Omit<ConditionalFormattingOptions, "ref">): this {
    this._sheet.addConditionalFormatting({ ...cf, ref: this._ref });
    return this;
  }

  cellIs(operator: CellIsOperators, formulae: (string | number)[], s?: Partial<ExcelStyle>): this {
    this._sheet.addCellIsRule(this._ref, operator, formulae, s);
    return this;
  }

  expression(formula: string, s?: Partial<ExcelStyle>): this {
    this._sheet.addExpressionRule(this._ref, formula, s);
    return this;
  }

  dataBar(color?: { argb?: string; theme?: number }): this {
    this._sheet.addDataBar(this._ref, color);
    return this;
  }

  colorScale(
    cfvo: {
      type: "min" | "max" | "num" | "percent" | "percentile" | "formula";
      value?: number | string;
    }[],
    colors?: { argb?: string; theme?: number }[],
  ): this {
    this._sheet.addColorScale(this._ref, cfvo, colors);
    return this;
  }

  iconSet(
    iconSet?: IconSetTypes,
    cfvo?: {
      type: "percent" | "num" | "percentile" | "formula";
      value?: number | string;
    }[],
    options?: { showValue?: boolean; reverse?: boolean },
  ): this {
    this._sheet.addIconSet(this._ref, iconSet, cfvo, options);
    return this;
  }

  top10(
    rank: number,
    options?: { percent?: boolean; bottom?: boolean; style?: Partial<ExcelStyle> },
  ): this {
    this._sheet.addTop10Rule(this._ref, rank, options);
    return this;
  }

  aboveAverage(options?: { aboveAverage?: boolean; style?: Partial<ExcelStyle> }): this {
    this._sheet.addAboveAverageRule(this._ref, options);
    return this;
  }

  containsText(text: string, operator?: ContainsTextOperators, s?: Partial<ExcelStyle>): this {
    this._sheet.addContainsTextRule(this._ref, text, operator, s);
    return this;
  }

  timePeriod(period: TimePeriodTypes, s?: Partial<ExcelStyle>): this {
    this._sheet.addTimePeriodRule(this._ref, period, s);
    return this;
  }

  merge(options?: { value?: CellValue; style?: CellStyle; height?: number }): this {
    this._sheet.merge(this._ref, options);
    return this;
  }

  formula(f: string): this {
    this._sheet.fillFormula(this._ref, f);
    return this;
  }
}

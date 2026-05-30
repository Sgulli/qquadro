import { describe, expect, it } from "vitest";
import {
  abs,
  add,
  and,
  average,
  averageif,
  averageifs,
  ceiling,
  concat,
  count,
  countif,
  countifs,
  cse,
  date,
  day,
  div,
  eomonth,
  F,
  filter,
  floor,
  hlookup,
  ifExpr,
  iferror,
  ifna,
  ifs,
  index,
  indirect,
  int,
  isblank,
  iserror,
  isnumber,
  istext,
  large,
  left,
  len,
  lower,
  match,
  max,
  mid,
  min,
  mod,
  month,
  mul,
  networkdays,
  not,
  now,
  offset,
  or,
  pct,
  power,
  proper,
  range,
  rank,
  rect,
  ref,
  right,
  round,
  rounddown,
  roundup,
  small,
  sort,
  sqrt,
  sub,
  subtotal,
  sum,
  sumif,
  sumifs,
  switchExpr,
  text,
  today,
  trim,
  trunc,
  unique,
  upper,
  vlookup,
  xlookup,
  year,
} from "../formulas.js";

describe("ref", () => {
  it("builds cell ref from letter + row", () => {
    expect(ref("A", 5)).toBe("A5");
  });

  it("builds cell ref from number + row (1-based col)", () => {
    expect(ref(1, 3)).toBe("A3");
    expect(ref(26, 1)).toBe("Z1");
    expect(ref(27, 1)).toBe("AA1");
  });

  it("handles multi-letter column names", () => {
    expect(ref("AB", 10)).toBe("AB10");
  });
});

describe("range", () => {
  it("builds column range", () => {
    expect(range("B", 2, 10)).toBe("B2:B10");
  });

  it("works with numeric column (1-based)", () => {
    expect(range(3, 1, 5)).toBe("C1:C5");
  });

  it("handles single-row range", () => {
    expect(range("A", 5, 5)).toBe("A5:A5");
  });
});

describe("rect", () => {
  it("builds rectangular range", () => {
    expect(rect("A", 1, "D", 10)).toBe("A1:D10");
  });

  it("works with mixed letter/number columns", () => {
    expect(rect(2, 1, "C", 5)).toBe("B1:C5");
  });
});

describe("aggregate formulas", () => {
  it("sum", () => {
    expect(sum("A1:A10")).toEqual({ formula: "SUM(A1:A10)" });
    expect(sum("A1:A10", 100)).toEqual({ formula: "SUM(A1:A10)", result: 100 });
  });

  it("average", () => {
    expect(average("B2:B8")).toEqual({ formula: "AVERAGE(B2:B8)" });
  });

  it("count", () => {
    expect(count("C1:C100")).toEqual({ formula: "COUNT(C1:C100)" });
  });

  it("max", () => {
    expect(max("D1:D12")).toEqual({ formula: "MAX(D1:D12)" });
  });

  it("min", () => {
    expect(min("E2:E20")).toEqual({ formula: "MIN(E2:E20)" });
  });
});

describe("arithmetic formulas", () => {
  it("add", () => {
    expect(add(1, 2)).toEqual({ formula: "1+2" });
    expect(add("A1", "B1", "C1")).toEqual({ formula: "A1+B1+C1" });
    expect(add(10)).toEqual({ formula: "10" });
  });

  it("sub", () => {
    expect(sub(10, 3)).toEqual({ formula: "10-3" });
    expect(sub("B5", "C5")).toEqual({ formula: "B5-C5" });
  });

  it("mul", () => {
    expect(mul(6, 7)).toEqual({ formula: "6*7" });
    expect(mul("A1", "B1")).toEqual({ formula: "A1*B1" });
  });

  it("div", () => {
    expect(div(10, 2)).toEqual({ formula: "10/2" });
    expect(div("A1", "B1")).toEqual({ formula: "A1/B1" });
    expect(div(5, 0).formula).toBe("5/0");
  });

  it("pct", () => {
    expect(pct("B5", "B3")).toEqual({ formula: "(B5-B3)/B3" });
    expect(pct(100, 50)).toEqual({ formula: "(100-50)/50" });
  });
});

describe("ifExpr", () => {
  it("builds simple IF with numbers", () => {
    expect(ifExpr("A1>10", 100, 0)).toEqual({ formula: "IF(A1>10,100,0)" });
  });

  it("quotes string values", () => {
    expect(ifExpr("B2>0", "Profit", "Loss")).toEqual({
      formula: 'IF(B2>0,"Profit","Loss")',
    });
  });

  it("nests formula values", () => {
    expect(ifExpr("C1>0", sum("D1:D10"), 0)).toEqual({
      formula: "IF(C1>0,SUM(D1:D10),0)",
    });
  });

  it("nests nested IF", () => {
    expect(ifExpr("A1>100", "High", ifExpr("A1>50", "Medium", "Low"))).toEqual({
      formula: 'IF(A1>100,"High",IF(A1>50,"Medium","Low"))',
    });
  });
});

describe("F namespace", () => {
  it("exposes all formula helpers", () => {
    expect(F.ref).toBe(ref);
    expect(F.range).toBe(range);
    expect(F.rect).toBe(rect);
    expect(F.sum).toBe(sum);
    expect(F.average).toBe(average);
    expect(F.count).toBe(count);
    expect(F.max).toBe(max);
    expect(F.min).toBe(min);
    expect(F.add).toBe(add);
    expect(F.sub).toBe(sub);
    expect(F.mul).toBe(mul);
    expect(F.div).toBe(div);
    expect(F.pct).toBe(pct);
    expect(F.if).toBe(ifExpr);
  });

  it("F.if is callable", () => {
    expect(F.if("X>Y", "yes", "no")).toEqual({ formula: 'IF(X>Y,"yes","no")' });
  });
});

describe("lookup & reference formulas", () => {
  it("vlookup", () => {
    expect(vlookup("Apple", "A1:C10", 2)).toEqual({
      formula: 'VLOOKUP("Apple",A1:C10,2,FALSE)',
    });
    expect(vlookup("Apple", "A1:C10", 3, true)).toEqual({
      formula: 'VLOOKUP("Apple",A1:C10,3,TRUE)',
    });
  });

  it("hlookup", () => {
    expect(hlookup("Q1", "A1:D5", 2)).toEqual({
      formula: 'HLOOKUP("Q1",A1:D5,2,FALSE)',
    });
  });

  it("index", () => {
    expect(index("A1:D10", 3)).toEqual({ formula: "INDEX(A1:D10,3)" });
    expect(index("A1:D10", 3, 2)).toEqual({ formula: "INDEX(A1:D10,3,2)" });
  });

  it("match", () => {
    expect(match("Apple", "A1:A10")).toEqual({
      formula: 'MATCH("Apple",A1:A10,0)',
    });
    expect(match(100, "B1:B10", 1)).toEqual({ formula: "MATCH(100,B1:B10,1)" });
  });

  it("xlookup", () => {
    expect(xlookup("Apple", "A1:A10", "B1:B10")).toEqual({
      formula: 'XLOOKUP("Apple",A1:A10,B1:B10)',
    });
    expect(xlookup("Apple", "A1:A10", "B1:B10", "Not found")).toEqual({
      formula: 'XLOOKUP("Apple",A1:A10,B1:B10,"Not found")',
    });
  });

  it("offset", () => {
    expect(offset("A1", 2, 3)).toEqual({ formula: "OFFSET(A1,2,3)" });
    expect(offset("A1", 2, 3, 5, 4)).toEqual({ formula: "OFFSET(A1,2,3,5,4)" });
  });

  it("indirect", () => {
    expect(indirect("A1")).toEqual({ formula: 'INDIRECT("A1")' });
    expect(indirect("Sheet1!A1")).toEqual({ formula: 'INDIRECT("Sheet1!A1")' });
  });
});

describe("conditional logic formulas", () => {
  it("and", () => {
    expect(and("A1>0", "B1>0")).toEqual({ formula: "AND(A1>0,B1>0)" });
    expect(and("A1>0", "B1>0", "C1>0")).toEqual({
      formula: "AND(A1>0,B1>0,C1>0)",
    });
  });

  it("or", () => {
    expect(or("A1>0", "B1>0")).toEqual({ formula: "OR(A1>0,B1>0)" });
  });

  it("not", () => {
    expect(not("A1>0")).toEqual({ formula: "NOT(A1>0)" });
  });

  it("switch", () => {
    expect(switchExpr("A1", 1, "One", 2, "Two", "Other")).toEqual({
      formula: 'SWITCH(A1,1,"One",2,"Two","Other")',
    });
  });

  it("ifs", () => {
    expect(ifs("A1>90", '"A"', "A1>80", '"B"', "TRUE", '"C"')).toEqual({
      formula: 'IFS(A1>90,"A",A1>80,"B",TRUE,"C")',
    });
  });

  it("iferror", () => {
    expect(iferror("A1/B1", 0)).toEqual({ formula: "IFERROR(A1/B1,0)" });
    expect(iferror("A1/B1", "Error")).toEqual({
      formula: 'IFERROR(A1/B1,"Error")',
    });
  });

  it("ifna", () => {
    expect(ifna("VLOOKUP(A1,B:C,2,0)", "Not found")).toEqual({
      formula: 'IFNA(VLOOKUP(A1,B:C,2,0),"Not found")',
    });
  });
});

describe("math formulas", () => {
  it("round", () => {
    expect(round(Math.PI, 2)).toEqual({ formula: "ROUND(3.141592653589793,2)" });
    expect(round("A1")).toEqual({ formula: "ROUND(A1,0)" });
  });

  it("roundup", () => {
    expect(roundup(3.14, 1)).toEqual({ formula: "ROUNDUP(3.14,1)" });
  });

  it("rounddown", () => {
    expect(rounddown(3.99, 0)).toEqual({ formula: "ROUNDDOWN(3.99,0)" });
  });

  it("abs", () => {
    expect(abs(-5)).toEqual({ formula: "ABS(-5)" });
    expect(abs("A1")).toEqual({ formula: "ABS(A1)" });
  });

  it("trunc", () => {
    expect(trunc(3.99)).toEqual({ formula: "TRUNC(3.99,0)" });
    expect(trunc("A1", 2)).toEqual({ formula: "TRUNC(A1,2)" });
  });

  it("int", () => {
    expect(int(3.99)).toEqual({ formula: "INT(3.99)" });
  });

  it("mod", () => {
    expect(mod(10, 3)).toEqual({ formula: "MOD(10,3)" });
  });

  it("ceiling", () => {
    expect(ceiling(3.2, 0.5)).toEqual({ formula: "CEILING(3.2,0.5)" });
  });

  it("floor", () => {
    expect(floor(3.8, 0.5)).toEqual({ formula: "FLOOR(3.8,0.5)" });
  });

  it("power", () => {
    expect(power(2, 3)).toEqual({ formula: "POWER(2,3)" });
  });

  it("sqrt", () => {
    expect(sqrt(16)).toEqual({ formula: "SQRT(16)" });
  });
});

describe("text formulas", () => {
  it("concat", () => {
    expect(concat("Hello", " ", "World")).toEqual({
      formula: 'CONCAT("Hello"," ","World")',
    });
    expect(concat("A1", "B1")).toEqual({ formula: "CONCAT(A1,B1)" });
  });

  it("text", () => {
    expect(text(1234.56, "#,##0.00")).toEqual({
      formula: 'TEXT(1234.56,"#,##0.00")',
    });
    expect(text("A1", "yyyy-mm-dd")).toEqual({
      formula: 'TEXT(A1,"yyyy-mm-dd")',
    });
  });

  it("left", () => {
    expect(left("Hello", 3)).toEqual({ formula: 'LEFT("Hello",3)' });
    expect(left("A1")).toEqual({ formula: "LEFT(A1,1)" });
  });

  it("right", () => {
    expect(right("Hello", 2)).toEqual({ formula: 'RIGHT("Hello",2)' });
  });

  it("mid", () => {
    expect(mid("Hello World", 7, 5)).toEqual({
      formula: 'MID("Hello World",7,5)',
    });
  });

  it("len", () => {
    expect(len("Hello")).toEqual({ formula: 'LEN("Hello")' });
    expect(len("A1")).toEqual({ formula: "LEN(A1)" });
  });

  it("trim", () => {
    expect(trim("  Hello  ")).toEqual({ formula: 'TRIM("  Hello  ")' });
  });

  it("upper", () => {
    expect(upper("hello")).toEqual({ formula: 'UPPER("hello")' });
  });

  it("lower", () => {
    expect(lower("HELLO")).toEqual({ formula: 'LOWER("HELLO")' });
  });

  it("proper", () => {
    expect(proper("hello world")).toEqual({ formula: 'PROPER("hello world")' });
  });
});

describe("aggregate formulas (conditional)", () => {
  it("sumif", () => {
    expect(sumif("A1:A10", ">100")).toEqual({
      formula: 'SUMIF(A1:A10,">100")',
    });
    expect(sumif("A1:A10", ">100", "B1:B10")).toEqual({
      formula: 'SUMIF(A1:A10,">100",B1:B10)',
    });
  });

  it("sumifs", () => {
    expect(sumifs("C1:C10", "A1:A10", ">100", "B1:B10", "<50")).toEqual({
      formula: 'SUMIFS(C1:C10,A1:A10,">100",B1:B10,"<50")',
    });
  });

  it("countif", () => {
    expect(countif("A1:A10", ">0")).toEqual({
      formula: 'COUNTIF(A1:A10,">0")',
    });
    expect(countif("A1:A10", "Apple")).toEqual({
      formula: 'COUNTIF(A1:A10,"Apple")',
    });
  });

  it("countifs", () => {
    expect(countifs("A1:A10", ">0", "B1:B10", "<100")).toEqual({
      formula: 'COUNTIFS(A1:A10,">0",B1:B10,"<100")',
    });
  });

  it("averageif", () => {
    expect(averageif("A1:A10", ">50")).toEqual({
      formula: 'AVERAGEIF(A1:A10,">50")',
    });
    expect(averageif("A1:A10", ">50", "B1:B10")).toEqual({
      formula: 'AVERAGEIF(A1:A10,">50",B1:B10)',
    });
  });

  it("averageifs", () => {
    expect(averageifs("C1:C10", "A1:A10", ">50", "B1:B10", "<100")).toEqual({
      formula: 'AVERAGEIFS(C1:C10,A1:A10,">50",B1:B10,"<100")',
    });
  });

  it("subtotal", () => {
    expect(subtotal(9, "A1:A10")).toEqual({ formula: "SUBTOTAL(9,A1:A10)" });
    expect(subtotal(1, "A1:A10", "B1:B10")).toEqual({
      formula: "SUBTOTAL(1,A1:A10,B1:B10)",
    });
  });
});

describe("date formulas", () => {
  it("now", () => {
    expect(now()).toEqual({ formula: "NOW()" });
  });

  it("today", () => {
    expect(today()).toEqual({ formula: "TODAY()" });
  });

  it("date", () => {
    expect(date(2026, 5, 30)).toEqual({ formula: "DATE(2026,5,30)" });
  });

  it("year", () => {
    expect(year("A1")).toEqual({ formula: "YEAR(A1)" });
  });

  it("month", () => {
    expect(month("A1")).toEqual({ formula: "MONTH(A1)" });
  });

  it("day", () => {
    expect(day("A1")).toEqual({ formula: "DAY(A1)" });
  });

  it("eomonth", () => {
    expect(eomonth("A1", 3)).toEqual({ formula: "EOMONTH(A1,3)" });
  });

  it("networkdays", () => {
    expect(networkdays("A1", "B1")).toEqual({ formula: "NETWORKDAYS(A1,B1)" });
    expect(networkdays("A1", "B1", "H1:H10")).toEqual({
      formula: "NETWORKDAYS(A1,B1,H1:H10)",
    });
  });
});

describe("info formulas", () => {
  it("isnumber", () => {
    expect(isnumber("A1")).toEqual({ formula: "ISNUMBER(A1)" });
  });

  it("istext", () => {
    expect(istext("A1")).toEqual({ formula: "ISTEXT(A1)" });
  });

  it("isblank", () => {
    expect(isblank("A1")).toEqual({ formula: "ISBLANK(A1)" });
  });

  it("iserror", () => {
    expect(iserror("A1")).toEqual({ formula: "ISERROR(A1)" });
  });
});

describe("rank formulas", () => {
  it("rank", () => {
    expect(rank(100, "A1:A10")).toEqual({ formula: "RANK(100,A1:A10,0)" });
    expect(rank("B1", "A1:A10", 1)).toEqual({ formula: "RANK(B1,A1:A10,1)" });
  });

  it("large", () => {
    expect(large("A1:A10", 1)).toEqual({ formula: "LARGE(A1:A10,1)" });
    expect(large("A1:A10", 3)).toEqual({ formula: "LARGE(A1:A10,3)" });
  });

  it("small", () => {
    expect(small("A1:A10", 1)).toEqual({ formula: "SMALL(A1:A10,1)" });
  });
});

describe("array formulas", () => {
  it("filter", () => {
    expect(filter("A1:C10", "B1:B10>100")).toEqual({
      formula: "FILTER(A1:C10,B1:B10>100)",
    });
    expect(filter("A1:C10", "B1:B10>100", "No results")).toEqual({
      formula: 'FILTER(A1:C10,B1:B10>100,"No results")',
    });
  });

  it("sort", () => {
    expect(sort("A1:C10")).toEqual({ formula: "SORT(A1:C10)" });
    expect(sort("A1:C10", 2)).toEqual({ formula: "SORT(A1:C10,2)" });
    expect(sort("A1:C10", 2, -1)).toEqual({ formula: "SORT(A1:C10,2,-1)" });
  });

  it("unique", () => {
    expect(unique("A1:A10")).toEqual({ formula: "UNIQUE(A1:A10)" });
    expect(unique("A1:A10", true)).toEqual({ formula: "UNIQUE(A1:A10,TRUE)" });
  });

  it("cse", () => {
    expect(cse("SUM(A1:A10*B1:B10)")).toEqual({
      formula: "{SUM(A1:A10*B1:B10)}",
    });
  });
});

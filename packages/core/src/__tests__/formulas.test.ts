import { describe, expect, it } from "vitest";
import {
  abs,
  add,
  and,
  average,
  averageif,
  ceiling,
  concat,
  count,
  countif,
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
  switchExpr,
  today,
  trim,
  unique,
  upper,
  vlookup,
  xlookup,
  year,
} from "../formulas.js";

const f = (s: string) => ({ formula: s });

describe("ref", () => {
  it('("A", 5) => A5', () => expect(ref("A", 5)).toBe("A5"));
  it("(1, 3) => A3", () => expect(ref(1, 3)).toBe("A3"));
  it("(26, 1) => Z1", () => expect(ref(26, 1)).toBe("Z1"));
  it("(27, 1) => AA1", () => expect(ref(27, 1)).toBe("AA1"));
  it('("AB", 10) => AB10', () => expect(ref("AB", 10)).toBe("AB10"));
});

describe("range", () => {
  it('("B", 2, 10) => B2:B10', () => expect(range("B", 2, 10)).toBe("B2:B10"));
  it("(3, 1, 5) => C1:C5", () => expect(range(3, 1, 5)).toBe("C1:C5"));
  it('("A", 5, 5) => A5:A5', () => expect(range("A", 5, 5)).toBe("A5:A5"));
});

describe("rect", () => {
  it('("A", 1, "D", 10) => A1:D10', () => expect(rect("A", 1, "D", 10)).toBe("A1:D10"));
  it('(2, 1, "C", 5) => B1:C5', () => expect(rect(2, 1, "C", 5)).toBe("B1:C5"));
});

describe("aggregate", () => {
  it("sum", () => expect(sum("A1:A10")).toEqual(f("SUM(A1:A10)")));
  it("average", () => expect(average("B2:B8")).toEqual(f("AVERAGE(B2:B8)")));
  it("count", () => expect(count("C1:C100")).toEqual(f("COUNT(C1:C100)")));
  it("max", () => expect(max("D1:D12")).toEqual(f("MAX(D1:D12)")));
  it("min", () => expect(min("E2:E20")).toEqual(f("MIN(E2:E20)")));
});

describe("arithmetic", () => {
  it("add", () => {
    expect(add(1, 2)).toEqual(f("1+2"));
    expect(add("A1", "B1", "C1")).toEqual(f("A1+B1+C1"));
  });
  it("sub", () => {
    expect(sub(10, 3)).toEqual(f("10-3"));
    expect(sub("B5", "C5")).toEqual(f("B5-C5"));
  });
  it("mul", () => {
    expect(mul(6, 7)).toEqual(f("6*7"));
    expect(mul("A1", "B1")).toEqual(f("A1*B1"));
  });
  it("div", () => {
    expect(div(10, 2)).toEqual(f("10/2"));
    expect(div("A1", "B1")).toEqual(f("A1/B1"));
  });
  it("pct", () => {
    expect(pct("B5", "B3")).toEqual(f("(B5-B3)/B3"));
    expect(pct(100, 50)).toEqual(f("(100-50)/50"));
  });
});

describe("ifExpr", () => {
  it("simple", () => expect(ifExpr("A1>10", 100, 0)).toEqual(f("IF(A1>10,100,0)")));
  it("strings", () =>
    expect(ifExpr("B2>0", "Profit", "Loss")).toEqual(f('IF(B2>0,"Profit","Loss")')));
  it("nested formula", () =>
    expect(ifExpr("C1>0", sum("D1:D10"), 0)).toEqual(f("IF(C1>0,SUM(D1:D10),0)")));
});

describe("lookup & reference", () => {
  it("vlookup", () =>
    expect(vlookup("Apple", "A1:C10", 2)).toEqual(f('VLOOKUP("Apple",A1:C10,2,false)')));
  it("hlookup", () => expect(hlookup("Q1", "A1:D5", 2)).toEqual(f('HLOOKUP("Q1",A1:D5,2,false)')));
  it("index", () => {
    expect(index("A1:D10", 3)).toEqual(f("INDEX(A1:D10,3)"));
    expect(index("A1:D10", 3, 2)).toEqual(f("INDEX(A1:D10,3,2)"));
  });
  it("match", () => expect(match("Apple", "A1:A10")).toEqual(f('MATCH("Apple",A1:A10,0)')));
  it("xlookup", () =>
    expect(xlookup("Apple", "A1:A10", "B1:B10")).toEqual(f('XLOOKUP("Apple",A1:A10,B1:B10)')));
  it("offset", () => expect(offset("A1", 2, 3)).toEqual(f("OFFSET(A1,2,3)")));
  it("indirect", () => expect(indirect("A1")).toEqual(f('INDIRECT("A1")')));
});

describe("conditional logic", () => {
  it("and", () => expect(and("A1>0", "B1>0")).toEqual(f("AND(A1>0,B1>0)")));
  it("or", () => expect(or("A1>0", "B1>0")).toEqual(f("OR(A1>0,B1>0)")));
  it("not", () => expect(not("A1>0")).toEqual(f("NOT(A1>0)")));
  it("switch", () => expect(switchExpr("A1", 1, "One")).toEqual(f('SWITCH(A1,1,"One")')));
  it("iferror", () => expect(iferror("A1/B1", 0)).toEqual(f("IFERROR(A1/B1,0)")));
  it("ifna", () => expect(ifna("A1", "NA")).toEqual(f('IFNA(A1,"NA")')));
});

describe("math", () => {
  it("round", () => expect(round(3.14, 2)).toEqual(f("ROUND(3.14,2)")));
  it("roundup", () => expect(roundup(3.14, 1)).toEqual(f("ROUNDUP(3.14,1)")));
  it("rounddown", () => expect(rounddown(3.99, 0)).toEqual(f("ROUNDDOWN(3.99,0)")));
  it("abs", () => expect(abs(-5)).toEqual(f("ABS(-5)")));
  it("int", () => expect(int(3.99)).toEqual(f("INT(3.99)")));
  it("mod", () => expect(mod(10, 3)).toEqual(f("MOD(10,3)")));
  it("ceiling", () => expect(ceiling(3.2, 0.5)).toEqual(f("CEILING(3.2,0.5)")));
  it("floor", () => expect(floor(3.8, 0.5)).toEqual(f("FLOOR(3.8,0.5)")));
  it("power", () => expect(power(2, 3)).toEqual(f("POWER(2,3)")));
  it("sqrt", () => expect(sqrt(16)).toEqual(f("SQRT(16)")));
});

describe("text", () => {
  it("concat", () =>
    expect(concat("Hello", " ", "World")).toEqual(f('CONCAT("Hello"," ","World")')));
  it("left", () => expect(left("Hello", 3)).toEqual(f('LEFT("Hello",3)')));
  it("right", () => expect(right("Hello", 2)).toEqual(f('RIGHT("Hello",2)')));
  it("mid", () => expect(mid("Hello World", 7, 5)).toEqual(f('MID("Hello World",7,5)')));
  it("len", () => expect(len("Hello")).toEqual(f('LEN("Hello")')));
  it("trim", () => expect(trim("  Hi  ")).toEqual(f('TRIM("  Hi  ")')));
  it("upper", () => expect(upper("hello")).toEqual(f('UPPER("hello")')));
  it("lower", () => expect(lower("HELLO")).toEqual(f('LOWER("HELLO")')));
  it("proper", () => expect(proper("hello world")).toEqual(f('PROPER("hello world")')));
});

describe("conditional aggregate", () => {
  it("sumif", () => expect(sumif("A1:A10", ">100")).toEqual(f('SUMIF(A1:A10,">100")')));
  it("countif", () => expect(countif("A1:A10", ">0")).toEqual(f('COUNTIF(A1:A10,">0")')));
  it("averageif", () => expect(averageif("A1:A10", ">50")).toEqual(f('AVERAGEIF(A1:A10,">50")')));
  it("subtotal", () => expect(subtotal(9, "A1:A10")).toEqual(f("SUBTOTAL(9,A1:A10)")));
});

describe("date", () => {
  it("now", () => expect(now()).toEqual(f("NOW()")));
  it("today", () => expect(today()).toEqual(f("TODAY()")));
  it("date", () => expect(date(2026, 5, 30)).toEqual(f("DATE(2026,5,30)")));
  it("year", () => expect(year("A1")).toEqual(f("YEAR(A1)")));
  it("month", () => expect(month("A1")).toEqual(f("MONTH(A1)")));
  it("day", () => expect(day("A1")).toEqual(f("DAY(A1)")));
  it("eomonth", () => expect(eomonth("A1", 3)).toEqual(f("EOMONTH(A1,3)")));
  it("networkdays", () => expect(networkdays("A1", "B1")).toEqual(f("NETWORKDAYS(A1,B1)")));
});

describe("info", () => {
  it("isnumber", () => expect(isnumber("A1")).toEqual(f("ISNUMBER(A1)")));
  it("istext", () => expect(istext("A1")).toEqual(f("ISTEXT(A1)")));
  it("isblank", () => expect(isblank("A1")).toEqual(f("ISBLANK(A1)")));
  it("iserror", () => expect(iserror("A1")).toEqual(f("ISERROR(A1)")));
});

describe("rank", () => {
  it("rank", () => expect(rank(100, "A1:A10")).toEqual(f("RANK(100,A1:A10,0)")));
  it("large", () => expect(large("A1:A10", 1)).toEqual(f("LARGE(A1:A10,1)")));
  it("small", () => expect(small("A1:A10", 1)).toEqual(f("SMALL(A1:A10,1)")));
});

describe("array", () => {
  it("filter", () =>
    expect(filter("A1:C10", "B1:B10>100")).toEqual(f("FILTER(A1:C10,B1:B10>100)")));
  it("sort", () => expect(sort("A1:C10", 2)).toEqual(f("SORT(A1:C10,2)")));
  it("unique", () => expect(unique("A1:A10")).toEqual(f("UNIQUE(A1:A10)")));
  it("cse", () => expect(cse("SUM(A1:A10*B1:B10)")).toEqual(f("{SUM(A1:A10*B1:B10)}")));
});

describe("F namespace", () => {
  it("exposes helpers", () => {
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
    expect(F.vlookup).toBe(vlookup);
    expect(F.match).toBe(match);
    expect(F.and).toBe(and);
    expect(F.or).toBe(or);
    expect(F.round).toBe(round);
  });
  it("F.if callable", () => expect(F.if("X>Y", "yes", "no")).toEqual(f('IF(X>Y,"yes","no")')));
});

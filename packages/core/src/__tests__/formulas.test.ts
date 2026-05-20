import { describe, expect, it } from "vitest";
import {
  add,
  average,
  count,
  div,
  F,
  ifExpr,
  max,
  min,
  mul,
  pct,
  range,
  rect,
  ref,
  sub,
  sum,
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
    expect(ifExpr("B2>0", "Profit", "Loss")).toEqual({ formula: 'IF(B2>0,"Profit","Loss")' });
  });

  it("nests formula values", () => {
    expect(ifExpr("C1>0", sum("D1:D10"), 0)).toEqual({ formula: "IF(C1>0,SUM(D1:D10),0)" });
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

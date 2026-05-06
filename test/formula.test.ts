import { describe, expect, it } from "vitest";
import { Worksheet } from "@flexsheet/core";
import {
  parseFormula,
  recalcWorksheet,
  setCellLiteralAndRecalc,
  setCellValueAndRecalc,
} from "@flexsheet/formula";

describe("parseFormula", () => {
  it("parses SUM(A1:A5)", () => {
    const ast = parseFormula("=SUM(A1:A5)");
    expect(ast.type).toBe("call");
    if (ast.type === "call") {
      expect(ast.name.toUpperCase()).toBe("SUM");
      expect(ast.args.length).toBe(1);
      expect(ast.args[0]?.type).toBe("range");
    }
  });

  it("parses A1+B1", () => {
    const ast = parseFormula("=A1+B1");
    expect(ast.type).toBe("binary");
    if (ast.type === "binary") {
      expect(ast.op).toBe("+");
      expect(ast.left.type).toBe("ref");
      expect(ast.right.type).toBe("ref");
    }
  });

  it("parses lowercase refs", () => {
    const ast = parseFormula("=a1+b2");
    expect(ast.type).toBe("binary");
  });

  it("parses absolute refs $K$9 and mixed A$1", () => {
    const ast = parseFormula("=$K$9+A$1");
    expect(ast.type).toBe("binary");
    if (ast.type === "binary") {
      expect(ast.left.type).toBe("ref");
      expect(ast.right.type).toBe("ref");
      if (ast.left.type === "ref") {
        expect(ast.left.row).toBe(8);
        expect(ast.left.col).toBe(10);
      }
    }
  });

  it("parses AVERAGE with dollar refs", () => {
    const ast = parseFormula("=AVERAGE($K$9,$K$10)");
    expect(ast.type).toBe("call");
    if (ast.type === "call") {
      expect(ast.args.length).toBe(2);
      expect(ast.args[0]?.type).toBe("ref");
      expect(ast.args[1]?.type).toBe("ref");
    }
  });
});

describe("evaluateAst + recalc", () => {
  it("evaluates SUM(A1:A5) with literals", () => {
    const sheet = new Worksheet("S", 20, 10);
    for (let r = 0; r < 5; r++) {
      sheet.setCellLiteral(r, 0, r + 1);
    }
    sheet.setCellFormula(5, 0, "=SUM(A1:A5)");
    recalcWorksheet(sheet);
    expect(sheet.getCell(5, 0).value).toBe(15);
  });

  it("evaluates =A1+B1", () => {
    const sheet = new Worksheet("S", 10, 10);
    sheet.setCellLiteral(0, 0, 3);
    sheet.setCellLiteral(0, 1, 7);
    sheet.setCellFormula(0, 2, "=A1+B1");
    recalcWorksheet(sheet);
    expect(sheet.getCell(0, 2).value).toBe(10);
  });

  it("evaluates AVERAGE($K$9,$K$10)", () => {
    const sheet = new Worksheet("S", 20, 20);
    sheet.setCellLiteral(8, 10, 4);
    sheet.setCellLiteral(9, 10, 8);
    sheet.setCellFormula(0, 0, "=AVERAGE($K$9,$K$10)");
    recalcWorksheet(sheet);
    expect(sheet.getCell(0, 0).value).toBe(6);
  });

  it("recalculates when dependency literal changes via setCellValueAndRecalc", () => {
    const sheet = new Worksheet("S", 10, 10);
    sheet.setCellLiteral(0, 0, 1);
    sheet.setCellFormula(0, 1, "=A1*10");
    recalcWorksheet(sheet);
    expect(sheet.getCell(0, 1).value).toBe(10);
    setCellValueAndRecalc(sheet, 0, 0, 5);
    expect(sheet.getCell(0, 1).value).toBe(50);
  });

  it("chains formula dependencies in one recalc pass", () => {
    const sheet = new Worksheet("S", 10, 10);
    sheet.setCellLiteral(0, 0, 2);
    sheet.setCellFormula(0, 1, "=A1+1");
    sheet.setCellFormula(0, 2, "=B1*2");
    recalcWorksheet(sheet);
    expect(sheet.getCell(0, 2).value).toBe(6);
  });

  it("detects circular reference", () => {
    const sheet = new Worksheet("S", 5, 5);
    sheet.setCellFormula(0, 0, "=B1");
    sheet.setCellFormula(0, 1, "=A1");
    recalcWorksheet(sheet);
    expect(sheet.getCell(0, 0).value).toBeNull();
    expect(sheet.getCell(0, 1).value).toBeNull();
  });
});

describe("setCellLiteralAndRecalc", () => {
  it("updates dependents after literal-only change", () => {
    const sheet = new Worksheet("S", 10, 10);
    sheet.setCellLiteral(0, 0, 1);
    sheet.setCellFormula(0, 1, "=A1+100");
    recalcWorksheet(sheet);
    expect(sheet.getCell(0, 1).value).toBe(101);
    setCellLiteralAndRecalc(sheet, 0, 0, 5);
    expect(sheet.getCell(0, 1).value).toBe(105);
  });
});

import type { CellScalar } from "@flexsheet/core";

export function cellScalarToEditString(value: CellScalar): string {
  if (value === null || value === "") {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value);
}

export function parseEditString(raw: string): CellScalar {
  const t = raw.trim();
  if (t === "") {
    return null;
  }
  if (t.startsWith("=")) {
    return t;
  }
  const u = t.toUpperCase();
  if (u === "TRUE") {
    return true;
  }
  if (u === "FALSE") {
    return false;
  }
  const n = Number(t);
  if (!Number.isNaN(n) && String(n) === t) {
    return n;
  }
  return t;
}

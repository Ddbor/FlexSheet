/**
 * TSV 纯文本：列用 \\t、行用 \\n；含制表符/换行/双引号的字段用 CSV 风格双引号包裹。
 */

export function escapeTsvField(text: string): string {
  if (/[\t\n\r"]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

export function serializeTsv(rows: readonly (readonly string[])[]): string {
  return rows.map((row) => row.map(escapeTsvField).join("\t")).join("\n");
}

/**
 * 解析剪贴板 TSV；支持引号字段内换行。
 */
export function parseTsv(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (normalized.length === 0) {
    return [];
  }
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let i = 0;
  let inQuotes = false;
  while (i < normalized.length) {
    const ch = normalized[i]!;
    if (inQuotes) {
      if (ch === '"') {
        if (normalized[i + 1] === '"') {
          cell += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      cell += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === "\t") {
      row.push(cell);
      cell = "";
      i++;
      continue;
    }
    if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      i++;
      continue;
    }
    cell += ch;
    i++;
  }
  row.push(cell);
  rows.push(row);
  return rows;
}

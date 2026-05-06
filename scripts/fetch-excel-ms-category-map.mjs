/**
 * 仅从 Microsoft「Excel 函数（按字母顺序）」索引页抓取函数名对应的文档分类标签，
 * 生成 `packages/flexsheet/src/chrome/excel-ms-category-map.generated.ts`。
 *
 *   node scripts/fetch-excel-ms-category-map.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "packages/flexsheet/src/chrome/excel-ms-category-map.generated.ts");

const INDEX_URL =
  "https://support.microsoft.com/zh-cn/office/excel-%E5%87%BD%E6%95%B0-%E6%8C%89%E5%AD%97%E6%AF%8D%E9%A1%BA%E5%BA%8F-b3944572-255d-4efb-bb96-c6d90033e188";

function decodeHtmlEntities(s) {
  return s
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}

/** @returns {{ name: string, msCategory: string }[]} */
function parseIndexRows(html) {
  const rows = [];
  const trChunks = html.split("<tr>");
  for (const chunk of trChunks) {
    const m = chunk.match(
      /<a\s+href="(\/zh-cn\/office\/[^"]+)"[^>]*class="ocpArticleLink"[^>]*>([^<]+)<\/a>/,
    );
    if (m === null) {
      continue;
    }
    const name = m[2].trim().toUpperCase();
    if (!/^[A-Z][A-Z0-9.]*$/.test(name)) {
      continue;
    }
    const blurbM = chunk.match(/<b\s+class="ocpRunInHead"[^>]*>([^<]*)<\/b>/);
    let msCategory = "其他";
    if (blurbM !== null) {
      let head = decodeHtmlEntities(blurbM[1].trim());
      head = head.replace(/[：:]\s*$/, "").trim();
      if (head.length > 0) {
        msCategory = head;
      }
    }
    rows.push({ name, msCategory });
  }
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    if (seen.has(r.name)) {
      continue;
    }
    seen.add(r.name);
    out.push(r);
  }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

function escapeTsString(s) {
  return s.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
}

const html = await (await fetch(INDEX_URL)).text();
const rows = parseIndexRows(html);

const lines = [];
lines.push("/** 由 `scripts/fetch-excel-ms-category-map.mjs` 根据 Microsoft 支持索引生成，请勿手改。 */");
lines.push("");
lines.push("/** 函数名 -> 文档「类型」分组标题（与索引页 ocpRunInHead 一致） */");
lines.push("export const EXCEL_MS_CATEGORY_BY_NAME: Readonly<Record<string, string>> = {");
for (const r of rows) {
  lines.push(`  '${escapeTsString(r.name)}': '${escapeTsString(r.msCategory)}',`);
}
lines.push("};");
lines.push("");

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, lines.join("\n"), "utf8");
console.log("Wrote", OUT, rows.length, "entries.");

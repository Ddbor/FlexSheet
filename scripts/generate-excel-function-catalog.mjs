/**
 * 从 Microsoft 支持「Excel 函数（按字母顺序）」索引页抓取函数名与详情页链接，
 * 再逐个请求详情页解析「说明 / 语法 / 参数」段落，生成
 * `packages/flexsheet/src/chrome/excel-function-catalog.generated.ts`。
 *
 * 若仅需刷新索引中的「类型」分组标签，可改跑 `scripts/fetch-excel-ms-category-map.mjs`（只需请求索引页一次）。
 *
 * 用法（仓库根目录）:
 *   node scripts/generate-excel-function-catalog.mjs
 *
 * 依赖：Node 18+（内置 fetch）。约 500 次请求，请间隔以免触发限流。
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "packages/flexsheet/src/chrome/excel-function-catalog.generated.ts");

const INDEX_URL =
  "https://support.microsoft.com/zh-cn/office/excel-%E5%87%BD%E6%95%B0-%E6%8C%89%E5%AD%97%E6%AF%8D%E9%A1%BA%E5%BA%8F-b3944572-255d-4efb-bb96-c6d90033e188";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function decodeHtmlEntities(s) {
  return s
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}

/** @returns {{ name: string, path: string, indexBlurb: string, msCategory: string }[]} */
function parseIndexHtml(html) {
  const rows = [];
  const trChunks = html.split("<tr>");
  for (const chunk of trChunks) {
    const m = chunk.match(
      /<a\s+href="(\/zh-cn\/office\/[^"]+)"[^>]*class="ocpArticleLink"[^>]*>([^<]+)<\/a>/,
    );
    if (m === null) {
      continue;
    }
    const hrefPath = m[1];
    const name = m[2].trim().toUpperCase();
    if (!/^[A-Z][A-Z0-9.]*$/.test(name)) {
      continue;
    }
    const blurbM = chunk.match(/<b\s+class="ocpRunInHead"[^>]*>([^<]*)<\/b>\s*([^<]*)/);
    const indexBlurb = blurbM !== null ? decodeHtmlEntities(`${blurbM[1].trim()} ${blurbM[2].trim()}`.trim()) : "";
    let msCategory = "其他";
    if (blurbM !== null) {
      let head = decodeHtmlEntities(blurbM[1].trim());
      head = head.replace(/[：:]\s*$/, "").trim();
      if (head.length > 0) {
        msCategory = head;
      }
    }
    rows.push({ name, path: hrefPath, indexBlurb, msCategory });
  }
  /** 同名保留首次（索引中偶发重复） */
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

function stripTags(html) {
  return decodeHtmlEntities(
    html
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim(),
  );
}

/**
 * @param {string} html
 * @returns {{ description: string, syntax: string, parameters: { name: string, text: string }[], variadic: boolean }}
 */
function parseDetailHtml(html, fnName) {
  const descM = html.match(/<h2[^>]*>\s*说明\s*<\/h2>\s*([\s\S]*?)<\/section>/i);
  let description = "";
  if (descM !== null) {
    const p = descM[1].match(/<p>([\s\S]*?)<\/p>/);
    description = p !== null ? stripTags(p[1]) : stripTags(descM[1]);
  }

  const synM = html.match(/<h2[^>]*>\s*语法\s*<\/h2>\s*([\s\S]*?)<\/section>/i);
  let syntax = `${fnName}(…)`;
  const parameters = [];
  if (synM !== null) {
    const sec = synM[1];
    const pLine = sec.match(/<p>\s*([A-Z][A-Z0-9.]*\([^<]*)\)\s*<\/p>/i);
    if (pLine !== null) {
      syntax = stripTags(pLine[1]) + ")";
    }
    const liBlocks = sec.split("<li>");
    for (let i = 1; i < liBlocks.length; i++) {
      const block = liBlocks[i];
      const bm = block.match(/<b[^>]*>([^<]+)<\/b>/);
      if (bm === null) {
        continue;
      }
      const pname = stripTags(bm[1]).replace(/^["']|["']$/g, "");
      const rest = block.slice(block.indexOf("</b>") + 4);
      const pm = rest.match(/<p>([\s\S]*?)<\/p>/);
      const text = pm !== null ? stripTags(pm[1]) : stripTags(rest);
      if (pname.length > 0 && text.length > 0) {
        parameters.push({ name: pname, text });
      }
    }
  }

  let variadic = false;
  if (parameters.length >= 2) {
    const nums = parameters.map((p) => /^number\s*\d+$/i.exec(p.name.trim()) !== null);
    const sameText =
      parameters.length >= 2 &&
      parameters.every((p) => p.text === parameters[0].text || parameters[0].text.includes(p.text.slice(0, 20)));
    if (nums.every(Boolean) && sameText) {
      variadic = true;
    }
  }
  if (/…|\.\.\./.test(syntax) || /,\s*\[/.test(syntax)) {
    variadic = true;
  }

  if (description.length === 0) {
    description = `Excel 函数 ${fnName}。`;
  }
  return { description, syntax, parameters, variadic };
}

function escapeTsString(s) {
  return s
    .replace(/\\/g, "\\\\")
    .replace(/'/g, "\\'")
    .replace(/\r?\n/g, "\\n")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}

async function main() {
  console.log("Fetching index…");
  const indexRes = await fetch(INDEX_URL);
  if (!indexRes.ok) {
    throw new Error(`Index HTTP ${indexRes.status}`);
  }
  const indexHtml = await indexRes.text();
  const indexRows = parseIndexHtml(indexHtml);
  console.log(`Found ${indexRows.length} functions from index.`);

  const catalog = [];
  let i = 0;
  for (const row of indexRows) {
    i += 1;
    const url = `https://support.microsoft.com${row.path}`;
    process.stdout.write(`\r[${i}/${indexRows.length}] ${row.name}…`);
    let detail = {
      description: row.indexBlurb || `Excel 函数 ${row.name}。`,
      syntax: `${row.name}(…)`,
      parameters: [],
      variadic: false,
    };
    try {
      await sleep(120);
      const r = await fetch(url);
      if (r.ok) {
        const h = await r.text();
        detail = parseDetailHtml(h, row.name);
        if (row.indexBlurb && detail.description.length < 8) {
          detail.description = row.indexBlurb;
        }
      }
    } catch {
      /* 保留索引摘要 */
    }
    catalog.push({
      name: row.name,
      supportPath: row.path,
      description: detail.description,
      syntax: detail.syntax,
      parameters: detail.parameters,
      variadic: detail.variadic,
    });
  }
  console.log("\nWriting", OUT);

  const lines = [];
  lines.push("/** 由 `scripts/generate-excel-function-catalog.mjs` 生成，请勿手改。 */");
  lines.push("");
  lines.push("export interface ExcelFunctionCatalogRow {");
  lines.push("  readonly name: string;");
  lines.push("  readonly supportPath: string;");
  lines.push("  readonly description: string;");
  lines.push("  readonly syntax: string;");
  lines.push("  readonly parameters: readonly { readonly name: string; readonly text: string }[];");
  lines.push("  /** 与 Excel 类似，最后一类参数可重复添加（如 number1、number2…） */");
  lines.push("  readonly variadic: boolean;");
  lines.push("}");
  lines.push("");
  lines.push(`export const EXCEL_FUNCTION_CATALOG: readonly ExcelFunctionCatalogRow[] = [`);
  for (const e of catalog) {
    const ps = e.parameters
      .map((p) => `{ name: '${escapeTsString(p.name)}', text: '${escapeTsString(p.text)}' }`)
      .join(", ");
    lines.push(
      `  { name: '${e.name}', supportPath: '${escapeTsString(e.supportPath)}', description: '${escapeTsString(e.description)}', syntax: '${escapeTsString(e.syntax)}', parameters: [${ps}], variadic: ${e.variadic ? "true" : "false"} },`,
    );
  }
  lines.push("];");
  lines.push("");

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, lines.join("\n"), "utf8");
  console.log("Done.", catalog.length, "entries.");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

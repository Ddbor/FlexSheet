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

/** 无浏览器 UA 时，批量请求易返回极短拦截页，解析会得到错误摘要。 */
const FETCH_HEADERS = {
  "User-Agent":
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
  "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function isLikelySupportArticleHtml(html) {
  return html.length > 8000 && /class="ocpArticleContent"/i.test(html);
}

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

/** Microsoft 支持页常见「本文介绍…公式语法和用法」占位首段，不宜作为唯一说明。 */
function isBoilerplateIntro(s) {
  return /^本文介绍\s+Microsoft\s+Excel\s+中\s+\S+\s+函数的公式语法和用法/.test(s.trim());
}

function collectParagraphs(htmlFragment) {
  const out = [];
  const re = /<p\b[^>]*>([\s\S]*?)<\/p>/gi;
  let m;
  while ((m = re.exec(htmlFragment)) !== null) {
    const t = stripTags(m[1]).trim();
    if (t.length > 0) {
      out.push(t);
    }
  }
  return out;
}

/**
 * 新版帮助：「语法」在 expando 标题 div 中，而非紧接在 `<h2>` 后的纯文本。
 * @param {string} html
 * @param {string} label 如「语法」
 * @returns {string | null} ocpExpandoBody 内部 HTML
 */
function extractOcpExpandoBodyForLabel(html, label) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const headRe = new RegExp(
    `<div[^>]*class="ocpExpandoHeadTitleContainer"[^>]*>\\s*${escaped}\\s*<\\/div>`,
    "i",
  );
  const lab = headRe.exec(html);
  if (lab === null) {
    return null;
  }
  const tail = html.slice(lab.index);
  const bodyOpen = tail.match(/<div[^>]*class="ocpExpandoBody"[^>]*>/i);
  if (bodyOpen === null) {
    return null;
  }
  const start = lab.index + bodyOpen.index + bodyOpen[0].length;
  const s = html.slice(start);
  let depth = 1;
  const tagRe = /<div\b[^>]*>|<\/div>/gi;
  let m;
  while ((m = tagRe.exec(s)) !== null) {
    if (m[0].toLowerCase() === "</div>") {
      depth -= 1;
      if (depth === 0) {
        return s.slice(0, m.index);
      }
    } else {
      depth += 1;
    }
  }
  return null;
}

/** @returns {string | null} */
function extractSyntaxSectionHtml(html) {
  const fromExpando = extractOcpExpandoBodyForLabel(html, "语法");
  if (fromExpando !== null) {
    return fromExpando;
  }
  /** 部分函数（如 DATE）将签名放在「技术细节」折叠块内。 */
  const tech = extractOcpExpandoBodyForLabel(html, "技术细节");
  if (tech !== null) {
    return tech;
  }
  /** 旧版排版：加粗「语法」小标题后紧跟签名（如 AND，无 expando）。 */
  const boldLab = html.match(
    /<b[^>]*class="ocpLegacyBold"[^>]*>\s*语法\s*<\/b>([\s\S]*?)(?=<h2\b|<section\b|<b[^>]*class="ocpLegacyBold"[^>]*>\s*示例\s*<\/b>|<\/article>)/i,
  );
  if (boldLab !== null) {
    return boldLab[1];
  }
  const legacy = html.match(/<h2[^>]*>[\s\S]{0,1200}?语法[\s\S]{0,200}?<\/h2>\s*([\s\S]*?)(?=<h2\b|<\/article>)/i);
  if (legacy !== null) {
    return legacy[1];
  }
  const section = html.match(/<h2[^>]*>\s*语法\s*<\/h2>\s*([\s\S]*?)<\/section>/i);
  return section !== null ? section[1] : null;
}

/**
 * 从语法区块 HTML 中提取形如 FUNC(...) 的首条语法行（跳过「使用…函数时」等说明段）。
 * @param {string} sec
 * @param {string} fnName
 */
function extractSyntaxLine(sec, fnName) {
  const fnEsc = fnName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const preferFn = new RegExp(`\\b${fnEsc}\\s*\\(`, "i");
  const anyFn = /\b[A-Za-z][A-Za-z0-9.]*\s*\(/;
  const normalize = (raw0) => {
    if (/^=\s*/.test(raw0)) {
      return null;
    }
    let raw = raw0.replace(/\s+/g, " ").trim();
    raw = raw.replace(/^语法[：:]\s*/i, "").replace(/^Syntax[：:]\s*/i, "").trim();
    const paren = raw.match(new RegExp(`(${fnEsc}\\s*\\([^)]*\\))`, "i"));
    if (paren !== null) {
      return paren[1].replace(/\s+/g, " ").trim();
    }
    if (preferFn.test(raw) || anyFn.test(raw)) {
      return raw;
    }
    return null;
  };
  const pBlocks = [...sec.matchAll(/<p\b[^>]*>([\s\S]*?)<\/p>/gi)];
  for (const m of pBlocks) {
    const t = normalize(stripTags(m[1]).trim());
    if (t !== null) {
      return t;
    }
  }
  const bold = [...sec.matchAll(/<b[^>]*class="ocpLegacyBold"[^>]*>([\s\S]*?)<\/b>/gi)];
  for (const m of bold) {
    const t = normalize(stripTags(m[1]).trim());
    if (t !== null) {
      return t;
    }
  }
  return null;
}

/** 语法未放在「语法」折叠块时的兜底（如仅 legacy 说明段）。 */
function extractSyntaxLineFromArticle(html, fnName) {
  const art = html.match(/<article[^>]*class="ocpArticleContent"[^>]*>([\s\S]*?)<\/article>/i);
  if (art === null) {
    return null;
  }
  const fnEsc = fnName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pBlocks = [...art[1].matchAll(/<p\b[^>]*>([\s\S]*?)<\/p>/gi)];
  for (const m of pBlocks) {
    const raw = stripTags(m[1])
      .replace(/\s+/g, " ")
      .trim();
    if (/^=\s*/.test(raw)) {
      continue;
    }
    const inline = raw.match(new RegExp(`\\b${fnEsc}\\s*\\([^)]*\\)`, "i"));
    if (inline !== null) {
      return inline[0].replace(/\s+/g, " ").trim();
    }
  }
  return null;
}

const SKIP_PARAM_TABLE_NAMES = new Set([
  "参数",
  "说明",
  "参数名称",
  "公式",
  "数据",
  "结果",
  "示例",
]);

/**
 * @param {string} sec 语法区块 HTML
 * @param {string} fnName
 * @returns {{ name: string, text: string }[]}
 */
function parseParametersFromSyntaxSection(sec, fnName) {
  const parameters = [];
  const tbodyM = sec.match(/<tbody\b[^>]*>([\s\S]*?)<\/tbody>/i);
  if (tbodyM !== null) {
    const rowChunks = tbodyM[1].split("<tr");
    for (let r = 1; r < rowChunks.length; r++) {
      const row = rowChunks[r];
      const tds = [...row.matchAll(/<td\b[^>]*>([\s\S]*?)<\/td>/gi)];
      if (tds.length < 2) {
        continue;
      }
      const nameCell = tds[0][1];
      const descCell = tds[1][1];
      const nm =
        nameCell.match(/<b[^>]*class="ocpRunInHead"[^>]*>([\s\S]*?)<\/b>/i) ??
        nameCell.match(/<b[^>]*class="ocpLegacyBold"[^>]*>([\s\S]*?)<\/b>/i);
      if (nm === null) {
        continue;
      }
      let pname = stripTags(nm[1])
        .replace(/\s*（[^）]*）\s*$/u, "")
        .replace(/\s*\([^)]*\)\s*$/u, "")
        .trim();
      pname = pname.replace(/^["']|["']$/g, "");
      if (SKIP_PARAM_TABLE_NAMES.has(pname)) {
        continue;
      }
      const text = collectParagraphs(descCell).join(" ").trim();
      if (pname.length > 0 && text.length > 0) {
        parameters.push({ name: pname, text });
      }
    }
  }
  if (parameters.length > 0) {
    return parameters;
  }
  const liBlocks = sec.split("<li>");
  for (let i = 1; i < liBlocks.length; i++) {
    const block = liBlocks[i];
    if (block.includes("<table")) {
      continue;
    }
    const bm = block.match(/<b[^>]*>([^<]+)<\/b>/);
    if (bm === null) {
      continue;
    }
    let pname = stripTags(bm[1]).replace(/^["']|["']$/g, "").trim();
    if (pname.toUpperCase() === fnName.toUpperCase()) {
      continue;
    }
    const rest = block.slice(block.indexOf("</b>") + 4);
    const pm = rest.match(/<p>([\s\S]*?)<\/p>/);
    const text = pm !== null ? stripTags(pm[1]) : stripTags(rest);
    if (pname.length > 0 && text.length > 0) {
      parameters.push({ name: pname, text });
    }
  }
  return parameters;
}

/**
 * @param {string} html
 * @returns {{ description: string, syntax: string, parameters: { name: string, text: string }[], variadic: boolean }}
 */
function parseDetailHtml(html, fnName) {
  /** 部分函数页无「说明」标题，正文在 ocpIntroduction（与 Microsoft 在线帮助一致）。 */
  let fromSection = "";
  const descM = html.match(/<h2[^>]*>\s*说明\s*<\/h2>\s*([\s\S]*?)<\/section>/i);
  if (descM !== null) {
    const paras = collectParagraphs(descM[1]);
    fromSection = paras.join(" ").trim();
    if (fromSection.length === 0) {
      fromSection = stripTags(descM[1]).trim();
    }
  }
  let fromIntro = "";
  const introSec = html.match(/<section[^>]*class="ocpIntroduction"[^>]*>([\s\S]*?)<\/section>/i);
  if (introSec !== null) {
    fromIntro = collectParagraphs(introSec[1]).join(" ").trim();
  }
  let description = "";
  if (isBoilerplateIntro(fromIntro) && fromSection.length > 0) {
    description = fromSection;
  } else if (fromSection.length > 0 && fromIntro.length > 0) {
    description = fromIntro.length >= fromSection.length ? fromIntro : fromSection;
  } else if (fromSection.length > 0) {
    description = fromSection;
  } else if (fromIntro.length > 0) {
    description = fromIntro;
  } else {
    const noteM = html.match(/<h2[^>]*>\s*备注\s*<\/h2>\s*([\s\S]*?)<\/section>/i);
    if (noteM !== null) {
      const paras = collectParagraphs(noteM[1]);
      description = paras.join(" ").trim();
    }
  }

  let syntax = `${fnName}(…)`;
  const parameters = [];
  const synSec = extractSyntaxSectionHtml(html);
  if (synSec !== null) {
    const synLine = extractSyntaxLine(synSec, fnName);
    if (synLine !== null) {
      syntax = synLine;
    }
    parameters.push(...parseParametersFromSyntaxSection(synSec, fnName));
  }
  if (syntax === `${fnName}(…)`) {
    const fallback = extractSyntaxLineFromArticle(html, fnName);
    if (fallback !== null) {
      syntax = fallback;
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

  const fnEscEnd = fnName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  syntax = syntax.replace(new RegExp(`^\\s*${fnEscEnd}\\s*\\(`, "i"), `${fnName}(`);

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
  const indexRes = await fetch(INDEX_URL, { headers: FETCH_HEADERS });
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
      await sleep(200);
      let h = "";
      for (let attempt = 0; attempt < 3; attempt++) {
        const r = await fetch(url, { headers: FETCH_HEADERS });
        if (r.ok) {
          h = await r.text();
          if (isLikelySupportArticleHtml(h)) {
            break;
          }
        }
        if (attempt < 2) {
          await sleep(attempt === 0 ? 1200 : 2800);
        }
      }
      if (isLikelySupportArticleHtml(h)) {
        detail = parseDetailHtml(h, row.name);
        const genericDesc = `Excel 函数 ${row.name}。`;
        if (
          row.indexBlurb &&
          (detail.description === genericDesc || detail.description.length < 4)
        ) {
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

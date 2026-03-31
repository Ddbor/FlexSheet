/** 仅保留 XML 1.0 允许的字符（避免 Excel 报「无法打开」或需修复）。 */
export function sanitizeXml10Text(text: string): string {
  let out = "";
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (
      cp === 0x9 ||
      cp === 0xa ||
      cp === 0xd ||
      (cp >= 0x20 && cp <= 0xd7ff) ||
      (cp >= 0xe000 && cp <= 0xfffd) ||
      (cp >= 0x10000 && cp <= 0x10ffff)
    ) {
      out += ch;
    }
  }
  return out;
}

export function escapeXml(text: string): string {
  const safe = sanitizeXml10Text(text);
  return safe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

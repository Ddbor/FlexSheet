/** 0-based列 → A, B, …, AA */
export function colIndexToLetters(col: number): string {
  if (col < 0) {
    return "";
  }
  let label = "";
  let i = col;
  while (i >= 0) {
    label = String.fromCharCode((i % 26) + 65) + label;
    i = Math.floor(i / 26) - 1;
  }
  return label;
}

export function formatCellRef(row: number, col: number): string {
  return `${colIndexToLetters(col)}${row + 1}`;
}

export function parseCellRef(ref: string): { row: number; col: number } | null {
  const m = ref.trim().match(/^([A-Za-z]+)(\d+)$/);
  if (m === null) {
    return null;
  }
  const col = lettersToColIndex(m[1]);
  const row = Number(m[2]) - 1;
  if (row < 0 || col < 0) {
    return null;
  }
  return { row, col };
}

function lettersToColIndex(letters: string): number {
  const u = letters.toUpperCase();
  let v = 0;
  for (let k = 0; k < u.length; k++) {
    const c = u.charCodeAt(k);
    if (c < 65 || c > 90) {
      return -1;
    }
    v = v * 26 + (c - 64);
  }
  return v - 1;
}

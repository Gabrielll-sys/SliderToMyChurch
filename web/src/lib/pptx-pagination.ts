import { normalizeLines } from "@/lib/missa";

const POINTS_PER_INCH = 72;

const DEFAULT_LINE_SPACING = 1.16;
const DEFAULT_MIN_FILL_RATIO = 0.62;
const DEFAULT_MIN_LAST_LINES = 2;

type PaginateInput = {
  lines: string[];
  fontSize: number;
  boxWidthIn: number;
  boxHeightIn: number;
  bold?: boolean;
  lineSpacing?: number;
  minFillRatio?: number;
  minLastLines?: number;
  hardMaxLines?: number;
};

export type PaginationResult = {
  pages: string[][];
  maxLinesPerPage: number;
  maxUnitsPerLine: number;
  estimatedLineHeightPt: number;
};

function charUnitWidth(char: string): number {
  if (char === " ") {
    return 0.33;
  }
  if (/[.,;:!'"`|]/.test(char)) {
    return 0.26;
  }
  if (/[-_/\\(){}\[\]]/.test(char)) {
    return 0.38;
  }
  if (/[MW@#%&]/.test(char)) {
    return 0.9;
  }
  if (/[0-9]/.test(char)) {
    return 0.57;
  }
  if (/[A-ZÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ]/.test(char)) {
    return 0.68;
  }
  if (/[a-záàâãäéèêëíìîïóòôõöúùûüçñ]/.test(char)) {
    return 0.55;
  }
  return 0.62;
}

function estimateTextUnits(text: string): number {
  let units = 0;
  for (const char of text) {
    units += charUnitWidth(char);
  }
  return units;
}

function normalizeSpace(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function splitLongWordByUnits(word: string, maxUnitsPerLine: number): string[] {
  if (!word) {
    return [];
  }

  if (estimateTextUnits(word) <= maxUnitsPerLine) {
    return [word];
  }

  const safeMaxUnits = Math.max(1.4, maxUnitsPerLine);
  const parts: string[] = [];
  let current = "";

  for (const char of word) {
    const suffix = current.length > 0 ? "-" : "";
    const tentative = `${current}${char}`;
    const tentativeUnits = estimateTextUnits(`${tentative}${suffix}`);

    if (tentativeUnits > safeMaxUnits && current.length > 0) {
      parts.push(`${current}-`);
      current = char;
    } else {
      current = tentative;
    }
  }

  if (current.length > 0) {
    parts.push(current);
  }

  return parts;
}

function wrapSingleLine(text: string, maxUnitsPerLine: number): string[] {
  const normalized = normalizeSpace(text);
  if (!normalized) {
    return [];
  }

  const words = normalized.split(" ");
  const wrapped: string[] = [];
  let current = "";

  for (const word of words) {
    const pieces = splitLongWordByUnits(word, maxUnitsPerLine);
    for (const piece of pieces) {
      const candidate = current ? `${current} ${piece}` : piece;
      if (estimateTextUnits(candidate) <= maxUnitsPerLine) {
        current = candidate;
        continue;
      }

      if (current) {
        wrapped.push(current);
      }
      current = piece;
    }
  }

  if (current) {
    wrapped.push(current);
  }

  return wrapped;
}

function wrapLines(lines: string[], maxUnitsPerLine: number): string[] {
  const wrapped: string[] = [];
  for (const line of lines) {
    wrapped.push(...wrapSingleLine(line, maxUnitsPerLine));
  }
  return wrapped;
}

function chunkLines(lines: string[], maxLinesPerPage: number): string[][] {
  if (lines.length === 0) {
    return [];
  }

  const pages: string[][] = [];
  for (let i = 0; i < lines.length; i += maxLinesPerPage) {
    pages.push(lines.slice(i, i + maxLinesPerPage));
  }
  return pages;
}

function rebalancePages(
  pages: string[][],
  maxLinesPerPage: number,
  minFillRatio: number,
  minLastLines: number,
): string[][] {
  if (pages.length <= 1) {
    return pages;
  }

  const minLinesPerPage = Math.max(
    1,
    Math.min(maxLinesPerPage - 1, Math.floor(maxLinesPerPage * minFillRatio)),
  );

  for (let i = pages.length - 1; i > 0; i -= 1) {
    while (
      pages[i].length < minLinesPerPage &&
      pages[i - 1].length > minLinesPerPage
    ) {
      const moved = pages[i - 1].pop();
      if (!moved) {
        break;
      }
      pages[i].unshift(moved);
    }
  }

  const last = pages[pages.length - 1];
  const prev = pages[pages.length - 2];
  while (last.length < minLastLines && prev.length > minLastLines) {
    const moved = prev.pop();
    if (!moved) {
      break;
    }
    last.unshift(moved);
  }

  return pages.filter((page) => page.length > 0);
}

export function paginateTextForSlide(input: PaginateInput): PaginationResult {
  const cleanLines = normalizeLines(input.lines);
  if (cleanLines.length === 0) {
    return {
      pages: [],
      maxLinesPerPage: 0,
      maxUnitsPerLine: 0,
      estimatedLineHeightPt: 0,
    };
  }

  const fontSize = Math.max(8, input.fontSize);
  const lineSpacing = input.lineSpacing ?? DEFAULT_LINE_SPACING;
  const minFillRatio = input.minFillRatio ?? DEFAULT_MIN_FILL_RATIO;
  const minLastLines = Math.max(1, input.minLastLines ?? DEFAULT_MIN_LAST_LINES);

  const widthPt = input.boxWidthIn * POINTS_PER_INCH;
  const heightPt = input.boxHeightIn * POINTS_PER_INCH;

  const horizontalPaddingPt = 10;
  const verticalPaddingPt = 10;
  const usableWidthPt = Math.max(40, widthPt - horizontalPaddingPt);
  const usableHeightPt = Math.max(40, heightPt - verticalPaddingPt);

  const glyphWidthFactor = input.bold ? 0.62 : 0.58;
  const maxUnitsPerLine = Math.max(
    8,
    usableWidthPt / Math.max(1, fontSize * glyphWidthFactor),
  );

  const estimatedLineHeightPt = Math.max(12, fontSize * lineSpacing);
  let maxLinesPerPage = Math.max(
    1,
    Math.floor(usableHeightPt / estimatedLineHeightPt),
  );

  if (input.hardMaxLines) {
    maxLinesPerPage = Math.max(1, Math.min(maxLinesPerPage, input.hardMaxLines));
  }

  const wrappedLines = wrapLines(cleanLines, maxUnitsPerLine);
  const initialPages = chunkLines(wrappedLines, maxLinesPerPage);
  const rebalanced = rebalancePages(
    initialPages,
    maxLinesPerPage,
    minFillRatio,
    minLastLines,
  );

  return {
    pages: rebalanced,
    maxLinesPerPage,
    maxUnitsPerLine,
    estimatedLineHeightPt,
  };
}

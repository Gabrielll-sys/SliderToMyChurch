import defaultTextsJson from "@/data/default_texts.json";
import textosFixosJson from "@/data/textos_fixos.json";

export type SectionType = "musica" | "leitura" | "aclamacao" | "palavra";

export const SECTION_TYPES: SectionType[] = [
  "musica",
  "leitura",
  "aclamacao",
  "palavra",
];

export type TextStyle = {
  fontFace: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  uppercase: boolean;
  lineSpacing: number;
  minFillRatio: number;
  minLastLines: number;
  hardMaxLines?: number;
};

export type SectionStyles = {
  title: TextStyle;
  refrain: TextStyle;
  verse: TextStyle;
  word: TextStyle;
  yellowTitle: TextStyle;
  whiteText: TextStyle;
  acclamation: TextStyle;
  antiphon: TextStyle;
};

export interface SectionState {
  id: string;
  name: string;
  canonicalId: string | null;
  type: SectionType;
  title: string;
  refrainLines: string[];
  verseLines: string[];
  wordLines: string[];
  yellowTitleLines: string[];
  whiteTextLines: string[];
  acclamationLines: string[];
  antiphonLines: string[];
  startWithRefrain: boolean;
  autoDetectRefrain: boolean;
  styles: SectionStyles;
}

export type SectionsState = Record<string, SectionState>;

export interface GeneratorOptions {
  includeCredo: boolean;
  includePreces: boolean;
  includeSanto: boolean;
  includeCordeiro: boolean;
  includeSantaLuzia: boolean;
  includeAvisos: boolean;
}

export interface GeneratePayload {
  presentationTitle: string;
  sectionOrder: string[];
  sections: SectionsState;
  options: GeneratorOptions;
}

export const DEFAULT_SECTION_ORDER = [
  "Entrada",
  "Ato Penitencial",
  "Glória",
  "Palavra",
  "1ª Leitura",
  "Salmo",
  "2ª Leitura",
  "Aclamação",
  "Oferendas",
  "Comunhão",
] as const;

const SECTION_TYPE_BY_NAME: Record<string, SectionType> = {
  Entrada: "musica",
  "Ato Penitencial": "musica",
  Glória: "musica",
  Palavra: "palavra",
  "1ª Leitura": "leitura",
  Salmo: "leitura",
  "2ª Leitura": "leitura",
  Aclamação: "aclamacao",
  Oferendas: "musica",
  Comunhão: "musica",
};

const defaultTexts = defaultTextsJson as Record<string, Record<string, unknown>>;
export const TEXTOS_FIXOS = textosFixosJson as {
  credo: string[];
  santa_luzia: string[];
};

export const DEFAULT_GENERATOR_OPTIONS: GeneratorOptions = {
  includeCredo: true,
  includePreces: true,
  includeSanto: true,
  includeCordeiro: true,
  includeSantaLuzia: true,
  includeAvisos: true,
};

function styleBase(fontSize: number, bold = true, uppercase = true): TextStyle {
  return {
    fontFace: "Arial",
    fontSize,
    bold,
    italic: false,
    uppercase,
    lineSpacing: 1.14,
    minFillRatio: 0.62,
    minLastLines: 2,
  };
}

function createDefaultStylesByType(type: SectionType): SectionStyles {
  const title = {
    ...styleBase(90, true, true),
    lineSpacing: 1.08,
    minFillRatio: 0.5,
    minLastLines: 1,
    hardMaxLines: 5,
  };

  const sectionStyles: SectionStyles = {
    title,
    refrain: { ...styleBase(80, true, true), hardMaxLines: 3 },
    verse: { ...styleBase(80, true, true), hardMaxLines: 3 },
    word: { ...styleBase(80, true, true), hardMaxLines: 6 },
    yellowTitle: {
      ...styleBase(90, true, true),
      lineSpacing: 1.12,
      minFillRatio: 0.55,
      minLastLines: 1,
      hardMaxLines: 5,
    },
    whiteText: { ...styleBase(90, true, true), hardMaxLines: 5 },
    acclamation: {
      ...styleBase(70, true, true),
      minFillRatio: 0.6,
    },
    antiphon: {
      ...styleBase(66, true, true),
      minFillRatio: 0.6,
    },
  };

  if (type === "palavra") {
    sectionStyles.word = { ...styleBase(80, true, true), hardMaxLines: 6 };
  }

  return sectionStyles;
}

export function createPythonPresetStyles(type: SectionType): SectionStyles {
  // Single source of truth for desktop-like defaults used by UI and API.
  return cloneStyles(createDefaultStylesByType(type));
}

export function cloneStyles(styles: SectionStyles): SectionStyles {
  return {
    title: { ...styles.title },
    refrain: { ...styles.refrain },
    verse: { ...styles.verse },
    word: { ...styles.word },
    yellowTitle: { ...styles.yellowTitle },
    whiteText: { ...styles.whiteText },
    acclamation: { ...styles.acclamation },
    antiphon: { ...styles.antiphon },
  };
}

export function normalizeLines(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map((item) => (typeof item === "string" ? item.trim() : ""))
    .filter((line) => line.length > 0);
}

export function splitBlockLines(block: string): string[] {
  return block
    .split(/\r?\n/g)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
}

export function splitTextareaBlocks(value: string): string[] {
  return value
    .split(/\r?\n\s*\r?\n/g)
    .map((block) => splitBlockLines(block).join("\n"))
    .filter((block) => block.length > 0);
}

export function splitTextareaBlocksForEditing(value: string): string[] {
  if (!value) {
    return [];
  }
  return value
    .replace(/\r/g, "")
    .split(/\n\s*\n/g)
    .filter((block) => block.length > 0);
}

export function joinTextareaBlocks(blocks: string[]): string {
  return blocks.join("\n\n");
}

export function blocksToLines(blocks: string[]): string[] {
  const lines: string[] = [];
  for (const block of blocks) {
    lines.push(...splitBlockLines(block));
  }
  return lines;
}

export function normalizeMusicBlocks(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }

  const blocks: string[] = [];
  for (const item of value) {
    if (typeof item === "string") {
      blocks.push(...splitTextareaBlocks(item));
      continue;
    }

    if (Array.isArray(item)) {
      const block = item
        .map((line) => (typeof line === "string" ? line.trim() : ""))
        .filter((line) => line.length > 0)
        .join("\n");
      if (block.length > 0) {
        blocks.push(block);
      }
    }
  }

  return blocks;
}

function normalizeTitle(value: unknown, fallback: string): string {
  if (typeof value !== "string") {
    return fallback;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

export function normalizeSectionType(value: unknown, fallback: SectionType): SectionType {
  if (typeof value !== "string") {
    return fallback;
  }
  return SECTION_TYPES.includes(value as SectionType) ? (value as SectionType) : fallback;
}

export function normalizeSectionId(name: string): string {
  const normalized = name
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return normalized || "secao";
}

export function buildUniqueSectionId(name: string, existingIds: string[]): string {
  const base = normalizeSectionId(name);
  let candidate = base;
  let index = 2;
  const used = new Set(existingIds);
  while (used.has(candidate)) {
    candidate = `${base}-${index}`;
    index += 1;
  }
  return candidate;
}

export function splitTextarea(value: string): string[] {
  return value
    .split(/\r?\n/g)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
}

export function splitTextareaForEditing(value: string): string[] {
  if (!value) {
    return [];
  }
  return value.replace(/\r/g, "").split("\n");
}

export function joinTextarea(lines: string[]): string {
  return lines.join("\n");
}

export function createEmptySection(
  id: string,
  name: string,
  type: SectionType,
): SectionState {
  const styles = createPythonPresetStyles(type);
  return {
    id,
    name,
    canonicalId: null,
    type,
    title: name.toUpperCase(),
    refrainLines: [],
    verseLines: [],
    wordLines: [],
    yellowTitleLines: type === "leitura" ? [name.toUpperCase()] : [],
    whiteTextLines: [],
    acclamationLines: [],
    antiphonLines: [],
    startWithRefrain: false,
    autoDetectRefrain: type === "musica",
    styles,
  };
}

export function buildInitialSections(): SectionsState {
  const sections: SectionsState = {};

  for (const name of DEFAULT_SECTION_ORDER) {
    const raw = defaultTexts[name] ?? {};
    const type = SECTION_TYPE_BY_NAME[name];
    const styles = createPythonPresetStyles(type);

    const section: SectionState = {
      id: name,
      name,
      canonicalId: name,
      type,
      title: normalizeTitle(raw.titulo, name.toUpperCase()),
      refrainLines: normalizeMusicBlocks(raw.refrao),
      verseLines: normalizeMusicBlocks(raw.versos),
      wordLines: normalizeLines(raw.texto),
      yellowTitleLines: normalizeLines(raw.titulo_amarelo),
      whiteTextLines: normalizeLines(raw.texto_branco),
      acclamationLines: normalizeLines(raw.aclamacao_texto),
      antiphonLines: normalizeLines(raw.antifona_texto),
      startWithRefrain: false,
      autoDetectRefrain: type === "musica",
      styles,
    };

    if (type === "leitura" && section.yellowTitleLines.length === 0) {
      section.yellowTitleLines = [name.toUpperCase()];
    }

    sections[section.id] = section;
  }

  return sections;
}

export function buildInitialSectionOrder(): string[] {
  return [...DEFAULT_SECTION_ORDER];
}

function normalizeStyleValue(
  incoming: unknown,
  fallback: TextStyle,
  {
    minFont = 8,
    maxFont = 120,
  }: { minFont?: number; maxFont?: number } = {},
): TextStyle {
  if (!incoming || typeof incoming !== "object") {
    return { ...fallback };
  }

  const source = incoming as Record<string, unknown>;
  const numberInRange = (value: unknown, min: number, max: number, def: number): number => {
    if (typeof value !== "number" || Number.isNaN(value)) {
      return def;
    }
    return Math.max(min, Math.min(max, value));
  };

  return {
    fontFace:
      typeof source.fontFace === "string" && source.fontFace.trim().length > 0
        ? source.fontFace.trim()
        : fallback.fontFace,
    fontSize: Math.round(numberInRange(source.fontSize, minFont, maxFont, fallback.fontSize)),
    bold: typeof source.bold === "boolean" ? source.bold : fallback.bold,
    italic: typeof source.italic === "boolean" ? source.italic : fallback.italic,
    uppercase:
      typeof source.uppercase === "boolean" ? source.uppercase : fallback.uppercase,
    lineSpacing: numberInRange(source.lineSpacing, 0.9, 1.8, fallback.lineSpacing),
    minFillRatio: numberInRange(source.minFillRatio, 0.3, 0.95, fallback.minFillRatio),
    minLastLines: Math.round(
      numberInRange(source.minLastLines, 1, 8, fallback.minLastLines),
    ),
    hardMaxLines:
      typeof source.hardMaxLines === "number" && Number.isFinite(source.hardMaxLines)
        ? Math.round(Math.max(1, Math.min(12, source.hardMaxLines)))
        : fallback.hardMaxLines,
  };
}

export function normalizeSectionState(
  incoming: unknown,
  fallback: SectionState,
): SectionState {
  if (!incoming || typeof incoming !== "object") {
    return {
      ...fallback,
      styles: cloneStyles(fallback.styles),
    };
  }

  const raw = incoming as Record<string, unknown>;
  const type = normalizeSectionType(raw.type, fallback.type);
  const defaultsForType = createPythonPresetStyles(type);

  const stylesRaw =
    raw.styles && typeof raw.styles === "object"
      ? (raw.styles as Record<string, unknown>)
      : {};

  return {
    ...fallback,
    name:
      typeof raw.name === "string" && raw.name.trim().length > 0
        ? raw.name.trim()
        : fallback.name,
    canonicalId:
      typeof raw.canonicalId === "string"
        ? raw.canonicalId
        : raw.canonicalId === null
          ? null
          : fallback.canonicalId,
    type,
    title:
      typeof raw.title === "string" && raw.title.trim().length > 0
        ? raw.title.trim()
        : fallback.title,
    refrainLines: normalizeMusicBlocks(raw.refrainLines),
    verseLines: normalizeMusicBlocks(raw.verseLines),
    wordLines: normalizeLines(raw.wordLines),
    yellowTitleLines: normalizeLines(raw.yellowTitleLines),
    whiteTextLines: normalizeLines(raw.whiteTextLines),
    acclamationLines: normalizeLines(raw.acclamationLines),
    antiphonLines: normalizeLines(raw.antiphonLines),
    startWithRefrain:
      typeof raw.startWithRefrain === "boolean"
        ? raw.startWithRefrain
        : fallback.startWithRefrain,
    autoDetectRefrain:
      typeof raw.autoDetectRefrain === "boolean"
        ? raw.autoDetectRefrain
        : fallback.autoDetectRefrain,
    styles: {
      title: normalizeStyleValue(stylesRaw.title, fallback.styles.title ?? defaultsForType.title),
      refrain: normalizeStyleValue(
        stylesRaw.refrain,
        fallback.styles.refrain ?? defaultsForType.refrain,
      ),
      verse: normalizeStyleValue(stylesRaw.verse, fallback.styles.verse ?? defaultsForType.verse),
      word: normalizeStyleValue(stylesRaw.word, fallback.styles.word ?? defaultsForType.word),
      yellowTitle: normalizeStyleValue(
        stylesRaw.yellowTitle,
        fallback.styles.yellowTitle ?? defaultsForType.yellowTitle,
      ),
      whiteText: normalizeStyleValue(
        stylesRaw.whiteText,
        fallback.styles.whiteText ?? defaultsForType.whiteText,
      ),
      acclamation: normalizeStyleValue(
        stylesRaw.acclamation,
        fallback.styles.acclamation ?? defaultsForType.acclamation,
      ),
      antiphon: normalizeStyleValue(
        stylesRaw.antiphon,
        fallback.styles.antiphon ?? defaultsForType.antiphon,
      ),
    },
  };
}

export function applyUppercase(lines: string[], enabled: boolean): string[] {
  if (!enabled) {
    return lines;
  }
  return lines.map((line) => line.toUpperCase());
}

export function normalizeRefrainLineKey(line: string): string {
  return line
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\s]/gu, "")
    .replace(/\s+/g, " ")
    .trim();
}

export function normalizeRefrainBlockKey(block: string): string {
  return splitBlockLines(block)
    .map(normalizeRefrainLineKey)
    .filter((line) => line.length > 0)
    .join("\n");
}

export function detectRefrainFromBlocks(blocks: string[]): string | null {
  const normalizedBlocks = blocks
    .map((block) => splitBlockLines(block).join("\n"))
    .filter((block) => block.length > 0);

  if (normalizedBlocks.length < 2) {
    return null;
  }

  const frequencies = new Map<string, { count: number; firstIndex: number; sample: string }>();
  normalizedBlocks.forEach((block, index) => {
    const key = normalizeRefrainBlockKey(block);
    if (key.length < 3) {
      return;
    }
    const current = frequencies.get(key);
    if (current) {
      current.count += 1;
      return;
    }
    frequencies.set(key, { count: 1, firstIndex: index, sample: block });
  });

  let best: { count: number; firstIndex: number; sample: string } | null = null;
  for (const candidate of frequencies.values()) {
    if (candidate.count < 2) {
      continue;
    }
    if (!best) {
      best = candidate;
      continue;
    }
    if (
      candidate.count > best.count ||
      (candidate.count === best.count && candidate.firstIndex < best.firstIndex)
    ) {
      best = candidate;
    }
  }

  return best?.sample ?? null;
}

export function detectRefrainFromLines(lines: string[]): string[] {
  const normalized = lines
    .map((line) => line.trim())
    .filter((line) => line.length >= 3);
  if (normalized.length < 2) {
    return [];
  }

  const counts = new Map<string, { count: number; sample: string }>();
  for (const line of normalized) {
    const key = normalizeRefrainLineKey(line);
    if (key.length < 3) {
      continue;
    }
    const current = counts.get(key);
    if (current) {
      current.count += 1;
    } else {
      counts.set(key, { count: 1, sample: line });
    }
  }

  let best: { count: number; sample: string } | null = null;
  for (const value of counts.values()) {
    if (!best || value.count > best.count) {
      best = value;
    }
  }

  if (!best || best.count < 2) {
    return [];
  }
  return [best.sample];
}


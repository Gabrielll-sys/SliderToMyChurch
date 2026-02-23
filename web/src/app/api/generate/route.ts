import fs from "node:fs";
import path from "node:path";

import { NextResponse } from "next/server";
import PptxGenJS from "pptxgenjs";

import {
  applyUppercase,
  blocksToLines,
  buildInitialSectionOrder,
  buildInitialSections,
  createEmptySection,
  DEFAULT_GENERATOR_OPTIONS,
  DEFAULT_SECTION_ORDER,
  normalizeSectionState,
  normalizeSectionType,
  splitBlockLines,
  TEXTOS_FIXOS,
  type GeneratePayload,
  type GeneratorOptions,
  type SectionState,
  type SectionsState,
  type TextStyle,
} from "@/lib/missa";
import { paginateTextForSlide } from "@/lib/pptx-pagination";

export const runtime = "nodejs";

const COLOR_BG = "000000";
const COLOR_YELLOW = "FFC000";
const COLOR_WHITE = "FFFFFF";

const BOX = { x: 0.2, y: 0.2, w: 15.6, h: 8.6 };

type TextBlockConfig = {
  color: string;
  fontFace: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  lineSpacing: number;
  minFillRatio: number;
  minLastLines: number;
  hardMaxLines?: number;
};

type ParsedPayload = GeneratePayload & { sectionOrder: string[] };

type FixedToken =
  | "CREDO"
  | "PRECES"
  | "SANTO"
  | "CORDEIRO"
  | "SANTA_LUZIA"
  | "AVISOS";

type GenerationItem =
  | { kind: "section"; sectionId: string }
  | { kind: "fixed"; token: FixedToken };

function createSlide(pptx: PptxGenJS): PptxGenJS.Slide {
  const slide = pptx.addSlide();
  slide.background = { color: COLOR_BG };
  return slide;
}

function toTextBlockConfig(style: TextStyle, color: string): TextBlockConfig {
  return {
    color,
    fontFace: style.fontFace,
    fontSize: style.fontSize,
    bold: style.bold,
    italic: style.italic,
    lineSpacing: style.lineSpacing,
    minFillRatio: style.minFillRatio,
    minLastLines: style.minLastLines,
    hardMaxLines: style.hardMaxLines,
  };
}

function addTextBlocks(
  pptx: PptxGenJS,
  lines: string[],
  config: TextBlockConfig,
): number {
  const pagination = paginateTextForSlide({
    lines,
    fontSize: config.fontSize,
    boxWidthIn: BOX.w,
    boxHeightIn: BOX.h,
    bold: config.bold,
    lineSpacing: config.lineSpacing,
    minFillRatio: config.minFillRatio,
    minLastLines: config.minLastLines,
    hardMaxLines: config.hardMaxLines,
  });

  if (pagination.pages.length === 0) {
    return 0;
  }

  for (const pageLines of pagination.pages) {
    const slide = createSlide(pptx);
    slide.addText(pageLines.join("\n"), {
      ...BOX,
      fontFace: config.fontFace,
      align: "center",
      valign: "middle",
      color: config.color,
      bold: config.bold,
      italic: config.italic,
      fontSize: config.fontSize,
      margin: 0.04,
      fit: "shrink",
    });
  }
  return pagination.pages.length;
}

function addStyledLines(
  pptx: PptxGenJS,
  lines: string[],
  style: TextStyle,
  color: string,
): number {
  return addTextBlocks(
    pptx,
    applyUppercase(lines, style.uppercase),
    toTextBlockConfig(style, color),
  );
}

function normalizeMusicBlocksForGeneration(blocks: string[]): string[] {
  const normalized = blocks
    .map((block) => splitBlockLines(block).join("\n"))
    .filter((block) => block.length > 0);

  // Compatibilidade com payload antigo (versos linha a linha sem blocos).
  if (normalized.length > 1 && normalized.every((block) => !block.includes("\n"))) {
    return [normalized.join("\n")];
  }
  return normalized;
}

function addAcclamationCombinedSlide(
  pptx: PptxGenJS,
  acclamationLines: string[],
  acclamationStyle: TextStyle,
  antiphonLines: string[],
  antiphonStyle: TextStyle,
): number {
  const acclamation = applyUppercase(acclamationLines, acclamationStyle.uppercase);
  const antiphon = applyUppercase(antiphonLines, antiphonStyle.uppercase);

  if (acclamation.length === 0 && antiphon.length === 0) {
    return 0;
  }

  const slide = createSlide(pptx);
  const rows: Array<{ text: string; style: TextStyle; color: string }> = [];
  for (const line of acclamation) {
    const clean = line.trim();
    if (clean.length === 0) {
      continue;
    }
    rows.push({ text: clean, style: acclamationStyle, color: COLOR_YELLOW });
  }
  if (rows.length > 0 && antiphon.some((line) => line.trim().length > 0)) {
    rows.push({
      text: " ",
      style: {
        ...antiphonStyle,
        fontSize: Math.max(8, Math.min(antiphonStyle.fontSize, 18)),
        bold: false,
        italic: false,
      },
      color: COLOR_WHITE,
    });
  }
  for (const line of antiphon) {
    const clean = line.trim();
    if (clean.length === 0) {
      continue;
    }
    rows.push({ text: clean, style: antiphonStyle, color: COLOR_WHITE });
  }

  const runs: PptxGenJS.TextProps[] = rows.map((row, index) => ({
    text: row.text,
    options: {
      color: row.color,
      fontFace: row.style.fontFace,
      fontSize: row.style.fontSize,
      bold: row.style.bold,
      italic: row.style.italic,
      breakLine: index < rows.length - 1,
    },
  }));

  if (runs.length === 0) {
    return 0;
  }

  slide.addText(runs, {
    ...BOX,
    align: "center",
    valign: "middle",
    margin: 0.04,
    fit: "shrink",
  });
  return 1;
}

function addTitle(pptx: PptxGenJS, title: string, style?: TextStyle): number {
  const titleLines = title
    .split(/\r?\n/g)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  const fallback: TextStyle = {
    fontFace: "Arial",
    fontSize: 90,
    bold: true,
    italic: false,
    uppercase: true,
    lineSpacing: 1.08,
    minFillRatio: 0.5,
    minLastLines: 1,
    hardMaxLines: 4,
  };

  const resolved = style ?? fallback;
  const sourceLines = applyUppercase(titleLines, resolved.uppercase);
  return addTextBlocks(pptx, sourceLines, toTextBlockConfig(resolved, COLOR_YELLOW));
}

function toBoolean(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function toOptions(value: unknown): GeneratorOptions {
  const raw = (value ?? {}) as Partial<GeneratorOptions>;

  return {
    includeCredo: toBoolean(
      raw.includeCredo,
      DEFAULT_GENERATOR_OPTIONS.includeCredo,
    ),
    includePreces: toBoolean(
      raw.includePreces,
      DEFAULT_GENERATOR_OPTIONS.includePreces,
    ),
    includeSanto: toBoolean(raw.includeSanto, DEFAULT_GENERATOR_OPTIONS.includeSanto),
    includeCordeiro: toBoolean(
      raw.includeCordeiro,
      DEFAULT_GENERATOR_OPTIONS.includeCordeiro,
    ),
    includeSantaLuzia: toBoolean(
      raw.includeSantaLuzia,
      DEFAULT_GENERATOR_OPTIONS.includeSantaLuzia,
    ),
    includeAvisos: toBoolean(
      raw.includeAvisos,
      DEFAULT_GENERATOR_OPTIONS.includeAvisos,
    ),
  };
}

function toSectionsAndOrder(
  sectionsValue: unknown,
  orderValue: unknown,
): { sections: SectionsState; sectionOrder: string[] } {
  const defaults = buildInitialSections();
  const rawSections =
    sectionsValue && typeof sectionsValue === "object"
      ? (sectionsValue as Record<string, unknown>)
      : {};

  const sections: SectionsState = {};

  for (const [defaultId, defaultSection] of Object.entries(defaults)) {
    sections[defaultId] = normalizeSectionState(rawSections[defaultId], defaultSection);
  }

  for (const [id, incomingValue] of Object.entries(rawSections)) {
    if (sections[id]) {
      continue;
    }
    // Keep compatibility with dynamic sections added from the web UI.
    const incoming =
      incomingValue && typeof incomingValue === "object"
        ? (incomingValue as Record<string, unknown>)
        : {};
    const name =
      typeof incoming.name === "string" && incoming.name.trim().length > 0
        ? incoming.name.trim()
        : id;
    const type = normalizeSectionType(incoming.type, "musica");
    const base = createEmptySection(id, name, type);
    sections[id] = normalizeSectionState(incomingValue, base);
  }

  const rawOrder = Array.isArray(orderValue)
    ? orderValue.filter((item): item is string => typeof item === "string")
    : buildInitialSectionOrder();

  const order: string[] = [];
  const seen = new Set<string>();
  // Preserve requested order first, then append missing canonical/custom sections.
  for (const id of rawOrder) {
    if (!sections[id] || seen.has(id)) {
      continue;
    }
    seen.add(id);
    order.push(id);
  }

  for (const defaultId of DEFAULT_SECTION_ORDER) {
    if (sections[defaultId] && !seen.has(defaultId)) {
      seen.add(defaultId);
      order.push(defaultId);
    }
  }

  for (const id of Object.keys(sections)) {
    if (!seen.has(id)) {
      seen.add(id);
      order.push(id);
    }
  }

  return { sections, sectionOrder: order };
}

function parsePayload(value: unknown): ParsedPayload {
  const raw = (value ?? {}) as Record<string, unknown>;
  const presentationTitle =
    typeof raw.presentationTitle === "string" && raw.presentationTitle.trim().length > 0
      ? raw.presentationTitle.trim()
      : "MISSA";

  const { sections, sectionOrder } = toSectionsAndOrder(
    raw.sections,
    raw.sectionOrder,
  );

  return {
    presentationTitle,
    sectionOrder,
    sections,
    options: toOptions(raw.options),
  };
}

function sanitizeTitleForFilename(title: string): string {
  const collapsed = title.replace(/\r?\n/g, " ").replace(/\s+/g, " ").trim();
  const safe = collapsed
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/[. ]+$/g, "")
    .slice(0, 120)
    .trim();
  return safe.length > 0 ? safe : "MISSA";
}

function toAsciiFilenameFallback(value: string): string {
  const ascii = value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\x20-\x7E]/g, "");
  const safe = ascii
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/[. ]+$/g, "")
    .slice(0, 120)
    .trim();
  return safe.length > 0 ? safe : "Missa";
}

function textHas(value: string, fragment: string): boolean {
  return value.toLowerCase().includes(fragment.toLowerCase());
}

function normalizeSectionLabel(value: string | null | undefined): string {
  if (!value) {
    return "";
  }
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ª/g, "a")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function isSecondReadingSection(section: SectionState): boolean {
  const labels = [
    normalizeSectionLabel(section.canonicalId),
    normalizeSectionLabel(section.name),
    normalizeSectionLabel(section.title),
  ];
  return labels.some((label) => label === "2a leitura" || label === "segunda leitura");
}

function isAcclamationAnchor(section: SectionState): boolean {
  if (section.type === "aclamacao") {
    return true;
  }
  return textHas(section.title, "aclama") && textHas(section.title, "evangelh");
}

function isOferendasAnchor(section: SectionState): boolean {
  if ((section.canonicalId ?? "").toLowerCase().startsWith("oferendas")) {
    return true;
  }
  return textHas(section.title, "oferend");
}

function isCommunionAnchor(section: SectionState): boolean {
  if ((section.canonicalId ?? "").toLowerCase().startsWith("comunh")) {
    return true;
  }
  return textHas(section.title, "comunh");
}

function buildGenerationSequence(payload: ParsedPayload): GenerationItem[] {
  const sequence: GenerationItem[] = [];

  for (const sectionId of payload.sectionOrder) {
    const section = payload.sections[sectionId];
    if (!section) {
      continue;
    }

    sequence.push({ kind: "section", sectionId });

    if (isAcclamationAnchor(section)) {
      if (payload.options.includeCredo) {
        sequence.push({ kind: "fixed", token: "CREDO" });
      }
      if (payload.options.includePreces) {
        sequence.push({ kind: "fixed", token: "PRECES" });
      }
    }

    if (isOferendasAnchor(section)) {
      if (payload.options.includeSanto) {
        sequence.push({ kind: "fixed", token: "SANTO" });
      }
      if (payload.options.includeCordeiro) {
        sequence.push({ kind: "fixed", token: "CORDEIRO" });
      }
    }

    if (payload.options.includeSantaLuzia && isCommunionAnchor(section)) {
      sequence.push({ kind: "fixed", token: "SANTA_LUZIA" });
    }
  }

  if (payload.options.includeSantaLuzia) {
    const hasCommunion = payload.sectionOrder.some((sectionId) => {
      const section = payload.sections[sectionId];
      return Boolean(section && isCommunionAnchor(section));
    });
    const hasSantaLuzia = sequence.some(
      (item) => item.kind === "fixed" && item.token === "SANTA_LUZIA",
    );
    const hasAnchor = sequence.some((item) => {
      if (item.kind === "fixed") {
        return item.token === "SANTO";
      }
      if (item.kind !== "section") {
        return false;
      }

      const section = payload.sections[item.sectionId];
      return Boolean(section && (isOferendasAnchor(section) || isAcclamationAnchor(section)));
    });

    if (!hasCommunion && !hasSantaLuzia && hasAnchor) {
      const findIndexByPriority = (): number => {
        const matchers: Array<(item: GenerationItem) => boolean> = [
          (item) => item.kind === "fixed" && item.token === "CORDEIRO",
          (item) => item.kind === "fixed" && item.token === "SANTO",
          (item) => {
            if (item.kind !== "section") {
              return false;
            }
            const section = payload.sections[item.sectionId];
            return Boolean(section && isOferendasAnchor(section));
          },
          (item) => item.kind === "fixed" && item.token === "PRECES",
          (item) => item.kind === "fixed" && item.token === "CREDO",
          (item) => {
            if (item.kind !== "section") {
              return false;
            }
            const section = payload.sections[item.sectionId];
            return Boolean(section && isAcclamationAnchor(section));
          },
        ];

        for (const matcher of matchers) {
          const index = sequence.findIndex(matcher);
          if (index >= 0) {
            return index;
          }
        }
        return -1;
      };

      const anchorIndex = findIndexByPriority();
      const santaItem: GenerationItem = { kind: "fixed", token: "SANTA_LUZIA" };
      if (anchorIndex >= 0) {
        sequence.splice(anchorIndex + 1, 0, santaItem);
      } else {
        const lastSectionIndex = sequence.reduce(
          (last, item, index) => (item.kind === "section" ? index : last),
          -1,
        );
        if (lastSectionIndex >= 0) {
          sequence.splice(lastSectionIndex + 1, 0, santaItem);
        } else {
          sequence.push(santaItem);
        }
      }
    }
  }

  if (payload.options.includeAvisos) {
    sequence.push({ kind: "fixed", token: "AVISOS" });
  }

  const deduped: GenerationItem[] = [];
  const seen = new Set<string>();
  for (const item of sequence) {
    const key = item.kind === "section" ? `section:${item.sectionId}` : `fixed:${item.token}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    deduped.push(item);
  }
  return deduped;
}

function addAvisosSlide(pptx: PptxGenJS): number {
  const imagePath = path.join(process.cwd(), "public", "AVISOS.png");
  if (!fs.existsSync(imagePath)) {
    return addTitle(pptx, "AVISOS");
  }

  const slide = createSlide(pptx);
  slide.addImage({ path: imagePath, x: 0, y: 0, w: 16, h: 9 });
  return 1;
}

function addFixedTextSection(
  pptx: PptxGenJS,
  title: string,
  lines: string[],
  contentColor: string,
  fontSize: number,
  linesPerSlide: number,
): number {
  const countTitle = addTitle(pptx, title);
  const validLines = lines
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (validLines.length === 0) {
    return countTitle;
  }

  const safeLinesPerSlide = Math.max(1, Math.floor(linesPerSlide));
  let countBody = 0;

  for (let index = 0; index < validLines.length; index += safeLinesPerSlide) {
    const block = validLines.slice(index, index + safeLinesPerSlide).join(" ").trim();
    if (!block) {
      continue;
    }

    const slide = createSlide(pptx);
    slide.addText(block, {
      ...BOX,
      fontFace: "Arial",
      align: "center",
      valign: "middle",
      color: contentColor,
      bold: true,
      italic: false,
      fontSize,
      margin: 0.04,
      fit: "shrink",
    });
    countBody += 1;
  }

  return countTitle + countBody;
}

function addChunkedStyledText(
  pptx: PptxGenJS,
  lines: string[],
  style: TextStyle,
  color: string,
  linesPerSlide: number,
): number {
  const validLines = applyUppercase(lines, style.uppercase)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (validLines.length === 0) {
    return 0;
  }

  const safeLinesPerSlide = Math.max(1, Math.floor(linesPerSlide));
  let slides = 0;
  for (let index = 0; index < validLines.length; index += safeLinesPerSlide) {
    const block = validLines.slice(index, index + safeLinesPerSlide).join(" ").trim();
    if (!block) {
      continue;
    }

    const slide = createSlide(pptx);
    slide.addText(block, {
      ...BOX,
      fontFace: style.fontFace,
      align: "center",
      valign: "middle",
      color,
      bold: style.bold,
      italic: style.italic,
      fontSize: style.fontSize,
      margin: 0.04,
      fit: "shrink",
    });
    slides += 1;
  }

  return slides;
}

function addFixedTokenSlides(pptx: PptxGenJS, token: FixedToken): number {
  switch (token) {
    case "CREDO":
      return addFixedTextSection(
        pptx,
        "ORAÇÃO DO CREDO",
        TEXTOS_FIXOS.credo,
        COLOR_WHITE,
        90,
        3,
      );
    case "PRECES":
      return addTitle(pptx, "PRECES");
    case "SANTO":
      return addTitle(pptx, "SANTO");
    case "CORDEIRO":
      return addTitle(pptx, "CORDEIRO");
    case "SANTA_LUZIA":
      return addFixedTextSection(
        pptx,
        "ORAÇÃO A SANTA LUZIA",
        TEXTOS_FIXOS.santa_luzia,
        COLOR_WHITE,
        90,
        4,
      );
    case "AVISOS":
      return addAvisosSlide(pptx);
    default:
      return 0;
  }
}

function addMusicSectionSlides(pptx: PptxGenJS, section: SectionState): number {
  const refrainLines = blocksToLines(section.refrainLines);
  const verseBlocks = normalizeMusicBlocksForGeneration(section.verseLines);
  let slides = addTitle(pptx, section.title, section.styles.title);

  if (section.startWithRefrain && refrainLines.length > 0) {
    slides += addStyledLines(
      pptx,
      refrainLines,
      section.styles.refrain,
      COLOR_YELLOW,
    );
  }

  if (verseBlocks.length > 0) {
    for (const block of verseBlocks) {
      slides += addStyledLines(
        pptx,
        splitBlockLines(block),
        section.styles.verse,
        COLOR_WHITE,
      );
      if (refrainLines.length > 0) {
        slides += addStyledLines(
          pptx,
          refrainLines,
          section.styles.refrain,
          COLOR_YELLOW,
        );
      }
    }
    return slides;
  }

  if (refrainLines.length > 0 && !section.startWithRefrain) {
    slides += addStyledLines(
      pptx,
      refrainLines,
      section.styles.refrain,
      COLOR_YELLOW,
    );
  }

  return slides;
}

function addReadingSectionSlides(pptx: PptxGenJS, section: SectionState): number {
  const hasWhiteText = section.whiteTextLines.some((line) => line.trim().length > 0);
  if (isSecondReadingSection(section) && !hasWhiteText) {
    return 0;
  }

  const leituraTitle =
    section.yellowTitleLines.length > 0
      ? section.yellowTitleLines
      : [section.title];

  let slides = addStyledLines(
    pptx,
    leituraTitle,
    section.styles.yellowTitle,
    COLOR_YELLOW,
  );
  slides += addStyledLines(
    pptx,
    section.whiteTextLines,
    section.styles.whiteText,
    COLOR_YELLOW,
  );
  return slides;
}

function addAcclamationSectionSlides(pptx: PptxGenJS, section: SectionState): number {
  let slides = addTitle(pptx, section.title, section.styles.title);
  slides += addAcclamationCombinedSlide(
    pptx,
    section.acclamationLines,
    section.styles.acclamation,
    section.antiphonLines,
    section.styles.antiphon,
  );
  return slides;
}

function addWordSectionSlides(pptx: PptxGenJS, section: SectionState): number {
  let slides = addTitle(pptx, section.title, section.styles.title);
  const linesPerSlide = section.styles.word.hardMaxLines ?? 6;
  slides += addChunkedStyledText(
    pptx,
    section.wordLines,
    section.styles.word,
    COLOR_YELLOW,
    linesPerSlide,
  );
  return slides;
}

function addSectionSlides(pptx: PptxGenJS, section: SectionState): number {
  switch (section.type) {
    case "musica":
      return addMusicSectionSlides(pptx, section);
    case "leitura":
      return addReadingSectionSlides(pptx, section);
    case "aclamacao":
      return addAcclamationSectionSlides(pptx, section);
    case "palavra":
      return addWordSectionSlides(pptx, section);
    default:
      return 0;
  }
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const payload = parsePayload(body);

    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "MASS_CUSTOM", width: 16, height: 9 });
    pptx.layout = "MASS_CUSTOM";
    pptx.author = "Slides To My Church";
    pptx.subject = "Celebração litúrgica";
    pptx.title = payload.presentationTitle;

    let generatedSlides = 0;

    generatedSlides += addTextBlocks(
      pptx,
      payload.presentationTitle.split(/\r?\n/g),
      {
        color: COLOR_YELLOW,
        fontFace: "Arial",
        fontSize: 90,
        bold: true,
        italic: false,
        hardMaxLines: 4,
        lineSpacing: 1.08,
        minFillRatio: 0.5,
        minLastLines: 1,
      },
    );

    const generationSequence = buildGenerationSequence(payload);

    for (const item of generationSequence) {
      if (item.kind === "fixed") {
        generatedSlides += addFixedTokenSlides(pptx, item.token);
        continue;
      }

      if (item.kind !== "section") {
        continue;
      }

      const section = payload.sections[item.sectionId];
      if (!section) {
        continue;
      }

      generatedSlides += addSectionSlides(pptx, section);
    }

    if (generatedSlides === 0) {
      generatedSlides += addTitle(pptx, "SEM CONTEÚDO");
    }

    const buffer = (await pptx.write({
      outputType: "nodebuffer",
    })) as Buffer;
    const bytes = new Uint8Array(buffer);
    const fileTitle = sanitizeTitleForFilename(payload.presentationTitle);
    const utf8Filename = `Slides ${fileTitle}.pptx`;
    const asciiFilename = `Slides ${toAsciiFilenameFallback(fileTitle)}.pptx`;

    return new NextResponse(bytes, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${asciiFilename}"; filename*=UTF-8''${encodeURIComponent(
          utf8Filename,
        )}`,
      },
    });
  } catch (error) {
    return NextResponse.json(
      {
        error: `Erro ao gerar apresentação: ${
          error instanceof Error ? error.message : "desconhecido"
        }`,
      },
      { status: 500 },
    );
  }
}


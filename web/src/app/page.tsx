"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import {
  buildInitialSectionOrder,
  buildInitialSections,
  buildUniqueSectionId,
  createEmptySection,
  DEFAULT_GENERATOR_OPTIONS,
  detectRefrainFromBlocks,
  joinTextareaBlocks,
  joinTextarea,
  normalizeSectionState,
  normalizeSectionType,
  normalizeRefrainBlockKey,
  splitTextareaBlocksForEditing,
  splitTextareaForEditing,
  type GeneratorOptions,
  type GeneratePayload,
  type SectionState,
  type SectionStyles,
  type SectionType,
  type SectionsState,
  type TextStyle,
} from "@/lib/missa";

const BOOLEAN_OPTION_KEYS: Array<keyof GeneratorOptions> = [
  "includeCredo",
  "includePreces",
  "includeSanto",
  "includeCordeiro",
  "includeSantaLuzia",
  "includeAvisos",
];

const BOOLEAN_OPTION_LABELS: Record<keyof GeneratorOptions, string> = {
  includeCredo: "Incluir Credo",
  includePreces: "Incluir Preces",
  includeSanto: "Incluir Santo",
  includeCordeiro: "Incluir Cordeiro",
  includeSantaLuzia: "Incluir Santa Luzia",
  includeAvisos: "Incluir Avisos",
};

const SECTION_TYPE_OPTIONS: Array<{ value: SectionType; label: string }> = [
  { value: "musica", label: "Música" },
  { value: "leitura", label: "Leitura" },
  { value: "aclamacao", label: "Aclamação" },
  { value: "palavra", label: "Palavra" },
];

const SECTION_TYPE_LABELS: Record<SectionType, string> = {
  musica: "Música",
  leitura: "Leitura",
  aclamacao: "Aclamação",
  palavra: "Palavra",
};

const SECTION_TYPE_ICONS: Record<SectionType, string> = {
  musica: "🎵",
  leitura: "📖",
  aclamacao: "📢",
  palavra: "💬",
};

const SECTION_TYPE_BADGE_CLASSES: Record<SectionType, string> = {
  musica: "bg-amber-100 text-amber-800 border-amber-300",
  leitura: "bg-sky-100 text-sky-800 border-sky-300",
  aclamacao: "bg-emerald-100 text-emerald-800 border-emerald-300",
  palavra: "bg-violet-100 text-violet-800 border-violet-300",
};

const FONT_OPTIONS = ["Arial", "Calibri", "Montserrat", "Segoe UI", "Times New Roman"];

type LineField =
  | "refrainLines"
  | "verseLines"
  | "wordLines"
  | "yellowTitleLines"
  | "whiteTextLines"
  | "acclamationLines"
  | "antiphonLines";

type StyleKey = keyof SectionStyles;

type EditorSpec = { field: LineField; label: string; styleKey: StyleKey };
type InsertPlacement = "before" | "after";

const EDITOR_SPECS: Record<SectionType, EditorSpec[]> = {
  musica: [
    { field: "refrainLines", label: "Refrão", styleKey: "refrain" },
    { field: "verseLines", label: "Versos", styleKey: "verse" },
  ],
  leitura: [
    { field: "yellowTitleLines", label: "Título amarelo", styleKey: "yellowTitle" },
    { field: "whiteTextLines", label: "Texto branco", styleKey: "whiteText" },
  ],
  aclamacao: [
    { field: "acclamationLines", label: "Aclamação", styleKey: "acclamation" },
    { field: "antiphonLines", label: "Antífona", styleKey: "antiphon" },
  ],
  palavra: [{ field: "wordLines", label: "Palavra", styleKey: "word" }],
};

const LOCAL_STORAGE_KEY = "slide-generator:editor-state:v1";
const DEFAULT_PRESENTATION_TITLE = "DOMINGO DA\nQUARESMA";

type StatusTone = "neutral" | "success" | "error" | "loading" | "warning";

type StatusState = {
  message: string;
  tone: StatusTone;
};

type PersistedEditorState = {
  presentationTitle: string;
  liturgiaDate: string;
  options: GeneratorOptions;
  sections: SectionsState;
  sectionOrder: string[];
};

const STATUS_TONE_CLASSES: Record<StatusTone, string> = {
  neutral: "text-stone-700",
  success: "text-emerald-700",
  error: "text-red-700",
  loading: "text-sky-700",
  warning: "text-amber-700",
};

function isIsoDate(value: unknown): value is string {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function normalizeOptions(value: unknown): GeneratorOptions {
  const raw =
    value && typeof value === "object" ? (value as Partial<GeneratorOptions>) : {};

  return {
    includeCredo:
      typeof raw.includeCredo === "boolean"
        ? raw.includeCredo
        : DEFAULT_GENERATOR_OPTIONS.includeCredo,
    includePreces:
      typeof raw.includePreces === "boolean"
        ? raw.includePreces
        : DEFAULT_GENERATOR_OPTIONS.includePreces,
    includeSanto:
      typeof raw.includeSanto === "boolean"
        ? raw.includeSanto
        : DEFAULT_GENERATOR_OPTIONS.includeSanto,
    includeCordeiro:
      typeof raw.includeCordeiro === "boolean"
        ? raw.includeCordeiro
        : DEFAULT_GENERATOR_OPTIONS.includeCordeiro,
    includeSantaLuzia:
      typeof raw.includeSantaLuzia === "boolean"
        ? raw.includeSantaLuzia
        : DEFAULT_GENERATOR_OPTIONS.includeSantaLuzia,
    includeAvisos:
      typeof raw.includeAvisos === "boolean"
        ? raw.includeAvisos
        : DEFAULT_GENERATOR_OPTIONS.includeAvisos,
  };
}

function normalizeSections(value: unknown): SectionsState {
  const defaults = buildInitialSections();
  const rawSections =
    value && typeof value === "object" ? (value as Record<string, unknown>) : {};

  const sections: SectionsState = {};

  for (const [id, fallback] of Object.entries(defaults)) {
    sections[id] = normalizeSectionState(rawSections[id], fallback);
  }

  for (const [id, incomingValue] of Object.entries(rawSections)) {
    if (sections[id]) {
      continue;
    }

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

  return sections;
}

function normalizeSectionOrderForState(
  orderValue: unknown,
  sections: SectionsState,
): string[] {
  const source = Array.isArray(orderValue)
    ? orderValue.filter((item): item is string => typeof item === "string")
    : [];

  const order: string[] = [];
  const seen = new Set<string>();

  for (const id of source) {
    if (!sections[id] || seen.has(id)) {
      continue;
    }
    seen.add(id);
    order.push(id);
  }

  for (const id of Object.keys(sections)) {
    if (seen.has(id)) {
      continue;
    }
    seen.add(id);
    order.push(id);
  }

  return order;
}

function isMusicBlockField(field: LineField): field is "refrainLines" | "verseLines" {
  return field === "refrainLines" || field === "verseLines";
}

function formatTodayISO(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
    d.getDate(),
  ).padStart(2, "0")}`;
}

function parseFilename(contentDisposition: string | null): string {
  if (!contentDisposition) {
    return "Missa.pptx";
  }

  const utf8Match = contentDisposition.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
  if (utf8Match?.[1]) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch {
      // Fallback to legacy filename parsing below.
    }
  }

  const quoted = contentDisposition.match(/filename="([^"]+)"/i);
  if (quoted?.[1]) {
    return quoted[1];
  }

  const plain = contentDisposition.match(/filename=([^;]+)/i);
  return plain?.[1]?.trim() ?? "Missa.pptx";
}

function moveItem(items: string[], id: string, delta: -1 | 1): string[] {
  const from = items.indexOf(id);
  if (from < 0 || from + delta < 0 || from + delta >= items.length) {
    return items;
  }
  const next = [...items];
  [next[from], next[from + delta]] = [next[from + delta], next[from]];
  return next;
}

function arraysEqual(a: string[], b: string[]): boolean {
  if (a.length !== b.length) {
    return false;
  }
  return a.every((item, index) => item === b[index]);
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

function StyleEditor({
  style,
  onPatch,
  label,
}: {
  style: TextStyle;
  onPatch: (patch: Partial<TextStyle>) => void;
  label?: string;
}) {
  return (
    <details className="group rounded-lg border border-stone-200 bg-stone-50">
      <summary className="flex cursor-pointer items-center gap-2 px-3 py-2 text-xs font-medium text-stone-500 select-none hover:text-stone-700">
        <span>⚙</span>
        <span>{label ?? "Ajustar estilo"}</span>
        <span className="ml-auto text-[10px] opacity-60 group-open:hidden">▸</span>
        <span className="ml-auto text-[10px] opacity-60 hidden group-open:inline">▾</span>
      </summary>
      <div className="space-y-2 border-t border-stone-200 p-2">
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-5">
          <label className="space-y-1 text-sm">
            <span className="font-medium text-stone-700">Fonte</span>
            <select
              value={style.fontFace}
              onChange={(event) => onPatch({ fontFace: event.target.value })}
              className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
              aria-label="Fonte do texto"
            >
              {FONT_OPTIONS.map((font) => (
                <option key={font}>{font}</option>
              ))}
            </select>
          </label>
          <label className="space-y-1 text-sm">
            <span className="font-medium text-stone-700">Tamanho</span>
            <input
              type="number"
              value={style.fontSize}
              min={8}
              max={120}
              onChange={(event) => {
                const next = Number(event.target.value);
                if (Number.isFinite(next)) {
                  onPatch({ fontSize: clamp(Math.round(next), 8, 120) });
                }
              }}
              className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
              aria-label="Tamanho da fonte"
            />
          </label>
          <label className="flex items-center gap-1 text-sm">
            <input
              type="checkbox"
              checked={style.bold}
              onChange={() => onPatch({ bold: !style.bold })}
            />
            Negrito
          </label>
          <label className="flex items-center gap-1 text-sm">
            <input
              type="checkbox"
              checked={style.italic}
              onChange={() => onPatch({ italic: !style.italic })}
            />
            Itálico
          </label>
          <label className="flex items-center gap-1 text-sm">
            <input
              type="checkbox"
              checked={style.uppercase}
              onChange={() => onPatch({ uppercase: !style.uppercase })}
            />
            Maiúsculas
          </label>
        </div>
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
          <label className="space-y-1 text-sm">
            <span className="font-medium text-stone-700">Espaçamento</span>
            <input
              type="number"
              step={0.01}
              min={0.9}
              max={1.8}
              value={style.lineSpacing}
              onChange={(event) => {
                const next = Number(event.target.value);
                if (Number.isFinite(next)) {
                  onPatch({ lineSpacing: clamp(next, 0.9, 1.8) });
                }
              }}
              className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
              aria-label="Espaçamento entre linhas"
            />
          </label>
          <label className="space-y-1 text-sm">
            <span className="font-medium text-stone-700">Preenchimento</span>
            <input
              type="number"
              step={0.01}
              min={0.3}
              max={0.95}
              value={style.minFillRatio}
              onChange={(event) => {
                const next = Number(event.target.value);
                if (Number.isFinite(next)) {
                  onPatch({ minFillRatio: clamp(next, 0.3, 0.95) });
                }
              }}
              className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
              aria-label="Preenchimento mínimo do slide"
            />
          </label>
          <label className="space-y-1 text-sm">
            <span className="font-medium text-stone-700">Mín. última página</span>
            <input
              type="number"
              min={1}
              max={8}
              value={style.minLastLines}
              onChange={(event) => {
                const next = Number(event.target.value);
                if (Number.isFinite(next)) {
                  onPatch({ minLastLines: clamp(Math.round(next), 1, 8) });
                }
              }}
              className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
              aria-label="Linhas mínimas na última página"
            />
          </label>
          <label className="space-y-1 text-sm">
            <span className="font-medium text-stone-700">Máx. por slide</span>
            <input
              type="number"
              min={1}
              max={12}
              value={style.hardMaxLines ?? ""}
              onChange={(event) => {
                const raw = event.target.value.trim();
                if (!raw) {
                  onPatch({ hardMaxLines: undefined });
                  return;
                }
                const next = Number(raw);
                if (Number.isFinite(next)) {
                  onPatch({ hardMaxLines: clamp(Math.round(next), 1, 12) });
                }
              }}
              className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
              placeholder="Auto"
              aria-label="Máximo de linhas por slide"
            />
          </label>
        </div>
      </div>
    </details>
  );
}

function sectionHasContent(section: SectionState): boolean {
  const fields: LineField[] = ["refrainLines", "verseLines", "wordLines", "yellowTitleLines", "whiteTextLines", "acclamationLines", "antiphonLines"];
  return fields.some((f) => {
    const val = section[f] as string[];
    return val && val.some((line) => line.trim().length > 0);
  });
}

type SectionEditorCardProps = {
  id: string;
  isOpen: boolean;
  canMoveUp: boolean;
  canMoveDown: boolean;
  section: SectionState;
  onToggleOpen: (id: string, open: boolean) => void;
  onMove: (id: string, delta: -1 | 1) => void;
  onDetectRefrain: (id: string) => void;
  onDelete: (id: string) => void;
  onUpdateTitle: (id: string, value: string) => void;
  onToggleStartWithRefrain: (id: string) => void;
  onToggleAutoDetectRefrain: (id: string) => void;
  onUpdateLines: (id: string, field: LineField, value: string) => void;
  onScheduleAutoDetectRefrain: (id: string) => void;
  onUpdateStyle: (id: string, key: StyleKey, patch: Partial<TextStyle>) => void;
};

function SectionEditorCard({
  id,
  isOpen,
  canMoveUp,
  canMoveDown,
  section,
  onToggleOpen,
  onMove,
  onDetectRefrain,
  onDelete,
  onUpdateTitle,
  onToggleStartWithRefrain,
  onToggleAutoDetectRefrain,
  onUpdateLines,
  onScheduleAutoDetectRefrain,
  onUpdateStyle,
}: SectionEditorCardProps) {
  const hasContent = sectionHasContent(section);

  return (
    <details
      open={isOpen}
      onToggle={(event) => onToggleOpen(id, event.currentTarget.open)}
      className="rounded-xl border border-stone-200 bg-white transition-shadow hover:shadow-sm"
    >
      <summary className="flex cursor-pointer items-center gap-2 px-4 py-3 text-sm font-semibold select-none">
        <span className={`inline-block h-2 w-2 rounded-full flex-shrink-0 ${
          hasContent ? "bg-emerald-500" : "bg-stone-300"
        }`} title={hasContent ? "Conteúdo preenchido" : "Seção vazia"} />
        <span className="flex-1">
          {section.title || section.name}
        </span>
        <span className={`inline-flex items-center gap-1 rounded-full border px-2 py-0.5 text-[11px] font-medium ${SECTION_TYPE_BADGE_CLASSES[section.type]}`}>
          {SECTION_TYPE_ICONS[section.type]} {SECTION_TYPE_LABELS[section.type]}
        </span>
        <span className="text-stone-400 text-xs ml-1">{isOpen ? "▾" : "▸"}</span>
      </summary>

      <div className="space-y-3 border-t border-stone-200 p-4">
        <div className="flex flex-wrap gap-2">
          <button
            onClick={() => onMove(id, -1)}
            disabled={!canMoveUp}
            title="Mover seção para cima"
            className={`rounded border border-stone-300 px-2.5 py-1 text-sm ${
              canMoveUp ? "hover:bg-stone-100" : "cursor-not-allowed opacity-40"
            }`}
          >
            ↑
          </button>
          <button
            onClick={() => onMove(id, 1)}
            disabled={!canMoveDown}
            title="Mover seção para baixo"
            className={`rounded border border-stone-300 px-2.5 py-1 text-sm ${
              canMoveDown ? "hover:bg-stone-100" : "cursor-not-allowed opacity-40"
            }`}
          >
            ↓
          </button>
          {section.type === "musica" && (
            <button
              onClick={() => onDetectRefrain(id)}
              className="rounded border border-amber-300 bg-amber-50 px-2 py-1 text-sm hover:bg-amber-100 transition-colors"
            >
              🔍 Detectar refrão
            </button>
          )}
          <button
            onClick={() => onDelete(id)}
            title="Excluir seção"
            className="ml-auto rounded border border-red-200 bg-red-50 px-2.5 py-1 text-sm text-red-600 hover:bg-red-100 transition-colors"
          >
            🗑
          </button>
        </div>

        <label className="space-y-1 text-sm">
          <span className="font-medium text-stone-700">Título da seção</span>
          <input
            value={section.title}
            onChange={(event) => onUpdateTitle(id, event.target.value)}
            className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
          />
        </label>

        {section.type === "musica" && (
          <div className="flex flex-wrap gap-3">
            <label className="flex items-center gap-2 text-sm">
              <input
                type="checkbox"
                checked={section.startWithRefrain}
                onChange={() => onToggleStartWithRefrain(id)}
              />
              Iniciar com refrão
            </label>
            <label className="flex items-center gap-2 text-sm">
              <input
                type="checkbox"
                checked={section.autoDetectRefrain}
                onChange={() => onToggleAutoDetectRefrain(id)}
              />
              Detectar e reorganizar refrão automaticamente ao colar
            </label>
          </div>
        )}

        <StyleEditor
          label="Estilo do título"
          style={section.styles.title}
          onPatch={(patch) => onUpdateStyle(id, "title", patch)}
        />

        {EDITOR_SPECS[section.type].map((spec) => {
          const isMusicField = section.type === "musica" && isMusicBlockField(spec.field);
          const value = isMusicField
            ? joinTextareaBlocks(section[spec.field] as string[])
            : joinTextarea(section[spec.field] as string[]);

          return (
            <div key={spec.field} className="space-y-1">
              <p className="text-sm font-semibold">{spec.label}</p>
              <p className="text-xs text-stone-500">
                {isMusicField ? "Separe blocos com uma linha em branco." : "Uma linha por frase."}
              </p>
              <textarea
                value={value}
                onChange={(event) => onUpdateLines(id, spec.field, event.target.value)}
                onPaste={() => {
                  if (isMusicField) {
                    onScheduleAutoDetectRefrain(id);
                  }
                }}
                className="min-h-32 w-full rounded border border-stone-300 p-2 text-sm"
                aria-label={`${spec.label} da seção ${section.title || section.name}`}
              />
              <StyleEditor
                label={`Estilo: ${spec.label}`}
                style={section.styles[spec.styleKey]}
                onPatch={(patch) => onUpdateStyle(id, spec.styleKey, patch)}
              />
            </div>
          );
        })}
      </div>
    </details>
  );
}

export default function Home() {
  const initialLiturgiaDate = useMemo(() => formatTodayISO(), []);
  const [sections, setSections] = useState<SectionsState>(() => buildInitialSections());
  const [sectionOrder, setSectionOrder] = useState<string[]>(() => buildInitialSectionOrder());
  const [openSectionIds, setOpenSectionIds] = useState<string[]>([]);
  const [presentationTitle, setPresentationTitle] = useState(DEFAULT_PRESENTATION_TITLE);
  const [liturgiaDate, setLiturgiaDate] = useState(initialLiturgiaDate);
  const [options, setOptions] = useState<GeneratorOptions>(DEFAULT_GENERATOR_OPTIONS);
  const [status, setStatus] = useState<StatusState>({ message: "Pronto.", tone: "neutral" });
  const [loadingLiturgia, setLoadingLiturgia] = useState(false);
  const [loadingPpt, setLoadingPpt] = useState(false);
  const [isStorageHydrated, setIsStorageHydrated] = useState(false);
  const [skipInitialLiturgiaImport, setSkipInitialLiturgiaImport] = useState(false);
  const [newSectionName, setNewSectionName] = useState("");
  const [newSectionType, setNewSectionType] = useState<SectionType>("musica");
  const [insertReferenceId, setInsertReferenceId] = useState("");
  const [insertPlacement, setInsertPlacement] = useState<InsertPlacement>("after");
  const newSectionNameInputRef = useRef<HTMLInputElement>(null);

  const setStatusMessage = useCallback((message: string, tone: StatusTone = "neutral") => {
    setStatus({ message, tone });
  }, []);

  const updateSection = (id: string, updater: (section: SectionState) => SectionState) => {
    setSections((current) => {
      const section = current[id];
      if (!section) {
        return current;
      }
      return { ...current, [id]: updater(section) };
    });
  };

  const removeSection = (id: string) => {
    setSections((current) => {
      if (!current[id]) {
        return current;
      }
      const next = { ...current };
      delete next[id];
      return next;
    });
    setSectionOrder((current) => current.filter((value) => value !== id));
    setOpenSectionIds((current) => current.filter((value) => value !== id));
  };

  const requestDeleteSection = (id: string) => {
    const section = sections[id];
    if (!section) {
      return;
    }

    const confirmed = window.confirm(
      `Excluir a seção "${section.title || section.name}"?\nEssa ação não pode ser desfeita.`,
    );
    if (!confirmed) {
      return;
    }

    removeSection(id);
    setStatusMessage("Seção excluída.", "warning");
  };

  const updateByCanonical = useCallback((
    canonicalId: string,
    updater: (section: SectionState) => SectionState,
  ) => {
    setSections((current) => {
      const id = Object.keys(current).find((item) => current[item]?.canonicalId === canonicalId);
      if (!id) {
        return current;
      }
      const section = current[id];
      return { ...current, [id]: updater(section) };
    });
  }, []);

  const updateLines = (id: string, field: LineField, value: string) => {
    updateSection(id, (section) => ({
      ...section,
      [field]:
        section.type === "musica" && isMusicBlockField(field)
          ? splitTextareaBlocksForEditing(value)
          : splitTextareaForEditing(value),
    }));
  };

  const updateStyle = (id: string, key: StyleKey, patch: Partial<TextStyle>) => {
    updateSection(id, (section) => ({
      ...section,
      styles: {
        ...section.styles,
        [key]: {
          ...section.styles[key],
          ...patch,
        },
      },
    }));
  };

  const addSection = () => {
    const name = newSectionName.trim();
    if (!name) {
      setStatusMessage("Informe um nome para a nova seção.", "error");
      newSectionNameInputRef.current?.focus();
      return;
    }
    const id = buildUniqueSectionId(name, sectionOrder);
    setSections((current) => ({
      ...current,
      [id]: createEmptySection(id, name, newSectionType),
    }));
    setSectionOrder((current) => {
      if (current.length === 0) {
        return [id];
      }
      const anchorId = insertReferenceId || current[current.length - 1];
      const anchorIndex = current.indexOf(anchorId);
      if (anchorIndex < 0) {
        return [...current, id];
      }
      const insertAt = insertPlacement === "before" ? anchorIndex : anchorIndex + 1;
      const next = [...current];
      next.splice(insertAt, 0, id);
      return next;
    });
    setOpenSectionIds((current) => (current.includes(id) ? current : [...current, id]));
    setInsertReferenceId(id);
    setNewSectionName("");
    setStatusMessage("Seção adicionada.", "success");
  };

  const withDetectedRefrainFromVerses = (section: SectionState): SectionState | null => {
    if (section.type !== "musica") {
      return null;
    }
    const detected = detectRefrainFromBlocks(section.verseLines);
    if (!detected) {
      return null;
    }
    const refrainKey = normalizeRefrainBlockKey(detected);
    const nextVerses = section.verseLines.filter(
      (block) => normalizeRefrainBlockKey(block) !== refrainKey,
    );
    const alreadyInRefrain = section.refrainLines.some(
      (block) => normalizeRefrainBlockKey(block) === refrainKey,
    );
    const nextRefrain = alreadyInRefrain
      ? section.refrainLines
      : [...section.refrainLines, detected];
    if (
      arraysEqual(nextRefrain, section.refrainLines) &&
      arraysEqual(nextVerses, section.verseLines)
    ) {
      return null;
    }
    return {
      ...section,
      refrainLines: nextRefrain,
      verseLines: nextVerses,
    };
  };

  const withDetectedRefrainFromRefrain = (section: SectionState): SectionState | null => {
    if (section.type !== "musica" || section.refrainLines.length === 0) {
      return null;
    }

    const detected = detectRefrainFromBlocks(section.refrainLines) ?? section.refrainLines[0];
    const refrainKey = normalizeRefrainBlockKey(detected);
    const nextRefrain = [detected];
    const nextVerses = section.refrainLines.filter(
      (block) => normalizeRefrainBlockKey(block) !== refrainKey,
    );
    if (
      arraysEqual(nextRefrain, section.refrainLines) &&
      arraysEqual(nextVerses, section.verseLines)
    ) {
      return null;
    }

    return {
      ...section,
      refrainLines: nextRefrain,
      verseLines: nextVerses,
    };
  };

  // Mantem a mesma priorizacao da aplicacao Python.
  const withAutoDetectedRefrain = (section: SectionState): SectionState | null => {
    if (section.type !== "musica") {
      return null;
    }
    const hasVerses = section.verseLines.length > 0;
    const hasRefrain = section.refrainLines.length > 0;
    if (hasVerses && !hasRefrain) {
      return withDetectedRefrainFromVerses(section);
    }
    if (hasRefrain && !hasVerses) {
      return withDetectedRefrainFromRefrain(section);
    }
    if (hasVerses && hasRefrain) {
      return withDetectedRefrainFromVerses(section);
    }
    return null;
  };

  const detectRefrain = (id: string) => {
    const section = sections[id];
    if (!section || section.type !== "musica") {
      return;
    }
    const next = withDetectedRefrainFromVerses(section);
    if (!next) {
      setStatusMessage("Não foi possível detectar refrão nesta seção.", "error");
      return;
    }
    updateSection(id, () => next);
    setStatusMessage("Refrão detectado e reorganizado.", "success");
  };

  const scheduleAutoDetectRefrain = (id: string) => {
    window.setTimeout(() => {
      setSections((current) => {
        const section = current[id];
        if (!section || section.type !== "musica" || !section.autoDetectRefrain) {
          return current;
        }
        const next = withAutoDetectedRefrain(section);
        if (!next) {
          return current;
        }
        return { ...current, [id]: next };
      });
    }, 0);
  };

  const applyLiturgiaData = useCallback((data: Record<string, string>) => {
    updateByCanonical("1ª Leitura", (section) => ({
      ...section,
      whiteTextLines: data.firstReadingRef ? [data.firstReadingRef] : [],
    }));
    updateByCanonical("Salmo", (section) => ({
      ...section,
      yellowTitleLines: data.psalmTitle ? [data.psalmTitle] : section.yellowTitleLines,
      whiteTextLines: data.psalmResponse ? [data.psalmResponse] : [],
    }));
    updateByCanonical("2ª Leitura", (section) => ({
      ...section,
      whiteTextLines: data.secondReadingRef ? [data.secondReadingRef] : [],
    }));
    updateByCanonical("Aclamação", (section) => ({
      ...section,
      acclamationLines: [data.gospelProclamation, data.gospelAcclamation].filter(Boolean),
      antiphonLines: [data.gospelAntiphon, data.gospelCitation].filter(Boolean),
    }));
  }, [updateByCanonical]);

  const fetchLiturgiaForDate = useCallback(async (date: string, silent = false) => {
    if (!silent) {
      setLoadingLiturgia(true);
      setStatusMessage("Buscando liturgia...", "loading");
    }
    try {
      const response = await fetch(`/api/liturgia?date=${date}`, { cache: "no-store" });
      const data = (await response.json()) as Record<string, string>;
      if (!response.ok) {
        throw new Error(data.error ?? "Falha ao buscar liturgia.");
      }

      applyLiturgiaData(data);
      if (!silent) {
        setStatusMessage("Liturgia importada.", "success");
      }
    } catch (error) {
      if (!silent) {
        setStatusMessage(
          error instanceof Error ? error.message : "Erro na liturgia.",
          "error",
        );
      }
    } finally {
      if (!silent) {
        setLoadingLiturgia(false);
      }
    }
  }, [applyLiturgiaData, setStatusMessage]);

  const fetchLiturgia = useCallback(async () => {
    await fetchLiturgiaForDate(liturgiaDate, false);
  }, [fetchLiturgiaForDate, liturgiaDate]);

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(LOCAL_STORAGE_KEY);
      if (!raw) {
        return;
      }

      const parsed = JSON.parse(raw) as Partial<PersistedEditorState>;
      const restoredSections = normalizeSections(parsed.sections);
      const restoredOrder = normalizeSectionOrderForState(parsed.sectionOrder, restoredSections);

      if (typeof parsed.presentationTitle === "string") {
        setPresentationTitle(parsed.presentationTitle);
      }
      if (isIsoDate(parsed.liturgiaDate)) {
        setLiturgiaDate(parsed.liturgiaDate);
      }

      setOptions(normalizeOptions(parsed.options));
      setSections(restoredSections);
      setSectionOrder(restoredOrder);
      setSkipInitialLiturgiaImport(true);
      setStatusMessage("Rascunho restaurado.", "success");
    } catch {
      window.localStorage.removeItem(LOCAL_STORAGE_KEY);
    } finally {
      setIsStorageHydrated(true);
    }
  }, [setStatusMessage]);

  useEffect(() => {
    if (!isStorageHydrated) {
      return;
    }
    const payload: PersistedEditorState = {
      presentationTitle,
      liturgiaDate,
      options,
      sections,
      sectionOrder,
    };

    try {
      window.localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(payload));
    } catch {
      // Ignore storage quota and availability errors.
    }
  }, [isStorageHydrated, liturgiaDate, options, presentationTitle, sectionOrder, sections]);

  useEffect(() => {
    if (!isStorageHydrated || skipInitialLiturgiaImport) {
      return;
    }
    void fetchLiturgiaForDate(initialLiturgiaDate, true);
  }, [
    fetchLiturgiaForDate,
    initialLiturgiaDate,
    isStorageHydrated,
    skipInitialLiturgiaImport,
  ]);

  useEffect(() => {
    setOpenSectionIds((current) => {
      const filtered = current.filter((id) => sectionOrder.includes(id));
      if (filtered.length > 0 || sectionOrder.length === 0) {
        return filtered;
      }
      return [sectionOrder[0]];
    });
  }, [sectionOrder]);

  useEffect(() => {
    if (sectionOrder.length === 0) {
      if (insertReferenceId !== "") {
        setInsertReferenceId("");
      }
      return;
    }
    if (!insertReferenceId || !sectionOrder.includes(insertReferenceId)) {
      setInsertReferenceId(sectionOrder[sectionOrder.length - 1] ?? "");
    }
  }, [insertReferenceId, sectionOrder]);

  const generatePptx = async () => {
    setLoadingPpt(true);
    setStatusMessage("Gerando apresentação...", "loading");
    const payload: GeneratePayload = { presentationTitle, sectionOrder, sections, options };
    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!response.ok) {
        throw new Error((await response.json()).error ?? "Falha ao gerar arquivo.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = parseFilename(response.headers.get("content-disposition"));
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.URL.revokeObjectURL(url);

      setStatusMessage("Download iniciado.", "success");
    } catch (error) {
      setStatusMessage(
        error instanceof Error ? error.message : "Erro ao gerar arquivo.",
        "error",
      );
    } finally {
      setLoadingPpt(false);
    }
  };

  const toggleOption = (key: keyof GeneratorOptions) => {
    setOptions((current) => ({ ...current, [key]: !(current[key] as boolean) }));
  };

  const moveSection = (id: string, delta: -1 | 1) => {
    setSectionOrder((current) => moveItem(current, id, delta));
  };

  const toggleSectionOpen = (id: string, open: boolean) => {
    setOpenSectionIds((current) => {
      const has = current.includes(id);
      if (open) {
        return has ? current : [...current, id];
      }
      if (!has) {
        return current;
      }
      return current.filter((item) => item !== id);
    });
  };

  const updateSectionTitle = (id: string, value: string) => {
    updateSection(id, (current) => ({ ...current, title: value }));
  };

  const toggleStartWithRefrain = (id: string) => {
    updateSection(id, (current) => ({
      ...current,
      startWithRefrain: !current.startWithRefrain,
    }));
  };

  const toggleAutoDetectRefrain = (id: string) => {
    updateSection(id, (current) => ({
      ...current,
      autoDetectRefrain: !current.autoDetectRefrain,
    }));
  };

  const restoreDefaults = async () => {
    const confirmed = window.confirm(
      "Restaurar o editor para o padrão inicial?\nIsso removerá o rascunho atual desta sessão.",
    );
    if (!confirmed) {
      return;
    }

    const nextSections = buildInitialSections();
    const nextOrder = buildInitialSectionOrder();

    setSections(nextSections);
    setSectionOrder(nextOrder);
    setOpenSectionIds(nextOrder.length > 0 ? [nextOrder[0]] : []);
    setPresentationTitle(DEFAULT_PRESENTATION_TITLE);
    setLiturgiaDate(initialLiturgiaDate);
    setOptions(DEFAULT_GENERATOR_OPTIONS);
    setNewSectionName("");
    setNewSectionType("musica");
    setInsertPlacement("after");
    setInsertReferenceId(nextOrder[nextOrder.length - 1] ?? "");
    setSkipInitialLiturgiaImport(false);

    try {
      window.localStorage.removeItem(LOCAL_STORAGE_KEY);
    } catch {
      // Ignore storage availability errors.
    }

    await fetchLiturgiaForDate(initialLiturgiaDate, false);
    setStatusMessage("Editor restaurado para o padrão inicial.", "warning");
  };

  const canAddSection = newSectionName.trim().length > 0;

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top_left,#f5f0d3_0%,#f4f7fb_45%,#eef1f7_100%)] p-3 pb-28 sm:p-4 sm:pb-24 md:p-8 md:pb-24">
      <main className="mx-auto flex max-w-6xl flex-col gap-4">
        <h1 className="sr-only">Editor de slides litúrgicos</h1>
        <section className="rounded-xl border border-stone-200 bg-white p-4">
          <label
            htmlFor="presentationTitle"
            className="mb-1 block text-sm font-semibold text-stone-800"
          >
            Título da apresentação
          </label>
          <textarea
            id="presentationTitle"
            value={presentationTitle}
            onChange={(event) => setPresentationTitle(event.target.value)}
            className="mb-2 min-h-20 w-full rounded border border-stone-300 p-2 text-sm"
            placeholder="Ex.: DOMINGO DA QUARESMA"
          />
          <p className="mb-3 text-xs text-stone-500">
            Use uma linha por bloco de título. Exemplo: linha 1 com o tema e linha 2 com o tempo
            litúrgico.
          </p>
          <div className="flex flex-col gap-3 sm:flex-row sm:flex-wrap sm:items-end">
            <label className="space-y-1 text-sm">
              <span className="font-medium text-stone-700">Data da liturgia</span>
              <input
                id="liturgiaDate"
                type="date"
                value={liturgiaDate}
                onChange={(event) => setLiturgiaDate(event.target.value)}
                className="rounded border border-stone-300 px-2 py-1 text-sm"
              />
            </label>
            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={fetchLiturgia}
                disabled={loadingLiturgia || loadingPpt}
                className="rounded border border-sky-300 bg-sky-50 px-3 py-2 text-sm font-semibold text-sky-800 hover:bg-sky-100 transition-colors"
              >
                {loadingLiturgia ? "Buscando..." : "📅 Buscar liturgia"}
              </button>
              <button
                onClick={generatePptx}
                disabled={loadingPpt || loadingLiturgia}
                className="hidden sm:inline-flex rounded-lg bg-gradient-to-r from-amber-500 to-amber-600 px-5 py-2 text-sm font-bold text-white shadow-md hover:from-amber-600 hover:to-amber-700 hover:shadow-lg transition-all disabled:opacity-50"
              >
                {loadingPpt ? "⏳ Gerando..." : "▶ Gerar .pptx"}
              </button>
            </div>
            <p
              role="status"
              aria-live="polite"
              className={`text-sm font-medium ${STATUS_TONE_CLASSES[status.tone]}`}
            >
              {status.message}
            </p>
          </div>
        </section>

        <section className="rounded-xl border border-stone-200 bg-white p-4">
          {/* Sub-block: Add new section */}
          <p className="mb-2 text-sm font-semibold text-stone-800">➕ Adicionar seção</p>
          <div className="mb-2 flex flex-col gap-2 sm:flex-row sm:flex-wrap sm:items-end">
            <label className="space-y-1 text-sm">
              <span className="font-medium text-stone-700">Nome</span>
              <input
                ref={newSectionNameInputRef}
                value={newSectionName}
                onChange={(event) => setNewSectionName(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    event.preventDefault();
                    void addSection();
                  }
                }}
                placeholder="Nova seção"
                className="rounded border border-stone-300 px-2 py-1 text-sm"
              />
            </label>
            <label className="space-y-1 text-sm">
              <span className="font-medium text-stone-700">Tipo</span>
              <select
                value={newSectionType}
                onChange={(event) => setNewSectionType(event.target.value as SectionType)}
                className="rounded border border-stone-300 px-2 py-1 text-sm"
              >
                {SECTION_TYPE_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
            <label className="space-y-1 text-sm">
              <span className="font-medium text-stone-700">Posição</span>
              <select
                value={insertPlacement}
                onChange={(event) => setInsertPlacement(event.target.value as InsertPlacement)}
                className="rounded border border-stone-300 px-2 py-1 text-sm"
              >
                <option value="after">Inserir após</option>
                <option value="before">Inserir antes</option>
              </select>
            </label>
            <label className="space-y-1 text-sm">
              <span className="font-medium text-stone-700">Referência</span>
              <select
                value={insertReferenceId}
                onChange={(event) => setInsertReferenceId(event.target.value)}
                className="rounded border border-stone-300 px-2 py-1 text-sm"
                disabled={sectionOrder.length === 0}
              >
                {sectionOrder.map((id) => (
                  <option key={id} value={id}>
                    {sections[id]?.name ?? id}
                  </option>
                ))}
              </select>
            </label>
            <button
              onClick={addSection}
              disabled={!canAddSection}
              className={`rounded px-3 py-1 text-sm font-semibold text-white ${
                canAddSection
                  ? "bg-stone-800 hover:bg-stone-900 transition-colors"
                  : "cursor-not-allowed bg-stone-400"
              }`}
            >
              + Adicionar
            </button>
          </div>
          <p className="text-xs text-stone-500">
            Dica: escolha a posição e a referência para inserir a nova seção no ponto certo da
            celebração.
          </p>

          {/* Separator */}
          <hr className="my-3 border-stone-200" />

          {/* Sub-block: Fixed liturgical elements */}
          <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-stone-500">Elementos fixos da celebração</p>
          <div className="mb-3 grid grid-cols-2 gap-x-4 gap-y-1 sm:flex sm:flex-wrap">
            {BOOLEAN_OPTION_KEYS.map((key) => (
              <label key={key} className="flex items-center gap-1.5 text-sm">
                <input type="checkbox" checked={Boolean(options[key])} onChange={() => toggleOption(key)} className="accent-amber-500" />
                {BOOLEAN_OPTION_LABELS[key]}
              </label>
            ))}
          </div>

          {/* Separator */}
          <hr className="my-3 border-stone-200" />

          {/* Sub-block: General actions */}
          <button
            onClick={() => void restoreDefaults()}
            className="rounded border border-stone-300 px-3 py-1 text-sm font-semibold text-stone-600 hover:bg-stone-100 transition-colors"
          >
            ↺ Restaurar padrão
          </button>
        </section>

        {sectionOrder.map((id, index) => {
          const section = sections[id];
          if (!section) {
            return null;
          }

          return (
            <SectionEditorCard
              key={id}
              id={id}
              isOpen={openSectionIds.includes(id)}
              canMoveUp={index > 0}
              canMoveDown={index < sectionOrder.length - 1}
              section={section}
              onToggleOpen={toggleSectionOpen}
              onMove={moveSection}
              onDetectRefrain={detectRefrain}
              onDelete={requestDeleteSection}
              onUpdateTitle={updateSectionTitle}
              onToggleStartWithRefrain={toggleStartWithRefrain}
              onToggleAutoDetectRefrain={toggleAutoDetectRefrain}
              onUpdateLines={updateLines}
              onScheduleAutoDetectRefrain={scheduleAutoDetectRefrain}
              onUpdateStyle={updateStyle}
            />
          );
        })}
      </main>

      {/* Floating Action Button — sempre visível */}
      <button
        onClick={generatePptx}
        disabled={loadingPpt || loadingLiturgia}
        className="fixed bottom-4 right-4 z-50 flex items-center gap-2 rounded-full bg-gradient-to-r from-amber-500 to-amber-600 px-4 py-2.5 text-xs font-bold text-white shadow-lg sm:bottom-6 sm:right-6 sm:px-6 sm:py-3 sm:text-sm hover:from-amber-600 hover:to-amber-700 hover:shadow-xl transition-all disabled:opacity-50 disabled:cursor-not-allowed"
        title="Gerar apresentação PowerPoint"
      >
        {loadingPpt ? "⏳ Gerando..." : "▶ Gerar .pptx"}
      </button>
    </div>
  );
}


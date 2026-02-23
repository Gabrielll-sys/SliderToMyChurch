"use client";

import { useCallback, useEffect, useMemo, useState } from "react";

import {
  buildInitialSectionOrder,
  buildInitialSections,
  buildUniqueSectionId,
  createEmptySection,
  DEFAULT_GENERATOR_OPTIONS,
  detectRefrainFromBlocks,
  joinTextareaBlocks,
  joinTextarea,
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
  { value: "musica", label: "Musica" },
  { value: "leitura", label: "Leitura" },
  { value: "aclamacao", label: "Aclamacao" },
  { value: "palavra", label: "Palavra" },
];

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
    { field: "refrainLines", label: "Refrao", styleKey: "refrain" },
    { field: "verseLines", label: "Versos", styleKey: "verse" },
  ],
  leitura: [
    { field: "yellowTitleLines", label: "Titulo amarelo", styleKey: "yellowTitle" },
    { field: "whiteTextLines", label: "Texto branco", styleKey: "whiteText" },
  ],
  aclamacao: [
    { field: "acclamationLines", label: "Aclamacao", styleKey: "acclamation" },
    { field: "antiphonLines", label: "Antifona", styleKey: "antiphon" },
  ],
  palavra: [{ field: "wordLines", label: "Palavra", styleKey: "word" }],
};

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
  const match = contentDisposition?.match(/filename="([^"]+)"/i);
  return match?.[1] ?? "Missa.pptx";
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
}: {
  style: TextStyle;
  onPatch: (patch: Partial<TextStyle>) => void;
}) {
  return (
    <div className="space-y-2 rounded-lg border border-stone-200 bg-stone-50 p-2">
      <div className="grid gap-2 sm:grid-cols-5">
        <select
          value={style.fontFace}
          onChange={(event) => onPatch({ fontFace: event.target.value })}
          className="rounded border border-stone-300 px-2 py-1 text-xs"
        >
          {FONT_OPTIONS.map((font) => (
            <option key={font}>{font}</option>
          ))}
        </select>
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
          className="rounded border border-stone-300 px-2 py-1 text-xs"
        />
        <label className="flex items-center gap-1 text-xs">
          <input
            type="checkbox"
            checked={style.bold}
            onChange={() => onPatch({ bold: !style.bold })}
          />
          Negrito
        </label>
        <label className="flex items-center gap-1 text-xs">
          <input
            type="checkbox"
            checked={style.italic}
            onChange={() => onPatch({ italic: !style.italic })}
          />
          Itálico
        </label>
        <label className="flex items-center gap-1 text-xs">
          <input
            type="checkbox"
            checked={style.uppercase}
            onChange={() => onPatch({ uppercase: !style.uppercase })}
          />
          Maiúsculas
        </label>
      </div>
      <div className="grid gap-2 sm:grid-cols-4">
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
          className="rounded border border-stone-300 px-2 py-1 text-xs"
          title="Espaçamento entre linhas"
        />
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
          className="rounded border border-stone-300 px-2 py-1 text-xs"
          title="Preenchimento mínimo do slide"
        />
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
          className="rounded border border-stone-300 px-2 py-1 text-xs"
          title="Linhas mínimas na última página"
        />
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
          className="rounded border border-stone-300 px-2 py-1 text-xs"
          title="Máximo de linhas por slide (vazio = automático)"
          placeholder="hard max"
        />
      </div>
    </div>
  );
}

type SectionEditorCardProps = {
  id: string;
  index: number;
  section: SectionState;
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
  index,
  section,
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
  return (
    <details open={index === 0} className="rounded-xl border border-stone-200 bg-white">
      <summary className="flex cursor-pointer justify-between px-4 py-2 text-sm font-semibold">
        <span>
          {(section.title || section.name)} ({section.type})
        </span>
        <span>{">"}</span>
      </summary>

      <div className="space-y-3 border-t border-stone-200 p-4">
        <div className="flex flex-wrap gap-2">
          <button
            onClick={() => onMove(id, -1)}
            className="rounded border border-stone-300 px-2 py-1 text-xs"
          >
            Mover Seção Para Cima
          </button>
          <button
            onClick={() => onMove(id, 1)}
            className="rounded border border-stone-300 px-2 py-1 text-xs"
          >
            Mover Seção Para Baixo
          </button>
          {section.type === "musica" && (
            <button
              onClick={() => onDetectRefrain(id)}
              className="rounded border border-amber-300 px-2 py-1 text-xs"
            >
              Detectar refrao
            </button>
          )}
          <button
            onClick={() => onDelete(id)}
            className="rounded border border-red-300 px-2 py-1 text-xs"
          >
            Excluir secao
          </button>
        </div>

        <input
          value={section.title}
          onChange={(event) => onUpdateTitle(id, event.target.value)}
          className="w-full rounded border border-stone-300 px-2 py-1 text-sm"
        />

        {section.type === "musica" && (
          <div className="flex flex-wrap gap-3">
            <label className="flex items-center gap-2 text-xs">
              <input
                type="checkbox"
                checked={section.startWithRefrain}
                onChange={() => onToggleStartWithRefrain(id)}
              />
              Iniciar com refrao
            </label>
            <label className="flex items-center gap-2 text-xs">
              <input
                type="checkbox"
                checked={section.autoDetectRefrain}
                onChange={() => onToggleAutoDetectRefrain(id)}
              />
              Detectar e reorganizar refrao automaticamente ao colar
            </label>
          </div>
        )}

        <StyleEditor
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
              <p className="text-xs font-semibold">{spec.label}</p>
              <textarea
                value={value}
                onChange={(event) => onUpdateLines(id, spec.field, event.target.value)}
                onPaste={() => {
                  if (isMusicField) {
                    onScheduleAutoDetectRefrain(id);
                  }
                }}
                className="min-h-32 w-full rounded border border-stone-300 p-2 text-sm"
              />
              <StyleEditor
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
  const [presentationTitle, setPresentationTitle] = useState("DOMINGO DA\nQUARESMA");
  const [liturgiaDate, setLiturgiaDate] = useState(initialLiturgiaDate);
  const [options, setOptions] = useState<GeneratorOptions>(DEFAULT_GENERATOR_OPTIONS);
  const [status, setStatus] = useState("Pronto.");
  const [loadingLiturgia, setLoadingLiturgia] = useState(false);
  const [loadingPpt, setLoadingPpt] = useState(false);
  const [newSectionName, setNewSectionName] = useState("");
  const [newSectionType, setNewSectionType] = useState<SectionType>("musica");
  const [insertReferenceId, setInsertReferenceId] = useState("");
  const [insertPlacement, setInsertPlacement] = useState<InsertPlacement>("after");

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
    setInsertReferenceId(id);
    setNewSectionName("");
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
      setStatus("Nao foi possivel detectar refrao.");
      return;
    }
    updateSection(id, () => next);
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
    }
    try {
      const response = await fetch(`/api/liturgia?date=${date}`, { cache: "no-store" });
      const data = (await response.json()) as Record<string, string>;
      if (!response.ok) {
        throw new Error(data.error ?? "Falha ao buscar liturgia.");
      }

      applyLiturgiaData(data);
      if (!silent) {
        setStatus("Liturgia importada.");
      }
    } catch (error) {
      if (!silent) {
        setStatus(error instanceof Error ? error.message : "Erro na liturgia.");
      }
    } finally {
      if (!silent) {
        setLoadingLiturgia(false);
      }
    }
  }, [applyLiturgiaData]);

  const fetchLiturgia = useCallback(async () => {
    await fetchLiturgiaForDate(liturgiaDate, false);
  }, [fetchLiturgiaForDate, liturgiaDate]);

  useEffect(() => {
    void fetchLiturgiaForDate(initialLiturgiaDate, true);
  }, [fetchLiturgiaForDate, initialLiturgiaDate]);

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

      setStatus("Download iniciado.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Erro ao gerar arquivo.");
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

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top_left,#f5f0d3_0%,#f4f7fb_45%,#eef1f7_100%)] p-4 md:p-8">
      <main className="mx-auto flex max-w-6xl flex-col gap-4">
        <section className="rounded-xl border border-stone-200 bg-white p-4">
          <textarea
            value={presentationTitle}
            onChange={(event) => setPresentationTitle(event.target.value)}
            className="mb-2 min-h-40 w-full rounded border border-stone-300 p-2 text-sm"
          />
          <div className="flex flex-wrap gap-2">
            <input
              type="date"
              value={liturgiaDate}
              onChange={(event) => setLiturgiaDate(event.target.value)}
              className="rounded border border-stone-300 px-2 py-1 text-sm"
            />
            <button
              onClick={fetchLiturgia}
              disabled={loadingLiturgia || loadingPpt}
              className="rounded bg-sky-700 px-3 py-1 text-sm font-semibold text-white"
            >
              {loadingLiturgia ? "Buscando..." : "Buscar liturgia"}
            </button>
            <button
              onClick={generatePptx}
              disabled={loadingPpt || loadingLiturgia}
              className="rounded bg-amber-500 px-3 py-1 text-sm font-semibold text-stone-900"
            >
              {loadingPpt ? "Gerando..." : "Gerar .pptx"}
            </button>
            <span className="text-sm text-stone-700">{status}</span>
          </div>
        </section>

        <section className="rounded-xl border border-stone-200 bg-white p-4">
          <div className="mb-2 flex flex-wrap gap-2">
            <input
              value={newSectionName}
              onChange={(event) => setNewSectionName(event.target.value)}
              placeholder="Nova secao"
              className="rounded border border-stone-300 px-2 py-1 text-sm"
            />
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
            <select
              value={insertPlacement}
              onChange={(event) => setInsertPlacement(event.target.value as InsertPlacement)}
              className="rounded border border-stone-300 px-2 py-1 text-sm"
            >
              <option value="after">Inserir apos</option>
              <option value="before">Inserir antes</option>
            </select>
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
            <button
              onClick={addSection}
              className="rounded bg-stone-800 px-3 py-1 text-sm font-semibold text-white"
            >
              Adicionar
            </button>
          </div>

          <div className="mb-2 flex flex-wrap gap-2">
            {BOOLEAN_OPTION_KEYS.map((key) => (
              <label key={key} className="flex items-center gap-1 text-xs">
                <input type="checkbox" checked={Boolean(options[key])} onChange={() => toggleOption(key)} />
                {BOOLEAN_OPTION_LABELS[key]}
              </label>
            ))}
          </div>
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
              index={index}
              section={section}
              onMove={moveSection}
              onDetectRefrain={detectRefrain}
              onDelete={removeSection}
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
    </div>
  );
}


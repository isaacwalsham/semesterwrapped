import React, { useEffect, useMemo, useRef, useState } from "react";
import html2canvas from "html2canvas";

const makeId = (prefix) => `${prefix}_${Math.random().toString(36).slice(2, 9)}`;

const DEFAULTS = {
  personName: "Your Name",
  university: "Your University",
  course: "BSc Biology",
  semester: "Semester 1",
  year: "Year 1 • 2025/26",
  headline: "My Semester Wrapped",
  caption: "Proud of my results this semester 🎓",
  linkedinHandle: "",

  showModuleMarks: true,
  showAssessmentsInBreakdown: true,

  autoModuleWeights: true,
  modules: [
    {
      id: makeId("m"),
      code: "CS301",
      title: "Machine Learning",
      credits: 0,
      weight: 100,
      autoComponentWeights: true,
      components: [
        { id: makeId("c"), type: "Coursework", name: "Coursework 1", weight: 40, mark: 0 },
        { id: makeId("c"), type: "Exam", name: "Exam", weight: 60, mark: 0 },
      ],
    },
  ],

  theme: {
    bg: "#070A12",
    card: "#0B1220",
    primary: "#7C3AED",
    accent: "#22C55E",
    text: "#E5E7EB",
    muted: "#A3A3A3",
  },
  template: "gradient",
  format: "linkedin_square",
};

function clampNumber(n, min, max) {
  const x = Number.isFinite(n) ? n : min;
  return Math.max(min, Math.min(max, x));
}

function safeText(s, max = 140) {
  return String(s ?? "").slice(0, max);
}

function formatPct(n) {
  if (!Number.isFinite(n)) return "–";
  return `${Math.round(n)}%`;
}

function hexAlpha(hex, a) {
  if (!hex || typeof hex !== "string" || !hex.startsWith("#") || hex.length !== 7) return hex;
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, ${clampNumber(a, 0, 1)})`;
}

function gradientBg(theme) {
  return `radial-gradient(1200px 700px at 10% 0%, ${hexAlpha(theme.primary, 0.62)}, transparent 55%),
          radial-gradient(1000px 700px at 90% 10%, ${hexAlpha(theme.accent, 0.52)}, transparent 55%),
          radial-gradient(900px 700px at 35% 95%, ${hexAlpha(theme.text, 0.10)}, transparent 60%),
          ${theme.bg}`;
}

function getDims(format) {
  const dims = {
    linkedin_square: { w: 1200, h: 1200, label: "LinkedIn (Square)" },
    instagram_story: { w: 1080, h: 1920, label: "Instagram (Story)" },
  };
  return dims[format] || dims.linkedin_square;
}

function getUkClassification(avg) {
  const n = Number(avg);
  if (!Number.isFinite(n)) return { label: "—", short: "—" };
  if (n >= 70) return { label: "First Class", short: "First Class" };
  if (n >= 60) return { label: "Upper Second (2:1)", short: "2:1" };
  if (n >= 50) return { label: "Lower Second (2:2)", short: "2:2" };
  if (n >= 40) return { label: "Third Class", short: "Third Class" };
  return { label: "Fail", short: "Fail" };
}

function normalizeWeights(items, key) {
  const sum = (items || []).reduce((a, it) => a + (Number(it?.[key]) || 0), 0);
  if (!sum) return items.map((it) => ({ ...it, __norm: 0 }));
  return items.map((it) => ({ ...it, __norm: ((Number(it?.[key]) || 0) / sum) * 100 }));
}

function evenSplit(n) {
  if (n <= 0) return [];
  const base = Math.floor(10000 / n) / 100; 

  const weights = Array.from({ length: n }, () => base);
  const total = weights.reduce((a, b) => a + b, 0);
  const rem = Math.round((100 - total) * 100) / 100;
  weights[0] = Math.round((weights[0] + rem) * 100) / 100;
  return weights;
}

function computeModuleMark(module) {
  const comps = module?.components || [];
  const active = comps.filter((c) => c && (c.name || c.type));
  if (active.length === 0) return { mark: NaN, weightSum: 0 };

  const sumW = active.reduce((a, c) => a + (Number(c.weight) || 0), 0);
  const denom = sumW > 0 ? sumW : 100;

  const weighted = active.reduce((acc, c) => {
    const w = Number(c.weight) || 0;
    const mk = Number(c.mark);
    const safeMk = Number.isFinite(mk) ? mk : 0;
    return acc + safeMk * (w / denom);
  }, 0);

  return { mark: clampNumber(weighted, 0, 100), weightSum: sumW };
}

function computeSemester(modules) {
  const ms = (modules || []).filter((m) => m && (m.code || m.title || (m.components || []).length));
  if (ms.length === 0) {
    return { avg: NaN, modules: [], best: null, count: 0 };
  }

  const withMarks = ms.map((m) => {
    const res = computeModuleMark(m);
    return { ...m, moduleMark: res.mark, __componentWeightSum: res.weightSum };
  });

  const norm = normalizeWeights(withMarks, "weight");
  const avg = norm.reduce((acc, m) => {
    const mk = Number(m.moduleMark);
    const safeMk = Number.isFinite(mk) ? mk : 0;
    return acc + safeMk * ((m.__norm || 0) / 100);
  }, 0);

  const best =
    withMarks
      .filter((m) => Number.isFinite(Number(m.moduleMark)))
      .slice()
      .sort((a, b) => (Number(b.moduleMark) || 0) - (Number(a.moduleMark) || 0))[0] || null;

  return { avg, modules: withMarks, best, count: withMarks.length };
}

function normalizeHandle(s) {
  const t = String(s || "").trim();
  if (!t) return "";
  return t.startsWith("@") ? t : `@${t}`;
}

function sanitizeFileName(name) {
  const raw = String(name || "").trim();
  const cleaned = raw
    .replace(/\.png$/i, "") 

    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "") // windows-illegal + control chars
    .replace(/\s+/g, " ")
    .trim();

  const safe = cleaned && cleaned.replace(/\./g, "").length ? cleaned : "semester-wrapped";
  return safe.slice(0, 80);
}

function Field({ label, children }) {
  return (
    <label style={{ display: "block" }}>
      <div
        style={{
          fontSize: 12,
          fontWeight: 850,
          letterSpacing: "-0.01em",
          opacity: 0.78,
          marginBottom: 6,
        }}
      >
        {label}
      </div>
      {children}
    </label>
  );
}

function Input(props) {
  return <input {...props} className={`sw-input ${props.className || ""}`} />;
}

function Select(props) {
  return <select {...props} className={`sw-input ${props.className || ""}`} />;
}

function Button({ variant = "primary", ...props }) {
  return <button {...props} className={`sw-btn sw-btn--${variant} ${props.className || ""}`} />;
}

function SizeHint({ format }) {
  const f = getDims(format);
  return (
    <div className="sw-sizehint" style={{ fontSize: 15 }}>
      Export size:{" "}
      <span className="sw-sizehint__pill" style={{ fontSize: 14 }}>
        {f.w}×{f.h}
      </span>{" "}
      <span className="sw-sizehint__muted">— {f.label}</span>
    </div>
  );
}

function ExportFrame({ format, children, containerRef }) {
  const { w, h } = getDims(format);
  const [box, setBox] = useState({ width: 520, height: 520 });
  const headRef = useRef(null);
  const [headerH, setHeaderH] = useState(44);

  useEffect(() => {
    const el = containerRef?.current;
    if (!el) return;

    if (typeof ResizeObserver === "undefined") {
      const update = () => {
        const rect = el.getBoundingClientRect?.();
        if (!rect) return;
        setBox({ width: rect.width, height: rect.height });
      };

      update();
      window.addEventListener("resize", update);
      return () => window.removeEventListener("resize", update);
    }

    const ro = new ResizeObserver((entries) => {
      const cr = entries[0]?.contentRect;
      if (!cr) return;
      setBox({ width: cr.width, height: cr.height });
    });

    ro.observe(el);
    return () => ro.disconnect();
  }, [containerRef]);

  useEffect(() => {
    const el = headRef.current;
    if (!el) return;

    const update = () => {
      const rect = el.getBoundingClientRect?.();
      if (!rect) return;
      setHeaderH(Math.max(44, Math.round(rect.height)));
    };

    update();

    if (typeof ResizeObserver === "undefined") {
      window.addEventListener("resize", update);
      return () => window.removeEventListener("resize", update);
    }

    const ro = new ResizeObserver(() => update());
    ro.observe(el);
    return () => ro.disconnect();
  }, [format]);

  const pad = 10;
  const availW = Math.max(200, box.width - pad * 2);
  const availH = Math.max(200, box.height - pad * 2 - headerH);
  const scale = Math.min(1, availW / w, availH / h);
  const scaledW = Math.round(w * scale);
  const scaledH = Math.round(h * scale);

  return (
    <div className="sw-preview">
      <div className="sw-preview__head" ref={headRef}>
        <div className="sw-preview__title">Preview</div>
        <SizeHint format={format} />
      </div>

      <div className="sw-preview__stage">
        <div style={{ width: scaledW, height: scaledH, display: "flex", justifyContent: "center", alignItems: "flex-start" }}>
          <div style={{ width: w, height: h, transform: `scale(${scale})`, transformOrigin: "top center" }}>{children}</div>
        </div>
      </div>
    </div>
  );
}

function WrappedCard({ state }) {
  const {
    theme,
    modules,
    headline,
    caption,
    personName,
    university,
    course,
    semester,
    year,
    format,
    template,
    linkedinHandle,
    showModuleMarks,
    showAssessmentsInBreakdown,
  } = state;

  const stats = useMemo(() => computeSemester(modules), [modules]);
  const cls = getUkClassification(stats.avg);

  const handleLine = useMemo(() => {
    const li = normalizeHandle(linkedinHandle);
    return li ? `LinkedIn - ${li}` : "";
  }, [linkedinHandle]);

  const rankedModules = useMemo(() => {
    return (stats.modules || []).slice().sort((a, b) => (Number(b.moduleMark) || 0) - (Number(a.moduleMark) || 0));
  }, [stats.modules]);

  const layout = useMemo(() => {
    if (format === "instagram_story") {
      return {
        innerPad: 62,
        innerGap: 22,
        headline: 76,
        name: 44,
        handle: 24,
        meta: 34,
        pct: 96,
        cls: 22,
        sectionTitle: 36,
        sectionSub: 30,
        empty: 28,
        footer: 52,
        hashtag: 52,
        footerMax: "70%",
        breakdownScale: 1.24,
      };
    }
    return {
      innerPad: 54,
      innerGap: 18,
      headline: 62,
      name: 34,
      handle: 20,
      meta: 28,
      pct: 72,
      cls: 18,
      sectionTitle: 28,
      sectionSub: 24,
      empty: 22,
      footer: 36,
      hashtag: 36,
      footerMax: "72%",
      breakdownScale: 1.08,
    };
  }, [format]);

  const scaled = (n) => Math.max(10, Math.round(n * layout.breakdownScale));
  const assessmentLimit = format === "instagram_story" ? 5 : 3;

  const breakdownSizing = useMemo(() => {
    const count = rankedModules.length;

    if (count <= 1) {
      return { rowPadY: 14, rowPadX: 16, rank: 16, code: 26, title: 30, weight: 24, mark: 34 };
    }
    if (count <= 3) {
      return { rowPadY: 13, rowPadX: 15, rank: 15, code: 24, title: 27, weight: 22, mark: 31 };
    }
    if (count <= 4) {
      return { rowPadY: 12, rowPadX: 14, rank: 14, code: 22, title: 25, weight: 21, mark: 29 };
    }
    if (count <= 7) {
      return { rowPadY: 10, rowPadX: 12, rank: 12, code: 19, title: 21, weight: 17, mark: 24 };
    }
    return { rowPadY: 8, rowPadX: 10, rank: 11, code: 17, title: 19, weight: 15, mark: 21 };
  }, [rankedModules.length]);

  const bg = template === "gradient" ? gradientBg(theme) : theme.bg;

  return (
    <div
      className="sw-card"
      style={{
        background: bg,
        color: theme.text,
        borderRadius: 28,
        overflow: "hidden",
        isolation: "isolate",
      }}
    >
      <div
        className="sw-card__grain"
        style={{
          backgroundImage:
            "repeating-linear-gradient(0deg, rgba(255,255,255,0.035) 0px, rgba(255,255,255,0.035) 1px, rgba(255,255,255,0) 2px, rgba(255,255,255,0) 4px)," +
            "repeating-linear-gradient(90deg, rgba(255,255,255,0.025) 0px, rgba(255,255,255,0.025) 1px, rgba(255,255,255,0) 2px, rgba(255,255,255,0) 5px)," +
            "radial-gradient(circle at 20% 30%, rgba(255,255,255,0.06), transparent 55%)," +
            "radial-gradient(circle at 70% 65%, rgba(0,0,0,0.08), transparent 55%)",
          opacity: 0.14,
          mixBlendMode: "normal",
        }}
      />

      <div className="sw-card__inner" style={{ padding: layout.innerPad, gap: layout.innerGap }}>
        <div className="sw-card__hero">
          <div className="sw-card__headline" style={{ fontSize: layout.headline }}>
            {safeText(headline, 42)}
          </div>
          <div className="sw-card__nameRow" style={{ marginTop: 12 }}>
            <div className="sw-card__name" style={{ fontSize: layout.name }}>
              {safeText(personName, 34)}
            </div>
            {handleLine ? (
              <div className="sw-card__handle" style={{ fontSize: layout.handle }}>
                {handleLine}
              </div>
            ) : null}
          </div>
          <div className="sw-card__meta" style={{ fontSize: layout.meta }}>
            <div className="sw-card__metaLine">
              <span className="sw-card__metaStrong">{safeText(university, 44)}</span>
              <span className="sw-card__dot">•</span>
              <span className="sw-card__metaMuted">{safeText(course, 48)}</span>
            </div>
            <div className="sw-card__metaLine">
              <span className="sw-card__metaMuted">{safeText(semester, 18)}</span>
              <span className="sw-card__dot">•</span>
              <span className="sw-card__metaMuted">{safeText(year, 16)}</span>
            </div>
          </div>

          <div className="sw-card__gradeRow">
            <div className="sw-card__pct" style={{ fontSize: layout.pct }}>
              {stats.count ? formatPct(stats.avg) : "—"}
            </div>
            <div className="sw-card__class" style={{ fontSize: layout.cls }}>
              {stats.count ? cls.label : "Add modules to calculate"}
            </div>
          </div>
        </div>

        <div className="sw-card__breakdown" style={{ marginTop: 10 }}>
          <div className="sw-card__sectionHead">
            <div className="sw-card__sectionTitle" style={{ fontSize: layout.sectionTitle, fontWeight: 900 }}>
              Module breakdown
            </div>
            <div className="sw-card__sectionSub" style={{ fontSize: layout.sectionSub, opacity: 0.85 }}>
              Ranked by module mark
            </div>
          </div>

          <div className="sw-card__list">
            {rankedModules.length === 0 ? (
              <div className="sw-card__empty" style={{ fontSize: layout.empty, opacity: 0.85 }}>
                Add modules and assessments to populate your wrap.
              </div>
            ) : (
              rankedModules.map((m, i) => (
                <div
                  key={m.id || `${m.code}-${i}`}
                  className="sw-card__row"
                  style={{
                    background: hexAlpha(theme.card, 0.55),
                    border: `1px solid ${hexAlpha(theme.text, 0.08)}`,
                    padding: `${scaled(breakdownSizing.rowPadY)}px ${scaled(breakdownSizing.rowPadX)}px`,
                  }}
                >
                  <div className="sw-card__rank" style={{ fontSize: scaled(breakdownSizing.rank) }}>
                    {String(i + 1).padStart(2, "0")}
                  </div>

                  <div className="sw-card__rowMain">
                    <div className="sw-card__rowTop">
                      <div className="sw-card__code" style={{ fontSize: scaled(breakdownSizing.code), fontWeight: 800 }}>
                        {safeText(m.code || `Module ${i + 1}`, 16)}
                      </div>
                    </div>
                    <div className="sw-card__title" style={{ fontSize: scaled(breakdownSizing.title), fontWeight: 900, lineHeight: 1.2 }}>
                      {safeText(m.title || "Untitled module", 44)}
                    </div>
                    {showAssessmentsInBreakdown ? (
                      <div className="sw-card__assessList" style={{ marginTop: scaled(6), gap: scaled(4) }}>
                        {(m.components || [])
                          .filter((c) => c && (c.name || c.type))
                          .slice(0, assessmentLimit)
                          .map((c, cIdx) => (
                            <div key={c.id || `${m.id || i}-a-${cIdx}`} className="sw-card__assessItem" style={{ fontSize: scaled(12) }}>
                              <span className="sw-card__assessName">{safeText(c.name || c.type || `Assessment ${cIdx + 1}`, 24)}</span>
                              <span className="sw-card__assessMeta">
                                {Number(c.weight) ? `${Math.round(Number(c.weight))}%` : ""}
                                {showModuleMarks ? ` • ${formatPct(Number(c.mark))}` : ""}
                              </span>
                            </div>
                          ))}
                        {(m.components || []).filter((c) => c && (c.name || c.type)).length > assessmentLimit ? (
                          <div className="sw-card__assessMore" style={{ fontSize: scaled(11) }}>
                            +{(m.components || []).filter((c) => c && (c.name || c.type)).length - assessmentLimit} more
                          </div>
                        ) : null}
                      </div>
                    ) : null}
                  </div>

                  <div className="sw-card__rowRight">
                    <div className="sw-card__weight" style={{ fontSize: scaled(breakdownSizing.weight) }}>
                      {Number(m.weight) ? `${Math.round(Number(m.weight))}%` : ""}
                    </div>
                    {showModuleMarks ? (
                      <div className="sw-card__mark" style={{ fontSize: scaled(breakdownSizing.mark) }}>
                        {formatPct(Number(m.moduleMark))}
                      </div>
                    ) : null}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>

        <div className="sw-card__footer" style={{ fontSize: layout.footer, opacity: 0.98, lineHeight: 1.3 }}>
          <div className="sw-card__footerLeft" style={{ maxWidth: layout.footerMax }}>
            {caption ? safeText(caption, 84) : ""}
          </div>
          <div className="sw-card__footerRight" style={{ color: theme.primary, fontSize: layout.hashtag, fontWeight: 900 }}>
            #SemesterWrapped
          </div>
        </div>
      </div>
    </div>
  );
}

export default function SemesterWrappedApp() {
  const [state, setState] = useState(DEFAULTS);
  const [activeTab, setActiveTab] = useState("info");
  const [isMobileLayout, setIsMobileLayout] = useState(false);

  const previewBoxRef = useRef(null);
  const previewCardRef = useRef(null);
  const editorBodyRef = useRef(null);

  const [fileName, setFileName] = useState(() => "semester-wrapped");

  const stats = useMemo(() => computeSemester(state.modules), [state.modules]);
  const dims = getDims(state.format);

  const set = (patch) => setState((s) => ({ ...s, ...patch }));
  const setTheme = (patch) => setState((s) => ({ ...s, theme: { ...s.theme, ...patch } }));

  useEffect(() => {
    const query = "(max-width: 900px), (hover: none) and (pointer: coarse)";
    const mql = window.matchMedia(query);
    const update = () => setIsMobileLayout(!!mql.matches);
    update();
    mql.addEventListener?.("change", update);
    return () => mql.removeEventListener?.("change", update);
  }, []);

  useEffect(() => {
    const root = document.getElementById("root");
    if (!root) return;
    if (isMobileLayout) root.classList.add("sw-root--mobile-mode");
    else root.classList.remove("sw-root--mobile-mode");
    return () => root.classList.remove("sw-root--mobile-mode");
  }, [isMobileLayout]);

  useEffect(() => {
    if (!state.autoModuleWeights) return;
    const n = state.modules.length;
    const weights = evenSplit(n);
    setState((s) => ({
      ...s,
      modules: s.modules.map((m, i) => ({ ...m, weight: weights[i] ?? 0 })),
    }));

  }, [state.autoModuleWeights, state.modules.length]);

  useEffect(() => {
    const el = editorBodyRef.current;
    if (!el || activeTab !== "modules") return;

    // Clamp scroll after module list shrinks to avoid blank panel viewport.
    const raf = requestAnimationFrame(() => {
      const maxScroll = Math.max(0, el.scrollHeight - el.clientHeight);
      if (el.scrollTop > maxScroll) el.scrollTop = maxScroll;
    });

    return () => cancelAnimationFrame(raf);
  }, [state.modules.length, activeTab]);

  const makeEmptyComponent = () => ({
    id: makeId("c"),
    type: "Overall",
    name: "Overall",
    weight: 100,
    mark: 0,
  });

  const makeEmptyModule = () => ({
    id: makeId("m"),
    code: "",
    title: "",
    credits: 0,
    weight: 100,
    autoComponentWeights: true,
    components: [makeEmptyComponent()],
  });

  const addModule = () =>
    setState((s) => ({
      ...s,
      modules: [...s.modules, makeEmptyModule()],
    }));

  const updateModule = (idx, patch) =>
    setState((s) => {
      const modules = s.modules.slice();
      modules[idx] = { ...modules[idx], ...patch };
      return { ...s, modules };
    });

  const removeModule = (idx) =>
    setState((s) => {
      const next = s.modules.filter((_, i) => i !== idx);
      return { ...s, modules: next.length ? next : [makeEmptyModule()] };
    });

  const addComponent = (moduleIdx) => {
    setState((s) => {
      const modules = s.modules.slice();
      const m = modules[moduleIdx];

      const nextComp = {
        id: makeId("c"),
        type: "Assignment",
        name: "Assignment",
        weight: 0,
        mark: 0,
      };

      const next = { ...m, components: [...(m.components || []), nextComp] };

      if (next.autoComponentWeights) {
        const weights = evenSplit(next.components.length);
        next.components = next.components.map((c, i) => ({ ...c, weight: weights[i] ?? 0 }));
      }

      modules[moduleIdx] = next;
      return { ...s, modules };
    });
  };

  const updateComponent = (moduleIdx, compIdx, patch) => {
    setState((s) => {
      const modules = s.modules.slice();
      const m = modules[moduleIdx];
      const components = (m.components || []).slice();
      components[compIdx] = { ...components[compIdx], ...patch };
      modules[moduleIdx] = { ...m, components };
      return { ...s, modules };
    });
  };

  const removeComponent = (moduleIdx, compIdx) => {
    setState((s) => {
      const modules = s.modules.slice();
      const m = modules[moduleIdx];
      const nextComps = (m.components || []).filter((_, i) => i !== compIdx);
      const ensured = nextComps.length ? nextComps : [makeEmptyComponent()];

      let nextModule = { ...m, components: ensured };
      if (nextModule.autoComponentWeights) {
        const weights = evenSplit(nextModule.components.length);
        nextModule.components = nextModule.components.map((c, i) => ({ ...c, weight: weights[i] ?? 0 }));
      }

      modules[moduleIdx] = nextModule;
      return { ...s, modules };
    });
  };

  const setAutoComponentWeights = (moduleIdx, enabled) => {
    setState((s) => {
      const modules = s.modules.slice();
      const m = modules[moduleIdx];
      let next = { ...m, autoComponentWeights: enabled };

      if (enabled) {
        const weights = evenSplit((next.components || []).length);
        next.components = (next.components || []).map((c, i) => ({ ...c, weight: weights[i] ?? 0 }));
      }

      modules[moduleIdx] = next;
      return { ...s, modules };
    });
  };

  function buildCaptionText() {
    const cls = getUkClassification(stats.avg);
    const linkedInLine = state.linkedinHandle ? `LinkedIn: ${normalizeHandle(state.linkedinHandle)}` : "";

    const caption =
      `📚 ${state.headline}\n` +
      `${state.personName}\n` +
      `${state.university} • ${state.course}\n` +
      `${state.semester} • ${state.year}\n\n` +
      `Results:\n` +
      `• Semester average: ${stats.count ? formatPct(stats.avg) : "–"}\n` +
      `• Classification: ${stats.count ? cls.label : "–"}\n` +
      `• Modules: ${stats.count}\n` +
      (linkedInLine ? `\n${linkedInLine}\n` : "\n") +
      `#SemesterWrapped #university #students`;

    return caption;
  }

  function copyCaption() {
    navigator.clipboard?.writeText(buildCaptionText());
  }

  async function exportPng() {
    const { w, h } = getDims(state.format);

    const sourceRoot = previewCardRef.current;
    const sourceCard = sourceRoot?.querySelector?.(".sw-card");
    if (!sourceRoot || !sourceCard) {
      console.error("Export failed: could not find preview card");
      alert("Export failed: preview is not ready.");
      return;
    }

    const finalName = `${sanitizeFileName(fileName)}.png`;

    document.body.classList.add("sw-exporting");
    let tempRoot = null;
    try {
      if (document.fonts?.ready) await document.fonts.ready;
      await new Promise((r) => requestAnimationFrame(r));

      const scale = 2;

      tempRoot = document.createElement("div");
      tempRoot.style.position = "fixed";
      tempRoot.style.left = "-20000px";
      tempRoot.style.top = "0";
      tempRoot.style.width = `${w}px`;
      tempRoot.style.height = `${h}px`;
      tempRoot.style.pointerEvents = "none";
      tempRoot.style.opacity = "1";
      tempRoot.style.zIndex = "-1";
      document.body.appendChild(tempRoot);

      const clonedCard = sourceCard.cloneNode(true);
      clonedCard.style.width = `${w}px`;
      clonedCard.style.height = `${h}px`;
      clonedCard.style.margin = "0";
      clonedCard.style.transform = "none";
      tempRoot.appendChild(clonedCard);

      const canvas = await html2canvas(clonedCard, {
        backgroundColor: null,
        scale,
        useCORS: false,
        allowTaint: false,
        logging: false,
        width: w,
        height: h,
        windowWidth: w,
        windowHeight: h,
        scrollX: 0,
        scrollY: 0,
        onclone: (doc) => {
          doc.body.classList.add("sw-exporting");
        },
      });

      const blob = await new Promise((resolve, reject) => {
        canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("toBlob returned null"))), "image/png");
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = finalName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Export failed:", err);
      alert("Export failed. Check console for details.");
    } finally {
      if (tempRoot && tempRoot.parentNode) tempRoot.parentNode.removeChild(tempRoot);
      document.body.classList.remove("sw-exporting");
    }
  }

  return (
    <div className={`sw-app ${isMobileLayout ? "sw-app--mobile-mode" : ""}`}>
      <div className={`sw-shell ${isMobileLayout ? "sw-shell--mobile-mode" : ""}`}>
        <div className="sw-topbar">
          <div className="sw-brand">
            <div className="sw-title">Semester Wrapped</div>
          </div>

          <div className="sw-actions">
            {}
            <div className="sw-file">
              <div className="sw-file__label">Filename</div>
              <div className="sw-file__row">
                <input
                  className="sw-input sw-file__input"
                  value={fileName}
                  onChange={(e) => setFileName(e.target.value)}
                  placeholder="semester-wrapped"
                />
              </div>
            </div>
            <Button
              variant="ghost"
              onClick={() => {
                setState(DEFAULTS);
                setActiveTab("info");
                setFileName("semester-wrapped");
              }}
            >
              Reset
            </Button>
            <Button onClick={exportPng}>Export PNG</Button>
          </div>
        </div>
      </div>

      <div className={`sw-workspace ${isMobileLayout ? "sw-workspace--mobile-mode" : ""}`}>
        {}
        <div className="sw-panel sw-panel--inputs">
          <div className="sw-panel-head">
            <div>
              <div className="sw-panel-title">Editor</div>
            </div>

            <div className="sw-tabs">
              <button className={`sw-tab ${activeTab === "info" ? "is-active" : ""}`} onClick={() => setActiveTab("info")} type="button">
                Info
              </button>
              <button className={`sw-tab ${activeTab === "modules" ? "is-active" : ""}`} onClick={() => setActiveTab("modules")} type="button">
                Modules
              </button>
              <button className={`sw-tab ${activeTab === "design" ? "is-active" : ""}`} onClick={() => setActiveTab("design")} type="button">
                Design
              </button>
            </div>
          </div>

          <div className="sw-panel-body" ref={editorBodyRef}>
            {activeTab === "info" ? (
              <div className="sw-form-grid">
                <Field label="Name">
                  <Input value={state.personName} onChange={(e) => set({ personName: e.target.value })} />
                </Field>
                <Field label="University">
                  <Input value={state.university} onChange={(e) => set({ university: e.target.value })} />
                </Field>
                <Field label="Course / Degree">
                  <Input value={state.course} onChange={(e) => set({ course: e.target.value })} />
                </Field>

                <Field label="Semester">
                  <Input value={state.semester} onChange={(e) => set({ semester: e.target.value })} placeholder="Semester 1" />
                </Field>
                <Field label="Course year (e.g., Year 1 • 2025/26)">
                  <Input value={state.year} onChange={(e) => set({ year: e.target.value })} placeholder="Year 1 • 2025/26" />
                </Field>

                <Field label="Headline">
                  <Input value={state.headline} onChange={(e) => set({ headline: e.target.value })} />
                </Field>

                <Field label="Caption (small footer line)">
                  <Input value={state.caption} onChange={(e) => set({ caption: e.target.value })} />
                </Field>

                <Field label="LinkedIn handle (optional)">
                  <Input value={state.linkedinHandle} onChange={(e) => set({ linkedinHandle: e.target.value })} placeholder="@yourname" />
                </Field>

                <div className="sw-inlineCard">
                  <div className="sw-inlineCard__top">
                    <div className="sw-inlineCard__title">Live result</div>
                    <div className="sw-inlineCard__pill">
                      {stats.count ? `${formatPct(stats.avg)} • ${getUkClassification(stats.avg).label}` : "Add modules to calculate"}
                    </div>
                  </div>
                </div>
              </div>
            ) : null}

            {activeTab === "modules" ? (
              <div className="sw-section">
                <div className="sw-sectionHead">
                  <div>
                    <div className="sw-sectionTitle">Modules</div>
                  </div>
                  <Button variant="ghost" onClick={addModule}>
                    + Add module
                  </Button>
                </div>

                <div className="sw-inlineCard sw-inlineCard--row">
                  <div>
                    <div className="sw-inlineCard__title">Auto module weights</div>
                    <div className="sw-inlineCard__sub">Splits modules equally and updates automatically.</div>
                  </div>
                  <label className="sw-toggle">
                    <input type="checkbox" checked={state.autoModuleWeights} onChange={(e) => set({ autoModuleWeights: e.target.checked })} />
                    <span>{state.autoModuleWeights ? "On" : "Off"}</span>
                  </label>
                </div>

                <div className="sw-modules">
                  {state.modules.map((m, idx) => {
                    const modCalc = computeModuleMark(m);
                    const modMark = Number.isFinite(modCalc.mark) ? modCalc.mark : 0;
                    const compSum = (m.components || []).reduce((a, c) => a + (Number(c.weight) || 0), 0);

                    return (
                      <details key={m.id || idx} className="sw-accordion" open={idx === 0}>
                        <summary className="sw-accordion__summary">
                          <div className="sw-accordion__left">
                            <div className="sw-accordion__title">
                              {(m.code || `Module ${idx + 1}`) + (m.title ? ` — ${m.title}` : "")}
                            </div>
                            <div className="sw-accordion__meta">
                              {state.autoModuleWeights ? "Auto weights" : `${Math.round(Number(m.weight) || 0)}% of semester`} •{" "}
                              {(m.components || []).length} assessments
                            </div>
                          </div>

                          <div className="sw-accordion__right">
                            <div className="sw-accordion__mark">{formatPct(modMark)}</div>
                          </div>
                        </summary>

                        <div className="sw-accordion__body">
                          <div className="sw-grid2">
                            <Field label="Module code">
                              <Input value={m.code} onChange={(e) => updateModule(idx, { code: e.target.value })} placeholder="e.g., CS301" />
                            </Field>
                            <Field label="Title">
                              <Input value={m.title} onChange={(e) => updateModule(idx, { title: e.target.value })} placeholder="e.g., Machine Learning" />
                            </Field>
                            <Field label="Weight (% of semester)">
                              <Input
                                type="number"
                                min={0}
                                max={100}
                                value={String(m.weight ?? "")}
                                disabled={state.autoModuleWeights}
                                onChange={(e) => updateModule(idx, { weight: clampNumber(Number(e.target.value), 0, 100) })}
                              />
                            </Field>
                            <Field label="Credits (optional)">
                              <Input
                                type="number"
                                min={0}
                                value={String(m.credits ?? "")}
                                onChange={(e) => updateModule(idx, { credits: clampNumber(Number(e.target.value), 0, 200) })}
                              />
                            </Field>
                          </div>

                          <div className="sw-inlineCard sw-inlineCard--row" style={{ marginTop: 12 }}>
                            <div>
                              <div className="sw-inlineCard__title">Calculated module mark</div>
                              <div className="sw-inlineCard__sub">Assessment weights sum: {Math.round(compSum)}%</div>
                            </div>
                            <div className="sw-inlineCard__pill">{formatPct(modMark)}</div>
                          </div>

                          <div className="sw-assessHead">
                            <div>
                              <div className="sw-sectionTitle" style={{ fontSize: 14 }}>
                                Assessments
                              </div>
                              <div className="sw-sectionSub">Weights contribute to the module mark.</div>
                            </div>
                            <div className="sw-assessHead__right">
                              <label className="sw-toggle">
                                <input type="checkbox" checked={!!m.autoComponentWeights} onChange={(e) => setAutoComponentWeights(idx, e.target.checked)} />
                                <span>Auto weights</span>
                              </label>
                              <Button variant="ghost" className="sw-assessAdd" onClick={() => addComponent(idx)}>
                                + Add
                              </Button>
                            </div>
                          </div>

                          <div className="sw-components">
                            {(m.components || []).map((c, cIdx) => (
                              <details key={c.id || cIdx} className="sw-subaccordion" open={cIdx === 0}>
                                <summary className="sw-subaccordion__summary">
                                  <div className="sw-subaccordion__left">
                                    <div className="sw-subaccordion__title">{c.name || c.type || `Assessment ${cIdx + 1}`}</div>
                                    <div className="sw-subaccordion__meta">
                                      {c.type} • {Math.round(Number(c.weight) || 0)}%
                                    </div>
                                  </div>
                                  <div className="sw-subaccordion__right">
                                    <div className="sw-subaccordion__mark">{formatPct(Number(c.mark))}</div>
                                  </div>
                                </summary>

                                <div className="sw-subaccordion__body">
                                  <div
                                    className="sw-grid5"
                                    style={{
                                      display: "grid",
                                      gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
                                      gap: 14,
                                      alignItems: "end",
                                    }}
                                  >
                                    <Field label="Type">
                                      <Select
                                        value={c.type}
                                        onChange={(e) => {
                                          const v = e.target.value;
                                          updateComponent(idx, cIdx, { type: v, name: c.name === c.type ? v : c.name });
                                        }}
                                        className="sw-minType"
                                      >
                                        <option value="Overall">Overall</option>
                                        <option value="Assignment">Assignment</option>
                                        <option value="Coursework">Coursework</option>
                                        <option value="Exam">Exam</option>
                                        <option value="Other">Other</option>
                                      </Select>
                                    </Field>

                                    <Field label="Label">
                                      <Input
                                        value={c.name}
                                        onChange={(e) => updateComponent(idx, cIdx, { name: e.target.value })}
                                        placeholder="e.g., Coursework 1"
                                        style={{ minWidth: 190 }}
                                      />
                                    </Field>

                                    <Field label="Weight (% of module)">
                                      <Input
                                        type="number"
                                        min={0}
                                        max={100}
                                        value={String(c.weight ?? "")}
                                        disabled={!!m.autoComponentWeights}
                                        onChange={(e) => updateComponent(idx, cIdx, { weight: clampNumber(Number(e.target.value), 0, 100) })}
                                      />
                                    </Field>

                                    <Field label="Mark (%)">
                                      <Input
                                        type="number"
                                        min={0}
                                        max={100}
                                        value={String(c.mark ?? "")}
                                        onChange={(e) => updateComponent(idx, cIdx, { mark: clampNumber(Number(e.target.value), 0, 100) })}
                                      />
                                    </Field>

                                    <div className="sw-delCell">
                                      {(m.components || []).length > 1 ? (
                                        <Button variant="danger" className="sw-x" onClick={() => removeComponent(idx, cIdx)} title="Remove assessment">
                                          ✕
                                        </Button>
                                      ) : null}
                                    </div>
                                  </div>
                                </div>
                              </details>
                            ))}
                          </div>

                          {state.modules.length > 1 ? (
                            <div style={{ marginTop: 14 }}>
                              <Button variant="danger" onClick={() => removeModule(idx)}>
                                Remove module
                              </Button>
                            </div>
                          ) : null}
                        </div>
                      </details>
                    );
                  })}
                </div>
              </div>
            ) : null}

            {activeTab === "design" ? (
              <div className="sw-section">
                <div className="sw-sectionHead">
                  <div>
                    <div className="sw-sectionTitle">Design</div>
                  </div>
                  <Button variant="ghost" onClick={() => setTheme(DEFAULTS.theme)}>
                    Reset colour scheme
                  </Button>
                </div>

                <div className="sw-form-grid">
                  <Field label="Template">
                    <Select value={state.template} onChange={(e) => set({ template: e.target.value })}>
                      <option value="classic">Classic</option>
                      <option value="gradient">Gradient</option>
                    </Select>
                  </Field>

                  <Field label="Export format">
                    <Select value={state.format} onChange={(e) => set({ format: e.target.value })}>
                      <option value="linkedin_square">LinkedIn square (1200×1200)</option>
                      <option value="instagram_story">Instagram story (1080×1920)</option>
                    </Select>
                  </Field>

                  <div className="sw-inlineCard">
                    <div className="sw-inlineCard__top">
                      <div className="sw-inlineCard__title">Visibility</div>
                    </div>
                    <label className="sw-toggle" style={{ marginTop: 10 }}>
                      <input type="checkbox" checked={!!state.showModuleMarks} onChange={(e) => set({ showModuleMarks: e.target.checked })} />
                      <span>Show module marks on the card</span>
                    </label>
                    <label className="sw-toggle" style={{ marginTop: 8 }}>
                      <input
                        type="checkbox"
                        checked={!!state.showAssessmentsInBreakdown}
                        onChange={(e) => set({ showAssessmentsInBreakdown: e.target.checked })}
                      />
                      <span>Show assessments in module breakdown</span>
                    </label>
                    <div className="sw-inlineCard__sub" style={{ marginTop: 8 }}>
                      Turn off marks and/or assessments to simplify what appears on the card.
                    </div>
                  </div>

                  <Field label="Primary">
                    <div className="sw-colorRow">
                      <input type="color" value={state.theme.primary} onChange={(e) => setTheme({ primary: e.target.value })} />
                      <Input value={state.theme.primary} onChange={(e) => setTheme({ primary: e.target.value })} />
                    </div>
                  </Field>

                  <Field label="Accent">
                    <div className="sw-colorRow">
                      <input type="color" value={state.theme.accent} onChange={(e) => setTheme({ accent: e.target.value })} />
                      <Input value={state.theme.accent} onChange={(e) => setTheme({ accent: e.target.value })} />
                    </div>
                  </Field>

                  <Field label="Background">
                    <div className="sw-colorRow">
                      <input type="color" value={state.theme.bg} onChange={(e) => setTheme({ bg: e.target.value })} />
                      <Input value={state.theme.bg} onChange={(e) => setTheme({ bg: e.target.value })} />
                    </div>
                  </Field>

                  <Field label="Card base">
                    <div className="sw-colorRow">
                      <input type="color" value={state.theme.card} onChange={(e) => setTheme({ card: e.target.value })} />
                      <Input value={state.theme.card} onChange={(e) => setTheme({ card: e.target.value })} />
                    </div>
                  </Field>
                </div>
              </div>
            ) : null}
          </div>
        </div>

        <div className="sw-panel sw-panel--preview">
          <div className="sw-preview-pad" ref={previewBoxRef}>
            <ExportFrame format={state.format} containerRef={previewBoxRef}>
              <div ref={previewCardRef} style={{ width: dims.w, height: dims.h }}>
                <WrappedCard state={state} />
              </div>
            </ExportFrame>
          </div>
        </div>
      </div>
    </div>
  );
}

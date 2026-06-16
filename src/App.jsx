import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import html2canvas from "html2canvas";

const makeId = (prefix) => `${prefix}_${Math.random().toString(36).slice(2, 9)}`;

function makeComponent(patch = {}) {
  return { id: makeId("c"), type: "Overall", name: "Overall", weight: 100, mark: 0, ...patch };
}

function makeModule(patch = {}) {
  return {
    id: makeId("m"),
    code: "",
    title: "",
    credits: 0,
    weight: 100,
    autoComponentWeights: true,
    components: [makeComponent()],
    ...patch,
  };
}

function makeSemester(patch = {}) {
  return { id: makeId("s"), label: "Semester 1", modules: [makeModule()], ...patch };
}

function makeYear(patch = {}) {
  return {
    id: makeId("y"),
    name: "Year 1",
    calendar: "2025/26",
    weight: 100,
    autoModuleWeights: true,
    semesters: [makeSemester()],
    ...patch,
  };
}

const DEFAULTS = {
  personName: "Your Name",
  university: "Your University",
  course: "BSc Biology",
  subtitle: "Class of 2026",
  headline: "My Semester Wrapped",
  caption: "Proud of my results this semester 🎓",
  linkedinHandle: "",

  showModuleMarks: true,
  showAssessmentsInBreakdown: true,

  years: [
    makeYear({
      name: "Year 1",
      calendar: "2024/25",
      weight: 0,
      autoModuleWeights: true,
      semesters: [
        makeSemester({
          label: "Semester 1",
          modules: [
            makeModule({
              code: "CS101",
              title: "Intro to Programming",
              weight: 100,
              autoComponentWeights: true,
              components: [
                makeComponent({ type: "Coursework", name: "Coursework 1", weight: 40, mark: 0 }),
                makeComponent({ type: "Exam", name: "Exam", weight: 60, mark: 0 }),
              ],
            }),
          ],
        }),
      ],
    }),
    makeYear({
      name: "Year 2",
      calendar: "2025/26",
      weight: 40,
      autoModuleWeights: true,
      semesters: [
        makeSemester({
          label: "Semester 1",
          modules: [
            makeModule({
              code: "CS201",
              title: "Algorithms",
              weight: 100,
              autoComponentWeights: true,
              components: [
                makeComponent({ type: "Coursework", name: "Coursework 1", weight: 40, mark: 0 }),
                makeComponent({ type: "Exam", name: "Exam", weight: 60, mark: 0 }),
              ],
            }),
          ],
        }),
      ],
    }),
    makeYear({
      name: "Year 3",
      calendar: "2026/27",
      weight: 60,
      autoModuleWeights: true,
      semesters: [
        makeSemester({
          label: "Semester 1",
          modules: [
            makeModule({
              code: "CS301",
              title: "Machine Learning",
              weight: 100,
              autoComponentWeights: true,
              components: [
                makeComponent({ type: "Coursework", name: "Coursework 1", weight: 40, mark: 0 }),
                makeComponent({ type: "Exam", name: "Exam", weight: 60, mark: 0 }),
              ],
            }),
          ],
        }),
      ],
    }),
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
  if (n >= 70) return { label: "First Class", short: "First" };
  if (n >= 60) return { label: "Upper Second (2:1)", short: "2:1" };
  if (n >= 50) return { label: "Lower Second (2:2)", short: "2:2" };
  if (n >= 40) return { label: "Third Class", short: "Third" };
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

// Flatten every module that belongs to a year (across all its semesters).
function yearModules(year) {
  return (year?.semesters || []).flatMap((s) => s?.modules || []);
}

// Average over a flat list of modules, normalised by module weight.
function computeModuleSet(modules) {
  const ms = (modules || []).filter((m) => m && (m.code || m.title || (m.components || []).length));
  if (ms.length === 0) {
    return { avg: NaN, modules: [], best: null, count: 0 };
  }

  const withMarks = ms.map((m) => {
    const res = computeModuleMark(m);
    return { ...m, moduleMark: res.mark };
  });

  const norm = normalizeWeights(withMarks, "weight");
  const avg = norm.reduce((acc, m) => {
    const mk = Number(m.moduleMark);
    const safeMk = Number.isFinite(mk) ? mk : 0;
    return acc + safeMk * ((m.__norm || 0) / 100);
  }, 0);

  return { avg, modules: withMarks, count: withMarks.length };
}

function computeYear(year) {
  return computeModuleSet(yearModules(year));
}

// Combine year averages by their degree weighting into one overall mark.
function computeDegree(years) {
  const ys = (years || []).map((y) => {
    const r = computeYear(y);
    return { ...y, yearAvg: r.avg, yearCount: r.count };
  });

  const contributing = ys.filter((y) => y.yearCount > 0);
  const norm = normalizeWeights(contributing, "weight");
  const normById = new Map(norm.map((y) => [y.id, y.__norm || 0]));
  const totalWeight = contributing.reduce((a, y) => a + (Number(y.weight) || 0), 0);

  const avg = contributing.length
    ? contributing.reduce((acc, y) => {
        const mk = Number.isFinite(y.yearAvg) ? y.yearAvg : 0;
        const share = totalWeight > 0 ? (normById.get(y.id) || 0) / 100 : 1 / contributing.length;
        return acc + mk * share;
      }, 0)
    : NaN;

  return { avg, years: ys, count: contributing.length, totalWeight };
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

    // eslint-disable-next-line no-control-regex -- strip Windows-illegal + control characters
    .replace(/[<>:"/\\|?* -]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  const safe = cleaned && cleaned.replace(/\./g, "").length ? cleaned : "semester-wrapped";
  return safe.slice(0, 80);
}

// Year-weighting presets. Each returns an array of weights for `n` years.
const YEAR_PRESETS = [
  { id: "equal", label: "Equal", build: (n) => evenSplit(n) },
  {
    id: "final",
    label: "Final year only",
    build: (n) => Array.from({ length: n }, (_, i) => (i === n - 1 ? 100 : 0)),
  },
  {
    id: "uk",
    label: "Standard UK",
    build: (n) => {
      if (n <= 1) return [100];
      if (n === 2) return [40, 60];
      return Array.from({ length: n }, (_, i) => {
        if (i === n - 1) return 60;
        if (i === n - 2) return 40;
        return 0;
      });
    },
  },
  {
    id: "meng",
    label: "Integrated (last 3)",
    build: (n) => {
      if (n < 3) return evenSplit(n);
      return Array.from({ length: n }, (_, i) => {
        if (i === n - 1) return 40;
        if (i === n - 2) return 40;
        if (i === n - 3) return 20;
        return 0;
      });
    },
  },
];

/* ============================================================
   Small UI atoms
   ============================================================ */
function Field({ label, hint, children }) {
  return (
    <label className="sw-field">
      <div className="sw-fieldLabel">{label}</div>
      {children}
      {hint ? <div className="sw-fieldHint">{hint}</div> : null}
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
  return <button type="button" {...props} className={`sw-btn sw-btn--${variant} ${props.className || ""}`} />;
}

function Toggle({ checked, onChange, label }) {
  return (
    <label className="sw-toggle">
      <input type="checkbox" checked={!!checked} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  );
}

// Flat, tappable list row — the building block that replaces nested accordions.
function ListRow({ title, sub, value, valueTone, onClick, chevron = true }) {
  return (
    <button type="button" className="sw-row" onClick={onClick}>
      <span className="sw-row__main">
        <span className="sw-row__title">{title}</span>
        {sub ? <span className="sw-row__sub">{sub}</span> : null}
      </span>
      {value != null ? <span className={`sw-row__value ${valueTone ? `sw-row__value--${valueTone}` : ""}`}>{value}</span> : null}
      {chevron ? <span className="sw-row__chev" aria-hidden="true">›</span> : null}
    </button>
  );
}

/* ============================================================
   Share view model + card rendering
   ============================================================ */

// Build the scope-specific content (hero + ordered blocks) for the card.
function buildShareView(state, share) {
  const years = state.years || [];
  const scope = share.scope;

  if (scope === "year" || scope === "semester") {
    const year = years.find((y) => y.id === share.yearId) || years[0];
    if (scope === "semester") {
      const sem = (year?.semesters || []).find((s) => s.id === share.semesterId) || (year?.semesters || [])[0];
      const set = computeModuleSet(sem?.modules || []);
      const ranked = (set.modules || [])
        .slice()
        .sort((a, b) => (Number(b.moduleMark) || 0) - (Number(a.moduleMark) || 0));
      return {
        scope,
        title: "Module breakdown",
        sub: `${safeText(year?.name || "", 18)}${sem?.label ? ` • ${safeText(sem.label, 18)}` : ""}`,
        hero: {
          pct: set.count ? formatPct(set.avg) : "—",
          classLabel: set.count ? getUkClassification(set.avg).label : "Add modules to calculate",
          context: "Semester average",
        },
        blocks: ranked.map((m, i) => moduleBlock(m, i, sem?.id, state.showAssessmentsInBreakdown, state.showModuleMarks)),
        semesterHeaders: {},
      };
    }

    // Year scope: grouped by semester
    const set = computeYear(year || {});
    const blocks = [];
    const semesterHeaders = {};
    let rank = 0;
    (year?.semesters || []).forEach((sem) => {
      const semSet = computeModuleSet(sem.modules || []);
      const header = {
        type: "semHead",
        key: `h_${sem.id}`,
        semesterId: sem.id,
        label: sem.label,
        avg: semSet.count ? formatPct(semSet.avg) : "—",
      };
      semesterHeaders[sem.id] = header;
      blocks.push(header);
      const ranked = (semSet.modules || [])
        .slice()
        .sort((a, b) => (Number(b.moduleMark) || 0) - (Number(a.moduleMark) || 0));
      ranked.forEach((m) => {
        blocks.push(moduleBlock(m, rank++, sem.id, false, state.showModuleMarks));
      });
    });
    return {
      scope,
      title: "Module breakdown",
      sub: `${safeText(year?.name || "Year", 18)}${year?.calendar ? ` • ${safeText(year.calendar, 12)}` : ""}`,
      hero: {
        pct: set.count ? formatPct(set.avg) : "—",
        classLabel: set.count ? getUkClassification(set.avg).label : "Add modules to calculate",
        context: "Year average",
      },
      blocks,
      semesterHeaders,
    };
  }

  // Course scope (default)
  const degree = computeDegree(years);
  const blocks = degree.years.map((y, i) => ({
    type: "yearRow",
    key: `y_${y.id}`,
    rank: i + 1,
    name: y.name || `Year ${i + 1}`,
    calendar: y.calendar,
    weight: Math.round(Number(y.weight) || 0),
    avg: y.yearCount ? formatPct(y.yearAvg) : "—",
    classLabel: y.yearCount ? getUkClassification(y.yearAvg).label : "No marks yet",
    showMark: state.showModuleMarks,
  }));
  return {
    scope,
    title: "Year breakdown",
    sub: "Weighted toward your degree",
    hero: {
      pct: degree.count ? formatPct(degree.avg) : "—",
      classLabel: degree.count ? getUkClassification(degree.avg).label : "Add modules to calculate",
      context: "Overall degree",
    },
    blocks,
    semesterHeaders: {},
  };
}

function moduleBlock(m, rank, semesterId, showAssess, showMark) {
  const active = (m.components || []).filter((c) => c && (c.name || c.type));
  const limit = 3;
  return {
    type: "moduleRow",
    key: `m_${m.id}`,
    semesterId,
    rank: rank + 1,
    code: m.code,
    title: m.title,
    weight: Math.round(Number(m.weight) || 0),
    mark: formatPct(Number(m.moduleMark)),
    showMark,
    assess: showAssess
      ? active.slice(0, limit).map((c) => ({
          name: c.name || c.type,
          weight: Number(c.weight) ? `${Math.round(Number(c.weight))}%` : "",
          mark: showMark ? formatPct(Number(c.mark)) : "",
        }))
      : [],
    moreCount: showAssess ? Math.max(0, active.length - limit) : 0,
  };
}

const CARD_LAYOUT = {
  linkedin_square: {
    innerPad: 54,
    innerGap: 16,
    headline: 58,
    name: 32,
    handle: 20,
    meta: 26,
    pct: 116,
    cls: 26,
    context: 19,
    sectionTitle: 28,
    sectionSub: 20,
    rowPadY: 16,
    rowPadX: 18,
    rank: 18,
    code: 22,
    title: 28,
    weight: 22,
    mark: 38,
    assess: 19,
    semHead: 24,
    footer: 30,
  },
  instagram_story: {
    innerPad: 64,
    innerGap: 22,
    headline: 76,
    name: 42,
    handle: 26,
    meta: 32,
    pct: 150,
    cls: 32,
    context: 24,
    sectionTitle: 36,
    sectionSub: 26,
    rowPadY: 20,
    rowPadX: 22,
    rank: 22,
    code: 26,
    title: 34,
    weight: 26,
    mark: 46,
    assess: 23,
    semHead: 30,
    footer: 38,
  },
};

function CardBlock({ block, layout, theme }) {
  if (block.type === "semHead") {
    return (
      <div className="sw-cardSemHead" style={{ borderColor: hexAlpha(theme.primary, 0.6) }}>
        <span className="sw-cardSemHead__label" style={{ fontSize: layout.semHead }}>
          {safeText(block.label, 26)}
          {block.cont ? " (cont.)" : ""}
        </span>
        <span className="sw-cardSemHead__avg" style={{ fontSize: layout.semHead, color: theme.accent }}>
          {block.avg}
        </span>
      </div>
    );
  }

  if (block.type === "yearRow") {
    return (
      <div
        className="sw-card__row"
        style={{ background: hexAlpha(theme.card, 0.55), border: `1px solid ${hexAlpha(theme.text, 0.08)}`, padding: `${layout.rowPadY}px ${layout.rowPadX}px` }}
      >
        <div className="sw-card__rank" style={{ fontSize: layout.rank }}>
          {String(block.rank).padStart(2, "0")}
        </div>
        <div className="sw-card__rowMain">
          <div className="sw-card__code" style={{ fontSize: layout.code }}>
            {safeText(block.calendar || "", 12)}
          </div>
          <div className="sw-card__title" style={{ fontSize: layout.title }}>
            {safeText(block.name, 28)}
          </div>
          <div className="sw-card__rowMeta" style={{ fontSize: layout.assess }}>
            {block.classLabel}
          </div>
        </div>
        <div className="sw-card__rowRight">
          <div className="sw-card__weight" style={{ fontSize: layout.weight }}>
            {block.weight}%
          </div>
          {block.showMark ? (
            <div className="sw-card__mark" style={{ fontSize: layout.mark, color: theme.accent }}>
              {block.avg}
            </div>
          ) : null}
        </div>
      </div>
    );
  }

  // moduleRow
  return (
    <div
      className="sw-card__row"
      style={{ background: hexAlpha(theme.card, 0.55), border: `1px solid ${hexAlpha(theme.text, 0.08)}`, padding: `${layout.rowPadY}px ${layout.rowPadX}px` }}
    >
      <div className="sw-card__rank" style={{ fontSize: layout.rank }}>
        {String(block.rank).padStart(2, "0")}
      </div>
      <div className="sw-card__rowMain">
        {block.code ? (
          <div className="sw-card__code" style={{ fontSize: layout.code }}>
            {safeText(block.code, 16)}
          </div>
        ) : null}
        <div className="sw-card__title" style={{ fontSize: layout.title }}>
          {safeText(block.title || "Untitled module", 44)}
        </div>
        {block.assess && block.assess.length ? (
          <div className="sw-card__assessList" style={{ fontSize: layout.assess }}>
            {block.assess.map((a, i) => (
              <div key={i} className="sw-card__assessItem">
                <span className="sw-card__assessName">{safeText(a.name, 26)}</span>
                <span className="sw-card__assessMeta">
                  {a.weight}
                  {a.weight && a.mark ? " • " : ""}
                  {a.mark}
                </span>
              </div>
            ))}
            {block.moreCount ? <div className="sw-card__assessMore">+{block.moreCount} more</div> : null}
          </div>
        ) : null}
      </div>
      <div className="sw-card__rowRight">
        {block.weight ? (
          <div className="sw-card__weight" style={{ fontSize: layout.weight }}>
            {block.weight}%
          </div>
        ) : null}
        {block.showMark ? (
          <div className="sw-card__mark" style={{ fontSize: layout.mark, color: theme.accent }}>
            {block.mark}
          </div>
        ) : null}
      </div>
    </div>
  );
}

// Renders one card (one page). `blocks` is the page slice; full-view when measuring.
function WrappedCard({ state, view, blocks, pageLabel }) {
  const { theme, headline, caption, personName, university, course, subtitle, format, template, linkedinHandle } = state;
  const layout = CARD_LAYOUT[format] || CARD_LAYOUT.linkedin_square;
  const handle = normalizeHandle(linkedinHandle);
  const bg = template === "gradient" ? gradientBg(theme) : theme.bg;

  return (
    <div className="sw-card" style={{ background: bg, color: theme.text }}>
      <div className="sw-card__grain" />
      <div className="sw-card__inner" style={{ padding: layout.innerPad, gap: layout.innerGap }}>
        <div className="sw-card__hero">
          <div className="sw-card__headline" style={{ fontSize: layout.headline }}>
            {safeText(headline, 42)}
          </div>
          <div className="sw-card__nameRow">
            <div className="sw-card__name" style={{ fontSize: layout.name }}>
              {safeText(personName, 34)}
            </div>
            {handle ? (
              <div className="sw-card__handle" style={{ fontSize: layout.handle }}>
                LinkedIn - {handle}
              </div>
            ) : null}
          </div>
          <div className="sw-card__meta" style={{ fontSize: layout.meta }}>
            <div className="sw-card__metaLine">
              <span className="sw-card__metaStrong">{safeText(university, 44)}</span>
              <span className="sw-card__dot">•</span>
              <span className="sw-card__metaMuted">{safeText(course, 48)}</span>
            </div>
            {subtitle ? (
              <div className="sw-card__metaLine">
                <span className="sw-card__metaMuted">{safeText(subtitle, 40)}</span>
              </div>
            ) : null}
          </div>

          <div className="sw-card__gradeRow">
            <div className="sw-card__pct" style={{ fontSize: layout.pct }}>
              {view.hero.pct}
            </div>
            <div className="sw-card__class" style={{ fontSize: layout.cls }}>
              <span className="sw-card__context" style={{ fontSize: layout.context, color: theme.accent }}>
                {view.hero.context}
              </span>
              {" • "}
              {view.hero.classLabel}
            </div>
          </div>
        </div>

        <div className="sw-card__breakdown">
          <div className="sw-card__sectionHead">
            <div className="sw-card__sectionTitle" style={{ fontSize: layout.sectionTitle }}>
              {view.title}
              {pageLabel ? <span className="sw-card__pageTag" style={{ fontSize: layout.sectionSub }}>{pageLabel}</span> : null}
            </div>
            <div className="sw-card__sectionSub" style={{ fontSize: layout.sectionSub }}>
              {view.sub}
            </div>
          </div>

          <div className="sw-card__list">
            {blocks.length === 0 ? (
              <div className="sw-card__empty" style={{ fontSize: layout.context }}>
                Add modules and assessments to populate your wrap.
              </div>
            ) : (
              blocks.map((b) => (
                <div className="sw-blk" key={b.key}>
                  <CardBlock block={b} layout={layout} theme={theme} />
                </div>
              ))
            )}
          </div>
        </div>

        <div className="sw-card__footer" style={{ fontSize: layout.footer }}>
          <div className="sw-card__footerLeft">{caption ? safeText(caption, 84) : ""}</div>
          <div className="sw-card__footerRight" style={{ color: theme.primary }}>
            #SemesterWrapped
          </div>
        </div>
      </div>
    </div>
  );
}

// Measure the full block list inside a hidden full-size card and split into pages that fit.
function paginateFromMeasure(measRoot, view) {
  const listEl = measRoot.querySelector(".sw-card__list");
  const breakdownEl = measRoot.querySelector(".sw-card__breakdown");
  const headEl = measRoot.querySelector(".sw-card__sectionHead");
  if (!listEl || !breakdownEl) return [view.blocks];

  const items = Array.from(listEl.children);
  if (items.length === 0) return [[]];

  const cs = getComputedStyle(listEl);
  const gap = parseFloat(cs.rowGap) || 12;
  const breakdownGap = parseFloat(getComputedStyle(breakdownEl).rowGap) || 0;
  const headH = headEl ? headEl.offsetHeight : 0;
  let avail = breakdownEl.clientHeight - headH - breakdownGap;

  // Reserve room for a possible continuation header on year-scope pages.
  let semHeadReserve = 0;
  if (view.scope === "year") {
    const firstHead = measRoot.querySelector(".sw-cardSemHead");
    semHeadReserve = (firstHead ? firstHead.offsetHeight : 0) + gap;
  }
  if (!(avail > 50)) return [view.blocks];

  const heights = items.map((el) => el.offsetHeight);

  const pages = [];
  let cur = [];
  let curH = 0;
  for (let i = 0; i < view.blocks.length; i++) {
    const h = heights[i] || 0;
    const budget = pages.length === 0 ? avail : avail - semHeadReserve;
    const add = (cur.length ? gap : 0) + h;
    if (cur.length && curH + add > budget) {
      pages.push(cur);
      cur = [];
      curH = 0;
    }
    cur.push(view.blocks[i]);
    curH += (cur.length > 1 ? gap : 0) + h;
  }
  if (cur.length) pages.push(cur);

  // Fix orphan semester headers (header as last block on a page → move to next page).
  for (let p = 0; p < pages.length - 1; p++) {
    const page = pages[p];
    while (page.length > 1 && page[page.length - 1].type === "semHead") {
      const moved = page.pop();
      pages[p + 1].unshift(moved);
    }
  }

  // Add continuation headers for pages (after the first) that start mid-semester.
  if (view.scope === "year") {
    for (let p = 1; p < pages.length; p++) {
      const first = pages[p][0];
      if (first && first.type === "moduleRow") {
        const header = view.semesterHeaders[first.semesterId];
        if (header) pages[p].unshift({ ...header, key: `${header.key}_cont_${p}`, cont: true });
      }
    }
  }

  return pages.length ? pages : [[]];
}

function SizeHint({ format }) {
  const f = getDims(format);
  return (
    <span className="sw-sizehint">
      <span className="sw-sizehint__pill">{f.w}×{f.h}</span> {f.label}
    </span>
  );
}

function ExportFrame({ format, containerRef, children }) {
  const { w, h } = getDims(format);
  const [box, setBox] = useState({ width: 520, height: 520 });

  useEffect(() => {
    const el = containerRef?.current;
    if (!el) return;
    if (typeof ResizeObserver === "undefined") {
      const update = () => {
        const r = el.getBoundingClientRect?.();
        if (r) setBox({ width: r.width, height: r.height });
      };
      update();
      window.addEventListener("resize", update);
      return () => window.removeEventListener("resize", update);
    }
    const ro = new ResizeObserver((entries) => {
      const cr = entries[0]?.contentRect;
      if (cr) setBox({ width: cr.width, height: cr.height });
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [containerRef]);

  const pad = 8;
  const availW = Math.max(160, box.width - pad * 2);
  const availH = Math.max(160, box.height - pad * 2);
  const scale = Math.min(1, availW / w, availH / h);

  return (
    <div className="sw-stage" style={{ width: Math.round(w * scale), height: Math.round(h * scale) }}>
      <div style={{ width: w, height: h, transform: `scale(${scale})`, transformOrigin: "top left" }}>{children}</div>
    </div>
  );
}

/* ============================================================
   Main app
   ============================================================ */
export default function SemesterWrappedApp() {
  const [state, setState] = useState(DEFAULTS);
  const [nav, setNav] = useState({ level: "course", yearId: null, semesterId: null, moduleId: null });
  const [editingCompId, setEditingCompId] = useState(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [share, setShare] = useState({ scope: "course", yearId: null, semesterId: null });
  const [pageIdx, setPageIdx] = useState(0);
  const [fileName, setFileName] = useState("semester-wrapped");
  const [captionCopied, setCaptionCopied] = useState(false);
  const [exporting, setExporting] = useState(false);

  const previewBoxRef = useRef(null);
  const previewCardRef = useRef(null);
  const measRef = useRef(null);
  const touchX = useRef(null);

  const dims = getDims(state.format);
  const set = (patch) => setState((s) => ({ ...s, ...patch }));
  const setTheme = (patch) => setState((s) => ({ ...s, theme: { ...s.theme, ...patch } }));

  // ---- Auto module weights (per year) ----
  const autoWeightSig = state.years.map((y) => (y.autoModuleWeights ? yearModules(y).length : -1)).join("|");
  useEffect(() => {
    setState((s) => {
      let changed = false;
      const years = s.years.map((y) => {
        if (!y.autoModuleWeights) return y;
        const weights = evenSplit(yearModules(y).length);
        let k = 0;
        const semesters = y.semesters.map((sem) => ({
          ...sem,
          modules: (sem.modules || []).map((m) => {
            const w = weights[k++] ?? 0;
            if (m.weight !== w) changed = true;
            return m.weight === w ? m : { ...m, weight: w };
          }),
        }));
        return { ...y, semesters };
      });
      return changed ? { ...s, years } : s;
    });
  }, [autoWeightSig]);

  // ---- id-addressed mutations ----
  const editYears = (fn) => setState((s) => ({ ...s, years: fn(s.years) }));

  const updateYear = (yid, patch) =>
    editYears((years) => years.map((y) => (y.id === yid ? { ...y, ...patch } : y)));

  const addYear = () =>
    setState((s) => {
      const y = makeYear({ name: `Year ${s.years.length + 1}`, calendar: "", weight: 0 });
      return { ...s, years: [...s.years, y] };
    });

  const removeYear = (yid) =>
    setState((s) => {
      const next = s.years.filter((y) => y.id !== yid);
      return { ...s, years: next.length ? next : [makeYear()] };
    });

  const applyYearPreset = (preset) =>
    editYears((years) => {
      const w = preset.build(years.length);
      return years.map((y, i) => ({ ...y, weight: w[i] ?? 0 }));
    });

  const mutateYear = (yid, fn) => editYears((years) => years.map((y) => (y.id === yid ? fn(y) : y)));

  const updateSemester = (yid, sid, patch) =>
    mutateYear(yid, (y) => ({ ...y, semesters: y.semesters.map((s) => (s.id === sid ? { ...s, ...patch } : s)) }));

  const addSemester = (yid) =>
    mutateYear(yid, (y) => ({
      ...y,
      semesters: [...y.semesters, makeSemester({ label: `Semester ${y.semesters.length + 1}`, modules: [makeModule()] })],
    }));

  const removeSemester = (yid, sid) =>
    mutateYear(yid, (y) => {
      const next = y.semesters.filter((s) => s.id !== sid);
      return { ...y, semesters: next.length ? next : [makeSemester()] };
    });

  const mutateSemester = (yid, sid, fn) =>
    mutateYear(yid, (y) => ({ ...y, semesters: y.semesters.map((s) => (s.id === sid ? fn(s) : s)) }));

  const addModule = (yid, sid) => mutateSemester(yid, sid, (s) => ({ ...s, modules: [...s.modules, makeModule()] }));

  const removeModule = (yid, sid, mid) =>
    mutateSemester(yid, sid, (s) => {
      const next = s.modules.filter((m) => m.id !== mid);
      return { ...s, modules: next.length ? next : [makeModule()] };
    });

  const mutateModule = (yid, sid, mid, fn) =>
    mutateSemester(yid, sid, (s) => ({ ...s, modules: s.modules.map((m) => (m.id === mid ? fn(m) : m)) }));

  const updateModule = (yid, sid, mid, patch) => mutateModule(yid, sid, mid, (m) => ({ ...m, ...patch }));

  const addComponent = (yid, sid, mid) =>
    mutateModule(yid, sid, mid, (m) => {
      const next = { ...m, components: [...(m.components || []), makeComponent({ type: "Assignment", name: "Assignment", weight: 0 })] };
      if (next.autoComponentWeights) {
        const w = evenSplit(next.components.length);
        next.components = next.components.map((c, i) => ({ ...c, weight: w[i] ?? 0 }));
      }
      return next;
    });

  const updateComponent = (yid, sid, mid, cid, patch) =>
    mutateModule(yid, sid, mid, (m) => ({ ...m, components: (m.components || []).map((c) => (c.id === cid ? { ...c, ...patch } : c)) }));

  const removeComponent = (yid, sid, mid, cid) =>
    mutateModule(yid, sid, mid, (m) => {
      const next = (m.components || []).filter((c) => c.id !== cid);
      let nm = { ...m, components: next.length ? next : [makeComponent()] };
      if (nm.autoComponentWeights) {
        const w = evenSplit(nm.components.length);
        nm.components = nm.components.map((c, i) => ({ ...c, weight: w[i] ?? 0 }));
      }
      return nm;
    });

  const setAutoComponentWeights = (yid, sid, mid, enabled) =>
    mutateModule(yid, sid, mid, (m) => {
      let nm = { ...m, autoComponentWeights: enabled };
      if (enabled) {
        const w = evenSplit((nm.components || []).length);
        nm.components = (nm.components || []).map((c, i) => ({ ...c, weight: w[i] ?? 0 }));
      }
      return nm;
    });

  // ---- current entities for the editor ----
  const curYear = state.years.find((y) => y.id === nav.yearId) || null;
  const curSem = curYear?.semesters.find((s) => s.id === nav.semesterId) || null;
  const curMod = curSem?.modules.find((m) => m.id === nav.moduleId) || null;

  // ---- share view + pagination ----
  const view = useMemo(() => buildShareView(state, share), [state, share]);
  const [pages, setPages] = useState([view.blocks]);

  useLayoutEffect(() => {
    if (!measRef.current) {
      setPages([view.blocks]);
      return;
    }
    const compute = () => {
      if (measRef.current) setPages(paginateFromMeasure(measRef.current, view));
    };
    compute();
    let cancelled = false;
    if (document.fonts?.ready) document.fonts.ready.then(() => !cancelled && compute());
    return () => {
      cancelled = true;
    };
  }, [view, state.format]);

  useEffect(() => {
    setPageIdx((p) => Math.min(p, Math.max(0, pages.length - 1)));
  }, [pages.length]);

  const safePage = Math.min(pageIdx, Math.max(0, pages.length - 1));
  const curBlocks = pages[safePage] || [];
  const pageLabel = pages.length > 1 ? `${safePage + 1} / ${pages.length}` : "";

  // ---- export ----
  async function renderCardToBlob() {
    const sourceCard = previewCardRef.current?.querySelector?.(".sw-card");
    if (!sourceCard) throw new Error("preview not ready");
    const { w, h } = getDims(state.format);
    if (document.fonts?.ready) await document.fonts.ready;
    await new Promise((r) => requestAnimationFrame(r));

    const tempRoot = document.createElement("div");
    Object.assign(tempRoot.style, { position: "fixed", left: "-20000px", top: "0", width: `${w}px`, height: `${h}px`, zIndex: "-1" });
    document.body.appendChild(tempRoot);
    const clone = sourceCard.cloneNode(true);
    Object.assign(clone.style, { width: `${w}px`, height: `${h}px`, margin: "0", transform: "none" });
    tempRoot.appendChild(clone);
    try {
      const canvas = await html2canvas(clone, { backgroundColor: null, scale: 2, useCORS: false, logging: false, width: w, height: h, windowWidth: w, windowHeight: h });
      return await new Promise((res, rej) => canvas.toBlob((b) => (b ? res(b) : rej(new Error("toBlob null"))), "image/png"));
    } finally {
      tempRoot.remove();
    }
  }

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  async function exportCurrent() {
    setExporting(true);
    document.body.classList.add("sw-exporting");
    try {
      const blob = await renderCardToBlob();
      const suffix = pages.length > 1 ? `-${safePage + 1}` : "";
      downloadBlob(blob, `${sanitizeFileName(fileName)}${suffix}.png`);
    } catch (e) {
      console.error("Export failed:", e);
      alert("Export failed. Check console.");
    } finally {
      document.body.classList.remove("sw-exporting");
      setExporting(false);
    }
  }

  async function exportAll() {
    if (pages.length <= 1) return exportCurrent();
    setExporting(true);
    document.body.classList.add("sw-exporting");
    try {
      for (let i = 0; i < pages.length; i++) {
        setPageIdx(i);
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
        const blob = await renderCardToBlob();
        downloadBlob(blob, `${sanitizeFileName(fileName)}-${i + 1}.png`);
      }
    } catch (e) {
      console.error("Export all failed:", e);
      alert("Export failed. Check console.");
    } finally {
      document.body.classList.remove("sw-exporting");
      setExporting(false);
    }
  }

  function buildCaptionText() {
    const degree = computeDegree(state.years);
    const cls = getUkClassification(degree.avg);
    const li = state.linkedinHandle ? `LinkedIn: ${normalizeHandle(state.linkedinHandle)}` : "";
    const yearLines = degree.years
      .map((y) => {
        const yc = getUkClassification(y.yearAvg);
        return `• ${y.name}${y.calendar ? ` (${y.calendar})` : ""}: ${y.yearCount ? `${formatPct(y.yearAvg)} ${yc.short}` : "–"} — ${Math.round(Number(y.weight) || 0)}%`;
      })
      .join("\n");
    return (
      `🎓 ${state.headline}\n${state.personName}\n${state.university} • ${state.course}\n` +
      (state.subtitle ? `${state.subtitle}\n` : "") +
      `\nOverall: ${degree.count ? `${formatPct(degree.avg)} • ${cls.label}` : "–"}\n\n${yearLines}\n` +
      (li ? `\n${li}\n` : "\n") +
      `#SemesterWrapped #university #students`
    );
  }

  async function copyCaption() {
    try {
      await navigator.clipboard?.writeText(buildCaptionText());
      setCaptionCopied(true);
      setTimeout(() => setCaptionCopied(false), 2000);
    } catch {
      alert("Couldn't copy caption.");
    }
  }

  // ---- navigation helpers ----
  const goCourse = () => setNav({ level: "course", yearId: null, semesterId: null, moduleId: null });
  const goYear = (yid) => setNav({ level: "year", yearId: yid, semesterId: null, moduleId: null });
  const goSemester = (yid, sid) => setNav({ level: "semester", yearId: yid, semesterId: sid, moduleId: null });
  const goModule = (yid, sid, mid) => setNav({ level: "module", yearId: yid, semesterId: sid, moduleId: mid });

  const degreeNow = useMemo(() => computeDegree(state.years), [state.years]);

  // Keep nav valid if the referenced entity disappears.
  useEffect(() => {
    if (nav.level === "year" && !curYear) goCourse();
    else if (nav.level === "semester" && (!curYear || !curSem)) curYear ? goYear(curYear.id) : goCourse();
    else if (nav.level === "module" && (!curYear || !curSem || !curMod)) {
      if (curYear && curSem) goSemester(curYear.id, curSem.id);
      else if (curYear) goYear(curYear.id);
      else goCourse();
    }
  }, [nav.level, curYear, curSem, curMod]);

  const breadcrumb = (
    <div className="sw-crumb">
      <button type="button" className="sw-crumb__item" onClick={goCourse}>Course</button>
      {curYear && nav.level !== "course" ? (
        <>
          <span className="sw-crumb__sep">/</span>
          <button type="button" className="sw-crumb__item" onClick={() => goYear(curYear.id)}>{curYear.name || "Year"}</button>
        </>
      ) : null}
      {curSem && (nav.level === "semester" || nav.level === "module") ? (
        <>
          <span className="sw-crumb__sep">/</span>
          <button type="button" className="sw-crumb__item" onClick={() => goSemester(curYear.id, curSem.id)}>{curSem.label || "Semester"}</button>
        </>
      ) : null}
      {curMod && nav.level === "module" ? (
        <>
          <span className="sw-crumb__sep">/</span>
          <span className="sw-crumb__item is-current">{curMod.code || curMod.title || "Module"}</span>
        </>
      ) : null}
    </div>
  );

  return (
    <div className="sw-app">
      <div className="sw-shell">
        <div className="sw-topbar">
          <div className="sw-title">Semester Wrapped</div>
          <div className="sw-actions">
            <Button variant="ghost" onClick={() => setSettingsOpen(true)}>Settings</Button>
            <Button
              variant="ghost"
              onClick={() => {
                setState(DEFAULTS);
                goCourse();
                setShare({ scope: "course", yearId: null, semesterId: null });
                setFileName("semester-wrapped");
              }}
            >
              Reset
            </Button>
          </div>
        </div>
      </div>

      <div className="sw-workspace">
        {/* ---------------- Preview ---------------- */}
        <div className="sw-panel sw-panel--preview">
          <div className="sw-preview__head">
            <div className="sw-preview__title">Share</div>
            <div className="sw-scope">
              <div className="sw-seg">
                {["course", "year", "semester"].map((sc) => (
                  <button
                    key={sc}
                    type="button"
                    className={`sw-seg__btn ${share.scope === sc ? "is-active" : ""}`}
                    onClick={() => {
                      const first = state.years[0];
                      setShare({
                        scope: sc,
                        yearId: share.yearId || first?.id || null,
                        semesterId: share.semesterId || first?.semesters?.[0]?.id || null,
                      });
                    }}
                  >
                    {sc === "course" ? "Course" : sc === "year" ? "Year" : "Semester"}
                  </button>
                ))}
              </div>
              {share.scope !== "course" ? (
                <Select
                  className="sw-scope__pick"
                  value={share.yearId || ""}
                  onChange={(e) => {
                    const y = state.years.find((yy) => yy.id === e.target.value);
                    setShare((sh) => ({ ...sh, yearId: e.target.value, semesterId: y?.semesters?.[0]?.id || null }));
                  }}
                >
                  {state.years.map((y) => (
                    <option key={y.id} value={y.id}>{y.name || "Year"}</option>
                  ))}
                </Select>
              ) : null}
              {share.scope === "semester" && curShareYear(state, share) ? (
                <Select
                  className="sw-scope__pick"
                  value={share.semesterId || ""}
                  onChange={(e) => setShare((sh) => ({ ...sh, semesterId: e.target.value }))}
                >
                  {curShareYear(state, share).semesters.map((s) => (
                    <option key={s.id} value={s.id}>{s.label || "Semester"}</option>
                  ))}
                </Select>
              ) : null}
            </div>
          </div>

          <div className="sw-preview__stage" ref={previewBoxRef}>
            <ExportFrame format={state.format} containerRef={previewBoxRef}>
              <div ref={previewCardRef} style={{ width: dims.w, height: dims.h }}>
                <WrappedCard state={state} view={view} blocks={curBlocks} pageLabel={pageLabel} />
              </div>
            </ExportFrame>
          </div>

          {pages.length > 1 ? (
            <div
              className="sw-pager"
              onTouchStart={(e) => (touchX.current = e.touches[0].clientX)}
              onTouchEnd={(e) => {
                if (touchX.current == null) return;
                const dx = e.changedTouches[0].clientX - touchX.current;
                if (dx < -40) setPageIdx((p) => Math.min(pages.length - 1, p + 1));
                if (dx > 40) setPageIdx((p) => Math.max(0, p - 1));
                touchX.current = null;
              }}
            >
              <button type="button" className="sw-pager__arrow" onClick={() => setPageIdx((p) => Math.max(0, p - 1))} disabled={safePage === 0}>‹</button>
              <div className="sw-pager__dots">
                {pages.map((_, i) => (
                  <button key={i} type="button" className={`sw-pager__dot ${i === safePage ? "is-active" : ""}`} onClick={() => setPageIdx(i)} aria-label={`Card ${i + 1}`} />
                ))}
              </div>
              <button type="button" className="sw-pager__arrow" onClick={() => setPageIdx((p) => Math.min(pages.length - 1, p + 1))} disabled={safePage === pages.length - 1}>›</button>
            </div>
          ) : null}

          <div className="sw-preview__actions">
            <div className="sw-file">
              <input className="sw-input" value={fileName} onChange={(e) => setFileName(e.target.value)} placeholder="semester-wrapped" />
            </div>
            <Button variant="ghost" onClick={copyCaption}>{captionCopied ? "Copied!" : "Caption"}</Button>
            {pages.length > 1 ? <Button variant="ghost" onClick={exportCurrent} disabled={exporting}>Export this</Button> : null}
            <Button onClick={exportAll} disabled={exporting}>
              {exporting ? "Exporting…" : pages.length > 1 ? `Export all (${pages.length})` : "Export PNG"}
            </Button>
          </div>
          <div className="sw-preview__hint">
            <SizeHint format={state.format} />
          </div>
        </div>

        {/* ---------------- Editor ---------------- */}
        <div className="sw-panel sw-panel--editor">
          <div className="sw-screen">
            <div className="sw-screen__head">
              {nav.level !== "course" ? (
                <button
                  type="button"
                  className="sw-back"
                  onClick={() => {
                    if (nav.level === "module") goSemester(curYear.id, curSem.id);
                    else if (nav.level === "semester") goYear(curYear.id);
                    else goCourse();
                  }}
                >
                  ‹ Back
                </button>
              ) : (
                <div className="sw-screen__title">Your course</div>
              )}
              {breadcrumb}
            </div>

            <div className="sw-screen__body">
              {/* -------- Course -------- */}
              {nav.level === "course" ? (
                <>
                  <div className="sw-result">
                    <div className="sw-result__label">Overall degree</div>
                    <div className="sw-result__value">
                      {degreeNow.count ? `${formatPct(degreeNow.avg)} • ${getUkClassification(degreeNow.avg).label}` : "Add modules to calculate"}
                    </div>
                  </div>

                  <div className="sw-block">
                    <div className="sw-block__head">
                      <div className="sw-block__title">Degree weighting</div>
                    </div>
                    <div className="sw-block__sub">How much each year counts toward your final classification.</div>
                    <div className="sw-presetRow">
                      {YEAR_PRESETS.map((p) => (
                        <Button key={p.id} variant="ghost" className="sw-btn--sm" onClick={() => applyYearPreset(p)}>{p.label}</Button>
                      ))}
                    </div>
                    <div className="sw-weightTable">
                      {state.years.map((y) => {
                        const r = computeYear(y);
                        return (
                          <div className="sw-weightRow" key={y.id}>
                            <div className="sw-weightRow__name">
                              <span className="sw-weightRow__title">{y.name || "Year"}</span>
                              <span className="sw-weightRow__meta">{r.count ? `${formatPct(r.avg)} • ${getUkClassification(r.avg).short}` : "No marks yet"}</span>
                            </div>
                            <div className="sw-weightRow__input">
                              <Input type="number" min={0} max={100} value={String(y.weight ?? "")} onChange={(e) => updateYear(y.id, { weight: clampNumber(Number(e.target.value), 0, 100) })} />
                              <span className="sw-weightRow__pct">%</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div className="sw-listHead">
                    <div className="sw-block__title">Years</div>
                  </div>
                  <div className="sw-list">
                    {state.years.map((y) => {
                      const r = computeYear(y);
                      return (
                        <ListRow
                          key={y.id}
                          title={y.name || "Year"}
                          sub={`${y.calendar ? `${y.calendar} • ` : ""}${y.semesters.length} sem • ${Math.round(Number(y.weight) || 0)}% of degree`}
                          value={r.count ? formatPct(r.avg) : "—"}
                          valueTone="accent"
                          onClick={() => goYear(y.id)}
                        />
                      );
                    })}
                  </div>
                  <Button variant="ghost" className="sw-addRow" onClick={addYear}>+ Add year</Button>
                </>
              ) : null}

              {/* -------- Year -------- */}
              {nav.level === "year" && curYear ? (
                <>
                  <div className="sw-fieldStack">
                    <Field label="Year name">
                      <Input value={curYear.name} onChange={(e) => updateYear(curYear.id, { name: e.target.value })} placeholder="Year 1" />
                    </Field>
                    <div className="sw-grid2">
                      <Field label="Academic year">
                        <Input value={curYear.calendar} onChange={(e) => updateYear(curYear.id, { calendar: e.target.value })} placeholder="2025/26" />
                      </Field>
                      <Field label="Weight (% of degree)">
                        <Input type="number" min={0} max={100} value={String(curYear.weight ?? "")} onChange={(e) => updateYear(curYear.id, { weight: clampNumber(Number(e.target.value), 0, 100) })} />
                      </Field>
                    </div>
                    <div className="sw-inlineToggle">
                      <div>
                        <div className="sw-inlineToggle__title">Auto module weights</div>
                        <div className="sw-inlineToggle__sub">Split all modules in this year equally.</div>
                      </div>
                      <Toggle checked={curYear.autoModuleWeights} onChange={(v) => updateYear(curYear.id, { autoModuleWeights: v })} label={curYear.autoModuleWeights ? "On" : "Off"} />
                    </div>
                  </div>

                  <div className="sw-listHead">
                    <div className="sw-block__title">Semesters</div>
                  </div>
                  <div className="sw-list">
                    {curYear.semesters.map((s) => {
                      const r = computeModuleSet(s.modules || []);
                      return (
                        <ListRow
                          key={s.id}
                          title={s.label || "Semester"}
                          sub={`${s.modules.length} module${s.modules.length === 1 ? "" : "s"}`}
                          value={r.count ? formatPct(r.avg) : "—"}
                          valueTone="accent"
                          onClick={() => goSemester(curYear.id, s.id)}
                        />
                      );
                    })}
                  </div>
                  <Button variant="ghost" className="sw-addRow" onClick={() => addSemester(curYear.id)}>+ Add semester</Button>
                  {state.years.length > 1 ? (
                    <div className="sw-removeRow">
                      <Button variant="danger" className="sw-btn--sm" onClick={() => { removeYear(curYear.id); goCourse(); }}>
                        <span className="sw-btn__icon" aria-hidden="true">✕</span> Remove year
                      </Button>
                    </div>
                  ) : null}
                </>
              ) : null}

              {/* -------- Semester -------- */}
              {nav.level === "semester" && curYear && curSem ? (
                <>
                  <div className="sw-fieldStack">
                    <Field label="Semester label">
                      <Input value={curSem.label} onChange={(e) => updateSemester(curYear.id, curSem.id, { label: e.target.value })} placeholder="Semester 1" />
                    </Field>
                  </div>

                  <div className="sw-listHead">
                    <div className="sw-block__title">Modules</div>
                  </div>
                  <div className="sw-list">
                    {curSem.modules.map((m) => {
                      const mk = computeModuleMark(m).mark;
                      return (
                        <ListRow
                          key={m.id}
                          title={m.code || m.title || "New module"}
                          sub={m.code && m.title ? m.title : `${(m.components || []).length} assessments`}
                          value={Number.isFinite(mk) ? formatPct(mk) : "—"}
                          valueTone="accent"
                          onClick={() => goModule(curYear.id, curSem.id, m.id)}
                        />
                      );
                    })}
                  </div>
                  <Button variant="ghost" className="sw-addRow" onClick={() => addModule(curYear.id, curSem.id)}>+ Add module</Button>
                  {curYear.semesters.length > 1 ? (
                    <div className="sw-removeRow">
                      <Button variant="danger" className="sw-btn--sm" onClick={() => { removeSemester(curYear.id, curSem.id); goYear(curYear.id); }}>
                        <span className="sw-btn__icon" aria-hidden="true">✕</span> Remove semester
                      </Button>
                    </div>
                  ) : null}
                </>
              ) : null}

              {/* -------- Module -------- */}
              {nav.level === "module" && curYear && curSem && curMod ? (
                <>
                  <div className="sw-fieldStack">
                    <div className="sw-grid2">
                      <Field label="Module code">
                        <Input value={curMod.code} onChange={(e) => updateModule(curYear.id, curSem.id, curMod.id, { code: e.target.value })} placeholder="CS301" />
                      </Field>
                      <Field label="Title">
                        <Input value={curMod.title} onChange={(e) => updateModule(curYear.id, curSem.id, curMod.id, { title: e.target.value })} placeholder="Machine Learning" />
                      </Field>
                    </div>
                    <div className="sw-grid2">
                      <Field label="Weight (% of year)" hint={curYear.autoModuleWeights ? "Auto — turn off in year settings" : undefined}>
                        <Input type="number" min={0} max={100} disabled={curYear.autoModuleWeights} value={String(curMod.weight ?? "")} onChange={(e) => updateModule(curYear.id, curSem.id, curMod.id, { weight: clampNumber(Number(e.target.value), 0, 100) })} />
                      </Field>
                      <Field label="Credits (optional)">
                        <Input type="number" min={0} value={String(curMod.credits ?? "")} onChange={(e) => updateModule(curYear.id, curSem.id, curMod.id, { credits: clampNumber(Number(e.target.value), 0, 200) })} />
                      </Field>
                    </div>
                    <div className="sw-result sw-result--sm">
                      <div className="sw-result__label">Module mark</div>
                      <div className="sw-result__value">{formatPct(computeModuleMark(curMod).mark)}</div>
                    </div>
                    <div className="sw-inlineToggle">
                      <div>
                        <div className="sw-inlineToggle__title">Auto assessment weights</div>
                        <div className="sw-inlineToggle__sub">Split assessments equally.</div>
                      </div>
                      <Toggle checked={curMod.autoComponentWeights} onChange={(v) => setAutoComponentWeights(curYear.id, curSem.id, curMod.id, v)} label={curMod.autoComponentWeights ? "On" : "Off"} />
                    </div>
                  </div>

                  <div className="sw-listHead">
                    <div className="sw-block__title">Assessments</div>
                  </div>
                  <div className="sw-list">
                    {curMod.components.map((c) => (
                      <ListRow
                        key={c.id}
                        title={c.name || c.type || "Assessment"}
                        sub={`${c.type} • ${Math.round(Number(c.weight) || 0)}% of module`}
                        value={formatPct(Number(c.mark))}
                        valueTone="accent"
                        onClick={() => setEditingCompId(c.id)}
                      />
                    ))}
                  </div>
                  <Button variant="ghost" className="sw-addRow" onClick={() => addComponent(curYear.id, curSem.id, curMod.id)}>+ Add assessment</Button>
                  {curSem.modules.length > 1 ? (
                    <div className="sw-removeRow">
                      <Button variant="danger" className="sw-btn--sm" onClick={() => { removeModule(curYear.id, curSem.id, curMod.id); goSemester(curYear.id, curSem.id); }}>
                        <span className="sw-btn__icon" aria-hidden="true">✕</span> Remove module
                      </Button>
                    </div>
                  ) : null}
                </>
              ) : null}
            </div>
          </div>
        </div>
      </div>

      {/* Hidden measurement card */}
      <div className="sw-measure" ref={measRef} aria-hidden="true">
        <div style={{ width: dims.w, height: dims.h }}>
          <WrappedCard state={state} view={view} blocks={view.blocks} pageLabel="" />
        </div>
      </div>

      {/* Assessment edit sheet */}
      {editingCompId && curMod ? (
        (() => {
          const c = curMod.components.find((x) => x.id === editingCompId);
          if (!c) return null;
          return (
            <div className="sw-sheet" onClick={(e) => { if (e.target === e.currentTarget) setEditingCompId(null); }}>
              <div className="sw-sheet__panel">
                <div className="sw-sheet__head">
                  <div className="sw-sheet__title">Edit assessment</div>
                  <button type="button" className="sw-sheet__close" onClick={() => setEditingCompId(null)}>Done</button>
                </div>
                <div className="sw-grid2">
                  <Field label="Type">
                    <Select value={c.type} onChange={(e) => updateComponent(curYear.id, curSem.id, curMod.id, c.id, { type: e.target.value, name: c.name === c.type ? e.target.value : c.name })}>
                      <option value="Overall">Overall</option>
                      <option value="Assignment">Assignment</option>
                      <option value="Coursework">Coursework</option>
                      <option value="Exam">Exam</option>
                      <option value="Other">Other</option>
                    </Select>
                  </Field>
                  <Field label="Label">
                    <Input value={c.name} onChange={(e) => updateComponent(curYear.id, curSem.id, curMod.id, c.id, { name: e.target.value })} placeholder="Coursework 1" />
                  </Field>
                  <Field label="Weight (% of module)" hint={curMod.autoComponentWeights ? "Auto" : undefined}>
                    <Input type="number" min={0} max={100} disabled={curMod.autoComponentWeights} value={String(c.weight ?? "")} onChange={(e) => updateComponent(curYear.id, curSem.id, curMod.id, c.id, { weight: clampNumber(Number(e.target.value), 0, 100) })} />
                  </Field>
                  <Field label="Mark (%)">
                    <Input type="number" min={0} max={100} value={String(c.mark ?? "")} onChange={(e) => updateComponent(curYear.id, curSem.id, curMod.id, c.id, { mark: clampNumber(Number(e.target.value), 0, 100) })} />
                  </Field>
                </div>
                {curMod.components.length > 1 ? (
                  <div className="sw-removeRow">
                    <Button variant="danger" className="sw-btn--sm" onClick={() => { removeComponent(curYear.id, curSem.id, curMod.id, c.id); setEditingCompId(null); }}>
                      <span className="sw-btn__icon" aria-hidden="true">✕</span> Remove assessment
                    </Button>
                  </div>
                ) : null}
              </div>
            </div>
          );
        })()
      ) : null}

      {/* Settings sheet (identity + design) */}
      {settingsOpen ? (
        <div className="sw-sheet" onClick={(e) => { if (e.target === e.currentTarget) setSettingsOpen(false); }}>
          <div className="sw-sheet__panel sw-sheet__panel--wide">
            <div className="sw-sheet__head">
              <div className="sw-sheet__title">Card settings</div>
              <button type="button" className="sw-sheet__close" onClick={() => setSettingsOpen(false)}>Done</button>
            </div>
            <div className="sw-sheet__scroll">
              <div className="sw-block__title">Details</div>
              <div className="sw-fieldStack">
                <Field label="Name"><Input value={state.personName} onChange={(e) => set({ personName: e.target.value })} /></Field>
                <Field label="University"><Input value={state.university} onChange={(e) => set({ university: e.target.value })} /></Field>
                <Field label="Course / Degree"><Input value={state.course} onChange={(e) => set({ course: e.target.value })} /></Field>
                <Field label="Subtitle"><Input value={state.subtitle} onChange={(e) => set({ subtitle: e.target.value })} placeholder="Class of 2026" /></Field>
                <Field label="Headline"><Input value={state.headline} onChange={(e) => set({ headline: e.target.value })} /></Field>
                <Field label="Caption (footer line)"><Input value={state.caption} onChange={(e) => set({ caption: e.target.value })} /></Field>
                <Field label="LinkedIn handle (optional)"><Input value={state.linkedinHandle} onChange={(e) => set({ linkedinHandle: e.target.value })} placeholder="@yourname" /></Field>
              </div>

              <div className="sw-block__title" style={{ marginTop: 18 }}>Design</div>
              <div className="sw-fieldStack">
                <div className="sw-grid2">
                  <Field label="Template">
                    <Select value={state.template} onChange={(e) => set({ template: e.target.value })}>
                      <option value="classic">Classic</option>
                      <option value="gradient">Gradient</option>
                    </Select>
                  </Field>
                  <Field label="Export format">
                    <Select value={state.format} onChange={(e) => set({ format: e.target.value })}>
                      <option value="linkedin_square">LinkedIn square</option>
                      <option value="instagram_story">Instagram story</option>
                    </Select>
                  </Field>
                </div>
                <Toggle checked={state.showModuleMarks} onChange={(v) => set({ showModuleMarks: v })} label="Show marks on the card" />
                <Toggle checked={state.showAssessmentsInBreakdown} onChange={(v) => set({ showAssessmentsInBreakdown: v })} label="Show assessments (semester view)" />
                <div className="sw-grid2">
                  <Field label="Primary"><div className="sw-colorRow"><input type="color" value={state.theme.primary} onChange={(e) => setTheme({ primary: e.target.value })} /><Input value={state.theme.primary} onChange={(e) => setTheme({ primary: e.target.value })} /></div></Field>
                  <Field label="Accent"><div className="sw-colorRow"><input type="color" value={state.theme.accent} onChange={(e) => setTheme({ accent: e.target.value })} /><Input value={state.theme.accent} onChange={(e) => setTheme({ accent: e.target.value })} /></div></Field>
                  <Field label="Background"><div className="sw-colorRow"><input type="color" value={state.theme.bg} onChange={(e) => setTheme({ bg: e.target.value })} /><Input value={state.theme.bg} onChange={(e) => setTheme({ bg: e.target.value })} /></div></Field>
                  <Field label="Card base"><div className="sw-colorRow"><input type="color" value={state.theme.card} onChange={(e) => setTheme({ card: e.target.value })} /><Input value={state.theme.card} onChange={(e) => setTheme({ card: e.target.value })} /></div></Field>
                </div>
                <Button variant="ghost" onClick={() => setTheme(DEFAULTS.theme)}>Reset colours</Button>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function curShareYear(state, share) {
  return state.years.find((y) => y.id === share.yearId) || state.years[0] || null;
}

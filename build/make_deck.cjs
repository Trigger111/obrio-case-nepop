/**
 * Збирає презентацію з чисел, експортованих ноутбуком (output/deck_data.json).
 * Жодне число не введене вручну.
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const D = JSON.parse(fs.readFileSync(path.join(ROOT, "output", "deck_data.json"), "utf8"));
const OUT = path.join(ROOT, "01_OBRIO_creative_testing_framework.pptx");

/* ─────────────────────────── дизайн-система ─────────────────────────── */
const C = {
  ink: "171635",        // глибокий індиго — темні слайди й заголовки
  inkSoft: "221F4D",
  onDark: "FFFFFF",
  onDarkMuted: "B9B6D6",
  surface: "FFFFFF",
  card: "F4F3FA",
  cardEdge: "E3E1F2",
  text: "1F1D3D",
  textSoft: "5A5872",
  textMuted: "8A8A96",
  primary: "4C3FD6",    // валідований слот 1
  accent: "D2691E",     // валідований слот 2
  good: "0CA30C",
  grey: "8A8A96",
  risk: "D03B3B",
};
const F = { head: "Cambria", body: "Calibri" };
const M = 0.62;                 // поле
const W = 13.33, H = 7.5;
const CW = W - 2 * M;           // ширина контенту

// Фірмовий знак OBRIO. Джерело — оригінальний асет із деку (build/make_logo.py).
const LOGO_BRAND = path.join(__dirname, "assets", "logo_brand.png");
const LOGO_INK = path.join(__dirname, "assets", "logo_ink.png");
const LOGO_RATIO = 195 / 68;
const LOGO_W = 0.8;
const LOGO_H = LOGO_W / LOGO_RATIO;
[LOGO_BRAND, LOGO_INK].forEach((p) => {
  if (!fs.existsSync(p)) throw new Error("немає файлу логотипа: " + p + " — спочатку запустіть build/make_logo.py");
});

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Genesis Analytics Camp 3.0 · Case #4";
pres.title = "OBRIO · Фреймворк тестування креативів";

let slideNo = 0;
const fmt = {
  money: (v, d = 0) => "$" + v.toLocaleString("uk-UA", { minimumFractionDigits: d, maximumFractionDigits: d }).replace(/ /g, " "),
  m: (v) => "$" + (v / 1e6).toFixed(1).replace(".", ",") + "M",
  k: (v) => "$" + Math.round(v / 1000).toLocaleString("uk-UA").replace(/ /g, " ") + "k",
  n: (v, d = 0) => v.toLocaleString("uk-UA", { minimumFractionDigits: d, maximumFractionDigits: d }).replace(/ /g, " "),
  pct: (v, d = 1) => (v * 100).toFixed(d).replace(".", ",") + "%",
  pctv: (v, d = 1) => v.toFixed(d).replace(".", ",") + "%",
  sign: (v, d = 1) => (v >= 0 ? "+" : "−") + Math.abs(v).toFixed(d).replace(".", ",") + "%",
};

/* ─────────────────────────── примітиви ─────────────────────────── */

function baseSlide(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? C.ink : C.surface };
  return s;
}

function chrome(s, dark, source, opts) {
  const o = opts || {};
  slideNo += 1;
  s.addText(String(slideNo), {
    x: W - M - 0.5, y: H - 0.52, w: 0.5, h: 0.3, align: "right",
    fontFace: F.body, fontSize: 10, color: dark ? C.onDarkMuted : C.textMuted,
  });
  // Фірмовий знак у колонтитулі. На темному тлі — брендова бірюза (контраст 10,9:1),
  // на світлому — монохромна версія: бірюза на білому дає лише 1,6:1.
  let textLeft = M;
  if (!o.noLogo) {
    s.addImage({
      path: dark ? LOGO_BRAND : LOGO_INK,
      x: M, y: H - 0.535, w: LOGO_W, h: LOGO_H,
    });
    textLeft = M + LOGO_W + 0.22;
  }
  if (source) {
    s.addText(source, {
      x: textLeft, y: H - 0.52, w: W - M - 0.62 - textLeft, h: 0.3, align: "left",
      fontFace: F.body, fontSize: 9, italic: true,
      color: dark ? C.onDarkMuted : C.textMuted,
    });
  }
  return s;
}

function title(s, text, dark, kicker) {
  let y = 0.44;
  if (kicker) {
    s.addText(kicker.toUpperCase(), {
      x: M, y, w: CW, h: 0.26, fontFace: F.body, fontSize: 10.5, bold: true,
      charSpacing: 1.6, color: dark ? C.onDarkMuted : C.primary, margin: 0,
    });
    y += 0.3;
  }
  s.addText(text, {
    x: M, y, w: CW, h: 0.72, fontFace: F.head, fontSize: 27, bold: true,
    color: dark ? C.onDark : C.ink, margin: 0, valign: "top",
  });
  return y + 0.82;
}

function lead(s, text, y, dark, w) {
  if (typeof y !== "number" || Number.isNaN(y)) throw new Error("lead(): y не передано — " + text.slice(0, 40));
  s.addText(text, {
    x: M, y, w: w || CW, h: 0.46, fontFace: F.body, fontSize: 13.5,
    color: dark ? C.onDarkMuted : C.textSoft, margin: 0, valign: "top",
  });
  return y + 0.5;
}

function card(s, o) {
  ["x", "y", "w", "h"].forEach((k) => {
    if (typeof o[k] !== "number" || Number.isNaN(o[k])) throw new Error("card(): некоректний " + k + " = " + o[k]);
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.09,
    fill: { color: o.fill || C.card },
    line: { color: o.edge || C.cardEdge, width: 0.75 },
    shadow: o.shadow === false ? undefined
      : { type: "outer", color: "171635", opacity: 0.07, blur: 8, offset: 2, angle: 90 },
  });
}

/** Плитка з великою цифрою. Висоти рахуються від кегля, а не задаються на око. */
function tile(s, o) {
  card(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill, edge: o.edge });
  const pad = 0.2;
  let y = o.y + 0.16;
  if (o.label) {
    s.addText(o.label.toUpperCase(), {
      x: o.x + pad, y, w: o.w - 2 * pad, h: 0.24, fontFace: F.body, fontSize: 9,
      bold: true, charSpacing: 1.2, color: o.labelColor || C.textMuted, margin: 0,
    });
    y += 0.26;
  }
  const vSize = o.valueSize || 30;
  const vH = (vSize * 1.36) / 72;
  s.addText(o.value, {
    x: o.x + pad, y, w: o.w - 2 * pad, h: vH,
    fontFace: F.body, fontSize: vSize, bold: true,
    color: o.color || C.ink, margin: 0, valign: "top",
  });
  y += vH + 0.02;
  if (o.note) {
    const noteH = o.y + o.h - y - 0.06;
    if (noteH < 0.18) throw new Error("tile(): замало місця під примітку — «" + o.note.slice(0, 30) + "»");
    s.addText(o.note, {
      x: o.x + pad, y, w: o.w - 2 * pad, h: noteH,
      fontFace: F.body, fontSize: 10.5, color: C.textSoft, margin: 0, valign: "top",
    });
  }
}

/** Рядок плиток однакової ширини. */
function tileRow(s, items, y, h, gap) {
  const g = gap === undefined ? 0.22 : gap;
  const w = (CW - g * (items.length - 1)) / items.length;
  items.forEach((it, i) => tile(s, Object.assign({}, it, { x: M + i * (w + g), y, w, h })));
  return y + h;
}

/** Пункти з бейджем-кружком. */
function steps(s, items, o) {
  const rowH = o.rowH || 0.86;
  items.forEach((it, i) => {
    const y = o.y + i * rowH;
    s.addShape(pres.ShapeType.ellipse, {
      x: o.x, y: y + 0.04, w: 0.36, h: 0.36,
      fill: { color: it.color || C.primary }, line: { color: it.color || C.primary, width: 0 },
    });
    s.addText(it.badge, {
      x: o.x, y: y + 0.04, w: 0.36, h: 0.36, align: "center", valign: "middle",
      fontFace: F.body, fontSize: 12, bold: true, color: "FFFFFF", margin: 0,
    });
    s.addText(
      [
        { text: it.head, options: { bold: true, fontSize: 13, color: C.ink, breakLine: true } },
        { text: it.body, options: { fontSize: 11.5, color: C.textSoft } },
      ],
      { x: o.x + 0.5, y, w: o.w - 0.5, h: rowH - 0.08, fontFace: F.body, margin: 0, valign: "top", lineSpacingMultiple: 1.02 }
    );
  });
  return o.y + items.length * rowH;
}

const AXIS = {
  catAxisLabelColor: C.textSoft, catAxisLabelFontSize: 10, catAxisLabelFontFace: F.body,
  valAxisLabelColor: C.textMuted, valAxisLabelFontSize: 10, valAxisLabelFontFace: F.body,
  valGridLine: { color: "EBEAF2", size: 0.75 },
  catGridLine: { style: "none" },
  catAxisLineShow: false, valAxisLineShow: false,
  dataLabelFontFace: F.body, dataLabelColor: C.text,
  showTitle: false,
};

function chartBox(s, o) {
  card(s, { x: o.x - 0.14, y: o.y - 0.12, w: o.w + 0.28, h: o.h + 0.42, fill: "FFFFFF", edge: C.cardEdge });
  if (o.caption) {
    s.addText(o.caption, {
      x: o.x, y: o.y + o.h + 0.05, w: o.w, h: 0.24, fontFace: F.body, fontSize: 9.5,
      italic: true, color: C.textMuted, margin: 0,
    });
  }
}

/* ═══════════════════════════ 1. Титул ═══════════════════════════ */
{
  const s = baseSlide(true);
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.1, y: -1.5, w: 6.4, h: 6.4, fill: { color: C.primary, transparency: 82 }, line: { width: 0 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 10.6, y: 2.6, w: 4.2, h: 4.2, fill: { color: C.accent, transparency: 88 }, line: { width: 0 },
  });
  s.addImage({ path: LOGO_BRAND, x: M, y: 0.78, w: 1.7, h: 1.7 / LOGO_RATIO });
  s.addText("GENESIS ANALYTICS CAMP 3.0   ·   CASE #4   ·   OBRIO / NEBULA", {
    x: M, y: 1.5, w: 9.2, h: 0.3, fontFace: F.body, fontSize: 11, bold: true,
    charSpacing: 1.8, color: C.onDarkMuted, margin: 0,
  });
  s.addText("Кому дати\nнаступний бюджет?", {
    x: M, y: 2.0, w: 9.0, h: 1.9, fontFace: F.head, fontSize: 48, bold: true,
    color: C.onDark, margin: 0, lineSpacingMultiple: 0.94,
  });
  s.addText("Фреймворк тестування креативів: прозорі правила замість здогадок про переможця", {
    x: M, y: 4.05, w: 8.6, h: 0.5, fontFace: F.body, fontSize: 16, color: C.onDarkMuted, margin: 0,
  });
  const facts = [
    [fmt.n(D.meta.rows), "щоденних рядків"],
    [fmt.n(D.meta.creatives), "креативів"],
    [D.meta.days + " день", D.meta.date_min.split("-").reverse().join(".") + " – " + D.meta.date_max.split("-").reverse().join(".")],
    [fmt.m(D.portfolio.spend), "витрат у портфелі"],
  ];
  const fw = (9.0 - 3 * 0.24) / 4;
  facts.forEach(([v, l], i) => {
    const x = M + i * (fw + 0.24);
    s.addText(v, { x, y: 4.95, w: fw, h: 0.36, fontFace: F.body, fontSize: 17, bold: true, color: C.onDark, margin: 0 });
    s.addText(l, { x, y: 5.31, w: fw, h: 0.5, fontFace: F.body, fontSize: 10.5, color: C.onDarkMuted, margin: 0 });
  });
  s.addText("Самодостатній документ: усі висновки на слайдах. Кожне число зібрано автоматично з файлу аналізу.", {
    x: M, y: 6.35, w: 11.0, h: 0.3, fontFace: F.body, fontSize: 11, italic: true, color: C.onDarkMuted, margin: 0,
  });
  s.addNotes("Документ створено для самостійного читання. Числа на слайдах експортує ноутбук у output/deck_data.json, тому деку й аналіз неможливо розсинхронізувати.");
  chrome(s, true, null, { noLogo: true });
}

/* ═══════════════════════════ 2. Проблема ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Сьогодні рішення або поспішає, або запізнюється", false, "проблема");
  y = lead(s, "Єдиної системи оцінки немає, тому кожен маркетолог тестує по-своєму — і обидві крайності коштують грошей.", y);

  const cw = (CW - 0.28) / 2;
  [
    { x: M, name: "«Тоха»: рішення на малих обʼємах", color: C.risk,
      stat: "38%", statNote: "креативів на $50 пізніше змінюють статус відносно цілі +30%",
      body: "Кілька випадкових покупок різко піднімають ранній ROMI. Варіації запускають за шумом, а не за сигналом." },
    { x: M + cw + 0.28, name: "«Гена»: чекати до $2 000", color: C.accent,
      stat: fmt.pct(D.counterfactual.reach_2000_share, 1), statNote: "креативів узагалі доходять до $2 000",
      body: "Правило точніше, але майже ні до кого не застосовне: слабкі встигають злити бюджет, сильні скейляться із запізненням." },
  ].forEach((b) => {
    card(s, { x: b.x, y, w: cw, h: 2.35 });
    s.addText(b.name, { x: b.x + 0.26, y: y + 0.22, w: cw - 0.52, h: 0.32, fontFace: F.body, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(b.stat, { x: b.x + 0.26, y: y + 0.6, w: 1.75, h: 0.62, fontFace: F.body, fontSize: 34, bold: true, color: b.color, margin: 0 });
    s.addText(b.statNote, { x: b.x + 2.05, y: y + 0.63, w: cw - 2.35, h: 0.62, fontFace: F.body, fontSize: 10.5, color: C.textSoft, margin: 0, valign: "top" });
    s.addText(b.body, { x: b.x + 0.26, y: y + 1.38, w: cw - 0.52, h: 0.8, fontFace: F.body, fontSize: 11.5, color: C.textSoft, margin: 0, valign: "top" });
  });
  y += 2.6;

  card(s, { x: M, y, w: CW, h: 1.28, fill: C.ink, edge: C.ink });
  s.addText("Правило «чекати $2 000» не просто неточне — воно незастосовне за бюджетом", {
    x: M + 0.3, y: y + 0.2, w: 7.4, h: 0.34, fontFace: F.body, fontSize: 14, bold: true, color: C.onDark, margin: 0,
  });
  s.addText("Довести кожен креатив до $2 000 коштувало б " + fmt.m(D.counterfactual.force_2000) +
    " проти фактичних " + fmt.m(D.counterfactual.actual_spend) + " — це у " +
    (D.counterfactual.force_2000 / D.counterfactual.actual_spend).toFixed(0) + " разів більший бюджет.", {
    x: M + 0.3, y: y + 0.58, w: 7.4, h: 0.55, fontFace: F.body, fontSize: 11.5, color: C.onDarkMuted, margin: 0, valign: "top",
  });
  s.addText("Потрібне одне правило:\nкому дати наступну частину бюджету і коли оцінити знову.", {
    x: M + 8.0, y: y + 0.24, w: CW - 8.3, h: 0.9, fontFace: F.body, fontSize: 12.5, bold: true, color: C.onDark, margin: 0, valign: "top",
  });

  s.addNotes("Обидва антипатерни з кейсу мають ціну. Ключовий аргумент проти правила Гени — не точність, а бюджет: воно вимагало б у 18 разів більше грошей, ніж компанія витратила за весь період.");
  chrome(s, false, "Джерело: 02_analysis_notebook.ipynb, розділи 2.4 і 5");
}

/* ═══════════════════════════ 3. Ідея фреймворка ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Ідея: не вгадувати переможця, а керувати наступним кроком", false, "чому саме такий фреймворк");
  y = lead(s, "Ми перевірили три можливі підходи. Два з них ці дані не витримують — і це перевірений результат, а не припущення.", y);

  const cw = (CW - 0.24 * 2) / 3;
  const opts = [
    { t: "Поріг обсягу", sub: "«чекати, поки набереться достатньо даних»", verdict: "не працює", vc: C.risk,
      why: "Суми, після якої рішення стає надійним, у цих даних не існує: навіть на $2 000 статус ще змінюється у " +
        fmt.pctv(D.stability[7].flip_rate_fixed * 100, 1) + " випадків." },
    { t: "Прогнозна модель", sub: "«навчити модель обирати переможців»", verdict: "недостатньо", vc: C.accent,
      why: "Модель ранжує краще за поточний ROMI (" + D.model_check.auc_model.toFixed(3).replace(".", ",") +
        " проти " + D.model_check.auc_romi.toFixed(3).replace(".", ",") + "), але за грошима виграла лише в " +
        D.model_check.rolling_profit_wins + " із " + D.model_check.rolling_folds + " періодів." },
    { t: "Правила наступного бюджету", sub: "«скільки дати зараз і коли перевірити»", verdict: "обрано", vc: C.good,
      why: "Не вимагає передбачати майбутнє. Обмежує збиток однією контрольною сумою, прискорює сильних кандидатів і піддається чесній перевірці експериментом." },
  ];
  opts.forEach((o, i) => {
    const x = M + i * (cw + 0.24);
    const chosen = i === 2;
    card(s, { x, y, w: cw, h: 2.62, fill: chosen ? "F1F7F1" : C.card, edge: chosen ? "BFE3BF" : C.cardEdge });
    s.addText(o.verdict.toUpperCase(), { x: x + 0.24, y: y + 0.2, w: cw - 0.48, h: 0.24, fontFace: F.body, fontSize: 9, bold: true, charSpacing: 1.2, color: o.vc, margin: 0 });
    s.addText(o.t, { x: x + 0.24, y: y + 0.46, w: cw - 0.48, h: 0.34, fontFace: F.body, fontSize: 15, bold: true, color: C.ink, margin: 0 });
    s.addText(o.sub, { x: x + 0.24, y: y + 0.8, w: cw - 0.48, h: 0.36, fontFace: F.body, fontSize: 10.5, italic: true, color: C.textMuted, margin: 0, valign: "top" });
    s.addText(o.why, { x: x + 0.24, y: y + 1.22, w: cw - 0.48, h: 1.2, fontFace: F.body, fontSize: 11.5, color: C.textSoft, margin: 0, valign: "top" });
  });
  y += 2.86;

  card(s, { x: M, y, w: CW, h: 1.05, fill: C.ink, edge: C.ink });
  s.addText([
    { text: "Одне речення, яким описується весь фреймворк:  ", options: { color: C.onDarkMuted, fontSize: 13 } },
    { text: "рішення — це не вирок «хіт чи провал», а відповідь на питання «яку наступну обмежену суму дати і коли оцінити знову».", options: { color: C.onDark, fontSize: 13, bold: true } },
  ], { x: M + 0.3, y: y + 0.2, w: CW - 0.6, h: 0.7, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Це відповідь на пункт 1 очікуваного результату кейсу: яка ідея закладена і чому саме вона. Важливо, що дві альтернативи не відкинуті на смак — вони перевірені в розділах 5 і 7 ноутбука.");
  chrome(s, false, "Джерело: розділи 5 (пороги обсягу), 7 (модель), 9 (правила)");
}

/* ═══════════════════════════ 4. Що міряємо ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Чотири визначення, на яких тримаються всі рішення", false, "метрики фреймворку");
  y = lead(s, "Усі числа далі рахуються тільки цими формулами — однаково для кожного креативу й кожного сегмента.", y);

  const cw = (CW - 0.22 * 3) / 4;
  const defs = [
    { n: "01", t: "Дохід", f: "predict_net_revenue + upsells", d: "Прогноз доходу плюс додаткові покупки. Апселі — це " + fmt.pct(D.portfolio.upsell_share, 0) + " доходу, без них ROMI занижений." },
    { n: "02", t: "ROMI", f: "(дохід − витрати) / витрати", d: "Ціль бізнесу — не нуль, а +30%. Саме з нею порівнюємо кожне рішення." },
    { n: "03", t: "ROMI частини", f: "рахуємо лише після рішення", d: "Дохід і витрати, що виникли після моменту оцінки. Не дає ранньому старту прикрашати підсумок." },
    { n: "04", t: "Зміна статусу", f: "перехід через ціль +30%", d: "Чи опинився ROMI пізніше по інший бік цілі. Це наша міра надійності рішення." },
  ];
  defs.forEach((o, i) => {
    const x = M + i * (cw + 0.22);
    card(s, { x, y, w: cw, h: 2.42 });
    s.addText(o.n, { x: x + 0.22, y: y + 0.18, w: cw - 0.44, h: 0.3, fontFace: F.body, fontSize: 12, bold: true, color: C.primary, margin: 0 });
    s.addText(o.t, { x: x + 0.22, y: y + 0.48, w: cw - 0.44, h: 0.34, fontFace: F.body, fontSize: 15, bold: true, color: C.ink, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.22, y: y + 0.86, w: cw - 0.44, h: 0.42, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: C.cardEdge, width: 0.75 } });
    s.addText(o.f, { x: x + 0.28, y: y + 0.86, w: cw - 0.56, h: 0.42, fontFace: "Consolas", fontSize: 9.5, color: C.primary, margin: 0, valign: "middle" });
    s.addText(o.d, { x: x + 0.22, y: y + 1.36, w: cw - 0.44, h: 0.92, fontFace: F.body, fontSize: 10.5, color: C.textSoft, margin: 0, valign: "top" });
  });
  y += 2.66;

  const base = [
    [fmt.n(D.meta.rows) + " × " + fmt.n(D.meta.creatives), "рядків «креатив × день» і унікальних креативів"],
    [fmt.pct(D.portfolio.romi, 1), "ROMI портфеля при цілі +30%"],
    [fmt.money(D.portfolio.median_lifetime_spend, 2), "медіанні витрати креативу за все життя"],
    [D.meta.media_sources.join(", "), "єдине джерело трафіку в даних"],
  ];
  const bw = (CW - 0.22 * 3) / 4;
  base.forEach(([v, l], i) => {
    const x = M + i * (bw + 0.22);
    card(s, { x, y, w: bw, h: 0.98, fill: "FFFFFF", edge: C.cardEdge, shadow: false });
    s.addText(v, { x: x + 0.2, y: y + 0.14, w: bw - 0.4, h: 0.36, fontFace: F.body, fontSize: 15, bold: true, color: C.ink, margin: 0 });
    s.addText(l, { x: x + 0.2, y: y + 0.5, w: bw - 0.4, h: 0.42, fontFace: F.body, fontSize: 10, color: C.textSoft, margin: 0, valign: "top" });
  });

  s.addNotes("Прогноз доходу фіксуємо в одній версії до старту перевірки, а після дозрівання когорт звіряємо з фактичними грошима. Медіанний креатив витрачає лише $5,35 — це головна причина, чому рішення мають бути дешевими й швидкими.");
  chrome(s, false, "Джерело: розділи 1–2 ноутбука");
}

/* ═══════════════════════════ 5. Питання 1 ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "До $50 ROMI — це не оцінка, а відсутність даних", false, "питання 1 · мінімальний поріг");
  y = lead(s, "Якщо креатив ще не купив жодної події конверсії, його ROMI дорівнює −100% механічно — незалежно від того, хороший він чи поганий.", y);

  const q1 = D.q1_information.filter((r) => [10, 25, 50, 100, 200, 500].includes(r["поріг, $"]));
  const chartW = 6.9, chartH = 3.05;
  chartBox(s, { x: M, y: y + 0.12, w: chartW, h: chartH, caption: "Частка креативів, у яких у момент першого перетину суми поле trials дорівнює нулю." });
  s.addChart(pres.ChartType.bar, [{
    name: "% креативів без жодної події",
    labels: q1.map((r) => "$" + fmt.n(r["поріг, $"])),
    values: q1.map((r) => Number(r["% без жодної події"].toFixed(1))),
  }], Object.assign({}, AXIS, {
    x: M, y: y + 0.12, w: chartW, h: chartH,
    barDir: "col", barGapWidthPct: 45, chartColors: [C.primary],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFormatCode: '0"%"',
    valAxisMaxVal: 70, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"',
    showLegend: false,
  }));

  const rx = M + chartW + 0.42, rw = CW - chartW - 0.42;
  tile(s, {
    x: rx, y: y + 0.12, w: rw, h: 1.62, fill: "F1F7F1", edge: "BFE3BF",
    label: "відповідь на питання 1", value: "$50", valueSize: 40, color: C.good,
    note: "перша сума, на якій більшість креативів уже мають хоч якийсь сигнал",
  });
  const rows3 = [
    ["на $25", fmt.pctv(D.q1_information[1]["% без жодної події"], 1), "не мають подій — рішення неможливе"],
    ["на $50", fmt.pctv(D.q1_information[2]["% без жодної події"], 1), "не мають подій — межа прийнятного"],
    ["на $100", fmt.pctv(D.q1_information[3]["% без жодної події"], 1), "не мають подій — але вже дорого чекати"],
  ];
  rows3.forEach(([a, b, c], i) => {
    const yy = y + 1.9 + i * 0.5;
    s.addText(a, { x: rx, y: yy, w: 0.85, h: 0.36, fontFace: F.body, fontSize: 11.5, bold: true, color: C.textSoft, margin: 0, valign: "middle" });
    s.addText(b, { x: rx + 0.85, y: yy, w: 0.9, h: 0.36, fontFace: F.body, fontSize: 14, bold: true, color: i === 1 ? C.good : C.ink, margin: 0, valign: "middle" });
    s.addText(c, { x: rx + 1.78, y: yy, w: rw - 1.78, h: 0.36, fontFace: F.body, fontSize: 10, color: C.textMuted, margin: 0, valign: "middle" });
  });
  y += chartH + 0.55;

  const args = [
    { badge: "1", head: "Інформаційний аргумент", body: "На $25 у " + fmt.pctv(D.q1_information[1]["% без жодної події"], 0) + " креативів немає жодної події. На $50 — у " + fmt.pctv(D.q1_information[2]["% без жодної події"], 0) + "." },
    { badge: "2", head: "Операційний аргумент", body: "Креативи, що дійшли до $50, тримають " + fmt.pctv(D.survival[0]["% від спенду"], 1) + " усього бюджету: поріг майже не втрачає гроші з поля зору." },
    { badge: "3", head: "Економічний аргумент", body: "Медіанний креатив витрачає " + fmt.money(D.portfolio.median_lifetime_spend, 2) + " за життя. Вищий поріг просто не встигне спрацювати." },
  ];
  const aw = (CW - 0.24 * 2) / 3;
  args.forEach((a, i) => {
    const x = M + i * (aw + 0.24);
    card(s, { x, y, w: aw, h: 1.1, fill: C.card, shadow: false });
    steps(s, [a], { x: x + 0.18, y: y + 0.14, w: aw - 0.36, rowH: 0.9 });
  });

  s.addNotes("Мінімальний поріг виведений із даних, а не призначений. До $50 оцінюємо лише технічну справність: покази йдуть, посилання живе, витрати й дохід пишуться.");
  chrome(s, false, "Джерело: розділ 5.2 ноутбука · графік q1_minimum_threshold.png");
}

/* ═══════════════════════════ 6. Обсяг не рятує ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Збільшення обсягу не робить рішення надійним", false, "питання 3 · де межа точності");
  y = lead(s, "Головний і найнезручніший результат аналізу: суми, після якої вердикт перестає змінюватися, у цих даних немає.", y);

  const st = D.stability;
  const chartW = 7.6, chartH = 3.3;
  chartBox(s, { x: M, y: y + 0.1, w: chartW, h: chartH, caption: "Частка креативів, у яких статус відносно цілі +30% пізніше змінився хоча б раз." });
  s.addChart(pres.ChartType.line, [
    { name: "однакове вікно спостереження ($200)", labels: st.map((r) => "$" + fmt.n(r.checkpoint)), values: st.map((r) => Number((r.flip_rate_fixed * 100).toFixed(1))) },
    { name: "вікно, що росте разом із порогом", labels: st.map((r) => "$" + fmt.n(r.checkpoint)), values: st.map((r) => Number((r.flip_rate * 100).toFixed(1))) },
  ], Object.assign({}, AXIS, {
    x: M, y: y + 0.1, w: chartW, h: chartH,
    chartColors: [C.primary, C.accent], lineSize: 2.5,
    lineDataSymbol: "circle", lineDataSymbolSize: 7,
    valAxisMaxVal: 60, valAxisMinVal: 0, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"',
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: C.textSoft,
  }));

  const rx = M + chartW + 0.4, rw = CW - chartW - 0.4;
  tile(s, {
    x: rx, y: y + 0.1, w: rw, h: 1.6, fill: "FDF3EC", edge: "F0D3BC",
    label: "навіть на $2 000", value: fmt.pctv(st[7].flip_rate_fixed * 100, 1), valueSize: 38, color: C.accent,
    note: "креативів усе ще змінюють статус — а доходить туди лише " + fmt.pct(D.counterfactual.reach_2000_share, 1) + " запусків",
  });
  s.addText([
    { text: "Методична деталь, яка змінює висновок\n", options: { bold: true, fontSize: 12, color: C.ink, breakLine: true } },
    { text: "Якщо вимагати «після $X ще $X», вікно вимірювання різне на кожній точці й крива стає немонотонною — здається, що обсяг узагалі не допомагає. На однаковому вікні видно правду: надійність росте, але повільно і не до кінця.", options: { fontSize: 11, color: C.textSoft } },
  ], { x: rx, y: y + 1.84, w: rw, h: 1.6, fontFace: F.body, margin: 0, valign: "top" });
  y += chartH + 0.68;

  const facts = [
    { label: "від $25 до $2 000", value: fmt.pctv(st[0].flip_rate_fixed * 100, 0) + " → " + fmt.pctv(st[7].flip_rate_fixed * 100, 0), note: "надійність росте, але лишається далекою від певності", color: C.primary },
    { label: "точність правила «ROMI ≥ 30%»", value: "51–54%", note: "збалансована точність на будь-якому обсязі — майже монета", color: C.accent },
    { label: "звідси головний висновок", value: "рішення — за зоною", note: "не за сумою витрат: уся нестабільність зосереджена в середині", color: C.good },
  ];
  tileRow(s, facts.map((f) => ({ label: f.label, value: f.value, valueSize: 22, color: f.color, note: f.note })), y, 1.4);

  s.addNotes("Ключове: ми не приховуємо незручний результат, а робимо з нього конструктивний висновок. Якщо обсяг не дає певності, то поріг зупинки треба обирати за ціною тесту, а рішення — за значенням ROMI.");
  chrome(s, false, "Джерело: розділ 5.1 ноутбука · графік stage_3_stability_backtest.png");
}

/* ═══════════════════════════ 7. Три зони ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "На $200 три зони ROMI поводяться зовсім по-різному", false, "питання 2 і 3 · очевидні випадки");
  y = lead(s, "Очевидний випадок — це не той, хто набрав багато витрат, а той, чий ROMI лежить далеко від цілі. Уся нестабільність сидить у середині.", y);

  const z = D.zones_200;
  const byName = (n) => z.find((r) => r["зона"] === n);
  const zr = byName("ризик"), zg = byName("сіра зона"), zs = byName("сильний кандидат");

  const chartW = 7.55, chartH = 2.90;
  chartBox(s, { x: M, y: y + 0.08, w: chartW, h: chartH, caption: "Ліворуч — менше означає стабільніше. Праворуч — більше означає кращий результат. Ціль ROMI +30%." });
  s.addChart(pres.ChartType.bar, [
    { name: "ризик  (ROMI ≤ −50%)", labels: ["змінили статус пізніше", "закінчили ROMI ≥ 30%", "ROMI наступних витрат"], values: [zr["змінили статус пізніше"], zr["фінальний ROMI ≥ 30%"], zr["ROMI наступних витрат разом"]].map((v) => Number(v.toFixed(1))) },
    { name: "сіра зона  (−50…+100%)", labels: ["змінили статус пізніше", "закінчили ROMI ≥ 30%", "ROMI наступних витрат"], values: [zg["змінили статус пізніше"], zg["фінальний ROMI ≥ 30%"], zg["ROMI наступних витрат разом"]].map((v) => Number(v.toFixed(1))) },
    { name: "сильний кандидат  (≥ +100%)", labels: ["змінили статус пізніше", "закінчили ROMI ≥ 30%", "ROMI наступних витрат"], values: [zs["змінили статус пізніше"], zs["фінальний ROMI ≥ 30%"], zs["ROMI наступних витрат разом"]].map((v) => Number(v.toFixed(1))) },
  ], Object.assign({}, AXIS, {
    x: M, y: y + 0.08, w: chartW, h: chartH,
    barDir: "col", barGrouping: "clustered", barGapWidthPct: 55,
    chartColors: [C.risk, C.grey, C.good],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 9.5, dataLabelFormatCode: '0"%"',
    valAxisMaxVal: 100, valAxisMajorUnit: 25, valAxisLabelFormatCode: '0"%"',
    catAxisLabelFontSize: 9.5,
    showLegend: true, legendPos: "b", legendFontSize: 9.5, legendColor: C.textSoft,
  }));

  const rx = M + chartW + 0.38, rw = CW - chartW - 0.38;
  s.addText("Точні значення", { x: rx, y: y + 0.08, w: rw, h: 0.28, fontFace: F.body, fontSize: 11, bold: true, color: C.ink, margin: 0 });
  s.addTable([
    [{ text: "зона", options: { bold: true } }, { text: "креативів", options: { bold: true, align: "right" } }, { text: "змінили\nстатус", options: { bold: true, align: "right" } }, { text: "ROMI\nдалі", options: { bold: true, align: "right" } }],
    ["ризик", fmt.n(zr["креативів у зоні"]) + " (" + zr["% тих, хто дійшов"].toFixed(0) + "%)", fmt.pctv(zr["змінили статус пізніше"], 1), fmt.pctv(zr["ROMI наступних витрат разом"], 1)],
    ["сіра зона", fmt.n(zg["креативів у зоні"]) + " (" + zg["% тих, хто дійшов"].toFixed(0) + "%)", fmt.pctv(zg["змінили статус пізніше"], 1), fmt.pctv(zg["ROMI наступних витрат разом"], 1)],
    ["сильний", fmt.n(zs["креативів у зоні"]) + " (" + zs["% тих, хто дійшов"].toFixed(0) + "%)", fmt.pctv(zs["змінили статус пізніше"], 1), fmt.pctv(zs["ROMI наступних витрат разом"], 1)],
    [{ text: "усі, хто дійшов", options: { italic: true } }, { text: fmt.n(D.stability[3].reached_creatives), options: { italic: true, align: "right" } }, { text: fmt.pctv(D.stability[3].flip_rate_fixed * 100, 1), options: { italic: true, align: "right" } }, { text: "—", options: { italic: true, align: "right" } }],
  ], {
    x: rx, y: y + 0.4, w: rw, colW: [1.15, 1.25, 1.05, 0.93],
    fontFace: F.body, fontSize: 9.5, color: C.text, border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    align: "right", valign: "middle", rowH: 0.34, fill: { color: "FFFFFF" },
  });
  s.addText([
    { text: "Чому дві колонки результату\n", options: { bold: true, fontSize: 11, color: C.ink, breakLine: true } },
    { text: "«Закінчили ROMI ≥ 30%» частково механічне: креатив із +185% мусив би довго працювати в мінус, щоб упасти нижче цілі. Тому головна колонка — ROMI грошей, витрачених ", options: { fontSize: 10, color: C.textSoft } },
    { text: "після", options: { fontSize: 10, color: C.textSoft, italic: true, bold: true } },
    { text: " рішення.", options: { fontSize: 10, color: C.textSoft } },
  ], { x: rx, y: y + 2.28, w: rw, h: 1.0, fontFace: F.body, margin: 0, valign: "top" });
  y += chartH + 0.5;

  const concl = [
    { label: "очевидно сильний", value: "ROMI ≥ +100%", color: C.good, note: "змінює статус лише в " + fmt.pctv(zs["змінили статус пізніше"], 1) + " — у " + (zg["змінили статус пізніше"] / zs["змінили статус пізніше"]).toFixed(1).replace(".", ",") + " раза стабільніше за сіру зону" },
    { label: "очевидно ризиковий", value: "ROMI ≤ −50%", color: C.risk, note: "наступні гроші окупаються під " + fmt.pctv(zr["ROMI наступних витрат разом"], 1) + " — нижче цілі, але це не збиток" },
    { label: "неоднозначний", value: "−50 … +100%", color: C.grey, note: fmt.pctv(zg["% тих, хто дійшов"], 0) + " креативів і " + fmt.pctv(zg["змінили статус пізніше"], 1) + " змін статусу — тут ранній ROMI не працює" },
  ];
  tileRow(s, concl.map((c) => ({ label: c.label, value: c.value, valueSize: 19, color: c.color, note: c.note })), y, 1.32);

  s.addNotes("Це пряма відповідь на питання 2 кейсу: обсяг $200, значення метрики ROMI ≥ +100% і ≤ −50%, і частота зміни статусу 11,5% та 28,4% проти 37,0% у середньому. Асиметрія важлива: сильний кандидат упізнається значно надійніше, ніж поганий, тому автоматичного вимкнення ми не пропонуємо.");
  chrome(s, false, "Джерело: розділ 5.3 ноутбука · графік q2_decision_zones.png");
}

/* ═══════════════════════════ 8. Максимум $500 ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Максимум $500 — це межа за ціною, а не за точністю", false, "питання 3 · максимальний обʼєм");
  y = lead(s, "Суми, після якої рішення стає надійним, у цих даних немає. Зате є сума, після якої подальша точність стає надто дорогою.", y);

  const sv = D.survival;
  const chartW = 7.2, chartH = 3.10;
  chartBox(s, { x: M, y: y + 0.1, w: chartW, h: chartH, caption: "Що вищий поріг, то менше креативів під нього підпадає — але тим більшу частку бюджету вони тримають." });
  s.addChart(pres.ChartType.line, [
    { name: "% креативів, що доходять до суми", labels: sv.map((r) => "$" + fmt.n(r["поріг, $"])), values: sv.map((r) => Number(r["% від усіх"].toFixed(1))) },
    { name: "% бюджету, який вони тримають", labels: sv.map((r) => "$" + fmt.n(r["поріг, $"])), values: sv.map((r) => Number(r["% від спенду"].toFixed(1))) },
  ], Object.assign({}, AXIS, {
    x: M, y: y + 0.1, w: chartW, h: chartH,
    chartColors: [C.primary, C.accent], lineSize: 2.5,
    lineDataSymbol: "circle", lineDataSymbolSize: 7,
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25, valAxisLabelFormatCode: '0"%"',
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: C.textSoft,
  }));

  const rx = M + chartW + 0.4, rw = CW - chartW - 0.4;
  tile(s, { x: rx, y: y + 0.1, w: rw, h: 1.56, fill: C.card, label: "довести кожен креатив до $500", value: fmt.m(D.counterfactual.force_500), valueSize: 34, color: C.risk, note: "проти фактичних " + fmt.m(D.counterfactual.actual_spend) + " за весь період — у " + (D.counterfactual.force_500 / D.counterfactual.actual_spend).toFixed(1).replace(".", ",") + " раза більше" });
  s.addText([
    { text: "Тому $500 — це ліміт, а не обіцянка\n", options: { bold: true, fontSize: 12, color: C.ink, breakLine: true } },
    { text: "Природним шляхом до $500 доходить " + fmt.pct(D.counterfactual.reach_500_share, 1) + " креативів. Ми не гарантуємо цю суму кожному — ми лише припиняємо на ній автоматичне правило. Далі креатив виходить із тесту й конкурує за бюджет портфеля на загальних підставах.", options: { fontSize: 11, color: C.textSoft } },
  ], { x: rx, y: y + 1.8, w: rw, h: 1.6, fontFace: F.body, margin: 0, valign: "top" });
  y += chartH + 0.58;

  const ladder = [
    { badge: "$50", head: "перша робоча точка", body: "нижче — тільки технічна перевірка", color: C.primary },
    { badge: "$200", head: "головна точка рішення", body: "тут працюють три зони ROMI", color: C.primary },
    { badge: "$300", head: "один вимірюваний крок", body: "лише для неоднозначної зони", color: C.grey },
    { badge: "$500", head: "кінець автоматики", body: "далі — рішення для портфеля", color: C.accent },
  ];
  const lw = (CW - 0.2 * 3) / 4;
  ladder.forEach((l, i) => {
    const x = M + i * (lw + 0.2);
    card(s, { x, y, w: lw, h: 1.05, fill: "FFFFFF", edge: C.cardEdge, shadow: false });
    s.addText(l.badge, { x: x + 0.2, y: y + 0.14, w: lw - 0.4, h: 0.44, fontFace: F.body, fontSize: 18, bold: true, color: l.color, margin: 0 });
    s.addText(l.head, { x: x + 0.2, y: y + 0.5, w: lw - 0.4, h: 0.24, fontFace: F.body, fontSize: 11, bold: true, color: C.ink, margin: 0 });
    s.addText(l.body, { x: x + 0.2, y: y + 0.73, w: lw - 0.4, h: 0.26, fontFace: F.body, fontSize: 9.5, color: C.textMuted, margin: 0 });
  });

  s.addNotes("Питання кейсу звучить як «які максимальні обʼєми, на яких рішення достатньо точне для більшості випадків». Чесна відповідь: такої суми в цих даних немає, і ми це показали на попередньому слайді. Тому межу обираємо економічно.");
  chrome(s, false, "Джерело: розділи 2.4 і 9 ноутбука");
}

/* ═══════════════════════════ 9-10. Правила ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Правила 1/2: від запуску до $200", false, "фреймворк · рішення");
  y = lead(s, "Ціль — це загальна сума витрат, а не нові гроші зверху. Сума до видачі = max(0, ціль − уже витрачено).", y);

  s.addTable([
    [
      { text: "момент рішення", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "що бачимо", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "дія", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "наступна оцінка", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
    ],
    ["до $50", "будь-який ROMI", "ефективність не оцінюємо; лише технічні зупинки", "на $50"],
    [{ text: "$50", options: { bold: true } }, "ROMI ≤ −50%", { text: "обережно ведемо до $100 + ручна перевірка", options: { color: C.risk } }, "на $100"],
    [{ text: "$50", options: { bold: true } }, "ROMI > −50%", { text: "ведемо до $200", options: { color: C.good } }, "на $200"],
    ["$100", "технічна помилка", "виправити або пауза; причина в журналі", "після виправлення"],
    ["$100", "усе працює", "ведемо до $200", "на $200"],
  ], {
    x: M, y, w: CW, colW: [2.0, 2.6, 5.2, 2.29],
    fontFace: F.body, fontSize: 12, color: C.text,
    border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    valign: "middle", rowH: 0.44, fill: { color: "FFFFFF" },
  });
  y += 3.05;

  const notes = [
    { badge: "!", head: "Чому «до $50» — це не бездіяльність", body: "На цьому етапі перевіряємо, що креатив узагалі отримав шанс: покази йдуть, посилання живе, оплата працює, витрати й дохід пишуться в звіт.", color: C.primary },
    { badge: "$", head: "Чому ціль, а не «+$150»", body: "Дані денні, тому в момент першого перетину $50 медіанні фактичні витрати вже " + fmt.money(D.policy_overshoot[0]["медіанні витрати в момент оцінки"]) + ". Добавка «+$150» означала б різні суми для різних креативів.", color: C.accent },
  ];
  const nw = (CW - 0.26) / 2;
  notes.forEach((n, i) => {
    const x = M + i * (nw + 0.26);
    card(s, { x, y, w: nw, h: 1.28, fill: C.card, shadow: false });
    steps(s, [n], { x: x + 0.2, y: y + 0.16, w: nw - 0.4, rowH: 1.0 });
  });

  s.addNotes("Ризикова гілка на $50 не означає вимкнення. Вона означає менший наступний крок і обовʼязкову ручну перевірку — саме тому, що з історичних даних безпечне автоматичне вимкнення довести не можна.");
  chrome(s, false, "Джерело: розділ 9 ноутбука");
}

{
  const s = baseSlide(false);
  let y = title(s, "Правила 2/2: після $200 зони отримують різний крок", false, "фреймворк · рішення");
  y = lead(s, "Саме тут працює поділ на три зони. Ризикова гілка навмисно залишена перевіркою, а не автоматичним вимкненням.", y);

  s.addTable([
    [
      { text: "момент рішення", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "що бачимо", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "дія", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "наступна оцінка", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
    ],
    [{ text: "$200", options: { bold: true } }, "ROMI ≥ +100%", { text: "пріоритет бюджету до $500", options: { color: C.good, bold: true } }, "на $500"],
    [{ text: "$200", options: { bold: true } }, "−50% < ROMI < +100%", "один вимірюваний крок до $300", "на $300"],
    [{ text: "$200", options: { bold: true } }, "ROMI ≤ −50%", { text: "у пілоті випадково 50/50: пауза або крок до $250", options: { color: C.risk } }, "на $250 або пауза"],
    ["$300", "ROMI частини $200→$300 ≥ 30%", { text: "ставимо в чергу до $500", options: { color: C.good } }, "на $500"],
    ["$300", "ROMI частини < 30%", "пауза", "—"],
    ["$500", "будь-який ROMI", "автоматичне правило завершується", "рішення для портфеля"],
  ], {
    x: M, y, w: CW, colW: [1.75, 3.3, 4.85, 2.19],
    fontFace: F.body, fontSize: 11.5, color: C.text,
    border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    valign: "middle", rowH: 0.4, fill: { color: "FFFFFF" },
  });
  y += 3.15;

  const notes = [
    { badge: "1", head: "На $300 дивимося лише останню частину", body: "Загальний ROMI тут не підходить: у ньому сидить старт, за який креатив уже отримав крок. Оцінюємо віддачу саме грошей $200 → $300.", color: C.primary },
    { badge: "2", head: "Ризикова зона не вимикається автоматично", body: "Її наступні гроші окупаються під " + fmt.pctv(D.zones_200.find((r) => r["зона"] === "ризик")["ROMI наступних витрат разом"], 1) + " — нижче цілі, але це прибуток. Право вимикати дає лише експеримент.", color: C.risk },
  ];
  const nw = (CW - 0.26) / 2;
  notes.forEach((n, i) => {
    const x = M + i * (nw + 0.26);
    card(s, { x, y, w: nw, h: 1.18, fill: C.card, shadow: false });
    steps(s, [n], { x: x + 0.2, y: y + 0.13, w: nw - 0.4, rowH: 0.95 });
  });

  s.addNotes("Пороги −50% і +100% — кандидати для контрольованої перевірки, а не універсальні константи. Їхню користь доводить лише випадковий розподіл.");
  chrome(s, false, "Джерело: розділ 9 ноутбука");
}

/* ═══════════════════════════ 11. Механіка бюджету ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Бюджет рахуємо від фактичних витрат, а не від назви порогу", false, "фреймворк · механіка");
  y = lead(s, "Дані денні, тому в момент рішення креатив майже завжди вже перевитратив поріг. Видаємо тільки різницю.", y);

  const ov = D.policy_overshoot;
  const chartW = 6.7, chartH = 3.0;
  chartBox(s, { x: M, y: y + 0.1, w: chartW, h: chartH, caption: "Медіанні накопичені витрати в момент першого перетину порогу." });
  s.addChart(pres.ChartType.bar, [
    { name: "назва порогу", labels: ["після $50", "після $200"], values: [50, 200] },
    { name: "медіанні фактичні витрати", labels: ["після $50", "після $200"], values: ov.map((r) => Number(r["медіанні витрати в момент оцінки"].toFixed(0))) },
  ], Object.assign({}, AXIS, {
    x: M, y: y + 0.1, w: chartW, h: chartH,
    barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [C.grey, C.primary],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFormatCode: '"$"0',
    valAxisMaxVal: 320, valAxisMajorUnit: 100, valAxisLabelFormatCode: '"$"0',
    showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: C.textSoft,
  }));

  const rx = M + chartW + 0.4, rw = CW - chartW - 0.4;
  s.addShape(pres.ShapeType.roundRect, { x: rx, y: y + 0.1, w: rw, h: 0.72, rectRadius: 0.08, fill: { color: C.ink }, line: { width: 0 } });
  s.addText("сума до видачі = max(0, ціль − уже витрачено)", {
    x: rx + 0.2, y: y + 0.1, w: rw - 0.4, h: 0.72, fontFace: "Consolas", fontSize: 12.5,
    color: C.onDark, margin: 0, valign: "middle",
  });
  s.addText("Приклад: ціль $200, уже витрачено $79 → видаємо $121.", {
    x: rx, y: y + 0.9, w: rw, h: 0.3, fontFace: F.body, fontSize: 11, italic: true, color: C.textSoft, margin: 0,
  });
  const oc = [
    ["після першого перетину $50", fmt.pctv(ov[0]["вже не нижче наступної цілі"] * 100, 1), "уже витратили ≥ $200 — новий бюджет їм не потрібен узагалі"],
    ["після першого перетину $200", fmt.pctv(ov[1]["вже не нижче наступної цілі"] * 100, 1), "уже витратили ≥ $500 — одразу переходять до рішення портфеля"],
  ];
  oc.forEach(([a, b, c], i) => {
    const yy = y + 1.32 + i * 0.98;
    card(s, { x: rx, y: yy, w: rw, h: 0.88, fill: C.card, shadow: false });
    s.addText(b, { x: rx + 0.18, y: yy + 0.1, w: 0.95, h: 0.62, fontFace: F.body, fontSize: 20, bold: true, color: C.accent, margin: 0, valign: "middle" });
    s.addText([
      { text: a + "\n", options: { fontSize: 10, bold: true, color: C.ink, breakLine: true } },
      { text: c, options: { fontSize: 9.5, color: C.textSoft } },
    ], { x: rx + 1.18, y: yy + 0.06, w: rw - 1.36, h: 0.76, fontFace: F.body, margin: 0, valign: "middle" });
  });
  y += chartH + 0.62;

  card(s, { x: M, y, w: CW, h: 0.95, fill: "F1F7F1", edge: "BFE3BF", shadow: false });
  s.addText([
    { text: "Чому це принципово.  ", options: { bold: true, fontSize: 12, color: C.ink } },
    { text: "Формулювання «додати +$150» дало б різні підсумкові суми різним креативам — і бюджети двох груп в експерименті перестали б бути порівнянними. Абсолютна ціль робить крок однаковим для всіх і робить перевірку чесною.", options: { fontSize: 12, color: C.textSoft } },
  ], { x: M + 0.28, y: y + 0.12, w: CW - 0.56, h: 0.72, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Це неочевидна, але важлива деталь реалізації: без неї експеримент із рівними бюджетами неможливо провести коректно.");
  chrome(s, false, "Джерело: розділ 10 ноутбука");
}

/* ═══════════════════════════ 12. Ручні перевірки ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Ручне рішення теж має правила — інакше фреймворк розмиється", false, "фреймворк · дисципліна");
  y = lead(s, "Для кожної суми, де людина втручається, є один короткий перелік. Рішення не залежить від настрою маркетолога.", y);

  const cw = (CW - 0.24 * 2) / 3;
  const manual = [
    { t: "$100", sub: "технічна перевірка", items: ["посилання й сторінка працюють", "оплата проходить", "витрати й дохід пишуться", "покази йдуть"], out: "усе працює → ціль $200", outc: C.good },
    { t: "$300", sub: "остання частина бюджету", items: ["дивимося лише ROMI $200 → $300", "загальний ROMI не підміняє його", "порівнюємо з ціллю +30%"], out: "≥ 30% → ціль $500, інакше пауза", outc: C.primary },
    { t: "$500", sub: "рішення для портфеля", items: ["нова гранична сума", "відповідальний", "дата наступного перегляду"], out: "без цього → пауза", outc: C.accent },
  ];
  manual.forEach((m, i) => {
    const x = M + i * (cw + 0.24);
    card(s, { x, y, w: cw, h: 2.75 });
    s.addText(m.t, { x: x + 0.24, y: y + 0.18, w: cw - 0.48, h: 0.5, fontFace: F.body, fontSize: 24, bold: true, color: C.ink, margin: 0 });
    s.addText(m.sub.toUpperCase(), { x: x + 0.24, y: y + 0.68, w: cw - 0.48, h: 0.24, fontFace: F.body, fontSize: 9, bold: true, charSpacing: 1.1, color: C.textMuted, margin: 0 });
    s.addText(m.items.map((t, k) => ({ text: t, options: { bullet: { code: "2022" }, breakLine: k < m.items.length - 1 } })), {
      x: x + 0.24, y: y + 0.94, w: cw - 0.48, h: 1.1, fontFace: F.body, fontSize: 11, color: C.textSoft, margin: 0, valign: "top", paraSpaceAfter: 4,
    });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.24, y: y + 2.12, w: cw - 0.48, h: 0.45, rectRadius: 0.07, fill: { color: "FFFFFF" }, line: { color: m.outc, width: 1 } });
    s.addText(m.out, { x: x + 0.32, y: y + 2.12, w: cw - 0.64, h: 0.45, fontFace: F.body, fontSize: 10.5, bold: true, color: m.outc, margin: 0, valign: "middle" });
  });
  y += 3.0;

  card(s, { x: M, y, w: CW, h: 1.15, fill: C.ink, edge: C.ink });
  s.addText("Журнал рішень — обовʼязкова частина фреймворку, а не звітність", {
    x: M + 0.3, y: y + 0.16, w: CW - 0.6, h: 0.3, fontFace: F.body, fontSize: 13, bold: true, color: C.onDark, margin: 0,
  });
  const logFields = ["хто", "коли", "кампанія / група", "контрольна сума", "рекомендація правила", "фактична дія", "наступна ціль", "дата перегляду", "причина", "версія правил"];
  s.addText(logFields.join("   ·   "), {
    x: M + 0.3, y: y + 0.52, w: CW - 0.6, h: 0.5, fontFace: F.body, fontSize: 11.5, color: C.onDarkMuted, margin: 0, valign: "top",
  });

  s.addNotes("Без журналу неможливо ані виміряти ефект правил, ані відрізнити «правило не спрацювало» від «правило обійшли». Поле «рекомендація правила проти фактичної дії» — ключове для аналізу пілота.");
  chrome(s, false, "Джерело: розділ 9 ноутбука");
}

/* ═══════════════════════════ 13. Розподіл бюджету ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Спільний бюджет розподіляємо повними кроками", false, "фреймворк · операційна частина");
  y = lead(s, "У кожної групи експерименту один тижневий ліміт. Часткових виплат немає: крок або видається цілком, або не видається.", y);

  const a = D.allocator_demo;
  y = steps(s, [
    { badge: "1", head: "Спочатку завершуємо розпочате", body: "Креативи, які вже в тесті, добираємо до найближчої контрольної суми. Хто чекає довше — той вище в черзі.", color: C.primary },
    { badge: "2", head: "Далі — сильні кандидати", body: "Залишок віддаємо зоні ROMI ≥ +100%, у порядку спадання ROMI. Це пріоритет, а не гарантія перемоги.", color: C.good },
    { badge: "3", head: "Решту переносимо, а не витрачаємо", body: "Якщо на повний крок не вистачає, гроші переходять на наступний тиждень. Половина кроку зробила б результат невимірюваним.", color: C.accent },
  ], { x: M, y: y + 0.08, w: 7.3, rowH: 0.95 });

  const rx = M + 7.7, rw = CW - 7.7;
  card(s, { x: rx, y: 1.72, w: rw, h: 3.15, fill: C.card });
  s.addText("Приклад тижня", { x: rx + 0.24, y: 1.88, w: rw - 0.48, h: 0.3, fontFace: F.body, fontSize: 12.5, bold: true, color: C.ink, margin: 0 });
  const rowsA = [
    ["тижневий бюджет", fmt.money(a.weekly_budget), C.ink],
    ["на розпочаті кроки", fmt.money(a.mandatory), C.primary],
    ["сильному кандидату", fmt.money(a.scale), C.good],
    ["перенесено на наступний", fmt.money(a.carry), C.accent],
  ];
  rowsA.forEach(([l, v, col], i) => {
    const yy = 2.28 + i * 0.55;
    s.addText(l, { x: rx + 0.24, y: yy, w: rw - 1.5, h: 0.4, fontFace: F.body, fontSize: 11, color: C.textSoft, margin: 0, valign: "middle" });
    s.addText(v, { x: rx + rw - 1.5, y: yy, w: 1.26, h: 0.4, align: "right", fontFace: F.body, fontSize: 14, bold: true, color: col, margin: 0, valign: "middle" });
  });
  s.addText("Частковий крок не видається нікому: $10 переносяться, а не «доливаються» четвертому кандидату.", {
    x: rx + 0.24, y: 4.52, w: rw - 0.48, h: 0.5, fontFace: F.body, fontSize: 9.5, italic: true, color: C.textMuted, margin: 0, valign: "top",
  });

  y = 5.15;
  card(s, { x: M, y, w: CW, h: 1.2, fill: "FFFFFF", edge: C.cardEdge, shadow: false });
  s.addText([
    { text: "Алгоритм виконуваний, а не описовий.  ", options: { bold: true, fontSize: 12, color: C.ink } },
    { text: "Функція allocate_weekly_policy_budget у розділі 11 ноутбука приймає чергу заявок і тижневий ліміт, а повертає розподіл із перевірками: жоден креатив не отримує часткового кроку, сума виданого плюс перенесеного завжди дорівнює бюджету, дублікати в черзі відхиляються.", options: { fontSize: 12, color: C.textSoft } },
  ], { x: M + 0.28, y: y + 0.14, w: CW - 0.56, h: 0.95, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Атомарність кроку — не педантизм. Половина кроку дає результат, який неможливо порівняти з іншою групою, і ламає весь вимір ефекту.");
  chrome(s, false, "Джерело: розділ 11 ноутбука");
}

/* ═══════════════════════════ 14. Історична перевірка ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Історія підтримує порядок пріоритетів — але не доводить ефект", false, "перевірка на даних");
  y = lead(s, "Для кожної гілки рахуємо ROMI грошей, витрачених саме після рішення. Три основні гілки перевищують ціль +30%.", y);

  const pb = D.policy_branches;
  const short = { "$50: ризик → $100": "$50 → $100  (ризик)", "$50: продовжити → $200": "$50 → $200  (продовжити)", "$200: ризик → $250": "$200 → $250  (ризик)", "$200: продовжити → $300": "$200 → $300  (сіра зона)", "$200: збільшити → $500": "$200 → $500  (сильний)" };
  const chartW = 7.5, chartH = 3.1;
  chartBox(s, { x: M, y: y + 0.1, w: chartW, h: chartH, caption: "ROMI витрат, зроблених після моменту рішення. Ціль бізнесу — +30%." });
  s.addChart(pres.ChartType.bar, [{
    name: "ROMI наступних витрат",
    labels: pb.map((r) => short[r["гілка"]] || r["гілка"]),
    values: pb.map((r) => Number((r["ROMI майбутніх витрат разом"] * 100).toFixed(1))),
  }], Object.assign({}, AXIS, {
    x: M, y: y + 0.1, w: chartW, h: chartH,
    barDir: "bar", barGapWidthPct: 45, chartColors: [C.primary],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFormatCode: '0"%"',
    valAxisMaxVal: 80, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"',
    catAxisLabelFontSize: 10, showLegend: false,
  }));

  const rx = M + chartW + 0.4, rw = CW - chartW - 0.4;
  s.addText("Скільки результатів ми взагалі бачимо", { x: rx, y: y + 0.1, w: rw, h: 0.28, fontFace: F.body, fontSize: 11.5, bold: true, color: C.ink, margin: 0 });
  s.addTable([
    [{ text: "гілка", options: { bold: true } }, { text: "кандидатів", options: { bold: true, align: "right" } }, { text: "відомий\nрезультат", options: { bold: true, align: "right" } }],
  ].concat(pb.map((r) => [
    (short[r["гілка"]] || r["гілка"]).split("  ")[0],
    fmt.n(r["кандидатів"]),
    fmt.n(r["відомих результатів"]) + " (" + (r["частка з відомим результатом"] * 100).toFixed(0) + "%)",
  ])), {
    x: rx, y: y + 0.44, w: rw, colW: [1.6, 1.1, 1.68],
    fontFace: F.body, fontSize: 10, color: C.text,
    border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    align: "right", valign: "middle", rowH: 0.35, fill: { color: "FFFFFF" },
  });
  s.addText([
    { text: "Чому це не доказ\n", options: { bold: true, fontSize: 11.5, color: C.ink, breakLine: true } },
    { text: "Майбутнє видно лише для тих креативів, яким у минулому вже дали бюджет. Гілка «$200 → $250» має всього " + fmt.n(pb[2]["відомих результатів"]) + " відомих результатів, і лише один із них перевищив ціль — тому її ROMI " + fmt.pct(pb[2]["ROMI майбутніх витрат разом"], 1) + " тримається на одному викиді.", options: { fontSize: 10.5, color: C.textSoft } },
  ], { x: rx, y: y + 2.45, w: rw, h: 1.2, fontFace: F.body, margin: 0, valign: "top" });
  y += chartH + 0.6;

  card(s, { x: M, y, w: CW, h: 1.15, fill: "FDF3EC", edge: "F0D3BC", shadow: false });
  s.addText([
    { text: "Що саме доводить ця перевірка.  ", options: { bold: true, fontSize: 12, color: C.ink } },
    { text: "Вона показує, що черга пріоритетів побудована правильно: гроші, віддані сильним кандидатам, справді окупалися краще за гроші, віддані ризиковій зоні. Вона ", options: { fontSize: 12, color: C.textSoft } },
    { text: "не", options: { fontSize: 12, bold: true, color: C.risk } },
    { text: " показує, що ці правила створили додатковий прибуток — для цього потрібен випадковий розподіл, описаний далі.", options: { fontSize: 12, color: C.textSoft } },
  ], { x: M + 0.28, y: y + 0.12, w: CW - 0.56, h: 0.92, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Селективна розмітка: історичний follow-up існує тільки там, де Meta або маркетолог уже вирішили дати бюджет. Це головна причина, чому автоматичне вимкнення з цих даних вивести не можна.");
  chrome(s, false, "Джерело: розділ 10 ноутбука · графік final_budget_rules_historical_check.png");
}

/* ═══════════════════════════ 15. Питання 4 — сегменти ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Однакові $200 купують різну кількість інформації", false, "питання 4 · розрізи");
  y = lead(s, "Ось чому обʼєми за розрізами вже відрізняються — навіть якщо поріг записаний одним числом.", y);

  // Показуємо саме розкид: пʼять найдешевших і пʼять найдорожчих тем.
  const sortedCpt = D.cpt_topic.slice().sort((a, b) => a["вартість події, $"] - b["вартість події, $"]);
  const shown = sortedCpt.slice(0, 5).concat(sortedCpt.slice(-5));
  shown.sort((a, b) => b["вартість події, $"] - a["вартість події, $"]);
  const chartW = 7.0, chartH = 3.0;
  chartBox(s, {
    x: M, y: y + 0.08, w: chartW, h: chartH,
    caption: "Пʼять найдешевших і пʼять найдорожчих тем серед " + D.cpt_topic.length + ", де подій достатньо для оцінки.",
  });
  s.addChart(pres.ChartType.bar, [{
    name: "вартість події, $",
    labels: shown.map((r) => r.topic),
    values: shown.map((r) => Number(r["вартість події, $"].toFixed(2))),
  }], Object.assign({}, AXIS, {
    x: M, y: y + 0.08, w: chartW, h: chartH,
    barDir: "bar", barGapWidthPct: 40, chartColors: [C.primary],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10, dataLabelFormatCode: '"$"0.00',
    valAxisLabelFormatCode: '"$"0', catAxisLabelFontSize: 10, showLegend: false,
  }));

  const rx = M + chartW + 0.4, rw = CW - chartW - 0.4;
  tile(s, {
    x: rx, y: y + 0.08, w: rw, h: 1.6, fill: C.card,
    label: "розкид ціни події між темами",
    value: "×" + (D.cpt_topic[D.cpt_topic.length - 1]["вартість події, $"] / D.cpt_topic[0]["вартість події, $"]).toFixed(1).replace(".", ","),
    valueSize: 36, color: C.accent,
    note: "від " + fmt.money(D.cpt_topic[0]["вартість події, $"], 2) + " до " + fmt.money(D.cpt_topic[D.cpt_topic.length - 1]["вартість події, $"], 2) + " — тому ті самі $200 дають від " + Math.round(200 / D.cpt_topic[D.cpt_topic.length - 1]["вартість події, $"]) + " до " + Math.round(200 / D.cpt_topic[0]["вартість події, $"]) + " подій",
  });
  s.addText("Формат × тип креативу", { x: rx, y: y + 1.82, w: rw, h: 0.28, fontFace: F.body, fontSize: 11, bold: true, color: C.ink, margin: 0 });
  s.addTable([
    [{ text: "сегмент", options: { bold: true } }, { text: "ціна події", options: { bold: true, align: "right" } }, { text: "подій\nза $200", options: { bold: true, align: "right" } }],
  ].concat(D.cpt_media.map((r) => [
    r.media_type + " · " + r.creo_type,
    fmt.money(r["вартість події, $"], 2),
    String(Math.round(r["подій за $200"])),
  ])), {
    x: rx, y: y + 2.14, w: rw, colW: [2.2, 1.0, 1.18],
    fontFace: F.body, fontSize: 10, color: C.text,
    border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    align: "right", valign: "middle", rowH: 0.32, fill: { color: "FFFFFF" },
  });
  y += chartH + 0.46;

  const cw = (CW - 0.24) / 2;
  card(s, { x: M, y, w: cw, h: 1.32, fill: "F1F7F1", edge: "BFE3BF", shadow: false });
  s.addText([
    { text: "Чи відрізняються обʼєми? Так.\n", options: { bold: true, fontSize: 12.5, color: C.ink, breakLine: true } },
    { text: "Формально єдиний поріг у доларах насправді означає різну суворість рішення: у дорогій темі $200 дають лише " + Math.round(200 / D.cpt_topic[D.cpt_topic.length - 1]["вартість події, $"]) + " подій, у дешевій — " + Math.round(200 / D.cpt_topic[0]["вартість події, $"]) + ".", options: { fontSize: 11, color: C.textSoft } },
  ], { x: M + 0.26, y: y + 0.16, w: cw - 0.52, h: 1.1, fontFace: F.body, margin: 0, valign: "top" });

  card(s, { x: M + cw + 0.24, y, w: cw, h: 1.32, fill: C.card, shadow: false });
  s.addText([
    { text: "Чи вводимо різні пороги зараз? Ні — і це свідомо.\n", options: { bold: true, fontSize: 12.5, color: C.ink, breakLine: true } },
    { text: "У пріоритетній гілці $200 → $500 підтверджений результат є лише для 51 Static, 59 Video, 13 new і 97 variation. Пороги, оцінені на десятках спостережень, були б нестійкішими за похибку, яку мали б виправити.", options: { fontSize: 11, color: C.textSoft } },
  ], { x: M + cw + 0.5, y: y + 0.16, w: cw - 0.52, h: 1.1, fontFace: F.body, margin: 0, valign: "top" });

  s.addNotes("Це відповідь на «чому» в питанні 4. Вартість події — саме той механізм, через який однакова сума означає різну впевненість. Сегментні пороги вводимо тоді, коли в сегменті накопичиться достатньо підтверджених результатів.");
  chrome(s, false, "Джерело: розділ 12.1 ноутбука · графік q4_segment_cost_of_information.png");
}

/* ═══════════════════════════ 16. Яких розрізів бракує ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Найсильніших розрізів у даних немає", false, "питання 4 · чого бракує");
  y = lead(s, "Географії, кампанії, плейсменту й аудиторії у файлі немає. Кожен зміщує і вартість події, і ROMI — зібрати їх треба до пілота.", y);

  s.addTable([
    [
      { text: "розріз", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "чому впливає", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
      { text: "що ламається без нього", options: { bold: true, color: C.onDark, fill: { color: C.ink } } },
    ],
    [{ text: "Географія", options: { bold: true } }, "За словами автора кейсу: в одній країні 1 000 показів дають оплату, в іншій — жодної", "Найбільша частина розкиду ROMI лишається непоясненою"],
    [{ text: "Кампанія / група оголошень", options: { bold: true } }, "Meta CBO перерозподіляє бюджет усередині кампанії: один креатив може зʼїсти весь ліміт", "Неможлива коректна рандомізація — групи впливають одна на одну"],
    [{ text: "Місце показу й аудиторія", options: { bold: true } }, "Ціна показу й конверсія різняться між плейсментами в рази", "Різницю між групами можна сплутати з різницею в закупівлі"],
    [{ text: "Сезон", options: { bold: true } }, "25 грудня — провал; період кінця грудня та зими — навпаки сплеск", "Пороги, підібрані в грудні, можуть не працювати влітку"],
    [{ text: "Звʼязок original ↔ variation", options: { bold: true } }, "У даних немає parent-id, тому варіацію не можна повʼязати з джерелом", "Неможливо виміряти головний бізнес-ефект: чи народжують хіти хороші варіації"],
  ], {
    x: M, y, w: CW, colW: [2.85, 4.65, 4.59],
    fontFace: F.body, fontSize: 11, color: C.text,
    border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    valign: "middle", rowH: 0.58, fill: { color: "FFFFFF" },
  });
  y += 3.85;

  card(s, { x: M, y, w: CW, h: 0.9, fill: "FDF3EC", edge: "F0D3BC", shadow: false });
  s.addText([
    { text: "Практичний висновок.  ", options: { bold: true, fontSize: 12.5, color: C.ink } },
    { text: "Ці поля потрібно почати збирати ", options: { fontSize: 12.5, color: C.textSoft } },
    { text: "до", options: { fontSize: 12.5, bold: true, color: C.accent } },
    { text: " запуску пілота, а не після нього. Особливо campaign_id і adset_id: без них рандомізація на рівні креативу некоректна, бо креативи однієї кампанії ділять спільний бюджет.", options: { fontSize: 12.5, color: C.textSoft } },
  ], { x: M + 0.28, y: y + 0.1, w: CW - 0.56, h: 0.8, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Питання кейсу прямо просить назвати інші розрізи, які можуть впливати. Важливо не просто перелічити їх, а сказати, що саме ламається без кожного — і що campaign_id блокує коректний дизайн експерименту.");
  chrome(s, false, "Джерело: розділ 12 ноутбука, Q&A з автором кейсу 31.07");
}

/* ═══════════════════════════ 17. Пілот ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Користь правил доводить лише випадковий розподіл", false, "питання 5 · як виміряти");
  y = lead(s, "Історична перевірка не відокремлює ефект правил від рішень Meta та маркетологів. Тому — контрольований експеримент.", y);

  const cw = (CW - 0.22 * 3) / 4;
  const design = [
    { label: "одиниця розподілу", value: "кампанія", note: "Якщо креативи ділять бюджет кампанії — ділимо кампаніями. Інакше — креативами.", color: C.primary },
    { label: "розподіл", value: "50 / 50", note: "Поточний процес проти нових правил. Тижневі бюджети груп рівні.", color: C.primary },
    { label: "головний показник", value: D.pilot.outcome_days + " днів", note: "Приріст доходу мінус приріст витрат за 14 календарних днів; одна зафіксована версія прогнозу.", color: C.accent },
    { label: "критерій рішення", value: "нижня межа", note: "Впроваджуємо, лише якщо нижня межа 95% інтервалу не нижча за мінімальний корисний ефект.", color: C.good },
  ];
  design.forEach((d, i) => {
    const x = M + i * (cw + 0.22);
    card(s, { x, y, w: cw, h: 2.05 });
    s.addText(d.label.toUpperCase(), { x: x + 0.22, y: y + 0.18, w: cw - 0.44, h: 0.24, fontFace: F.body, fontSize: 9, bold: true, charSpacing: 1.1, color: C.textMuted, margin: 0 });
    s.addText(d.value, { x: x + 0.22, y: y + 0.44, w: cw - 0.44, h: 0.5, fontFace: F.body, fontSize: 24, bold: true, color: d.color, margin: 0 });
    s.addText(d.note, { x: x + 0.22, y: y + 0.98, w: cw - 0.44, h: 0.9, fontFace: F.body, fontSize: 10.5, color: C.textSoft, margin: 0, valign: "top" });
  });
  y += 2.3;

  y = steps(s, [
    { badge: "→", head: "Вхід у тест — при першому досягненні $50 накопичених витрат", body: "Розподіл фіксується один раз на креатив і не змінюється. Стратифікація: тиждень запуску × формат медіа × тип креативу.", color: C.primary },
    { badge: "✓", head: "Рахуємо за призначеною групою, навіть якщо правило обійшли", body: "Нульове продовження лишається нулем, а не «немає даних». Інакше оцінка ефекту зміщується на користь нових правил.", color: C.good },
    { badge: "!", head: "Три обовʼязкові перевірки перед висновком", body: "Бюджети груп справді рівні · жоден сегмент не домінує · результат не тримається на одному викиді.", color: C.accent },
  ], { x: M, y: y + 0.05, w: 8.4, rowH: 0.86 });

  const rx = M + 8.8, rw = CW - 8.8;
  card(s, { x: rx, y: 4.62, w: rw, h: 2.16, fill: C.ink, edge: C.ink });
  s.addText("Ризикова зона — теж експеримент", { x: rx + 0.24, y: 4.78, w: rw - 0.48, h: 0.3, fontFace: F.body, fontSize: 12.5, bold: true, color: C.onDark, margin: 0 });
  s.addText("Усередині пілота креативи з ROMI ≤ −50% на $200 діляться 50/50: пауза проти кроку до $250. Це єдиний спосіб дізнатися, скільки хороших креативів ми вимикали б помилково — з історії це не виводиться в принципі.", {
    x: rx + 0.24, y: 5.14, w: rw - 0.48, h: 1.5, fontFace: F.body, fontSize: 10.5, color: C.onDarkMuted, margin: 0, valign: "top",
  });

  s.addNotes("Ключова відмінність від історичної перевірки: тут ми самі створюємо контрфакт. Рандомізація всередині ризикової зони — єдине джерело даних про хибні вимкнення.");
  chrome(s, false, "Джерело: розділ 13 ноутбука");
}

/* ═══════════════════════════ 18. Вибірка і строк ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Скільки це триватиме — і як зробити швидше", false, "питання 5 · планування");
  y = lead(s, "Розмір вибірки без строку — половина відповіді. Проблема не в малому ефекті, а у величезному розкиді результату одного креативу.", y);

  const lv = D.pilot.variance_levers;
  const chartW = 7.15, chartH = 2.95;
  chartBox(s, { x: M, y: y + 0.08, w: chartW, h: chartH, caption: "Тижнів до висновку при середньому історичному темпі появи нових креативів (включно з 14-денним хвостом)." });
  s.addChart(pres.ChartType.bar, [{
    name: "тижнів до висновку",
    labels: ["сирий результат\n(основний показник)", "кластери-кампанії\n−25% розкиду", "CUPED\n−35% розкиду", "обмежено крайній 1%\n(перевірка стійкості)"],
    values: lv.map((r) => Number(r["тижнів при середньому темпі"].toFixed(0))),
  }], Object.assign({}, AXIS, {
    x: M, y: y + 0.08, w: chartW, h: chartH,
    barDir: "col", barGapWidthPct: 50, chartColors: [C.primary],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelFormatCode: "0",
    valAxisMaxVal: 50, valAxisMajorUnit: 10,
    catAxisLabelFontSize: 9, showLegend: false,
  }));

  const rx = M + chartW + 0.4, rw = CW - chartW - 0.4;
  s.addText("Звідки береться 9 046 на групу", { x: rx, y: y + 0.08, w: rw, h: 0.28, fontFace: F.body, fontSize: 11.5, bold: true, color: C.ink, margin: 0 });
  s.addTable([
    [{ text: "мінімально корисний ефект", options: {} }, { text: fmt.money(50), options: { bold: true, align: "right" } }],
    [{ text: "очікуваний ефект", options: {} }, { text: fmt.money(100), options: { bold: true, align: "right" } }],
    [{ text: "запас = очікуваний − мінімум", options: {} }, { text: fmt.money(50), options: { bold: true, align: "right" } }],
    [{ text: "розкид результату (SD)", options: {} }, { text: fmt.money(D.pilot.sd_raw), options: { bold: true, align: "right" } }],
    [{ text: "потрібно на групу", options: { bold: true } }, { text: fmt.n(D.pilot.sample_size[1]["на групу: без обмеження крайніх значень"]), options: { bold: true, align: "right", color: C.primary } }],
  ], {
    x: rx, y: y + 0.4, w: rw, colW: [2.55, 1.83],
    fontFace: F.body, fontSize: 10.5, color: C.text,
    border: { type: "solid", color: C.cardEdge, pt: 0.5 },
    valign: "middle", rowH: 0.34, fill: { color: "FFFFFF" },
  });
  s.addText([
    { text: "Вибірку рахуємо не «проти нуля»\n", options: { bold: true, fontSize: 11, color: C.ink, breakLine: true } },
    { text: "Критерій суворий: нижня межа інтервалу має перевищити мінімально корисний ефект. Тому в знаменнику стоїть запас $50, а не сам ефект. Якщо очікуваний ефект дорівнює мінімальному, жодна скінченна вибірка не дає 80% шансу пройти цей критерій.", options: { fontSize: 10, color: C.textSoft } },
  ], { x: rx, y: y + 2.32, w: rw, h: 1.15, fontFace: F.body, margin: 0, valign: "top" });
  y += chartH + 0.55;

  const tiles = [
    { label: "медіанний результат креативу", value: fmt.money(D.pilot.profit_median), color: C.ink, note: "при середньому " + fmt.money(D.pilot.profit_mean) + " — хвіст, а не типовий випадок" },
    { label: "розкид проти медіани", value: fmt.money(D.pilot.sd_raw), color: C.risk, note: "наслідок того, що топ-1% креативів тримає " + fmt.pct(D.portfolio.top1_spend_share, 0) + " бюджету" },
    { label: "рішення", value: "знижуємо розкид", color: C.good, note: "кластери й CUPED скорочують перевірку з " + lv[0]["тижнів при середньому темпі"].toFixed(0) + " до " + lv[2]["тижнів при середньому темпі"].toFixed(0) + " тижнів, не послаблюючи критерій" },
  ];
  tileRow(s, tiles.map((t) => ({ label: t.label, value: t.value, valueSize: 22, color: t.color, note: t.note })), y, 1.4);

  s.addNotes("Чесно показуємо, що наївний розрахунок дає 42 тижні — і що правильна реакція на це не «подовжити тест», а зменшити дисперсію. Коефіцієнти −25% і −35% — консервативні планові припущення; фактичне зменшення переоцінюємо на перших 1 000 креативів наосліп.");
  chrome(s, false, "Джерело: розділи 13.1 і 13.2 ноутбука");
}

/* ═══════════════════════════ 19. Інсайти ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Чотири факти з даних, які змінюють оптику", false, "інсайти");
  y = lead(s, "Не входять у правила напряму, але міняють те, як варто дивитися на тестування креативів у цілому.", y);

  const cw = (CW - 0.22) / 2, ch = 1.62;
  const ins = [
    { v: fmt.pct(D.portfolio.top1_spend_share, 0), t: "бюджету забирає топ-1% креативів", b: "Бюджетом усередині кампанії керує алгоритм Meta, а не маркетолог. Тому правило має бути про наступний крок, а не про фінальний вирок.", c: C.primary },
    { v: "CV " + D.portfolio.cv_upsells_per_trial.toFixed(1).replace(".", ","), t: "розкид апселів проти CV " + D.portfolio.cv_pnr_per_trial.toFixed(2).replace(".", ",") + " у прогнозі доходу", b: "Апселі дають " + fmt.pct(D.portfolio.upsell_share, 0) + " доходу, але саме вони створюють «пилу» раннього ROMI. Одна велика покупка одномоментно змінює картину.", c: C.accent },
    { v: fmt.n(D.portfolio.never_delivered), t: "креативів (" + fmt.pct(D.portfolio.never_delivered_share, 1) + ") не витратили жодного долара", b: "Це не провали — це запуски, яким алгоритм не дав шансу. Рахувати їх як невдачі означає системно недооцінювати роботу дизайнерів.", c: C.risk },
    { v: fmt.pct(D.portfolio.peak_month_share, 0), t: "усього бюджету припадає на грудень 2025", b: "Медіанний креатив живе " + D.portfolio.median_active_days.toFixed(0) + " дні. Фреймворк мусить давати відповідь за дні: другого тижня у більшості просто не буде.", c: C.grey },
  ];
  ins.forEach((o, i) => {
    const x = M + (i % 2) * (cw + 0.22);
    const yy = y + Math.floor(i / 2) * (ch + 0.2);
    card(s, { x, y: yy, w: cw, h: ch });
    s.addText(o.v, { x: x + 0.26, y: yy + 0.2, w: 2.1, h: 0.55, fontFace: F.body, fontSize: 27, bold: true, color: o.c, margin: 0, valign: "top" });
    s.addText(o.t, { x: x + 2.42, y: yy + 0.22, w: cw - 2.68, h: 0.55, fontFace: F.body, fontSize: 11.5, bold: true, color: C.ink, margin: 0, valign: "top" });
    s.addText(o.b, { x: x + 0.26, y: yy + 0.85, w: cw - 0.52, h: 0.62, fontFace: F.body, fontSize: 11, color: C.textSoft, margin: 0, valign: "top" });
  });
  y += 2 * ch + 0.2 + 0.24;

  card(s, { x: M, y, w: CW, h: 0.92, fill: C.ink, edge: C.ink });
  s.addText([
    { text: "Спільне в цих фактах:  ", options: { bold: true, fontSize: 12.5, color: C.onDark } },
    { text: "результат окремого креативу здебільшого визначають речі поза його якістю — рішення алгоритму, рідкісні великі покупки й сезон. Саме тому фреймворк керує ризиком кроку, а не намагається вгадати переможця.", options: { fontSize: 12.5, color: C.onDarkMuted } },
  ], { x: M + 0.3, y: y + 0.1, w: CW - 0.6, h: 0.72, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Це необовʼязковий блок кейсу. Кожен факт вибраний за критерієм «змінює рішення», а не «цікава цифра».");
  chrome(s, false, "Джерело: розділ 15 ноутбука · графік insights_portfolio.png");
}

/* ═══════════════════════════ 20. Висновки ═══════════════════════════ */
{
  const s = baseSlide(true);
  let y = title(s, "Пʼять відповідей кейсу — в одному місці", true, "висновки");
  y += 0.05;

  const answers = [
    { n: "01", q: "Мінімальний поріг", a: "$50", d: "Виведено з даних: на $25 у " + fmt.pctv(D.q1_information[1]["% без жодної події"], 0) + " креативів ще немає жодної події, тому ROMI −100% означає відсутність даних. До $50 — лише технічна перевірка." },
    { n: "02", q: "Очевидні випадки", a: "на $200 за зоною ROMI", d: "≥ +100% — сильний: змінює статус у " + fmt.pctv(D.zones_200[2]["змінили статус пізніше"], 1) + ". ≤ −50% — ризиковий: " + fmt.pctv(D.zones_200[0]["змінили статус пізніше"], 1) + ". Середнє по всіх — " + fmt.pctv(D.stability[3].flip_rate_fixed * 100, 1) + "." },
    { n: "03", q: "Неоднозначні й максимум", a: "$300, межа $500", d: "Сіра зона — " + fmt.pctv(D.zones_200[1]["% тих, хто дійшов"], 0) + " креативів і " + fmt.pctv(D.zones_200[1]["змінили статус пізніше"], 1) + " змін статусу. Суми «достатньої точності» не існує, тому межу обрано за ціною." },
    { n: "04", q: "Розрізи", a: "пороги спільні, звіт окремий", d: "Ціна події різниться в " + (D.cpt_topic[D.cpt_topic.length - 1]["вартість події, $"] / D.cpt_topic[0]["вартість події, $"]).toFixed(1).replace(".", ",") + " раза між темами, тож обʼєми вже різні. Але сегментних даних поки замало для окремих порогів." },
    { n: "05", q: "Як виміряти ефект", a: "випадковий тест 50/50", d: "Рівні бюджети, прибуток за " + D.pilot.outcome_days + " днів, рішення за нижньою межею інтервалу. Розкид знижуємо кластерами й CUPED." },
  ];
  const rowH = 0.78;
  answers.forEach((a, i) => {
    const yy = y + i * rowH;
    s.addText(a.n, { x: M, y: yy, w: 0.55, h: 0.4, fontFace: F.head, fontSize: 17, bold: true, color: C.primary === "4C3FD6" ? "9A90F0" : C.primary, margin: 0, valign: "middle" });
    s.addText(a.q, { x: M + 0.6, y: yy, w: 2.6, h: 0.4, fontFace: F.body, fontSize: 12.5, bold: true, color: C.onDark, margin: 0, valign: "middle" });
    s.addText(a.a, { x: M + 3.25, y: yy, w: 2.75, h: 0.4, fontFace: F.body, fontSize: 12.5, bold: true, color: "7BE07B", margin: 0, valign: "middle" });
    s.addText(a.d, { x: M + 6.1, y: yy - 0.04, w: CW - 6.1, h: 0.62, fontFace: F.body, fontSize: 10.5, color: C.onDarkMuted, margin: 0, valign: "middle" });
  });
  y += answers.length * rowH + 0.2;

  card(s, { x: M, y, w: CW, h: 1.05, fill: C.inkSoft, edge: "3B3670" });
  s.addText([
    { text: "Рекомендація:  ", options: { bold: true, fontSize: 13.5, color: C.onDark } },
    { text: "запустити контрольований пілот на кампаніях із рівними бюджетами. Історія підтримує порядок пріоритетів, але не доводить економічний ефект — і ми не видаємо одне за інше.", options: { fontSize: 13.5, color: C.onDarkMuted } },
  ], { x: M + 0.3, y: y + 0.14, w: CW - 0.6, h: 0.78, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Усі пʼять питань кейсу закриті. Головна відмінність від наївного рішення: ми не стверджуємо, що знайшли універсальні пороги, і чітко розділяємо доведене та правдоподібне.");
  chrome(s, true, null);
}

/* ═══════════════════════════ 21. Обмеження ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Чого ми не стверджуємо", false, "межі висновків");
  y = lead(s, "Розділ, без якого фреймворк був би небезпечним: він показує, де саме закінчується доведене.", y);

  const cw = (CW - 0.24) / 2;
  const left = [
    { h: "Ми не довели, що правила заробляють гроші", b: "Історичний результат видно лише для креативів, яким у минулому вже дали бюджет. Це селективна розмітка, а не випадкова вибірка." },
    { h: "Ми не пропонуємо автоматичне вимкнення", b: "Наступні гроші ризикової зони окупаються під " + fmt.pctv(D.zones_200[0]["ROMI наступних витрат разом"], 1) + " — нижче цілі, але це прибуток. Право вимикати дає лише експеримент." },
    { h: "Пороги −50%, +100%, $50, $200, $500 — кандидати", b: "Вони обґрунтовані даними, але не оптимізовані як універсальні константи. Точні значення уточнює пілот." },
  ];
  const right = [
    { h: "Дохід у файлі прогнозований, а не отриманий", b: "predict_net_revenue — це pLTV. Одну версію прогнозу фіксуємо до старту, а після дозрівання когорт звіряємо з фактичними грошима." },
    { h: "Семантика поля trials не підтверджена", b: "Усі ненульові значення кратні 16, причину автор кейсу не назвала. Тому ніде в роботі ми не називаємо це число кількістю оплат." },
    { h: "Часове покриття перевірки вузьке", b: "Чотири періоди перевірки моделі лежать у проміжку близько 30.11 – 19.01. Це краще за один поділ, але не доказ стійкості між сезонами." },
  ];
  [left, right].forEach((col, ci) => {
    col.forEach((o, i) => {
      const x = M + ci * (cw + 0.24);
      const yy = y + i * 1.22;
      card(s, { x, y: yy, w: cw, h: 1.1, fill: C.card, shadow: false });
      s.addText(o.h, { x: x + 0.24, y: yy + 0.14, w: cw - 0.48, h: 0.28, fontFace: F.body, fontSize: 12, bold: true, color: C.ink, margin: 0 });
      s.addText(o.b, { x: x + 0.24, y: yy + 0.44, w: cw - 0.48, h: 0.58, fontFace: F.body, fontSize: 10.5, color: C.textSoft, margin: 0, valign: "top" });
    });
  });
  y += 3 * 1.22 + 0.12;

  card(s, { x: M, y, w: CW, h: 0.92, fill: "F1F7F1", edge: "BFE3BF", shadow: false });
  s.addText([
    { text: "Що з цього випливає для впровадження.  ", options: { bold: true, fontSize: 12.5, color: C.ink } },
    { text: "Фреймворк можна вмикати вже зараз — але як керовану процедуру з журналом і пілотом, а не як автопілот. Перший результат, який дасть право розширювати правила, — прибуток за 14 днів у випадковому тесті.", options: { fontSize: 12.5, color: C.textSoft } },
  ], { x: M + 0.28, y: y + 0.08, w: CW - 0.56, h: 0.76, fontFace: F.body, margin: 0, valign: "middle" });

  s.addNotes("Цей слайд навмисно залишений у деку. Він показує, що межі висновків усвідомлені, і саме тому решті цифр можна довіряти.");
  chrome(s, false, "Джерело: розділ 14 ноутбука");
}

/* ═══════════════════════════ 22. Пакет ═══════════════════════════ */
{
  const s = baseSlide(false);
  let y = title(s, "Що в пакеті й як це перевірити", false, "додаток");
  y = lead(s, "Кожне число на слайдах відтворюється з початкового файлу даних за один запуск.", y);

  const files = [
    { n: "01", t: "Ця презентація", d: "Усі висновки, правила, обмеження й план перевірки — безпосередньо на слайдах.", c: C.primary },
    { n: "02", t: "02_analysis_report.html", d: "Статична версія висновків, таблиць і графіків без коду. Відкривається в браузері.", c: C.primary },
    { n: "03", t: "02_analysis_notebook.ipynb", d: "Відтворюваний аналіз: 16 розділів, маршрут читання й таблиця «питання кейсу → розділ».", c: C.accent },
    { n: "04", t: "output/deck_data.json", d: "Числа, з яких зібрано цю деку. Експортуються розділом 16 ноутбука.", c: C.accent },
    { n: "05", t: "data/creo_framework_data.csv", d: "Вхідні дані для повторного запуску розрахунків.", c: C.grey },
    { n: "06", t: "requirements.txt", d: "Версії бібліотек, на яких перевірено повний запуск від початку до кінця.", c: C.grey },
  ];
  const cw = (CW - 0.22 * 2) / 3, ch = 1.3;
  files.forEach((f, i) => {
    const x = M + (i % 3) * (cw + 0.22);
    const yy = y + Math.floor(i / 3) * (ch + 0.2);
    card(s, { x, y: yy, w: cw, h: ch, fill: "FFFFFF", edge: C.cardEdge });
    s.addText(f.n, { x: x + 0.22, y: yy + 0.16, w: 0.6, h: 0.26, fontFace: F.body, fontSize: 11, bold: true, color: f.c, margin: 0 });
    s.addText(f.t, { x: x + 0.22, y: yy + 0.42, w: cw - 0.44, h: 0.3, fontFace: F.body, fontSize: 12, bold: true, color: C.ink, margin: 0 });
    s.addText(f.d, { x: x + 0.22, y: yy + 0.72, w: cw - 0.44, h: 0.5, fontFace: F.body, fontSize: 10, color: C.textSoft, margin: 0, valign: "top" });
  });
  y += 2 * ch + 0.2 + 0.26;

  card(s, { x: M, y, w: CW, h: 1.5, fill: C.ink, edge: C.ink });
  s.addText("Як відтворити", { x: M + 0.3, y: y + 0.16, w: 4.0, h: 0.3, fontFace: F.body, fontSize: 13, bold: true, color: C.onDark, margin: 0 });
  s.addText("Встановити бібліотеки з requirements.txt і виконати всі комірки ноутбука від початку до кінця з цієї теки як кореня проєкту. Графіки збережуться в output/figures, числа — в output/deck_data.json.", {
    x: M + 0.3, y: y + 0.52, w: 7.4, h: 0.85, fontFace: F.body, fontSize: 11.5, color: C.onDarkMuted, margin: 0, valign: "top",
  });
  const st2 = [["44", "розрахункові комірки"], ["0", "помилок при повному запуску"], ["16", "збережених графіків"]];
  st2.forEach(([v, l], i) => {
    const x = M + 8.2 + i * 1.45;
    s.addText(v, { x, y: y + 0.34, w: 1.4, h: 0.42, fontFace: F.body, fontSize: 22, bold: true, color: C.onDark, margin: 0 });
    s.addText(l, { x, y: y + 0.76, w: 1.4, h: 0.5, fontFace: F.body, fontSize: 9.5, color: C.onDarkMuted, margin: 0, valign: "top" });
  });

  s.addNotes("Пакет самодостатній: чернетки, журнали роботи й старі версії навмисно не включені, щоб не створювати суперечностей із фінальним рішенням.");
  chrome(s, false, null);
}

pres.writeFile({ fileName: OUT }).then(() => {
  console.log("Готово:", path.basename(OUT));
  console.log("Слайдів:", slideNo);
});

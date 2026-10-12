const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI';
pptx.company = 'AI Course';
pptx.subject = '用最通俗的方式理解 AI 原理与企业落地';
pptx.title = '看懂 AI，才能用好 AI';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Cambria',
  bodyFontFace: 'Arial',
  lang: 'zh-CN'
};

const W = 13.333;
const H = 7.5;
const C = {
  ink: '10212B',
  navy: '0B1020',
  indigo: '1A2454',
  teal: '1B998B',
  mint: '34D399',
  ice: '6EE7F9',
  coral: 'FF7A59',
  amber: 'FFC857',
  red: 'E85D75',
  white: 'FFFFFF',
  paper: 'F7FAFC',
  pale: 'EEF5F5',
  line: 'D6E2E6',
  muted: '637381',
  soft: 'E8EFF2',
  darkMuted: 'A9B5C6'
};

// PingFang SC is installed on the target macOS/LibreOffice environment and
// keeps Chinese glyphs visible in both the deck and the rendered QA images.
const FONT = { head: 'Noto Sans CJK SC', body: 'Noto Sans CJK SC' };

function addBg(slide, color = C.paper) {
  slide.background = { color };
}

function addTitle(slide, title, kicker, dark = false) {
  const color = dark ? C.white : C.ink;
  const muted = dark ? C.darkMuted : C.muted;
  if (kicker) slide.addText(kicker.toUpperCase(), { x: 0.65, y: 0.38, w: 4.8, h: 0.24, fontFace: FONT.body, fontSize: 11, bold: true, charSpacing: 1.8, color: dark ? C.ice : C.teal, margin: 0 });
  slide.addText(title, { x: 0.65, y: 0.66, w: 9.8, h: 0.58, fontFace: FONT.head, fontSize: 27, bold: true, color, margin: 0, breakLine: false, fit: 'shrink' });
  slide.addText('看懂原理 → 会算账 → 设边界 → 留人复核', { x: 8.55, y: 0.47, w: 4.1, h: 0.22, fontFace: FONT.body, fontSize: 9.5, color: muted, align: 'right', margin: 0 });
}

function addFooter(slide, n, dark = false) {
  slide.addText(String(n).padStart(2, '0'), { x: 12.25, y: 7.08, w: 0.45, h: 0.18, fontFace: FONT.body, fontSize: 9, bold: true, color: dark ? C.darkMuted : C.muted, align: 'right', margin: 0 });
}

function addNotes(slide, text) { slide.addNotes(text); }

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.12,
    fill: { color: opts.fill || C.white, transparency: opts.transparency || 0 },
    line: { color: opts.line || C.line, transparency: opts.lineTransparency ?? 0, width: opts.lineWidth || 1 },
    shadow: opts.shadow ? { type: 'outer', color: '7D8A91', blur: 1, angle: 45, distance: 1, opacity: 0.12 } : undefined
  });
}

function pill(slide, text, x, y, w, color, fill, fs = 11) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.32, rectRadius: 0.16, fill: { color: fill }, line: { color: fill } });
  slide.addText(text, { x, y: y + 0.06, w, h: 0.17, fontFace: FONT.body, fontSize: fs, bold: true, color, align: 'center', margin: 0 });
}

function dot(slide, x, y, r, fill, line = fill) {
  slide.addShape(pptx.ShapeType.ellipse, { x, y, w: r, h: r, fill: { color: fill }, line: { color: line, width: 1 } });
}

function line(slide, x1, y1, x2, y2, color = C.line, width = 1.5, dash = 'solid', begin = 'none', end = 'none') {
  slide.addShape(pptx.ShapeType.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color, width, dash, beginArrowType: begin, endArrowType: end } });
}

function arrow(slide, x, y, w, h, color = C.teal) {
  slide.addShape(pptx.ShapeType.chevron, { x, y, w, h, fill: { color }, line: { color } });
}

function label(slide, text, x, y, w, color = C.muted, fs = 11, opts = {}) {
  slide.addText(text, { x, y, w, h: opts.h || 0.24, fontFace: opts.fontFace || FONT.body, fontSize: fs, bold: !!opts.bold, italic: !!opts.italic, color, align: opts.align || 'left', valign: opts.valign || 'mid', margin: opts.margin ?? 0, breakLine: false, fit: 'shrink' });
}

function body(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, { x, y, w, h, fontFace: FONT.body, fontSize: opts.fs || 16, color: opts.color || C.ink, bold: !!opts.bold, breakLine: false, margin: opts.margin ?? 0, valign: opts.valign || 'top', fit: 'shrink', paraSpaceAfterPt: opts.paraSpaceAfterPt || 6, bullet: opts.bullet });
}

// 01 — Cover
{
  const s = pptx.addSlide(); addBg(s, C.navy);
  s.addShape(pptx.ShapeType.ellipse, { x: 8.3, y: 0.5, w: 4.1, h: 4.1, fill: { color: C.indigo, transparency: 22 }, line: { color: C.ice, transparency: 55, width: 1.2 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.2, y: 1.4, w: 2.3, h: 2.3, fill: { color: C.teal, transparency: 18 }, line: { color: C.ice, transparency: 25, width: 1.5 } });
  s.addText('看懂 AI，\n才能用好 AI', { x: 0.75, y: 1.1, w: 6.7, h: 1.6, fontFace: FONT.head, fontSize: 34, bold: true, color: C.white, margin: 0, breakLine: false, fit: 'shrink' });
  s.addText('从函数、神经网络到企业流程的一条故事线', { x: 0.78, y: 3.0, w: 6.0, h: 0.34, fontFace: FONT.body, fontSize: 17, color: C.ice, margin: 0 });
  s.addText('AI 不是魔法，是一台从例子里调旋钮、再按旋钮给答案的概率机器。', { x: 0.78, y: 4.05, w: 5.8, h: 0.55, fontFace: FONT.body, fontSize: 15, color: C.darkMuted, margin: 0, breakLine: false, fit: 'shrink' });
  // right-side function story
  label(s, '输入', 8.22, 5.33, 0.85, C.darkMuted, 11, { align: 'center' });
  label(s, '模型', 10.1, 5.33, 0.85, C.darkMuted, 11, { align: 'center' });
  label(s, '行动', 11.95, 5.33, 0.85, C.darkMuted, 11, { align: 'center' });
  dot(s, 8.45, 4.75, 0.45, C.ice); dot(s, 10.32, 4.75, 0.45, C.teal); dot(s, 12.18, 4.75, 0.45, C.coral);
  line(s, 8.9, 4.98, 10.3, 4.98, C.ice, 2, 'solid', 'none', 'triangle');
  line(s, 10.77, 4.98, 12.17, 4.98, C.teal, 2, 'solid', 'none', 'triangle');
  s.addText('fθ(x)', { x: 9.48, y: 2.15, w: 2.0, h: 0.55, fontFace: FONT.head, fontSize: 30, bold: true, italic: true, color: C.white, align: 'center', margin: 0 });
  s.addText('企业的关键问题不是“要不要 AI”，\n而是“把它放在哪个工位”。', { x: 8.0, y: 5.85, w: 4.4, h: 0.56, fontFace: FONT.body, fontSize: 15, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
  s.addText('企业 AI 原理入门 · 2026', { x: 0.78, y: 6.75, w: 3.0, h: 0.2, fontFace: FONT.body, fontSize: 10, color: C.darkMuted, margin: 0 });
  addFooter(s, 1, true);
  addNotes(s, '开场：今天不从术语背诵开始，而是从一个问题开始——AI 到底在做什么？理解它的工作方式，才能知道哪些工作适合交给它，哪些工作必须保留人。');
}

// 02 — Story map
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '先讲一个企业故事：新人小林如何变成“数字同事”', '01 · 先建立直觉');
  s.addText('给他 10 万条历史订单、客服对话和质检照片。先学习，再接单；答错要复盘，不能直接签字。', { x: 0.68, y: 1.42, w: 11.8, h: 0.44, fontFace: FONT.body, fontSize: 16, color: C.ink, margin: 0, fit: 'shrink' });
  const steps = [
    ['看资料', '数据', '从过去案例里找规律', C.ice],
    ['练习', '训练', '把“旋钮”调到合适', C.teal],
    ['接单', '推理', '拿新问题算一个答案', C.coral],
    ['复核', '闭环', '错了记录，继续改进', C.amber]
  ];
  steps.forEach((it, i) => {
    const x = 0.72 + i * 3.12;
    card(s, x, 2.35, 2.55, 2.45, { fill: C.white, line: it[3], lineWidth: 1.6, shadow: true });
    dot(s, x + 0.22, 2.58, 0.32, it[3]);
    label(s, String(i + 1), x + 0.22, 2.64, 0.32, C.navy, 12, { align: 'center', bold: true });
    label(s, it[0], x + 0.68, 2.56, 1.58, C.ink, 19, { bold: true });
    pill(s, it[1], x + 0.22, 3.16, 0.78, C.navy, it[3], 10);
    body(s, it[2], x + 0.22, 3.75, 1.95, 0.56, { fs: 14, color: C.muted });
    if (i < steps.length - 1) arrow(s, x + 2.66, 3.35, 0.34, 0.28, C.line);
  });
  card(s, 0.72, 5.38, 11.95, 0.85, { fill: C.indigo, line: C.indigo });
  s.addText('这就把三个问题串起来了：AI 怎样学？为什么会错？企业该把哪类工作交给它？', { x: 1.03, y: 5.65, w: 11.25, h: 0.28, fontFace: FONT.body, fontSize: 17, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 2);
  addNotes(s, '用“新人小林”贯穿全场。数据像他的工作档案，训练像岗前练习，推理像接单做事，复核像绩效复盘。AI 的价值来自整个闭环，而不是某一次回答。');
}

// 03 — AI definition
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '一句话定义 AI：给定函数，先找旋钮，再算答案', '02 · 训练与推理');
  s.addText('把模型想成一个函数：', { x: 0.78, y: 1.42, w: 2.6, h: 0.3, fontFace: FONT.body, fontSize: 16, color: C.muted, margin: 0 });
  s.addText('y = f(x; θ)', { x: 3.0, y: 1.25, w: 3.0, h: 0.55, fontFace: FONT.head, fontSize: 28, bold: true, italic: true, color: C.indigo, margin: 0, align: 'center' });
  card(s, 0.78, 2.16, 5.55, 3.35, { fill: C.white, line: C.ice, lineWidth: 1.5, shadow: true });
  pill(s, '训练', 1.06, 2.47, 0.88, C.navy, C.ice, 12);
  body(s, '给很多“输入 → 正确输出”的例子，\n反过来把参数 θ 调到合适。', 1.06, 3.05, 4.7, 0.78, { fs: 20, bold: true, color: C.ink });
  label(s, '像调收音机：一点点拧旋钮，让杂音变少。', 1.06, 4.32, 4.55, C.muted, 13);
  const trainPts = [[1.2, 4.9], [2.2, 4.68], [3.1, 4.83], [4.0, 4.55], [5.0, 4.38]];
  trainPts.forEach((p, i) => { dot(s, p[0], p[1], 0.16, i === 4 ? C.coral : C.teal); if (i < trainPts.length - 1) line(s, p[0] + 0.15, p[1] + 0.08, trainPts[i + 1][0], trainPts[i + 1][1] + 0.08, C.line, 1.5); });
  card(s, 6.95, 2.16, 5.58, 3.35, { fill: C.white, line: C.coral, lineWidth: 1.5, shadow: true });
  pill(s, '推理', 7.23, 2.47, 0.88, C.white, C.coral, 12);
  body(s, '参数 θ 固定，给一个新输入 x，\n计算出答案 ŷ。', 7.23, 3.05, 4.75, 0.78, { fs: 20, bold: true, color: C.ink });
  label(s, '像拿着已经调好的收音机，听一条新频道。', 7.23, 4.32, 4.85, C.muted, 13);
  s.addShape(pptx.ShapeType.roundRect, { x: 7.34, y: 4.74, w: 4.58, h: 0.52, rectRadius: 0.14, fill: { color: C.pale }, line: { color: C.pale } });
  label(s, '新问题 x', 7.55, 4.89, 0.95, C.indigo, 13, { bold: true });
  arrow(s, 8.7, 4.87, 0.45, 0.25, C.teal);
  label(s, '模型 fθ', 9.38, 4.89, 1.15, C.indigo, 13, { bold: true });
  arrow(s, 10.75, 4.87, 0.45, 0.25, C.teal);
  label(s, '答案 ŷ', 11.42, 4.89, 0.7, C.coral, 13, { bold: true });
  s.addText('训练反复减少误差，在验证表现上逼近一组能泛化的参数；推理时参数通常冻结。', { x: 0.95, y: 6.02, w: 11.35, h: 0.36, fontFace: FONT.body, fontSize: 14.5, color: C.indigo, bold: true, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 3);
  addNotes(s, '核心类比：训练是反过来找参数，推理是参数固定后代入新输入。这个区分能帮助企业理解：训练是一次性投入，推理是每次调用都会产生的成本与时延。');
}

// 04 — training vs inference
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '训练 vs 推理：一个是在改公式，一个是在代公式', '03 · 两个过程');
  card(s, 0.75, 1.5, 5.8, 4.72, { fill: C.white, line: C.ice, lineWidth: 1.5, shadow: true });
  pill(s, '训练：像学厨师', 1.04, 1.8, 1.62, C.navy, C.ice, 12);
  const tx = [
    ['1', '看很多菜谱和历史订单', C.indigo],
    ['2', '试做，和标准答案比差多少', C.teal],
    ['3', '根据“差距”调味、改参数', C.coral]
  ];
  tx.forEach((it, i) => { const y = 2.55 + i * 0.83; dot(s, 1.1, y, 0.34, it[2]); label(s, it[0], 1.1, y + 0.07, 0.34, C.white, 12, { align: 'center', bold: true }); body(s, it[1], 1.65, y + 0.02, 4.2, 0.3, { fs: 16, bold: i === 2, color: C.ink }); if (i < 2) line(s, 1.27, y + 0.37, 1.27, y + 0.67, C.line, 1.2, 'dash'); });
  label(s, '特点：慢、贵、集中发生；决定“学会什么”。', 1.05, 5.35, 4.95, C.muted, 13);
  card(s, 6.8, 1.5, 5.8, 4.72, { fill: C.white, line: C.coral, lineWidth: 1.5, shadow: true });
  pill(s, '推理：像接单做菜', 7.09, 1.8, 1.72, C.white, C.coral, 12);
  const ix = [
    ['1', '收到一个新问题', C.indigo],
    ['2', '按已经调好的配方计算', C.teal],
    ['3', '每个 token 都要付算力', C.coral]
  ];
  ix.forEach((it, i) => { const y = 2.55 + i * 0.83; dot(s, 7.15, y, 0.34, it[2]); label(s, it[0], 7.15, y + 0.07, 0.34, C.white, 12, { align: 'center', bold: true }); body(s, it[1], 7.7, y + 0.02, 4.2, 0.3, { fs: 16, bold: i === 2, color: C.ink }); if (i < 2) line(s, 7.32, y + 0.37, 7.32, y + 0.67, C.line, 1.2, 'dash'); });
  label(s, '特点：每问一次都花钱；长上下文、并发会放大成本。', 7.1, 5.35, 4.95, C.muted, 13);
  s.addText('企业决策：训练决定能力边界；推理决定日常账单。', { x: 1.05, y: 6.58, w: 11.3, h: 0.32, fontFace: FONT.body, fontSize: 17, bold: true, color: C.indigo, align: 'center', margin: 0 });
  addFooter(s, 4);
  addNotes(s, '把训练和推理分开讲，听众就能理解为什么“模型训练很贵”与“每次调用也要付钱”是两笔不同的账。');
}

// 05 — neuron linear + activation
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '一个神经元：线性层负责“混合”，激活函数负责“开闸”', '04 · 神经网络');
  card(s, 0.75, 1.55, 6.05, 4.75, { fill: C.white, line: C.line, lineWidth: 1.2, shadow: true });
  pill(s, '线性层：y = Wx + b', 1.05, 1.86, 2.05, C.navy, C.ice, 12);
  label(s, '把每个线索乘上权重，再加总', 1.05, 2.38, 3.6, C.muted, 13);
  const inputs = [['客户说“太贵”', 'x₁', C.ice], ['近 30 天退货率', 'x₂', C.teal], ['季节因素', 'x₃', C.coral]];
  inputs.forEach((it, i) => { const y = 3.02 + i * 0.74; s.addShape(pptx.ShapeType.roundRect, { x: 1.05, y, w: 2.08, h: 0.42, rectRadius: 0.1, fill: { color: C.pale }, line: { color: C.pale } }); label(s, it[0], 1.19, y + 0.11, 1.42, C.ink, 12, { bold: true }); label(s, it[1], 2.57, y + 0.11, 0.35, it[2], 11, { bold: true, align: 'right' }); line(s, 3.22, y + 0.21, 4.25, 3.76, it[2], 1.8, 'solid', 'none', 'triangle'); });
  s.addShape(pptx.ShapeType.ellipse, { x: 4.15, y: 3.35, w: 1.08, h: 1.08, fill: { color: C.indigo }, line: { color: C.indigo } });
  label(s, '加权\n求和', 4.15, 3.61, 1.08, C.white, 12, { align: 'center', bold: true, h: 0.4 });
  line(s, 5.24, 3.9, 6.1, 3.9, C.indigo, 1.8, 'solid', 'none', 'triangle');
  s.addShape(pptx.ShapeType.roundRect, { x: 6.05, y: 3.45, w: 0.5, h: 0.9, rectRadius: 0.15, fill: { color: C.coral }, line: { color: C.coral } });
  label(s, '输出', 6.06, 3.76, 0.48, C.white, 10, { align: 'center', bold: true });
  label(s, '权重只是“线索旋钮”的直觉比喻；真实重要性要靠评测验证。', 1.05, 5.55, 5.15, C.muted, 12.5);
  card(s, 7.1, 1.55, 5.45, 4.75, { fill: C.white, line: C.coral, lineWidth: 1.5, shadow: true });
  pill(s, '激活：ReLU', 7.4, 1.86, 1.28, C.white, C.coral, 12);
  label(s, '负数关掉，正数放行：ReLU(z)=max(0,z)', 7.4, 2.38, 4.25, C.muted, 13);
  line(s, 8.0, 5.48, 11.82, 5.48, C.line, 1.2);
  line(s, 10.02, 5.88, 10.02, 2.98, C.line, 1.2);
  line(s, 8.38, 5.48, 10.02, 5.48, C.coral, 3.2);
  line(s, 10.02, 5.48, 11.42, 3.35, C.teal, 3.2);
  label(s, '−', 7.9, 5.6, 0.28, C.muted, 15, { bold: true, align: 'center' }); label(s, '+', 11.85, 5.6, 0.28, C.muted, 15, { bold: true, align: 'center' });
  label(s, '关', 8.72, 5.55, 0.5, C.coral, 12, { bold: true }); label(s, '开', 10.82, 3.0, 0.5, C.teal, 12, { bold: true });
  s.addText('一层很简单；很多层把简单规则叠成复杂表达。', { x: 7.42, y: 5.82, w: 4.5, h: 0.3, fontFace: FONT.body, fontSize: 14.5, bold: true, color: C.indigo, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 5);
  addNotes(s, '不要把“神经元”讲成会思考的小人。它就是一个小计算单元：先做加权求和，再由激活函数决定是否放行。线性层混合信息，激活函数提供非线性。');
}

// 06 — activation intuition
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '为什么需要激活函数？把直线折成能拐弯的路', '05 · 非线性');
  s.addText('没有激活：叠很多层，仍然像一把直尺。', { x: 0.78, y: 1.46, w: 4.9, h: 0.28, fontFace: FONT.body, fontSize: 16, color: C.ink, bold: true, margin: 0 });
  card(s, 0.78, 2.0, 5.45, 3.95, { fill: C.white, line: C.line, shadow: true });
  line(s, 1.35, 5.35, 5.67, 5.35, C.line, 1.2); line(s, 1.55, 5.55, 1.55, 2.58, C.line, 1.2);
  line(s, 1.65, 5.1, 5.5, 2.85, C.indigo, 3.5);
  label(s, '只能用一条直线\n切分世界', 2.25, 3.25, 2.7, C.indigo, 18, { align: 'center', bold: true, h: 0.58 });
  label(s, '猫 / 狗、流失 / 不流失、\n合格 / 不合格……很多边界不是直的。', 1.3, 5.62, 4.3, C.muted, 12, { align: 'center', h: 0.45 });
  s.addText('加上激活：每个神经元都能在不同区域换一条规则。', { x: 6.92, y: 1.46, w: 5.55, h: 0.28, fontFace: FONT.body, fontSize: 16, color: C.ink, bold: true, margin: 0 });
  card(s, 6.92, 2.0, 5.55, 3.95, { fill: C.white, line: C.teal, lineWidth: 1.4, shadow: true });
  line(s, 7.48, 5.35, 11.9, 5.35, C.line, 1.2); line(s, 7.68, 5.55, 7.68, 2.58, C.line, 1.2);
  line(s, 7.78, 5.35, 9.35, 5.35, C.coral, 3.2);
  line(s, 9.35, 5.35, 11.57, 3.02, C.teal, 3.2);
  dot(s, 9.22, 5.22, 0.25, C.amber);
  label(s, '一个折点很简单；\n很多折点就能逼近复杂曲线。', 8.25, 3.22, 3.45, C.teal, 18, { align: 'center', bold: true, h: 0.62 });
  label(s, '模型因此可以“弯着理解”现实，而不是只会画直线。', 7.55, 5.62, 4.2, C.muted, 12, { align: 'center', h: 0.28 });
  card(s, 2.28, 6.36, 8.75, 0.5, { fill: C.indigo, line: C.indigo });
  s.addText('线性层负责组合信息；激活函数打破线性限制；多层组合才有复杂表达能力。', { x: 2.48, y: 6.51, w: 8.35, h: 0.17, fontFace: FONT.body, fontSize: 13.5, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 6);
  addNotes(s, '可以用“画地图”类比：只有直尺，能画的边界很有限；有了折点和弯路，模型才有机会描述复杂问题。激活函数提供表达能力，但不保证模型就一定学会。');
}

// 07 — network structure
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '网络结构 = 分工设计：很多简单层协作，才描述复杂问题', '06 · 网络结构');
  s.addText('像一条流水线：深度是工序，宽度是每道工序的人手。结构决定“先看什么，再组合什么”。', { x: 0.78, y: 1.4, w: 11.6, h: 0.34, fontFace: FONT.body, fontSize: 16, color: C.ink, margin: 0, fit: 'shrink' });
  const cols = [
    { x: 0.85, title: '输入', sub: '像素 / 词 / 数值', color: C.ice, h: 1.0 },
    { x: 3.15, title: '局部线索', sub: '边缘 / 关键词', color: C.teal, h: 1.6 },
    { x: 5.45, title: '组合特征', sub: '形状 / 短语', color: C.indigo, h: 2.18 },
    { x: 7.75, title: '任务判断', sub: '分类 / 生成', color: C.coral, h: 2.75 }
  ];
  cols.forEach((it, i) => {
    const y = 4.72 - it.h;
    card(s, it.x, y, 1.72, it.h, { fill: C.white, line: it.color, lineWidth: 1.6, shadow: true });
    dot(s, it.x + 0.65, y + 0.22, 0.42, it.color);
    label(s, String(i + 1), it.x + 0.65, y + 0.31, 0.42, C.white, 13, { align: 'center', bold: true });
    label(s, it.title, it.x + 0.16, y + it.h - 0.76, 1.4, C.ink, 16, { bold: true, align: 'center' });
    label(s, it.sub, it.x + 0.12, y + it.h - 0.43, 1.48, C.muted, 11, { align: 'center' });
    if (i < cols.length - 1) arrow(s, it.x + 1.82, y + it.h - 0.56, 0.38, 0.28, C.line);
  });
  // parallel and skip connections mini diagrams
  card(s, 10.0, 1.9, 2.65, 1.28, { fill: C.pale, line: C.line });
  label(s, '并联：从多个角度看', 10.22, 2.12, 2.2, C.indigo, 13, { bold: true, align: 'center' });
  [2.48, 2.77, 3.06].forEach((y, i) => { dot(s, 10.35, y, 0.12, [C.ice, C.teal, C.coral][i]); line(s, 10.48, y + 0.06, 11.1, 2.77, [C.ice, C.teal, C.coral][i], 1.2); }); dot(s, 11.1, 2.68, 0.18, C.indigo);
  card(s, 10.0, 3.45, 2.65, 1.28, { fill: C.pale, line: C.line });
  label(s, '跳跃：让原信息直达', 10.22, 3.68, 2.2, C.indigo, 13, { bold: true, align: 'center' });
  line(s, 10.35, 4.37, 11.08, 4.05, C.teal, 1.8, 'solid', 'none', 'triangle'); line(s, 10.35, 3.89, 11.08, 4.05, C.coral, 1.8, 'solid', 'none', 'triangle'); dot(s, 11.03, 3.93, 0.24, C.indigo);
  s.addText('不是越大越好：容量、数据、成本和泛化要匹配。', { x: 0.92, y: 6.1, w: 8.7, h: 0.3, fontFace: FONT.body, fontSize: 16, bold: true, color: C.indigo, margin: 0 });
  addFooter(s, 7);
  addNotes(s, '网络结构就是“分工设计”。串联负责逐层抽象，并联负责多视角，跳跃连接负责保留原信息。企业不要只问模型多大，要问结构是否匹配任务。');
}

// 08 — Transformer
{
  const s = pptx.addSlide(); addBg(s, C.navy); addTitle(s, 'Transformer：让每个词先问“我该关注谁？”', '07 · 上下文', true);
  s.addText('传统按顺序读，长句容易忘；Transformer 用注意力让每个词都能回头看上下文。', { x: 0.78, y: 1.44, w: 11.7, h: 0.34, fontFace: FONT.body, fontSize: 16, color: C.darkMuted, margin: 0 });
  const toks = ['小明', '把', '书', '放进', '抽屉', '因为', '它', '太重'];
  const xs = [0.82, 2.05, 2.9, 3.78, 5.1, 6.5, 7.7, 8.72];
  toks.forEach((t, i) => { const emph = t === '它'; s.addShape(pptx.ShapeType.roundRect, { x: xs[i], y: 3.0, w: t.length > 1 ? 0.9 : 0.62, h: 0.52, rectRadius: 0.12, fill: { color: emph ? C.coral : C.indigo }, line: { color: emph ? C.coral : C.ice, transparency: emph ? 0 : 35, width: 1.1 } }); label(s, t, xs[i], 3.16, t.length > 1 ? 0.9 : 0.62, C.white, 15, { align: 'center', bold: emph }); });
  // attention lines
  line(s, 8.0, 2.96, 3.35, 2.96, C.coral, 3.0, 'solid', 'none', 'triangle');
  line(s, 8.0, 2.86, 5.55, 2.86, C.amber, 1.7, 'solid', 'none', 'triangle');
  line(s, 8.0, 2.76, 1.45, 2.76, C.ice, 1.1, 'solid', 'none', 'triangle');
  label(s, '“它”更该关注“书”', 8.35, 2.18, 2.25, C.coral, 14, { bold: true });
  card(s, 0.82, 4.35, 3.52, 1.35, { fill: C.indigo, line: C.ice, lineWidth: 1.2 });
  label(s, 'Q · Query', 1.07, 4.65, 1.2, C.ice, 15, { bold: true });
  label(s, '我在找什么？', 1.07, 5.08, 2.7, C.white, 15, { bold: true });
  card(s, 4.62, 4.35, 3.52, 1.35, { fill: C.indigo, line: C.teal, lineWidth: 1.2 });
  label(s, 'K · Key', 4.87, 4.65, 1.2, C.mint, 15, { bold: true });
  label(s, '谁和我最匹配？', 4.87, 5.08, 2.7, C.white, 15, { bold: true });
  card(s, 8.42, 4.35, 3.52, 1.35, { fill: C.indigo, line: C.coral, lineWidth: 1.2 });
  label(s, 'V · Value', 8.67, 4.65, 1.2, C.coral, 15, { bold: true });
  label(s, '对方带来什么信息？', 8.67, 5.08, 2.9, C.white, 15, { bold: true });
  s.addText('Q 与 K 打分 → softmax 变成关注比例 → 按比例汇总 V', { x: 0.95, y: 5.91, w: 11.4, h: 0.22, fontFace: FONT.body, fontSize: 12.5, color: C.amber, align: 'center', bold: true, margin: 0, fit: 'shrink' });
  s.addText('位置 → 因果多头注意力（Q/K/V）→ 残差/归一化 → FFN（线性+激活）→ 重复多层', { x: 0.82, y: 6.25, w: 11.7, h: 0.25, fontFace: FONT.body, fontSize: 12.5, color: C.ice, align: 'center', bold: true, margin: 0, fit: 'shrink' });
  addFooter(s, 8, true);
  addNotes(s, 'Transformer 的关键不是“更聪明”，而是更擅长处理上下文。注意力可以理解为：每个词都在决定自己应该参考哪些词。Q/K/V 是“我在找什么、谁匹配、对方带来什么”。');
}

// 09 — GPT
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, 'GPT：每次只猜下一个词，却能写成一段话', '08 · 生成模型');
  s.addText('GPT = Generative Pre-trained Transformer：先学语言的统计规律，再按要求逐 token（文字片段）接龙。', { x: 0.78, y: 1.43, w: 11.6, h: 0.32, fontFace: FONT.body, fontSize: 15.5, color: C.ink, margin: 0, fit: 'shrink' });
  card(s, 0.78, 2.0, 7.25, 3.82, { fill: C.white, line: C.line, shadow: true });
  pill(s, '预训练 → 对齐 → 逐 token 生成', 1.08, 2.3, 2.55, C.navy, C.ice, 11.5);
  label(s, '对齐 = 学会按要求回答、守规则，不等于把新知识写进脑子。', 1.08, 2.72, 5.55, C.indigo, 12, { bold: true });
  const seq = ['客户投诉：', '鞋底', '开胶，', '建议', '先'];
  seq.forEach((t, i) => { const x = 1.08 + i * 1.12; s.addShape(pptx.ShapeType.roundRect, { x, y: 3.17, w: 0.98, h: 0.52, rectRadius: 0.1, fill: { color: i === seq.length - 1 ? C.coral : C.pale }, line: { color: i === seq.length - 1 ? C.coral : C.line } }); label(s, t, x, 3.33, 0.98, i === seq.length - 1 ? C.white : C.ink, 13, { align: 'center', bold: i === seq.length - 1 }); if (i < seq.length - 1) arrow(s, x + 1.0, 3.31, 0.24, 0.22, C.line); });
  label(s, '候选 token 概率（仅展示部分）：', 1.1, 4.14, 3.0, C.muted, 13, { bold: true });
  const cand = [['给', '0.42', C.teal], ['建议', '0.28', C.ice], ['联系', '0.16', C.amber], ['直接', '0.08', C.coral]];
  cand.forEach((it, i) => { const y = 4.58 + i * 0.3; label(s, it[0], 1.25, y, 0.55, C.ink, 12, { bold: true }); s.addShape(pptx.ShapeType.roundRect, { x: 2.0, y: y + 0.04, w: 3.2 * Number(it[1]), h: 0.13, rectRadius: 0.06, fill: { color: it[2] }, line: { color: it[2] } }); label(s, it[1], 5.55, y, 0.55, C.muted, 11, { align: 'right' }); });
  card(s, 8.45, 2.0, 4.05, 3.82, { fill: C.indigo, line: C.indigo, shadow: true });
  label(s, '上下文窗口', 8.82, 2.38, 2.8, C.ice, 16, { bold: true });
  label(s, '参数是长期学到的经验；\n桌面是这次能直接参考的材料。\n未提供的私有 / 最新资料，不能可靠推断。', 8.82, 2.95, 3.05, C.white, 14.5, { bold: true, h: 1.0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 8.82, y: 4.1, w: 3.1, h: 0.72, rectRadius: 0.12, fill: { color: C.navy }, line: { color: C.ice, transparency: 35 } });
  label(s, '温度 ↑ = 更发散\n温度 ↓ = 更稳定', 9.0, 4.28, 2.7, C.ice, 13, { align: 'center', bold: true, h: 0.38 });
  s.addText('流畅度很高 ≠ 事实一定正确。每一步都从概率分布按策略选一个 token：可以选最高，也可以采样。', { x: 0.95, y: 6.3, w: 11.35, h: 0.34, fontFace: FONT.body, fontSize: 15.5, bold: true, color: C.coral, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 9);
  addNotes(s, 'GPT 的训练目标是预测下一个 token，生成时通常一步一步解码。因果遮罩保证当前位置只看已经出现的内容；温度主要改变抽样随机性，不会增加事实知识或纠错能力。');
}

// 10 — RSI
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, 'RSI：AI 参与改进“造 AI 的流程”', '09 · 前沿方向', false);
  pill(s, 'Recursive Self-Improvement · 递归自我改进', 0.78, 1.38, 3.2, C.navy, C.amber, 11);
  label(s, '这是研究方向，不是企业可以直接放开的“无限自我进化”。', 4.2, 1.46, 6.3, C.muted, 13);
  // loop nodes
  const loop = [
    { x: 1.2, y: 3.0, title: '发现弱点', sub: '哪里总出错？', c: C.ice },
    { x: 4.0, y: 2.02, title: '提方案', sub: '生成数据 / 代码', c: C.teal },
    { x: 7.35, y: 2.02, title: '跑实验', sub: '在沙盒里验证', c: C.indigo },
    { x: 10.15, y: 3.0, title: '评估更新', sub: '指标更好才保留', c: C.coral }
  ];
  loop.forEach((it) => { s.addShape(pptx.ShapeType.ellipse, { x: it.x, y: it.y, w: 1.55, h: 1.55, fill: { color: it.c }, line: { color: C.white, width: 1.2 } }); label(s, it.title, it.x + 0.08, it.y + 0.47, 1.39, C.white, 16, { align: 'center', bold: true }); label(s, it.sub, it.x + 0.12, it.y + 0.95, 1.31, C.white, 10, { align: 'center', h: 0.26 }); });
  line(s, 2.75, 3.72, 4.1, 3.0, C.teal, 2.1, 'solid', 'none', 'triangle');
  line(s, 5.55, 2.77, 7.3, 2.77, C.teal, 2.1, 'solid', 'none', 'triangle');
  line(s, 8.9, 3.0, 10.1, 3.72, C.teal, 2.1, 'solid', 'none', 'triangle');
  line(s, 10.15, 4.65, 2.0, 4.65, C.amber, 2, 'dash', 'none', 'triangle');
  label(s, '下一轮用已通过评估的版本继续迭代', 4.75, 4.78, 3.9, C.amber, 12, { align: 'center', bold: true });
  card(s, 2.32, 5.65, 8.75, 0.62, { fill: C.white, line: C.amber, lineWidth: 1.3 });
  s.addText('企业今天能借鉴的是“有界自我改进”：固定指标 + 沙盒权限 + 人审批 + 可回滚；聊天反馈不会自动改写基础模型权重。', { x: 2.48, y: 5.78, w: 8.45, h: 0.28, fontFace: FONT.body, fontSize: 13.5, bold: true, color: C.indigo, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 10);
  addNotes(s, '这里的 RSI 按 Recursive Self-Improvement 解释。企业可做的是让 AI 帮忙生成测试、数据和改进方案，但必须由固定外部指标、独立留出集、沙盒、权限、人审批和回滚机制来约束。改进代码、工具或数据不等于基础模型权重自动更新，也不保证能力无限增长。');
}

// 11 — Limits
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, 'AI 的边界：会说 ≠ 说得对', '10 · 局限性');
  s.addText('模型越流畅，越需要把“能不能核验”写进流程。', { x: 0.78, y: 1.4, w: 5.3, h: 0.3, fontFace: FONT.body, fontSize: 16, color: C.ink, bold: true, margin: 0 });
  const cards = [
    ['幻觉', '把猜测说成事实', '放大镜', C.coral],
    ['过时', '训练没覆盖新信息', '时钟', C.amber],
    ['脆弱', '换个问法就失灵', '裂缝', C.teal],
    ['偏差 / 安全', '数据偏见、注入、泄漏', '盾牌', C.indigo]
  ];
  cards.forEach((it, i) => { const x = 0.8 + (i % 2) * 3.1; const y = 2.05 + Math.floor(i / 2) * 1.75; card(s, x, y, 2.72, 1.34, { fill: C.white, line: it[3], lineWidth: 1.5, shadow: true }); dot(s, x + 0.22, y + 0.23, 0.42, it[3]); label(s, String(i + 1), x + 0.22, y + 0.33, 0.42, C.white, 12, { align: 'center', bold: true }); label(s, it[0], x + 0.82, y + 0.23, 1.55, C.ink, 16, { bold: true }); label(s, it[1], x + 0.82, y + 0.68, 1.55, C.muted, 12, { h: 0.34 }); });
  card(s, 7.15, 2.0, 5.25, 3.9, { fill: C.indigo, line: C.indigo, shadow: true });
  label(s, '可靠性三角', 7.53, 2.38, 2.1, C.ice, 17, { bold: true });
  const cx = 9.82, cy = 4.25;
  s.addShape(pptx.ShapeType.triangle, { x: 8.42, y: 3.05, w: 2.8, h: 2.3, rotate: 0, fill: { color: C.navy, transparency: 20 }, line: { color: C.ice, width: 1.2 } });
  dot(s, cx - 0.14, 3.25, 0.28, C.mint); dot(s, 8.55, 5.07, 0.28, C.amber); dot(s, 11.0, 5.07, 0.28, C.coral);
  label(s, '正确', 9.62, 2.95, 0.6, C.mint, 12, { bold: true, align: 'center' }); label(s, '速度', 8.05, 5.38, 0.6, C.amber, 12, { bold: true, align: 'center' }); label(s, '成本', 11.07, 5.38, 0.6, C.coral, 12, { bold: true, align: 'center' });
  label(s, '通常要在三者之间做取舍，\n而不是同时拉满。', 8.28, 3.9, 3.15, C.white, 16, { align: 'center', bold: true, h: 0.58 });
  s.addText('关键结果必须可核验、可追责：让 AI 先做草稿和分析，人保留最后判断。', { x: 0.95, y: 6.3, w: 11.3, h: 0.32, fontFace: FONT.body, fontSize: 16, bold: true, color: C.indigo, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 11);
  addNotes(s, '用四个词把局限讲清：幻觉、过时、脆弱、偏差与安全。可靠性三角表示在同一任务与资源约束下，准确率、速度、成本通常需要取舍；缓存、批处理和模型组合可以改变这个平衡。');
}

// 12 — Cost
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '成本不是只有“买模型”：真正贵的是整条链路', '11 · 成本');
  s.addText('估算公式：调用次数 ×（输入 token×输入单价 + 输出 token×输出单价）+ 训练/微调 + 数据/集成/人工 + 错误成本。', { x: 0.78, y: 1.42, w: 11.6, h: 0.34, fontFace: FONT.body, fontSize: 14.5, color: C.ink, bold: true, margin: 0, fit: 'shrink' });
  const costItems = [
    ['训练 / 微调', '需要时才发生的算力和实验', C.indigo, 2.7],
    ['推理', '每次调用、长上下文、并发', C.coral, 3.2],
    ['数据', '清洗、标注、知识库', C.teal, 2.45],
    ['集成', '权限、监控、评测、上线', C.indigo, 2.1],
    ['人工', '复核、运营、培训', C.amber, 1.6],
    ['错误', '返工、合规、客户损失', C.red, 2.85]
  ];
  costItems.forEach((it, i) => { const y = 2.15 + i * 0.58; label(s, it[0], 0.95, y + 0.08, 0.82, C.ink, 12.5, { bold: true }); s.addShape(pptx.ShapeType.roundRect, { x: 1.9, y, w: it[3], h: 0.3, rectRadius: 0.11, fill: { color: it[2] }, line: { color: it[2] } }); label(s, it[1], 5.35, y + 0.08, 3.2, C.muted, 11.5); });
  label(s, '条形长度只是成本项示意，不代表实际占比。', 1.0, 5.72, 6.4, C.muted, 11, { italic: true });
  card(s, 8.98, 2.15, 3.45, 3.64, { fill: C.white, line: C.line, shadow: true });
  label(s, '按任务匹配工具', 9.3, 2.5, 2.8, C.indigo, 18, { bold: true, align: 'center' });
  const toolRows = [
    ['规则', '稳定、便宜', C.teal],
    ['小模型', '固定、高频、低风险', C.ice],
    ['大模型', '复杂 / 不确定，先算单位成本', C.coral],
    ['人工', '高风险、不可错', C.amber]
  ];
  toolRows.forEach((it, i) => { const y = 3.15 + i * 0.52; dot(s, 9.45, y + 0.06, 0.18, it[2]); label(s, it[0], 9.76, y, 0.72, C.ink, 13, { bold: true }); label(s, it[1], 10.6, y, 1.45, C.muted, 12); });
  s.addText('先算单位经济性，再扩张调用量。', { x: 1.2, y: 6.1, w: 11.0, h: 0.32, fontFace: FONT.body, fontSize: 17, bold: true, color: C.indigo, align: 'center', margin: 0 });
  addFooter(s, 12);
  addNotes(s, '企业容易漏算三笔钱：数据准备、集成治理、错误代价。用“任务—工具匹配”可以避免用最贵的模型处理最简单的问题。');
}

// 13 — scenario matrix
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '企业选场景：先看重复度，再看可控性', '12 · 场景选择');
  label(s, '优先高频、可量化、可复核、数据已经存在的任务。', 0.8, 1.4, 6.2, C.ink, 16, { bold: true });
  // axes and quadrants
  const x0 = 2.1, y0 = 2.05, w = 8.9, h = 3.7;
  s.addShape(pptx.ShapeType.rect, { x: x0, y: y0, w, h, fill: { color: C.white }, line: { color: C.line, width: 1.2 } });
  s.addShape(pptx.ShapeType.rect, { x: x0 + w / 2, y: y0, w: w / 2, h: h / 2, fill: { color: 'E8FFF5', transparency: 10 }, line: { color: 'E8FFF5', transparency: 100 } });
  s.addShape(pptx.ShapeType.rect, { x: x0, y: y0 + h / 2, w: w / 2, h: h / 2, fill: { color: 'FFF6DD', transparency: 12 }, line: { color: 'FFF6DD', transparency: 100 } });
  line(s, x0 + w / 2, y0, x0 + w / 2, y0 + h, C.line, 1.1); line(s, x0, y0 + h / 2, x0 + w, y0 + h / 2, C.line, 1.1);
  label(s, '重复度高 →', x0 + w - 1.6, y0 + h + 0.2, 1.55, C.muted, 12, { align: 'right', bold: true });
  s.addText('可控性 ↑', { x: 0.9, y: 2.82, w: 0.7, h: 0.52, fontFace: FONT.body, fontSize: 12, color: C.muted, bold: true, rotate: 270, margin: 0 });
  label(s, '绿区：可自动化', x0 + 4.75, y0 + 0.28, 2.2, C.teal, 15, { bold: true, align: 'center' });
  label(s, '黄区：副驾驶', x0 + 0.25, y0 + 3.12, 2.25, C.amber, 14, { bold: true, align: 'center' });
  label(s, '红区：人做最终决定', x0 + 5.0, y0 + 2.03, 2.9, C.coral, 14, { bold: true, align: 'center' });
  const tags = [
    ['客服草稿', 7.05, 2.72, C.teal], ['报表摘要', 8.85, 3.15, C.teal], ['知识问答', 6.25, 3.05, C.teal],
    ['采购预测', 3.0, 4.02, C.amber], ['质检初筛', 4.72, 4.7, C.amber], ['人事 / 财务审批', 8.35, 4.78, C.coral], ['对外承诺', 6.65, 5.2, C.coral]
  ];
  tags.forEach((it) => pill(s, it[0], it[1], it[2], Math.max(1.1, it[0].length * 0.22), C.navy, it[3], 10));
  s.addText('先让 AI 做副驾驶，再根据评测结果逐步放权。', { x: 2.1, y: 6.28, w: 8.85, h: 0.3, fontFace: FONT.body, fontSize: 17, bold: true, color: C.indigo, align: 'center', margin: 0 });
  addFooter(s, 13);
  addNotes(s, '场景选择不要从“模型有多强”开始，而要从“任务是否重复、风险是否可控、数据是否可用”开始。高风险任务让 AI 做草稿或提醒，人做最终决定。');
}

// 14 — maturity path
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '落地路线：先做副驾驶，再做流程，再做系统', '13 · 企业落地');
  s.addText('每一步都要有基线、有评测、有回滚；不要一开始就追求“全自动”。', { x: 0.78, y: 1.42, w: 11.5, h: 0.32, fontFace: FONT.body, fontSize: 16, color: C.ink, bold: true, margin: 0 });
  const stairs = [
    { x: 0.95, y: 4.92, w: 3.1, h: 0.72, n: '01', title: '副驾驶', sub: '写草稿、查资料、做摘要', c: C.ice },
    { x: 4.35, y: 3.95, w: 3.1, h: 1.69, n: '02', title: '流程', sub: '自动分派、校验、提醒\n人保留关键节点', c: C.teal },
    { x: 7.75, y: 2.95, w: 3.1, h: 2.69, n: '03', title: '系统', sub: '接入业务系统\n按权限行动\n持续评测与审计', c: C.indigo }
  ];
  stairs.forEach((it) => { s.addShape(pptx.ShapeType.roundRect, { x: it.x, y: it.y, w: it.w, h: it.h, rectRadius: 0.13, fill: { color: it.c }, line: { color: C.white, width: 1.1 } }); label(s, it.n, it.x + 0.25, it.y + 0.22, 0.55, C.navy, 15, { bold: true }); label(s, it.title, it.x + 0.92, it.y + 0.2, 1.65, C.white, 20, { bold: true }); label(s, it.sub, it.x + 0.3, it.y + 0.68, it.w - 0.6, C.white, 13, { h: it.h - 0.7 }); });
  line(s, 1.08, 5.96, 10.7, 5.96, C.line, 1.2, 'dash');
  ['基线 KPI', '可信数据', '离线评测', '小流量', '复盘'].forEach((t, i) => { const x = 1.12 + i * 2.35; dot(s, x, 6.24, 0.16, [C.coral, C.teal, C.indigo, C.amber, C.coral][i]); label(s, t, x + 0.25, 6.18, 1.0, C.muted, 11, { bold: true }); });
  addFooter(s, 14);
  addNotes(s, '落地成熟度是逐级放权，而不是一次性自动化。先从副驾驶开始建立评测基线，再把稳定的动作放进流程，最后才考虑系统级行动。');
}

// 15 — management actions
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '把原理变成管理动作：六个问题决定成败', '14 · 管理清单');
  const centerX = 6.05, centerY = 3.45;
  s.addShape(pptx.ShapeType.ellipse, { x: centerX - 1.0, y: centerY - 0.7, w: 2.0, h: 1.4, fill: { color: C.indigo }, line: { color: C.indigo } });
  label(s, 'AI 项目', centerX - 0.72, centerY - 0.18, 1.45, C.white, 21, { align: 'center', bold: true });
  const ring = [
    { x: 1.0, y: 1.65, t: '选场景', sub: '哪里最重复？', c: C.ice },
    { x: 4.0, y: 1.02, t: '给数据', sub: '它凭什么知道？', c: C.teal },
    { x: 8.25, y: 1.02, t: '设护栏', sub: '哪些事不能做？', c: C.amber },
    { x: 10.2, y: 3.55, t: '算 ROI', sub: '省下的钱够不够？', c: C.coral },
    { x: 7.95, y: 5.05, t: '留人工', sub: '错了谁发现？', c: C.indigo },
    { x: 2.1, y: 5.05, t: '复盘', sub: '数据如何变好？', c: C.mint }
  ];
  ring.forEach((it) => { line(s, centerX, centerY, it.x + 0.78, it.y + 0.5, C.line, 1.1, 'dash'); s.addShape(pptx.ShapeType.roundRect, { x: it.x, y: it.y, w: 1.56, h: 0.98, rectRadius: 0.12, fill: { color: it.c }, line: { color: it.c } }); label(s, it.t, it.x + 0.12, it.y + 0.2, 1.32, C.navy, 16, { align: 'center', bold: true }); label(s, it.sub, it.x + 0.08, it.y + 0.57, 1.4, C.navy, 10.5, { align: 'center' }); });
  s.addText('把“理解原理”翻译成流程：每个 AI 项目都要能回答这六个问题。', { x: 1.3, y: 6.47, w: 10.6, h: 0.28, fontFace: FONT.body, fontSize: 16, color: C.indigo, bold: true, align: 'center', margin: 0 });
  addFooter(s, 15);
  addNotes(s, '这页把技术原理翻译成管理动作：没有数据就没有可靠推理，没有护栏就没有可控自动化，没有 ROI 就没有持续投入。');
}

// 16 — close
{
  const s = pptx.addSlide(); addBg(s, C.navy);
  s.addText('中心思想', { x: 0.8, y: 0.72, w: 2.0, h: 0.3, fontFace: FONT.body, fontSize: 12, bold: true, charSpacing: 2, color: C.ice, margin: 0 });
  s.addText('把 AI 放在\n“可验证的工位”', { x: 0.8, y: 1.35, w: 6.35, h: 1.55, fontFace: FONT.head, fontSize: 34, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  s.addText('懂原理 → 会算账 → 设边界 → 留人复核 → 持续迭代', { x: 0.83, y: 3.38, w: 6.15, h: 0.35, fontFace: FONT.body, fontSize: 18, color: C.mint, bold: true, margin: 0, fit: 'shrink' });
  const qs = [
    ['它凭什么知道？', '数据和上下文是否可信？', C.ice],
    ['错了谁发现？', '有没有评测、复核和回滚？', C.amber],
    ['省下的钱够吗？', '单位成本是否小于业务价值？', C.coral]
  ];
  qs.forEach((it, i) => { const x = 7.65; const y = 1.2 + i * 1.45; card(s, x, y, 4.5, 1.02, { fill: C.indigo, line: it[2], lineWidth: 1.4 }); dot(s, x + 0.28, y + 0.28, 0.42, it[2]); label(s, String(i + 1), x + 0.28, y + 0.37, 0.42, C.navy, 13, { align: 'center', bold: true }); label(s, it[0], x + 0.9, y + 0.2, 2.9, C.white, 17, { bold: true }); label(s, it[1], x + 0.9, y + 0.61, 3.2, C.darkMuted, 11.5); });
  s.addText('AI 是概率引擎和放大器，不是责任承担者。\n企业竞争力来自“数据 + 流程 + 评测 + 人”的闭环。', { x: 0.83, y: 5.45, w: 6.3, h: 0.7, fontFace: FONT.body, fontSize: 17, color: C.white, bold: true, margin: 0, fit: 'shrink' });
  s.addText('谢谢', { x: 0.83, y: 6.7, w: 1.0, h: 0.25, fontFace: FONT.head, fontSize: 15, color: C.darkMuted, margin: 0 });
  addFooter(s, 16, true);
  addNotes(s, '结尾留三个问题：它凭什么知道？错了谁发现？省下的钱够吗？这三个问题能把 AI 从演示效果拉回到企业治理和经营结果。');
}

pptx.writeFile({ fileName: '/Users/minster/code/ai/ai-course/ai_from_principles_to_practice.pptx' });

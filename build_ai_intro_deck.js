const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');
const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI';
pptx.company = 'AI Course';
pptx.subject = '神经网络、Transformer、GPT 与 RSI';
pptx.title = '看懂 AI，才能用好 AI';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'Hiragino Sans GB W6', bodyFontFace: 'Hiragino Sans GB W3', lang: 'zh-CN' };

const FONT = { head: 'Hiragino Sans GB W6', body: 'Hiragino Sans GB W3' };
const C = {
  ink: '10212B', navy: '0B1020', indigo: '1A2454', teal: '1B998B', mint: '34D399',
  ice: '6EE7F9', coral: 'FF7A59', amber: 'FFC857', red: 'E85D75', white: 'FFFFFF',
  paper: 'F7FAFC', pale: 'EEF5F5', line: 'D6E2E6', muted: '637381', darkMuted: 'A9B5C6',
  soft: 'E8EFF2', greenBg: 'E8FFF5', yellowBg: 'FFF6DD'
};

function addBg(slide, color = C.paper) { slide.background = { color }; }

function addTitle(slide, title, kicker, dark = false) {
  slide.addText(kicker.toUpperCase(), { x: 0.65, y: 0.32, w: 6.5, h: 0.23, fontFace: FONT.body, fontSize: 10.5, bold: true, charSpacing: 1.4, color: dark ? C.ice : C.teal, margin: 0 });
  slide.addText(title, { x: 0.65, y: 0.64, w: 12, h: 0.54, fontFace: FONT.head, fontSize: 26, bold: true, color: dark ? C.white : C.ink, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.line, { x: 0.65, y: 1.3, w: 1.15, h: 0, line: { color: dark ? C.ice : C.teal, width: 2.8 } });
}

function addFooter(slide, n, dark = false) {
  slide.addText(String(n).padStart(2, '0'), { x: 12.25, y: 7.08, w: 0.45, h: 0.18, fontFace: FONT.body, fontSize: 9, bold: true, color: dark ? C.darkMuted : C.muted, align: 'right', margin: 0 });
}

function addNotes(slide, text) { slide.addNotes(text); }

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: opts.fill || C.white, transparency: opts.transparency || 0 },
    line: { color: opts.line || C.line, width: opts.lineWidth || 1, transparency: opts.lineTransparency || 0 },
    shadow: opts.shadow ? { type: 'outer', color: '7D8A91', blur: 1, angle: 45, distance: 1, opacity: 0.12 } : undefined
  });
}

function label(slide, text, x, y, w, color = C.muted, fs = 12, opts = {}) {
  slide.addText(text, { x, y, w, h: opts.h || 0.25, fontFace: opts.fontFace || FONT.body, fontSize: fs, bold: !!opts.bold, italic: !!opts.italic, color, align: opts.align || 'left', valign: opts.valign || 'mid', margin: opts.margin ?? 0, fit: 'shrink' });
}

function pill(slide, text, x, y, w, fill, color = C.navy, fs = 11) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.31, rectRadius: 0.15, fill: { color: fill }, line: { color: fill } });
  label(slide, text, x, y + 0.05, w, color, fs, { align: 'center', bold: true });
}

function dot(slide, x, y, r, fill, lineColor = fill) {
  slide.addShape(pptx.ShapeType.ellipse, { x, y, w: r, h: r, fill: { color: fill }, line: { color: lineColor, width: 1 } });
}

function line(slide, x1, y1, x2, y2, color = C.line, width = 1.5, dash = 'solid', end = 'none') {
  slide.addShape(pptx.ShapeType.line, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color, width, dash, endArrowType: end } });
}

function arrow(slide, x, y, w = 0.36, h = 0.24, color = C.teal) {
  slide.addShape(pptx.ShapeType.chevron, { x, y, w, h, fill: { color }, line: { color } });
}

function step(slide, x, y, n, title, sub, fill) {
  dot(slide, x, y, 0.37, fill);
  label(slide, String(n), x, y + 0.07, 0.37, C.navy, 12, { align: 'center', bold: true });
  label(slide, title, x + 0.55, y + 0.01, 1.85, C.ink, 15, { bold: true });
  label(slide, sub, x + 0.55, y + 0.29, 2.8, C.muted, 11.5);
}

function addGrid(slide, x, y, cell, rows, cols, highlight = null) {
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const active = highlight && r >= highlight.r && r < highlight.r + highlight.h && c >= highlight.c && c < highlight.c + highlight.w;
      slide.addShape(pptx.ShapeType.rect, { x: x + c * cell, y: y + r * cell, w: cell - 0.01, h: cell - 0.01, fill: { color: active ? C.coral : ((r + c) % 3 === 0 ? C.ice : C.soft), transparency: active ? 0 : 18 }, line: { color: C.white, transparency: 35, width: 0.5 } });
    }
  }
}

// 01 — Cover
{
  const s = pptx.addSlide(); addBg(s, C.navy);
  s.addText('看懂 AI，\n才能用好 AI', { x: 0.8, y: 1.0, w: 5.8, h: 1.45, fontFace: FONT.head, fontSize: 34, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  s.addText('从神经网络、Transformer、GPT 到 RSI', { x: 0.83, y: 2.85, w: 6.0, h: 0.32, fontFace: FONT.body, fontSize: 17, color: C.ice, margin: 0 });
  s.addText('AI 不是魔法，是一套从例子里调参数，再对新输入做计算的系统。', { x: 0.83, y: 3.55, w: 5.8, h: 0.55, fontFace: FONT.body, fontSize: 15, color: C.darkMuted, margin: 0, fit: 'shrink' });
  const flow = [{ x: 7.4, title: '输入', sub: '图片 / 文字 / 数据', c: C.ice }, { x: 9.45, title: '模型', sub: '参数 + 计算结构', c: C.teal }, { x: 11.5, title: '输出', sub: '预测 / 生成 / 动作', c: C.coral }];
  flow.forEach((it, i) => { card(s, it.x, 2.2, 1.5, 1.55, { fill: C.indigo, line: it.c, lineWidth: 1.5 }); dot(s, it.x + 0.56, 2.48, 0.38, it.c); label(s, it.title, it.x + 0.15, 3.0, 1.2, C.white, 16, { align: 'center', bold: true }); label(s, it.sub, it.x + 0.1, 3.38, 1.3, C.darkMuted, 10.5, { align: 'center' }); if (i < flow.length - 1) arrow(s, it.x + 1.58, 2.85, 0.38, 0.26, it.c); });
  label(s, '训练：调整参数', 7.45, 4.55, 1.45, C.ice, 12, { align: 'center', bold: true });
  label(s, '推理：使用参数', 9.5, 4.55, 1.45, C.mint, 12, { align: 'center', bold: true });
  label(s, '复核：检查结果', 11.55, 4.55, 1.45, C.coral, 12, { align: 'center', bold: true });
  addFooter(s, 1, true); addNotes(s, '开场先给路线：先理解模型怎样算，再理解它怎样学，最后看为什么还需要工具、评估和 RSI。');
}

// 02 — Function and training/inference
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, 'AI 的核心：先找到一组参数，再用它处理新输入', '01 · 训练与推理');
  s.addText('把模型写成一个函数：', { x: 0.8, y: 1.55, w: 2.5, h: 0.25, fontFace: FONT.body, fontSize: 16, color: C.muted, margin: 0 });
  s.addText('ŷ = f(x; θ)', { x: 3.1, y: 1.36, w: 3.1, h: 0.55, fontFace: FONT.head, fontSize: 28, bold: true, italic: true, color: C.indigo, margin: 0, align: 'center' });
  card(s, 0.8, 2.15, 5.7, 3.65, { fill: C.white, line: C.ice, lineWidth: 1.6, shadow: true });
  pill(s, '训练：参数会变', 1.1, 2.48, 1.5, C.ice, C.navy, 12);
  step(s, 1.12, 3.02, 1, '输入 + 答案', '看历史样本', C.ice);
  step(s, 1.12, 3.64, 2, '前向预测', '得到 ŷ', C.teal);
  step(s, 1.12, 4.26, 3, '计算损失', '比较 ŷ 和答案', C.amber);
  step(s, 1.12, 4.88, 4, '反传求梯度 → 优化器更新', '学习率决定步长', C.coral);
  line(s, 1.3, 5.32, 5.7, 5.32, C.amber, 2, 'dash', 'triangle');
  label(s, '训练是一条反馈回路', 3.0, 5.42, 2.1, C.amber, 12, { align: 'center', bold: true });
  card(s, 6.85, 2.15, 5.7, 3.65, { fill: C.white, line: C.coral, lineWidth: 1.6, shadow: true });
  pill(s, '推理：参数通常不变', 7.15, 2.48, 1.72, C.coral, C.white, 12);
  card(s, 7.18, 3.35, 1.35, 0.78, { fill: C.pale, line: C.line }); label(s, '新输入 x', 7.18, 3.62, 1.35, C.indigo, 14, { align: 'center', bold: true });
  arrow(s, 8.75, 3.62, 0.45, 0.28, C.teal);
  card(s, 9.35, 3.2, 1.55, 1.08, { fill: C.indigo, line: C.indigo }); label(s, 'fθ', 9.35, 3.54, 1.55, C.white, 22, { align: 'center', bold: true });
  arrow(s, 11.1, 3.62, 0.45, 0.28, C.coral);
  card(s, 11.7, 3.35, 0.78, 0.78, { fill: C.coral, line: C.coral }); label(s, 'ŷ', 11.7, 3.62, 0.78, C.white, 18, { align: 'center', bold: true });
  label(s, '推理是在代入公式，不是在现场重新训练。', 7.2, 4.85, 4.75, C.muted, 14, { align: 'center', bold: true });
  addFooter(s, 2); addNotes(s, '这页把训练和推理放在同一张图中。训练是“预测—比较—更新”的反馈回路；推理是参数固定后，给新输入做一次前向计算。');
}

// 03 — Neuron
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '一个神经元：先混合线索，再决定是否放行', '02 · 神经网络');
  card(s, 0.75, 1.7, 6.0, 4.8, { fill: C.white, line: C.line, shadow: true });
  label(s, '具体算一遍', 1.05, 2.0, 1.6, C.indigo, 18, { bold: true });
  label(s, 'z = 0.8x₁ + 0.1x₂ − 0.4x₃ + b', 1.05, 2.47, 4.9, C.ink, 19, { bold: true });
  const vals = [['客户说“太贵”', 'x₁ = 1.0', '× 0.8', C.ice], ['近 30 天退货率', 'x₂ = 0.5', '× 0.1', C.teal], ['季节因素', 'x₃ = 0.2', '× −0.4', C.coral]];
  vals.forEach((it, i) => { const y = 3.13 + i * 0.62; pill(s, it[0], 1.05, y, 1.75, C.pale, C.ink, 10.5); label(s, it[1], 3.0, y + 0.06, 1.0, C.indigo, 12, { bold: true }); label(s, it[2], 4.25, y + 0.06, 0.85, it[3], 12, { bold: true }); });
  label(s, 'z = 0.8 + 0.05 − 0.08 + b', 1.05, 5.18, 4.2, C.ink, 15, { bold: true });
  label(s, 'ReLU(z) = max(0, z)', 1.05, 5.6, 3.5, C.coral, 15, { bold: true });
  card(s, 7.1, 1.7, 5.45, 4.8, { fill: C.white, line: C.coral, shadow: true });
  label(s, '为什么需要激活函数？', 7.4, 2.0, 3.5, C.indigo, 18, { bold: true });
  line(s, 7.85, 5.35, 11.9, 5.35, C.line, 1.2); line(s, 9.55, 5.75, 9.55, 2.95, C.line, 1.2);
  line(s, 8.1, 5.35, 9.55, 5.35, C.coral, 3.5); line(s, 9.55, 5.35, 11.25, 3.15, C.teal, 3.5);
  label(s, 'z < 0：输出 0', 7.7, 5.55, 1.55, C.coral, 12, { bold: true }); label(s, 'z > 0：原样放行', 10.1, 3.0, 1.85, C.teal, 12, { bold: true });
  label(s, '没有非线性，多层线性计算仍可合并成一层。', 7.48, 6.0, 4.7, C.muted, 13, { align: 'center', bold: true });
  addFooter(s, 3); addNotes(s, '把公式落到一个数值过程：输入乘权重、加总、经过 ReLU。激活函数提供折点，让多层网络能够表达弯曲的关系。');
}

// 04 — CNN derived from vision
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '卷积网络：把“人看局部线索”总结成结构', '03 · 网络结构');
  s.addText('人看一张图时，不需要同时处理每个像素；先找局部线索，再把线索组合成整体。卷积网络把这个现象写进了结构。', { x: 0.8, y: 1.52, w: 11.8, h: 0.42, fontFace: FONT.body, fontSize: 15.5, color: C.ink, margin: 0, fit: 'shrink' });
  card(s, 0.8, 2.18, 3.2, 3.75, { fill: C.white, line: C.ice, lineWidth: 1.5, shadow: true });
  label(s, '人类视觉：先看一小块', 1.1, 2.48, 2.55, C.indigo, 16, { bold: true, align: 'center' });
  addGrid(s, 1.28, 3.08, 0.34, 7, 7, { r: 2, c: 3, h: 3, w: 3 });
  label(s, '局部感受野', 1.25, 5.58, 2.45, C.coral, 14, { align: 'center', bold: true });
  arrow(s, 4.28, 3.8, 0.48, 0.32, C.teal);
  card(s, 4.9, 2.18, 3.35, 3.75, { fill: C.white, line: C.teal, lineWidth: 1.5, shadow: true });
  label(s, '工程化结构：卷积核扫描', 5.2, 2.48, 2.75, C.indigo, 16, { bold: true, align: 'center' });
  addGrid(s, 5.35, 3.13, 0.33, 5, 5, { r: 1, c: 1, h: 3, w: 3 });
  label(s, '3×3 小窗口反复共享同一组权重', 5.18, 5.38, 2.85, C.teal, 12, { align: 'center', bold: true, h: 0.43 });
  arrow(s, 8.55, 3.8, 0.48, 0.32, C.teal);
  card(s, 9.18, 2.18, 3.4, 3.75, { fill: C.white, line: C.coral, lineWidth: 1.5, shadow: true });
  label(s, '逐层组合：从边缘到整体', 9.48, 2.48, 2.8, C.indigo, 16, { bold: true, align: 'center' });
  const stages = [['边缘', C.ice], ['角点 / 纹理', C.teal], ['眼睛 / 轮廓', C.indigo], ['物体', C.coral]];
  stages.forEach((it, i) => { const y = 3.1 + i * 0.55; pill(s, it[0], 9.55 + i * 0.16, y, 1.2 + i * 0.25, it[1], i === 2 ? C.white : C.navy, 11); if (i < stages.length - 1) arrow(s, 10.9 + i * 0.16, y + 0.08, 0.28, 0.18, C.line); });
  label(s, '结论：结构不是凭空发明，而是对现实现象的工程化总结。', 2.05, 6.35, 9.2, C.indigo, 16, { align: 'center', bold: true });
  addFooter(s, 4); addNotes(s, '重点讲清用户要求的结论：网络结构来自对现实规律的总结。人类视觉先看局部，再组合整体；CNN 用局部感受野、共享卷积核、池化和逐层特征组合来实现类似的计算偏置。');
}

// 05 — CNN pipeline
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '卷积网络怎样一步步把图片变成分类结果', '04 · 从现象到实现');
  const items = [['28×28 图片', '原始像素', C.ice], ['3×3 卷积', '识别局部边缘', C.teal], ['特征图', '边缘位置保留下来', C.indigo], ['池化', '压缩尺寸，保留强信号', C.amber], ['全连接', '综合线索做分类', C.coral]];
  items.forEach((it, i) => { const x = 0.7 + i * 2.5; card(s, x, 2.35, 2.05, 2.25, { fill: C.white, line: it[2], lineWidth: 1.5, shadow: true }); dot(s, x + 0.83, 2.65, 0.38, it[2]); label(s, String(i + 1), x + 0.83, 2.73, 0.38, C.navy, 12, { align: 'center', bold: true }); label(s, it[0], x + 0.15, 3.25, 1.75, C.ink, 16, { align: 'center', bold: true }); label(s, it[1], x + 0.18, 3.82, 1.7, C.muted, 11.5, { align: 'center', h: 0.38 }); if (i < items.length - 1) arrow(s, x + 2.12, 3.35, 0.3, 0.22, C.line); });
  label(s, '局部 → 组合 → 压缩 → 判断', 3.65, 5.4, 5.8, C.indigo, 22, { align: 'center', bold: true });
  label(s, '不同任务采用不同结构：图像常利用空间局部性；文本需要处理序列和远距离关系。', 1.0, 6.2, 11.2, C.muted, 14, { align: 'center', bold: true });
  addFooter(s, 5); addNotes(s, '这一页作为上一页的“工程流水线”补充。让听众看到结构里的每个部件为什么存在，而不是只记住 CNN 三个字。');
}

// 06 — Transformer attention
{
  const s = pptx.addSlide(); addBg(s, C.navy); addTitle(s, 'Transformer：让每个词计算“我该参考谁”', '05 · 注意力', true);
  s.addText('双向编码器示意：小明把书放进抽屉，因为它太重。这里“它”可以参考整句上下文。', { x: 0.78, y: 1.5, w: 11.7, h: 0.34, fontFace: FONT.body, fontSize: 15.5, color: C.darkMuted, margin: 0 });
  const toks = ['小明', '把', '书', '放进', '抽屉', '因为', '它', '太重'];
  toks.forEach((t, i) => { const x = 0.95 + i * 1.02; const emph = t === '它'; card(s, x, 2.35, 0.82, 0.5, { fill: emph ? C.coral : C.indigo, line: emph ? C.coral : C.ice, lineWidth: 1 }); label(s, t, x, 2.5, 0.82, C.white, 14, { align: 'center', bold: emph }); });
  label(s, '“它”', 7.08, 3.02, 0.82, C.coral, 15, { align: 'center', bold: true });
  const refs = [['书', '0.70', C.coral], ['抽屉', '0.12', C.amber], ['太重', '0.10', C.ice], ['其他', '0.08', C.teal]];
  refs.forEach((it, i) => { const y = 3.45 + i * 0.42; label(s, it[0], 5.25, y, 0.75, C.white, 12, { align: 'right', bold: true }); s.addShape(pptx.ShapeType.roundRect, { x: 6.2, y: y + 0.06, w: 2.55 * Number(it[1]), h: 0.16, rectRadius: 0.08, fill: { color: it[2] }, line: { color: it[2] } }); label(s, it[1], 8.95, y, 0.55, it[2], 11, { bold: true }); });
  card(s, 9.7, 3.14, 2.75, 1.75, { fill: C.indigo, line: C.ice, lineWidth: 1.2 });
  label(s, 'Q · 我在找什么', 9.95, 3.42, 2.2, C.ice, 13, { bold: true });
  label(s, 'K · 谁最匹配', 9.95, 3.82, 2.2, C.mint, 13, { bold: true });
  label(s, 'V · 带来什么信息', 9.95, 4.22, 2.2, C.coral, 13, { bold: true });
  s.addText('Q 与 K 打分 → softmax 变成比例 → 按比例汇总 V；数字仅为教学示意。', { x: 1.1, y: 5.7, w: 11.0, h: 0.28, fontFace: FONT.body, fontSize: 15, color: C.amber, align: 'center', bold: true, margin: 0 });
  addFooter(s, 6, true); addNotes(s, '点击节奏：先出现句子，再强调“它”，最后出现权重和 Q/K/V。本页明确是双向编码器示意；GPT 的 decoder-only 因果注意力会屏蔽未来位置。注意力比例是教学示意，不是真实模型读出的测量值。');
}

// 07 — Transformer block
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, 'Transformer 的一层：先交换信息，再加工表示', '06 · 架构');
  const blocks = [['token + 位置', '顺序信息', C.ice], ['多头注意力', '跨位置交换信息', C.teal], ['残差 + 归一化', '保留并稳定信号', C.indigo], ['前馈网络', '逐位置非线性加工', C.coral], ['重复多层', '形成上下文表示', C.amber]];
  blocks.forEach((it, i) => { const x = 0.7 + i * 2.5; card(s, x, 2.25, 2.08, 2.0, { fill: C.white, line: it[2], lineWidth: 1.5, shadow: true }); dot(s, x + 0.84, 2.55, 0.38, it[2]); label(s, String(i + 1), x + 0.84, 2.63, 0.38, C.navy, 12, { align: 'center', bold: true }); label(s, it[0], x + 0.16, 3.16, 1.76, C.ink, 15, { align: 'center', bold: true }); label(s, it[1], x + 0.15, 3.67, 1.78, C.muted, 11.5, { align: 'center', h: 0.35 }); if (i < blocks.length - 1) arrow(s, x + 2.13, 3.05, 0.3, 0.22, C.line); });
  card(s, 2.25, 5.15, 8.85, 0.78, { fill: C.indigo, line: C.indigo });
  label(s, '注意力负责“位置之间互相看”，FFN 负责“每个位置自己加工”。', 2.55, 5.41, 8.25, C.white, 16, { align: 'center', bold: true });
  label(s, '训练时可以并行处理已知序列；生成时仍要一个 token 接一个 token。', 1.1, 6.35, 11.0, C.muted, 14, { align: 'center', bold: true });
  addFooter(s, 7); addNotes(s, 'Transformer 不是只有注意力。完整一层还包括位置处理、残差、归一化和前馈网络。用这一页把部件关系讲完整。');
}

// 08 — GPT versus Transformer
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, 'GPT 相对 Transformer：进步主要在“选择 + 训练 + 规模化”', '07 · GPT');
  card(s, 0.8, 1.75, 5.5, 4.55, { fill: C.white, line: C.line, shadow: true });
  pill(s, 'Transformer：架构家族', 1.1, 2.08, 2.0, C.indigo, C.white, 12);
  const tr = [['结构', '可以是编码器、解码器，或两者组合'], ['注意力', '让位置之间交换上下文信息'], ['原始论文', '主要验证序列转换与机器翻译']];
  tr.forEach((it, i) => { const y = 2.85 + i * 0.72; label(s, it[0], 1.12, y, 0.7, C.indigo, 13, { bold: true }); label(s, it[1], 2.0, y, 3.85, C.muted, 13, { h: 0.38 }); });
  card(s, 6.95, 1.75, 5.55, 4.55, { fill: C.indigo, line: C.indigo, shadow: true });
  pill(s, 'GPT：一套生成式配方', 7.25, 2.08, 1.95, C.coral, C.white, 12);
  const gp = [['结构选择', '选择 decoder-only；因果遮罩继承自 Transformer'], ['预训练', 'GPT-1：无标签文本预训练，再做任务微调'], ['规模扩展', 'GPT-2/3：放大数据与模型，支持零 / 少样本'], ['后训练 / 系统', '指令偏好训练 + 工具 / 状态，塑造助手体验']];
  gp.forEach((it, i) => { const y = 2.78 + i * 0.65; dot(s, 7.28, y + 0.06, 0.18, [C.ice, C.teal, C.amber, C.coral][i]); label(s, it[0], 7.65, y, 1.2, C.white, 12.5, { bold: true }); label(s, it[1], 8.95, y, 2.95, C.darkMuted, 12.5, { h: 0.35 }); });
  s.addText('因果遮罩和自回归生成已在原始 Transformer 解码器中存在；GPT 的主线增量是预训练、迁移、规模化，以及后续助手对齐。', { x: 1.0, y: 6.45, w: 11.25, h: 0.35, fontFace: FONT.body, fontSize: 14.5, color: C.indigo, bold: true, align: 'center', margin: 0, fit: 'shrink' });
  addFooter(s, 8); addNotes(s, '用户要求讲清 GPT 相对于 Transformer 的进步。要避免说成“GPT 改进了注意力公式”：Transformer 是架构家族，GPT 的关键贡献是采用 decoder-only 生成结构，并配合 next-token 预训练、规模化训练和后续对齐，做成可通用的生成模型路线。来源：Transformer 原始论文 https://arxiv.org/html/1706.03762v7；GPT-1 https://cdn.openai.com/research-covers/language-unsupervised/language_understanding_paper.pdf；GPT-2 https://cdn.openai.com/better-language-models/language_models_are_unsupervised_multitask_learners.pdf；GPT-3 https://arxiv.org/abs/2005.14165；InstructGPT https://arxiv.org/abs/2203.02155。');
}

// 09 — GPT generation
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, 'GPT 怎样生成一段话：每一步都只决定下一个 token', '08 · 自回归生成');
  label(s, '提示词：客户投诉鞋底开胶，', 0.85, 1.65, 3.7, C.indigo, 17, { bold: true });
  const seq = ['建议', '先', '拍照', '并', '保留'];
  seq.forEach((t, i) => { const x = 0.95 + i * 1.32; card(s, x, 2.25, 1.05, 0.6, { fill: i === 0 ? C.coral : C.pale, line: i === 0 ? C.coral : C.line }); label(s, t, x, 2.45, 1.05, i === 0 ? C.white : C.ink, 15, { align: 'center', bold: true }); if (i < seq.length - 1) arrow(s, x + 1.1, 2.45, 0.27, 0.2, C.line); });
  label(s, '当前前缀', 0.95, 3.15, 1.1, C.muted, 12, { bold: true });
  const cand = [['建议', '42%', C.teal], ['联系', '28%', C.ice], ['直接', '16%', C.amber], ['其他', '14%', C.coral]];
  cand.forEach((it, i) => { const y = 3.58 + i * 0.42; label(s, it[0], 1.45, y, 0.62, C.ink, 12, { bold: true }); s.addShape(pptx.ShapeType.roundRect, { x: 2.25, y: y + 0.06, w: 3.2 * Number(it[1].replace('%', '')) / 42, h: 0.17, rectRadius: 0.08, fill: { color: it[2] }, line: { color: it[2] } }); label(s, it[1], 5.7, y, 0.55, it[2], 11, { bold: true }); });
  card(s, 7.25, 1.65, 5.2, 3.95, { fill: C.indigo, line: C.indigo, shadow: true });
  label(s, '一次生成的真实过程', 7.58, 2.0, 3.0, C.ice, 17, { bold: true });
  const gen = [['1', '处理提示词', 'token → embedding → Transformer'], ['2', '选择下一个 token', '贪心、采样、温度等策略'], ['3', '追加到上下文', '把“建议”放回输入'], ['4', '继续下一步', '直到结束或达到长度限制']];
  gen.forEach((it, i) => { const y = 2.72 + i * 0.62; dot(s, 7.62, y + 0.03, 0.26, [C.ice, C.teal, C.amber, C.coral][i]); label(s, it[0], 7.62, y + 0.08, 0.26, C.navy, 10, { align: 'center', bold: true }); label(s, it[1], 8.05, y, 1.2, C.white, 13, { bold: true }); label(s, it[2], 9.35, y, 2.65, C.darkMuted, 11.5, { h: 0.32 }); });
  label(s, '上下文是“这次能直接参考的桌面”，参数是长期训练得到的能力。', 1.0, 6.18, 11.2, C.indigo, 15, { align: 'center', bold: true });
  addFooter(s, 9); addNotes(s, 'GPT 生成不是一次性写完整答案，而是每步预测一个 token，再把它追加到上下文。可以点击逐步讲解；本页的条形是教学示意，不是实际模型输出。');
}

// 10 — Context limits
{
  const s = pptx.addSlide(); addBg(s, C.navy); addTitle(s, '上下文窗口为什么有限：算力、缓存、位置和有效利用都在约束它', '09 · Transformer 局限', true);
  const cards = [['注意力计算', 'n 个 token 的全注意力约有 n² 个位置配对。长度翻倍，配对约变四倍。', C.coral], ['KV cache', '历史 K/V 会复用，但 token 越长，缓存越大；生成每一步仍要访问历史。', C.amber], ['位置泛化', '能算出更远的位置，不代表模型训练过这种距离，更不代表一定用得好。', C.ice], ['有效利用', '资料放进窗口后，模型仍可能漏掉中间证据、多个条件或远处关系。', C.mint]];
  cards.forEach((it, i) => { const x = 0.8 + (i % 2) * 6.15; const y = 1.85 + Math.floor(i / 2) * 2.15; card(s, x, y, 5.45, 1.65, { fill: C.indigo, line: it[2], lineWidth: 1.3 }); dot(s, x + 0.3, y + 0.32, 0.36, it[2]); label(s, String(i + 1), x + 0.3, y + 0.4, 0.36, C.navy, 12, { align: 'center', bold: true }); label(s, it[0], x + 0.88, y + 0.23, 2.2, C.white, 16, { bold: true }); label(s, it[1], x + 0.88, y + 0.72, 4.1, C.darkMuted, 12.5, { h: 0.58 }); });
  s.addText('“支持更长输入”不等于“无限记忆”，也不等于“所有证据都能可靠综合”。', { x: 1.0, y: 6.45, w: 11.3, h: 0.3, fontFace: FONT.body, fontSize: 15.5, color: C.amber, bold: true, align: 'center', margin: 0 });
  addFooter(s, 10, true); addNotes(s, '这一页把上下文限制拆成四个独立原因。FlashAttention 可以优化内存访问，但不会让标准稠密注意力的配对计算、KV cache 和有效利用问题消失。来源：Transformer 原始论文 https://arxiv.org/html/1706.03762v7；FlashAttention https://arxiv.org/abs/2205.14135；Hugging Face KV cache https://huggingface.co/docs/transformers/cache_explanation；Lost in the Middle https://arxiv.org/abs/2307.03172。');
}

// 11 — Why RSI
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '为什么需要 RSI：单次改进太慢，人工试错无法覆盖所有环节', '10 · RSI 的动机');
  label(s, '先看痛点，再看闭环。RSI 是对迭代瓶颈的工程回应。', 0.8, 1.56, 11.8, C.ink, 15.5, { bold: true });
  const pain = [['模型参数静态', '训练结束后，系统不会因为一次失败自动重写自己。', C.coral], ['人工迭代慢', '数据、提示、代码、评测常由人逐项试错。', C.amber], ['历史越来越长', '实验记录在外部，但每次能读到的上下文有限。', C.ice], ['改进难以验证', '一个分数变好，不代表真实任务和旧能力都变好。', C.teal]];
  pain.forEach((it, i) => { const x = 0.8 + (i % 2) * 3.55; const y = 2.35 + Math.floor(i / 2) * 1.45; card(s, x, y, 3.1, 1.05, { fill: C.white, line: it[2], lineWidth: 1.4, shadow: true }); label(s, it[0], x + 0.2, y + 0.16, 2.55, C.indigo, 15, { bold: true }); label(s, it[1], x + 0.2, y + 0.52, 2.65, C.muted, 10.5, { h: 0.34 }); });
  arrow(s, 7.35, 3.62, 0.6, 0.35, C.coral);
  card(s, 8.25, 2.35, 4.25, 2.95, { fill: C.indigo, line: C.indigo, shadow: true });
  label(s, '需要一个持续改进闭环', 8.6, 2.75, 3.55, C.white, 20, { bold: true, align: 'center' });
  const needs = ['发现弱点', '提出候选改动', '实际跑实验', '外部评估后保留'];
  needs.forEach((t, i) => { const y = 3.38 + i * 0.48; dot(s, 8.86, y + 0.05, 0.16, [C.ice, C.teal, C.amber, C.coral][i]); label(s, t, 9.2, y, 2.6, C.darkMuted, 13, { bold: true }); });
  label(s, '通过评估的版本进入下一轮，成为下一轮改进的基础。', 1.2, 6.35, 10.7, C.indigo, 15.5, { align: 'center', bold: true });
  addFooter(s, 11); addNotes(s, 'RSI 前先讲为什么需要它：模型静态、人工试错慢、上下文和实验历史有限、评估不容易。RSI 试图把发现问题、提出改动、实验和筛选纳入持续循环。来源：RSI 概念综述 https://arxiv.org/html/2609.11873v3；AIDE² https://arxiv.org/html/2609.26457v1。');
}

// 12 — RSI loop
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, 'RSI：把“发现问题 → 验证改动”变成受控循环', '11 · RSI 原理');
  const nodes = [{ x: 0.95, y: 3.2, t: '发现弱点', sub: '失败日志 / 评测', c: C.ice }, { x: 3.6, y: 2.15, t: '提出改动', sub: '代码 / 数据 / 策略', c: C.teal }, { x: 6.55, y: 2.15, t: '沙盒实验', sub: '固定预算执行', c: C.indigo }, { x: 9.45, y: 3.2, t: '外部评估', sub: '留出任务 + 回归', c: C.coral }];
  nodes.forEach((it, i) => { s.addShape(pptx.ShapeType.ellipse, { x: it.x, y: it.y, w: 1.65, h: 1.65, fill: { color: it.c }, line: { color: C.white, width: 1.2 } }); label(s, it.t, it.x + 0.1, it.y + 0.48, 1.45, C.white, 16, { align: 'center', bold: true }); label(s, it.sub, it.x + 0.1, it.y + 1.0, 1.45, C.white, 10.5, { align: 'center', h: 0.28 }); if (i < nodes.length - 1) line(s, it.x + 1.65, it.y + 0.8, nodes[i + 1].x, nodes[i + 1].y + 0.8, C.teal, 2.4, 'solid', 'triangle'); });
  line(s, 10.25, 4.9, 2.0, 4.9, C.amber, 2.2, 'dash', 'triangle');
  label(s, '通过才保留；失败就回滚或换候选', 4.2, 5.02, 4.6, C.amber, 13, { align: 'center', bold: true });
  card(s, 2.1, 5.75, 9.2, 0.55, { fill: C.white, line: C.line });
  label(s, '边界：权限、沙盒、独立评估、版本记录、可回滚。改进代码不等于基础模型权重自动更新。', 2.35, 5.94, 8.7, C.indigo, 12.5, { align: 'center', bold: true });
  addFooter(s, 12); addNotes(s, '点击节奏：四个节点依次出现，再出现回路和护栏。RSI 是工程闭环，不是一个新的神经网络层。改进对象可以是代理代码、工具、数据、模型参数或改进策略，必须说明具体改了哪一层。来源：RSI 概念综述 https://arxiv.org/html/2609.11873v3；AIDE² https://arxiv.org/html/2609.26457v1；AlphaEvolve 官方介绍 https://deepmind.google/blog/alphaevolve-a-gemini-powered-coding-agent-for-designing-advanced-algorithms/。');
}

// 13 — Limits
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, 'Transformer + RSI 的边界：循环变快，不等于自动变可靠', '12 · 局限性');
  card(s, 0.8, 1.75, 5.7, 4.65, { fill: C.white, line: C.indigo, lineWidth: 1.5, shadow: true });
  pill(s, 'Transformer 仍受限', 1.1, 2.08, 1.65, C.indigo, C.white, 12);
  const tl = [['窗口', '输入和生成总量有限'], ['成本', '长上下文、KV cache 和注意力变贵'], ['利用', '放进去不等于找对、记牢、综合对'], ['事实', '生成流畅不等于真实正确']];
  tl.forEach((it, i) => { const y = 2.84 + i * 0.68; dot(s, 1.13, y + 0.06, 0.18, [C.coral, C.amber, C.ice, C.teal][i]); label(s, it[0], 1.5, y, 0.65, C.indigo, 13, { bold: true }); label(s, it[1], 2.35, y, 3.45, C.muted, 12.5, { h: 0.35 }); });
  card(s, 6.85, 1.75, 5.65, 4.65, { fill: C.indigo, line: C.indigo, shadow: true });
  pill(s, 'RSI 仍受限', 7.15, 2.08, 1.35, C.coral, C.white, 12);
  const rl = [['评测投机', '分数变好，不等于真实任务变好'], ['误判传播', '错误候选被接受，会进入下一轮'], ['资源成本', '候选越多，训练和验证越贵'], ['能力边界', '改代理代码不会自动扩大模型窗口']];
  rl.forEach((it, i) => { const y = 2.84 + i * 0.68; dot(s, 7.18, y + 0.06, 0.18, [C.amber, C.coral, C.ice, C.mint][i]); label(s, it[0], 7.55, y, 0.9, C.white, 13, { bold: true }); label(s, it[1], 8.65, y, 3.25, C.darkMuted, 12.5, { h: 0.35 }); });
  s.addText('真正的进步需要：独立评估 + 可追溯版本 + 外部事实 + 人类责任。', { x: 1.1, y: 6.62, w: 11.1, h: 0.27, fontFace: FONT.body, fontSize: 15.5, color: C.indigo, align: 'center', bold: true, margin: 0 });
  addFooter(s, 13); addNotes(s, '这页把 Transformer 的限制和 RSI 的限制分开，避免“加一个自我改进循环就解决所有问题”的误解。尤其强调固定基础模型的 RSI 可以提高工作流，但不会自动改变窗口或模型权重。');
}

// 14 — Cost
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, 'AI 项目的成本：模型只是账单的一部分', '13 · 成本');
  label(s, '真正的成本来自整条链路：数据、训练、推理、集成、人工复核和错误代价。', 0.8, 1.55, 11.8, C.ink, 15.5, { bold: true });
  const rows = [['数据', '清洗、标注、知识库与更新', C.ice, 3.0], ['训练 / 微调', '一次性或周期性的实验算力', C.indigo, 2.35], ['推理', '输入输出 token、并发、长上下文', C.coral, 3.75], ['集成治理', '权限、监控、评测、日志和回滚', C.teal, 2.95], ['人工复核', '高风险结果需要人做最后判断', C.amber, 2.45], ['错误代价', '返工、客户损失、合规风险', C.red, 3.3]];
  rows.forEach((it, i) => { const y = 2.18 + i * 0.58; label(s, it[0], 0.95, y + 0.06, 1.1, C.ink, 13, { bold: true }); s.addShape(pptx.ShapeType.roundRect, { x: 2.15, y, w: it[3], h: 0.28, rectRadius: 0.12, fill: { color: it[2] }, line: { color: it[2] } }); label(s, it[1], 6.25, y + 0.05, 4.9, C.muted, 12); });
  card(s, 9.85, 2.05, 2.45, 3.72, { fill: C.white, line: C.indigo, shadow: true });
  label(s, '先算单位经济性', 10.1, 2.45, 1.95, C.indigo, 17, { align: 'center', bold: true });
  label(s, '每次调用成本\n× 调用次数\n+ 集成与复核\n+ 错误代价', 10.18, 3.3, 1.75, C.ink, 14, { align: 'center', bold: true, h: 1.45 });
  label(s, '再决定是否扩大调用量。', 10.1, 5.18, 1.95, C.coral, 12, { align: 'center', bold: true });
  addFooter(s, 14); addNotes(s, '这页保留原稿的经营视角，但把示意条明确标成成本项，不伪装成真实占比。重点是模型价格之外还有数据、治理、人工和错误成本。');
}

// 15 — Scenario matrix
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '企业选场景：先看重复度，再看可控性', '14 · 场景选择');
  label(s, '优先选择高频、可量化、可复核、数据已经存在的任务。', 0.8, 1.55, 8.0, C.ink, 15.5, { bold: true });
  const x0 = 1.85, y0 = 2.15, w = 8.95, h = 3.55;
  s.addShape(pptx.ShapeType.rect, { x: x0, y: y0, w, h, fill: { color: C.white }, line: { color: C.line, width: 1.2 } });
  s.addShape(pptx.ShapeType.rect, { x: x0 + w / 2, y: y0, w: w / 2, h: h / 2, fill: { color: C.greenBg, transparency: 5 }, line: { color: C.greenBg, transparency: 100 } });
  s.addShape(pptx.ShapeType.rect, { x: x0, y: y0 + h / 2, w: w / 2, h: h / 2, fill: { color: C.yellowBg, transparency: 6 }, line: { color: C.yellowBg, transparency: 100 } });
  line(s, x0 + w / 2, y0, x0 + w / 2, y0 + h, C.line, 1.1); line(s, x0, y0 + h / 2, x0 + w, y0 + h / 2, C.line, 1.1);
  label(s, '重复度高 →', x0 + w - 1.6, y0 + h + 0.2, 1.55, C.muted, 12, { align: 'right', bold: true });
  label(s, '可控性高 ↑', x0 - 1.25, y0 + 0.05, 1.0, C.muted, 12, { bold: true });
  label(s, '绿区：可以自动化', x0 + 4.9, y0 + 0.22, 2.5, C.teal, 15, { bold: true, align: 'center' });
  label(s, '黄区：AI 做副驾驶', x0 + 0.55, y0 + 2.95, 2.7, C.amber, 14, { bold: true, align: 'center' });
  label(s, '红区：人做最终决定', x0 + 5.0, y0 + 2.0, 2.9, C.coral, 14, { bold: true, align: 'center' });
  pill(s, '客服草稿', 7.0, 2.8, 1.25, C.teal, C.navy, 10); pill(s, '报表摘要', 8.65, 3.1, 1.25, C.teal, C.navy, 10); pill(s, '知识问答', 6.25, 3.25, 1.25, C.teal, C.navy, 10);
  pill(s, '采购预测', 2.65, 4.15, 1.25, C.amber, C.navy, 10); pill(s, '质检初筛', 4.45, 4.62, 1.25, C.amber, C.navy, 10); pill(s, '财务审批', 8.6, 4.75, 1.25, C.coral, C.navy, 10); pill(s, '对外承诺', 6.65, 5.15, 1.25, C.coral, C.navy, 10);
  label(s, '先让 AI 做副驾驶，再根据评测结果逐步放权。', 2.1, 6.3, 8.85, C.indigo, 16, { align: 'center', bold: true });
  addFooter(s, 15); addNotes(s, '坐标轴直接回答场景选择：重复度高、可控性高的任务优先自动化；高风险事项让 AI 做草稿、提醒或初筛，人做最终决定。');
}

// 16 — Rollout
{
  const s = pptx.addSlide(); addBg(s, C.pale); addTitle(s, '落地路线：先做副驾驶，再做流程，再做系统', '15 · 企业落地');
  label(s, '每一步都要有基线、有评测、有回滚，不要一开始就追求全自动。', 0.8, 1.55, 11.8, C.ink, 15.5, { bold: true });
  const stages = [{ x: 0.95, y: 4.7, w: 3.1, h: 0.9, n: '01', title: '副驾驶', sub: '写草稿、查资料、做摘要', c: C.ice }, { x: 4.35, y: 3.75, w: 3.1, h: 1.85, n: '02', title: '流程', sub: '自动分派、校验、提醒\n人保留关键节点', c: C.teal }, { x: 7.75, y: 2.75, w: 3.1, h: 2.85, n: '03', title: '系统', sub: '接入业务系统\n按权限行动\n持续评测与审计', c: C.indigo }];
  stages.forEach((it) => { s.addShape(pptx.ShapeType.roundRect, { x: it.x, y: it.y, w: it.w, h: it.h, rectRadius: 0.12, fill: { color: it.c }, line: { color: C.white, width: 1.1 } }); label(s, it.n, it.x + 0.25, it.y + 0.22, 0.55, C.navy, 15, { bold: true }); label(s, it.title, it.x + 0.92, it.y + 0.2, 1.65, C.white, 20, { bold: true }); label(s, it.sub, it.x + 0.3, it.y + 0.68, it.w - 0.6, C.white, 13, { h: it.h - 0.7 }); });
  line(s, 1.1, 5.98, 10.7, 5.98, C.line, 1.2, 'dash');
  ['基线 KPI', '可信数据', '离线评测', '小流量', '复盘'].forEach((t, i) => { const x = 1.12 + i * 2.35; dot(s, x, 6.27, 0.16, [C.coral, C.teal, C.indigo, C.amber, C.coral][i]); label(s, t, x + 0.25, 6.2, 1.0, C.muted, 11, { bold: true }); });
  addFooter(s, 16); addNotes(s, '成熟度是逐级放权。先建立稳定的评测基线和复核机制，再把通过验证的动作纳入流程，最后才考虑系统级行动。');
}

// 17 — Management checklist
{
  const s = pptx.addSlide(); addBg(s); addTitle(s, '把原理变成管理动作：六个问题决定成败', '16 · 管理清单');
  const qs = [['选场景', '哪里最重复？', C.ice], ['给数据', '它凭什么知道？', C.teal], ['设护栏', '哪些事不能做？', C.amber], ['算 ROI', '省下的钱够不够？', C.coral], ['留人工', '错了谁发现？', C.indigo], ['复盘', '数据如何变好？', C.mint]];
  qs.forEach((it, i) => { const x = 1.0 + (i % 3) * 4.05; const y = 2.0 + Math.floor(i / 3) * 1.65; card(s, x, y, 3.25, 1.15, { fill: C.white, line: it[2], lineWidth: 1.5, shadow: true }); dot(s, x + 0.28, y + 0.32, 0.42, it[2]); label(s, String(i + 1), x + 0.28, y + 0.4, 0.42, C.navy, 12, { align: 'center', bold: true }); label(s, it[0], x + 0.9, y + 0.2, 1.7, C.indigo, 16, { bold: true }); label(s, it[1], x + 0.9, y + 0.62, 1.9, C.muted, 12); });
  label(s, '把“理解原理”翻译成流程：每个 AI 项目都要能回答这六个问题。', 1.3, 6.2, 10.6, C.indigo, 16, { align: 'center', bold: true });
  addFooter(s, 17); addNotes(s, '用六个问题替代抽象的 AI 愿景：场景、数据、护栏、ROI、人工和复盘。它们对应模型输入、系统权限、评测和运营闭环。');
}

// 18 — Close
{
  const s = pptx.addSlide(); addBg(s, C.navy);
  s.addText('把 AI 放在\n可验证的工位', { x: 0.8, y: 1.15, w: 6.2, h: 1.35, fontFace: FONT.head, fontSize: 34, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  s.addText('懂原理 → 会评估 → 设边界 → 再逐步放权', { x: 0.83, y: 3.2, w: 6.0, h: 0.34, fontFace: FONT.body, fontSize: 18, color: C.mint, bold: true, margin: 0 });
  const qs = [['它凭什么知道？', '数据和上下文可信吗？', C.ice], ['错了谁发现？', '有测试、复核和回滚吗？', C.amber], ['省下的钱够吗？', '单位成本小于业务价值吗？', C.coral]];
  qs.forEach((it, i) => { const y = 1.3 + i * 1.45; card(s, 7.65, y, 4.55, 1.02, { fill: C.indigo, line: it[2], lineWidth: 1.4 }); dot(s, 7.95, y + 0.3, 0.4, it[2]); label(s, String(i + 1), 7.95, y + 0.38, 0.4, C.navy, 12, { align: 'center', bold: true }); label(s, it[0], 8.58, y + 0.18, 2.8, C.white, 16, { bold: true }); label(s, it[1], 8.58, y + 0.58, 3.1, C.darkMuted, 11.5); });
  s.addText('AI 是概率引擎和放大器，不是责任承担者。', { x: 0.83, y: 5.55, w: 6.0, h: 0.3, fontFace: FONT.body, fontSize: 17, color: C.white, bold: true, margin: 0 });
  addFooter(s, 18, true); addNotes(s, '收尾回到三个问题：它凭什么知道？错了谁发现？省下的钱够吗？理解原理的目的，是把 AI 放到可验证、可追责、可迭代的工位。');
}

async function addFadeTransitions(fileName) {
  const zip = await JSZip.loadAsync(fs.readFileSync(fileName));
  const slideNames = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  await Promise.all(slideNames.map(async name => {
    const xml = await zip.file(name).async('string');
    if (xml.includes('<p:transition')) return;
    zip.file(name, xml.replace('</p:cSld>', '</p:cSld><p:transition spd="med" advClick="1"><p:fade/></p:transition>'));
  }));
  fs.writeFileSync(fileName, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

const output = path.join(__dirname, 'ai_from_principles_to_practice.pptx');
pptx.writeFile({ fileName: output }).then(() => addFadeTransitions(output)).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

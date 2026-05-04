const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "何成";
pres.title = "水冷恒温箱的设计与实现";

// === Color Palette ===
const C = {
  navy: "1A2744",
  blue: "2B4C7E",
  accent: "3498DB",
  lightBlue: "E8F0FE",
  white: "FFFFFF",
  dark: "1A1A2E",
  gray: "5A6A7A",
  lightGray: "F0F2F5",
  green: "27AE60",
  orange: "E67E22",
  tableHead: "1A2744",
  tableAlt: "F5F7FA",
  border: "D0D7E2",
};

// === Helper Functions ===
const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

function addSlideHeader(slide, title, subtitle) {
  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  // Title
  slide.addText(title, {
    x: 0.6, y: 0.25, w: 8.8, h: 0.55, fontSize: 26, fontFace: "Arial Black",
    color: C.navy, bold: true, margin: 0,
  });
  // Separator line
  slide.addShape(pres.shapes.LINE, { x: 0.6, y: 0.85, w: 1.2, h: 0, line: { color: C.accent, width: 3 } });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6, y: 0.9, w: 8.8, h: 0.35, fontSize: 12, fontFace: "Calibri",
      color: C.gray, italic: true, margin: 0,
    });
  }
  // Page number area (bottom right)
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: C.navy } });
}

function addPageNum(slide, num, total) {
  slide.addText(`${num} / ${total}`, {
    x: 8.5, y: 5.37, w: 1.2, h: 0.22, fontSize: 9, fontFace: "Calibri",
    color: C.white, align: "right", margin: 0,
  });
}

// ============================================================
// SLIDE 1: Cover
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.dark };
  // Decorative top bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  // University name
  slide.addText("合肥工业大学", {
    x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 20, fontFace: "Calibri",
    color: C.accent, align: "center", charSpacing: 8, margin: 0,
  });
  slide.addText("仪器科学与光电工程学院", {
    x: 0.5, y: 1.0, w: 9, h: 0.45, fontSize: 14, fontFace: "Calibri",
    color: "8899AA", align: "center", charSpacing: 4, margin: 0,
  });
  // Title
  slide.addText("水冷恒温箱的设计与实现", {
    x: 0.8, y: 1.8, w: 8.4, h: 1.0, fontSize: 36, fontFace: "Arial Black",
    color: C.white, align: "center", bold: true, margin: 0,
  });
  // Subtitle line
  slide.addText("本科毕业设计答辩", {
    x: 0.8, y: 2.75, w: 8.4, h: 0.45, fontSize: 16, fontFace: "Calibri",
    color: C.accent, align: "center", margin: 0,
  });
  // Decorative line
  slide.addShape(pres.shapes.LINE, { x: 3.5, y: 3.35, w: 3, h: 0, line: { color: C.accent, width: 2 } });
  // Info block
  const infoOpts = { fontSize: 13, fontFace: "Calibri", color: "AABBCC", align: "center", margin: 0 };
  slide.addText("专　业：测控技术与仪器", { x: 1, y: 3.6, w: 8, h: 0.35, ...infoOpts });
  slide.addText("答辩人：何　成　　　　指导教师：李维诗", { x: 1, y: 3.95, w: 8, h: 0.35, ...infoOpts });
  slide.addText("2025年6月", { x: 1, y: 4.4, w: 8, h: 0.35, ...infoOpts, fontSize: 12, color: "778899" });
  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: C.accent } });
}

// ============================================================
// SLIDE 2: Research Background
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "研究背景与课题意义", "Background & Significance");
  addPageNum(slide, 2, 16);

  // Left column: 3 cards
  const cards = [
    { title: "精密测量需求", body: "光学检测、微纳加工等场景要求\n温度波动 < 0.1℃，振动与气流极低" },
    { title: "传统风冷方案不足", body: "风扇引入气流扰动、机械振动和噪声\n难以满足高精度恒温需求" },
    { title: "本课题方案", body: "提出无风扰外循环水冷恒温方案\n目标：高稳定、低扰动、良好均匀性" },
  ];
  cards.forEach((c, i) => {
    const yPos = 1.25 + i * 1.3;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: yPos, w: 5.0, h: 1.1,
      fill: { color: C.lightGray }, shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: yPos, w: 0.07, h: 1.1, fill: { color: C.accent },
    });
    slide.addText(c.title, {
      x: 0.9, y: yPos + 0.08, w: 4.5, h: 0.32,
      fontSize: 14, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
    });
    slide.addText(c.body, {
      x: 0.9, y: yPos + 0.42, w: 4.5, h: 0.6,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });

  // Right column: key point callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.0, y: 1.25, w: 3.6, h: 4.0,
    fill: { color: C.navy }, shadow: makeShadow(),
  });
  slide.addText("核心问题", {
    x: 6.2, y: 1.45, w: 3.2, h: 0.4,
    fontSize: 16, fontFace: "Arial Black", color: C.accent, bold: true, margin: 0,
  });
  slide.addText([
    { text: "如何在不引入强制气流的条件下，实现箱内温度的高精度、高稳定性控制？", options: { breakLine: true, fontSize: 13, color: C.white } },
    { text: "\n", options: { breakLine: true, fontSize: 8 } },
    { text: "解决思路", options: { breakLine: true, fontSize: 14, color: C.accent, bold: true } },
    { text: "TEC制冷/制热 + 外循环水冷传热 + 高导热内胆均温", options: { fontSize: 12, color: "AABBCC" } },
  ], { x: 6.2, y: 2.05, w: 3.2, h: 2.8, fontFace: "Calibri", valign: "top", margin: 0 });
}

// ============================================================
// SLIDE 3: Research Status
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "国内外研究现状与存在问题", "Literature Review");
  addPageNum(slide, 3, 16);

  // Comparison table
  const tableRows = [
    [
      { text: "方案类型", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "优点", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "不足", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
    ],
    [
      { text: "风冷恒温箱", options: { fill: { color: C.tableAlt } } },
      { text: "结构简单、成本低", options: { fill: { color: C.tableAlt } } },
      { text: "气流扰动大、振动噪声", options: { fill: { color: C.tableAlt } } },
    ],
    [
      { text: "液冷恒温箱", options: {} },
      { text: "换热效率高", options: {} },
      { text: "多为浸没式，不适用空气空间", options: {} },
    ],
    [
      { text: "TEC控温系统", options: { fill: { color: C.tableAlt } } },
      { text: "响应快、无运动部件", options: { fill: { color: C.tableAlt } } },
      { text: "制冷量有限、需散热", options: { fill: { color: C.tableAlt } } },
    ],
  ];
  slide.addTable(tableRows, {
    x: 0.6, y: 1.3, w: 8.8, h: 2.0,
    fontSize: 12, fontFace: "Calibri", color: C.navy,
    border: { type: "solid", pt: 0.5, color: C.border },
    colW: [2.2, 3.3, 3.3],
    valign: "middle",
  });

  // Gap analysis box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.6, w: 8.8, h: 1.5,
    fill: { color: C.lightBlue }, shadow: makeShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.6, w: 0.07, h: 1.5, fill: { color: C.accent },
  });
  slide.addText("研究空白", {
    x: 0.9, y: 3.7, w: 4, h: 0.35,
    fontSize: 14, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  slide.addText([
    { text: "• 面向空气空间的无风扰水冷恒温箱研究较少", options: { breakLine: true } },
    { text: "• 现有方案多依赖风扇强制对流，难以兼顾低扰动与高均匀性", options: { breakLine: true } },
    { text: "• 缺乏针对大惯性水冷系统的控制策略研究", options: {} },
  ], {
    x: 0.9, y: 4.1, w: 8.2, h: 0.95,
    fontSize: 12, fontFace: "Calibri", color: C.gray, margin: 0,
  });
}

// ============================================================
// SLIDE 4: Research Goals & Specs
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "研究目标与技术指标", "Research Objectives & Specifications");
  addPageNum(slide, 4, 16);

  // Goal statement
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.25, w: 8.8, h: 0.65,
    fill: { color: C.navy },
  });
  slide.addText("设计一台面向精密测量场景的外循环水冷无风扰恒温箱，实现高稳定、低扰动的恒温环境", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.55,
    fontSize: 13, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", margin: 0,
  });

  // Specs table
  const specRows = [
    [
      { text: "技术指标", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "设计目标", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
    ],
    [
      { text: "目标温度", options: { fill: { color: C.tableAlt } } },
      { text: "20℃ 附近", options: { fill: { color: C.tableAlt } } },
    ],
    [
      { text: "调节时间", options: {} },
      { text: "≤ 120 min", options: {} },
    ],
    [
      { text: "最大超调量", options: { fill: { color: C.tableAlt } } },
      { text: "≤ 0.3℃", options: { fill: { color: C.tableAlt } } },
    ],
    [
      { text: "15h 温度波动", options: {} },
      { text: "< 0.1℃", options: {} },
    ],
    [
      { text: "空间均匀度", options: { fill: { color: C.tableAlt } } },
      { text: "< 0.1℃", options: { fill: { color: C.tableAlt } } },
    ],
    [
      { text: "扰动恢复时间", options: {} },
      { text: "≤ 30 min", options: {} },
    ],
  ];
  slide.addTable(specRows, {
    x: 1.5, y: 2.15, w: 7.0, h: 3.0,
    fontSize: 13, fontFace: "Calibri", color: C.navy,
    border: { type: "solid", pt: 0.5, color: C.border },
    colW: [3.5, 3.5],
    valign: "middle", align: "center",
  });
}

// ============================================================
// SLIDE 5: System Overview
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "系统总体方案", "System Architecture");
  addPageNum(slide, 5, 16);

  // Three pillars
  const pillars = [
    { title: "机械结构", items: "三层复合箱体\n保温层+水冷层+内胆\n无风扇设计" },
    { title: "硬件电路", items: "STM32主控\nPT100多点测温\nTEC驱动+水泵驱动" },
    { title: "软件控制", items: "积分分离PID\n中值+滑动平均滤波\n串口通信+OLED显示" },
  ];
  pillars.forEach((p, i) => {
    const xPos = 0.6 + i * 3.1;
    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 2.8, w: 2.8, h: 2.3,
      fill: { color: C.lightGray }, shadow: makeShadow(),
    });
    // Top accent
    slide.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 2.8, w: 2.8, h: 0.06, fill: { color: C.accent },
    });
    // Title
    slide.addText(p.title, {
      x: xPos + 0.15, y: 2.95, w: 2.5, h: 0.4,
      fontSize: 15, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
    });
    // Items
    slide.addText(p.items, {
      x: xPos + 0.15, y: 3.4, w: 2.5, h: 1.5,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0, lineSpacingMultiple: 1.5,
    });
  });

  // Top: system block diagram (text-based)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.25, w: 8.8, h: 1.3,
    fill: { color: C.lightBlue },
  });
  // Flow: TEC -> Water Loop -> Box Wall -> Inner Chamber
  const blocks = [
    { label: "TEC\n制冷/制热", x: 0.9 },
    { label: "外循环\n水路系统", x: 3.0 },
    { label: "箱壁\n换热层", x: 5.1 },
    { label: "高导热\n内胆", x: 7.2 },
  ];
  blocks.forEach((b) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: b.x, y: 1.45, w: 1.7, h: 0.85,
      fill: { color: C.navy }, shadow: makeShadow(),
    });
    slide.addText(b.label, {
      x: b.x, y: 1.45, w: 1.7, h: 0.85,
      fontSize: 12, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", margin: 0,
    });
  });
  // Arrows between blocks
  [2.65, 4.75, 6.85].forEach((ax) => {
    slide.addText("→", {
      x: ax, y: 1.55, w: 0.4, h: 0.65,
      fontSize: 22, fontFace: "Calibri", color: C.accent, align: "center", valign: "middle", margin: 0,
    });
  });
}

// ============================================================
// SLIDE 6: Mechanical Structure & Heat Transfer
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "机械结构与传热设计", "Mechanical Design & Heat Transfer");
  addPageNum(slide, 6, 16);

  // Left: cross-section diagram (text-based layered structure)
  slide.addText("箱体剖面结构示意", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  // Outer layer - insulation
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 4.0, h: 3.2,
    fill: { color: "D5E8D4" }, line: { color: "82B366", width: 1.5 },
  });
  slide.addText("保温层（聚氨酯发泡）", {
    x: 0.9, y: 1.8, w: 3.8, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: "4A7A3A", margin: 0,
  });
  // Middle layer - water cooling
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 1.3, y: 2.25, w: 3.1, h: 2.3,
    fill: { color: "DAE8FC" }, line: { color: "6C8EBF", width: 1.5 },
  });
  slide.addText("水冷换热层（蛇形紫铜管）", {
    x: 1.4, y: 2.3, w: 2.9, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: "3A6A9A", margin: 0,
  });
  // Inner layer - high conductivity liner
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 1.8, y: 2.75, w: 2.1, h: 1.4,
    fill: { color: "FFF2CC" }, line: { color: "D6B656", width: 1.5 },
  });
  slide.addText("高导热铝内胆\n300×250×200mm\n有效容积 ~15L", {
    x: 1.85, y: 2.85, w: 2.0, h: 1.1,
    fontSize: 10, fontFace: "Calibri", color: "8A6D2B", align: "center", valign: "middle", margin: 0,
  });

  // Right: key design points
  slide.addText("设计要点", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const points = [
    { title: "三层复合结构", desc: "保温层隔热 → 水冷层换热 → 内胆均温" },
    { title: "无风扇设计", desc: "箱内无强制对流，消除气流扰动" },
    { title: "热量传递路径", desc: "TEC → 循环水 → 蛇形铜管 → 箱壁 → 内胆辐射/自然对流" },
    { title: "密封与保温", desc: "聚氨酯发泡填充，硅胶密封条，减少热泄漏" },
  ];
  points.forEach((p, i) => {
    const yPos = 1.75 + i * 0.82;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: yPos, w: 4.1, h: 0.7,
      fill: { color: i % 2 === 0 ? C.lightGray : C.lightBlue },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: yPos, w: 0.06, h: 0.7, fill: { color: C.accent },
    });
    slide.addText(p.title, {
      x: 5.75, y: yPos + 0.05, w: 3.7, h: 0.28,
      fontSize: 12, fontFace: "Calibri", color: C.navy, bold: true, margin: 0,
    });
    slide.addText(p.desc, {
      x: 5.75, y: yPos + 0.35, w: 3.7, h: 0.28,
      fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });
}

// ============================================================
// SLIDE 7: Water Circuit & Temperature Sensors
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "水路布局与测温点设计", "Water Circuit & Sensor Layout");
  addPageNum(slide, 7, 16);

  // Left: water circuit description
  slide.addText("蛇形水路布局", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  // Schematic box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.7, w: 4.5, h: 2.6,
    fill: { color: C.lightBlue },
  });
  // Draw serpentine pattern with text
  slide.addText([
    { text: "┌──────────────────┐", options: { breakLine: true } },
    { text: "│  ═══╗  ╔═══╗  ╔══│ ← 底板", options: { breakLine: true } },
    { text: "│  ═══╝  ╚═══╝  ╚══│", options: { breakLine: true } },
    { text: "├──────────────────┤", options: { breakLine: true } },
    { text: "│ ║    侧壁    ║  │", options: { breakLine: true } },
    { text: "│ ║  蛇形铜管  ║  │", options: { breakLine: true } },
    { text: "│ ║            ║  │", options: { breakLine: true } },
    { text: "└──────────────────┘", options: {} },
  ], {
    x: 0.8, y: 1.8, w: 4.1, h: 2.3,
    fontSize: 10, fontFace: "Consolas", color: C.blue, margin: 0, lineSpacingMultiple: 1.2,
  });
  // Water circuit features
  slide.addText([
    { text: "• 单路串联蛇形紫铜管", options: { breakLine: true } },
    { text: "• 覆盖底板、侧壁、后壁", options: { breakLine: true } },
    { text: "• 紫铜管导热系数高，换热均匀", options: {} },
  ], {
    x: 0.6, y: 4.45, w: 4.5, h: 0.8,
    fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
  });

  // Right: 5-point temperature monitoring
  slide.addText("五点 PT100 测温布置", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  // Box top view
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: 1.75, w: 3.5, h: 2.5,
    fill: { color: "FFF9E6" }, line: { color: C.border, width: 1 },
  });
  slide.addText("箱内俯视图", {
    x: 5.9, y: 1.8, w: 1.5, h: 0.25,
    fontSize: 9, fontFace: "Calibri", color: C.gray, italic: true, margin: 0,
  });
  // 5 sensor points
  const sensors = [
    { label: "T1", x: 6.2, y: 2.15 },
    { label: "T2", x: 8.6, y: 2.15 },
    { label: "T3", x: 7.4, y: 2.85 },
    { label: "T4", x: 6.2, y: 3.55 },
    { label: "T5", x: 8.6, y: 3.55 },
  ];
  sensors.forEach((s) => {
    slide.addShape(pres.shapes.OVAL, {
      x: s.x, y: s.y, w: 0.35, h: 0.35,
      fill: { color: C.accent },
    });
    slide.addText(s.label, {
      x: s.x, y: s.y, w: 0.35, h: 0.35,
      fontSize: 9, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
    });
  });

  // Sensor notes
  slide.addText([
    { text: "• 五点分布于箱内不同空间位置", options: { breakLine: true } },
    { text: "• 监测空间温度均匀性", options: { breakLine: true } },
    { text: "• T3 为中心点，作为控制反馈信号", options: {} },
  ], {
    x: 5.5, y: 4.45, w: 4.2, h: 0.8,
    fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
  });
}

// ============================================================
// SLIDE 8: Hardware System Design
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "硬件系统设计", "Hardware Design");
  addPageNum(slide, 8, 16);

  // Central MCU block
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 3.8, y: 2.4, w: 2.4, h: 1.2,
    fill: { color: C.navy }, shadow: makeShadow(),
  });
  slide.addText("STM32F103C8T6\n主控芯片", {
    x: 3.8, y: 2.4, w: 2.4, h: 1.2,
    fontSize: 13, fontFace: "Arial Black", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
  });

  // Surrounding modules
  const modules = [
    { label: "PT100 + ADS1220\n温度采集模块", x: 0.5, y: 1.3, w: 2.5, h: 0.9, color: "E8F5E9", accent: C.green },
    { label: "TEC 双向\nH桥驱动", x: 0.5, y: 2.6, w: 2.5, h: 0.9, color: "FFF3E0", accent: C.orange },
    { label: "水泵驱动\nMOS管控制", x: 0.5, y: 3.9, w: 2.5, h: 0.9, color: "E3F2FD", accent: C.accent },
    { label: "OLED 显示\n实时温度/状态", x: 7.0, y: 1.3, w: 2.5, h: 0.9, color: "F3E5F5", accent: "9C27B0" },
    { label: "串口通信\n上位机数据上传", x: 7.0, y: 2.6, w: 2.5, h: 0.9, color: "E8EAF6", accent: "3F51B5" },
    { label: "电源管理\n12V/5V/3.3V", x: 7.0, y: 3.9, w: 2.5, h: 0.9, color: "FFEBEE", accent: "E53935" },
  ];
  modules.forEach((m) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: m.x, y: m.y, w: m.w, h: m.h,
      fill: { color: m.color }, shadow: makeShadow(),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: m.x, y: m.y, w: 0.06, h: m.h, fill: { color: m.accent },
    });
    slide.addText(m.label, {
      x: m.x + 0.15, y: m.y, w: m.w - 0.2, h: m.h,
      fontSize: 11, fontFace: "Calibri", color: C.navy, valign: "middle", margin: 0,
    });
  });

  // Arrows from left modules to MCU
  [1.75, 3.05, 4.35].forEach((ay) => {
    slide.addText("→", {
      x: 3.1, y: ay, w: 0.6, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.accent, align: "center", valign: "middle", margin: 0,
    });
  });
  // Arrows from MCU to right modules
  [1.75, 3.05, 4.35].forEach((ay) => {
    slide.addText("→", {
      x: 6.3, y: ay, w: 0.6, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.accent, align: "center", valign: "middle", margin: 0,
    });
  });

  // Bottom note
  slide.addText("系统具备多点温度采集、闭环控制、数据显示与上传功能", {
    x: 0.6, y: 5.0, w: 8.8, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center", italic: true, margin: 0,
  });
}

// ============================================================
// SLIDE 9: Software & Control Algorithm
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "软件流程与控制算法", "Software & Control Algorithm");
  addPageNum(slide, 9, 16);

  // Left: software flow
  slide.addText("软件主流程", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const flowSteps = [
    "系统初始化",
    "硬件自检",
    "多通道温度采集",
    "中值滤波 + 滑动平均",
    "积分分离PID运算",
    "TEC/水泵驱动输出",
    "OLED显示 + 串口上传",
  ];
  flowSteps.forEach((step, i) => {
    const yPos = 1.75 + i * 0.48;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.0, y: yPos, w: 3.6, h: 0.38,
      fill: { color: i === 4 ? C.accent : C.navy },
    });
    slide.addText(step, {
      x: 1.0, y: yPos, w: 3.6, h: 0.38,
      fontSize: 11, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", margin: 0,
    });
    if (i < flowSteps.length - 1) {
      slide.addText("▼", {
        x: 2.5, y: yPos + 0.35, w: 0.6, h: 0.18,
        fontSize: 8, fontFace: "Calibri", color: C.accent, align: "center", margin: 0,
      });
    }
  });

  // Right: control algorithm details
  slide.addText("积分分离PID控制", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  // Algorithm card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 1.7, w: 4.1, h: 1.8,
    fill: { color: C.lightBlue },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 1.7, w: 0.06, h: 1.8, fill: { color: C.accent },
  });
  slide.addText([
    { text: "核心思想：", options: { bold: true, breakLine: true } },
    { text: "|e(t)| > ε 时，去掉积分项，避免大超调", options: { breakLine: true } },
    { text: "|e(t)| ≤ ε 时，引入积分项，消除稳态误差", options: { breakLine: true } },
    { text: "\n", options: { breakLine: true, fontSize: 6 } },
    { text: "适用特点：", options: { bold: true, breakLine: true } },
    { text: "水冷系统大惯性、纯滞后特性", options: {} },
  ], {
    x: 5.75, y: 1.8, w: 3.7, h: 1.6,
    fontSize: 11, fontFace: "Calibri", color: C.navy, margin: 0,
  });

  // PID parameters
  slide.addText("PID参数整定结果", {
    x: 5.5, y: 3.7, w: 4.1, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.navy, bold: true, margin: 0,
  });
  const pidRows = [
    [
      { text: "参数", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "数值", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
    ],
    [{ text: "Kp", options: { fill: { color: C.tableAlt } } }, { text: "85", options: { fill: { color: C.tableAlt } } }],
    [{ text: "Ki", options: {} }, { text: "0.035", options: {} }],
    [{ text: "Kd", options: { fill: { color: C.tableAlt } } }, { text: "200", options: { fill: { color: C.tableAlt } } }],
    [{ text: "积分分离阈值 ε", options: {} }, { text: "0.8℃", options: {} }],
  ];
  slide.addTable(pidRows, {
    x: 5.5, y: 4.05, w: 4.1, h: 1.15,
    fontSize: 11, fontFace: "Calibri", color: C.navy,
    border: { type: "solid", pt: 0.5, color: C.border },
    colW: [2.2, 1.9], valign: "middle", align: "center",
  });
}

// ============================================================
// SLIDE 10: Simulation Analysis
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "系统建模与仿真分析", "Modeling & Simulation (MATLAB/Simulink)");
  addPageNum(slide, 10, 16);

  // Left: model description
  slide.addText("一阶惯性纯滞后模型", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.7, w: 4.5, h: 1.2,
    fill: { color: C.lightGray },
  });
  slide.addText([
    { text: "G(s) = K · e^(-τs) / (Ts + 1)", options: { breakLine: true, fontSize: 16, bold: true, fontFace: "Consolas", color: C.navy } },
    { text: "\n", options: { breakLine: true, fontSize: 6 } },
    { text: "水冷恒温箱具有大惯性、纯滞后特性\n传统PID易产生较大超调和振荡", options: { fontSize: 11, color: C.gray } },
  ], {
    x: 0.8, y: 1.8, w: 4.1, h: 1.0,
    fontFace: "Calibri", margin: 0, valign: "middle",
  });

  // Comparison table
  slide.addText("常规PID vs 积分分离PID 仿真对比", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const simRows = [
    [
      { text: "性能指标", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "常规PID", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "积分分离PID", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "改善效果", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
    ],
    [
      { text: "最大超调", options: { fill: { color: C.tableAlt } } },
      { text: "较大", options: { fill: { color: C.tableAlt } } },
      { text: "~0.10℃", options: { fill: { color: C.tableAlt }, bold: true, color: C.green } },
      { text: "显著减小", options: { fill: { color: C.tableAlt }, color: C.green } },
    ],
    [
      { text: "调节时间", options: {} },
      { text: "较长", options: {} },
      { text: "~86.7 min", options: { bold: true, color: C.accent } },
      { text: "明显缩短", options: { color: C.green } },
    ],
    [
      { text: "15h波动", options: { fill: { color: C.tableAlt } } },
      { text: "—", options: { fill: { color: C.tableAlt } } },
      { text: "~0.077℃", options: { fill: { color: C.tableAlt }, bold: true, color: C.accent } },
      { text: "满足要求", options: { fill: { color: C.tableAlt }, color: C.green } },
    ],
    [
      { text: "扰动恢复", options: {} },
      { text: "较慢", options: {} },
      { text: "~18.3 min", options: { bold: true, color: C.accent } },
      { text: "恢复更快", options: { color: C.green } },
    ],
  ];
  slide.addTable(simRows, {
    x: 0.6, y: 3.5, w: 8.8, h: 1.7,
    fontSize: 12, fontFace: "Calibri", color: C.navy,
    border: { type: "solid", pt: 0.5, color: C.border },
    colW: [2.2, 2.0, 2.3, 2.3], valign: "middle", align: "center",
  });

  // Right top: conclusion callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 1.25, w: 4.1, h: 1.65,
    fill: { color: C.navy }, shadow: makeShadow(),
  });
  slide.addText("仿真结论", {
    x: 5.7, y: 1.35, w: 3.7, h: 0.3,
    fontSize: 14, fontFace: "Arial Black", color: C.accent, bold: true, margin: 0,
  });
  slide.addText([
    { text: "积分分离PID有效抑制超调", options: { breakLine: true } },
    { text: "缩短调节时间", options: { breakLine: true } },
    { text: "提高稳态精度", options: { breakLine: true } },
    { text: "验证了控制策略的可行性", options: {} },
  ], {
    x: 5.7, y: 1.75, w: 3.7, h: 1.0,
    fontSize: 12, fontFace: "Calibri", color: C.white, margin: 0, lineSpacingMultiple: 1.4,
  });
}

// ============================================================
// SLIDE 11: Prototype & Test Platform
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "样机制作与实验平台", "Prototype & Test Platform");
  addPageNum(slide, 11, 16);

  // Left: prototype components
  slide.addText("样机组成", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const components = [
    { name: "三层复合箱体", desc: "保温层 + 水冷层 + 铝内胆" },
    { name: "水路系统", desc: "蛇形紫铜管 + 水泵 + 水箱" },
    { name: "TEC制冷模组", desc: "TEC片 + 散热器 + 散热风扇" },
    { name: "控制电路板", desc: "STM32 + ADS1220 + 驱动电路" },
    { name: "上位机测试平台", desc: "串口数据采集 + 实时曲线显示" },
  ];
  components.forEach((c, i) => {
    const yPos = 1.7 + i * 0.65;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: yPos, w: 4.5, h: 0.55,
      fill: { color: i % 2 === 0 ? C.lightGray : C.lightBlue },
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.75, y: yPos + 0.1, w: 0.35, h: 0.35,
      fill: { color: C.accent },
    });
    slide.addText(String(i + 1), {
      x: 0.75, y: yPos + 0.1, w: 0.35, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
    });
    slide.addText(c.name, {
      x: 1.25, y: yPos + 0.02, w: 1.8, h: 0.25,
      fontSize: 12, fontFace: "Calibri", color: C.navy, bold: true, margin: 0,
    });
    slide.addText(c.desc, {
      x: 1.25, y: yPos + 0.28, w: 3.6, h: 0.22,
      fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });

  // Right: experiment list
  slide.addText("实验项目", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const experiments = [
    "温度采集精度验证",
    "阶跃响应实验",
    "长时间稳定性实验（15h）",
    "空间均匀性实验（五点）",
    "扰动恢复实验（开门15s）",
  ];
  experiments.forEach((exp, i) => {
    const yPos = 1.75 + i * 0.6;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: yPos, w: 4.1, h: 0.48,
      fill: { color: C.navy }, shadow: makeShadow(),
    });
    slide.addText(`${i + 1}   ${exp}`, {
      x: 5.7, y: yPos, w: 3.7, h: 0.48,
      fontSize: 12, fontFace: "Calibri", color: C.white, valign: "middle", margin: 0,
    });
  });

  // Bottom note
  slide.addText("完成样机装配、水路搭建、PCB焊接、嵌入式程序开发与调试", {
    x: 0.6, y: 5.0, w: 8.8, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: C.gray, align: "center", italic: true, margin: 0,
  });
}

// ============================================================
// SLIDE 12: Experiment Results 1 - Accuracy & Step Response
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "实验结果（一）测温精度与阶跃响应", "Results: Accuracy & Step Response");
  addPageNum(slide, 12, 16);

  // Left: temperature accuracy
  slide.addText("测温精度验证", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  // Big number callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.7, w: 4.5, h: 1.3,
    fill: { color: C.lightBlue }, shadow: makeShadow(),
  });
  slide.addText("±0.02℃", {
    x: 0.6, y: 1.75, w: 4.5, h: 0.7,
    fontSize: 40, fontFace: "Arial Black", color: C.accent, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("各通道在20℃附近的测量误差", {
    x: 0.6, y: 2.5, w: 4.5, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center", margin: 0,
  });
  slide.addText([
    { text: "• 采用PT100四线制 + ADS1220高精度ADC", options: { breakLine: true } },
    { text: "• 五通道一致性良好", options: { breakLine: true } },
    { text: "• 满足高精度温控反馈需求", options: {} },
  ], {
    x: 0.8, y: 3.15, w: 4.1, h: 0.8,
    fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
  });

  // Right: step response
  slide.addText("阶跃响应实验", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  // Step response key data
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.5, y: 1.7, w: 4.1, h: 2.5,
    fill: { color: C.lightGray }, shadow: makeShadow(),
  });
  // Two big numbers side by side
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.7, y: 1.85, w: 1.8, h: 1.0,
    fill: { color: C.navy },
  });
  slide.addText("95 min", {
    x: 5.7, y: 1.85, w: 1.8, h: 0.65,
    fontSize: 28, fontFace: "Arial Black", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("调节时间", {
    x: 5.7, y: 2.5, w: 1.8, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: "AABBCC", align: "center", margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.6, y: 1.85, w: 1.8, h: 1.0,
    fill: { color: C.navy },
  });
  slide.addText("0.15℃", {
    x: 7.6, y: 1.85, w: 1.8, h: 0.65,
    fontSize: 28, fontFace: "Arial Black", color: C.accent, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("最大超调", {
    x: 7.6, y: 2.5, w: 1.8, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: "AABBCC", align: "center", margin: 0,
  });

  // Description
  slide.addText([
    { text: "实验条件：", options: { bold: true, breakLine: true } },
    { text: "室温 ~24.8℃ → 目标 20℃", options: { breakLine: true } },
    { text: "\n", options: { breakLine: true, fontSize: 4 } },
    { text: "系统从室温降至20℃目标温度，约95min达到稳态，最大超调约0.15℃，响应过程平稳", options: {} },
  ], {
    x: 5.7, y: 3.05, w: 3.7, h: 1.0,
    fontSize: 11, fontFace: "Calibri", color: C.navy, margin: 0,
  });

  // Bottom highlight bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.3, w: 8.8, h: 0.55,
    fill: { color: "E8F5E9" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.3, w: 0.06, h: 0.55, fill: { color: C.green },
  });
  slide.addText("结论：测温精度和阶跃响应均优于设计目标（调节时间 ≤120min，超调 ≤0.3℃）", {
    x: 0.85, y: 4.3, w: 8.3, h: 0.55,
    fontSize: 12, fontFace: "Calibri", color: C.navy, valign: "middle", bold: true, margin: 0,
  });
}

// ============================================================
// SLIDE 13: Experiment Results 2 - Stability & Uniformity
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "实验结果（二）稳定性与空间均匀性", "Results: Stability & Uniformity");
  addPageNum(slide, 13, 16);

  // Top: two big metric cards
  // Card 1: 15h stability
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.25, w: 4.3, h: 2.0,
    fill: { color: C.navy }, shadow: makeShadow(),
  });
  slide.addText("15h 长时间稳定性", {
    x: 0.8, y: 1.35, w: 3.9, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: "AABBCC", margin: 0,
  });
  slide.addText("0.094℃", {
    x: 0.6, y: 1.75, w: 4.3, h: 0.7,
    fontSize: 42, fontFace: "Arial Black", color: C.accent, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("峰谷波动（目标 < 0.1℃）", {
    x: 0.6, y: 2.5, w: 4.3, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: "8899AA", align: "center", margin: 0,
  });
  // Check mark
  slide.addText("✓ 达标", {
    x: 3.3, y: 2.85, w: 1.2, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.green, bold: true, align: "right", margin: 0,
  });

  // Card 2: spatial uniformity
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.25, w: 4.3, h: 2.0,
    fill: { color: C.navy }, shadow: makeShadow(),
  });
  slide.addText("五点空间均匀性", {
    x: 5.3, y: 1.35, w: 3.9, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: "AABBCC", margin: 0,
  });
  slide.addText("0.076℃", {
    x: 5.1, y: 1.75, w: 4.3, h: 0.7,
    fontSize: 42, fontFace: "Arial Black", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("最大温差（目标 < 0.1℃）", {
    x: 5.1, y: 2.5, w: 4.3, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: "8899AA", align: "center", margin: 0,
  });
  slide.addText("✓ 达标", {
    x: 7.8, y: 2.85, w: 1.2, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.green, bold: true, align: "right", margin: 0,
  });

  // Bottom: analysis cards
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.5, w: 4.3, h: 1.5,
    fill: { color: C.lightBlue },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.5, w: 0.06, h: 1.5, fill: { color: C.accent },
  });
  slide.addText("稳定性分析", {
    x: 0.85, y: 3.55, w: 3.8, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.navy, bold: true, margin: 0,
  });
  slide.addText([
    { text: "• 连续运行15小时无明显漂移", options: { breakLine: true } },
    { text: "• 温度波动控制在0.094℃以内", options: { breakLine: true } },
    { text: "• 积分分离PID有效消除稳态误差", options: {} },
  ], {
    x: 0.85, y: 3.9, w: 3.8, h: 0.9,
    fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 3.5, w: 4.3, h: 1.5,
    fill: { color: C.lightGray },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 3.5, w: 0.06, h: 1.5, fill: { color: C.accent },
  });
  slide.addText("均匀性分析", {
    x: 5.35, y: 3.55, w: 3.8, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.navy, bold: true, margin: 0,
  });
  slide.addText([
    { text: "• 五点温度分布一致性良好", options: { breakLine: true } },
    { text: "• 高导热内胆有效扩散壁面热量", options: { breakLine: true } },
    { text: "• 无风扇条件下均匀性优于0.1℃", options: {} },
  ], {
    x: 5.35, y: 3.9, w: 3.8, h: 0.9,
    fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
  });
}

// ============================================================
// SLIDE 14: Experiment Results 3 - Disturbance & Comparison
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "实验结果（三）扰动恢复与指标对比", "Results: Disturbance Recovery & Summary");
  addPageNum(slide, 14, 16);

  // Left: disturbance recovery
  slide.addText("扰动恢复实验", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.7, w: 4.5, h: 1.6,
    fill: { color: C.lightGray }, shadow: makeShadow(),
  });
  // Big number
  slide.addText("22 min", {
    x: 0.6, y: 1.8, w: 4.5, h: 0.7,
    fontSize: 38, fontFace: "Arial Black", color: C.accent, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("恢复时间（目标 ≤ 30min）", {
    x: 0.6, y: 2.5, w: 4.5, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.gray, align: "center", margin: 0,
  });
  slide.addText([
    { text: "实验条件：", options: { bold: true, breakLine: true } },
    { text: "开门15s引入环境扰动后关门", options: { breakLine: true } },
    { text: "系统自动恢复至稳态，约22min", options: {} },
  ], {
    x: 0.8, y: 2.85, w: 4.1, h: 0.4,
    fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0,
  });

  // Right: comprehensive comparison table
  slide.addText("设计目标 vs 实验结果", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const compRows = [
    [
      { text: "指标", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "目标", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "实测", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
      { text: "", options: { bold: true, color: "FFFFFF", fill: { color: C.tableHead } } },
    ],
    [
      { text: "调节时间", options: { fill: { color: C.tableAlt } } },
      { text: "≤120min", options: { fill: { color: C.tableAlt } } },
      { text: "95min", options: { fill: { color: C.tableAlt }, bold: true, color: C.accent } },
      { text: "✓", options: { fill: { color: C.tableAlt }, color: C.green, bold: true } },
    ],
    [
      { text: "最大超调", options: {} },
      { text: "≤0.3℃", options: {} },
      { text: "0.15℃", options: { bold: true, color: C.accent } },
      { text: "✓", options: { color: C.green, bold: true } },
    ],
    [
      { text: "15h波动", options: { fill: { color: C.tableAlt } } },
      { text: "<0.1℃", options: { fill: { color: C.tableAlt } } },
      { text: "0.094℃", options: { fill: { color: C.tableAlt }, bold: true, color: C.accent } },
      { text: "✓", options: { fill: { color: C.tableAlt }, color: C.green, bold: true } },
    ],
    [
      { text: "空间均匀度", options: {} },
      { text: "<0.1℃", options: {} },
      { text: "0.076℃", options: { bold: true, color: C.accent } },
      { text: "✓", options: { color: C.green, bold: true } },
    ],
    [
      { text: "扰动恢复", options: { fill: { color: C.tableAlt } } },
      { text: "≤30min", options: { fill: { color: C.tableAlt } } },
      { text: "22min", options: { fill: { color: C.tableAlt }, bold: true, color: C.accent } },
      { text: "✓", options: { fill: { color: C.tableAlt }, color: C.green, bold: true } },
    ],
  ];
  slide.addTable(compRows, {
    x: 5.3, y: 1.7, w: 4.4, h: 2.4,
    fontSize: 12, fontFace: "Calibri", color: C.navy,
    border: { type: "solid", pt: 0.5, color: C.border },
    colW: [1.2, 1.1, 1.1, 0.5], valign: "middle", align: "center",
  });

  // Bottom conclusion bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.5, w: 8.8, h: 0.6,
    fill: { color: C.navy },
  });
  slide.addText("所有设计指标均已达到，系统可在20℃附近稳定运行，满足精密测量恒温需求", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.6,
    fontSize: 14, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
  });
}

// ============================================================
// SLIDE 15: Summary, Innovation, Outlook
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  addSlideHeader(slide, "总结、创新点与展望", "Conclusion & Future Work");
  addPageNum(slide, 15, 16);

  // Left column: summary + innovation
  slide.addText("研究成果", {
    x: 0.6, y: 1.25, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.65, w: 4.5, h: 1.2,
    fill: { color: C.lightBlue },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.65, w: 0.06, h: 1.2, fill: { color: C.accent },
  });
  slide.addText([
    { text: "• 完成了水冷恒温箱的方案设计与样机制作", options: { breakLine: true } },
    { text: "• 实现了无风扰条件下的高精度温度控制", options: { breakLine: true } },
    { text: "• 各项指标均达到设计目标", options: { breakLine: true } },
    { text: "• 验证了外循环水冷恒温方案的可行性", options: {} },
  ], {
    x: 0.85, y: 1.72, w: 4.0, h: 1.05,
    fontSize: 11, fontFace: "Calibri", color: C.navy, margin: 0, lineSpacingMultiple: 1.3,
  });

  // Innovation points
  slide.addText("创新与特色", {
    x: 0.6, y: 3.05, w: 4.5, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const innovations = [
    "提出适用于精密测量的外循环水冷无风扰恒温方案",
    "采用三层复合箱体结构提升壁面传热均匀性",
    "针对大惯性长滞后对象采用积分分离PID策略",
    "通过样机实验验证了方案可行性",
  ];
  innovations.forEach((inn, i) => {
    const yPos = 3.45 + i * 0.42;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: yPos + 0.05, w: 0.25, h: 0.25,
      fill: { color: C.accent },
    });
    slide.addText(String(i + 1), {
      x: 0.7, y: yPos + 0.05, w: 0.25, h: 0.25,
      fontSize: 9, fontFace: "Calibri", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
    });
    slide.addText(inn, {
      x: 1.1, y: yPos, w: 3.9, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.navy, valign: "middle", margin: 0,
    });
  });

  // Right column: limitations & outlook
  slide.addText("不足与展望", {
    x: 5.5, y: 1.25, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Arial Black", color: C.navy, bold: true, margin: 0,
  });
  const limitations = [
    { title: "装配精度", desc: "盘管贴合与装配工艺仍可优化" },
    { title: "模型精度", desc: "工程简化模型可进一步精细化辨识" },
    { title: "控制算法", desc: "可研究Smith预估、自适应控制、模型预测控制等" },
    { title: "长期验证", desc: "更高精度工况下的长期可靠性有待验证" },
  ];
  limitations.forEach((lim, i) => {
    const yPos = 1.7 + i * 0.85;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: yPos, w: 4.1, h: 0.72,
      fill: { color: i % 2 === 0 ? C.lightGray : C.lightBlue },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: yPos, w: 0.06, h: 0.72, fill: { color: C.orange },
    });
    slide.addText(lim.title, {
      x: 5.75, y: yPos + 0.06, w: 3.7, h: 0.28,
      fontSize: 12, fontFace: "Calibri", color: C.navy, bold: true, margin: 0,
    });
    slide.addText(lim.desc, {
      x: 5.75, y: yPos + 0.36, w: 3.7, h: 0.28,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });
}

// ============================================================
// SLIDE 16: Thank You / Q&A
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.dark };
  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

  // Decorative line
  slide.addShape(pres.shapes.LINE, { x: 3.5, y: 1.6, w: 3, h: 0, line: { color: C.accent, width: 2 } });

  // Main text
  slide.addText("感谢各位老师聆听", {
    x: 0.5, y: 1.8, w: 9, h: 0.8,
    fontSize: 34, fontFace: "Arial Black", color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
  });
  slide.addText("恳请批评指正", {
    x: 0.5, y: 2.6, w: 9, h: 0.5,
    fontSize: 20, fontFace: "Calibri", color: C.accent, align: "center", valign: "middle", margin: 0,
  });

  // Decorative line
  slide.addShape(pres.shapes.LINE, { x: 3.5, y: 3.3, w: 3, h: 0, line: { color: C.accent, width: 2 } });

  // Info
  slide.addText("答辩人：何成　　　指导教师：李维诗", {
    x: 1, y: 3.7, w: 8, h: 0.4,
    fontSize: 14, fontFace: "Calibri", color: "8899AA", align: "center", margin: 0,
  });
  slide.addText("合肥工业大学 · 仪器科学与光电工程学院 · 测控技术与仪器", {
    x: 1, y: 4.1, w: 8, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: "667788", align: "center", margin: 0,
  });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: C.accent } });
}

// ============================================================
// Write file
// ============================================================
pres.writeFile({ fileName: "水冷恒温箱答辩PPT.pptx" })
  .then(() => console.log("PPT generated: 水冷恒温箱答辩PPT.pptx"))
  .catch((err) => console.error("Error:", err));
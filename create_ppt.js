const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaFlask, FaExclamationTriangle, FaCheckCircle, FaCog, FaWater,
  FaTachometerAlt, FaThermometerHalf, FaTools, FaBookOpen, FaSearch,
  FaChartLine, FaClipboardList, FaShieldAlt, FaBolt, FaWrench
} = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

const C = {
  navy: "0C2340",
  darkBlue: "1B3A5C",
  teal: "0D7377",
  lightTeal: "14919B",
  accent: "E8A838",
  white: "FFFFFF",
  offWhite: "F4F6F8",
  lightGray: "E2E8F0",
  medGray: "8896A7",
  darkGray: "334155",
  text: "1E293B",
  red: "DC2626",
  orange: "EA580C",
  green: "059669",
  lightRed: "FEF2F2",
  lightOrange: "FFF7ED",
  lightGreen: "F0FDF4",
};

const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "化工实验教学";
  pres.title = "流体流动阻力实验 - 教学研讨";

  const icons = {
    flask: await iconToBase64Png(FaFlask, "#FFFFFF", 256),
    warning: await iconToBase64Png(FaExclamationTriangle, "#DC2626", 256),
    check: await iconToBase64Png(FaCheckCircle, "#059669", 256),
    cog: await iconToBase64Png(FaCog, "#0D7377", 256),
    water: await iconToBase64Png(FaWater, "#0D7377", 256),
    tacho: await iconToBase64Png(FaTachometerAlt, "#0D7377", 256),
    thermo: await iconToBase64Png(FaThermometerHalf, "#0D7377", 256),
    tools: await iconToBase64Png(FaTools, "#E8A838", 256),
    book: await iconToBase64Png(FaBookOpen, "#FFFFFF", 256),
    search: await iconToBase64Png(FaSearch, "#0D7377", 256),
    chart: await iconToBase64Png(FaChartLine, "#0D7377", 256),
    clipboard: await iconToBase64Png(FaClipboardList, "#FFFFFF", 256),
    shield: await iconToBase64Png(FaShieldAlt, "#DC2626", 256),
    bolt: await iconToBase64Png(FaBolt, "#E8A838", 256),
    wrench: await iconToBase64Png(FaWrench, "#E8A838", 256),
    warningWhite: await iconToBase64Png(FaExclamationTriangle, "#FFFFFF", 256),
    checkWhite: await iconToBase64Png(FaCheckCircle, "#FFFFFF", 256),
    cogWhite: await iconToBase64Png(FaCog, "#FFFFFF", 256),
    toolsWhite: await iconToBase64Png(FaTools, "#FFFFFF", 256),
    waterWhite: await iconToBase64Png(FaWater, "#FFFFFF", 256),
  };

  // ==================== SLIDE 1: Title ====================
  let slide = pres.addSlide();
  slide.background = { color: C.navy };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });
  slide.addImage({ data: icons.flask, x: 4.5, y: 0.6, w: 1, h: 1 });
  slide.addText("流体流动阻力实验", {
    x: 0.5, y: 1.7, w: 9, h: 1.2,
    fontSize: 40, fontFace: "Microsoft YaHei", color: C.white,
    bold: true, align: "center", margin: 0
  });
  slide.addText("教学研讨 · 实验操作详析与安全保障", {
    x: 0.5, y: 2.9, w: 9, h: 0.6,
    fontSize: 20, fontFace: "Microsoft YaHei", color: C.accent,
    align: "center", margin: 0
  });
  slide.addShape(pres.shapes.LINE, {
    x: 3.5, y: 3.7, w: 3, h: 0, line: { color: C.medGray, width: 1 }
  });
  slide.addText("化工原理实验课程", {
    x: 0.5, y: 4.0, w: 9, h: 0.5,
    fontSize: 14, fontFace: "Microsoft YaHei", color: C.medGray,
    align: "center", margin: 0
  });
  slide.addText("涵盖：操作规程 · 数据记录 · 问题分析 · 安全措施", {
    x: 0.5, y: 4.5, w: 9, h: 0.4,
    fontSize: 12, fontFace: "Microsoft YaHei", color: C.medGray,
    align: "center", margin: 0
  });

  // ==================== SLIDE 2: Outline ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.book, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("研讨内容概览", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const outlineItems = [
    { num: "01", title: "实验核心内容与测量原理", desc: "直管阻力 \u03BB-Re 关系 / 局部阻力系数 \u03B6 测定" },
    { num: "02", title: "实验装置关键参数", desc: "管径、管长、流量计量程、传感器规格" },
    { num: "03", title: "分步操作规程", desc: "从开机到关机的每一步阀门与按键操作" },
    { num: "04", title: "实验条件与数据记录", desc: "温度、流量范围、压差读数、记录表格设计" },
    { num: "05", title: "常见问题诊断", desc: "气泡、读数漂移、流量计失准等典型故障" },
    { num: "06", title: "安全风险与应对", desc: "电气安全、水锤效应、传感器过载保护" },
  ];

  outlineItems.forEach((item, i) => {
    const y = 1.15 + i * 0.72;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: y, w: 8.8, h: 0.62, fill: { color: C.white },
      shadow: makeShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: y, w: 0.08, h: 0.62, fill: { color: C.teal }
    });
    slide.addText(item.num, {
      x: 0.85, y: y, w: 0.6, h: 0.62,
      fontSize: 20, fontFace: "Consolas", color: C.teal, bold: true, valign: "middle", margin: 0
    });
    slide.addText(item.title, {
      x: 1.55, y: y, w: 4, h: 0.62,
      fontSize: 15, fontFace: "Microsoft YaHei", color: C.text, bold: true, valign: "middle", margin: 0
    });
    slide.addText(item.desc, {
      x: 5.5, y: y, w: 3.7, h: 0.62,
      fontSize: 11, fontFace: "Microsoft YaHei", color: C.medGray, valign: "middle", margin: 0
    });
  });

  // ==================== SLIDE 3: Core Experiment Content ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.clipboard, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("实验核心内容", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  // Left card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.15, w: 4.3, h: 4.1, fill: { color: C.white }, shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.15, w: 4.3, h: 0.55, fill: { color: C.teal }
  });
  slide.addImage({ data: icons.waterWhite, x: 0.7, y: 1.22, w: 0.35, h: 0.35 });
  slide.addText("实验 A：直管摩擦阻力", {
    x: 1.15, y: 1.15, w: 3.5, h: 0.55,
    fontSize: 15, fontFace: "Microsoft YaHei", color: C.white, bold: true, valign: "middle", margin: 0
  });
  slide.addText([
    { text: "测量目标", options: { bold: true, color: C.teal, fontSize: 13, breakLine: true } },
    { text: "在不同流速下测定光滑管与粗糙管的直管压降 \u0394P，计算摩擦系数 \u03BB，建立 \u03BB\u2013Re 双对数关系曲线。", options: { fontSize: 12, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "核心公式", options: { bold: true, color: C.teal, fontSize: 13, breakLine: true } },
    { text: "\u03BB = 2d\u00B7\u0394P / (\u03C1\u00B7l\u00B7u\u00B2)", options: { fontSize: 13, fontFace: "Consolas", color: C.darkGray, breakLine: true } },
    { text: "Re = d\u00B7u\u00B7\u03C1 / \u03BC", options: { fontSize: 13, fontFace: "Consolas", color: C.darkGray, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "关键参量", options: { bold: true, color: C.teal, fontSize: 13, breakLine: true } },
    { text: "光滑管：d = 0.008 m，l = 1.7 m\n粗糙管：d = 0.010 m，l = 1.7 m\n流量范围：10\u20131000 L/h\n数据组数：15\u201320 组", options: { fontSize: 11, color: C.text } },
  ], { x: 0.7, y: 1.85, w: 3.9, h: 3.2, valign: "top", lineSpacingMultiple: 1.1, margin: 0 });

  // Right card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.15, w: 4.3, h: 4.1, fill: { color: C.white }, shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.15, w: 4.3, h: 0.55, fill: { color: C.accent }
  });
  slide.addImage({ data: icons.cogWhite, x: 5.4, y: 1.22, w: 0.35, h: 0.35 });
  slide.addText("实验 B：局部阻力系数", {
    x: 5.85, y: 1.15, w: 3.5, h: 0.55,
    fontSize: 15, fontFace: "Microsoft YaHei", color: C.white, bold: true, valign: "middle", margin: 0
  });
  slide.addText([
    { text: "测量目标", options: { bold: true, color: C.accent, fontSize: 13, breakLine: true } },
    { text: "在大流量条件下测定阀门（特定开度）的局部阻力系数 \u03B6。", options: { fontSize: 12, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "核心公式", options: { bold: true, color: C.accent, fontSize: 13, breakLine: true } },
    { text: "\u03B6 = 2\u00B7\u0394P_local / (\u03C1\u00B7u\u00B2)", options: { fontSize: 13, fontFace: "Consolas", color: C.darkGray, breakLine: true } },
    { text: "h = 2(Pb\u2212Pb\u2032) \u2212 (Pa\u2212Pa\u2032)", options: { fontSize: 13, fontFace: "Consolas", color: C.darkGray, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "关键参量", options: { bold: true, color: C.accent, fontSize: 13, breakLine: true } },
    { text: "局部阻力管段：d = 0.020 m\n主管径：0.042 m\n测压点：远端（6）+ 近端（7, 15）\n需在大流量下测 5\u20138 组数据", options: { fontSize: 11, color: C.text } },
  ], { x: 5.4, y: 1.85, w: 3.9, h: 3.2, valign: "top", lineSpacingMultiple: 1.1, margin: 0 });

  // ==================== SLIDE 4: Equipment Parameters ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.cogWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("装置关键参数一览", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const paramHeaders = [
    [
      { text: "部件名称", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "编号", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "规格参数", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "操作要点", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
    ]
  ];
  const paramData = [
    ["光滑管", "17", "d = 0.008 m，l = 1.7 m", "用于湍流区 \u03BB\u2013Re 关系测定"],
    ["粗糙管", "18", "d = 0.010 m，l = 1.7 m", "对比不同 \u03B5/d 的影响"],
    ["局部阻力管段", "10", "d = 0.020 m", "阀门在特定开度下测 \u03B6"],
    ["大流量转子流量计", "16", "LZB-25，100\u20131000 L/h\n精度 1.5 级", "大 Re 段使用，注意浮子稳定"],
    ["小流量转子流量计", "15", "LZB-10，10\u2013100 L/h\n精度 2.5 级", "小 Re 段使用，避免层流数据缺失"],
    ["压差传感器", "12", "LXWY，0\u2013200 kPa", "缓慢调阀，防瞬时过载"],
    ["倒置 U 型管", "22", "配合放空阀 21", "小压差时使用，需彻底排气"],
    ["温度计", "25", "测水温", "每组数据前记录，不可用室温代替"],
  ];
  const paramRows = paramData.map(row => row.map(cell => ({
    text: cell, options: { fontSize: 10.5, fontFace: "Microsoft YaHei", color: C.text, valign: "middle", align: "center" }
  })));

  slide.addTable([...paramHeaders, ...paramRows], {
    x: 0.4, y: 1.1, w: 9.2,
    colW: [1.6, 0.7, 2.8, 4.1],
    border: { pt: 0.5, color: C.lightGray },
    rowH: Array(paramRows.length + 1).fill(0.5),
    autoPage: false
  });

  // ==================== SLIDE 5: Procedure 1/3 ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.toolsWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("分步操作规程（1/3）\u2014\u2014 开机与准备", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 24, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const steps1 = [
    { step: "1", title: "检查水箱液位", detail: "打开水箱盖目视检查，水位应在水箱 3/4 高度。不足时从自来水管补水，关闭水箱放水阀（23）。" },
    { step: "2", title: "关闭全部流量控制阀", detail: "将大流量调节阀（14）和小流量调节阀（27）旋至全关位置（顺时针旋紧）。确认放水阀（3、4、24）关闭。" },
    { step: "3", title: "初始化测压系统", detail: "检查倒置 U 型管放空阀（21）状态；关闭 U 型管进出水阀（11）。开启压差传感器电源，预热 10\u201315 分钟。记录传感器零点读数 P\u2080。" },
    { step: "4", title: "选通实验管路", detail: "直管实验：打开光滑管测压阀（9、19），关闭粗糙管测压阀（8、20）和局部阻力测压阀（6、7、15）。\n局部阻力实验：反之操作，确保仅一条管路处于通路状态。" },
    { step: "5", title: "启动离心泵", detail: "确认出口阀（14、27）全关 \u2192 按下泵启动按钮 \u2192 泵运转平稳后方可进行下一步。注意：本装置为自吸式泵，无需灌泵。" },
  ];

  steps1.forEach((s, i) => {
    const y = 1.1 + i * 0.88;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 0.78, fill: { color: C.white }, shadow: makeShadow()
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y + 0.14, w: 0.5, h: 0.5, fill: { color: C.teal }
    });
    slide.addText(s.step, {
      x: 0.7, y: y + 0.14, w: 0.5, h: 0.5,
      fontSize: 16, fontFace: "Consolas", color: C.white, bold: true, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.title, {
      x: 1.4, y: y + 0.02, w: 2.2, h: 0.78,
      fontSize: 13, fontFace: "Microsoft YaHei", color: C.teal, bold: true, valign: "middle", margin: 0
    });
    slide.addText(s.detail, {
      x: 3.6, y: y + 0.02, w: 5.7, h: 0.78,
      fontSize: 10.5, fontFace: "Microsoft YaHei", color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.05
    });
  });

  // ==================== SLIDE 6: Procedure 2/3 ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.toolsWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("分步操作规程（2/3）\u2014\u2014 排气与测量", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 24, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const steps2 = [
    { step: "6", title: "排除系统气泡", detail: "缓慢开启大流量阀（14），流量升至 800\u20131000 L/h，持续 2\u20133 分钟冲刷管路。逐个开启/关闭各测压阀，观察 U 型管液面：若两侧液面齐平则无气泡；若不齐平则反复开关阀门排气。" },
    { step: "7", title: "验证零点", detail: "关闭大流量阀（14），使流量为零。此时压差传感器应回到零点值 P\u2080（\u00B10.05 kPa 以内），倒置 U 型管两臂液面齐平。若偏差过大，重新排气。" },
    { step: "8", title: "沿程阻力数据采集", detail: "从最大流量开始逐步关小调节阀（先用大流量阀 14，再切换至小流量阀 27）。每档流量稳定 30\u201360 秒后，同时记录：转子流量计读数 Q、压差传感器读数 \u0394P（或 U 型管液面差 \u0394h）、水温 T。共采集 15\u201320 组。" },
    { step: "9", title: "切换管路", detail: "完成光滑管测量后：关闭光滑管测压阀（9、19）\u2192 打开粗糙管测压阀（8、20）\u2192 重复排气步骤 \u2192 重复数据采集。" },
  ];

  steps2.forEach((s, i) => {
    const y = 1.1 + i * 1.08;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 0.95, fill: { color: C.white }, shadow: makeShadow()
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y + 0.22, w: 0.5, h: 0.5, fill: { color: C.teal }
    });
    slide.addText(s.step, {
      x: 0.7, y: y + 0.22, w: 0.5, h: 0.5,
      fontSize: 16, fontFace: "Consolas", color: C.white, bold: true, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.title, {
      x: 1.4, y: y + 0.02, w: 2.0, h: 0.95,
      fontSize: 13, fontFace: "Microsoft YaHei", color: C.teal, bold: true, valign: "middle", margin: 0
    });
    slide.addText(s.detail, {
      x: 3.4, y: y + 0.02, w: 5.9, h: 0.95,
      fontSize: 10.5, fontFace: "Microsoft YaHei", color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.05
    });
  });

  // ==================== SLIDE 7: Procedure 3/3 ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.toolsWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("分步操作规程（3/3）\u2014\u2014 局部阻力与关机", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 24, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const steps3 = [
    { step: "10", title: "局部阻力测量", detail: "关闭直管测压阀（8、9、19、20）\u2192 打开局部阻力测压阀（远端 6 + 近端 7、15）\u2192 排气 \u2192 在大流量段（400\u20131000 L/h）采集 5\u20138 组数据。需同时读取四个测压点：Pa、Pa\u2032、Pb、Pb\u2032，计算 h = 2(Pb\u2212Pb\u2032) \u2212 (Pa\u2212Pa\u2032)。" },
    { step: "11", title: "关机程序", detail: "\u2460 缓慢关闭大流量阀（14）至全关\n\u2461 按下泵停止按钮\n\u2462 关闭压差传感器电源\n\u2463 关闭 U 型管进出水阀（11）\n\u2464 记录最终水温\n\u2465 整理实验台面" },
  ];

  steps3.forEach((s, i) => {
    const y = 1.1 + i * 1.3;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 1.15, fill: { color: C.white }, shadow: makeShadow()
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y + 0.32, w: 0.5, h: 0.5, fill: { color: C.accent }
    });
    slide.addText(s.step, {
      x: 0.7, y: y + 0.32, w: 0.5, h: 0.5,
      fontSize: 16, fontFace: "Consolas", color: C.white, bold: true, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.title, {
      x: 1.4, y: y + 0.02, w: 2.0, h: 1.15,
      fontSize: 14, fontFace: "Microsoft YaHei", color: C.accent, bold: true, valign: "middle", margin: 0
    });
    slide.addText(s.detail, {
      x: 3.4, y: y + 0.02, w: 5.9, h: 1.15,
      fontSize: 11, fontFace: "Microsoft YaHei", color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1
    });
  });

  // Key reminder box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.9, w: 9, h: 1.4, fill: { color: "FEF3C7" }, shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.9, w: 0.08, h: 1.4, fill: { color: C.accent }
  });
  slide.addImage({ data: icons.bolt, x: 0.8, y: 4.05, w: 0.4, h: 0.4 });
  slide.addText("操作关键提醒", {
    x: 1.35, y: 3.95, w: 3, h: 0.4,
    fontSize: 14, fontFace: "Microsoft YaHei", color: C.accent, bold: true, margin: 0
  });
  slide.addText([
    { text: "\u2460 切换管路时必须先关闭原管路测压阀再打开新管路，防止串压损坏传感器", options: { fontSize: 11, breakLine: true } },
    { text: "\u2461 调节阀门速度宜慢不宜快，尤其在大流量段，快速关阀可引起水锤效应", options: { fontSize: 11, breakLine: true } },
    { text: "\u2462 每次切换管路后必须重新排气并验证零点", options: { fontSize: 11, breakLine: true } },
    { text: "\u2463 先测大流量再逐步减小，可以复用排气效果，减少操作时间", options: { fontSize: 11 } },
  ], { x: 0.8, y: 4.45, w: 8.4, h: 0.8, color: C.text, fontFace: "Microsoft YaHei", margin: 0, lineSpacingMultiple: 1.1 });

  // ==================== SLIDE 8: Experimental Conditions ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.clipboard, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("实验条件与参数设定", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const condHeaders = [
    [
      { text: "参数", options: { fill: { color: C.darkBlue }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "光滑管实验", options: { fill: { color: C.darkBlue }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "粗糙管实验", options: { fill: { color: C.darkBlue }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "局部阻力实验", options: { fill: { color: C.darkBlue }, color: C.white, bold: true, fontSize: 12, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
    ]
  ];
  const condData = [
    ["管内径 d (m)", "0.008", "0.010", "0.020"],
    ["管长 l (m)", "1.7", "1.7", "\u2014"],
    ["流量范围 (L/h)", "10\u20131000", "10\u20131000", "400\u20131000"],
    ["流量计选择", "LZB-10 + LZB-25", "LZB-10 + LZB-25", "LZB-25"],
    ["数据组数", "15\u201320", "15\u201320", "5\u20138"],
    ["压差测量方式", "传感器 / U 型管", "传感器 / U 型管", "传感器"],
    ["水温范围 (\u00B0C)", "15\u201330（实测）", "15\u201330（实测）", "15\u201330（实测）"],
    ["流速范围 u (m/s)", "\u2248 0.055\u20135.5", "\u2248 0.035\u20133.5", "\u2248 0.35\u20130.88"],
    ["Re 范围", "\u2248 440\u201344000", "\u2248 350\u201335000", "\u2248 7000\u201317600"],
  ];
  const condRows = condData.map((row, ri) => row.map(cell => ({
    text: cell, options: {
      fontSize: 10.5, fontFace: "Microsoft YaHei", color: C.text, valign: "middle", align: "center",
      fill: { color: ri % 2 === 0 ? C.offWhite : C.white }
    }
  })));

  slide.addTable([...condHeaders, ...condRows], {
    x: 0.4, y: 1.1, w: 9.2,
    colW: [2.0, 2.4, 2.4, 2.4],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.42, ...Array(condData.length).fill(0.42)],
    autoPage: false
  });

  slide.addText("注：流速 u 由 u = Q/(\u03C0d\u00B2/4) 估算；Re 按 20\u00B0C 水的运动粘度 \u03BD \u2248 1.006\u00D710\u207B\u2076 m\u00B2/s 计算。实际值以测量水温查表为准。", {
    x: 0.5, y: 5.1, w: 9, h: 0.4,
    fontSize: 10, fontFace: "Microsoft YaHei", color: C.medGray, italic: true, margin: 0
  });

  // ==================== SLIDE 9: Data Table - Straight Pipe ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.clipboard, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("数据记录表：直管摩擦阻力", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  slide.addText("管路类型：\u25A1 光滑管 (d = 0.008 m)    \u25A1 粗糙管 (d = 0.010 m)          管长 l = 1.7 m          水温 T = ____\u00B0C", {
    x: 0.5, y: 1.0, w: 9, h: 0.4,
    fontSize: 11, fontFace: "Microsoft YaHei", color: C.text, margin: 0
  });

  const dataHeaders = [
    [
      { text: "序号", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "流量计\n读数 Q\n(L/h)", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 9, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "流速 u\n(m/s)", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 9, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "压差 \u0394P\n传感器\n(kPa)", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 9, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "U 型管\n\u0394h (mm)", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 9, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "Re", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "\u03BB", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "lg Re", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "lg \u03BB", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
    ]
  ];
  const emptyRows = [];
  for (let i = 1; i <= 8; i++) {
    emptyRows.push([
      { text: String(i), options: { fontSize: 10, fontFace: "Microsoft YaHei", color: C.medGray, align: "center", valign: "middle", fill: { color: i % 2 === 0 ? C.white : C.offWhite } } },
      ...Array(8).fill({ text: "", options: { fontSize: 10, align: "center", valign: "middle", fill: { color: i % 2 === 0 ? C.white : C.offWhite } } })
    ]);
  }

  slide.addTable([...dataHeaders, ...emptyRows], {
    x: 0.3, y: 1.5, w: 9.4,
    colW: [0.6, 1.1, 0.9, 1.2, 1.0, 1.2, 0.9, 1.0, 1.0],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.6, ...Array(8).fill(0.4)],
    autoPage: false
  });

  slide.addText("计算公式：u = Q/(3600\u00B7\u03C0\u00B7d\u00B2/4)；Re = u\u00B7d/\u03BD；\u03BB = 2\u00B7d\u00B7\u0394P/(\u03C1\u00B7l\u00B7u\u00B2)。后三列为计算列，实验中先记录 Q 和 \u0394P。", {
    x: 0.5, y: 5.1, w: 9, h: 0.4,
    fontSize: 9.5, fontFace: "Microsoft YaHei", color: C.medGray, italic: true, margin: 0
  });

  // ==================== SLIDE 10: Data Table - Local Resistance ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.clipboard, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("数据记录表：局部阻力系数", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  slide.addText("管径 d = 0.020 m          阀门类型：________          阀门开度：________          水温 T = ____\u00B0C", {
    x: 0.5, y: 1.0, w: 9, h: 0.4,
    fontSize: 11, fontFace: "Microsoft YaHei", color: C.text, margin: 0
  });

  const localHeaders = [
    [
      { text: "序号", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "Q (L/h)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "u (m/s)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "Pa (kPa)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "Pa\u2032(kPa)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "Pb (kPa)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "Pb\u2032(kPa)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "h (J/kg)", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "\u03B6", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 10, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
    ]
  ];
  const localEmpty = [];
  for (let i = 1; i <= 6; i++) {
    localEmpty.push([
      { text: String(i), options: { fontSize: 10, fontFace: "Microsoft YaHei", color: C.medGray, align: "center", valign: "middle", fill: { color: i % 2 === 0 ? C.white : C.offWhite } } },
      ...Array(8).fill({ text: "", options: { fontSize: 10, align: "center", valign: "middle", fill: { color: i % 2 === 0 ? C.white : C.offWhite } } })
    ]);
  }
  localEmpty.push([
    { text: "平均", options: { fontSize: 10, fontFace: "Microsoft YaHei", color: C.text, bold: true, align: "center", valign: "middle", fill: { color: "FEF3C7" } } },
    ...Array(7).fill({ text: "", options: { fontSize: 10, align: "center", valign: "middle", fill: { color: "FEF3C7" } } }),
    { text: "\u03B6\u0304 =", options: { fontSize: 10, fontFace: "Microsoft YaHei", color: C.text, bold: true, align: "center", valign: "middle", fill: { color: "FEF3C7" } } },
  ]);

  slide.addTable([...localHeaders, ...localEmpty], {
    x: 0.3, y: 1.5, w: 9.4,
    colW: [0.6, 1.0, 0.9, 1.1, 1.1, 1.1, 1.1, 1.2, 0.8],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.55, ...Array(7).fill(0.45)],
    autoPage: false
  });

  slide.addText("计算：h = 2(Pb\u2212Pb\u2032) \u2212 (Pa\u2212Pa\u2032)；\u03B6 = 2h / u\u00B2。取各组 \u03B6 的算术平均值作为最终结果。", {
    x: 0.5, y: 5.15, w: 9, h: 0.35,
    fontSize: 9.5, fontFace: "Microsoft YaHei", color: C.medGray, italic: true, margin: 0
  });

  // ==================== SLIDE 11: Common Problems 1/2 ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.warningWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("常见问题诊断（1/2）", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const problems1 = [
    {
      problem: "压差读数波动不稳",
      cause: "管路中残留气泡；调阀过快引起流量脉动；转子流量计浮子卡滞。",
      solution: "重新大流量冲刷排气；缓慢调阀等待 30 秒以上；轻敲流量计外壳使浮子自由浮动。"
    },
    {
      problem: "U 型管液面不齐（流量为零时）",
      cause: "导压管内有残余气泡；U 型管放空阀（21）未操作到位。",
      solution: "打开放空阀（21）排出 U 型管顶部空气，再关闭。反复 2\u20133 次至液面齐平。"
    },
    {
      problem: "\u03BB\u2013Re 曲线偏离 Blasius 方程",
      cause: "温度未实测（粘度偏差）；小流量段流量计精度不足；管壁结垢改变粗糙度。",
      solution: "严格实测水温并查物性表；小流量用 LZB-10（精度 2.5 级）；实验前检查管路清洁度。"
    },
  ];

  problems1.forEach((p, i) => {
    const y = 1.1 + i * 1.45;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 1.32, fill: { color: C.white }, shadow: makeShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 0.08, h: 1.32, fill: { color: C.red }
    });
    slide.addText(p.problem, {
      x: 0.8, y: y + 0.05, w: 8.5, h: 0.3,
      fontSize: 14, fontFace: "Microsoft YaHei", color: C.red, bold: true, margin: 0
    });
    slide.addText([
      { text: "原因：", options: { bold: true, color: C.teal, fontSize: 11 } },
      { text: p.cause, options: { color: C.text, fontSize: 11, breakLine: true } },
      { text: "对策：", options: { bold: true, color: C.green, fontSize: 11 } },
      { text: p.solution, options: { color: C.text, fontSize: 11 } },
    ], { x: 0.8, y: y + 0.4, w: 8.5, h: 0.85, fontFace: "Microsoft YaHei", margin: 0, lineSpacingMultiple: 1.15 });
  });

  // ==================== SLIDE 12: Common Problems 2/2 ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.warningWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("常见问题诊断（2/2）", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const problems2 = [
    {
      problem: "传感器示数突变或超量程",
      cause: "切换管路时未先关旧管路测压阀即开新管路阀，导致串压；调阀过于猛烈。",
      solution: "严格执行先关后开切换顺序；大流量段调阀角度不超过 5\u00B0/次；若传感器过载，需关机断电后重新校零。"
    },
    {
      problem: "局部阻力系数 \u03B6 离散度大",
      cause: "小流量下 \u0394P 太小、相对误差大；四个测压点之间有气泡。",
      solution: "局部阻力实验仅在大流量段（Q \u2265 400 L/h）进行，保证 \u0394P 具有足够的测量精度；逐个检查测压点排气。"
    },
    {
      problem: "实验时间过长 / 效率低",
      cause: "频繁切换管路和反复排气；记录数据时未分工。",
      solution: "先集中完成同一管路的全部流量点再切换；安排组员分工：一人调阀、一人读流量计、一人读压差、一人记录水温和时间。"
    },
  ];

  problems2.forEach((p, i) => {
    const y = 1.1 + i * 1.45;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 1.32, fill: { color: C.white }, shadow: makeShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 0.08, h: 1.32, fill: { color: C.orange }
    });
    slide.addText(p.problem, {
      x: 0.8, y: y + 0.05, w: 8.5, h: 0.3,
      fontSize: 14, fontFace: "Microsoft YaHei", color: C.orange, bold: true, margin: 0
    });
    slide.addText([
      { text: "原因：", options: { bold: true, color: C.teal, fontSize: 11 } },
      { text: p.cause, options: { color: C.text, fontSize: 11, breakLine: true } },
      { text: "对策：", options: { bold: true, color: C.green, fontSize: 11 } },
      { text: p.solution, options: { color: C.text, fontSize: 11 } },
    ], { x: 0.8, y: y + 0.4, w: 8.5, h: 0.85, fontFace: "Microsoft YaHei", margin: 0, lineSpacingMultiple: 1.15 });
  });

  // ==================== SLIDE 13: Safety Hazards ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.red } });
  slide.addImage({ data: icons.warningWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("安全风险分析与防护措施", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const safetyHeaders = [
    [
      { text: "风险类别", options: { fill: { color: C.red }, color: C.white, bold: true, fontSize: 11, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "风险描述", options: { fill: { color: C.red }, color: C.white, bold: true, fontSize: 11, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "后果", options: { fill: { color: C.red }, color: C.white, bold: true, fontSize: 11, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
      { text: "防护措施", options: { fill: { color: C.red }, color: C.white, bold: true, fontSize: 11, fontFace: "Microsoft YaHei", align: "center", valign: "middle" } },
    ]
  ];
  const safetyData = [
    ["电气安全", "湿手触碰泵开关或传感器电源", "触电", "操作电气设备前擦干双手；电源线远离水源；实验区地面保持干燥"],
    ["水锤效应", "快速关闭大流量阀（14）", "管路接头松脱\n传感器损坏", "关阀速度不超过 5\u00B0/次，间隔 3 秒；关机前先将流量降至最低"],
    ["传感器过载", "串压或瞬时压力冲击\n超过 200 kPa", "传感器永久损坏\n数据失真", "切换管路时先关后开；大流量 \u0394P 用传感器时关闭 U 型管阀门（11）"],
    ["地面湿滑", "水箱溢流或接头渗漏\n导致地面积水", "滑倒摔伤", "实验前检查所有接头密封；穿防滑鞋；地面有水立即擦干"],
    ["玻璃器件", "U 型管为玻璃材质\n碰撞或温度骤变可碎裂", "玻璃碎伤", "U 型管周围不放置硬物；不用热水冲洗；操作时动作轻柔"],
    ["泵空转", "水箱液位过低导致\n泵抽空", "泵叶轮磨损\n机械密封烧毁", "液位不低于 1/4 箱高；运行中注意泵声音是否异常"],
  ];
  const safetyRows = safetyData.map((row, ri) => row.map((cell, ci) => ({
    text: cell, options: {
      fontSize: 9.5, fontFace: "Microsoft YaHei", color: C.text, valign: "middle",
      align: ci === 0 ? "center" : "left",
      fill: { color: ri % 2 === 0 ? C.lightRed : C.white }
    }
  })));

  slide.addTable([...safetyHeaders, ...safetyRows], {
    x: 0.3, y: 1.05, w: 9.4,
    colW: [1.2, 2.2, 1.5, 4.5],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.42, ...Array(safetyData.length).fill(0.65)],
    autoPage: false
  });

  // ==================== SLIDE 14: Data Analysis Guide ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.clipboard, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("数据处理与结果分析指导", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  // Left card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.15, w: 4.3, h: 4.1, fill: { color: C.white }, shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.15, w: 4.3, h: 0.5, fill: { color: C.teal }
  });
  slide.addText("\u03BB\u2013Re 双对数曲线", {
    x: 0.7, y: 1.15, w: 3.9, h: 0.5,
    fontSize: 14, fontFace: "Microsoft YaHei", color: C.white, bold: true, valign: "middle", margin: 0
  });
  slide.addText([
    { text: "1. 绘图方法", options: { bold: true, color: C.teal, fontSize: 12, breakLine: true } },
    { text: "以 lg Re 为横坐标、lg \u03BB 为纵坐标，在双对数坐标纸上描点。光滑管和粗糙管数据分别绘制。", options: { fontSize: 11, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "2. 与 Blasius 方程比较", options: { bold: true, color: C.teal, fontSize: 12, breakLine: true } },
    { text: "对光滑管湍流区（3000 < Re < 10\u2075）：\n\u03BB = 0.3164 \u00B7 Re\u207B\u2070\u00B7\u00B2\u2075\n在同一图上绘出理论线，对比实验数据偏差。", options: { fontSize: 11, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "3. 层流区验证", options: { bold: true, color: C.teal, fontSize: 12, breakLine: true } },
    { text: "Re < 2100 段应呈直线，斜率 = \u22121（即 \u03BB = 64/Re）。若偏离说明未达充分发展层流或存在自然对流干扰。", options: { fontSize: 11, color: C.text } },
  ], { x: 0.7, y: 1.8, w: 3.9, h: 3.3, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

  // Right card
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.15, w: 4.3, h: 4.1, fill: { color: C.white }, shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.15, w: 4.3, h: 0.5, fill: { color: C.accent }
  });
  slide.addText("局部阻力系数处理", {
    x: 5.4, y: 1.15, w: 3.9, h: 0.5,
    fontSize: 14, fontFace: "Microsoft YaHei", color: C.white, bold: true, valign: "middle", margin: 0
  });
  slide.addText([
    { text: "1. 逐组计算 \u03B6", options: { bold: true, color: C.accent, fontSize: 12, breakLine: true } },
    { text: "由四个测压点读数计算：\nh = 2(Pb\u2212Pb\u2032) \u2212 (Pa\u2212Pa\u2032)\n\u03B6 = 2h / u\u00B2", options: { fontSize: 11, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "2. 求算术平均值", options: { bold: true, color: C.accent, fontSize: 12, breakLine: true } },
    { text: "\u03B6\u0304 = \u03A3\u03B6\u1D62 / n\n同时计算标准偏差 \u03C3 评估数据离散度。若某组 \u03B6 偏离平均值超过 2\u03C3，可视为异常数据剔除后重新计算。", options: { fontSize: 11, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "3. 结果对比", options: { bold: true, color: C.accent, fontSize: 12, breakLine: true } },
    { text: "将实测 \u03B6\u0304 与文献值（如闸阀全开 \u03B6 \u2248 0.17、半开 \u03B6 \u2248 4.5 等）进行比较，分析偏差原因。", options: { fontSize: 11, color: C.text } },
  ], { x: 5.4, y: 1.8, w: 3.9, h: 3.3, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

  // ==================== SLIDE 15: Precision Tips ====================
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addImage({ data: icons.checkWhite, x: 0.5, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("提高测量精度的技巧", {
    x: 1.2, y: 0.1, w: 8, h: 0.7,
    fontSize: 26, fontFace: "Microsoft YaHei", color: C.white, bold: true, margin: 0
  });

  const tips = [
    { icon: icons.thermo, title: "温度控制", text: "每 3\u20134 组数据记录一次水温。温度变化 > 2\u00B0C 时需分段计算物性参数（\u03C1、\u03BC）。不可用室温代替水温\u2014\u2014管路循环中水温会因泵做功而上升。" },
    { icon: icons.tacho, title: "流量计选择", text: "Q < 100 L/h 时使用小流量计 LZB-10（切断阀 26 切换）；Q > 100 L/h 时使用大流量计 LZB-25。跨量程过渡区读两个流量计取一致值。" },
    { icon: icons.search, title: "压差测量策略", text: "小流量（低 Re）时用倒置 U 型管\u2014\u2014灵敏度高于传感器。大流量时关闭 U 型管阀门（11）防止液面被吹出，改用传感器。两种方式在重叠流量段应相互校验。" },
    { icon: icons.water, title: "排气判据", text: "判据不是凭感觉，而是：(1) 零流量时传感器读数 = P\u2080 \u00B1 0.05 kPa；(2) U 型管两臂液面差 \u2264 1 mm。不满足时继续排气。" },
  ];

  tips.forEach((t, i) => {
    const y = 1.1 + i * 1.1;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 0.95, fill: { color: C.white }, shadow: makeShadow()
    });
    slide.addImage({ data: t.icon, x: 0.7, y: y + 0.22, w: 0.5, h: 0.5 });
    slide.addText(t.title, {
      x: 1.4, y: y + 0.05, w: 2, h: 0.95,
      fontSize: 14, fontFace: "Microsoft YaHei", color: C.teal, bold: true, valign: "middle", margin: 0
    });
    slide.addText(t.text, {
      x: 3.3, y: y + 0.05, w: 6, h: 0.95,
      fontSize: 10.5, fontFace: "Microsoft YaHei", color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1
    });
  });

  // ==================== SLIDE 16: Summary ====================
  slide = pres.addSlide();
  slide.background = { color: C.navy };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

  slide.addText("教学要点总结", {
    x: 0.5, y: 0.5, w: 9, h: 0.8,
    fontSize: 32, fontFace: "Microsoft YaHei", color: C.accent, bold: true, align: "center", margin: 0
  });

  const summaryItems = [
    "操作规范化：阀门操作遵循先关后开、慢调稳读原则",
    "排气彻底性：以定量判据（传感器零点、U 型管液面差）验证，而非目视估计",
    "数据可靠性：实测水温、选对流量计量程、重叠区交叉校验",
    "安全意识：干手操作电气、缓慢调阀防水锤、湿滑地面即时处理",
    "结果分析：\u03BB\u2013Re 曲线与 Blasius 方程的定量比较、局部阻力系数的统计检验",
  ];

  summaryItems.forEach((item, i) => {
    const y = 1.6 + i * 0.7;
    slide.addImage({ data: icons.checkWhite, x: 0.8, y: y + 0.08, w: 0.35, h: 0.35 });
    slide.addText(item, {
      x: 1.3, y: y, w: 8, h: 0.55,
      fontSize: 14, fontFace: "Microsoft YaHei", color: C.white, valign: "middle", margin: 0
    });
  });

  slide.addShape(pres.shapes.LINE, {
    x: 2.5, y: 5.0, w: 5, h: 0, line: { color: C.medGray, width: 1 }
  });
  slide.addText("流体流动阻力实验 \u00B7 教学研讨", {
    x: 0.5, y: 5.1, w: 9, h: 0.35,
    fontSize: 12, fontFace: "Microsoft YaHei", color: C.medGray, align: "center", margin: 0
  });

  await pres.writeFile({ fileName: "./fluid_resistance_experiment.pptx" });
  console.log("PPTX created successfully!");
}

main().catch(err => { console.error(err); process.exit(1); });

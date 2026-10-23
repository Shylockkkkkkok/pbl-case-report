/**
 * pptxgenjs 版式库 —— PBL 汇报 PPT（16:9，13.333 × 7.5 英寸）
 *
 * 用法：
 *   const { createDeck } = require("./ppt_helpers");
 *   const D = createDeck();
 *   const s = D.contentSlide("临床证据 · 1", 4);
 *   D.title(s, "证据：睡眠剥夺直接改变甲状腺轴");
 *   D.chartCard(s, D.M, D.TOP, 7.35, D.CONTENT_H, "chart_effect",
 *               "assets/chart_effect.png", "数据来源：……");
 *   D.refLine(s, "[2] Jain SP, et al. J Endocrinol Invest. 2026.");
 *   D.pptx.writeFile({ fileName: "out.pptx" });
 *
 * 运行前：export NODE_PATH="C:/Users/86188/AppData/Roaming/npm/node_modules"
 *
 * ⚠️ 画线时绝不能出现负的 w/h（PowerPoint 会打不开文件），用 lineShape()。
 */
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// ---------------- 设计系统：夜晚 × 高代谢 ----------------
const C = {
  night: "0E2233", night2: "17384F", night3: "24506B",
  teal: "1C7293", tealLt: "4FA3BF", amber: "E8871E",
  red: "C0392B", green: "0D9488", purple: "7A4E9B",
  white: "FFFFFF", bg: "F7FAFC", card: "FFFFFF", line: "DCE5EC",
  text: "1B2A38", text2: "4A5C6B", text3: "8296A6",
};
const FONT = "Noto Sans SC";              // 正文
const DISPLAY = "Source Han Sans CN";     // 标题

const W = 13.333, H = 7.5;
const M = 0.62;             // 左右页边距
const TITLE_Y = 0.40;
const TOP = 1.28;           // 内容区上沿
const BOT = 6.60;           // 内容区下沿
const CONTENT_H = BOT - TOP;

const sh = () => ({ type: "outer", color: "0E2233", blur: 7, offset: 2.5, angle: 90, opacity: 0.10 });
const shSoft = () => ({ type: "outer", color: "0E2233", blur: 4, offset: 1.5, angle: 90, opacity: 0.07 });

/** 图片真实宽高比表；用 PIL/sharp 读像素后填进来 */
const AR = {};

function fitBox(key, bx, by, bw, bh) {
  const ar = AR[key];
  let dw = bw, dh = dw / ar;
  if (dh > bh) { dh = bh; dw = dh * ar; }
  return { x: bx + (bw - dw) / 2, y: by + (bh - dh) / 2, w: dw, h: dh };
}

async function iconPng(Icon, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

/**
 * 画一条带方向的线。自动用 flipH/flipV 表达方向，避免负的 cx/cy。
 */
function lineShape(s, pptx, x1, y1, x2, y2, opts = {}) {
  const dx = x2 - x1, dy = y2 - y1;
  s.addShape(pptx.ShapeType.line, {
    x: Math.min(x1, x2), y: Math.min(y1, y2),
    w: Math.abs(dx), h: Math.abs(dy),
    flipH: dx < 0, flipV: dy < 0,
    line: {
      color: opts.color || C.amber,
      width: opts.width || 2.2,
      endArrowType: opts.arrow === false ? undefined : "triangle",
    },
  });
}

async function createDeck({ icons = {}, author = "", title = "" } = {}) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = author;
  pptx.title = title;

  // 预渲染图标： icons = { moon: [FaMoon, C.amberLt], ... }
  const I = {};
  for (const [k, [Icon, color]] of Object.entries(icons)) {
    I[k] = await iconPng(Icon, "#" + color);
  }

  function contentSlide(sectionLabel, pageNo) {
    const s = pptx.addSlide();
    s.background = { color: C.bg };
    s.addText(sectionLabel, {
      x: M, y: 0.20, w: 7, h: 0.24, fontFace: FONT, fontSize: 11,
      color: C.text3, valign: "middle",          // 不要加 charSpacing
    });
    s.addText(String(pageNo).padStart(2, "0"), {
      x: W - M - 1.0, y: 0.20, w: 1.0, h: 0.24, fontFace: FONT,
      fontSize: 11, color: C.text3, align: "right", valign: "middle",
    });
    return s;
  }

  /** 标题：左侧琥珀色竖条是视觉母题，不要用下划线 */
  function title(s, text) {
    s.addShape(pptx.ShapeType.rect, {
      x: M, y: TITLE_Y + 0.06, w: 0.085, h: 0.46, fill: { color: C.amber },
    });
    s.addText(text, {
      x: M + 0.22, y: TITLE_Y, w: W - 2 * M - 0.3, h: 0.58,
      fontFace: DISPLAY, fontSize: 26, bold: true, color: C.night, valign: "middle",
    });
  }

  /** 页脚参考文献：每页只要引用了文献就必须有 */
  function refLine(s, text) {
    s.addShape(pptx.ShapeType.line, {
      x: M, y: 6.76, w: W - 2 * M, h: 0, line: { color: C.line, width: 0.75 },
    });
    s.addText(
      [{ text: "参考文献  ", options: { color: C.amber, bold: true } },
       { text: text, options: { color: C.text3 } }],
      { x: M, y: 6.80, w: W - 2 * M, h: 0.62, fontFace: FONT, fontSize: 8.8,
        valign: "top", lineSpacingMultiple: 1.18 }
    );
  }

  function card(s, x, y, w, h, fill) {
    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0.09,
      fill: { color: fill || C.card },
      line: { color: C.line, width: 0.75 },
      shadow: shSoft(),
    });
  }

  function iconHead(s, icon, x, y, w, label, color) {
    s.addShape(pptx.ShapeType.ellipse, { x, y, w: 0.34, h: 0.34, fill: { color } });
    s.addImage({ data: icon, x: x + 0.055, y: y + 0.055, w: 0.23, h: 0.23 });
    s.addText(label, {
      x: x + 0.44, y, w: w - 0.44, h: 0.34, fontFace: FONT, fontSize: 13.5,
      bold: true, color: C.night, valign: "middle",
    });
  }

  /**
   * 图表卡片：按真实宽高比缩放后居中，图注紧贴图片底边 —— 保证永不重叠。
   * key 必须在 AR 表里有对应宽高比。
   */
  function chartCard(s, x, y, w, h, key, path, captionText) {
    card(s, x, y, w, h);
    const capH = 0.30, padTop = 0.28, padX = 0.30;
    const boxW = w - 2 * padX;
    const boxH = h - padTop - capH - 0.22;
    const ar = AR[key];
    let dw = boxW, dh = dw / ar;
    if (dh > boxH) { dh = boxH; dw = dh * ar; }
    const dx = x + (w - dw) / 2;
    const dy = y + padTop + (boxH - dh) / 2;
    s.addImage({ path, x: dx, y: dy, w: dw, h: dh });
    s.addText(captionText, {
      x: x + 0.2, y: dy + dh + 0.06, w: w - 0.4, h: capH,
      fontFace: FONT, fontSize: 10.4, color: C.text3, align: "center", valign: "middle",
    });
  }

  /** 右侧纵向要点卡：pts = [[标题, 正文, 颜色, 图标key, 可选角标], ...] */
  function pointCards(s, x, y, w, cardH, gap, pts) {
    pts.forEach(([t, d, col, ic, tag], i) => {
      const cy = y + i * (cardH + gap);
      card(s, x, cy, w, cardH);
      s.addShape(pptx.ShapeType.rect, {
        x, y: cy + 0.14, w: 0.075, h: cardH - 0.28, fill: { color: col },
      });
      if (I[ic]) s.addImage({ data: I[ic], x: x + 0.22, y: cy + 0.19, w: 0.25, h: 0.25 });
      s.addText(
        [{ text: t + "  ", options: { bold: true, color: col, fontSize: 13.4 } },
         { text: tag || "", options: { color: C.text3, fontSize: 10.2 } }],
        { x: x + 0.58, y: cy + 0.15, w: w - 0.78, h: 0.33, fontFace: FONT, valign: "middle" });
      s.addText(d, {
        x: x + 0.26, y: cy + 0.54, w: w - 0.50, h: cardH - 0.66,
        fontFace: FONT, fontSize: 12.2, color: C.text, valign: "top",
        lineSpacingMultiple: 1.22,
      });
    });
  }

  /** 深色满版页（封面 / 参考文献 / 致谢）常用装饰 */
  function glow(s, x, y, d, color) {
    s.addShape(pptx.ShapeType.ellipse, {
      x, y, w: d, h: d, fill: { color }, line: { color, width: 0 },
    });
  }

  return {
    pptx, I,
    contentSlide, title, refLine, card, iconHead, chartCard, pointCards, glow,
    line: (s, x1, y1, x2, y2, o) => lineShape(s, pptx, x1, y1, x2, y2, o),
  };
}

module.exports = { createDeck, C, FONT, DISPLAY, AR, fitBox, lineShape, W, H, M, TOP, BOT, CONTENT_H };

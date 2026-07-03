#!/usr/bin/env node
/**
 * make_profile_docs.js — 마크앤컴퍼니 1페이지 기업 프로필 생성기
 *
 * Usage:
 *   node make_profile_docs.js \
 *     --companies /path/companies.json \
 *     --config    /path/config.json \
 *     --out       /path/out \
 *     [--ext      /path/ext_data.json]
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, ExternalHyperlink,
  BorderStyle, WidthType, ShadingType,
  VerticalAlign, Header, Footer, PageNumber, TabStopType, TabStopPosition,
  PageBreak, ImageRun,
} = require("docx");

// ─── CLI args ───────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
function getArg(name) {
  const i = args.indexOf(name);
  return i !== -1 ? args[i + 1] : null;
}
const companiesPath = getArg("--companies");
const configPath    = getArg("--config");
const outDir        = getArg("--out");
const extPath       = getArg("--ext");

if (!companiesPath || !configPath || !outDir) {
  console.error("Usage: node make_profile_docs.js --companies <path> --config <path> --out <dir> [--ext <path>]");
  process.exit(1);
}

const companies = JSON.parse(fs.readFileSync(companiesPath, "utf-8"));
const cfg       = JSON.parse(fs.readFileSync(configPath, "utf-8"));
const ext       = extPath && fs.existsSync(extPath) ? JSON.parse(fs.readFileSync(extPath, "utf-8")) : {};

// ─── Design Tokens ──────────────────────────────────────────────────────────
const FONT         = "Pretendard";
const ACCENT       = "1F3B57";   // 딥 네이비
const ACCENT_LIGHT = "EDF1F5";   // 연한 블루 틴트
const SUBTLE       = "6B7280";   // 미드 그레이
const LINE         = "D7DCE3";   // 연한 그레이
const PANEL_FILL   = "F6F4ED";   // 웜 아이보리
const PANEL_ACCENT = "8A7B4E";   // 머티드 골드
const BODY         = "26282B";
const LINK_COLOR   = "1155CC";

const PAGE_WIDTH    = 12240;
const PAGE_HEIGHT   = 15840;
const MARGIN        = 1080;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 10080
const LABEL_W       = 2400;
const VALUE_W       = CONTENT_WIDTH - LABEL_W; // 7680

// ─── Helpers ────────────────────────────────────────────────────────────────
function isEmpty(v) { return v === null || v === undefined || v === "" || v === "-" || v === 0 || v === "0"; }
function val(v, fallback = "-") { return isEmpty(v) ? fallback : String(v); }

function fmtWon(n) {
  if (isEmpty(n)) return null;
  const num = Number(n);
  if (Number.isNaN(num)) return String(n);
  return num.toLocaleString("ko-KR") + "원";
}

function today() {
  const d = new Date();
  return `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")}`;
}

// ─── Typography primitives ──────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, ...opts });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 250, after: 105 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 4 } },
    children: [run(text, { bold: true, size: 24, color: ACCENT })],
  });
}

function bodyPara(text) {
  const lines = String(text).split("\n");
  return lines.map((line, i) =>
    new Paragraph({
      spacing: { after: i === lines.length - 1 ? 70 : 40, line: 300 },
      children: [run(line, { size: 21, color: BODY })],
    })
  );
}

function bulletList(items) {
  if (!items.length) {
    return [new Paragraph({ spacing: { after: 60 }, children: [run("-", { size: 21, color: SUBTLE })] })];
  }
  return items.map(t =>
    new Paragraph({
      numbering: { reference: "profile-bullets", level: 0 },
      spacing: { after: 60 },
      children: [run(t, { size: 21, color: BODY })],
    })
  );
}

// ─── 성장 지표 (자산·부채·자본·고용 성장률처럼 "라벨 + 기간별 증감 2개" 형태의 컴팩트 표) ──
const GROWTH_UP = "1E8E5A";
const GROWTH_DOWN = "C0392B";

function fmtPct(v) {
  if (v === null || v === undefined || v === "" || isNaN(Number(v))) return null;
  const n = Number(v) * 100;
  const sign = n > 0 ? "+" : "";
  return `${sign}${n.toFixed(1)}%`;
}

function growthCell(rawValue, compact) {
  const size = compact ? 15 : 19;
  const pct = fmtPct(rawValue);
  if (pct === null) return [run("-", { size, color: SUBTLE })];
  const n = Number(rawValue);
  const color = n >= 0 ? GROWTH_UP : GROWTH_DOWN;
  const arrow = n >= 0 ? "▲ " : "▼ ";
  return [run(arrow + pct, { size, bold: true, color })];
}

function growthIndicatorRow(label, val1, val2, labelW, colW) {
  const b = { style: BorderStyle.SINGLE, size: 2, color: LINE };
  const vm = 70; // 컴팩트 표: 기본 오버뷰 행(90)보다 좁게
  const cell = (children, width, right) => new TableCell({
    width: { size: width, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
    verticalAlign: VerticalAlign.CENTER, margins: { top: vm, bottom: vm, left: 160, right: right ?? 160 },
    children: [new Paragraph({ alignment: right === 0 ? undefined : undefined, children })],
  });
  return new TableRow({ children: [
    new TableCell({
      width: { size: labelW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      shading: { fill: ACCENT_LIGHT, type: ShadingType.CLEAR },
      verticalAlign: VerticalAlign.CENTER, margins: { top: vm, bottom: vm, left: 160, right: 120 },
      children: [new Paragraph({ children: [run(label, { bold: true, size: 19, color: ACCENT })] })],
    }),
    cell(growthCell(val1), colW),
    cell(growthCell(val2), colW),
  ]});
}

// cfg.growthIndicators: { caption, rows: [{ label, fields: [field1, field2] }] }
function buildGrowthIndicators(cfg, company) {
  if (!cfg || !cfg.rows || !cfg.rows.length) return [];
  const labelW = 2400;
  const colW = (CONTENT_WIDTH - labelW) / 2;
  const rows = cfg.rows.map(r => growthIndicatorRow(r.label, company[r.fields[0]], company[r.fields[1]], labelW, colW));
  const out = [];
  if (cfg.caption) {
    out.push(new Paragraph({
      spacing: { after: 40 },
      children: [run(cfg.caption, { size: 15, color: SUBTLE })],
    }));
  }
  out.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [labelW, colW, colW],
    rows,
  }));
  return out;
}


// ─── 데이터 매트릭스 표 (예: 고용/재무 현황 — "구분 + 연도별 값 + 증감률" 여러 행) ──────
function fmtEok(won) {
  if (won === null || won === undefined || won === "" || isNaN(Number(won))) return null;
  const n = Number(won);
  if (n === 0) return null; // 정확히 0원인 경우만 "-"로 표시
  const abs = Math.abs(n);
  // 금액 규모에 맞는 단위를 자동 선택 (억 단위로 강제 표기하면 소액 기업은 "0.0억"처럼 실제 규모가 안 보이므로)
  if (abs >= 1e8) return `${(n / 1e8).toFixed(1)}억`;
  if (abs >= 1e6) return `${(n / 1e6).toFixed(1)}백만원`;
  if (abs >= 1e4) return `${(n / 1e4).toFixed(1)}만원`;
  return `${n.toLocaleString("ko-KR")}원`;
}

function fmtNum(v, unit = "") {
  if (v === null || v === undefined || v === "" || isNaN(Number(v))) return null;
  const n = Number(v);
  const s = Number.isInteger(n) ? String(n) : n.toFixed(1);
  return `${s}${unit}`;
}

function dataMatrixCell(rawValue, type, unit, compact) {
  const size = compact ? 15 : 18;
  if (type === "percent") return growthCell(rawValue, compact);
  const text = type === "money" ? fmtEok(rawValue) : fmtNum(rawValue, unit);
  if (text === null) return [run("-", { size, color: SUBTLE })];
  return [run(text, { size, color: BODY })];
}

function dataMatrixRow(rowCfg, company, labelW, colW, isHeader, compact) {
  const b = { style: BorderStyle.SINGLE, size: 2, color: LINE };
  const vm = compact ? 32 : 60;
  const headSize = compact ? 14 : 17;
  const labelSize = compact ? 15 : 18;
  if (isHeader) {
    const headCell = (text, width) => new TableCell({
      width: { size: width, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      shading: { fill: ACCENT_LIGHT, type: ShadingType.CLEAR },
      verticalAlign: VerticalAlign.CENTER, margins: { top: vm, bottom: vm, left: 80, right: 80 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(text, { bold: true, size: headSize, color: ACCENT })] })],
    });
    return new TableRow({ children: [headCell(rowCfg.label ?? "구분", labelW), ...rowCfg.columns.map(c => headCell(c, colW))] });
  }
  const cells = [
    new TableCell({
      width: { size: labelW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      shading: { fill: ACCENT_LIGHT, type: ShadingType.CLEAR },
      verticalAlign: VerticalAlign.CENTER, margins: { top: vm, bottom: vm, left: 140, right: 100 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(rowCfg.label, { bold: true, size: labelSize, color: ACCENT })] })],
    }),
    ...rowCfg.fields.map((f, i) => new TableCell({
      width: { size: colW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      verticalAlign: VerticalAlign.CENTER, margins: { top: vm, bottom: vm, left: 80, right: 80 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: dataMatrixCell(company[f], rowCfg.types[i], rowCfg.unit, compact) })],
    })),
  ];
  return new TableRow({ children: cells });
}

// tableCfg: { heading | caption, chart: { path, aspect, widthPx }, columns: [...라벨], rows: [...](선택), compact }
// heading: 일반 섹션 헤딩(굵은 밑줄) / caption: 작은 회색 캡션 한 줄(공간 절약용) / compact: 폰트·여백 축소
// chart가 있으면 표 바로 위에 이미지(전체 폭 기준 자동 리사이즈)를 삽입
// rows가 없으면(=차트만 있는 시계열 차트 등) 표는 생략하고 헤딩+차트만 렌더링
function buildDataTable(tableCfg, company) {
  if (!tableCfg) return [];
  const hasRows = tableCfg.rows && tableCfg.rows.length;
  const out = [];
  if (tableCfg.heading) out.push(sectionHeading(tableCfg.heading));
  else if (tableCfg.caption) out.push(new Paragraph({ spacing: { before: 100, after: 40 }, children: [run(tableCfg.caption, { bold: true, size: 19, color: ACCENT })] }));

  if (tableCfg.chart && fs.existsSync(tableCfg.chart.path)) {
    const widthPx = tableCfg.chart.widthPx ?? 620;
    const heightPx = Math.round(widthPx / (tableCfg.chart.aspect ?? 3.8));
    out.push(new Paragraph({
      spacing: { after: 90 },
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({ data: fs.readFileSync(tableCfg.chart.path), transformation: { width: widthPx, height: heightPx }, type: "png" })],
    }));
  } else if (tableCfg.chartField) {
    // 기업마다 다른 차트 이미지를 쓸 때: company[chartField]에 저장된 경로를 사용
    // (데이터가 부족해 차트를 못 만든 기업은 company[chartField]가 null이라 자동으로 스킵됨)
    const chartPath = company[tableCfg.chartField];
    if (chartPath && fs.existsSync(chartPath)) {
      const widthPx = tableCfg.chartWidthPx ?? 620;
      const heightPx = Math.round(widthPx / (tableCfg.chartAspect ?? 3.8));
      out.push(new Paragraph({
        spacing: { after: 90 },
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: fs.readFileSync(chartPath), transformation: { width: widthPx, height: heightPx }, type: "png" })],
      }));
    }
  }

  if (!hasRows) return out; // 표 없이 헤딩+차트만

  const n = tableCfg.columns.length;
  const labelW = 2200;
  const colW = (CONTENT_WIDTH - labelW) / n;
  const compact = !!tableCfg.compact;
  const headerRow = dataMatrixRow({ columns: tableCfg.columns }, company, labelW, colW, true, compact);
  const dataRows = tableCfg.rows.map(r => dataMatrixRow(r, company, labelW, colW, false, compact));
  out.push(new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [labelW, ...Array(n).fill(colW)],
    rows: [headerRow, ...dataRows],
  }));
  return out;
}



// spanUnits: 1 = 기본 한 줄 높이, 2 = 두 줄 높이(예: 모집분야를 크게 표시해 반대편 표와 높이를 맞출 때)
function rowVMargin(spanUnits) {
  // spanUnits=1일 때 90(기존값)이 되도록 보정한 공식: 210*n - 120
  return Math.round(210 * spanUnits - 120);
}

function overviewRow(label, valueRuns, labelW = LABEL_W, valueW = VALUE_W, spanUnits = 1, extraTwips = 0) {
  const b = { style: BorderStyle.SINGLE, size: 2, color: LINE };
  const vm = rowVMargin(spanUnits);
  const extraTop = Math.floor(extraTwips / 2);
  const extraBottom = extraTwips - extraTop;
  return new TableRow({ children: [
    new TableCell({
      width: { size: labelW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      shading: { fill: ACCENT_LIGHT, type: ShadingType.CLEAR },
      verticalAlign: VerticalAlign.CENTER, margins: { top: vm + extraTop, bottom: vm + extraBottom, left: 160, right: 120 },
      children: [new Paragraph({ children: [run(label, { bold: true, size: 20, color: ACCENT })] })],
    }),
    new TableCell({
      width: { size: valueW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      verticalAlign: VerticalAlign.CENTER, margins: { top: vm + extraTop, bottom: vm + extraBottom, left: 160, right: 160 },
      children: [new Paragraph({ children: valueRuns })],
    }),
  ]});
}

// 좌우 표의 줄 수(단위 높이 합)가 다를 때 남는 만큼 채우는 빈 행 (내용 없이 테두리만 정상 표시 — 표 하단선을 실제로 맞추기 위함)
function emptyOverviewRow(labelW, valueW, spanUnits = 1) {
  const b = { style: BorderStyle.SINGLE, size: 2, color: LINE };
  const vm = rowVMargin(spanUnits);
  return new TableRow({ children: [
    new TableCell({
      width: { size: labelW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      margins: { top: vm, bottom: vm, left: 160, right: 120 },
      children: [new Paragraph({ children: [run("", { size: 20 })] })],
    }),
    new TableCell({
      width: { size: valueW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      margins: { top: vm, bottom: vm, left: 160, right: 160 },
      children: [new Paragraph({ children: [run("", { size: 20 })] })],
    }),
  ]});
}

// 실측(pt 단위) 기반 정밀 보정용 필러 행 — 목표 총 높이(twips)를 정확히 맞춤
// (텍스트 1줄 기본 높이 ≈244twips를 빼고 나머지를 margin으로 채움). 테두리는 일반 행과 동일하게 표시.
const EMPTY_LINE_TWIPS = 244;
function emptyOverviewRowExact(labelW, valueW, totalTwips) {
  const b = { style: BorderStyle.SINGLE, size: 2, color: LINE };
  const pad = Math.max(0, Math.round(totalTwips - EMPTY_LINE_TWIPS));
  const top = Math.floor(pad / 2);
  const bottom = pad - top;
  return new TableRow({ children: [
    new TableCell({
      width: { size: labelW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      margins: { top, bottom, left: 160, right: 120 },
      children: [new Paragraph({ children: [run("", { size: 20 })] })],
    }),
    new TableCell({
      width: { size: valueW, type: WidthType.DXA }, borders: { top: b, bottom: b, left: b, right: b },
      margins: { top, bottom, left: 160, right: 160 },
      children: [new Paragraph({ children: [run("", { size: 20 })] })],
    }),
  ]});
}

// 오버뷰 표를 1개(full-width) 또는 여러 개(side-by-side, 공간 절약) 그룹으로 렌더링
// cfg.overviewGroups: [{ rows: [...] }, { rows: [...] }] → 그룹 수만큼 나란히 배치
// 각 row에 "span": 2 를 주면 해당 행만 두 배 높이로 표시됨 (반대편 표와 높이를 맞출 때 사용)
// cfg.overviewRows(구버전 flat 배열)는 자동으로 단일 그룹으로 취급되어 그대로 동작(하위호환)
function buildOverviewGroups(groups, company) {
  if (!groups || !groups.length) return [];

  if (groups.length === 1) {
    return [new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [LABEL_W, VALUE_W],
      rows: groups[0].rows.map(rowCfg => overviewRow(rowCfg.label, resolveValueRuns(rowCfg, company), LABEL_W, VALUE_W, rowCfg.span || 1)),
    })];
  }

  // 2개 이상 그룹: 나란히 배치 (표 자체가 길어지는 것을 방지, 공간 절약)
  const GAP = 360;
  const n = groups.length;
  const groupW = (CONTENT_WIDTH - GAP * (n - 1)) / n;
  const miniLabelW = 2100;
  const miniValueW = groupW - miniLabelW;
  const none = { style: BorderStyle.NONE };
  const wrapBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

  // 단위 높이 합(각 row의 span 합)으로 그룹 간 높이를 비교 — 단순 row 개수가 아니라 span까지 반영
  const groupUnits = groups.map(g => g.rows.reduce((sum, r) => sum + (r.span || 1), 0));
  const maxUnits = Math.max(...groupUnits);

  const cells = [];
  const colWidths = [];
  const exactFiller = company.__overviewFillerTwips; // { groupIndex: extraTwips } - 실측 기반 정밀 보정용 (선택적)
  groups.forEach((g, i) => {
    const extra = exactFiller && exactFiller[i] > 0 ? exactFiller[i] : 0;
    // 보정값이 있으면 마지막 행에 몰아주지 않고 모든 행에 균등하게 나눠서 자연스럽게 꽉 채움
    const perRowExtra = extra > 0 ? Math.floor(extra / g.rows.length) : 0;
    let usedExtra = 0;
    const dataRows = g.rows.map((rowCfg, idx) => {
      let rowExtra = perRowExtra;
      if (idx === g.rows.length - 1) rowExtra = extra - usedExtra; // 나머지(반올림 오차)는 마지막 행에서 흡수
      usedExtra += rowExtra;
      return overviewRow(rowCfg.label, resolveValueRuns(rowCfg, company), miniLabelW, miniValueW, rowCfg.span || 1, rowExtra);
    });
    let fillerRows = [];
    if (extra === 0) {
      const remaining = maxUnits - groupUnits[i];
      fillerRows = remaining > 0 ? [emptyOverviewRow(miniLabelW, miniValueW, remaining)] : [];
    }
    cells.push(new TableCell({
      width: { size: groupW, type: WidthType.DXA },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Table({
        width: { size: groupW, type: WidthType.DXA },
        columnWidths: [miniLabelW, miniValueW],
        rows: [...dataRows, ...fillerRows],
      })],
    }));
    colWidths.push(groupW);
    if (i < n - 1) {
      cells.push(new TableCell({
        width: { size: GAP, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [new Paragraph({ children: [] })],
      }));
      colWidths.push(GAP);
    }
  });

  return [new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: colWidths,
    borders: wrapBorders,
    rows: [new TableRow({ children: cells })],
  })];
}

function resolveValueRuns(rowCfg, company) {
  const raw = company[rowCfg.field];
  if (rowCfg.type === "money") {
    const fmt = fmtWon(raw);
    return [run(fmt ?? "-", { size: 20, color: fmt ? BODY : SUBTLE })];
  }
  if (rowCfg.type === "link") {
    if (isEmpty(raw)) return [run("-", { size: 20, color: SUBTLE })];
    return [new ExternalHyperlink({ link: raw, children: [run(raw, { size: 20, color: LINK_COLOR, underline: {} })] })];
  }
  // plain (default)
  return [run(val(raw), { size: 20, color: BODY })];
}

// ─── External data panel ─────────────────────────────────────────────────────
function buildPanel(company, panelCfg) {
  if (!panelCfg || !panelCfg.enabled) return [];
  const key = String(company[panelCfg.linkKey] ?? "");
  const rec = ext[key];
  if (!rec) return [];

  const fields = panelCfg.fields || {};
  const aiSummary = !isEmpty(rec[fields.aiSummary]) ? String(rec[fields.aiSummary]) : null;
  const investors = !isEmpty(rec[fields.investors]) ? String(rec[fields.investors]) : null;
  const link      = !isEmpty(rec[fields.link])      ? String(rec[fields.link])      : null;
  if (!aiSummary && !investors && !link) return [];

  const panelChildren = [];
  if (aiSummary) {
    panelChildren.push(new Paragraph({
      spacing: { after: investors || link ? 70 : 0 },
      children: [
        run("AI 요약  ", { size: 19, bold: true, color: PANEL_ACCENT }),
        run(aiSummary, { size: 19, color: "3A3A3A" }),
      ],
    }));
  }
  if (investors) {
    panelChildren.push(new Paragraph({
      spacing: { after: link ? 70 : 0 },
      children: [
        run("주요 투자자  ", { size: 19, bold: true, color: PANEL_ACCENT }),
        run(investors, { size: 19, color: BODY }),
      ],
    }));
  }
  if (link) {
    panelChildren.push(new Paragraph({
      spacing: { after: 0 },
      children: [new ExternalHyperlink({ link, children: [
        run(`${panelCfg.title ?? "외부 데이터"}에서 기업정보 보기 ↗`, { size: 18, color: LINK_COLOR, underline: {} }),
      ]})],
    }));
  }

  const pb = { style: BorderStyle.SINGLE, size: 2, color: "E3DCC6" };
  return [
    new Paragraph({
      spacing: { before: 280, after: 80 },
      children: [
        run(panelCfg.title ?? "외부 데이터", { size: 18, bold: true, color: PANEL_ACCENT }),
        run(`  ${panelCfg.subtitle ?? ""}`, { size: 15, color: SUBTLE }),
      ],
    }),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [CONTENT_WIDTH],
      rows: [new TableRow({ children: [new TableCell({
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        borders: { top: pb, bottom: pb, left: pb, right: pb },
        shading: { fill: PANEL_FILL, type: ShadingType.CLEAR },
        margins: { top: 130, bottom: 130, left: 180, right: 180 },
        children: panelChildren,
      })]})],
    }),
  ];
}

// ─── 2-column section (e.g. 관심 트랙 / 기대사항) ────────────────────────────
function twoColumnSection(sectionCfg, company) {
  if (!sectionCfg) return [];
  const halfW = (CONTENT_WIDTH - 400) / 2;
  const left  = (sectionCfg.leftFields  ?? []).map(f => company[f]).filter(Boolean);
  const right = (sectionCfg.rightFields ?? []).map(f => company[f]).filter(Boolean);
  const none  = { style: BorderStyle.NONE };
  const borders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

  return [
    sectionHeading(sectionCfg.heading ?? ""),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [halfW, 400, halfW], borders,
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: halfW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 120 },
          children: [
            new Paragraph({ spacing: { after: 70 }, children: [run(sectionCfg.leftLabel ?? "", { bold: true, size: 20, color: "111315" })] }),
            ...bulletList(left),
          ],
        }),
        new TableCell({ width: { size: 400, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [new Paragraph({ children: [] })] }),
        new TableCell({
          width: { size: halfW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 120, right: 0 },
          children: [
            new Paragraph({ spacing: { after: 70 }, children: [run(sectionCfg.rightLabel ?? "", { bold: true, size: 20, color: "111315" })] }),
            ...bulletList(right),
          ],
        }),
      ]})],
    }),
  ];
}

// ─── Document builder ────────────────────────────────────────────────────────
function buildDoc(company) {
  const children = [];
  const dateStr = today();

  // Title
  children.push(
    new Paragraph({ spacing: { after: 60 }, children: [run(val(company[cfg.nameField]), { size: 38, bold: true, color: "111315" })] }),
    new Paragraph({
      spacing: { after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: ACCENT, space: 6 } },
      children: [run(val(company[cfg.subtitleField] ?? ""), { size: 23, bold: true, color: ACCENT })],
    })
  );

  // Overview table (overviewGroups 있으면 나란히 배치, 없으면 overviewRows 단일 표로 하위호환)
  const overviewGroups = cfg.overviewGroups ?? (cfg.overviewRows ? [{ rows: cfg.overviewRows }] : []);
  children.push(...buildOverviewGroups(overviewGroups, company));

  // 길이 편차가 큰 필드(예: 성과키워드)는 좌우 2단 표 밖, 전체 폭으로 별도 표시
  // (2단 표 안에 두면 줄바꿈 발생 시 반대편 표와 높이가 어긋나기 때문)
  if (cfg.wideOverviewRows && cfg.wideOverviewRows.length) {
    children.push(new Paragraph({ spacing: { after: 0 }, children: [] }));
    children.push(...buildOverviewGroups([{ rows: cfg.wideOverviewRows }], company));
  }

  // 성장 지표 (자산·부채·자본·고용 성장률 등, 있을 때만)
  if (cfg.growthIndicators) {
    children.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
    children.push(...buildGrowthIndicators(cfg.growthIndicators, company));
  }

  // 표 바로 아래 여백 (표와 본문 섹션이 너무 붙어 보이지 않도록)
  children.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

  // External panel (right under overview)
  children.push(...buildPanel(company, cfg.panel));

  // Text sections
  for (const sec of (cfg.sections ?? [])) {
    children.push(sectionHeading(sec.heading));
    children.push(...bodyPara(val(company[sec.field], sec.emptyText ?? "-")));
  }

  // Two-column section
  children.push(...twoColumnSection(cfg.twoColumnSection, company));

  // Optional single text section (e.g. 기타)
  const optSec = cfg.optionalTextSection;
  if (optSec && !isEmpty(company[optSec.field])) {
    children.push(sectionHeading(optSec.heading));
    children.push(...bodyPara(val(company[optSec.field])));
  }

  // Link section (e.g. 사업계획서)
  const linkSec = cfg.linkSection;
  if (linkSec) {
    children.push(sectionHeading(linkSec.heading));
    const url = company[linkSec.field];
    children.push(new Paragraph({
      spacing: { after: 0 },
      children: isEmpty(url)
        ? [run("-", { size: 21, color: SUBTLE })]
        : [new ExternalHyperlink({ link: url, children: [run(linkSec.linkLabel ?? "파일 열기 ↗", { size: 21, bold: true, color: LINK_COLOR, underline: {} })] })],
    }));
  }

  // 데이터 표 (고용/재무 현황 등)
  // dataTablesCondField: 이 필드가 있는데 company[필드]가 falsy면 dataTables 전체를 스킵
  //   (예: innoclaw 매칭이 안 된 지원기업은 성장 데이터 자체가 없으므로 2페이지를 만들지 않고 1페이지로 끝냄)
  // dataTablesInline: true면 페이지브레이크 없이 1페이지 흐름에 바로 이어붙임 (공간이 빠듯할 때 compact:true 와 함께 사용)
  // 기본값(false): 항목이 많아 1페이지에 안 들어가는 경우를 위해 페이지를 나눠서 표시
  const hasGrowthData = !cfg.dataTablesCondField || !!company[cfg.dataTablesCondField];
  if (cfg.dataTables && cfg.dataTables.length && hasGrowthData) {
    if (!cfg.dataTablesInline) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      children.push(new Paragraph({
        spacing: { after: 200 },
        children: [run(val(company[cfg.nameField]), { bold: true, size: 30, color: "111315" }), run("  성장 데이터", { size: 21, color: SUBTLE })],
      }));
    }
    cfg.dataTables.forEach((t, i) => {
      children.push(...buildDataTable(t, company));
      if (i < cfg.dataTables.length - 1) children.push(new Paragraph({ spacing: { after: cfg.dataTablesInline ? 60 : 110 }, children: [] }));
    });
  }

  return new Document({
    creator: "Mark & Company",
    title: `${val(company[cfg.nameField])} - ${cfg.programName ?? "프로필"}`,
    styles: { default: { document: { run: { font: FONT, size: 21, color: BODY } } } },
    numbering: { config: [{ reference: "profile-bullets", levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 320, hanging: 220 } } },
    }]}]},
    sections: [{
      properties: { page: { size: { width: PAGE_WIDTH, height: PAGE_HEIGHT }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } },
      ...(cfg.header?.enabled === false ? {} : { headers: { default: new Header({ children: [new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [run(cfg.headerLeft ?? "", { size: 16, color: SUBTLE }), run(`\t${dateStr} 작성`, { size: 16, color: SUBTLE })],
      })]})} }),
      ...(cfg.footer?.enabled === false ? {} : { footers: { default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          run(`${cfg.footerText ?? ""}  — `, { size: 15, color: SUBTLE }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 15, color: SUBTLE }),
        ],
      })]})} }),
      children,
    }],
  });
}

// ─── Main ────────────────────────────────────────────────────────────────────
async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  for (const company of companies) {
    const doc = buildDoc(company);
    const buf = await Packer.toBuffer(doc);
    const companyName = String(company[cfg.fileNameField] ?? "unknown").replace(/[\\/:*?"<>|]/g, "_");
    const fname = `${cfg.outputPrefix ?? "프로필"}_${companyName}.docx`;
    const fpath = path.join(outDir, fname);
    fs.writeFileSync(fpath, buf);
    console.log("✅", fname);
  }
  console.log(`\n총 ${companies.length}개 파일 생성 완료 → ${outDir}`);
}

main().catch(e => { console.error("❌", e.message); process.exit(1); });

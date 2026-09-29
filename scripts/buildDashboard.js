/*
 * 매일 1회 실행되는 자동화의 메인 스크립트.
 *   1) dashboard.jsx 에서 최신 데이터 엔진(dataEngine.mjs)을 다시 뽑아냄
 *   2) 구글 시트(판매데이터 / 일일리포트)를 서비스 계정으로 불러와 xlsx로 내려받음
 *   3) 브라우저에서 "엑셀 업로드"할 때와 동일한 로직으로 기존 DEFAULT_DATA에 병합
 *   4) 병합된 데이터를 dashboard.jsx의 DEFAULT_DATA 자리에 다시 써 넣음
 *   5) esbuild로 번들을 다시 만들고 index.html을 재생성함
 *
 * 변경 사항이 전혀 없으면(=시트에 새로 반영된 달이 없으면) index.html을 다시 쓰지 않고 조용히 끝남
 * (깃허브 액션 쪽에서 "달라진 파일이 있을 때만 커밋"하도록 되어 있어 이중 안전장치)
 */
import fs from "node:fs";
import path from "node:path";
import { execSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.join(__dirname, "..");
const DASHBOARD_JSX = path.join(REPO_ROOT, "dashboard.jsx");
const ENGINE_MJS = path.join(__dirname, "dataEngine.mjs");
const ENTRY_JSX = path.join(__dirname, "..", "build", "entry.jsx");
const BUNDLE_JS = path.join(__dirname, "..", "build", "bundle.js");
const INDEX_HTML = path.join(REPO_ROOT, "index.html");

function log(msg) {
  console.log(`[buildDashboard] ${msg}`);
}

// 구글 드라이브 "내보내기(export)" API로 받은 xlsx에는 셀 "메모(노트)"가 빠져있어서(댓글과 달리
// 노트는 export 결과물에 포함되지 않음), 구글 시트 API로 직접 읽어온 노트 텍스트(notesGrid)를
// 워크북에 다시 "주입"해준 뒤 재직렬화함 — 이후 단계(processExcelFiles)는 브라우저에서 xlsx를
// 업로드했을 때와 완전히 동일한 코드로, 이 주입된 코멘트를 정상적으로 읽어들이게 됨.
function injectMemoNotesIntoWorkbook(buffer, notesGrid, MONTH_CODES, findMonthHeaderRow, log) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  let injectedCount = 0;
  const dbg = (m) => { if (log) log("    [주입 진단] " + m); };

  dbg(`엑셀 탭 목록: ${wb.SheetNames.join(", ")}`);
  dbg(`시트 API가 돌려준 탭 목록: ${Object.keys(notesGrid).join(", ")}`);

  wb.SheetNames.forEach((sheetName) => {
    const notesForSheet = notesGrid[sheetName];
    if (!notesForSheet) { dbg(`"${sheetName}": 시트 API 응답에 같은 이름의 탭이 없어서 건너뜀`); return; }
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws["!ref"]) { dbg(`"${sheetName}": 빈 시트`); return; }

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const hit = findMonthHeaderRow(rows);
    if (!hit) { dbg(`"${sheetName}": JAN~DEC 헤더 행을 못 찾음`); return; }
    const headerRow = rows[hit.r];
    const monthCols = {};
    let searchFrom = hit.c;
    MONTH_CODES.forEach((m) => {
      const idx = headerRow.findIndex((cell, ci) => ci >= searchFrom && cell != null && String(cell).trim().toUpperCase() === m);
      if (idx >= 0) { monthCols[m] = idx; searchFrom = idx + 1; }
    });
    const labelCol = Math.max(0, hit.c - 1);
    let customOrderRowIdx = -1;
    for (let r = hit.r + 1; r < rows.length; r++) {
      const cell = rows[r] ? rows[r][labelCol] : null;
      if (cell != null && String(cell).toLowerCase().includes("custom order")) { customOrderRowIdx = r; break; }
    }
    if (customOrderRowIdx < 0) { dbg(`"${sheetName}": "Custom Order" 라벨 행을 못 찾음 (헤더행=${hit.r}, 라벨열=${labelCol})`); return; }

    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rowOffset = range.s.r, colOffset = range.s.c;
    dbg(`"${sheetName}": 헤더행=${hit.r}, Custom Order행=${customOrderRowIdx}, range 시작(r,c)=(${rowOffset},${colOffset}), 시트API 행 수=${notesForSheet.length}`);

    MONTH_CODES.forEach((m) => {
      const col = monthCols[m];
      if (col == null) return;
      const absRow = customOrderRowIdx + rowOffset;
      const absCol = col + colOffset;
      const rowArr = notesForSheet[absRow];
      const note = rowArr && rowArr[absCol];
      if (m === "SEP" || m === "OCT") {
        dbg(`${m}: absRow=${absRow}, absCol=${absCol}, 해당 행 시트API 길이=${rowArr ? rowArr.length : "행 자체 없음"}, note=${note ? JSON.stringify(note.slice(0, 40)) : "없음"}`);
      }
      if (!note) return;
      const addr = XLSX.utils.encode_cell({ r: absRow, c: absCol });
      if (!ws[addr]) ws[addr] = { t: "s", v: "" };
      ws[addr].c = [{ a: "sdc-dashboard-bot", t: note }];
      injectedCount++;
    });
  });

  return { buffer: XLSX.write(wb, { type: "buffer", bookType: "xlsx" }), injectedCount };
}

async function main() {
  log("1/5 dataEngine.mjs 재생성 중...");
  execSync(`node "${path.join(__dirname, "extract-engine.js")}" "${DASHBOARD_JSX}" "${ENGINE_MJS}"`, { stdio: "inherit" });

  const engine = await import(`${ENGINE_MJS}?t=${Date.now()}`); // 캐시 무시
  const { DEFAULT_DATA, processExcelFiles, mergeMonthlySeries, mergeMonthlyByKeyMap, mergeVendorMonthlyDetail, extractYearFields, CHANNEL_KEYS, EXTRA_CHANNEL_COLORS, mergeRecordsByMonth, MONTH_CODES, findMonthHeaderRow } = engine;

  log("2/5 구글 시트에서 최신 데이터 내려받는 중...");
  const { fetchSheetsAsExcelFiles, fetchCellNotesGrid } = await import("./fetchSheets.js");
  const files = await fetchSheetsAsExcelFiles();

  log("2-B/5 Custom Order 메모(노트)를 구글 시트 API로 직접 불러와 반영 중...");
  try {
    const notesGrid = await fetchCellNotesGrid(process.env.SELLING_DATA_SHEET_ID);
    const sellingFile = files.find((f) => f.name === "selling-data.xlsx");
    if (sellingFile) {
      const { buffer, injectedCount } = injectMemoNotesIntoWorkbook(sellingFile.buffer, notesGrid, MONTH_CODES, findMonthHeaderRow, log);
      sellingFile.buffer = buffer;
      log(`  ✓ Custom Order 메모(노트) ${injectedCount}개 셀에서 확인되어 반영함`);
    }
  } catch (e) {
    log(`  ⚠️ Custom Order 메모(노트)를 구글 시트 API로 불러오는 중 오류 발생 (건너뜀): ${e.message}`);
  }

  log("3/5 엑셀 업로드와 동일한 로직으로 병합 중...");
  const existingCurrentYear = DEFAULT_DATA.currentYear || new Date().getFullYear();
  const { bundles, messages } = await processExcelFiles(files, existingCurrentYear);
  messages.forEach((m) => log("  " + m));

  if (bundles.length === 0) {
    log("인식할 수 있는 데이터가 없어 종료합니다 (시트 형식을 확인하세요).");
    return;
  }

  const detectedYears = bundles.map((b) => b.year).filter((y) => y != null);
  const newCurrentYear = detectedYears.length > 0 ? Math.max(existingCurrentYear, ...detectedYears) : existingCurrentYear;

  let nextTop = { ...DEFAULT_DATA };
  let nextPriorYears = { ...(DEFAULT_DATA.priorYears || {}) };
  let nextExtraChannels = [...(DEFAULT_DATA.extraChannels || [])];

  if (newCurrentYear !== existingCurrentYear) {
    nextPriorYears = { ...nextPriorYears, [existingCurrentYear]: extractYearFields(nextTop) };
  }

  bundles.forEach((b) => {
    const y = b.year || newCurrentYear;
    if (b.fields.newChannels) {
      b.fields.newChannels.forEach((nc) => {
        if (!nextExtraChannels.some((c) => c.key === nc.key)) {
          const color = EXTRA_CHANNEL_COLORS[nextExtraChannels.length % EXTRA_CHANNEL_COLORS.length];
          nextExtraChannels = [...nextExtraChannels, { key: nc.key, label: nc.label, color }];
        }
      });
    }
    if (y === newCurrentYear) {
      const mergedFields = { ...b.fields };
      delete mergedFields.newChannels;
      if (b.fields.channels) {
        const allKeys = new Set([...CHANNEL_KEYS, ...Object.keys(nextTop.channels || {}), ...Object.keys(b.fields.channels)]);
        const mergedChannels = {};
        allKeys.forEach((k) => {
          mergedChannels[k] = mergeMonthlySeries(nextTop.channels ? nextTop.channels[k] : null, b.fields.channels[k]);
        });
        mergedFields.channels = mergedChannels;
      }
      if (b.fields.vendorMonthly) mergedFields.vendorMonthly = mergeMonthlyByKeyMap(nextTop.vendorMonthly, b.fields.vendorMonthly);
      if (b.fields.vendorMonthlyDetail) mergedFields.vendorMonthlyDetail = mergeVendorMonthlyDetail(nextTop.vendorMonthlyDetail, b.fields.vendorMonthlyDetail);
      if (b.fields.itemVendorMap) mergedFields.itemVendorMap = { ...(nextTop.itemVendorMap || {}), ...b.fields.itemVendorMap };
      if (b.fields.posMonthlyQty) mergedFields.posMonthlyQty = mergeMonthlySeries(nextTop.posMonthlyQty, b.fields.posMonthlyQty);
      if (b.fields.posInvoiceCounts) mergedFields.posInvoiceCounts = mergeMonthlySeries(nextTop.posInvoiceCounts, b.fields.posInvoiceCounts);
      if (b.fields.visits) mergedFields.visits = mergeMonthlySeries(nextTop.visits, b.fields.visits);
      if (b.fields.contacts) mergedFields.contacts = mergeMonthlySeries(nextTop.contacts, b.fields.contacts);
      if (b.fields.sold) mergedFields.sold = mergeMonthlySeries(nextTop.sold, b.fields.sold);
      // customOrders/marketEvents: 메모 방식은 달마다 셀이 따로 있어서, 아직 새 형식으로 안 옮긴
      // 달은 이번 실행에 아예 안 잡힐 수 있음 -- 그런 달의 기존 데이터까지 통째로 지워지지 않도록
      // "이번에 새로 읽어온 달"만 교체하고 나머지 달은 그대로 유지함.
      if (b.fields.customOrders) mergedFields.customOrders = mergeRecordsByMonth(nextTop.customOrders, b.fields.customOrders);
      if (b.fields.marketEvents) mergedFields.marketEvents = mergeRecordsByMonth(nextTop.marketEvents, b.fields.marketEvents);
      // customLeads: 확정 주문과 달리 "지금 진행 중인 문의"를 나타내는 현재 상태 정보라, 탭이든
      // 메모든 이번에 읽어온 값으로 통째로 교체함 (메모 방식은 파서가 이미 가장 최근 달의 메모만
      // 반영해서 넘겨줌 -- 옛날 달의 리드가 계속 누적되어 남지 않도록).
      if (b.fields.customLeads) mergedFields.customLeads = b.fields.customLeads;
      nextTop = { ...nextTop, ...mergedFields };
    } else {
      const prev = nextPriorYears[y] || {};
      const mergedFields = { ...b.fields };
      delete mergedFields.newChannels;
      if (b.fields.channels) {
        const allKeys = new Set([...CHANNEL_KEYS, ...Object.keys(prev.channels || {}), ...Object.keys(b.fields.channels)]);
        const mergedChannels = {};
        allKeys.forEach((k) => {
          mergedChannels[k] = mergeMonthlySeries(prev.channels ? prev.channels[k] : null, b.fields.channels[k]);
        });
        mergedFields.channels = mergedChannels;
      }
      nextPriorYears = { ...nextPriorYears, [y]: { ...prev, ...mergedFields } };
    }
  });

  const next = {
    ...nextTop,
    currentYear: newCurrentYear,
    priorYears: nextPriorYears,
    extraChannels: nextExtraChannels,
    lastUploadedAt: new Date().toISOString(),
  };

  if (JSON.stringify(next) === JSON.stringify(DEFAULT_DATA)) {
    log("데이터에 변경 사항이 없습니다. index.html을 다시 만들지 않고 종료합니다.");
    return;
  }

  log("4/5 dashboard.jsx의 DEFAULT_DATA를 갱신하는 중...");
  writeNewDefaultData(DASHBOARD_JSX, next);

  log("5/5 esbuild로 번들 재생성 후 index.html 작성 중...");
  await rebuildIndexHtml();

  log("완료.");
}

function writeNewDefaultData(dashboardPath, nextData) {
  const src = fs.readFileSync(dashboardPath, "utf-8");
  const lines = src.split("\n");
  const startIdx = lines.findIndex((l) => /^const DEFAULT_DATA = \{/.test(l));
  if (startIdx === -1) throw new Error("dashboard.jsx 에서 'const DEFAULT_DATA = {' 를 찾지 못했습니다.");
  let endIdx = -1;
  for (let i = startIdx + 1; i < lines.length; i++) {
    if (/^\};\s*$/.test(lines[i])) { endIdx = i; break; }
  }
  if (endIdx === -1) throw new Error("dashboard.jsx 에서 DEFAULT_DATA 블록의 닫는 '};' 를 찾지 못했습니다.");

  const newBlock = "const DEFAULT_DATA = " + JSON.stringify(nextData, null, 2) + ";";
  const newLines = [...lines.slice(0, startIdx), newBlock, ...lines.slice(endIdx + 1)];
  fs.writeFileSync(dashboardPath, newLines.join("\n"), "utf-8");
}

async function rebuildIndexHtml() {
  const buildDir = path.join(REPO_ROOT, "build");
  fs.mkdirSync(buildDir, { recursive: true });
  fs.writeFileSync(
    ENTRY_JSX,
    `import React from "react";
import { createRoot } from "react-dom/client";
import Dashboard from "../dashboard.jsx";

const root = createRoot(document.getElementById("root"));
root.render(React.createElement(Dashboard));
`,
    "utf-8"
  );

  execSync(
    `npx esbuild "${ENTRY_JSX}" --bundle --format=iife --jsx=automatic --minify --outfile="${BUNDLE_JS}"`,
    { stdio: "inherit", cwd: REPO_ROOT }
  );

  const bundleCode = fs.readFileSync(BUNDLE_JS, "utf-8");
  const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>SDC Gift Shop Sales Dashboard</title>
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  #root { min-height: 100vh; }
</style>
</head>
<body>
<div id="root"></div>
<script>
${bundleCode}
</script>
</body>
</html>
`;
  fs.writeFileSync(INDEX_HTML, html, "utf-8");
}

main().catch((e) => {
  console.error("[buildDashboard] 실패:", e);
  process.exit(1);
});

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

// 구글 드라이브 "내보내기(export)"로 받은 xlsx는 셀 "메모(노트)"가 아예 빠져 있을 뿐 아니라, 빈
// 선행 행/열을 잘라내는 등 실제 라이브 시트와 좌표(행/열 번호)가 어긋날 수 있다는 것까지
// 확인됨 (xlsx 쪽에서 찾은 위치를 시트 API 응답에 그대로 대입했더니 전혀 다른 셀을 가리켰음).
// 그래서 이제 "어느 행이 헤더/Custom Order 행인지" 탐지 자체를 시트 API가 돌려준 그리드
// 안에서 처음부터 다시 함 — values/notes가 완전히 같은 좌표계라 어긋날 일이 없음.
function parseCustomOrderMemoFromSheetsGrid(grid, MONTH_CODES, findMonthHeaderRow, parseMemoOrderLine, MEMO_LEAD_PREFIX_RE, log) {
  const dbg = (m) => { if (log) log("    [시트API 기반 탐지] " + m); };
  let best = null;

  Object.keys(grid).forEach((sheetName) => {
    const { values, notes } = grid[sheetName];
    if (!values || values.length === 0) return;
    const hit = findMonthHeaderRow(values);
    if (!hit) return;
    const headerRow = values[hit.r];
    const monthCols = {};
    let searchFrom = hit.c;
    MONTH_CODES.forEach((m) => {
      const idx = headerRow.findIndex((cell, ci) => ci >= searchFrom && cell != null && String(cell).trim().toUpperCase() === m);
      if (idx >= 0) { monthCols[m] = idx; searchFrom = idx + 1; }
    });
    const labelCol = Math.max(0, hit.c - 1);
    let customOrderRowIdx = -1;
    for (let r = hit.r + 1; r < values.length; r++) {
      const cell = values[r] ? values[r][labelCol] : null;
      if (cell != null && String(cell).toLowerCase().includes("custom order")) { customOrderRowIdx = r; break; }
    }
    if (customOrderRowIdx < 0) { dbg(`"${sheetName}": 헤더행=${hit.r}은 찾았지만 Custom Order 라벨 행을 못 찾음`); return; }
    dbg(`"${sheetName}": 헤더행=${hit.r}, Custom Order행=${customOrderRowIdx} — 이 시트를 대상으로 사용`);

    const orders = [];
    const leads = [];
    let unparsedCount = 0;
    let orderIdCounter = 1;
    let lastOrderMonth = -1;
    let anyRecognized = false;
    const monthDebug = [];

    MONTH_CODES.forEach((m, month) => {
      const col = monthCols[m];
      if (col == null) return;
      const note = notes[customOrderRowIdx] && notes[customOrderRowIdx][col];
      if (!note) { monthDebug.push(`${m}:메모없음`); return; }
      note.split("\n").forEach((line) => {
        const trimmed = line.trim();
        if (!trimmed) return;
        if (MEMO_LEAD_PREFIX_RE.test(trimmed)) { anyRecognized = true; return; }
        if (parseMemoOrderLine(trimmed) != null) { anyRecognized = true; lastOrderMonth = month; }
      });
    });
    const leadStartMonth = Math.max(lastOrderMonth, 0);

    MONTH_CODES.forEach((m, month) => {
      const col = monthCols[m];
      if (col == null) return;
      const note = notes[customOrderRowIdx] && notes[customOrderRowIdx][col];
      if (!note) return;
      const lines = note.split("\n").map((l) => l.trim()).filter((l) => l.length > 0);
      let recognizedThisMonth = 0, leadsThisMonth = 0, unparsedBefore = unparsedCount;
      const groups = new Map();
      lines.forEach((line) => {
        if (MEMO_LEAD_PREFIX_RE.test(line)) {
          leadsThisMonth++;
          if (month >= leadStartMonth) leads.push(line.replace(MEMO_LEAD_PREFIX_RE, "").trim());
          return;
        }
        const parsed = parseMemoOrderLine(line);
        if (!parsed) { unparsedCount++; return; }
        recognizedThisMonth++;
        const key = parsed.customer.toLowerCase();
        if (!groups.has(key)) {
          groups.set(key, {
            id: orderIdCounter++, groupId: `${month}-${key}`, month,
            customer: parsed.customer, category: parsed.category, product: parsed.product,
            qty: parsed.qty, amount: parsed.amount, note: "",
          });
        } else {
          const existing = groups.get(key);
          if (parsed.product) existing.product = existing.product ? `${existing.product}, ${parsed.product}` : parsed.product;
          if (parsed.qty != null) existing.qty = (existing.qty || 0) + parsed.qty;
          existing.amount += parsed.amount;
          if (!existing.category && parsed.category) existing.category = parsed.category;
        }
      });
      groups.forEach((o) => orders.push(o));
      monthDebug.push(`${m}:${note.length}자/${lines.length}줄(인식${recognizedThisMonth},리드${leadsThisMonth},미인식${unparsedCount - unparsedBefore})`);
    });

    const result = { orders, leads: anyRecognized ? leads : null, unparsedCount, debug: monthDebug };
    // Custom Order 행을 찾은 첫 시트를 사용 (판매 데이터 시트는 보통 하나뿐)
    if (!best) best = result;
  });

  return best;
}

async function main() {
  log("1/5 dataEngine.mjs 재생성 중...");
  execSync(`node "${path.join(__dirname, "extract-engine.js")}" "${DASHBOARD_JSX}" "${ENGINE_MJS}"`, { stdio: "inherit" });

  const engine = await import(`${ENGINE_MJS}?t=${Date.now()}`); // 캐시 무시
  const { DEFAULT_DATA, processExcelFiles, mergeMonthlySeries, mergeMonthlyByKeyMap, mergeVendorMonthlyDetail, extractYearFields, CHANNEL_KEYS, EXTRA_CHANNEL_COLORS, mergeRecordsByMonth, MONTH_CODES, findMonthHeaderRow, parseMemoOrderLine, MEMO_LEAD_PREFIX_RE } = engine;

  log("2/5 구글 시트에서 최신 데이터 내려받는 중...");
  const { fetchSheetsAsExcelFiles, fetchCustomOrderMemoGrid } = await import("./fetchSheets.js");
  const files = await fetchSheetsAsExcelFiles();

  log("2-B/5 Custom Order 메모(노트)를 구글 시트 API로 직접 불러와 파싱 중...");
  let sheetsApiMemoResult = null;
  try {
    const grid = await fetchCustomOrderMemoGrid(process.env.SELLING_DATA_SHEET_ID);
    sheetsApiMemoResult = parseCustomOrderMemoFromSheetsGrid(grid, MONTH_CODES, findMonthHeaderRow, parseMemoOrderLine, MEMO_LEAD_PREFIX_RE, log);
    if (sheetsApiMemoResult) {
      log(`  🔍 Custom Order memo 점검(시트API): ${sheetsApiMemoResult.debug.join(" / ")}`);
      if (sheetsApiMemoResult.orders.length > 0) log(`  ✓ Custom Order memo(s) recognized via Sheets API (${sheetsApiMemoResult.orders.length} order line(s))`);
      if (sheetsApiMemoResult.leads != null) log(`  ✓ Custom Order lead memo(s) via Sheets API: ${sheetsApiMemoResult.leads.length}건`);
      if (sheetsApiMemoResult.unparsedCount > 0) log(`  ℹ️ ${sheetsApiMemoResult.unparsedCount}줄 형식 불일치로 건너뜀`);
    } else {
      log("  ⚠️ 시트API 응답에서 Custom Order 행을 가진 시트를 찾지 못함");
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

  // 판매 데이터 번들에 Custom Orders/Leads 탭이 없어서 아직 못 채워졌다면, 시트 API로 직접
  // 파싱한 메모 결과로 채워넣음 (탭이 있으면 탭이 우선이라 이미 채워져 있으므로 덮어쓰지 않음).
  if (sheetsApiMemoResult) {
    const salesBundle = bundles.find((b) => b.type === "sales");
    if (salesBundle) {
      if (!salesBundle.fields.customOrders && sheetsApiMemoResult.orders.length > 0) {
        salesBundle.fields.customOrders = sheetsApiMemoResult.orders;
      }
      if (!salesBundle.fields.customLeads && sheetsApiMemoResult.leads != null) {
        salesBundle.fields.customLeads = sheetsApiMemoResult.leads;
      }
    }
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

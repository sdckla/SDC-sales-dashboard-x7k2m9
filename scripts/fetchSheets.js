/* Google Sheets(비공개) → 엑셀(xlsx) 버퍼로 내려받기
 * 서비스 계정(Service Account) 인증을 사용합니다. 시트를 공개로 바꿀 필요가 없습니다.
 * 필요 환경변수:
 *   GCP_SERVICE_ACCOUNT_KEY   - 서비스 계정 JSON 키 전체 내용 (문자열)
 *   SELLING_DATA_SHEET_ID     - "판매 데이터(Monthly Sales/Details/POS Data)" 구글시트 ID
 *   DAILY_REPORT_SHEET_ID     - "일일 리포트(Daily Report)" 구글시트 ID
 */
import { google } from "googleapis";

function getAuth() {
  const raw = process.env.GCP_SERVICE_ACCOUNT_KEY;
  if (!raw) throw new Error("환경변수 GCP_SERVICE_ACCOUNT_KEY 가 설정되지 않았습니다.");
  let credentials;
  try {
    credentials = JSON.parse(raw);
  } catch (e) {
    throw new Error("GCP_SERVICE_ACCOUNT_KEY 값이 올바른 JSON이 아닙니다: " + e.message);
  }
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/spreadsheets.readonly",
    ],
  });
}

// 구글시트를 xlsx 포맷으로 export 해서 Buffer로 받아옴 (Drive API 사용 — 서비스 계정이
// 그 시트에 "뷰어" 이상 권한으로 공유되어 있어야 함)
async function exportSheetAsXlsxBuffer(drive, spreadsheetId, label) {
  if (!spreadsheetId) throw new Error(`${label} 시트 ID가 설정되지 않았습니다.`);
  const res = await drive.files.export(
    {
      fileId: spreadsheetId,
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(res.data);
}

export async function fetchSheetsAsExcelFiles() {
  const auth = getAuth();
  const drive = google.drive({ version: "v3", auth });

  const sellingBuf = await exportSheetAsXlsxBuffer(drive, process.env.SELLING_DATA_SHEET_ID, "판매 데이터(SELLING_DATA_SHEET_ID)");
  const dailyBuf = await exportSheetAsXlsxBuffer(drive, process.env.DAILY_REPORT_SHEET_ID, "일일 리포트(DAILY_REPORT_SHEET_ID)");

  // processExcelFiles()가 기대하는 형태: { name, buffer }
  return [
    { name: "selling-data.xlsx", buffer: sellingBuf },
    { name: "daily-report.xlsx", buffer: dailyBuf },
  ];
}

// 구글 드라이브 API로 시트를 xlsx로 "내보내기(export)" 하면 셀에 달린 "메모(노트)"가
// 통째로 빠져버림 (댓글/스레드댓글과 달리 노트는 export 결과물에 포함되지 않음). Custom Order
// 메모 자동 인식 기능을 쓰려면 노트 내용이 반드시 필요하므로, 구글 시트 API(v4)로 별도로
// 직접 읽어옴 — Sheets API는 노트를 정상적으로 돌려주고, JSON이라 인코딩 문제도 없음.
// 반환값: { [시트탭이름]: string[][] } — 각 셀의 note 텍스트(없으면 null), row/col은 0부터 시작하는
// 절대 좌표(A1 기준)라서 xlsx로 읽은 워크북의 range.s.r/range.s.c 오프셋을 더한 주소와 그대로 대응됨.
export async function fetchCellNotesGrid(spreadsheetId, log) {
  const dbg = (m) => { if (log) log("    [시트API 원본 진단] " + m); };
  const auth = getAuth();
  const sheetsApi = google.sheets({ version: "v4", auth });
  // fields 마스크 없이, includeGridData만으로 전체를 받아옴 (마스크 문법 문제로 데이터가
  // 잘리는 경우를 배제하기 위해 — 응답 크기가 커도 시트 하나 정도는 문제없음).
  const res = await sheetsApi.spreadsheets.get({
    spreadsheetId,
    includeGridData: true,
  });
  const sheets = res.data.sheets || [];
  dbg(`응답에 포함된 시트 수: ${sheets.length}`);
  const grid = {};
  sheets.forEach((sheet) => {
    const title = sheet.properties && sheet.properties.title;
    if (!title) return;
    const dataArr = sheet.data || [];
    const rowData = (dataArr[0] && dataArr[0].rowData) || [];
    dbg(`"${title}": sheet.data 배열 길이=${dataArr.length}, rowData 길이=${rowData.length}`);
    if (title === "Monthly Sales" && rowData[5]) {
      const sampleCell = (rowData[5].values || [])[10];
      dbg(`"${title}" 6번째 행(0-index 5) 11번째 칸(0-index 10) 원본: ${JSON.stringify(sampleCell)}`);
    }
    grid[title] = rowData.map((row) => (row.values || []).map((cell) => (cell && cell.note) || null));
  });
  return grid;
}

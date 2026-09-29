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
// 반환값: { [시트탭이름]: { values: (string|null)[][], notes: (string|null)[][] } }
// values/notes는 같은 구글 시트 API 응답에서 같은 인덱스로 뽑아내므로 행/열 번호가 서로 100%
// 일치함 (반면 구글 드라이브로 "내보낸" xlsx 파일은 빈 선행 행/열을 잘라내는 등 실제 시트와
// 좌표가 어긋날 수 있어서, 그 xlsx에서 찾은 위치를 이 API 응답에 그대로 대입하면 엉뚱한 셀을
// 가리키는 문제가 있었음 — 그래서 이제 "어느 행이 Custom Order 행인지"도 이 API 응답 자체에서
// 직접 다시 찾음).
export async function fetchCustomOrderMemoGrid(spreadsheetId) {
  const auth = getAuth();
  const sheetsApi = google.sheets({ version: "v4", auth });
  const res = await sheetsApi.spreadsheets.get({
    spreadsheetId,
    includeGridData: true,
  });
  const sheets = res.data.sheets || [];
  const grid = {};
  sheets.forEach((sheet) => {
    const title = sheet.properties && sheet.properties.title;
    if (!title) return;
    const dataArr = sheet.data || [];
    const rowData = (dataArr[0] && dataArr[0].rowData) || [];
    const values = rowData.map((row) => (row.values || []).map((cell) => (cell && cell.formattedValue) || null));
    const notes = rowData.map((row) => (row.values || []).map((cell) => (cell && cell.note) || null));
    grid[title] = { values, notes };
  });
  return grid;
}

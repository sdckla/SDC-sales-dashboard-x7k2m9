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
    scopes: ["https://www.googleapis.com/auth/drive.readonly"],
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

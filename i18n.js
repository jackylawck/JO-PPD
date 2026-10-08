// i18n.js - Jumbo Orient Punch Problem Detector 國際化語言包
const i18nData = {
  zh: {
    title: "東淦打卡異常偵測工具",
    subtitle: "Jumbo Orient Punch Problem Detector (JO-PPD)",
    privacyBadge: "純前端離線解析・零數據外流風險",
    langBtn: "English",

    // 指引與說明區塊
    guideTitle: "系統說明與操作指引",
    guideIntro: "本工具為東淦工程有限公司（Jumbo Orient Contracting Limited）自主研發之考勤完整性審計引擎。專門用於每月審計考勤原始記錄（Attendance Record），在瀏覽器本地記憶體內自動識別未閉環、漏打卡、打卡脫節等異常，秒級產出標準化 Problem Log 報表。",
    guideStepTitle: "操作步驟：",
    guideSteps: [
      "從考勤系統導出月度考勤原始報表（.xls 或 .xlsx 格式）。",
      "將檔案拖曳至下方上傳區，系統即刻在瀏覽器本地執行向量化審計。",
      "檢視異常數據預覽表格，點擊「下載 Problem.xlsx 報表」匯出標準化日誌。"
    ],
    guideRulesTitle: "核心審計邏輯：",
    guideRules: [
      "審計範圍：嚴格審計月薪職員（E 開頭工號），自動排除地盤工友（W 開頭工號）。",
      "週期截數：自動套用每月 25 號截數日，避免跨期未平帳數據干擾。",
      "缺勤衝突：實際有打卡紀錄（Clk > 0）卻被系統標記為缺勤（ABS）。",
      "漏打下班：上班有打卡但下班無卡且離場時間為 0（Time Out = 0）。",
      "外勤真空：特殊狀態標記（如 TRAIN / SITE）但實際打卡完全空白。",
      "打卡脫節：打卡時間與系統認可進出場時間脫節大於等於 30 分鐘。",
      "工時流失：無請假登記下，早退與遲到累計工時損失大於等於 1.5 小時。"
    ],

    // 上傳區域
    dropTitle: "點擊選擇或拖曳考勤 Excel 檔到此處",
    dropSubtitle: "支援原廠系統導出之 .xls 及 .xlsx 檔案",
    dropNotice: "資安承諾：所有數據運算均在您的本機瀏覽器內完成，絕不傳輸至任何外部伺服器。",

    // 狀態與按鈕
    processing: "正在解析檔案並執行本地完整性審計...",
    complete: "審計完成！共篩選出 {count} 筆打卡異常記錄。",
    downloadBtn: "下載 Problem.xlsx 報表",
    noData: "未偵測到符合條件的異常打卡記錄。",

    // 預覽表格標頭
    colDept: "部門代碼",
    colDeptName: "部門名稱",
    colEmpCode: "工號",
    colName: "姓名",
    colDate: "日期",
    colDay: "星期",
    colClocking: "原始打卡 (Clk1-4)",
    colIn: "Time IN",
    colOut: "Time OUT"
  },

  en: {
    title: "JO Punch Problem Detector",
    subtitle: "Jumbo Orient Punch Problem Detector (JO-PPD)",
    privacyBadge: "Client-Side Offline Engine • Zero Data Leakage Risk",
    langBtn: "繁體中文",

    // Guide Section
    guideTitle: "Instructions & Operational Guide",
    guideIntro: "Developed independently by Jumbo Orient Contracting Limited, this system is an enterprise-grade punch integrity audit engine. It inspects raw monthly attendance records directly within browser memory, identifying unclosed punches and roster discrepancies to export standardized Problem Log workbooks instantly.",
    guideStepTitle: "Operational Steps:",
    guideSteps: [
      "Export the monthly raw attendance file (.xls or .xlsx) from the attendance software.",
      "Drag and drop the file into the upload zone below for instant local evaluation.",
      "Review the parsed anomalies and click 'Download Problem.xlsx' to save the report."
    ],
    guideRulesTitle: "Core Audit Logic:",
    guideRules: [
      "Staff Scope: Strictly audits salaried staff (E-code), excluding site workers (W-code).",
      "Cut-off Cycle: Enforces the 25th cut-off cycle to prevent cross-period spillover.",
      "Conflict Absence: Punch recorded (Clk > 0) but status marked as ABS.",
      "Missed Checkout: Valid check-in punch but no check-out punch (Time Out = 0).",
      "Unpunched Duty: Official duty marked (e.g. TRAIN / SITE) but punches are completely blank.",
      "Time Discrepancy: Discrepancy between actual punch and system record >= 30 mins.",
      "Severe Lost Time: Unapproved cumulative early leave or lateness >= 1.5 hours."
    ],

    // Upload Zone
    dropTitle: "Click to select or drag attendance Excel file here",
    dropSubtitle: "Supports standard system exported .xls and .xlsx files",
    dropNotice: "Security Assurance: Data is processed strictly within local memory. Zero data is transmitted externally.",

    // Status & Actions
    processing: "Parsing workbook and executing local audit rules...",
    complete: "Audit complete! Identified {count} punch anomaly records.",
    downloadBtn: "Download Problem.xlsx Report",
    noData: "No punch anomalies detected.",

    // Table Headers
    colDept: "Dept Code",
    colDeptName: "Dept Name",
    colEmpCode: "Emp Code",
    colName: "Name",
    colDate: "Date",
    colDay: "Day",
    colClocking: "Actual Clocking (Clk1-4)",
    colIn: "Time IN",
    colOut: "Time OUT"
  }
};

let currentLang = 'zh';

function setLanguage(lang) {
  currentLang = lang;
  const t = i18nData[lang];

  document.getElementById('ui-title').innerText = t.title;
  document.getElementById('ui-subtitle').innerText = t.subtitle;
  document.getElementById('ui-privacy-badge').innerText = t.privacyBadge;
  document.getElementById('ui-lang-btn').innerText = t.langBtn;

  document.getElementById('ui-guide-title').innerText = t.guideTitle;
  document.getElementById('ui-guide-intro').innerText = t.guideIntro;
  document.getElementById('ui-guide-step-title').innerText = t.guideStepTitle;
  document.getElementById('ui-guide-rules-title').innerText = t.guideRulesTitle;

  const stepList = document.getElementById('ui-guide-steps');
  stepList.innerHTML = t.guideSteps.map(s => `<li>${s}</li>`).join('');

  const rulesList = document.getElementById('ui-guide-rules');
  rulesList.innerHTML = t.guideRules.map(r => `<li>${r}</li>`).join('');

  document.getElementById('ui-drop-title').innerText = t.dropTitle;
  document.getElementById('ui-drop-subtitle').innerText = t.dropSubtitle;
  document.getElementById('ui-drop-notice').innerText = t.dropNotice;
  document.getElementById('download-btn').innerText = t.downloadBtn;

  document.getElementById('th-dept').innerText = t.colDept;
  document.getElementById('th-dept-name').innerText = t.colDeptName;
  document.getElementById('th-emp').innerText = t.colEmpCode;
  document.getElementById('th-name').innerText = t.colName;
  document.getElementById('th-date').innerText = t.colDate;
  document.getElementById('th-day').innerText = t.colDay;
  document.getElementById('th-clk').innerText = t.colClocking;
  document.getElementById('th-in').innerText = t.colIn;
  document.getElementById('th-out').innerText = t.colOut;

  if (window.lastProblemCount !== undefined) {
    document.getElementById('status-text').innerText = t.complete.replace('{count}', window.lastProblemCount);
  }
}

function toggleLanguage() {
  setLanguage(currentLang === 'zh' ? 'en' : 'zh');
}

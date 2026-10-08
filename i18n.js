// i18n.js - Jumbo Orient Punch Problem Detector 國際化語言包
const i18nData = {
  zh: {
    title: "東淦打卡異常偵測工具",
    subtitle: "Jumbo Orient Punch Problem Detector (JO-PPD)",
    privacyBadge: "純前端離線解析・零數據外流風險",
    langBtn: "English",

    guideTitle: "系統說明與操作指引",
    guideIntro: "本工具為東淦工程有限公司（Jumbo Orient Contracting Limited）自主研發之考勤異常快篩引擎。系統自動將「全天缺勤 (ABS)」與「漏打收工卡」列為最高優先級置頂高亮，並提供一鍵分類過濾。",
    guideStepTitle: "操作步驟：",
    guideSteps: [
      "從考勤系統導出月度考勤原始報表（.xls 或 .xlsx 格式）。",
      "將檔案拖曳至下方上傳區，系統即刻在瀏覽器本地秒級完成審計。",
      "系統預設將「缺勤」與「漏打卡」置頂顯示，亦可點擊上方標籤快速切換分類，點擊按鈕匯出標準 Problem.xlsx 報表。"
    ],
    guideRulesTitle: "公司業務審計基準：",
    guideRules: [
      "🚨 最高優先：全天缺勤（系統記為 ABS）或有打卡卻被標記為 ABS。",
      "⚠️ 次高優先：漏打卡（上班有打卡，但無收工打卡記錄）。",
      "外勤無打卡：標記為 TRAIN / SITE / MEET 等特殊任務但全日打卡空白。",
      "遲到判定：首度上班打卡遲過 9:00（Clk1 > 9.00）。",
      "早退判定：星期一至五，收工打卡早過 17:30（Last_Clk < 17.30 且非半日假）。"
    ],

    dropTitle: "點擊選擇或拖曳考勤 Excel 檔到此處",
    dropSubtitle: "支援原廠系統導出之 .xls 及 .xlsx 檔案",
    dropNotice: "資安承諾：所有數據運算均在您的本機瀏覽器內完成，絕不傳輸至任何外部伺服器。",

    processing: "正在解析檔案並執行本地業務規則審計...",
    complete: "審計完成！共篩選出 {count} 筆打卡異常記錄（缺勤與漏打卡已置頂）。",
    downloadBtn: "下載 Problem.xlsx 報表",
    noData: "太棒了！未偵測到任何異常打卡記錄。",

    filterAll: "全部異常",
    filterAbs: "🚨 全天缺勤 (ABS)",
    filterMissed: "⚠️ 漏打卡 (Missed Checkout)",
    filterLate: "遲過 9:00",
    filterEarly: "早過 17:30",

    colDept: "部門代碼",
    colDeptName: "部門名稱",
    colEmpCode: "工號",
    colName: "姓名",
    colDate: "日期",
    colDay: "星期",
    colClocking: "原始打卡 (Clk1-4)",
    colIn: "Time IN",
    colOut: "Time OUT",
    colReason: "異常原因 (Reason)"
  },

  en: {
    title: "JO Punch Problem Detector",
    subtitle: "Jumbo Orient Punch Problem Detector (JO-PPD)",
    privacyBadge: "Client-Side Offline Engine • Zero Data Leakage Risk",
    langBtn: "繁體中文",

    guideTitle: "Instructions & Operational Guide",
    guideIntro: "Developed independently by Jumbo Orient Contracting Limited, this system audits attendance records with automated priority ranking—highlighting Absent (ABS) and Missed Checkout at the top.",
    guideStepTitle: "Operational Steps:",
    guideSteps: [
      "Export the monthly raw attendance file (.xls or .xlsx) from the attendance software.",
      "Drag and drop the file into the upload zone below for instant local evaluation.",
      "Severe anomalies (ABS and Missed Checkout) are highlighted at the top. Use category tabs to filter or click Download to save the Problem.xlsx report."
    ],
    guideRulesTitle: "Corporate Audit Benchmarks:",
    guideRules: [
      "🚨 High Priority: Absence (marked as ABS) or punch recorded but marked as ABS.",
      "⚠️ High Priority: Missed Checkout (checked in but missing checkout punch).",
      "Unpunched Duty: Official duties (TRAIN / SITE / MEET) with zero punches.",
      "Late Check-in: First check-in punch after 09:00 AM (Clk1 > 9.00).",
      "Early Leave: Checked out before 17:30 on weekdays without approved leave."
    ],

    dropTitle: "Click to select or drag attendance Excel file here",
    dropSubtitle: "Supports standard system exported .xls and .xlsx files",
    dropNotice: "Security Assurance: Data is processed strictly within local memory. Zero data is transmitted externally.",

    processing: "Parsing workbook and executing corporate audit rules...",
    complete: "Audit complete! Identified {count} punch anomalies (ABS & Missed Punches pinned to top).",
    downloadBtn: "Download Problem.xlsx Report",
    noData: "Great! No punch anomalies detected.",

    filterAll: "All Anomalies",
    filterAbs: "🚨 Absence (ABS)",
    filterMissed: "⚠️ Missed Checkout",
    filterLate: "Late > 9:00",
    filterEarly: "Early < 17:30",

    colDept: "Dept Code",
    colDeptName: "Dept Name",
    colEmpCode: "Emp Code",
    colName: "Name",
    colDate: "Date",
    colDay: "Day",
    colClocking: "Actual Clocking (Clk1-4)",
    colIn: "Time IN",
    colOut: "Time OUT",
    colReason: "Reason"
  }
};

let currentLang = 'zh';

function setSafeText(id, text) {
  const el = document.getElementById(id);
  if (el) el.innerText = text;
}

function setLanguage(lang) {
  currentLang = lang;
  const t = i18nData[lang];

  setSafeText('ui-title', t.title);
  setSafeText('ui-subtitle', t.subtitle);
  setSafeText('ui-privacy-badge', t.privacyBadge);
  setSafeText('ui-lang-btn', t.langBtn);

  setSafeText('ui-guide-title', t.guideTitle);
  setSafeText('ui-guide-intro', t.guideIntro);
  setSafeText('ui-guide-step-title', t.guideStepTitle);
  setSafeText('ui-guide-rules-title', t.guideRulesTitle);

  const stepList = document.getElementById('ui-guide-steps');
  if (stepList) stepList.innerHTML = t.guideSteps.map(s => `<li>${s}</li>`).join('');

  const rulesList = document.getElementById('ui-guide-rules');
  if (rulesList) rulesList.innerHTML = t.guideRules.map(r => `<li>${r}</li>`).join('');

  setSafeText('ui-drop-title', t.dropTitle);
  setSafeText('ui-drop-subtitle', t.dropSubtitle);
  setSafeText('ui-drop-notice', t.dropNotice);
  setSafeText('download-btn', t.downloadBtn);

  setSafeText('tab-all', t.filterAll);
  setSafeText('tab-abs', t.filterAbs);
  setSafeText('tab-missed', t.filterMissed);
  setSafeText('tab-late', t.filterLate);
  setSafeText('tab-early', t.filterEarly);

  setSafeText('th-dept', t.colDept);
  setSafeText('th-dept-name', t.colDeptName);
  setSafeText('th-emp', t.colEmpCode);
  setSafeText('th-name', t.colName);
  setSafeText('th-date', t.colDate);
  setSafeText('th-day', t.colDay);
  setSafeText('th-clk', t.colClocking);
  setSafeText('th-in', t.colIn);
  setSafeText('th-out', t.colOut);
  setSafeText('th-reason', t.colReason);

  if (window.lastProblemCount !== undefined) {
    setSafeText('status-text', t.complete.replace('{count}', window.lastProblemCount));
  }
}

function toggleLanguage() {
  setLanguage(currentLang === 'zh' ? 'en' : 'zh');
}

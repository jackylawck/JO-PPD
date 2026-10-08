// i18n.js - Jumbo Orient Punch Problem Detector 國際化語言包
const i18nData = {
  zh: {
    title: "東淦打卡異常偵測工具",
    subtitle: "Jumbo Orient Punch Problem Detector (JO-PPD)",
    privacyBadge: "純前端離線解析・零數據外流風險",
    langBtn: "English",

    // 指引與說明區塊
    guideTitle: "系統說明與操作指引",
    guideIntro: "本工具為東淦工程有限公司（Jumbo Orient Contracting Limited）自主研發之考勤異常快篩引擎。以公司實際排班彈性（最遲 9:00 返工、最早 17:30 收工）為基準，在瀏覽器本地自動標記打卡異常與原因，產出標準化 Problem Log 報表。",
    guideStepTitle: "操作步驟：",
    guideSteps: [
      "從考勤系統導出月度考勤原始報表（.xls 或 .xlsx 格式）。",
      "將檔案拖曳至下方上傳區，系統即刻在瀏覽器本地秒級完成審計。",
      "系統預設將「漏打卡」與「缺勤」置頂顯示，亦可點擊上方標籤快速切換分類，點擊按鈕匯出標準 Problem.xlsx 報表。"
    ],
    guideRulesTitle: "公司業務審計基準：",
    guideRules: [
      "⚠️ 最高優先：漏打卡（整天剛好只有一次打卡記錄，且非正當外勤）。",
      "🚨 次高優先：全天缺勤（整天無打卡且系統標記為 ABS）。",
      "正當外勤豁免：SITE（出地盤）、MEET（開會）、TRAIN（培訓）等經審批活動全面放行，不視為異常。",
      "遲到判定：首度上班打卡遲過 9:00（Clk1 > 9.00）。",
      "早退判定：星期一至五，收工打卡早過 17:30（Last_Clk < 17.30 且非半日假）。"
    ],

    // 🛡️ 企業級法規與國際標準合規宣示 (Governance & Compliance)
    complianceTitle: "🏛️ 企業級數據私隱、資安與科技治理承諾 (Compliance & Governance)",
    complianceIntro: "本系統架構貫徹 Privacy by Design（從設計著手保護私隱）原則，嚴格符合中外法規標準與國際認證體系，保障員工考勤數據之最高私隱安全：",
    complianceBadges: [
      { name: "香港 Cap. 486 PDPO", desc: "完全符合香港《個人資料（私隱）條例》，零未授權數據轉移" },
      { name: "歐盟 GDPR / EU AI Act", desc: "純確定性可解釋算法，無黑箱偏見；完全杜絕跨國境數據外洩" },
      { name: "國家網信辦法規", desc: "落實《個人信息保護法》(PIPL) 與數據安全規範，本地封閉記憶體運算" },
      { name: "ISO/IEC 27001 & 27701", desc: "遵循國際資安與隱私資訊管理標準，不留存任何磁碟日誌或 Cookie" },
      { name: "ISO/IEC 42001 AIMS", desc: "符合負責任人工智慧治理體系規範，規則透明、結果具 100% 可追溯性" }
    ],

    // 上傳區域
    dropTitle: "點擊選擇或拖曳考勤 Excel 檔到此處",
    dropSubtitle: "支援原廠系統導出之 .xls 及 .xlsx 檔案",
    dropNotice: "資安承諾：所有數據運算均在您的本機瀏覽器內完成，絕不傳輸至任何外部伺服器。",

    // 狀態與按鈕
    processing: "正在解析檔案並執行本地業務規則審計...",
    complete: "審計完成！共篩選出 {count} 筆打卡異常記錄（漏打卡與缺勤已置頂）。",
    downloadBtn: "下載 Problem.xlsx 報表",
    noData: "太棒了！未偵測到任何異常打卡記錄。",

    // 快速過濾標籤
    filterAll: "全部異常",
    filterMissed: "⚠️ 漏打卡 (Missed Checkout)",
    filterAbs: "🚨 全天缺勤 (ABS)",
    filterLate: "遲過 9:00",
    filterEarly: "早過 17:30",

    // 預覽表格標頭
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

    // Guide Section
    guideTitle: "Instructions & Operational Guide",
    guideIntro: "Developed independently by Jumbo Orient Contracting Limited, this system audits attendance records against corporate working-hour standards (latest check-in 09:00, earliest check-out 17:30). Processing executes locally within browser memory.",
    guideStepTitle: "Operational Steps:",
    guideSteps: [
      "Export the monthly raw attendance file (.xls or .xlsx) from the attendance software.",
      "Drag and drop the file into the upload zone below for instant local evaluation.",
      "Missed checkouts and unexcused absences are pinned to the top. Filter by category or export the standardized Problem.xlsx."
    ],
    guideRulesTitle: "Corporate Audit Benchmarks:",
    guideRules: [
      "⚠️ Top Priority: Missed Checkout (exactly one punch recorded without approved out-of-office duty).",
      "🚨 High Priority: Full-day Absence (zero punches and recorded as ABS).",
      "Approved Duty Exemption: SITE, MEET, TRAIN, and other approved duties are fully exempt.",
      "Late Check-in: First check-in punch after 09:00 AM (Clk1 > 9.00).",
      "Early Leave: Checked out before 17:30 on weekdays without approved leave."
    ],

    // Compliance Section
    complianceTitle: "🏛️ Enterprise Data Privacy, InfoSec & AI Governance (Compliance)",
    complianceIntro: "Built strictly on the principle of Privacy by Design, this system complies with international data privacy laws and rigorous governance frameworks to guarantee absolute employee data security:",
    complianceBadges: [
      { name: "Hong Kong Cap. 486 PDPO", desc: "Fully adheres to Personal Data (Privacy) Ordinance; zero unauthorized transfers." },
      { name: "EU GDPR & EU AI Act", desc: "100% deterministic & explainable logic with zero cross-border transfer risk." },
      { name: "CAC / PIPL Framework", desc: "Meets Personal Information Protection Law through isolated client memory execution." },
      { name: "ISO/IEC 27001 & 27701", desc: "Compliant with InfoSec and Privacy Information Management; zero disk logging." },
      { name: "ISO/IEC 42001 AIMS", desc: "Aligns with Responsible AI Management Systems ensuring audit traceability." }
    ],

    // Upload Zone
    dropTitle: "Click to select or drag attendance Excel file here",
    dropSubtitle: "Supports standard system exported .xls and .xlsx files",
    dropNotice: "Security Assurance: Data is processed strictly within local memory. Zero data is transmitted externally.",

    // Status & Actions
    processing: "Parsing workbook and executing corporate audit rules...",
    complete: "Audit complete! Identified {count} punch anomalies (Missed Punches & ABS pinned to top).",
    downloadBtn: "Download Problem.xlsx Report",
    noData: "Great! No punch anomalies detected.",

    // Fast Filter Tabs
    filterAll: "All Anomalies",
    filterMissed: "⚠️ Missed Checkout",
    filterAbs: "🚨 Absence (ABS)",
    filterLate: "Late > 9:00",
    filterEarly: "Early < 17:30",

    // Table Headers
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

  // 渲染合規專區卡片
  setSafeText('ui-compliance-title', t.complianceTitle);
  setSafeText('ui-compliance-intro', t.complianceIntro);
  const complianceContainer = document.getElementById('ui-compliance-badges');
  if (complianceContainer) {
    complianceContainer.innerHTML = t.complianceBadges.map(b => `
      <div class="bg-white p-3 rounded-lg border border-gray-200 shadow-sm flex flex-col justify-between">
        <span class="text-xs font-bold text-blue-800 bg-blue-50 px-2 py-0.5 rounded w-max mb-1.5">${b.name}</span>
        <p class="text-[11px] text-gray-600 leading-tight">${b.desc}</p>
      </div>
    `).join('');
  }

  setSafeText('ui-drop-title', t.dropTitle);
  setSafeText('ui-drop-subtitle', t.dropSubtitle);
  setSafeText('ui-drop-notice', t.dropNotice);
  setSafeText('download-btn', t.downloadBtn);

  setSafeText('tab-all', t.filterAll);
  setSafeText('tab-missed', t.filterMissed);
  setSafeText('tab-abs', t.filterAbs);
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

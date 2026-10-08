// app.js - Jumbo Orient Punch Problem Detector 核心業務審計引擎
let allProblemRows = [];
let currentFilter = 'ALL';
let outputWorkbook = null;
let targetFileName = "Problem.xlsx";

document.addEventListener("DOMContentLoaded", () => {
  if (typeof setLanguage === "function") setLanguage('zh');

  const fileInput = document.getElementById('file-input');
  const dropArea = document.getElementById('drop-area');
  const downloadBtn = document.getElementById('download-btn');

  if (dropArea && fileInput) {
    dropArea.addEventListener('click', () => fileInput.click());
    dropArea.addEventListener('dragover', (e) => { e.preventDefault(); dropArea.classList.add('bg-blue-50'); });
    dropArea.addEventListener('dragleave', () => dropArea.classList.remove('bg-blue-50'));
    dropArea.addEventListener('drop', (e) => {
      e.preventDefault();
      dropArea.classList.remove('bg-blue-50');
      if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length) handleFile(e.target.files[0]);
    });
  }

  if (downloadBtn) {
    downloadBtn.addEventListener('click', () => {
      if (!outputWorkbook) return;
      XLSX.writeFile(outputWorkbook, targetFileName);
    });
  }
});

function handleFile(file) {
  if (!file) return;
  const statusPanel = document.getElementById('status-panel');
  const statusText = document.getElementById('status-text');

  if (statusPanel) statusPanel.classList.remove('hidden');
  if (statusText) statusText.innerText = i18nData[currentLang].processing;

  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    processAttendance(workbook, file.name);
  };
  reader.readAsArrayBuffer(file);
}

function processAttendance(wb, filename) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  // 1. 動態識別 Period 檔名與截數日 (25號)
  targetFileName = "Problem.xlsx";
  let cutoffLimitDate = null;
  for (let r = 0; r < Math.min(10, rawData.length); r++) {
    const line = (rawData[r] || []).join(" ");
    const match = line.match(/Period\s*:\s*:?\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s*To\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})/i);
    if (match) {
      const [, d1, m1, y1, d2, m2, y2] = match;
      targetFileName = `${parseInt(d1)}-${parseInt(d2)}.${parseInt(m2)}.${y2}.Problem.xlsx`;
      cutoffLimitDate = new Date(parseInt(y2), parseInt(m2) - 1, 25);
      break;
    }
  }

  // 2. 定位表頭行
  let headerRow = -1;
  for (let i = 0; i < Math.min(15, rawData.length); i++) {
    if (rawData[i] && rawData[i].some(cell => String(cell).includes("Department Code"))) {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    alert(currentLang === 'zh' ? "無法找到 'Department Code' 表頭！" : "Cannot find 'Department Code' header!");
    return;
  }

  const rows = rawData.slice(headerRow + 2);
  const targetDepts = ['PCD', 'HOF', 'PMD', 'POD', 'QSD', 'WMD', 'ADD', 'CPD', 'SED'];
  const rawProblems = [];

  rows.forEach(r => {
    const dept = String(r[0] || '').trim().toUpperCase();
    const deptName = String(r[1] || '').trim();
    const empCode = String(r[2] || '').trim().toUpperCase();
    const name = String(r[3] || '').trim();
    const desig = String(r[4] || '').trim();
    
    // 僅審計 E 職員與目標業務部門
    if (!empCode.startsWith('E') || !targetDepts.includes(dept)) return;

    const dateStr = String(r[5] || '').trim();
    const rowDate = new Date(dateStr);
    if (cutoffLimitDate && !isNaN(rowDate) && rowDate > cutoffLimitDate) return;

    const day = String(r[6] || '').trim();
    const clk1 = parseFloat(r[7]) || 0;
    const clk2 = parseFloat(r[8]) || 0;
    const clk3 = parseFloat(r[9]) || 0;
    const clk4 = parseFloat(r[10]) || 0;
    const timeIn = parseFloat(r[11]) || 0;
    const timeOut = parseFloat(r[12]) || 0;
    const shift = String(r[13] || '').trim().toUpperCase();
    const normal = parseFloat(r[14]) || 0;
    const reasonRaw = String(r[23] || '').trim().toUpperCase();

    // 排除休假/公眾假期
    if (shift === 'REST' || reasonRaw === 'RES' || reasonRaw === 'PH') return;

    const lastClk = Math.max(clk1, clk2, clk3, clk4);

    // 業務審計判定與優先級
    let anomalyReason = "";
    let category = "";
    let priority = 99;

    // 🚨 優先級 1: 全天缺勤 (ABS)
    if (reasonRaw === 'ABS') {
      anomalyReason = clk1 > 0 ? "缺勤但有打卡 (ABS with Punch)" : "全天缺勤 (ABS)";
      category = "ABS";
      priority = 1;
    }
    // ⚠️ 優先級 2: 漏打收工卡
    else if (clk1 > 0 && lastClk === clk1 && (timeOut === 0 || lastClk < 13.00)) {
      anomalyReason = "漏打收工卡 (Missed Checkout)";
      category = "MISSED";
      priority = 2;
    }
    // 優先級 3: 特殊任務無實質打卡
    else if (clk1 === 0 && lastClk === 0 && ['TRAIN', 'SITE', 'MEET', 'EVENT', 'OTHER'].includes(reasonRaw)) {
      anomalyReason = `特殊任務無打卡 (${reasonRaw})`;
      category = "MISSED";
      priority = 3;
    }
    // 優先級 4: 遲到 (遲過 9:00 返工)
    else if (clk1 > 9.00 && !reasonRaw) {
      anomalyReason = `遲過九點 (Late: ${clk1.toFixed(2)} > 9:00)`;
      category = "LATE";
      priority = 4;
    }
    // 優先級 5: 早退 (早過 17:30 收工，星期一至五標準日)
    else if (day !== 'Saturday' && normal >= 7.0 && lastClk > 0 && lastClk < 17.30 && !reasonRaw) {
      anomalyReason = `早過五點半 (Early: ${lastClk.toFixed(2)} < 17:30)`;
      category = "EARLY";
      priority = 5;
    }

    if (anomalyReason) {
      rawProblems.push({
        dept, deptName, empCode, name, desig,
        date: dateStr, day, clk1, clk2, clk3, clk4,
        timeIn: (clk1 === 0 && lastClk === 0) ? 0 : timeIn,
        timeOut: (clk1 === 0 && lastClk === 0) ? 0 : timeOut,
        reason: anomalyReason,
        category: category,
        priority: priority
      });
    }
  });

  // 排序：ABS 最前，漏打卡第二
  rawProblems.sort((a, b) => a.priority - b.priority);
  allProblemRows = rawProblems;

  window.lastProblemCount = allProblemRows.length;
  const statusText = document.getElementById('status-text');
  if (statusText) {
    statusText.innerText = i18nData[currentLang].complete.replace('{count}', allProblemRows.length);
  }

  applyFilter('ALL');
  buildOutputExcel(allProblemRows);
}

function applyFilter(filterKey) {
  currentFilter = filterKey;
  document.querySelectorAll('.filter-tab').forEach(tab => {
    tab.classList.remove('bg-gray-900', 'text-white');
    tab.classList.add('bg-gray-100', 'text-gray-700');
  });

  const activeTabMap = {
    'ALL': 'tab-all',
    'ABS': 'tab-abs',
    'MISSED': 'tab-missed',
    'LATE': 'tab-late',
    'EARLY': 'tab-early'
  };

  const activeTab = document.getElementById(activeTabMap[filterKey]);
  if (activeTab) {
    activeTab.classList.remove('bg-gray-100', 'text-gray-700');
    activeTab.classList.add('bg-gray-900', 'text-white');
  }

  const filtered = filterKey === 'ALL' 
    ? allProblemRows 
    : allProblemRows.filter(r => r.category === filterKey);

  renderPreview(filtered);
}

function renderPreview(rows) {
  const previewBody = document.getElementById('preview-body');
  if (!previewBody) return;

  previewBody.innerHTML = '';
  if (rows.length === 0) {
    previewBody.innerHTML = `<tr><td colspan="10" class="p-6 text-center text-gray-400">${i18nData[currentLang].noData}</td></tr>`;
    return;
  }

  rows.forEach(r => {
    const tr = document.createElement('tr');
    
    let rowClass = "hover:bg-gray-50 transition border-b";
    let badgeClass = "px-2 py-0.5 rounded text-xs font-semibold";

    if (r.category === 'ABS') {
      rowClass = "bg-red-50 hover:bg-red-100 transition border-b border-red-100";
      badgeClass += " bg-red-100 text-red-800 border border-red-300 font-bold";
    } else if (r.category === 'MISSED') {
      rowClass = "bg-yellow-50 hover:bg-yellow-100 transition border-b border-yellow-100";
      badgeClass += " bg-yellow-100 text-yellow-800 border border-yellow-300 font-bold";
    } else if (r.category === 'LATE') {
      badgeClass += " bg-blue-50 text-blue-700 border border-blue-200";
    } else {
      badgeClass += " bg-purple-50 text-purple-700 border border-purple-200";
    }

    tr.className = rowClass;
    tr.innerHTML = `
      <td class="px-3 py-2 font-mono font-medium">${r.dept}</td>
      <td class="px-3 py-2 text-gray-600">${r.deptName}</td>
      <td class="px-3 py-2 font-mono font-bold text-gray-800">${r.empCode}</td>
      <td class="px-3 py-2 font-medium text-gray-900">${r.name}</td>
      <td class="px-3 py-2">${r.date}</td>
      <td class="px-3 py-2">${r.day}</td>
      <td class="px-3 py-2 font-mono text-gray-500">${r.clk1} / ${r.clk2} / ${r.clk3} / ${r.clk4}</td>
      <td class="px-3 py-2 font-mono font-bold text-gray-700">${r.timeIn.toFixed(2)}</td>
      <td class="px-3 py-2 font-mono font-bold text-gray-700">${r.timeOut.toFixed(2)}</td>
      <td class="px-3 py-2"><span class="${badgeClass}">${r.reason}</span></td>
    `;
    previewBody.appendChild(tr);
  });
}

function buildOutputExcel(rows) {
  const header1 = ['Department Code', 'Department Name', 'Emp Code', 'Name', 'Designation', 'Date', 'Day', 'Actual Clocking', '', '', '', 'Time', '', 'Reason'];
  const header2 = ['', '', '', '', '', '', '', '', '', '', '', 'IN', 'OUT', ''];
  const aoa = [header1, header2];

  rows.forEach(r => {
    aoa.push([
      r.dept, r.deptName, r.empCode, r.name, r.desig,
      r.date, r.day, r.clk1, r.clk2, r.clk3, r.clk4, r.timeIn, r.timeOut, r.reason
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = [
    { s: { r: 0, c: 7 }, e: { r: 0, c: 10 } },
    { s: { r: 0, c: 11 }, e: { r: 0, c: 12 } }
  ];

  outputWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputWorkbook, ws, "Problem.Log");
}

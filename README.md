# Jumbo Orient Punch Problem Detector (JO-PPD)
### 東淦打卡異常偵測工具・企業級純前端考勤審計工作站

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-Client--Side%20Web%20(SheetJS)-green)
![Privacy](https://img.shields.io/badge/Data%20Privacy-Zero--Data%20Retention-emerald)
![Jurisdiction](https://img.shields.io/badge/Jurisdiction-Hong%20Kong%20SAR-orange)

---

## 🌐 項目線上系統 (Live Workstation)
* **GitHub Pages 部署網址:** [https://jackylawck.github.io/JO-PPD/](https://jackylawck.github.io/JO-PPD/)
* **官方企業網站:** [https://www.jumboorient.com.hk/](https://www.jumboorient.com.hk/)

---

## 📖 繁體中文項目簡介

### 1. 系統定位與研發背景
**Jumbo Orient Punch Problem Detector (JO-PPD / 東淦打卡異常偵測工具)** 是由**東淦工程有限公司（Jumbo Orient Contracting Limited）**自主研發的企業級無伺服器考勤完整性快篩工具。

本系統專為解決每月份人事考勤審計中的大量重複性工序而設計。針對原廠考勤報表中龐雜的打卡原始數據，系統透過確定性（Deterministic）業務規則，秒級提取打卡異常、漏簽與缺勤記錄，並支援一鍵導出對齊原廠規格之 `Problem.xlsx` 報表。

### 2. 核心架構：零數據留存 (Zero-Data Retention)
* **純客戶端本地記憶體沙盒運算 (Client-Side Sandboxing):** 所有 Excel 文件的解析、運算、過濾與二進制重組均在使用者本機瀏覽器的記憶體中完成。
* **零伺服器傳輸 (Zero Network Transmission):** 系統不設任何後端 API，不傳輸任何員工姓名、工號或考勤打卡數據至外部雲端。
* **從設計著手保護隱私 (Privacy by Design):** 關閉或刷新瀏覽器分頁即刻銷毀記憶體數據，杜絕資料外洩與跨國境傳輸風險。

### 3. 公司業務審計基準 (Corporate Audit Rules)
系統依據公司現行出勤指引與彈性工時標準進行精準診斷：
1. **⚠️ 漏打收工卡 (Missed Checkout - 置頂高亮):**
   * 首度打卡有效（$\text{Clk}_1 > 0$ 或 $\text{Time IN} > 0$），但系統收工時間完全未記錄（$\text{Time OUT} = 0.00$），且下半天無有效打卡。
   * **智慧豁免機制:** 若系統已有合法收工時間（$\text{Time OUT} \ge 17.00$），即使實體打卡槽位用滿亦自動豁免，杜絕誤報。
2. **🚨 全天缺勤 (Absence - 置頂高亮):**
   * 系統標記為 `ABS` 且整天打卡為空（$0$ 次打卡）。若有打卡則獨立標註為「缺勤但有打卡」。
3. **正當外勤全面放行 (Approved Out-of-Office Exemption):**
   * 標註有地盤外勤（`SITE`）、業務會議（`MEET`）、內部培訓（`TRAIN`）等已審批活動全面放行，不視為打卡異常。
4. **遲過九點 (Late Check-in):**
   * 首度上班打卡遲過上午 9:00（$\text{Clk}_1 > 9.00$）且無請假單。
5. **早過五點半 (Early Leave):**
   * 星期一至五標準工時工作日，收工打卡與系統記錄皆早於下午 5:30（$\max(\text{Last\_Clk}, \text{Time\_Out}) < 17.30$）且無請假單。

---

## 🌐 English Project Overview

### 1. System Mission & Scope
The **Jumbo Orient Punch Problem Detector (JO-PPD)** is an enterprise-grade client-side attendance integrity inspection workstation independently designed and maintained for **Jumbo Orient Contracting Limited**.

It eliminates manual screening overhead across thousands of monthly timecard entries. Operating on strict deterministic logic, it instantly filters punch anomalies, missed checkout punches, and unverified absences, exporting a standardized audit workbook (`Problem.xlsx`) fully compatible with legacy systems.

### 2. Architecture: Zero-Data Retention
* **Client-Side Sandbox Execution:** All Excel reading, array parsing, rule verification, and workbook rebuilding occur purely in-memory via SheetJS inside the user's browser endpoint.
* **Zero Network Traffic:** Zero backend servers, zero API requests, and zero data logging. Employee Personal Identifiable Information (PII) never traverses external networks.
* **Privacy by Design:** All in-memory structures are cleared instantly upon tab closure, adhering to modern zero-trust enterprise security baselines.

### 3. Audit Logic Benchmarks
1. **⚠️ Missed Checkout (Top Priority Alert):**
   * Valid clock-in recorded, but checkout is completely empty ($\text{Time OUT} = 0.00$) with no afternoon clocking.
   * **Smart Exemption:** Automatically excludes false alarms if $\text{Time OUT}$ is already recognized ($\ge 17.00$) by administrative adjustment.
2. **🚨 Unexcused Absence (ABS - Top Priority Alert):**
   * Tagged as `ABS` with zero punch entries. Records with punches are differentiated for review.
3. **Approved Official Duty Immunity:**
   * Legitimate off-site assignments (`SITE`, `MEET`, `TRAIN`, etc.) are unconditionally exempted.
4. **Late Check-in:** First punch after 09:00 AM ($\text{Clk}_1 > 9.00$) without approved leave.
5. **Early Leave:** Final checkout before 17:30 PM on standard weekdays without approved leave.

---

## 📂 專案檔案結構 (Repository Structure)

```text
JO-PPD/
├── index.html          # 結構骨架、SEO 宣傳標籤、Google 驗證與視窗容器
├── app.js              # 核心業務審計引擎、SheetJS 解析、分類排序與 Excel 導出
├── i18n.js             # 國際化雙語模組 (全繁體中文 / 全英文切換及合規說明)
├── COMPLIANCE.md       # 企業級法規適用性判定與科技治理白皮書
├── TERMS.md            # 法律使用條款、知識產權防護與審計免責聲明
├── LICENSE             # MIT 開源授權協議條款
├── .gitignore          # 嚴格禁止任何考勤原檔 (*.xls, *.xlsx) 提交至雲端
└── README.md           # 專案說明文檔

```

---

## 🏛️ 企業治理與法規宣告 (Governance & Compliance)

本項目技術特徵為「純確定性邏輯（Deterministic Logic）」，具備 100% 數學可解釋性，不包含黑箱概率推論或生成模型。詳細法律與標準對照請查閱專題文件：

* **[合規架構白皮書 (COMPLIANCE.md)](https://github.com/jackylawck/JO-PPD/blob/77e8f20313cd3f15bac70ede126b8a046c4437b1/COMPLIANCE.md):**
闡明對香港 Cap. 486《個人資料（私隱）條例》、ISO/IEC 27001、ISO/IEC 27701 的完全合規實踐，並依法出具對 EU AI Act、ISO/IEC 42001、GDPR 及國家網信辦算法備案的**非適用性宣告（Statement of Non-Applicability）**。
* **[法律條款與審計免責 (TERMS.md)](https://github.com/jackylawck/JO-PPD/blob/77e8f20313cd3f15bac70ede126b8a046c4437b1/TERMS.md):**
明定本系統屬於「管理輔助診斷工具」，確立人事主管人工複核（Human-in-the-Loop）原則，免除演算法直接作為法律紀律處分憑證之連帶責任。


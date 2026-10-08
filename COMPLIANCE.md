# 企業級法規適用性判定、數據主權與科技治理白皮書
## Compliance, Data Sovereignty & Tech Governance Whitepaper

**項目名稱 (Project):** Jumbo Orient Punch Problem Detector (JO-PPD / 東淦打卡異常偵測工具)  
**研發機構 (Organization):** 東淦工程有限公司 (Jumbo Orient Contracting Limited)  
**系統架構師 (Lead Architect):** 羅子淇 (Jacky Law, F.I.H.R.M., FHKIoD)  
**生效日期 (Effective Date):** 2026-10-01  
**版本 (Version):** 1.0.0 (Production Release)

---

## 1. 執行摘要 (Executive Summary)

**中文：**  
本系統專為東淦工程有限公司內部人力資源管理設計，旨在審計內部考勤打卡數據之完整性。本白皮書旨在從法律合規（Legal Compliance）、數據主權（Data Sovereignty）、資訊安全（InfoSec）與負責任科技治理（Responsible Tech Governance）四大維度，對系統技術邊界出具正式合規聲明。本系統底層架構採取「純客戶端本地記憶體沙盒運算（Client-Side In-Memory Execution）」，不設立任何伺服器端後端，完全切斷個人資料（PII）向外傳輸之途徑。

**English:**  
This system is an enterprise internal audit utility engineered for Jumbo Orient Contracting Limited to inspect monthly punch records for integrity anomalies. This whitepaper establishes a formal declaration regarding Legal Compliance, Data Sovereignty, Information Security (InfoSec), and Responsible Technology Governance. Architected on a strict "Client-Side In-Memory Sandboxed Execution" model with zero backend storage, the system inherently eliminates any transmission vector of Personally Identifiable Information (PII) beyond the local browser endpoint.

---

## 2. 法律與標準適用性矩陣 (Compliance Applicability Matrix)

| 規範 / 標準 (Framework / Standard) | 適用狀態 (Status) | 技術與治理實踐 (Implementation & Rationale) |
| :--- | :--- | :--- |
| **香港《個人資料（私隱）條例》(Cap. 486 PDPO)** | **全面合規 (Fully Compliant)** | 貫徹保障資料原則（DPP 1 至 DPP 4）。數據完全留存於用戶本機記憶體，關閉分頁即刻銷毀，無任何未授權存取或第三方洩漏。 |
| **ISO/IEC 27001 (資訊安全管理體系)** | **全面合規 (Fully Compliant)** | 最小特權與零外部攻擊面原則。純靜態託管於 GitHub Pages，不提供任何 API 入口或外部寫入權限，無服務器被滲透之風險。 |
| **ISO/IEC 27701 (隱私資訊管理體系)** | **全面合規 (Fully Compliant)** | 貫徹「從設計著手保護隱私（Privacy by Design）」理念。所有員工考勤數據之處理週期均控制於單機瀏覽器沙盒，不收集、不持久化任何 Cookie 或日誌。 |
| **香港《僱傭條例》(Cap. 57)** | **合規支持 (Compliance Enabler)** | 確定性記錄比對，不具備自動處分或扣薪功能。僅作為客觀數據輔助，保障員工合法休假與法定權益，遵守人工最終複核原則。 |
| **歐盟 EU AI Act (人工智慧法案)** | **非適用宣告 (SoNA - Non-Applicable)** | **出具法定非適用性宣告**：系統依據固定布林運算子判定，非自主學習或機器學習黑箱系統，不構成歐盟 AI 法案管轄之 AI 系統。 |
| **ISO/IEC 42001 (人工智慧管理體系)** | **非適用宣告 (SoNA - Non-Applicable)** | 系統未採納任何生成式模型或概率型推論技術，不屬於 AIMS（AI Management System）之審計受體，但全面貫徹其可解釋性核心精神。 |
| **歐盟通用數據保護條例 (GDPR)** | **安全豁免 (Exempt by Architecture)** | 本項目為香港境內本土工程營運使用，且不具備雲端傳輸能力，無任何跨境傳輸（Cross-Border Data Transfer）或境外受試者資料收集行為。 |
| **國家網信辦《生成式AI服務管理辦法》** | **非適用宣告 (SoNA - Non-Applicable)** | 本項目不面向大眾提供生成文本、圖像或語音服務，亦不具備任何模型訓練，毋須履行演算法備案手續。 |

---

## 3. 技術特徵與確定性治理 (Deterministic Engineering Architecture)

### 3.1 零數據留存架構 (Zero-Data Retention Sandbox)
* **無網絡後端 (No Backend Endpoint):** 系統運作完全依賴瀏覽器載入之 JavaScript（SheetJS），所有 Excel 表格解碼、過濾與二進制重構皆於使用者的 RAM 中執行。
* **無持久化存儲 (No Persistence):** 系統不使用 `localStorage`、`sessionStorage`、`IndexedDB`，亦不植入任何跨站追蹤 Cookie。分頁刷新或關閉時，所有考勤記憶體空間立即釋放。
* **靜態 CSP 嚴格阻斷 (Strict Content Security Policy):** 網頁標頭實施嚴格 CSP，阻斷所有未授權之外部腳本載入與 `connect-src` 外部回傳。

### 3.2 確定性審計規則（非黑箱演算法）
本系統之核心判定邏輯嚴格建構於具備 100% 數學確定性之二元邏輯（Deterministic Rule-Based Logic）：
1. **缺勤邏輯:** $\text{Status} = \text{ABS}$
2. **漏打卡邏輯:** $\text{Valid\_Punch\_Count} = 1 \land \text{Time\_Out} = 0 \land \text{Duty} \notin \{\text{Approved Exemptions}\}$
3. **遲到邏輯:** $\text{Clk}_1 > 9.00 \land \text{Duty} \notin \{\text{Approved Exemptions}\}$
4. **早退邏輯:** $\max(\text{Last\_Clk}, \text{Time\_Out}) < 17.30 \land \text{Duty} \notin \{\text{Approved Exemptions}\}$

系統**不使用**任何具機率性、隨機性或權重調整之神經網絡或預測模型，徹底消除「演算法偏見（Algorithmic Bias）」與「AI 幻覺（Hallucination）」，任何輸入皆可重現且具備完全之審計回溯性（Audit Traceability）。

---

## 4. 監管聯絡與問責 (Accountability & Contact)

如對本項目之科技架構、數據合規或法規適用性有任何治理諮詢，請聯絡：

* **企業實體:** 東淦工程有限公司 (Jumbo Orient Contracting Limited)
* **管治顧問 / 架構師:** 羅子淇 (Jacky Law, Senior HR Manager & Corporate Governance Advisor)
* **官方網址:** [https://www.jumboorient.com.hk/](https://www.jumboorient.com.hk/)
* **開放源碼庫:** [https://github.com/jackylawck/JO-PPD](https://github.com/jackylawck/JO-PPD)


## 一句话

从 UCI 的 `online_retail_II.xlsx` 出发，用 Python 做清洗与业务分析，用本机 Excel 生成带**原生 PivotTable** 的工作簿和截图，再打成 ZIP，并用 Outlook 准备提交邮件（默认只预览、不真发）。

```bash
python main.py                  # 全流程
python main.py --reuse-clean    # 跳过读原始 xlsx，复用已清洗 parquet
```

---

## 1. 作业要什么（对照清单）

| # | 作业要求 | 主模块 | 主要产物 |
|---|----------|--------|----------|
| 1 | 数据准备：UCI 文件、清洗、审计 | `data_processing.py` | parquet + `output/audit/` |
| 2 | 六类业务分析 | `analysis.py` | Summary KPI + 分析表 |
| 3 | 异常 Top 20 | `analysis.detect_anomalies` | Anomalies sheet |
| 4 | Excel：Clean_Data / Summary / 6×原生 Pivot / Anomalies | `excel_automation.py` | `Retail_Analysis.xlsx` |
| 5 | 6 张 Pivot PNG | `export_pivot_pngs` | `output/png/` |
| 6 | Outlook：主题、摘要、正文内嵌 2 图、附 xlsx | `email_automation.py` | 预览 HTML；过大则 EmailSafe ZIP |
| 7 | 工程：模块化、校验/日志、无硬编码密钥、requirements+README、关 Excel/Outlook；ZIP 包 | `main` + `validation` + `packaging` | `{Name}_PythonAssessment.zip` |

---

## 2. 代码框架（先讲清楚结构）

### 2.1 设计原则

- **`main.py` 只编排**，不写业务公式。
- **口径只认一处**：`src/metrics.py`（v1.6.0），分析 / Summary / 邮件共用。
- **配置与密钥进 `.env`**：姓名、自测邮箱；路径相对仓库根目录（`config.PROJECT_ROOT`）。
- **重活可跳过**：`--reuse-clean` / `--skip-excel` / `--email-preview-only`，方便迭代。
- **默认不发信**；`--send` 必须有 `--send-to` 或 `POC_TEST_EMAIL`，**从不**自动用招聘方 `SUBMISSION_EMAIL`。

### 2.2 目录与职责

```text
python-retail-assessment/
├── main.py                 # CLI + Step 1→6 编排
├── requirements.txt
├── README.md
├── .env.example            # CANDIDATE_NAME / POC_TEST_EMAIL（无密钥）
├── docs/
│   ├── ASSIGNMENT_WALKTHROUGH.md   # 本文
│   └── PIPELINE_FLOW.md            # 函数级细节
├── src/
│   ├── config.py           # 路径、ensure_dirs、读 .env
│   ├── logging_utils.py    # 控制台 + output/logs/pipeline.log
│   ├── metrics.py          # Gross / Net / 退货率 / 异常合同
│   ├── data_processing.py  # 读表、跨表去重、enrich、审计
│   ├── analysis.py         # 六类分析 + 异常 + AnalysisBundle
│   ├── excel_automation.py # xlsxwriter 底稿 + xlwings 原生 Pivot + PNG
│   ├── validation.py       # 数据/分析/工作簿/PNG 断言
│   ├── packaging.py        # 完整 ZIP / EmailSafe ZIP
│   └── email_automation.py # HTML 预览 + Outlook appscript
├── tests/                  # 合成数据单测（不启 Excel）
├── scripts/                # 早期 PoC（非正式提交主路径）
├── data/
│   ├── raw/                # online_retail_II.xlsx（gitignore）
│   └── processed/          # clean_transactions.parquet
└── output/                 # xlsx / png / zip / audit / email_preview
```

### 2.3 `main.py` 怎么串起来

```text
parse_args → setup_logging → ensure_dirs
    │
    ├─ Step 1  prepare_dataset() 或 load_prepared()     → PrepResult(clean, audit)
    ├─ Step 2  run_analyses(clean)                      → AnalysisBundle
    │            写 analysis_summary.json
    ├─ Step 3  build_workbook(clean, bundle, audit)     → Retail_Analysis.xlsx + PNGs
    ├─ Step 4  run_validations(...)                     → 失败则 return 1
    ├─ Step 5  build_submission_zip + build_email_safe_zip
    └─ Step 6  write_dry_run_preview
                 └─ 若 --send → send_submission_email(..., send=True)
```

对象在步骤之间的传递：

| 对象 | 谁产出 | 谁消费 |
|------|--------|--------|
| `PrepResult.clean` | `data_processing` | `analysis`、`excel`、`validation` |
| `PrepResult.audit_summary` | `data_processing` | `excel` Summary、校验 |
| `AnalysisBundle` | `analysis.run_analyses` | Excel、邮件、校验 |
| `workbook` Path | `excel.build_workbook` | 校验、打包、邮件 |
| ZIP Path | `packaging` | 邮件附件 |

`AnalysisBundle` 里大致有：`summary`（KPI dict）、六张分析 DataFrame、`anomalies`、`narrative`（给 Summary/邮件用的说明文字）。

### 2.4 模块 ↔ 关键函数（口述速查）

| 模块 | 关键入口 | 内部还会碰到 |
|------|----------|--------------|
| `data_processing` | `prepare_dataset` / `load_prepared` | `load_raw` → `apply_cross_sheet_dedupe` → `enrich` → 审计 JSON/CSV |
| `metrics` | `MetricsContract` / `describe_contract` | Gross、`is_return_line` 子类、Net、return_line_rate、异常配额 |
| `analysis` | `run_analyses` | `product_/country_/monthly_/returns_/customer_/sales_timing` + `detect_anomalies` + `build_summary` |
| `excel_automation` | `build_workbook` | `_write_base_workbook_*` → xlwings `PIVOT_SPECS` → `export_pivot_pngs` |
| `validation` | `run_validations` | prep / analyses / workbook（含 Pivot、Grand Total）/ pngs |
| `packaging` | `build_submission_zip` | `_add_common_files`（源码+README+docs+PNG+audit） |
| `email_automation` | `write_dry_run_preview` / `send_submission_email` | `build_email_html`、attachments 策略 |

---

## 3. 怎么一步步做（故事 + 一点细节）

### 第 0 步：搭架子

- 按作业拆模块（上表）；日志进 `output/logs/`。
- `.env`：`CANDIDATE_NAME`（ZIP 名与邮件主题）、`POC_TEST_EMAIL`（自测）、`SUBMISSION_EMAIL`（仅备忘，不自动发）。
- 依赖：`pandas` / `pyarrow` / `xlsxwriter` / `xlwings` / `appscript` / `Pillow` / `python-dotenv`（见 `requirements.txt`）。

### 第 1 步：数据准备

**入口：** `prepare_dataset(input_path?)` → `PrepResult`

1. **定位文件** `find_input_workbook`：`data/raw/online_retail_II.xlsx` 或 `--input`。
2. **`load_raw`**：读 Year 2009-2010 / 2010-2011 两个 sheet，合并，加 `source_sheet` / `source_row` / `load_row_id`（约 106 万行）。
3. **`apply_cross_sheet_dedupe`**：按业务主键跨表去重，保留较新 sheet；删除行写入 `removed_cross_sheet_clones.csv`。清洗后约 **1,044,848** 行。
4. **`enrich`（打标，不乱删交易）** 关键列：

   | 列 / 概念 | 含义 |
   |-----------|------|
   | `line_revenue` | Quantity × Price |
   | `is_cancellation` | Invoice 以 `C` 开头 |
   | `is_return_line` | **作业 return indicator**：非 C 且 Quantity&lt;0 |
   | `is_confirmed_return` | 子类 Price&gt;0 → 进 `return_value`，扣 Net |
   | `is_stock_adjustment` | 子类其余（多为 Price≤0）→ 只计件数，**不扣 Net** |
   | `return_line_class` | `confirmed_return` / `stock_adjustment` |
   | `year_month` / weekday / hour | 时段与月度分析 |
   | `product_label` | `StockCode — Description`（产品 Pivot 行） |
   | `Country_NonUK` | UK 置空，供国家 Pivot |
   | distinct helpers | 发票/客户等，供「看似 distinct count」的 Sum 透视 |

5. **落盘**：`data/processed/clean_transactions.parquet`；审计含 `data_quality_summary.json`、`revenue_reconciliation.json`、`changed_records.csv` 等。

**口述重点：** 取消 ≠ 退货；退货再拆「确认退货」vs「库存调整」。恒等式：`is_return_line` 行数 = confirmed + adjustment。

### 第 2 步：六类分析 + 异常

**入口：** `run_analyses(clean)` → `AnalysisBundle`；KPI 写入 `output/analysis_summary.json`。

| 函数 | 细节 |
|------|------|
| `product_performance` | 非取消正销量收入；Top 产品进 Summary |
| `country_performance` | **排除 UK**；AOV；缺 Customer ID 收入占比 |
| `monthly_performance` + `complete_months` | 按全局日期跨度标完整月；best/worst 只用完整月（本数据边缘月约 2009-12、2011-12） |
| `returns_performance` | 仅 `units_sold ≥ 50`；排序键 **`return_line_rate`** = (退货件数+调整件数)/销量；旁注取消 |
| `customer_value` | 仅已知 Customer ID；收入与发票数 → AOV |
| `sales_timing` | 星期 × 小时矩阵；`peak_weekday` / `peak_hour` **分开**聚合 |
| `detect_anomalies` | 正价销售 IQR×3 + 规则；配额约 **12 金额 / 4 调整 / 4 客户** → Top 20；阈值写入 `audit/anomaly_thresholds.json` |
| `build_summary` | Gross、return_value、cancellation_value、Net、Net after cancellations、Top 结论等 |

**口径（讲数字时用）：**

```text
Gross          = 非取消正价行收入
return_value   = 确认退货金额（本数据集上常为 0）
Net            = Gross − return_value
cancellation_value = |C 发票收入|
Net after canc = Gross − return_value − cancellation_value
                 （约 £20.4m → £18.9m，取消约 £1.47m）
```

### 第 3 步：Excel + 原生 Pivot + PNG

**入口：** `build_workbook(clean, bundle, audit)`

**Phase A — 底稿（XlsxWriter）**

- 列筛选 + `CLEAN_DATA_EXCEL_RENAME`：Excel 里显示 `Revenue (GBP)` 等，而不是 `inv_flag_*`。
- 写入：`Clean_Data`、`Summary`、`Anomalies`、六张隐藏的汇总数据源、六个空 Pivot sheet。

**Phase B — 本机 Excel（xlwings）**

- 按 `PIVOT_SPECS` 建**原生** PivotTable：

  | Sheet | 行字段 | 值 / 备注 |
  |-------|--------|-----------|
  | Pivot_Product_Revenue | `product_label` | 收入等；藏 bottom GT |
  | Pivot_Country_Performance | 非英国国家（前五） | Pivot 内 AOV 值字段；藏 GT |
  | Pivot_Monthly_Performance | Year-month | 藏 GT |
  | Pivot_Returns | Product（源=`Returns_Source`） | 含 rate；**藏 GT**（rate 不可加总） |
  | Pivot_Customer_Value | Customer ID | 已知客户汇总 + Pivot 内 AOV；藏 GT |
  | Pivot_Sales_Timing | Day × Hour | 收入矩阵；可留 GT |

- 刷新后 `finally` 关闭 `wb` / `wb2` 并 `app.quit()`。

**PNG：** `export_pivot_pngs`；逐张截取完整原生 Pivot 区域，排名范围已在 Python 数据源中限定。PNG 失败会 **中断 pipeline**，避免旧图混进邮件。

工作簿保留百万行 Clean_Data，并使用小型隐藏表作为 Pivot 数据源。

### 第 4 步：校验

**入口：** `run_validations(clean, audit, bundle, workbook, check_pngs=…)`

检查示例：行数与审计一致、UK 不在国家分析、peak 字段存在、六张 Pivot 存在、非加总 sheet 无 Grand Total 行、6 张 PNG 非空。失败 → `main` 返回 1。

### 第 5 步：打包

| ZIP | 函数 | 内容 |
|-----|------|------|
| `{CANDIDATE_NAME}_PythonAssessment.zip` | `build_submission_zip` | 源码、requirements、README、docs、xlsx、png、audit、summary — **正式交这个** |
| `…_EmailSafe.zip` | `build_email_safe_zip` | 同上但**无 xlsx**，给邮箱用 |

包内包含测试和流程文档。原始 UCI 文件和本地配置 `.env` 由运行者单独准备。

### 第 6 步：邮件

1. 主题：`Python Assessment Submission - <Candidate Name>`
2. `build_email_html`：短说明 + 摘要 + **两张图 data-URI 内嵌**（另附 PNG 文件防 Gmail 剥图）
3. xlsx 超过 18MB 时 CLI 阻止发送。完整 ZIP 包含 xlsx。预览目录另提供 CID 内嵌的本地 EML 草稿
4. 默认 `write_dry_run_preview` → `output/email_preview/`；确认后  
   `python main.py --email-preview-only --send --send-to you@example.com`  
   发送有 try/except，失败写 `submission_receipt.json` 的 `status=failed`

### 第 7 步：工程收尾

- 模块化 + 日志 + 校验 + 错误返回码  
- `requirements.txt` + README  
- `pytest tests/`：合成数据测清洗/指标/退货率等（不启 Excel）  
- Excel / Outlook 对象有关闭路径  

---

## 4. 端到端一张图

```text
UCI xlsx
  → data_processing（去重 + enrich + 审计）→ parquet
  → analysis（六类 + Top20 + summary）
  → excel_automation（底稿 + 原生 Pivot + PNG）
  → validation
  → packaging（完整 ZIP + EmailSafe）
  → email_automation（预览 / 可选 Outlook 发送）
```

---

## 5. 常被问的决策

1. **取消 ≠ 退货** — `C…` 发票 vs 非 C 且 Qty&lt;0；只有确认退货扣 Net。  
2. **为何两个 Net** — 作业看退货；另报扣取消后的参考净额（≈ £18.9m）。  
3. **为何 xlsx 巨大** — 真 Pivot + 全量 Clean_Data。  
4. **邮件为何不附 xlsx** — 邮箱体积；完整包在 ZIP。  
5. **为何要本机 Excel/Outlook** — 作业要求客户端自动化（xlwings + appscript）。  
6. **Returns 为何藏 Grand Total** — 表上有预计算 rate，Sum of rates 无意义。

---

## 6. 演示怎么跑 / 看什么

```bash
cd /path/to/python-retail-assessment
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env          # 填 CANDIDATE_NAME
# 放置 data/raw/online_retail_II.xlsx
python main.py                # 或 --reuse-clean
python -m pytest tests/ -q
```

| 想看 | 打开 |
|------|------|
| KPI | `output/analysis_summary.json` 或 xlsx → Summary |
| 透视 | `output/Retail_Analysis.xlsx` |
| 截图 | `output/png/` |
| 提交包 | `output/{姓名}_PythonAssessment.zip` |
| 邮件预览 | `output/email_preview/submission_email_preview.html` |
| 口径定义 | `src/metrics.py` |

---

## 7. 和 PIPELINE_FLOW 的分工

| 文档 | 适合 |
|------|------|
| **本文** | 口述：框架 + 作业步骤 + 关键细节 |
| [PIPELINE_FLOW.md](PIPELINE_FLOW.md) | 逐函数、逐文件对照 |

---

*指标以 `src/metrics.py`（v1.6.0）与 Summary 表为准。*

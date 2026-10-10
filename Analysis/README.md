<div align="center">

# -- ! Delivery Delay Analysis ! --
### *Finding the Service Type & Hub with the Greatest Delivery-Delay Burden using Excel, SQL & Python*
**Data Analysis Practical Exam — Set A | Student ID: `12827`**

[![Excel](https://img.shields.io/badge/Excel-XLOOKUP%20%7C%20SUMIFS%20%7C%20Pivot-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![SQL](https://img.shields.io/badge/SQL-JOIN%20%7C%20GROUP%20BY%20%7C%20HAVING-4479A1?style=for-the-badge&logo=mysql&logoColor=white)](https://www.w3schools.com/sql/)
[![Python](https://img.shields.io/badge/Python-pandas%20%7C%20matplotlib-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Analysis](https://img.shields.io/badge/Analysis-Delay%20Incidence%20Rate-FF6F00?style=for-the-badge&logo=databricks&logoColor=white)](https://github.com/)

<br/>

> *"On-time delivery is a promise — every delayed day is a promise broken."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🗂️ Dataset & Data Dictionary](#️-dataset--data-dictionary)
- [🧹 Data Cleaning & Metric Definitions](#-data-cleaning--metric-definitions)
- [📗 Task 1 — Excel Analysis](#-task-1--excel-analysis)
- [🗄️ Task 2 — SQL Analysis](#️-task-2--sql-analysis)
- [🐍 Task 3 — Python Analysis](#-task-3--python-analysis)
- [🔗 Cross-Tool Reconciliation](#-cross-tool-reconciliation)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [💡 Recommendation & Limitation](#-recommendation--limitation)
- [🎥 Video Explanation](#-video-explanation)
- [🏆 Advantages](#-advantages)
- [📚 References](#-references)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

**Delivery Delay Analysis** is a data analysis project that studies how much delivery delay each **service type, route, hub, and month** contributes. The same dataset is analysed independently in **Excel**, **SQL**, and **Python**, and the results are cross-checked so that every tool gives the same answer.

This project is designed to:
- Clean raw data by detecting and removing an exact duplicate record (13 rows → 12 rows)
- Join a fact table (`deliveries`) with a lookup table (`routes`) using `route_id`
- Create a derived field (`delay_days`) to measure delivery delay
- Compare service types, routes, hubs, and months using Excel formulas, SQL queries, and Python code
- Reconcile aggregate values across all three tools

> **Business Question:** *Which service type has the greatest delivery-delay burden, and which hub needs priority attention?*

---

## 🎯 Problem Statement

> **Objective:** Identify the service type with the heaviest delay burden and the hub that needs priority attention.

A delivery network runs four routes across three hubs. Each row is a **month-end route/hub delivery summary** with the promised days and the actual days taken. The analysis answers two business questions:

1. **Which service type (Express or Standard) has the greatest delivery-delay burden?**
2. **Which hub (Chennai, Delhi, or Mumbai) needs priority attention?**

| 📂 Dataset | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| `deliveries.csv` | Fact table | 13 rows (12 unique + 1 duplicate) of delivery records |
| `routes.csv` | Lookup table | 4 rows mapping each route to its service type |

The goal is to demonstrate **practical data analysis skills** using Excel, SQL, and Python on one consistent dataset.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧹 **Duplicate Removal** | Detects and removes 1 exact duplicate row (record 12) — 13 rows become 12 |
| 🔗 **Lookup Join** | Links deliveries to routes using `route_id` (XLOOKUP / JOIN / `merge`) |
| ⏱️ **Delay Field** | `delay_days = MAX(actual_days − promised_days, 0)` |
| 📊 **Excel PivotTable** | Total delay_days by service type × month (Jan → Feb → Mar) |
| 🗄️ **SQL Queries** | Three labelled queries (S2a, S2b, S2c) plus a LEFT JOIN integrity check |
| 🐍 **Python Pipeline** | Load → clean → merge → assert → derive → summarise → plot → export |
| ✅ **Validation Checks** | Row-count assertions and zero-unmatched-key checks in every tool |
| 🔁 **Reproducible** | Relative paths and run-in-order scripts that work on any computer |

---

## 🏗️ Project Structure

```
📦 data-analysis-set-a-YOUR-STUDENT-ID/
│
├── 📁 data/
│   └── 📁 raw/
│       ├── 📄 deliveries.csv         ← Raw fact file (13 rows incl. 1 duplicate)
│       └── 📄 routes.csv             ← Lookup file (4 rows)
│
├── 📁 excel/
│   └── 📗 analysis.xlsx              ← Raw, Lookup, Clean, Summary sheets
│
├── 📁 sql/
│   ├── 📄 setup.sql                  ← CREATE TABLE + INSERT (12 + 4 rows)
│   └── 📄 queries.sql                ← S2a, S2b, S2c + integrity check
│
├── 📁 python/
│   └── 🐍 analysis.ipynb             ← Cleaning, merge, derivation, chart, exports
│
├── 📁 outputs/
│   ├── 📄 clean_data.csv             ← Merged 12-row clean dataset
│   ├── 📄 python_summary.csv         ← Service-type summary
│   ├── 🖼️ python_chart.png           ← Monthly total delay_days chart
│   └── 📁 sql/                       ← Saved SQL query results
│       ├── 📄 s2a_delay_by_service_type.csv
│       ├── 📄 s2b_routes_delay_over_8.csv
│       ├── 📄 s2c_top_two_hubs.csv
│       └── 📄 s3_integrity_check.csv
│
├── 📄 requirements.txt               ← Python packages (pandas, matplotlib)
├── 📄 .gitignore                     ← Excludes environments, caches, credentials
└── 📄 README.md                      ← Project documentation
```

---

## 🔄 Project Workflow

```
Raw Data (deliveries.csv + routes.csv)
              │
              ▼
┌─────────────────────────────────┐
│  Verify Raw Files               │  ← 13 delivery rows, 4 route rows
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│  Remove Exact Duplicate         │  ← Record 12 appears twice → 12 rows remain
└────────────────┬────────────────┘
                 │
      ┌──────────┼──────────┐
      ▼          ▼          ▼
┌──────────┐ ┌──────────┐ ┌───────────┐
│  EXCEL   │ │   SQL    │ │  PYTHON   │
│ XLOOKUP  │ │ JOIN     │ │ merge()   │
│ MAX()    │ │ CASE WHEN│ │ clip()    │
│ SUMIFS   │ │ GROUP BY │ │ groupby   │
│ Pivot    │ │ HAVING   │ │ matplotlib│
└────┬─────┘ └────┬─────┘ └────┬──────┘
     │            │            │
     └────────────┼────────────┘
                  ▼
┌─────────────────────────────────┐
│  Cross-Tool Reconciliation      │  ← Same value in Excel, SQL & Python
└────────────────┬────────────────┘
                 │
                 ▼
     Findings + Recommendation ✅
```

---

## 🗂️ Dataset & Data Dictionary

### 📄 `deliveries.csv` — Fact Table

| Column | Type | Meaning |
|--------|------|---------|
| `record_id` | Integer | Unique record number |
| `month` | Text (ordered) | Month of the record — **Jan → Feb → Mar** |
| `route_id` | Text | Route key (R1–R4), links to `routes.csv` |
| `hub` | Text | Hub handling the delivery: Chennai, Delhi, or Mumbai |
| `promised_days` | Numeric | Days promised for delivery |
| `actual_days` | Numeric | Days actually taken |

### 📄 `routes.csv` — Lookup Table

| Column | Type | Meaning |
|--------|------|---------|
| `route_id` | Text | Unique route key (R1–R4) |
| `route` | Text | Route name |
| `service_type` | Text | Service type: Express or Standard |

### 🛣️ Route Mapping

| route_id | Route | Service Type |
|----------|-------|--------------|
| R1 | Metro Link | Express |
| R2 | City Dash | Express |
| R3 | Highway Freight | Standard |
| R4 | Rural Feeder | Standard |

> One `routes` row maps to many `deliveries` rows (**one-to-many** on `route_id`).

---

## 🧹 Data Cleaning & Metric Definitions

### 🧼 Cleaning Steps

| Step | Action | Result |
|------|--------|--------|
| 1️⃣ | Load raw `deliveries.csv` and `routes.csv` unchanged | 13 + 4 rows |
| 2️⃣ | Identify the exact duplicate (record `12, Mar, R4, Mumbai, 6, 15`) | 1 duplicate found |
| 3️⃣ | Remove the duplicate | **13 → 12 rows** |
| 4️⃣ | Set data types (`record_id` integer, days numeric, `month` text) | Types confirmed |
| 5️⃣ | Join deliveries to routes on `route_id` | 0 unmatched keys |
| 6️⃣ | Add `delay_days` | 34 total delay days |

### 📐 Metric Definitions

| Metric | Rule |
|--------|------|
| ⏱️ **delay_days** | `MAX(actual_days − promised_days, 0)` — zero when on time or early |
| 📉 **Delay Incidence Rate** | `Records where actual_days > promised_days ÷ All records × 100` |
| 📊 **Share of Total Delay** | `Group's summed delay_days ÷ Overall total delay_days × 100` |

> ⚠️ Rates are calculated from **underlying counts**, never by averaging subgroup percentages. Numeric results are shown to **two decimal places**. Summed `delay_days` represents **cumulative record-level delay**, not unique parcels to expedite today.

---

## 📗 Task 1 — Excel Analysis

### 📝 1. Workbook Guide — `excel/analysis.xlsx`

| Sheet | Purpose |
|-------|---------|
| 📄 **Raw** | Original 13-row `deliveries.csv`, pasted unchanged, with the before row count |
| 🔍 **Lookup** | The 4-row `routes.csv` |
| 🧹 **Clean** | 12 unique rows, `service_type` and `delay_days` columns, before (13) / after (12) counts |
| 📊 **Summary** | SUMIFS hub table, PivotTable, and column chart |

---

### 🔍 2. Service Type Lookup (XLOOKUP)

```excel
=XLOOKUP(C2, Lookup!A:A, Lookup!C:C)
```

*Pulls `service_type` from the Lookup sheet using `route_id`. `INDEX/MATCH` works equally well.*

---

### ⏱️ 3. Delay Days

```excel
=MAX(F2-E2, 0)
```

*Applied to all 12 rows of the Clean sheet (F = actual_days, E = promised_days).*

---

### 🔢 4. Total Delay by Hub (SUMIFS)
```

| Hub | Total delay_days |
|-----|------------------|
| 🏙️ Chennai | 36 |
| 🏙️ Delhi | 36 |
| 🏙️ Mumbai | 36 |
```
---

### 📊 5. PivotTable & Chart

- **Rows:** `service_type`  |  **Columns:** `month` (ordered Jan → Feb → Mar)  |  **Values:** Sum of `delay_days`

<img width="807" height="207" alt="Screenshot 2026-10-10 115811" src="https://github.com/user-attachments/assets/e34489a1-a69a-46a1-a7d6-15811ae41f2b" />

- A **column chart** with a title, axis labels, and legend is placed on the Summary sheet.

<img width="937" height="587" alt="Screenshot 2026-10-10 115859" src="https://github.com/user-attachments/assets/87516c39-ee94-4e53-bf31-b82fbeb87c93" />


> All formulas and the PivotTable stay **live and editable** — no screenshots-only results.

---

## 🗄️ Task 2 — SQL Analysis

### 📝 6. Setup — `sql/setup.sql`

- Defines `routes` and `deliveries` with primary keys and a **foreign key** `deliveries.route_id → routes.route_id`
- Loads exactly **4 route rows** and **12 delivery rows** (the duplicate is excluded)
- Starts with a comment stating the **SQL dialect and version**

```sql
-- SQL Dialect & Version: [Your SQL Engine & Version, e.g. MySQL 8.0]

CREATE TABLE routes (
    route_id     INT  PRIMARY KEY,
    route        TEXT NOT NULL,
    service_type TEXT NOT NULL
);

CREATE TABLE deliveries (
    record_id     INT         PRIMARY KEY,
    month         TEXT  NOT NULL,
    route_id      INT  NOT NULL,
    hub           TEXT NOT NULL,
    promised_days INT         NOT NULL,
    actual_days   INT         NOT NULL,
    FOREIGN KEY (route_id) REFERENCES routes(route_id)
);
```

---

### 📊 7. S2a — Total Delay Days by Service Type

```sql
SELECT r.service_type,
       SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days ELSE 0 END) AS total_delay_days
FROM deliveries d
JOIN routes r ON d.route_id = r.route_id
GROUP BY r.service_type
ORDER BY total_delay_days DESC;
```

**Result:**

| service_type | total_delay_days |
|--------------|------------------|
| Standard | 22 |
| Express | 12 |

---

### 🚨 8. S2b — Routes with Summed Delay Above 8

```sql
SELECT r.route,
       SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days ELSE 0 END) AS total_delay_days
FROM deliveries d
JOIN routes r ON d.route_id = r.route_id
GROUP BY r.route
HAVING SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days ELSE 0 END) > 8
ORDER BY total_delay_days DESC;
```

**Result:**

| route | total_delay_days |
|-------|------------------|
| Rural Feeder | 14 |
| Metro Link | 9 |

> Highway Freight totals exactly 8, so it is **not** above the threshold.

---

### 🏅 9. S2c — Top Two Hubs by Delay

```sql
SELECT d.hub,
       SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days ELSE 0 END) AS total_delay_days
FROM deliveries d
GROUP BY d.hub
ORDER BY total_delay_days DESC, d.hub ASC
LIMIT 2;
```

**Result:**

| hub | total_delay_days |
|-----|------------------|
| Mumbai | 15 |
| Delhi | 14 |

---

### 🔎 10. Data Integrity Check (LEFT JOIN)

```sql
SELECT COUNT(*) AS unmatched_keys
FROM deliveries d
LEFT JOIN routes r ON d.route_id = r.route_id
WHERE r.route_id IS NULL;
```

*Confirms every `route_id` in the fact table matches a lookup row — **`unmatched_keys = 0`**.*

---

### ▶️ 11. SQL Execution Order

| Step | File | Action |
|------|------|--------|
| 1️⃣ | `sql/setup.sql` | Create tables and load 4 + 12 rows |
| 2️⃣ | `sql/queries.sql` | Run S2a → S2b → S2c → integrity check |
| 3️⃣ | `outputs/sql/` | Save each result as a labelled CSV |

> **SQL Dialect & Version:** `[Fill in — e.g. MySQL 8.0 / SQL Server 2022]`

---

## 🐍 Task 3 — Python Analysis

### 📝 12. Setup & Run

```bash
pip install -r requirements.txt
jupyter notebook python/analysis.ipynb
```

*Open the notebook and use **Kernel → Restart & Run All**. The first cell detects the repository root automatically, so all paths stay relative and no edits are needed on another machine.*

---

### 📥 13. Load, Clean & Merge

```python
deliveries = pd.read_csv("data/raw/deliveries.csv")
routes = pd.read_csv("data/raw/routes.csv")

deliveries = deliveries.drop_duplicates()                 # 13 → 12 rows
df = deliveries.merge(routes, on="route_id", how="left")

assert len(df) == 12
assert df["service_type"].isna().sum() == 0               # no unmatched keys
```

---

### ⏱️ 14. Delay Field & Service-Type Summary

```python
df["delay_days"] = (df["actual_days"] - df["promised_days"]).clip(lower=0)

grp = df.groupby("service_type")
summary = pd.DataFrame({
    "total_delay_days": grp["delay_days"].sum(),
    "total_records": grp["record_id"].count(),
    "delayed_records": (df["actual_days"] > df["promised_days"]).groupby(df["service_type"]).sum(),
}).reset_index()
summary["delay_incidence_rate_pct"] = (summary["delayed_records"] / summary["total_records"] * 100).round(2)
```

**Service-Type Summary:**

| service_type | total_delay_days | total_records | delayed_records | delay_incidence_rate_pct |
|--------------|------------------|---------------|-----------------|--------------------------|
| Standard | 22 | 6 | 5 | 83.33 |
| Express | 12 | 6 | 4 | 66.67 |

**Route with the Greatest Delay:**

| Item | Value |
|------|-------|
| 🛣️ **Route** | Rural Feeder |
| ⏱️ **Summed delay_days** | 14 |
| 📊 **Overall total delay_days** | 34 |
| 📉 **Share of overall total** | **41.18%** |

---

### 📈 15. Chart & Exports

```python
monthly = df.groupby("month")["delay_days"].sum().reindex(["Jan", "Feb", "Mar"])

plt.bar(monthly.index, monthly.values)
plt.title("Monthly Total Delay Days")
plt.xlabel("Month")
plt.ylabel("Total Delay Days")
plt.savefig("outputs/python_chart.png", dpi=150)

df.to_csv("outputs/clean_data.csv", index=False)
summary.to_csv("outputs/python_summary.csv", index=False)
```

<img width="995" height="617" alt="Screenshot 2026-10-10 120102" src="https://github.com/user-attachments/assets/58f0ffc5-ee7a-4d96-9d2d-8cf9878e969c" />

| Month | Total delay_days |
|-------|------------------|
| Jan | 8 |
| Feb | 9 |
| Mar | 17 |

---

## 🔗 Cross-Tool Reconciliation

**Aggregate checked:** Total `delay_days` for the **Standard** service type.

| Tool | Method | Result |
|------|--------|--------|
| 📗 **Excel** | PivotTable — Standard Grand Total | 22 |
| 🗄️ **SQL** | S2a query | 22 |
| 🐍 **Python** | Service-type summary | 22 |

**Second check — overall total delay_days:** Excel PivotTable grand total = SQL (S2a: 22 + 12) = Python (`df["delay_days"].sum()`) = **34**.

> ✅ All three tools agree. Rounding note: `[Write "No rounding differences — all values are whole numbers" or add any difference you notice]`.

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | `[365 / 2019+]` | XLOOKUP, MAX, SUMIFS, PivotTable, charts |
| 🗄️ **SQL** | `[Engine & version]` | Table design, JOIN, GROUP BY, HAVING, LIMIT |
| 🐍 **Python** | `[3.8+]` | Core scripting language |
| 🐼 **pandas** | `[version]` | Data loading, cleaning, merging, aggregation |
| 📊 **matplotlib** | `[version]` | Charting |
| 🔀 **Git / GitHub** | Latest | Version control and submission |

---

## 📈 Results & Insights

After running the analysis, the following findings are produced:

- 🧹 **Clean Dataset** — 13 raw rows reduced to 12 unique records
- ⏱️ **Total Delay** — **34 delay days** across all 12 records (average 2.83 per record)
- 📦 **Service Type Burden** — **Standard** carries **22 delay days (64.71%)** with an **83.33%** delay incidence rate, versus **Express** with 12 delay days (35.29%) and 66.67%
- 🏙️ **Priority Hub** — **Mumbai** has the highest delay at **15 delay days (44.12%)**, followed closely by **Delhi** with 14 (41.18%); Chennai has 5 (14.71%)
- 🛣️ **Worst Route** — **Rural Feeder** has **14 delay days**, which is **41.18%** of the overall total
- 📅 **Monthly Trend** — Delay rises sharply: **8** (Jan), **9** (Feb), **17** (Mar), so March alone holds **50.00%** of all delay days
- 🚨 **Single Largest Delay** — Record 12 (Mar, Rural Feeder, Mumbai) took **9 days longer** than promised

---

## 💡 Recommendation & Limitation

### 🎯 Recommendation

Give **priority attention to the Mumbai hub** and to the **Standard service routes**, especially **Rural Feeder**, which causes 41.18% of all delay days. Review what happened in **March**, since delay doubled compared with February.

### ⚠️ Limitation

The dataset is **synthetic and very small** (12 month-end summary records). Summed `delay_days` is **cumulative record-level delay**, not the number of unique parcels to expedite today, so the findings show direction rather than statistical proof.

---

## 🎥 Video Explanation

| Item | Detail |
|------|--------|
| 🔗 **Video URL** | `[Paste your unlisted YouTube / Google Drive link here]` |
| ⏱️ **Duration** | `[mm:ss]` |
| 👁️ **Access** | Anyone with the link can view |

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | Applies Excel, SQL, and Python to one simple, easy-to-follow dataset |
| 🔁 **Reproducible** | Relative paths and ordered run steps work on any computer |
| ✅ **Validated** | Row-count assertions and zero-unmatched-key checks in each tool |
| 🔗 **Cross-Verified** | Key results match across Excel, SQL, and Python |
| 🧮 **Count-Based Rates** | Rates come from underlying counts, not averaged percentages |
| 🧩 **Modular** | Each tool lives in its own folder and can be run independently |
| 📖 **Well Documented** | Data dictionary, metric definitions, and run steps included |
| 🧪 **Extensible** | Easy to add more months, hubs, or routes |

---

## 📚 References

- `[List any external code, tutorials, or resources you used — or write "None"]`

> **Authorship Declaration:** *All work in this repository is my own except where cited.*

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Sneha Gupta

> *"Data tells the story — analysis helps us listen."*

**🎓 Role:** BBA Student | Data Analysis Learner\
**🆔 Student ID:** `12827` | **📝 Assigned Set:** Set A\
**📍 Location:** India\
**🛠️ Skills:** Excel · SQL · Python · Data Cleaning · Data Analysis

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 🏫 **Red & White Skill Education** — For the practical exam and dataset
- 📚 [Python Official Docs](https://docs.python.org/3/) — Official Python language reference
- 🐼 [pandas Documentation](https://pandas.pydata.org/docs/) — Data analysis library reference
- 📊 [matplotlib Documentation](https://matplotlib.org/stable/index.html) — Plotting reference
- 🗄️ [W3Schools SQL](https://www.w3schools.com/sql/) — Beginner SQL reference
- 📗 [Microsoft Excel Support](https://support.microsoft.com/excel) — Excel formula and PivotTable help
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: 10 October, 2026*

</div>

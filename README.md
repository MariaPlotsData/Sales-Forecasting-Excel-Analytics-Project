# 📊 Sales Pipeline Forecasting & Excel Analytics Project

**Course / Program:** Data Analytics Bootcamp — Excel & Prompt Engineering Module by Brave Analytics
* **Tools Used:** Microsoft Excel (`IFS`, `SUMIF`, `COUNTIF`, `COUNTA`, `LEFT`, `FIND`, `CONCATENATE`, Data Cleaning, Financial Forecasting)  
* **Company Context:** HLP Manufacturing INC. (2025 Sales Performance & Forecasting)

---

## 📌 Executive Summary

This project showcases a complete sales pipeline analysis and forecasting model for **HLP Manufacturing INC.** across 23 enterprise client accounts worldwide. 

Using raw sales data, advanced Excel formulas were applied to clean text fields, categorize regional territories, perform conditional accounting, evaluate sales target thresholds, and build a monthly revenue forecasting model.

---

## 🛠️ Key Technical Deliverables & Formula Solutions

### 1. Data Manipulation & String Extraction
* **Agent Full Name Concatenation:** Merged separate first and last name fields into a unified full name column.
  ```excel
  =CONCAT(C7," ",D7)
* **Regionalization Extraction** Extracted the primary country/region code from full territory strings (e.g., converting "US - Northeast" to "US" or "EMEA - France" to "EMEA").
  ```excel
  =LEFT(E7,FIND(" ",E7)-1)
### 2.Logical Categorization & Target Evaluation
* **Performance Remarks (Tiering)** Categorized deal sizes based on percentage contribution to the total pipeline ($3,659,150 target scope).
  ```excel
  =IFS(I7<0.04,"Low",I7<=0.045,"Medium",I7>0.045,"High")
* **Target Achievement Verification** Applied logical checks to evaluate whether forecast deal amounts met baseline quota requirements ($146,000 threshold).
  ```excel
  =IF(G7>=146000,"YES","NO")
### 3. Conditional Aggregation & Summary Analytics
* **Total Portfolio Records** Evaluated complete record entries across customer accounts.
  ```excel
  =COUNTA($B$7:$B$29)
* **Regional Sales Aggregation** Calculated total forecast amounts contributed per region (e.g, US).
  ```excel
  =SUMIF($L$7:$L$29,"US",$G$7:$G$29)
* **Sales Agent Deal Count** Calculated total deal volume managed by specific account executives (e.g., Rowan Murphy).
  ```excel
  =COUNTIF($N$7:$N$29,"Rowan Murphy")
* **Sales forecast per quarter** Calculated the sales forecast in quarterly basis (e.g., Q1)
  ```excel
  =SUMIF($H$7:$H$29,"January",$G$7:$G$29)+SUMIF($H$7:$H$29,"February",$G$7:$G$29)+SUMIF($H$7:$H$29,"March",$G$7:$G$29)

---


### 📈 Sales Performance Summary

---

| Metric / Dimension | Total Value | Key Findings |
| :--- | :--- | :--- |
| **Total Pipeline Amount** | **$3,659,150** | 23 Active Client Accounts |
| **Quota Attainment Rate** | **86.95% (20 / 23)** | 20 accounts exceeded the $145k deal threshold |
| **Top Performing Region** | **US ($1,246,950)** | Represents **34%** of total forecast revenue |
  
### Regional Revenue Contribution
  
* **US Region:** $1,246,950 (34% of total)
* **EMEA Region:** $1,001,500 (28% of total)
* **APSA Region:** $962,700 (26% of total)
* **Canada Region:** $448,000 (12% of total)

---

### 🎓 Learning Outcomes

---

 1. Mastered text manipulation (LEFT, FIND, CONCATENATE) for data cleaning.
 2. Implemented dynamic conditional logic (IFS, IF) for automatic data labelling.
 3. Created flexible summary reports using conditional statistical functions (SUMIF, COUNTIF, COUNTA).
 4. Structured a clean, executive-ready Excel workbook with clear data hierarchy and documentation.

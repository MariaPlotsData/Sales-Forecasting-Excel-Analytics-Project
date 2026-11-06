# 📊 Sales Pipeline Forecasting & Excel Analytics Project

**Author:** Maria Alia Aguila  
**Course / Program:** Data Analytics Bootcamp — Excel & Prompt Engineering Module  
**Tools Used:** Microsoft Excel (`IFS`, `SUMIF`, `COUNTIF`, `COUNTA`, `LEFT`, `FIND`, `CONCATENATE`, Data Cleaning, Financial Forecasting)  
**Company Context:** HLP Manufacturing INC. (2025 Sales Performance & Forecasting)

---

## 📌 Executive Summary

This project showcases a complete sales pipeline analysis and forecasting model for **HLP Manufacturing INC.** across 23 enterprise client accounts worldwide. 

Using raw sales data, advanced Excel formulas were applied to clean text fields, categorize regional territories, perform conditional accounting, evaluate sales target thresholds, and build a monthly revenue forecasting model.

---

## 🛠️ Key Technical Deliverables & Formula Solutions

### 1. Data Manipulation & String Extraction
* **Agent Full Name Concatenation:** Merged separate first and last name fields into a unified full name column.
  ```excel
  =CONCATENATE(B6, " ", C6)

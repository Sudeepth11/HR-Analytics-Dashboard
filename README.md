# 👥 HR Analytics Dashboard | Power BI Project

An interactive 3-page Power BI dashboard that turns raw employee records into a clear view of **headcount, performance rating, promotion and retrenchment risk**, all the way down to the individual employee ID. An Excel workbook with engineered columns and attrition KPIs supports the analysis.

![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?logo=powerbi&logoColor=black)
![DAX](https://img.shields.io/badge/DAX-Measures-blue)
![Excel](https://img.shields.io/badge/Excel-Formulas%20%26%20Summary-217346?logo=microsoftexcel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-HR%20Analytics-purple)

---

## 📌 Table of Contents
1. [Project Overview](#-project-overview)
2. [Business Questions](#-business-questions)
3. [Dataset](#-dataset)
4. [Calculated Columns (Excel)](#-calculated-columns-excel)
5. [KPI & Attrition Measures (Excel)](#-kpi--attrition-measures-excel)
6. [DAX Measures (Power BI)](#-dax-measures-power-bi)
7. [Dashboard Pages](#-dashboard-pages)
8. [Key Insights](#-key-insights)
9. [Recommendations](#-recommendations)
10. [Repository Structure](#-repository-structure)
11. [How to Use](#-how-to-use)
12. [Tools & Skills](#-tools--skills)
13. [Author](#-author)

---

## 🎯 Project Overview

HR leaders usually have to dig through spreadsheets to find who needs a promotion, who is at risk, and where attrition is coming from. This project gives them one report that answers those questions at a glance, plus an **Action page** listing the exact employees who need a decision this cycle.

**Headline numbers (Power BI report)**

| Total Employees | Male | Female | Due for Promotion | Due for Retrenchment | Active Workers |
|:-:|:-:|:-:|:-:|:-:|:-:|
| **1,470** | **882** | **588** | **72** | **117** | **1,233** |

---

## ❓ Business Questions

1. Which departments and job roles carry the most promotion vs. retrenchment risk?
2. How is the workforce performing, and where does the low/high rating split concentrate?
3. Which specific employees (by employee number) need action this cycle?
4. In the underlying HR dataset, what drives attrition and flight-risk scoring?

---

## 📂 Dataset

**File:** `HR_Attrition_Dashboard.xlsx`

| Sheet | Purpose |
|---|---|
| `Data` / `HRData` | 1,478 employee rows × 41 columns (raw fields + engineered columns) |
| `Summary` | Live Excel formulas for attrition KPIs and breakdowns |
| `Dashboard` | Excel-side workforce overview |

**Key columns**

| Column | Type | Description |
|---|---|---|
| EmpID | Text | Unique employee identifier |
| Age / AgeGroup | Number / Text | Age and derived age band |
| Attrition | Yes/No | Whether the employee has left |
| Department / JobRole / JobLevel | Text / Number | Organisational position |
| MonthlyIncome / AnnualIncome | Decimal | Monthly pay and derived yearly pay |
| OverTime | Yes/No | Regularly works overtime |
| JobSatisfaction / WorkLifeBalance | 1–4 | Self-reported ratings |
| YearsatCompany / YearsSincePromotion | Number | Tenure and time since last promotion |
| DistanceFromHome(KM) | Number | Commute distance |
| PerformanceRating | Number | Manager-assigned rating |

**Quick facts from the data**

- **1,478** employee rows · **41** columns
- **239** employees with Attrition = Yes (**16.2%**)
- Average monthly income: **₹49,514**

---

## 🧱 Calculated Columns (Excel)

Nine engineered columns, built with Excel formulas inside `HRData`:

| Column | Logic | Purpose |
|---|---|---|
| AnnualIncome | `MonthlyIncome × 12` | Yearly pay for salary banding |
| SalarySlab | Bucket of annual income | 0-3 / 3-6 / 6-10 / 10+ LPA |
| AgeGroup | Bucket of Age | 18-25 / 26-35 / 36-45 / 46-55 / 55+ |
| TenureBand | Bucket of YearsatCompany | 0-1 / 2-3 / 4-5 / 6-10 / 10+ yrs |
| AttritionFlag | `IF(Attrition="Yes",1,0)` | Numeric flag for summing/averaging |
| PromotionOverdue | `IF(YearsSincePromotion>=3,"Overdue","Recent")` | Flags long waits for advancement |
| IncomeVsRoleAvg | `(Income − role average) ÷ role average` | How far pay sits above/below the role's average |
| FlightRiskScore | OverTime=Yes + WorkLifeBalance≤2 + JobSatisfaction≤2 + YearsSincePromotion≥3 | Composite 0–4 retention-risk score |
| FlightRiskTier | Bucket of FlightRiskScore | Minimal / Low / Medium / High Risk |

---

## 📐 KPI & Attrition Measures (Excel)

Live formulas on the `Summary` sheet:

```excel
Total Employees      =COUNTA(Data!G2:G1479)
Employees Left       =COUNTIF(Data!D2:D1479,"Yes")
Attrition Rate       =Employees Left / Total Employees
Avg Monthly Income   =AVERAGE(Data!S2:S1479)
Avg Tenure (yrs)     =AVERAGE(Data!AE2:AE1479)
Avg Tenure - Left    =AVERAGEIF(Data!D2:D1479,"Yes",Data!AE2:AE1479)
Avg Tenure - Active  =AVERAGEIF(Data!D2:D1479,"No",Data!AE2:AE1479)
Attrition by Category =IFERROR(Left / Total, 0)
```

The attrition-rate-by-category pattern is reused for Department, OverTime, Age Group, Job Role, Work-Life Balance, Marital Status and Salary Slab.

---

## 🧮 DAX Measures (Power BI)

Stored in the `Dax Measure` table and used across all pages:

| Measure | What it shows |
|---|---|
| `Total_Employees` | Headcount on every KPI card |
| `Male` / `Female` | Gender split |
| `Due for Promotion` | Employees flagged ready for advancement |
| `Due for Retrenchment` | Employees flagged at risk of role reduction |
| `Not Due For Promotion` | Complement of Due for Promotion |
| `Active_Workers` | Currently active employees |
| `Low Performance %` / `High Rating` | Share of employees at each rating tier |

> ℹ️ DAX expressions are compiled inside the `.pbix` model. Open the file in Power BI Desktop to view the exact formulas.

---

## 📊 Dashboard Pages

Three pages navigated with the **Home / Details / Action** button menu.

### 🏠 Home — Headcount, satisfaction & action snapshot
- **Cards:** Total_Employees, Male, Female
- **Clustered column:** Due for Retrenchment / Promotion by Department
- **Column chart:** Total_Employees by Job Satisfaction Level
- **Pie chart:** Total_Employees by Over Time (71.7% No · 28.3% Yes)
- **Pivot table:** Job Role × headcount, promotion and retrenchment counts

### 🔍 Details — Tenure, seniority & commute profile
- **Cards:** Low / High Rating, Active Workers, Due for Promotion / Not Due / Due for Retrenchment
- **Bar chart:** Total_Employees by Years At Company
- **Column chart:** Total_Employees by Job Level
- **Donut chart:** Total_Employees by Distance From Office (63.95% Very Close)

### ✅ Action — Named employees needing a decision
- **Table:** Employee number list for **Due for Promotion** (72)
- **Table:** Employee number list for **Due for Retrenchment** (117)

**Job Role breakdown (Home page pivot)**

| Job Role | Employees | Due for Promotion | Due for Retrenchment |
|---|:-:|:-:|:-:|
| Healthcare Representative | 131 | 16 | 13 |
| Human Resources | 52 | – | 1 |
| Laboratory Technician | 259 | 3 | 5 |
| Manager | 102 | 22 | 44 |
| Manufacturing Director | 145 | 4 | 9 |
| Research Director | 80 | 8 | 20 |
| Research Scientist | 292 | 3 | 5 |
| Sales Executive | 326 | 16 | 20 |
| Sales Representative | 83 | – | – |
| **Total** | **1,470** | **72** | **117** |

### Screenshots
> Add exported screenshots of each page to an `images/` folder and they will show here.

![Home Page](images/home_page.png)
![Details Page](images/details_page.png)
![Action Page](images/action_page.png)

---

## 💡 Key Insights

1. **Retrenchment risk is concentrated, not spread evenly.** Manager is the riskiest role: of 102 managers, **44 are Due for Retrenchment and 22 Due for Promotion**, so about 65% of the role is flagged for action.
2. **Performance ratings skew heavily low.** **84.63%** of employees carry a Low rating vs. **15.37%** High, which points to a rating-scale or calibration issue rather than isolated cases.
3. **Overtime and attrition move together.** In the source dataset, employees working overtime leave at **31.0%** vs. **10.3%** for those who don't (419 vs. 1,059 employees).
4. **Younger employees leave most.** Attrition is **36.6%** for ages 18-25 vs. 9.1% for ages 36-45. Single employees leave at 21.6% vs. 11.4% for married.
5. **Commute distance is not the constraint.** 940 of 1,470 employees (**63.95%**) live Very Close to the office, so proximity is unlikely to explain the promotion/retrenchment split.

---

## ✅ Recommendations

- **Audit the Manager role first**, since it has the highest combined promotion and retrenchment rate
- **Review the rating scale**: an 85/15 low/high split is unlikely to be purely performance-driven
- **Cross-reference FlightRiskTier with the Action page** before finalising retrenchment decisions
- **Look at overtime workloads**, given the 3x attrition gap
- **Automate the Summary-sheet refresh** so attrition KPIs update as soon as the data does

---

## 🗂 Repository Structure

```
HR-Analytics-Dashboard/
│
├── HR_Analytics_Dashboard.pbix                # Power BI project (main deliverable)
├── HR_Attrition_Dashboard.xlsx                # Source data + calculated columns + KPI summary
├── HR_Analytics_Dashboard.pdf                 # Exported dashboard (3 pages)
├── HR_Analytics_Dashboard_Presentation.pptx   # Project walkthrough deck
├── images/                                    # Dashboard screenshots (optional)
└── README.md
```

---

## ▶️ How to Use

1. Clone or download this repository
2. Open `HR_Analytics_Dashboard.pbix` in **Power BI Desktop**
3. If prompted, point the data source to your local copy of `HR_Attrition_Dashboard.xlsx` (*Home → Transform data → Data source settings*)
4. Click **Refresh**, then use the **Home / Details / Action** buttons to move between pages
5. No Power BI Desktop? View `HR_Analytics_Dashboard.pdf` for a static copy

---

## 🛠 Tools & Skills

- **Power BI Desktop**: data modelling, report design, page navigation
- **DAX**: reusable measures for headcount, promotion and retrenchment
- **Microsoft Excel**: calculated columns, KPI formulas, attrition breakdowns
- **Skills demonstrated:** HR analytics, feature engineering, KPI design, dashboard storytelling, insight generation

---

## 👤 Author

**Sudeepth Sasikumar**
B.Tech Computer Science | Aspiring Data Analyst
📧 sudeepth203@gmail.com
🔗 [LinkedIn](https://www.linkedin.com/in/your-profile) · [GitHub](https://github.com/your-username)

⭐ If you found this project useful, consider giving it a star!

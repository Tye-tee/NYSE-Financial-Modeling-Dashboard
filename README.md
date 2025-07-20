# NYSE-Financial-Modeling-Dashboard

A dynamic financial model built in **Excel** to analyze and forecast financial performance for companies listed on the **New York Stock Exchange (NYSE)**. This model integrates advanced Excel functions (`INDEX`, `MATCH`, `OFFSET`) and interactive charts for scenario forecasting and risk analysis.

---
##  What's Inside
-  [**NYSE Data Analysis.xlsx**](https://1drv.ms/x/c/18593756cf7a242d/Eab9qFLd3jxFgSreaqO8YrIB04dXq1wC34I_5-e-FBVyOA?e=VpjZID) — The dynamic financial model (Cloud Access Link)
-  Dynamic P&L Dashboards by Ticker
-  Scenario Forecasting with base, strong, and weak cases
-  Sensitivity Analysis using revenue, gross, and operating margins

##  Features
- Drop-down driven dynamic dashboards
- Financial KPIs: Revenue, Cost, Gross Profit, EBIT
- Sector-based risk/volatility model using standard deviation
- Automated scenario forecasting model
- Clean, visual analytics using Excel charts

##  Tools & Functions Used
- Excel formulas: `INDEX`, `MATCH`, `OFFSET`, `FILTER`
- Charts: Bar
- Data validation for dynamic interactivity
- Standard deviation calculations for risk analysis

##  Preview
>  <img width="1300" height="567" alt="image" src="https://github.com/user-attachments/assets/83134ce9-5146-4444-a124-538cfe6877f0" />

## Problem & Business Impact

Financial analysts and investors often face challenges in assessing company performance and forecasting future growth due to fragmented data, lack of dynamic scenario modeling, and static reporting tools. Static models often fail to capture the uncertainty and variability in business outcomes, limiting the ability to make timely and strategic decisions. This project addresses that gap by developing a flexible forecasting solution for NYSE-listed companies:

- Users can select any company via its NYSE ticker symbol.
- The model calculates key financial & operational metrics using summary statistics
-  Uses historical data trends to forecast future outcomes
- It generates forecasts using three customizable scenarios: strong, weak, and base cases.
- Forecasting spans 1–2 years ahead, giving teams a forward-looking view of potential revenue paths
- Communicating findings through interactive dashboards and visual reports


## Analysis Workflow
Data Cleaning & Preparation

- Started with a pre-cleaned NYSE dataset provided by the Udacity Business Analytics team (original Kaggle dataset).

- Performed additional cleaning in Excel:

- Replaced missing values and removed special characters using Find and Replace.
  
- Standardized ticker symbols and applied Data Validation to eliminate duplicates.

- Selected relevant sub-industries for deeper analysis: Biotechnology, P&C Insurance, Life & Health, and Pharmaceuticals.
________________________________________
2. Exploratory Data Analysis (EDA)
- Generated a summary statistics table for Total Revenue grouped by industry and sub-industry:

- Included Sum, Average, Median, Standard Deviation, Minimum, Maximum, Variance, and Range.
  
- Used Pivot Tables and Filter functions to isolate and compare financial performance of the selected sub-industries.
  
- Created bar charts to visualize average Total Revenue over a 4-year period (2013–2016), highlighting key revenue trends and outliers by sector.
________________________________________
3. Profit & Loss (P&L) Dashboard – Dynamic
-	Built a dynamic P&L dashboard in Excel that updates based on selected ticker symbol via drop-down menu.

- Dashboard included dynamic calculations of: - Gross Profit - Operating Profit - Cost of Goods Sold (COGS)
- Used INDEX-MATCH, INDIRECT, and named ranges to ensure dynamic response to user-selected tickers.
________________________________________
4. Forecasting Model – Scenario Planning
-	Selected Celgene (Ticker: CELG) as the company for scenario-based financial forecasting.
-	Key business question: "How do we expect revenue growth and operating profit to change under different market conditions?"
-	Built a two-year forecast model with:
-	Three scenarios: Strong, Base, and Weak.
-	Scenario selection via drop-down list, dynamically changing all related assumptions.
- Applied OFFSET and MATCH functions to update Revenue Growth, Gross Margin, and Operating Margin across scenarios.
- Combined formulas with multiplication logic to dynamically model Gross Profit and Operating Profit outcomes under each case.



This solution enables financial teams to compare outcomes across multiple economic conditions, empowering product, finance, and strategy teams to make informed, data-driven decisions on how to optimize for growth, mitigate risk, or reallocate resources.

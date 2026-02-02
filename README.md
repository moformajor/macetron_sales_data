# Macetron Sales Dashboard

## Project Overview

The **Macetron Sales Dashboard** is an automated Excel-based analytics solution designed to help stakeholders gain clear, actionable insights into sales performance. The dashboard integrates data collection, cleaning, analysis, and visualization into a single, dynamic tool that supports fast and informed decision-making.

The solution is fully self-sufficient and requires no ongoing support from data professionals once deployed. Users can interact with the dashboard immediately to monitor performance and identify trends.

---

## Project Objectives

The primary objective of this project is to answer key business questions raised by the client, including:

- **Product Performance**
  - Which products generate the highest revenue?
  - Which products deliver the highest profit?
  - What costs are associated with top-performing products?

- **Financial Trends**
  - How do revenue, cost, and profit change over time?
  - Are there identifiable growth or decline patterns?

- **Order Status Analysis**
  - How do order statuses (Completed, Pending, Returned) evolve over time?
  - What operational or customer behavior patterns can be identified?

---


## Data Cleaning

### Standardizing Formats
Uniform date, time, and numeric formats were applied across all datasets to ensure consistency and compatibility during analysis.

### Validating Data
Data validation checks were performed to identify and correct errors such as invalid entries, improper spelling, and formatting issues.

### Removing Duplicates
Duplicate records were identified and removed to preserve data integrity.

### Handling Missing Values
Missing values were addressed using appropriate methods such as imputation, interpolation, or removal, depending on their impact on the analysis.

### Correcting Inconsistencies
Inconsistencies in spelling, capitalization, and data entry were corrected to improve data quality and reliability.

---

## Data Processing

### Merging Tables
Multiple datasets were combined using appropriate keys to create a unified and comprehensive view of the sales data.

### Creating Calculated Columns
Additional columns were created using formulas and logical operations to support deeper analysis, including profit and performance metrics.

### Applying VLOOKUP
Lookup functions were used to link related data across tables, enriching the dataset with relevant contextual information.

### Filtering and Sorting
Data was filtered and sorted to focus on relevant subsets, enabling targeted insights and faster analysis.

---

## Data Analysis

### Descriptive Statistics
A comprehensive statistical summary was conducted to uncover key patterns and trends in the dataset. This included assessing distributions, variability, and detecting outliers to ensure data consistency before further analysis.

### Hypothesis Testing (t-Test)
A two-sample t-test was performed to evaluate the statistical significance of differences between selected groups based on the client’s business questions. This strengthened the validity of insights drawn from the data.

---


## Data Collection

### Form Design
A structured Excel input form was designed to ensure accurate and consistent data entry. Writable fields were clearly defined, while non-input areas were locked to prevent accidental modifications.

### VBA Automation
A VBA macro was implemented to handle form submission. This automation captures user inputs and transfers them directly into the database, reducing manual effort and minimizing errors.

### Auto-Population & Data Storage
Calculated fields were automatically populated using predefined logic. Submitted data was stored securely in the database to maintain accuracy and traceability.

### Confirmation & Reset
After each submission, a confirmation pop-up is displayed to the user, and the form resets automatically to allow seamless entry of new records for future use.

---

## Dashboard Development

### KPI Creation Using Pivot Tables
Pivot tables were designed to calculate key performance indicators such as revenue, cost, profit, order volume, and order status distribution.

### Dashboard Framework Design
The dashboard layout and structure were planned in advance and then implemented in Excel to ensure clarity, usability, and effective data storytelling.

### Chart Development
Relevant charts were created to visually represent KPIs and insights, making trends and performance patterns easy to interpret.

### Interactive Slicers
Slicers were added to enable interactive filtering across time periods, products, cities, and order statuses.

### Final Review & Validation
A final quality check was conducted to ensure data accuracy, visual consistency, and full functionality of all dashboard components.

---

## Dashboard Features

- **Automated Data Processing**
  - End-to-end automation from data entry to visualization
  - Minimal manual intervention required

- **Dynamic Visualizations**
  - Interactive charts and pivot tables
  - Filters for time period, product category, and order status

- **Key Metrics Tracked**
  - Total Revenue
  - Total Cost
  - Profit
  - Order Volume
  - Order Status Distribution

- **User-Friendly Design**
  - Designed for non-technical users
  - Clear layout for quick decision-making

---

## Tools & Technologies

- **Microsoft Excel**
  - Power Query
  - VBA (Macros)
  - Pivot Tables & Charts
  - Excel formulas
  - Interactive slicers and dashboards

---

## Business Value

This dashboard enables Macetron to:

- Identify top-performing and underperforming products
- Monitor profitability and cost efficiency
- Track order fulfillment and operational trends
- Make informed, data-driven business decisions
- Reduce dependency on analysts for recurring reports

---

## How to Use

1. Open the Excel dashboard file.
2. Refresh the data (if connected to an updated source).
3. Use slicers and filters to explore insights.
4. Review KPIs and visual summaries for decision-making.

---

## Future Improvements (Optional)

- Integration with live sales databases
- Migration to Power BI for real-time reporting
- Predictive sales and demand forecasting
- Customer segmentation and advanced analytics

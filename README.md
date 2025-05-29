# A Dynamic Sales Management System In Excel

Welcome to my Excel-powered sales analytics project!

This project simulates a real-world retail sales operation where customer orders are recorded via a user-friendly sales form. From the moment a sale is entered, every part of the workbook updates automatically — from KPIs and sales tracking per retail outlet, to customer behavior analytics and delivery performance analysis. All of this is built in Microsoft Excel with automation via VBA (macros), dynamic formulas, pivot tables, and dashboards.

---

## Why This Project?

As a Certified Accounting Technician and aspiring data analyst, I developed this Excel-based solution to:

- Automate sales tracking across multiple branches and sales reps  
- Use Excel to deliver interactive business intelligence without third-party tools  
- Analyze delivery performance and customer satisfaction via order status metrics  
- Improve my ability to design integrated Excel systems for real-world business decisions  
- Showcase my ability to bridge financial acumen with data analysis for business impact  

---

## Tech Stack

| Tool                                           | Purpose                                                                      |
| ----------------------------------------------|-------------------------------------------------------------------------------|
| **Microsoft Excel (Macros Enabled)**           | Core platform for sales tracking, automation, KPI monitoring, and dashboards |
| **VBA (Sales Form)**                           | Automates data collection through a structured entry form                    |
| **Pivot Tables & Charts**                      | Used for trend analysis and interactivity                                    |
| **Excel Formulas (XLOOKUP, IF, SUMIFS, etc.)** | Drives dynamic KPIs and visualizations                                       |

---

## Dataset Overview

# Source of Data: 
The raw data used for this project was gotten from kaggle while a small number of data were generated as samples through the **Sales Form**, which serves as the entry point and updates all analytical worksheets.

### Fields Captured:

| Field Name               | Description                                  |
|--------------------------|----------------------------------------------|
| **Date**                 | Date of sale entry                           |
| **Customer Name**        | Name of the buyer                            |
| **Customer Type**        | Category (Individual, Retailer, Distributor) |
| **Product Category**     | e.g., Beverages, Toiletries                  |
| **Product Name**         | The specific product sold                    |
| **Unit Price**           | Price per unit sold                          |
| **Quantity Sold**        | Number of units purchased                    |
| **Branch**               | Location of the retail store                 |
| **Sales Representative** | Name of the employee that handled the sale   |
| **Delivery Date & Time** | When the order was delivered                 |
| **Order Status**         | Whether the order was Completed or Returned  |

---

## Project Breakdown

### Phase 1: Data Entry & Collection

- Sales data is entered via a VBA-powered **Sales Form**  
- Auto-populates a master sales table  
- Drop-downs for products, customer types, branches, and reps  
- All lists are editable via the **Settings Sheet**  

---

### Phase 2: KPI Dashboard (KPI Sheet)

| Metric                     | Description                                             |
|----------------------------|---------------------------------------------------------|
| **Total Revenue**          | Sum of all `Unit Price × Quantity`                      |
| **Quantity Sold**          | Total units sold                                        |
| **Sales Completion Rate**  | % of orders marked "Completed"                          |
| **Return Rate**            | % of returned orders                                    |
| **Top Product**            | Most sold product based on quantity                     |
| **Top Performing Branch**  | Branch with highest revenue                             |
| **Sales by Customer Type** | Segmented sales volume (e.g., Retailers vs Individuals) |
| **Avg Sales per Rep**      | Mean revenue per salesperson                            |

---

### Phase 3: Retail Branch Performance (Retail Store Sales Sheet)

| Metric                               | Insight                                       |
|--------------------------------------|-----------------------------------------------|
| **Total Revenue per Branch**         | Used to rank branch performance               |
| **Order Completion vs Return Count** | Helps assess service quality by location      |
| **Average Sale Value**               | Indicates customer spending behavior          |
| **Branch Filter**                    | Enables deep-dive into specific branch trends |

---

### Phase 4: Dashboard

The interactive Excel dashboard shows:

- **Revenue Trends Over Time**  
- **Sales by Product Category**  
- **Branch Comparison Charts**  
- **Top Products Sold**  
- **Customer Type Breakdown**  
- **Sales Rep Leaderboard**
- **Slicers to reflect the above KPIs based on major categories like Month, Country & product category.**

The entirety of the dashboard was built in Excel using pivot tables and slicers.

---

### Phase 5: Delivery Time vs Order Status (Analysis Sheet)

**Goal:** Determine if **delivery time** affects whether an order is **Completed or Returned**.

| Analysis Done                         | Description                                              |
|---------------------------------------|----------------------------------------------------------|
| **Delivery Duration Calculation**     | Based on time between sale and actual delivery           |
| **Average Delivery Time by Status**   | Compares Completed vs Returned orders                    |
| **Status Segmentation by Time Range** | Bins delivery times to explore correlations              |
| **Visual Analysis (Charts)**          | Bar/line graphs comparing delay groups with return rates |

## Using T-test:T-Test: Two-Sample Assuming Unequal Variances
- Developed null Hypothesis (H₀): Delivery time does not influence order status &
- Alternative Hypothesis (H₁): Orders with longer delivery times are more likely to be returned.

In overall, the descriptive analysis and t-test revealed that a quantitative relationship exists between delivery time and customer satisfaction, measured by whether the order is returned. As a result, actionable operational insights were obtained which would be implemented in order to reduce return rates and improve customer experience.

**Key Insight:** Longer delivery times tend to have higher return rates — highlighting areas for operational improvement.

---

## Settings Sheet – Source of Automation

All drop-downs and dynamic references in the form and formulas are controlled from this sheet:

| Component          | Use                                     |
|--------------------|------------------------------------------|
| **Product List**   | With associated categories and prices   |
| **Branches**       | Editable list of outlets                |
| **Sales Reps**     | Add/remove rep names                    |
| **Customer Types** | e.g., Individual, Retailer, Distributor |

---

## File Structure

| File                           | Description                                    |
|--------------------------------|------------------------------------------------|
| `sales_data_Akinfewa.xlsm`     | Main macro-enabled workbook                    |
| `README.md`                    | Project documentation                          |

---

## Key Highlights of the Project

- Fully automated data entry system with a macro-powered form  
- Real-time KPI dashboard & performance tracking  
- Retail store-level analytics  
- Delivery analysis to assess operational efficiency  
- Built with 100% Excel: no external software needed  

---

## Known Limitations

- No integration with external databases  
- Analysis limited to manually entered sales data  
- No real-time alerts or email automation (can be added with Power Automate)  

---

## 👤 Prepared By

**Otun Oluwapelumi Ayodele (AAT)**  
Certified Accounting Technician | Data Enthusiast  
📧 Email: [oluwapelumiotun@gmail.com](mailto:oluwapelumiotun@gmail.com)  
🕊️ Twitter: [@FiscalMindAcct](https://twitter.com/FiscalMindAcct)  
💼 LinkedIn: [oluwapelumiotun](https://linkedin.com/in/oluwapelumiotun) 

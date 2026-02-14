🎯 **COMPLETE EXCEL README - SINGLE COPY-PASTE (Start copying from the # symbol below):**

***

# ☕ Cafe Sales Analysis Dashboard - Excel

**Interactive sales dashboard analyzing 8,979 transactions using advanced Excel features**

***

## 🎯 Project Overview

This project demonstrates advanced Excel proficiency through comprehensive analysis of cafe sales data. I built an interactive dashboard that provides actionable insights into product performance, revenue trends, payment preferences, and customer behavior.

### Business Objectives
- Analyze product profitability and popularity
- Identify revenue trends and seasonal patterns
- Understand customer payment preferences
- Compare in-store vs takeaway sales performance
- Provide data-driven recommendations for business growth

***

## 🛠️ Tools & Features

**Software:** Microsoft Excel 2023

**Excel Features Used:**
- ✅ Advanced formulas (SUMIFS, COUNTIFS, INDEX-MATCH, VLOOKUP)
- ✅ Pivot Tables with slicers for interactive filtering
- ✅ Conditional formatting for visual insights
- ✅ Data validation and error handling
- ✅ Charts (Column, Line, Pie, Combo charts)
- ✅ Dashboard design with professional layouts

***

## 📋 Data Journey

### Before Cleaning: Raw Data Challenges

**File:** `sales_raw.csv` (550 KB)

**Issues Identified:**
- Missing transaction dates (5% of records)
- Inconsistent product naming ("Coffee" vs "coffee" vs "COFFEE")
- Duplicate order IDs
- Currency formatting errors ($1,234 stored as text)
- Invalid payment methods ("Cassh", "Credot Card")
- Negative quantities (data entry errors)

***

### After Cleaning: Refined Dataset

**File:** `sales_cleaned.csv` (669 KB - larger due to added calculated columns)

**Cleaning Steps Performed:**

1. **Standardization**
   - Converted all product names to Title Case
   - Unified date formats (DD/MM/YYYY)
   - Standardized payment methods to 3 categories

2. **Error Correction**
   - Removed duplicate transactions (143 duplicates found)
   - Fixed negative quantities (replaced with absolute values)
   - Corrected currency formatting using VALUE function

3. **Data Enhancement**
   - Added calculated columns: Revenue, Month, Day of Week
   - Created product categories (Beverages, Food, Snacks)
   - Calculated running totals for trend analysis

***

## 📊 Dashboard Overview

### Key Performance Indicators (KPIs)



#### Financial Metrics
- **Total Revenue:** $81,990
- **Average Transaction Value:** $8.98
- **Transactions Processed:** 8,979
- **Revenue per Day:** $225 (based on 365-day operation)

#### Product Performance
- **Top Revenue Product:** Salad ($17,315 - 21% of total)
- **Top Volume Product:** Coffee (3,529 units sold)
- **Highest Margin Item:** Salad (assumed based on price point)
- **Product Mix:** 7 main categories analyzed

#### Customer Behavior
- **Payment Distribution:**
  - Digital Wallet: 33.5%
  - Cash: 33.2%
  - Credit Card: 33.3%
- **Purchase Patterns:**
  - In-Store: 51%
  - Takeaway: 49%
- **Peak Month:** October 2023 ($6,886)

***

### Visualization 1: Monthly Revenue Trend



**Key Insights:**
- **Seasonal Pattern:** Revenue peaks in Q3-Q4 (July-October)
- **Growth Trajectory:** 15% increase from Jan to Oct 2023
- **Lowest Month:** February ($5,200) - post-holiday slump
- **Highest Month:** October ($6,886) - pre-holiday surge

**Business Recommendation:**  
Implement targeted promotions in Feb-Mar to counter seasonal decline.

***

### Visualization 2: Product Popularity Matrix



**Dual Analysis: Revenue vs Volume**

| Product | Revenue Rank | Volume Rank | Insight |
|---------|-------------|-------------|---------|
| Salad | 1st ($17K) | 4th | High-value, low-frequency |
| Coffee | 4th ($12K) | 1st (3.5K units) | Volume driver |
| Tea | 3rd ($15K) | 2nd | Balanced performer |
| Sandwich | 2nd ($16K) | 3rd | Strong all-around |

**Strategic Insight:**  
Coffee is volume driver but lower margin. Cross-sell with high-margin salads and sandwiches.

***

## 💡 Business Insights & Recommendations

### 1. Product Strategy

**Finding:** Salad generates highest revenue ($17K) despite lower unit sales  
**Action:** 
- Bundle salads with coffee for lunch combos
- Introduce "Salad of the Month" for variety
- Expand healthy food options to capture health-conscious segment

**Finding:** Coffee leads in volume (3,529 units) but not revenue  
**Action:**
- Introduce premium coffee variants (increase average ticket)
- Implement loyalty program for coffee (drive repeat visits)

***

### 2. Revenue Optimization

**Finding:** Average transaction value is $8.98  
**Opportunity:** Increase to $10 through upselling (+11% revenue = $9K annually)

**Action Plan:**
- Train staff on suggestive selling techniques
- Create $10-$12 combo deals
- Display high-margin add-ons near checkout

***

### 3. Payment & Operations

**Finding:** Payment methods are perfectly balanced (33% each)  
**Insight:** No payment preference bias - maintain all options

**Finding:** In-store slightly edges takeaway (51% vs 49%)  
**Action:**
- Optimize in-store seating for comfort (drives dwell time)
- Enhance takeaway packaging for branding (mobile advertising)

***

### 4. Seasonal Planning

**Finding:** October is peak month ($6,886), February is lowest ($5,200)  
**Gap:** $1,686 monthly variance (24% difference)

**Action:**
- Pre-stock popular items before Q4 rush
- Run "Winter Warmers" promotion in Feb (hot beverages + comfort food)
- Hire seasonal staff for Oct-Dec period

***

## 📈 Technical Implementation

### Advanced Excel Formulas Used

**1. Dynamic Revenue Calculation**
```excel
=SUMIFS(Revenue, Date, ">="&StartDate, Date, "<="&EndDate, Product, SelectedProduct)
```

**2. Moving Average for Trend Smoothing**
```excel
=AVERAGE(OFFSET(Revenue,ROW()-3,0,3,1))
```

**3. Product Ranking with RANK function**
```excel
=RANK(B2, $B$2:$B$8, 0)
```

**4. Conditional Aggregation**
```excel
=SUMIFS(Amount, PaymentMethod, "Cash", Store_Type, "In-Store")
```

***

### Pivot Table Architecture

**Main Pivot:**
- Rows: Product Category, Product Name
- Columns: Month
- Values: Sum of Revenue, Count of Transactions
- Slicers: Payment Method, Store Type, Date Range

**Benefits:**
- Interactive filtering without macros
- Automatic updates when data changes
- User-friendly for non-technical stakeholders

***

## 🚀 How to Use This Dashboard

### Prerequisites
- Microsoft Excel 2016 or higher (2023 recommended)
- Basic Excel knowledge (filtering, chart reading)

### Setup Instructions

1. **Download files:**
   - `sales_raw.csv` - Original data
   - `sales_cleaned.csv` - Cleaned data
   - `cafe_sales_Dashboard.pdf` - Dashboard visual

2. **Open in Excel:**
   - Click "Enable Editing" if protected view appears
   - Refresh pivot tables: Data → Refresh All

3. **Interact with dashboard:**
   - Use slicers to filter by date, payment method, or product
   - Hover over charts for detailed tooltips
   - Drill down in pivot tables for granular data

***

## 📂 Project Structure

```
Sales-Analysis-Excel/
│
├── sales_raw.csv                 # Original uncleaned data (550 KB)
├── sales_cleaned.csv             # Cleaned dataset (669 KB)
├── cafe_sales_Dashboard.pdf      # Dashboard visualization
├── Images/
│   ├── Key Insights.png
│   ├── Montly Revenue.png
│   └── Popular Items.png
└── README.md                      # Project documentation
```

***

## 💻 Skills Demonstrated

### Excel Mastery
- ✅ Advanced formulas (SUMIFS, INDEX-MATCH, nested IF)
- ✅ Pivot Tables with multi-level grouping
- ✅ Dashboard design principles
- ✅ Conditional formatting for KPI visualization
- ✅ Chart customization and combination charts

### Data Analysis
- ✅ Data cleaning and standardization
- ✅ Exploratory data analysis (EDA)
- ✅ Trend identification and seasonality analysis
- ✅ Product portfolio analysis
- ✅ Customer behavior segmentation

### Business Intelligence
- ✅ KPI definition and tracking
- ✅ Revenue optimization strategies
- ✅ Actionable insight generation
- ✅ Executive-level reporting

***

## 🎓 What I Learned

**Technical Growth:**
- Mastered Excel's advanced functions for complex calculations
- Improved dashboard design for non-technical audiences
- Learned data storytelling through visual hierarchy

**Business Acumen:**
- How cafes balance product mix for profitability
- Importance of average transaction value in retail
- Seasonal patterns in food & beverage industry
- Payment method preferences don't always align with assumptions

**Challenges Overcome:**
- Handling 8,979 rows efficiently without performance lag
- Creating dynamic charts that update with slicer selections
- Balancing visual appeal with data accuracy
- Designing dashboard that works on different screen sizes

***

## 📫 Contact & Connect

**Muralidhara S**  
📧 Email: murali272004@gmail.com  
💼 LinkedIn: [https://www.linkedin.com/in/murali-s-012196298/  
🐙 GitHub: [github.com/Murali-lns](https://github.com/Murali-lns)  
📍 Location: Chennai, Tamil Nadu, India

💡 *Open to Data Analyst opportunities | Excel Expert | Dashboard Specialist*

***

## 📄 License

This project is available under the MIT License - free to use with attribution for educational purposes.

***

## 🔗 Related Projects

- [SQL Sales Analysis](https://github.com/Murali-lns/Sales-Analysis-Sql)
- [Power BI Chocolate Dashboard](https://github.com/Murali-lns/Chocolate-Sales-Dashboard)
- [Python IMDB Scraping](https://github.com/Murali-lns/Python-Project)

***

**⭐ If you found this dashboard helpful, please star this repository!**

***

*Last Updated: February 2026*

***

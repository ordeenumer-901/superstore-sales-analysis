"""
Superstore Sales Analysis
=========================
Author: [Your Name]
Date: 2024

Business Question:
    We have 4 years of sales data. Where are we making money,
    where are we losing it, and what should we do differently?

This script:
    1. Loads and explores the data
    2. Loads it into SQLite for SQL analysis
    3. Answers 5 key business questions
    4. Exports results for the dashboard
"""

import pandas as pd
import sqlite3
import json
import os

# ─────────────────────────────────────────────
# STEP 1: LOAD & EXPLORE THE DATA
# ─────────────────────────────────────────────
# Why: Before any analysis, we need to understand what we're working with.
# We check shape, nulls, date range, and value distributions.

print("=" * 60)
print("STEP 1: Loading & Exploring the Data")
print("=" * 60)

df = pd.read_excel('data/superstore.xlsx', sheet_name='Cleaned Data Set')
df['Year'] = df['Order Date1'].dt.year
df['Month'] = df['Order Date1'].dt.to_period('M').astype(str)

print(f"Rows: {len(df):,}")
print(f"Columns: {df.shape[1]}")
print(f"Date Range: {df['Order Date1'].min().date()} to {df['Order Date1'].max().date()}")
print(f"Null Values: {df.isnull().sum().sum()} (zero = clean data)")
print(f"\nCategories: {df['Category1'].unique().tolist()}")
print(f"Regions: {df['Region1'].unique().tolist()}")
print(f"Segments: {df['Segment1'].unique().tolist()}")
print(f"\nSales Range: ${df['Sales'].min():,} – ${df['Sales'].max():,}")
print(f"Avg Sale: ${df['Sales'].mean():,.0f} | Median Sale: ${df['Sales'].median():,.0f}")
print(f"Max Discount: {df['Discount'].max():.0%}  ← flag: anything >100% needs review")


# ─────────────────────────────────────────────
# STEP 2: LOAD INTO SQLITE
# ─────────────────────────────────────────────
# Why: SQL is the language of data teams. Loading into SQLite lets us
# write proper queries against the data, just like a real analyst would
# with a company database.

print("\n" + "=" * 60)
print("STEP 2: Loading Data into SQLite")
print("=" * 60)

conn = sqlite3.connect('data/superstore.db')
df.to_sql('orders', conn, if_exists='replace', index=False)
print("✓ Data loaded into SQLite table: 'orders'")
print(f"  {len(df):,} rows written")


# ─────────────────────────────────────────────
# STEP 3: SQL ANALYSIS — 5 BUSINESS QUESTIONS
# ─────────────────────────────────────────────

print("\n" + "=" * 60)
print("STEP 3: SQL Analysis — Answering Business Questions")
print("=" * 60)

results = {}


# ── Q1: Revenue & Profit Trend by Year ──────
# Why: The first thing any stakeholder wants to know is "are we growing?"
# Year-over-year revenue and profit tells the big picture story.

q1 = """
SELECT
    Year,
    SUM(Sales)  AS Total_Revenue,
    SUM(Profit) AS Total_Profit,
    ROUND(CAST(SUM(Profit) AS FLOAT) / SUM(Sales) * 100, 1) AS Profit_Margin_Pct
FROM orders
GROUP BY Year
ORDER BY Year
"""
df_q1 = pd.read_sql(q1, conn)
results['yearly_trend'] = df_q1.to_dict(orient='records')
print("\nQ1: Revenue & Profit by Year")
print(df_q1.to_string(index=False))


# ── Q2: Profit by Category & Sub-Category ───
# Why: Not all products are equal. This tells us which product lines
# are actually driving the business vs. just generating revenue noise.

q2 = """
SELECT
    Category1       AS Category,
    Sub_Category1   AS Sub_Category,
    SUM(Sales)      AS Revenue,
    SUM(Profit)     AS Profit,
    ROUND(CAST(SUM(Profit) AS FLOAT) / SUM(Sales) * 100, 1) AS Margin_Pct,
    COUNT(*)        AS Order_Count
FROM orders
GROUP BY Category1, Sub_Category1
ORDER BY Profit DESC
"""
# SQLite uses column names as-is; use the actual column name
q2 = """
SELECT
    Category1         AS Category,
    "Sub-Category1"   AS Sub_Category,
    SUM(Sales)        AS Revenue,
    SUM(Profit)       AS Profit,
    ROUND(CAST(SUM(Profit) AS FLOAT) / SUM(Sales) * 100, 1) AS Margin_Pct,
    COUNT(*)          AS Order_Count
FROM orders
GROUP BY Category1, "Sub-Category1"
ORDER BY Profit DESC
"""
df_q2 = pd.read_sql(q2, conn)
results['category_profit'] = df_q2.to_dict(orient='records')
print("\nQ2: Profit by Category & Sub-Category")
print(df_q2.to_string(index=False))


# ── Q3: Regional Performance ─────────────────
# Why: Geography matters. If one region is consistently underperforming,
# that's a resource allocation and strategy conversation.

q3 = """
SELECT
    Region1   AS Region,
    SUM(Sales)  AS Revenue,
    SUM(Profit) AS Profit,
    ROUND(CAST(SUM(Profit) AS FLOAT) / SUM(Sales) * 100, 1) AS Margin_Pct,
    COUNT(DISTINCT "Customer Name1") AS Unique_Customers,
    COUNT(*) AS Orders
FROM orders
GROUP BY Region1
ORDER BY Profit DESC
"""
df_q3 = pd.read_sql(q3, conn)
results['regional'] = df_q3.to_dict(orient='records')
print("\nQ3: Regional Performance")
print(df_q3.to_string(index=False))


# ── Q4: Top 10 Customers by Profit ──────────
# Why: In most businesses, a small number of customers drive most revenue.
# Knowing who they are helps with retention and relationship strategy.

q4 = """
SELECT
    "Customer Name1" AS Customer,
    Segment1         AS Segment,
    COUNT(*)         AS Orders,
    SUM(Sales)       AS Revenue,
    SUM(Profit)      AS Profit,
    ROUND(CAST(SUM(Profit) AS FLOAT) / SUM(Sales) * 100, 1) AS Margin_Pct
FROM orders
GROUP BY "Customer Name1", Segment1
ORDER BY Profit DESC
LIMIT 10
"""
df_q4 = pd.read_sql(q4, conn)
results['top_customers'] = df_q4.to_dict(orient='records')
print("\nQ4: Top 10 Customers by Profit")
print(df_q4.to_string(index=False))


# ── Q5: Discount Impact on Margin ───────────
# Why: Discounts feel like a growth lever but often destroy margin.
# This query shows exactly what happens to profit at each discount tier.

q5 = """
SELECT
    CASE
        WHEN Discount = 0         THEN '0% - No Discount'
        WHEN Discount <= 0.10     THEN '1–10%'
        WHEN Discount <= 0.20     THEN '11–20%'
        WHEN Discount <= 0.30     THEN '21–30%'
        WHEN Discount <= 0.50     THEN '31–50%'
        ELSE '50%+'
    END AS Discount_Tier,
    COUNT(*)   AS Orders,
    SUM(Sales) AS Revenue,
    SUM(Profit) AS Profit,
    ROUND(CAST(SUM(Profit) AS FLOAT) / SUM(Sales) * 100, 1) AS Margin_Pct
FROM orders
GROUP BY Discount_Tier
ORDER BY MIN(Discount)
"""
df_q5 = pd.read_sql(q5, conn)
results['discount_impact'] = df_q5.to_dict(orient='records')
print("\nQ5: Discount Impact on Profit Margin")
print(df_q5.to_string(index=False))


# ── BONUS: Monthly trend for dashboard chart ─
q_monthly = """
SELECT
    Month,
    Year,
    SUM(Sales)  AS Revenue,
    SUM(Profit) AS Profit
FROM orders
GROUP BY Month, Year
ORDER BY Month
"""
df_monthly = pd.read_sql(q_monthly, conn)
results['monthly_trend'] = df_monthly.to_dict(orient='records')

# Segment breakdown
q_seg = """
SELECT Segment1 AS Segment, SUM(Sales) AS Revenue, SUM(Profit) AS Profit
FROM orders GROUP BY Segment1
"""
df_seg = pd.read_sql(q_seg, conn)
results['segment'] = df_seg.to_dict(orient='records')


# ─────────────────────────────────────────────
# STEP 4: EXPORT RESULTS
# ─────────────────────────────────────────────
# Why: We export the results as JSON so the dashboard can read them.
# This decouples the analysis from the visualization — a real-world pattern.

print("\n" + "=" * 60)
print("STEP 4: Exporting Results")
print("=" * 60)

with open('visuals/results.json', 'w') as f:
    json.dump(results, f, indent=2)
print("✓ Results exported to visuals/results.json")

conn.close()
print("\n✓ Analysis complete. Open visuals/dashboard.html to view the dashboard.")

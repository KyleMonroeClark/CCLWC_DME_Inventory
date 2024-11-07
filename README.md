# Lymphedema Clinic & DME Supplier Data Management

This project provides a data management solution for Central Coast Lymphedema and Wound Center, a clinic specializing in lymphedema treatment and DME (Durable Medical Equipment) supplies. The solution helps manage patient orders, inventory, and financial analysis, streamlining operations for compression garments, pneumatic devices, and wound care bandaging.

## Table of Contents
1. [Project Overview](#project-overview)
2. [Folder Structure](#folder-structure)
3. [SQL Database Structure](#sql-database-structure)
4. [SQL Queries](#sql-queries)
5. [Excel Macros & Formulas](#excel-macros--formulas)
6. [Usage Instructions](#usage-instructions)

## Project Overview

This project simulates operations at a lymphedema clinic and DME supplier, managing the ordering, inventory, and revenue of compression garments and other medical supplies. The system includes tools to track Medicare and Non-Medicare orders, update inventory, and analyze financial performance.

## Folder Structure

The project is organized into the following folders:

- **`Excel Workbook/`**: Contains a copy of the main Excel workbook with confidential patient orders removed.
- **`data/`**: Includes two CSV files:
  - `Inventory.csv`: Inventory sheet data.
  - `Orders.csv`: Order data with anonymized or made-up entries to maintain patient confidentiality.
- **`sql_queries/`**: Contains SQL queries for analyzing inventory and orders data.
- **`VBA_Macros/`**: Holds the `.bas` file with all VBA macros used for automating tasks.
- **`Images/`**: Screenshots showing project functionality (e.g., dropdown lists, pivot tables, macro-enabled sheets).
- **`README.md`**: This file.

## SQL Database Structure

The data has been organized into two main CSV files:

1. **Inventory.csv** – Holds product data for compression garments and related DME items.
2. **Orders.csv** – Contains order details, including patient information, part IDs, quantities, costs, and profits. For confidentiality, the data includes made-up entries.

## SQL Queries

The `sql_queries/` folder contains SQL scripts to analyze the data and provide useful business insights:

1. **Inventory Overview**: Summarizes inventory across various categories.
2. **Orders Summary (by Insurance Type)**: Shows a breakdown of orders by insurance type.
3. **Part ID Inventory Usage**: Tracks usage of each part ID, including remaining inventory levels.
4. **Inventory Ordering Suggestions**: Recommends parts to reorder based on current stock and optimal inventory levels.
5. **Profitability of Orders (by Date)**: Analyzes profitability over time, based on expected revenue and actual profits recorded.

## Excel Macros & Formulas

The Excel workbook includes complex macros and formulas to automate processes:

- **Patient Order Tracking (Medicare and Non-Medicare)**: 
  - Dropdown lists for Part IDs, with data populated through `VLOOKUP` and `IF` formulas to display HCPCS reimbursement, garment descriptions, and pricing details.
  - Macro buttons for "Clear Contents" and "Print and Record," which reset sheet data or log orders in the **Order List** and adjust inventory levels.
- **Order List Sheet**:
  - Logs patient details, part IDs, quantities, and financial metrics (expected revenue, recorded revenue, and actual profit).
  - Expected revenue calculation based on Medicare reimbursement (80% of HCPCS rates) or set prices for non-Medicare orders.
- **Inventory Sheet**:
  - Tracks each item’s Part ID, HCPCS code, Medicare pricing, category, body location, compression grade, and more.
  - Includes formulas to suggest reorder quantities based on minimum and optimal inventory levels.
- **Inventory Recording Sheet**:
  - Allows quick adjustments for new inventory received. Macro buttons include:
    - **Clear Contents**: Resets the sheet’s data fields.
    - **Record Inventory**: Updates inventory counts when new shipments arrive.

## Images

The `Images/` folder includes visual references to demonstrate key functionality:

1. **Dropdown List for Part ID**: Example from the Medicare Proof of Delivery form.
2. **Inventory Ordering Pivot Table**: Screenshot of the pivot table showing inventory reorder suggestions.
3. **Inventory Recording Table**: Display of the Inventory Recording sheet, including macro buttons for easy stock management.

## Usage Instructions

To explore this project:
1. **SQL Queries** – Import the data from the `data/` folder into a relational database and run the SQL scripts in `sql_queries/` to analyze the data.
2. **Excel Workbook** – Open the workbook in `Excel Workbook/` to view and test the automated ordering and inventory management processes. VBA macros are included for workflow automation, and images provide visual references.

This project offers a comprehensive solution for managing a medical clinic’s DME operations, combining data automation, inventory tracking, and profitability analysis.

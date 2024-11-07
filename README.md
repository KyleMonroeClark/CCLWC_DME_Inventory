# Lymphedema Clinic & DME Supplier Data Management

This project is a data management solution for Central Coast Lymphedema and Wound Center. It manages patient orders for compression garments, pneumatic compression devices, and wound care bandaging. The solution includes tools to track Medicare and Non-Medicare orders, manage inventory, and analyze financial metrics like expected revenue and profitability.

## Table of Contents
1. [Project Overview](#project-overview)
2. [File Structure](#file-structure)
3. [SQL Database Structure](#sql-database-structure)
4. [SQL Queries](#sql-queries)
5. [Excel Macros & Formulas](#excel-macros--formulas)
6. [Dashboard Insights](#dashboard-insights)
7. [Usage Instructions](#usage-instructions)

## Project Overview

This project simulates the operations of a Lymphedema clinic and Durable Medical Equipment (DME) supplier. The clinic provides compression garments, pneumatic compression devices, and wound care bandaging to patients. The project helps track orders, inventory, and revenue (for comissions)
## File Structure

The project is structured into the following files and directories:

- **`Excel Workbook/`**: A copy of the excel workbook, modified to delete all orders for patient confidentiality.
- **`data/`**: Contains raw .csv files for SQL queries. Data gotten from the inventory sheet in the Excel file and made up data to protect patient information in the orders sheet.
- **`sql_queries/`**: Contains SQL queries to analyze the data stored in a relational database.
- **`VBA Macros & Formulas/`**: Contains code for VBA modules, and examples of some of the complex formulas used for this project.
- **`Images/`**: Contains images (mostly screenshots) meant to give an idea of how the project works.
- **`README.md`**: This file.

## SQL Database Structure

The data is organized into two main .csv files:
1. **Inventory**: Stores information about the products (compression garments, pneumatic devices, etc.).
2. **Orders**: Contains all the order details, including patient information, part IDs, quantities, costs, and profits. For patient confidentiality, I have created fake entries to fill out this field.

## SQL Queries

This section contains SQL queries designed to analyze the data and provide buisness insights. The following queries are included:

1. **Inventory Overview**: Summarizes inventory across categories.
2. **Orders Summary (by Insurance Type)**: Gives a breakdown of orders by insurance type.
3. **Part ID Inventory Usage**: Tracks how much of each part ID has been ordered and what's left in inventory.
4. **Inventory Ordering Suggestions**: Recommends parts to reorder based on current stock levels and optimal inventory.
5. **Profitability of Orders (by Date)**: Analyzes the profitability of orders over time.

## Excel Macros & Formulas

The Excel workbooks used in this project include complex macros and formulas that automate tasks such as:
- Calculating expected revenue based on Medicare reimbursement rates.
- Updating inventory counts when orders are processed.
- Adding new orders to the order list and updating the inventory.

## Images
Includes the following images:
- Image of drop down list of Part ID's for the Medicare Proof of Delivery
- Image of pivot table used primary to track which garments to order to replenish inventory (including macro enabled buttons).
- Image of table used to track inventory that we have recieved, before stocking (including macro enabled buttons).
- PDF of example Proof of Delivery hiding

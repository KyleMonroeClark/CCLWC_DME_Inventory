Top_Ordered_Parts.sql
Queries the Inventory and Orders tables to find the most ordered parts, including details such as manufacturer, model, and total quantity ordered. Only parts ordered more than once are displayed, sorted by highest total.

Monthly_Order_Summary.sql
Groups part order quantities by month, showing each part’s monthly total from the Orders table.

Profit_Summary.sql
Provides a monthly summary of expected versus actual profit based on orders. Uses the Date of Service column to group data by month.

Demographics_and_Orders.sql
Returns mailing addresses of patients with orders over six months old, along with a breakdown of items ordered. Includes part IDs, quantity, and a description for each item ordered.

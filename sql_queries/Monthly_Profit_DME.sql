SELECT 
    DATE_FORMAT(Date_of_Service, '%Y-%m') AS Order_Month,
    SUM(Expected_Revenue) AS Total_Expected_Profit,
    SUM(Actual_Profit) AS Total_Actual_Profit
FROM Orders
GROUP BY Order_Month
ORDER BY Order_Month;

SELECT 
    i.Part_ID,
    i.Manufacturer,
    i.Model,
    i.Compression_Grade,
    i.Size,
    i.Color,
    i.Style,
    DATE_FORMAT(o.Date_of_Service, '%Y-%m') AS Order_Month,
    SUM(
        CASE 
            WHEN o.Part_ID1 = i.Part_ID THEN o.Quantity1 / i.Units_Per_Package
            WHEN o.Part_ID2 = i.Part_ID THEN o.Quantity2 / i.Units_Per_Package
            WHEN o.Part_ID3 = i.Part_ID THEN o.Quantity3 / i.Units_Per_Package
            WHEN o.Part_ID4 = i.Part_ID THEN o.Quantity4 / i.Units_Per_Package
            ELSE 0
        END
    ) AS Total_Quantity_Ordered
FROM Inventory i
LEFT JOIN Orders o
    ON i.Part_ID IN (o.Part_ID1, o.Part_ID2, o.Part_ID3, o.Part_ID4)
GROUP BY i.Part_ID, Order_Month, i.Manufacturer, i.Model, i.Compression_Grade, i.Size, i.Color, i.Style
HAVING Total_Quantity_Ordered > 1
ORDER BY Order_Month, Total_Quantity_Ordered DESC;

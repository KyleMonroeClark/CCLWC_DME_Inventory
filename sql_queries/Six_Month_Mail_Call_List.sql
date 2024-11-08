SELECT 
    o.Patient_ID,
    p.Mailing_Address,
    p.City,
    p.State,
    p.Zip,
    o.Date_of_Service,
    CASE 
        WHEN o.Part_ID1 IS NOT NULL THEN o.Part_ID1
        WHEN o.Part_ID2 IS NOT NULL THEN o.Part_ID2
        WHEN o.Part_ID3 IS NOT NULL THEN o.Part_ID3
        WHEN o.Part_ID4 IS NOT NULL THEN o.Part_ID4
    END AS Part_ID,
    CASE 
        WHEN o.Part_ID1 IS NOT NULL THEN o.Quantity1
        WHEN o.Part_ID2 IS NOT NULL THEN o.Quantity2
        WHEN o.Part_ID3 IS NOT NULL THEN o.Quantity3
        WHEN o.Part_ID4 IS NOT NULL THEN o.Quantity4
    END AS Quantity,
    i.Description
FROM Orders o
JOIN Patients p ON o.Patient_ID = p.Patient_ID
JOIN Inventory i ON i.Part_ID IN (o.Part_ID1, o.Part_ID2, o.Part_ID3, o.Part_ID4)
WHERE DATEDIFF(CURDATE(), o.Date_of_Service) >= 180
ORDER BY o.Date_of_Service;

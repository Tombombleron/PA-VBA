SELECT [Employee ID],
        FirstName & " " & LastName as FullName,
        CostCentreCode,
        Department,
Band
FROM ((ee_Employees_t a
LEFT JOIN ee_IDCostCentre b
ON (a.IDCostCentre = b.IDCostCentre))
LEFT JOIN DeptIDs c
ON (a.IDDept = c.TypIDdept))
LEFT JOIN BandIDs d
ON (a.IDBand = d.IDBand)
WHERE a.LastName IN ('NAME_HERE');

SELECT  a.[Employee ID] AS EmployeeID,
        a.FirstName & " " & a.LastName AS FullName,
        Nz(a.MySingle, '[MySingle not found]') AS MySingle,
        NZ(a.EmailAdd, '[email not found]') AS EmailAddress,
        a.IsActive,
        e.CostCentreCode AS CostCentre,
        Nz(f.Department, '[department not found]') AS Department,
        b.BusinessArea,
        c.LocationName,
        d.[Band]
FROM ((((ee_EMPLOYEES_t AS a
LEFT JOIN ee_IDBusinessArea AS b
ON a.IDBusinessArea = b.IDBusinessArea)
LEFT JOIN ee_IDLocation AS c
ON a.IDLocation = c.IDLocation)
LEFT JOIN ee_IDBand AS d
ON a.IDBand = d.IDBand)
LEFT JOIN ee_IDCostCentre AS e
ON a.IDCostCentre = e.IDCostCentre)
LEFT JOIN ee_IDDept AS f
ON a.IDDept = f.IDDept
WHERE (((a.[Employee ID]) Like "*" & Forms!cr_Cards_Search_f.empnoParam & "*")
AND (([FirstName] & " " & [LastName]) Like "*" & Forms!cr_Cards_Search_f.empnaParam & "*"))
ORDER BY a.IsActive, b.BusinessArea DESC , d.[Band], c.LocationName;

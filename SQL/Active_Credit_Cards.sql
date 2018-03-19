SELECT a.[Card Number],
        b.[Explanation] AS Status,
        Trim(Nz(d.[FirstName]) & "  " & Nz(d.[LastName], '[name not found]')) AS [Cardholder Name],
        Nz(d.[EmailAdd], '[email not found]') AS [Email Address],
        Nz(e.[CostCentreCode], '[cost centre not found]') AS CostCentre,
        c.Currency
FROM (((cr_CREDITCARDS_t AS a
LEFT JOIN ee_EMPLOYEES_t AS d
ON d.[Employee ID] = a.EmployeeID)
LEFT JOIN TypIDstat AS b
ON b.IDStatus = a.IDStatus)
LEFT JOIN cr_IDCurrency AS c
ON c.IDCurrency = a.IDCurrency)
LEFT JOIN ee_IDCostCentre AS e
ON e.IDCostCentre=d.IDCostCentre
WHERE b.Explanation NOT IN ('Closed', 'Replaced As Lost or Stolen', 'Replaced By Fraud')
ORDER BY c.Currency DESC , b.Explanation;

SELECT DISTINCT ( b.[employee Id]) AS EmployeeID, b.[FirstName] & " " & b.[LastName] AS FullName,
                  b.[EmailAdd] AS EmailAddress,
                  '*' & RIGHT(c.[Card Number], 6) AS CardNumber,
                  a.DateDeducted,
                  a.AmountDeducted,
                  b.IsActive
FROM (cr_SalaryDeduction_History_q AS a
LEFT JOIN ee_EMPLOYEES_t AS b
ON b.[Employee ID] = a.EmployeeID)
LEFT JOIN cr_CREDITCARDS_t AS c ON
b.[Employee ID]  = c.EmployeeID
WHERE c.[IDStatus] NOT IN ('6', '5', '4', '3')
ORDER BY a.DateDeducted DESC , b.IsActive;

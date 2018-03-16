SELECT a.[Card Number], '*' & RIGHT(a.[Card Number], 4) AS LastFourDigits,
        '*' & RIGHT(a.[Card Number], 6) AS LastSixDigits,
        Nz(b.[FirstName], '[name not found]') & " " & b.[LastName] AS FullName,
        Nz(b.[EmailAdd], '[email not found]') AS EmailAddress,
        b.IsActive,
        c.Explanation,
        e.CardLimit,
        d.Currency
FROM (((cr_CREDITCARDS_t AS a
LEFT JOIN ee_EMPLOYEES_t AS b
ON a.[EmployeeID] = b.[Employee ID])
LEFT JOIN TypIDstat AS c
ON c.IDStatus = a.IDStatus)
LEFT JOIN cr_IDCurrency AS d
ON d.IDCurrency = a.IDCurrency)
LEFT JOIN cr_IDCardLimit AS e
ON b.IDBand = e.IDBand
WHERE ((([FirstName] & " " & [Lastname]) Like "*" & Forms!cr_Cards_Search_f.empParam & "*")
And (a.[Card Number] Like '*' & Forms!cr_Cards_Search_f.empParam1 & '*'))
ORDER BY b.IsActive, c.Explanation DESC , a.IDCurrency, b.FirstName;

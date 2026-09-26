SELECT
  Sum([Quantity] * [UnitPrice]) AS OrderTotal,
  Employees.FullNameFNLN AS Expr1
FROM
  (
    Employees
    INNER JOIN Orders ON Employees.EmployeeID = Orders.EmployeeID
  )
  INNER JOIN OrderDetails ON Orders.OrderID = OrderDetails.OrderID
WHERE
  (
    (
      (
        Year([OrderDate])
      ) = Year(
        Date()
      )
    )
  )
GROUP BY
  Employees.FullNameFNLN;

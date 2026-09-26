SELECT
  Orders.EmployeeID,
  Sum(qryOrderTotal.OrderTotal) AS OrderTotal,
  [Employees].[FirstName] & " " & [Employees].[LastName] AS FullNameFNLN,
  Format([OrderDate], "mmm-yyyy") AS MonthYear,
  Format([OrderDate], "yyyy-mm") AS MonthYearSort
FROM
  Employees
  INNER JOIN (
    qryOrderTotal
    INNER JOIN Orders ON qryOrderTotal.OrderID = Orders.OrderID
  ) ON Employees.EmployeeID = Orders.EmployeeID
WHERE
  (
    (
      (Orders.OrderDate) Between reportParameterStartDate()
      And reportParameterEndDate()
    )
  )
GROUP BY
  Orders.EmployeeID,
  [Employees].[FirstName] & " " & [Employees].[LastName],
  Format([OrderDate], "mmm-yyyy"),
  Format([OrderDate], "yyyy-mm");

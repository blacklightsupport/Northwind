SELECT
  Employees.EmployeeID,
  [Employees].[FirstName] & " " & [Employees].[LastName] AS FullNameFNLN,
  Employees.EmailAddress
FROM
  Employees
ORDER BY
  [Employees].[FirstName] & " " & [Employees].[LastName];

SELECT
  Employees.EmployeeID,
  [employees].[FirstName] & " " & [employees].[LastName] AS FullNameFNLN,
  Employees.PrimaryPhone,
  Employees.SecondaryPhone,
  Left([FirstName], 1) AS EmailGroup
FROM
  Employees
ORDER BY
  [employees].[FirstName] & " " & [employees].[LastName];

SELECT
  EmployeePrivileges.*
FROM
  Employees
  INNER JOIN EmployeePrivileges ON Employees.EmployeeID = EmployeePrivileges.EmployeeID
ORDER BY
  Employees.FirstName,
  Employees.LastName;

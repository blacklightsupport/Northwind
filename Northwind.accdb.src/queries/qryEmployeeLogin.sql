SELECT
  qryEmployees.EmployeeID,
  qryEmployees.FirstName,
  qryEmployees.LastName,
  qryEmployees.EmailAddress,
  qryEmployees.JobTitle,
  qryEmployees.WindowsUserName,
  qryEmployees.FullNameFNLN AS FullName
FROM
  qryEmployees;

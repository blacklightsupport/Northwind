SELECT
  qryEmployees.EmployeeID,
  qryEmployees.fullnamefnln AS Supervisor
FROM
  qryEmployees
WHERE
  (
    (
      (qryEmployees.EmployeeID) <> [Forms]![frmEmployeeList]![EmployeeID]
    )
  );

SELECT
  qryOrderList.*
FROM
  qryOrderList
WHERE
  (
    (
      (qryOrderList.EmployeeID) = [Parent].[EmployeeID]
    )
  )
ORDER BY
  qryOrderList.OrderDate DESC;

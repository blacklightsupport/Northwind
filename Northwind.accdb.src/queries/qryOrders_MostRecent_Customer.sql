SELECT
  qryOrderList.*
FROM
  qryOrderList
WHERE
  (
    (
      (qryOrderList.CustomerID) = [Parent]![Parent]![CustomerID]
    )
  )
ORDER BY
  qryOrderList.OrderDate DESC;

SELECT
  MRU.*
FROM
  MRU
WHERE
  (
    (
      (MRU.EmployeeID) = Get_UserID()
    )
  )
ORDER BY
  MRU.MRU_ID;

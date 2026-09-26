SELECT
  Companies.CompanyID,
  Companies.CompanyName
FROM
  Companies
WHERE
  (
    (
      (Companies.CompanyTypeID) = 1
    )
  )
ORDER BY
  Companies.CompanyName;

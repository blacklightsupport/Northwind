SELECT
  Companies.CompanyID,
  Companies.CompanyName
FROM
  Companies
WHERE
  (
    (
      (Companies.CompanyTypeID) = 2
    )
  )
ORDER BY
  Companies.CompanyName;

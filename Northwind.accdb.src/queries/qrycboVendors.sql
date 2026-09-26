SELECT
  Companies.CompanyID,
  Companies.CompanyName
FROM
  Companies
WHERE
  (
    (
      (Companies.CompanyTypeID) = 3
    )
  )
ORDER BY
  Companies.CompanyName;

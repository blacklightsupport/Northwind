SELECT
  Companies.*,
  [Address] & Space(2)& [City] & ", " & [StateAbbrev] & Space(2)& [Zip] AS BusinessAddress
FROM
  Companies
ORDER BY
  Companies.CompanyName;

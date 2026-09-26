SELECT
  Companies.CompanyID,
  Companies.CompanyName,
  CompanyTypes.CompanyType,
  Companies.BusinessPhone,
  Companies.Address,
  Companies.City,
  Companies.StateAbbrev,
  Companies.Zip,
  Companies.Website,
  Companies.Notes,
  TaxStatus.TaxStatus,
  Companies.AddedBy,
  Companies.AddedOn,
  Companies.ModifiedBy,
  Companies.ModifiedOn,
  [Address] & Space(2)& [City] & ", " & [StateAbbrev] & Space(2)& [Zip] AS BusinessAddress,
  Companies.CompanyTypeID,
  Companies.StandardTaxStatusID
FROM
  TaxStatus
  INNER JOIN (
    CompanyTypes
    INNER JOIN Companies ON CompanyTypes.CompanyTypeID = Companies.CompanyTypeID
  ) ON TaxStatus.TaxStatusID = Companies.StandardTaxStatusID
ORDER BY
  Companies.CompanyName;

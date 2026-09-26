SELECT
  Products.ProductID,
  Products.ProductName,
  Products.StandardUnitCost
FROM
  Products
  INNER JOIN ProductVendors ON Products.ProductID = ProductVendors.ProductID
WHERE
  (
    (
      (ProductVendors.VendorID) = [Parent]![VendorID]
    )
  )
ORDER BY
  Products.ProductName;

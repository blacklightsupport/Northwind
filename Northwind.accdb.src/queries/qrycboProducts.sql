SELECT
  Products.ProductID,
  Products.ProductName,
  Products.UnitPrice,
  Products.ProductCategoryID
FROM
  Products
WHERE
  (
    (
      (Products.ProductCategoryID) = [Form]![cboProductCategories]
    )
  )
ORDER BY
  Products.ProductName;

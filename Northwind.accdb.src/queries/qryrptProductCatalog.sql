SELECT
  ProductCategories.ProductCategoryID,
  ProductCategories.ProductCategoryName,
  ProductCategories.ProductCategoryCode,
  Products.ProductName,
  Products.ProductID,
  Products.QuantityPerUnit,
  Products.UnitPrice,
  qryTotalSalesByProduct.TotalSales,
  ProductCategories.ProductCategoryDesc,
  ProductCategories.ProductCategoryImage
FROM
  ProductCategories
  INNER JOIN (
    Products
    LEFT JOIN qryTotalSalesByProduct ON Products.ProductID = qryTotalSalesByProduct.ProductID
  ) ON ProductCategories.ProductCategoryID = Products.ProductCategoryID;

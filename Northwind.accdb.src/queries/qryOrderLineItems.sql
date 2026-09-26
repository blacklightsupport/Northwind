SELECT
  OrderDetails.*,
  Products.ProductName,
  Products.ProductCategoryID,
  ProductCategories.ProductCategoryName
FROM
  (
    ProductCategories
    INNER JOIN Products ON ProductCategories.ProductCategoryID = Products.ProductCategoryID
  )
  INNER JOIN OrderDetails ON Products.ProductID = OrderDetails.ProductID
ORDER BY
  OrderDetails.OrderDetailID;

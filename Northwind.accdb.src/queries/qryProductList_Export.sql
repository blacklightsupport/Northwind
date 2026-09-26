SELECT
  Products.ProductID AS [Product ID],
  Products.ProductCode AS [Product Code],
  Products.ProductName AS Product,
  CDbl(
    ProductAllocated([Products].[ProductID])
  ) AS [Allocated Inventory],
  ProductToSell([Products].[ProductID]) AS [Inventory To Sell],
  ProductOnOrder([Products].[ProductID]) AS [Qty On Order],
  Products.ReorderLevel AS [Reorder Level],
  Products.TargetLevel AS [Target Level],
  Products.MinimumReorderQuantity AS [Min Reorder Qty],
  Products.Discontinued
FROM
  Products
ORDER BY
  Products.ProductName;

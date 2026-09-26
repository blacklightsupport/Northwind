SELECT
  Products.ProductID,
  ProductCategories.ProductCategoryName,
  Products.ProductCode,
  Products.ProductName,
  Products.ProductDescription,
  Products.StandardUnitCost,
  ProductNoStock([Products].[ProductID]) AS NoStock,
  ProductAllocated([Products].[ProductID]) AS Allocated,
  ProductToSell([Products].[ProductID]) AS ToSell,
  ProductOnOrder([Products].[ProductID]) AS QuantityOnOrder,
  Products.MinimumReorderQuantity,
  Products.UnitPrice,
  Products.ReorderLevel,
  Products.TargetLevel,
  Products.QuantityPerUnit,
  Products.Discontinued,
  Products.ProductCategoryID,
  Products.AddedBy,
  Products.AddedOn,
  Products.ModifiedBy,
  Products.ModifiedOn
FROM
  ProductCategories
  INNER JOIN Products ON ProductCategories.ProductCategoryID = Products.ProductCategoryID;

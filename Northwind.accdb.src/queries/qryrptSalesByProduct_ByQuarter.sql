SELECT
  Products.ProductName,
  Products.ProductID,
  Format([OrderDate], "q-yyyy") AS QuarterYear,
  Sum(
    [OrderDetails].[Quantity] *(1 - [OrderDetails].[Discount]) * [OrderDetails].[UnitPrice]
  ) AS OrderTotal
FROM
  Products
  INNER JOIN (
    Orders
    INNER JOIN OrderDetails ON Orders.OrderID = OrderDetails.OrderID
  ) ON Products.ProductID = OrderDetails.ProductID
WHERE
  (
    (
      (Orders.OrderDate) Between reportParameterStartDate()
      And reportParameterEndDate()
    )
  )
GROUP BY
  Products.ProductName,
  Products.ProductID,
  Format([OrderDate], "q-yyyy");

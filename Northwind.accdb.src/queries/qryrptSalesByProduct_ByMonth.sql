SELECT
  Products.ProductName,
  Products.ProductID,
  Format([OrderDate], "mmm-yyyy") AS MonthYear,
  Format([OrderDate], "yyyy-mm") AS MonthYearSort,
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
  Format([OrderDate], "mmm-yyyy"),
  Format([OrderDate], "yyyy-mm");

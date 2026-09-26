SELECT
  OrderDetails.ProductID,
  Sum(
    [Quantity] *(1 - [Discount]) * [UnitPrice]
  ) AS TotalSales
FROM
  OrderDetails
GROUP BY
  OrderDetails.ProductID;

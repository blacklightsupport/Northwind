SELECT
  Orders.OrderID,
  Sum(
    [Quantity] *(1 - [Discount]) * [UnitPrice]
  ) AS OrderTotal
FROM
  Orders
  LEFT JOIN OrderDetails ON Orders.OrderID = OrderDetails.OrderID
GROUP BY
  Orders.OrderID;

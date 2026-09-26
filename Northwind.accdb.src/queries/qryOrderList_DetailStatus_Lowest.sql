SELECT
  OrderDetails.OrderID,
  Min(OrderDetailStatus.SortOrder) AS MinOfSortOrder
FROM
  OrderDetailStatus
  INNER JOIN OrderDetails ON OrderDetailStatus.OrderDetailStatusID = OrderDetails.OrderDetailStatusID
GROUP BY
  OrderDetails.OrderID
ORDER BY
  OrderDetails.OrderID,
  Min(OrderDetailStatus.SortOrder);

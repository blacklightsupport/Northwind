SELECT
  Orders.*,
  OrderStatus.OrderStatusName
FROM
  OrderStatus
  INNER JOIN Orders ON OrderStatus.OrderStatusID = Orders.OrderStatusID;

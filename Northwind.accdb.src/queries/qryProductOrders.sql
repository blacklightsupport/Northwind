SELECT
  OrderDetails.OrderID,
  OrderDetails.ProductID,
  Orders.OrderDate,
  OrderDetails.Quantity,
  OrderDetails.UnitPrice,
  [OrderDetails].[Quantity] * [OrderDetails].[UnitPrice] AS ExtendedPrice,
  OrderStatus.OrderStatusName AS OrderStatus,
  OrderDetailStatus.OrderDetailStatusName AS ProductStatus
FROM
  Products
  INNER JOIN (
    OrderStatus
    INNER JOIN (
      Orders
      INNER JOIN (
        OrderDetailStatus
        INNER JOIN OrderDetails ON OrderDetailStatus.OrderDetailStatusID = OrderDetails.OrderDetailStatusID
      ) ON Orders.OrderID = OrderDetails.OrderID
    ) ON OrderStatus.OrderStatusID = Orders.OrderStatusID
  ) ON Products.ProductID = OrderDetails.ProductID;

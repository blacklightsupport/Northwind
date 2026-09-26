SELECT
  PurchaseOrders.StatusID,
  PurchaseOrderDetails.ProductID,
  Sum(PurchaseOrderDetails.Quantity) AS Quantity
FROM
  PurchaseOrders
  INNER JOIN PurchaseOrderDetails ON PurchaseOrders.PurchaseOrderID = PurchaseOrderDetails.PurchaseOrderID
GROUP BY
  PurchaseOrders.StatusID,
  PurchaseOrderDetails.ProductID;

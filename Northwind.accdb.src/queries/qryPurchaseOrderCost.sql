SELECT
  PurchaseOrders.PurchaseOrderID,
  Sum([Quantity] * [UnitCost]) AS ExtendedCost
FROM
  PurchaseOrders
  LEFT JOIN PurchaseOrderDetails ON PurchaseOrders.PurchaseOrderID = PurchaseOrderDetails.PurchaseOrderID
GROUP BY
  PurchaseOrders.PurchaseOrderID;

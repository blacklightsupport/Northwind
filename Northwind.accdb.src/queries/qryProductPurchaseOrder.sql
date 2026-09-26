SELECT
  Products.ProductID,
  PurchaseOrderDetails.PurchaseOrderID,
  PurchaseOrderStatus.StatusID,
  PurchaseOrderStatus.StatusName,
  PurchaseOrderStatus.SortOrder,
  PurchaseOrderDetails.Quantity,
  PurchaseOrderDetails.UnitCost,
  [Quantity] * [UnitCost] AS ExtendedCost,
  Companies.CompanyName,
  PurchaseOrders.SubmittedDate,
  PurchaseOrderDetails.ReceivedDate
FROM
  PurchaseOrderStatus
  RIGHT JOIN (
    (
      Companies
      INNER JOIN PurchaseOrders ON Companies.CompanyID = PurchaseOrders.VendorID
    )
    INNER JOIN (
      Products
      INNER JOIN PurchaseOrderDetails ON Products.ProductID = PurchaseOrderDetails.ProductID
    ) ON PurchaseOrders.PurchaseOrderID = PurchaseOrderDetails.PurchaseOrderID
  ) ON PurchaseOrderStatus.StatusID = PurchaseOrders.StatusID
ORDER BY
  PurchaseOrderStatus.SortOrder,
  PurchaseOrders.SubmittedDate;

SELECT
  PurchaseOrders.PurchaseOrderID,
  PurchaseOrderStatus.StatusName,
  qrycboVendors.CompanyName,
  Nz(
    [qryPurchaseOrderCost].[ExtendedCost]
  ) + Nz([PurchaseOrders].[ShippingFee]) + Nz([PurchaseOrders].[TaxAmount]) AS TotalCost,
  [qrycboEmployees-Submitted].FullNameFNLN AS Submitter,
  PurchaseOrders.SubmittedDate,
  [qrycboEmployees-Approved].FullNameFNLN AS Approver,
  PurchaseOrders.ApprovedDate,
  PurchaseOrders.PaymentDate
FROM
  PurchaseOrderStatus
  INNER JOIN (
    qrycboEmployees AS [qrycboEmployees-Approved]
    RIGHT JOIN (
      qrycboEmployees AS [qrycboEmployees-Submitted]
      RIGHT JOIN (
        (
          qrycboVendors
          INNER JOIN PurchaseOrders ON qrycboVendors.CompanyID = PurchaseOrders.VendorID
        )
        INNER JOIN qryPurchaseOrderCost ON PurchaseOrders.PurchaseOrderID = qryPurchaseOrderCost.PurchaseOrderID
      ) ON [qrycboEmployees-Submitted].EmployeeID = PurchaseOrders.SubmittedByID
    ) ON [qrycboEmployees-Approved].EmployeeID = PurchaseOrders.ApprovedByID
  ) ON PurchaseOrderStatus.StatusID = PurchaseOrders.StatusID
ORDER BY
  PurchaseOrders.PurchaseOrderID DESC;

SELECT
  PurchaseOrders.VendorID AS myCompanyID,
  PurchaseOrders.PurchaseOrderID,
  PurchaseOrders.SubmittedByID,
  PurchaseOrders.SubmittedDate,
  PurchaseOrders.ApprovedByID,
  PurchaseOrders.ApprovedDate,
  PurchaseOrders.StatusID,
  PurchaseOrders.ReceivedDate,
  PurchaseOrders.ShippingFee,
  PurchaseOrders.TaxAmount,
  PurchaseOrders.PaymentDate,
  PurchaseOrders.PaymentAmount,
  PurchaseOrders.PaymentMethod,
  PurchaseOrders.Notes,
  PurchaseOrders.AddedBy,
  PurchaseOrders.AddedOn,
  PurchaseOrders.ModifiedBy,
  PurchaseOrders.ModifiedOn,
  SumbittedBy.FullNameFNLN AS SubmittedBy,
  ApprovedBy.FullNameFNLN AS ApprovedBy,
  PurchaseOrderStatus.StatusName,
  qryPurchaseOrderCost.ExtendedCost
FROM
  PurchaseOrderStatus
  INNER JOIN (
    (
      (
        PurchaseOrders
        INNER JOIN qrycboEmployees AS SumbittedBy ON PurchaseOrders.SubmittedByID = SumbittedBy.EmployeeID
      )
      LEFT JOIN qrycboEmployees AS ApprovedBy ON PurchaseOrders.ApprovedByID = ApprovedBy.EmployeeID
    )
    LEFT JOIN qryPurchaseOrderCost ON PurchaseOrders.PurchaseOrderID = qryPurchaseOrderCost.PurchaseOrderID
  ) ON PurchaseOrderStatus.StatusID = PurchaseOrders.StatusID;

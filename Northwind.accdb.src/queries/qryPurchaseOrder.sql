SELECT
  PurchaseOrders.*,
  PurchaseOrderStatus.StatusName,
  qryEmployees.FullNameFNLN AS SubmittedBy,
  qryEmployees_1.FullNameFNLN AS ApprovedBy
FROM
  PurchaseOrderStatus
  INNER JOIN (
    (
      PurchaseOrders
      LEFT JOIN qryEmployees ON PurchaseOrders.SubmittedByID = qryEmployees.EmployeeID
    )
    LEFT JOIN qryEmployees AS qryEmployees_1 ON PurchaseOrders.ApprovedByID = qryEmployees_1.EmployeeID
  ) ON PurchaseOrderStatus.StatusID = PurchaseOrders.StatusID;

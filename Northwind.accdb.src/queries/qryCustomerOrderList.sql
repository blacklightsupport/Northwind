SELECT
  Orders.OrderID,
  Orders.EmployeeID,
  Orders.CustomerID AS myCompanyID,
  Orders.OrderDate,
  Orders.ShippedDate,
  Orders.ShipperID,
  Orders.ShippingFee,
  Orders.TaxRate,
  Orders.TaxStatusID,
  Orders.PaymentMethod,
  Orders.PaidDate,
  Orders.Notes,
  Orders.OrderStatusID,
  Orders.AddedBy,
  Orders.AddedOn,
  Orders.ModifiedBy,
  Orders.ModifiedOn,
  qrycboEmployees.FullNameFNLN AS EmployeeName,
  qrycboShippers.CompanyName AS ShipperName,
  qryOrderTotal.OrderTotal,
  qrycboOrderStatus.OrderStatusName
FROM
  (
    (
      (
        Orders
        INNER JOIN qrycboEmployees ON Orders.EmployeeID = qrycboEmployees.EmployeeID
      )
      LEFT JOIN qrycboShippers ON Orders.ShipperID = qrycboShippers.CompanyID
    )
    LEFT JOIN qryOrderTotal ON Orders.OrderID = qryOrderTotal.OrderID
  )
  INNER JOIN qrycboOrderStatus ON Orders.OrderStatusID = qrycboOrderStatus.OrderStatusID
ORDER BY
  Orders.OrderID DESC;

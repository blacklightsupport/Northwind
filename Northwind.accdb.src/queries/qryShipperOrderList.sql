SELECT
  Orders.OrderID,
  qrycboEmployees.FullNameFNLN AS EmployeeFNLN,
  Orders.OrderDate,
  qryOrderTotal.OrderTotal,
  OrderStatus.OrderStatusName,
  Orders.CustomerID,
  Orders.ShipperID AS myCompanyID,
  qrycboCustomers.CompanyName AS CustomerName
FROM
  OrderStatus
  INNER JOIN (
    (
      (
        Orders
        INNER JOIN qrycboEmployees ON Orders.EmployeeID = qrycboEmployees.EmployeeID
      )
      INNER JOIN qryOrderTotal ON Orders.OrderID = qryOrderTotal.OrderID
    )
    INNER JOIN qrycboCustomers ON Orders.CustomerID = qrycboCustomers.CompanyID
  ) ON OrderStatus.OrderStatusID = Orders.OrderStatusID
ORDER BY
  Orders.OrderID;

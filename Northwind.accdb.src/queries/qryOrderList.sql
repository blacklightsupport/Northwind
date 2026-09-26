SELECT
  Orders.*,
  qrycboEmployees.FullNameFNLN,
  qryCustomers.CompanyName,
  qryOrderTotal.OrderTotal,
  OrderStatus.OrderStatusName,
  qryOrderList_DetailStatus.OrderDetailStatusName,
  DateValue([OrderDate]) AS OrderDateOnly
FROM
  OrderStatus
  INNER JOIN (
    (
      qrycboEmployees
      INNER JOIN (
        (
          qryCustomers
          INNER JOIN Orders ON qryCustomers.CompanyID = Orders.CustomerID
        )
        INNER JOIN qryOrderTotal ON Orders.OrderID = qryOrderTotal.OrderID
      ) ON qrycboEmployees.EmployeeID = Orders.EmployeeID
    )
    LEFT JOIN qryOrderList_DetailStatus ON Orders.OrderID = qryOrderList_DetailStatus.OrderID
  ) ON OrderStatus.OrderStatusID = Orders.OrderStatusID
ORDER BY
  Orders.OrderID DESC;

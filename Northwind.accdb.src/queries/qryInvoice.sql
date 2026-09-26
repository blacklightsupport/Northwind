SELECT
  qryCustomers.CompanyName,
  qryCustomers.Address,
  qryCustomers.City,
  qryCustomers.StateAbbrev,
  qryCustomers.Zip,
  Orders.*,
  qryEmployees.FullNameFNLN AS SalesPerson,
  OrderDetails.ProductID,
  OrderDetails.Quantity,
  OrderDetails.UnitPrice,
  OrderDetails.Discount,
  Products.ProductCode,
  Products.ProductName
FROM
  Products
  INNER JOIN (
    (
      qryEmployees
      INNER JOIN (
        qryCustomers
        INNER JOIN Orders ON qryCustomers.CompanyID = Orders.CustomerID
      ) ON qryEmployees.EmployeeID = Orders.EmployeeID
    )
    INNER JOIN OrderDetails ON Orders.OrderID = OrderDetails.OrderID
  ) ON Products.ProductID = OrderDetails.ProductID;

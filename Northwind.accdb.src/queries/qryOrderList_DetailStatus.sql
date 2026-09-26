SELECT
  qryOrderList_DetailStatus_Lowest.OrderID,
  OrderDetailStatus.OrderDetailStatusName
FROM
  qryOrderList_DetailStatus_Lowest
  INNER JOIN OrderDetailStatus ON qryOrderList_DetailStatus_Lowest.MinOfSortOrder = OrderDetailStatus.SortOrder;

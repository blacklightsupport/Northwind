SELECT
  StockTake.ProductID,
  Max(StockTake.StockTakeDate) AS MaxOfStockTakeDate
FROM
  StockTake
GROUP BY
  StockTake.ProductID;

CREATE TABLE [StockTake] (
  [StockTakeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StockTakeDate] DATETIME,
  [ProductID] LONG CONSTRAINT [New_New_ProductsStockTake] REFERENCES [Products] ([ProductID]),
  [QuantityOnHand] SHORT,
  [ExpectedQuantity] LONG,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

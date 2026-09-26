CREATE TABLE [Products] (
  [ProductID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProductCode] VARCHAR (20) CONSTRAINT [ProductCode] UNIQUE,
  [ProductName] VARCHAR (50) CONSTRAINT [ProductName] UNIQUE,
  [ProductDescription] LONGTEXT,
  [StandardUnitCost] CURRENCY,
  [UnitPrice] CURRENCY,
  [ReorderLevel] SHORT,
  [TargetLevel] SHORT,
  [QuantityPerUnit] VARCHAR (50),
  [Discontinued] BIT,
  [MinimumReorderQuantity] SHORT,
  [ProductCategoryID] LONG CONSTRAINT [New_New_ProductCategories_NEWProducts] REFERENCES [ProductCategories] ([ProductCategoryID]),
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

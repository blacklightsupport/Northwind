CREATE TABLE [PurchaseOrderStatus] (
  [StatusID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StatusName] VARCHAR (50) CONSTRAINT [StatusName] UNIQUE,
  [SortOrder] BYTE CONSTRAINT [SortOrder] UNIQUE,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

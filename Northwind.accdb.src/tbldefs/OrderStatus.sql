CREATE TABLE [OrderStatus] (
  [OrderStatusID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [OrderStatusCode] VARCHAR (5) CONSTRAINT [StatusCode] UNIQUE,
  [OrderStatusName] VARCHAR (50) CONSTRAINT [StatusName] UNIQUE,
  [SortOrder] BYTE CONSTRAINT [SortOrder] UNIQUE,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

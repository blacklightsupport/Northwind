CREATE TABLE [PurchaseOrderDetails] (
  [PurchaseOrderDetailID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [PurchaseOrderID] LONG CONSTRAINT [New_New_PurchaseOrdersPurchaseOrderDetails] REFERENCES [PurchaseOrders] ([PurchaseOrderID]) ON DELETE CASCADE ,
  [ProductID] LONG CONSTRAINT [New_New_ProductsPurchaseOrderDetails] REFERENCES [Products] ([ProductID]),
  [Quantity] SHORT,
  [UnitCost] CURRENCY,
  [ReceivedDate] DATETIME,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME,
   CONSTRAINT 
)

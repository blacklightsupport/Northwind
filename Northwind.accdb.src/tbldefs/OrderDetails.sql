CREATE TABLE [OrderDetails] (
  [OrderDetailID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [OrderID] LONG CONSTRAINT [New_New_OrdersOrderDetails] REFERENCES [Orders] ([OrderID]) ON DELETE CASCADE ,
  [ProductID] LONG CONSTRAINT [New_New_ProductsOrderDetails] REFERENCES [Products] ([ProductID]),
  [Quantity] SHORT,
  [UnitPrice] CURRENCY,
  [Discount] SINGLE,
  [OrderDetailStatusID] LONG CONSTRAINT [New_New_OrderDetailsStatusOrderDetails] REFERENCES [OrderDetailStatus] ([OrderDetailStatusID]),
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME,
   CONSTRAINT 
)

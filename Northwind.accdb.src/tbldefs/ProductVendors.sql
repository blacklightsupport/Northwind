CREATE TABLE [ProductVendors] (
  [ProductVendorID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProductID] LONG CONSTRAINT [New_New_ProductsProductVendors] REFERENCES [Products] ([ProductID]),
  [VendorID] LONG,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME,
   CONSTRAINT 
)

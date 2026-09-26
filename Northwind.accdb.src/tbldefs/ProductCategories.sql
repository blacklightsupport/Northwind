CREATE TABLE [ProductCategories] (
  [ProductCategoryID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProductCategoryName] VARCHAR (255) CONSTRAINT [ProductCategory] UNIQUE,
  [ProductCategoryCode] VARCHAR (3) CONSTRAINT [ProductCategoryCode] UNIQUE,
  [ProductCategoryDesc] VARCHAR (255),
  [ProductCategoryImage] VARCHAR,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

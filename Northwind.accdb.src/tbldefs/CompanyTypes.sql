CREATE TABLE [CompanyTypes] (
  [CompanyTypeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [CompanyType] VARCHAR (50) CONSTRAINT [CompanyType] UNIQUE,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

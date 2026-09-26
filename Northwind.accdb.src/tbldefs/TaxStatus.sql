CREATE TABLE [TaxStatus] (
  [TaxStatusID] BYTE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [TaxStatus] VARCHAR (50) CONSTRAINT [TaxStatus] UNIQUE,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

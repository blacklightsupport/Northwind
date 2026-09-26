CREATE TABLE [Titles] (
  [Title] VARCHAR (20) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

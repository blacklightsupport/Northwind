CREATE TABLE [Companies] (
  [CompanyID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [CompanyName] VARCHAR (50),
  [CompanyTypeID] LONG CONSTRAINT [New_New_CompanyTypesCompanies] REFERENCES [CompanyTypes] ([CompanyTypeID]),
  [BusinessPhone] VARCHAR (20),
  [Address] VARCHAR (255),
  [City] VARCHAR (255),
  [StateAbbrev] VARCHAR (2) CONSTRAINT [New_New_StatesCompanies] REFERENCES [States] ([StateAbbrev]),
  [Zip] VARCHAR (10),
  [Website] LONGTEXT,
  [Notes] LONGTEXT,
  [StandardTaxStatusID] BYTE,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME,
   CONSTRAINT 
)

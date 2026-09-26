CREATE TABLE [Contacts] (
  [ContactID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [CompanyID] LONG CONSTRAINT [New_New_CompaniesContacts] REFERENCES [Companies] ([CompanyID]),
  [LastName] VARCHAR (30),
  [FirstName] VARCHAR (20),
  [EmailAddress] VARCHAR (255),
  [JobTitle] VARCHAR (50),
  [PrimaryPhone] VARCHAR (20),
  [SecondaryPhone] VARCHAR (20),
  [Notes] LONGTEXT,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME,
   CONSTRAINT ,
   CONSTRAINT 
)

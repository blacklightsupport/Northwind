CREATE TABLE [Privileges] (
  [PrivilegeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [PrivilegeName] VARCHAR (50) CONSTRAINT [PrivilegeName] UNIQUE,
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME
)

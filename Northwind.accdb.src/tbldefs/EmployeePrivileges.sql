CREATE TABLE [EmployeePrivileges] (
  [EmployeePrivilegeID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [EmployeeID] LONG CONSTRAINT [New_New_EmployeesEmployeePrivileges] REFERENCES [Employees] ([EmployeeID]) ON DELETE CASCADE ,
  [PrivilegeID] LONG CONSTRAINT [New_New_PrivilegesEmployeePrivileges] REFERENCES [Privileges] ([PrivilegeID]),
  [AddedBy] VARCHAR (255),
  [AddedOn] DATETIME,
  [ModifiedBy] VARCHAR (255),
  [ModifiedOn] DATETIME,
   CONSTRAINT 
)

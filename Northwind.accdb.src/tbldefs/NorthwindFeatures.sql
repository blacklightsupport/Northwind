CREATE TABLE [NorthwindFeatures] (
  [NorthwindFeaturesID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ItemName] VARCHAR (255),
  [Description] VARCHAR (255),
  [Navigation] VARCHAR (255),
  [LearnMore] LONGTEXT,
  [HelpKeywords] VARCHAR (255),
  [OpenMethod] LONG
)

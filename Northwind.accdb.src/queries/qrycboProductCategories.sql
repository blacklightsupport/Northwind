select DISTINCT
  1 as theClause,
  0 as ProductCategoryID,
  "<All>" as ProductCategoryName
FROM
  ProductCategories
UNION ALL
SELECT
  2 as theClause,
  ProductCategories.ProductCategoryID,
  ProductCategories.ProductCategoryName
FROM
  ProductCategories
ORDER BY
  theClause,
  ProductCategoryName;

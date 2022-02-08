SELECT cc1.*
FROM CC_Findings AS cc1
WHERE cc1.ID IN
(
  SELECT TOP 1 ID
  FROM CC_Findings as cc2
  WHERE cc2.TaskID = cc1.TaskID
  ORDER BY cc2.[Status Date] DESC
);

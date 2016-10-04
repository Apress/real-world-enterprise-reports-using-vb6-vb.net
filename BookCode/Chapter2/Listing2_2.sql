DROP TABLE  #Dealer, #temp

CREATE TABLE #temp (level int, ID int)
CREATE TABLE #Dealer (seq int identity, level int, ID int)

DECLARE @level int, @curr int
SELECT TOP 1 @level=1, @curr=ID 
FROM Dealer 
WHERE ID=SponsorID

INSERT INTO #temp (level, ID) 
VALUES (@level, @curr)

WHILE (@level > 0) BEGIN

  IF EXISTS(SELECT * 
	FROM #temp 
	WHERE level=@level) 
BEGIN
    SELECT TOP 1 @curr=ID 
    FROM #temp
    WHERE level=@level

    INSERT #Dealer (level, ID) 
    VALUES (@level, @curr)

    DELETE #temp
    WHERE level=@level 
    AND ID=@curr

    INSERT #temp
    SELECT @level+1, ID
    FROM Dealer
    WHERE SponsorID=@curr
    AND SponsorID <> ID

    IF (@@ROWCOUNT > 0) SET @level=@level+1
  END ELSE
    SET @level=@level-1

END

SELECT REPLICATE(CHAR(9),level)+i.LastName + ', ' + 
i.FirstName AS DealerName, level
FROM #Dealer d JOIN Dealer i ON d.ID = i.ID
ORDER BY seq

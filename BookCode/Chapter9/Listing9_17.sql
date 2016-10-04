CREATE PROCEDURE sp_PagedEquipment
(
	@Page int,
	@RecsPerPage int
)

AS

/*
We don ’t need the #\of rows inserted
into our temporary table so turn off
the count
*/

SET NOCOUNT ON

/*Create a temporary table */
CREATE TABLE #TempEquipment
(
	ID int IDENTITY,
	EID int,
	INV_NUMBER int,
	Bldg varchar(5),
	Area varchar(5),
	OnHand int
)


/*
Insert the rows from tblEquipment into the temp table
by selecting from the master table
*/

INSERT INTO #TempEquipment (EID,INV_NUMBER,Bldg,Area,OnHand)
SELECT EID,INV_NUMBER,Bldg,Area,OnHand FROM tblEquipment ORDER BY INV_NUMBER


/*Find first and last records that we want to return */

DECLARE @FirstRec int,@LastRec int


/*
Calculate the starting record for this “page ”
The way we calculate the page sets is quite simple.
use the page number,subtract 1 (which would position
your cursor at the end of the prior page,and then
multiply it by the number of records being returned
by this cursor.
*/

SELECT @FirstRec =(@Page -1)*@RecsPerPage


/*Calculate the ending record for this “page ”*/

SELECT @LastRec =(@Page *@RecsPerPage +1)


/*
Return the paged set of records including a column
which contains a Boolean indicating if there are more
records left
*/

SELECT *,
MoreRecords =
(
SELECT COUNT(*)
FROM #TempEquipment E
WHERE E.ID >=@LastRec
)
FROM #TempEquipment
WHERE ID >@FirstRec AND ID <@LastRec

/*Turn COUNT back on */
SET NOCOUNT OFF
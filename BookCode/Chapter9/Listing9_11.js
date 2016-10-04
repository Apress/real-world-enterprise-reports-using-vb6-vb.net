function tblInsertRow()
{

	var oTR;		// New Table Row Object
	var oTD;		// New Table Data Object
	var i;			// Loop Counter
		
	// Using the Table ID (tblDemo) object we can 
	// insert a new table row.  Doing this will
	// return a Table Row (TR) object
	oTR = tblDemo.insertRow();
	
	for (i = 1; i < 4; i++) 
	{
			
		//Using the Table Row Object Insert a Cell
		oTD = oTR.insertCell();			
		oTD.align = "center"
		
		// Using the Rows calculate new cell information
		oTD.innerText="Element " + (tblDemo.rows.length -1) + ":" + i;	

	}
}

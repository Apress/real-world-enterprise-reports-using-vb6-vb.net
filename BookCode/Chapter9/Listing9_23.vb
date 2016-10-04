Dim objCol1 As New HyperLinkColumn()
Dim objCol2 As New BoundColumn()
Dim objCol3 As New BoundColumn()
Dim objCol4 As New BoundColumn()
Dim objStyle As New Style()

With objCol1
    .DataTextField = "ConventionName"
    .HeaderText = "Convention"
    .DataNavigateUrlField = "ID"
    .DataNavigateUrlFormatString = "convinfo.aspx?id={0}"
    .SortExpression = "ConventionName"
End With

With objCol2
    .DataField = "StartDate"
    .HeaderText = "Start Date"
    .DataFormatString = "{0:MM/dd/yyyy}"
    .SortExpression = "StartDate"
End With

With objCol3
    .DataField = "City"
    .HeaderText = "City"
    .SortExpression = "City"
End With

With objCol4
    .DataField = "State"
    .HeaderText = "State"
    .SortExpression = "State"
End With

With DataGrid1
     .AllowPaging = True
     .PageSize = 5
    .AllowSorting = True
    .AutoGenerateColumns = False
    .Columns.Add (objCol1)
    .Columns.Add (objCol2)
    .Columns.Add (objCol3)
    .Columns.Add (objCol4)
    .BackColor = objStyle.BackColor.Gray
    .ForeColor = objStyle.BackColor.Black
    .GridLines = GridLines.None
    .Width = Unit.Pixel(600)
    .GridLines = GridLines.None
    .CellPadding = 5
    .CellSpacing = 5
    .BorderWidth = Unit.Point(10)
    .BorderColor = Color.Red
End With

Call BindToDataSource

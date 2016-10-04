<%

Dim iReport
Dim cSQL
Dim cCriteria
Dim cProduct
Dim cDept
Dim cSourceID
Dim iClosed
Dim cOrderDateFrom
Dim cOrderDateTo
Dim cShipDateFrom
Dim cShipDateTo
			
iReport = Request.Form("Report")

Select Case iReport

	Case 0

		With Request
			cProduct = .Form("lstProduct")
			cDept = .Form("lstDepartment")
			cSourceID = .Form("cboSource")
			iClosed = .Form("cboClosed")
			
			cOrderDateFrom = _
				.Form("txtShipMonthFrom") & "/" & _
				.Form("txtShipDayFrom") & "/" & _
				.Form("txtShipYearFrom")
			
			cOrderDateTo = _
				.Form("txtShipMonthTo") & "/" & _
				.Form("txtShipDayTo") & "/" & _
				.Form("txtShipYearTo")

			cShipDateFrom = _
				.Form("txtOrderMonthFrom") & "/" & _
				.Form("txtOrderDayFrom") & "/" & _
				.Form("txtOrderYearFrom")
				
			cShipDateTo = _
				.Form("txtOrderMonthTo") & "/" & _
				.Form("txtOrderDayTo") & "/" & _
				.Form("txtOrderYearTo")
		End With				
		
		cSQL = GetListBoxSQL(cProduct, "productid")
		cSQL = cSQL & GetListBoxSQL(cDept, "deptid")
		cSQL = cSQL & GetComboBoxSQL(cSourceID, "sourceid")
		cSQL = cSQL & GetCheckBoxSQL(iClosed, "closed")			
		cSQL = cSQL & GetDateSQL("orderdate", cOrderDateFrom, cOrderDateTo)
		cSQL = cSQL & GetDateSQL("shipdate", cShipDateFrom, cShipDateTo)

		cCriteria = GetListBoxCriteria(cProduct, "Product")
		cCriteria = cCriteria & GetListBoxCriteria(cDept, "Department")
		cCriteria = cCriteria & GetComboBoxCriteria(cSourceID, "Source")
		cCriteria = cCriteria & GetCheckBoxCriteria(iClosed, "Closed")			
		cCriteria = cCriteria & GetDateCriteria("Order Date", _
		                        cOrderDateFrom, cOrderDateTo)
		cCriteria = cCriteria & GetDateCriteria("Ship Date", _
		                        cShipDateFrom, cShipDateTo)

End Select

If cSQL <> "" Then
    cSQL = Mid(cSQL, 1, Len(cSQL) - Len(" AND "))
End If

If cCriteria <> "" Then
    cCriteria = Mid(cCriteria, 1, Len(cCriteria) - Len(" AND "))
End If

Response.Write cSQL
Response.Write "<P>"
Response.Write cCriteria
    
%>

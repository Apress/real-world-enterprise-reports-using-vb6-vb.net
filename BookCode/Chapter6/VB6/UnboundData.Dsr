VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} UnboundData 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14052
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24786
   _ExtentY        =   14499
   SectionData     =   "UnboundData.dsx":0000
End
Attribute VB_Name = "UnboundData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_DataInitialize()
    
    Fields.Add "ExtendedPrice"

End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If EOF Then
        Exit Sub
    End If
    
    Fields("ExtendedPrice").Value = Format(Fields("UnitPrice").Value * _
                                    Fields("Quantity").Value, "###,##0.00")
        
End Sub


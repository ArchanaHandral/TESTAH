VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARPRGProdPlan 
   Caption         =   "Production Planning Form"
   ClientHeight    =   13290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17205
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   30348
   _ExtentY        =   23442
   SectionData     =   "ARPRGProdPlan.dsx":0000
End
Attribute VB_Name = "ARPRGProdPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This is the implementation of the Computerization Order Form for a Production Run Group (PRG).</summary>
'</comment>

Option Explicit

Private Sub Detail_Format()
    Dim i As Long
    Dim arrFilePath() As String
    
    Me.RunBarcode = "[" & ProductionGroup.Barcode_Id & "]"
    Me.Label3Barcode = "[" & ProductionGroup.Barcode_Id & "]"
    Me.otxtLabel3JobNumber = gJobNumber
    Me.otxtLabel3IDNumber = gLabelId

    For i = 1 To PlanningList.count
        If i = 1 Then
            Me.otxtLabel3PDRs.Text = PlanningList.Item(i).Form_Idectification
        Else
            Me.otxtLabel3PDRs.Text = Me.otxtLabel3PDRs & ", " & PlanningList.Item(i).Form_Idectification
        End If
    Next i

    Me.otxtLabel3QtyProduced = ProductionGroup.Qty_Requested & " CS LABELS" & " + " & ProductionGroup.Samples_Requested - ProductionGroup.Clintrak_Samples & " CLIENT SAMPLES"
    Me.otxtLabel3Description = ProductionGroup.description
    Me.otxtStockNo = ProductionGroup.stock & " - " & ProductionGroup.Stock_Desc
    If ProductionGroup.DigitalLabelParts <> "" Then
        Me.otxtStockNo = Me.otxtStockNo & vbCrLf & "(" & ProductionGroup.DigitalLabelParts & ")"
    End If
    Me.otxtOnsertPressDie = ProductionGroup.OnsertPressDie
        
    ' DW 2010-002 added to remove "-" when there is no description for "N/A"
    Select Case ProductionGroup.Scratch_Stock
        Case "N/A"
            Me.otxtScratchStockNo = ProductionGroup.Scratch_Stock
        Case Else
            Me.otxtScratchStockNo = ProductionGroup.Scratch_Stock & " - " & ProductionGroup.Scratch_Stock_Desc
    End Select

    're-get label description here.  This description doesn't have the treatment group name
    Me.otxtDesc = ProductionGroup.description
    Me.otxtClient = gClientName
    Me.otxtProtocol = gProtocol
    Me.otxtIDNumber = gLabelId
    Me.otxtJobNo = gJobNumber
    Me.otxtFormID = ProductionGroup.Form_Identification
    Me.otxtQtyRequested = ProductionGroup.Qty_Requested
    Me.otxtQtySamples = ProductionGroup.Samples_Requested - ProductionGroup.Clintrak_Samples
    Me.otxtFileName = ProductionGroup.File_Name
    
    ' Linebreak filename portion of path
    arrFilePath = Split(Me.otxtFileName.Text, "\")
    arrFilePath(UBound(arrFilePath)) = vbCrLf & arrFilePath(UBound(arrFilePath))
    Me.otxtFileName.Text = Join(arrFilePath, "\")
    
    Me.otxtProducedBy = ProductionGroup.Produced_By
    Me.otxtDateProduced = ProductionGroup.Produced_Date & " " & gClintrakLocations(CStr(ProductionGroup.Clintrak_Location_Id)).Time_Zone_Display
    Me.otxtRandId = gRandIDNumber
    
    'md 2006-020
    If gReOrientFlag Then
        Me.lblReorient.Visible = True
    Else
        Me.lblReorient.Visible = False
    End If

End Sub

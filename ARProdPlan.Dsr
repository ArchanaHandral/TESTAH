VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARProdPlan 
   Caption         =   "Computerization Order Form"
   ClientHeight    =   12930
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   17970
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   31697
   _ExtentY        =   22807
   SectionData     =   "ARProdPlan.dsx":0000
End
Attribute VB_Name = "ARProdPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<comment>
' <summary>
' This is the implementation of the Computerization Order Form for a Production Run (PDR).</summary>
'</comment>
Option Explicit

Private madoData As nADOData.CADOData

Private Sub Detail_Format()
    Dim strPlainText As String
    Dim arrFilePath() As String
    
    If booReplacement Then
        Me.lblReplacements.Visible = True
        If gReprintFile_Type = "REPLACEMENT" Then
            Me.lblReplacements.Caption = "** REPLACEMENTS **"
            Me.oLbl1Clintrak.Caption = "MATERIAL SHIPPING TAG (Replacements)"
            Me.oLbl2Clintrak.Caption = "MATERIAL SHIPPING TAG (Replacements)"
            Me.oLbl3Clintrak.Caption = "IN-PROCESS (Replacements)"
            Me.otxtGroupName = gGroupName
        Else
            Me.lblReplacements.Caption = "** RESUPPLY **"
            Me.oLbl1Clintrak.Caption = "MATERIAL SHIPPING TAG (ReSupply)"
            Me.oLbl2Clintrak.Caption = "MATERIAL SHIPPING TAG (ReSupply)"
            Me.oLbl3Clintrak.Caption = "IN-PROCESS (ReSupply)"
            Me.otxtGroupName = "N/A"
        End If
    Else
        Me.lblReplacements.Visible = False
        Me.oLbl1Clintrak.Caption = "MATERIAL SHIPPING TAG"
        Me.oLbl2Clintrak.Caption = "MATERIAL SHIPPING TAG"
        Me.oLbl3Clintrak.Caption = "IN-PROCESS"
        Me.otxtGroupName = gGroupName
    End If

    Me.RunBarcode = "[" & ProductionRun.Barcode_Id & "]"
    Me.otxtLabel1CustomerName = gClientName
    Me.otxtLabel1JobNumber = gJobNumber
    Me.otxtLabel1IDNumber = gLabelId
    Me.otxtLabel1Protocol = gProtocol
    Me.otxtLabel1Desc = ProductionRun.Prod_Description
    Me.otxtLabel1Qty = ProductionRun.Samples_Requested - ProductionRun.Clintrak_Samples & " CLIENT SAMPLES"
    Me.bcSamples.Message = ProductionRun.Barcode_Id
    Me.otxtLabel2CustomerName = gClientName
    Me.otxtLabel2JobNumber = gJobNumber
    Me.otxtLabel2IDNumber = gLabelId
    Me.otxtLabel2Protocol = gProtocol
    Me.otxtLabel2Desc = ProductionRun.Prod_Description
    Me.otxtLabel2Qty = ProductionRun.Qty_Requested & " CS LABELS"
    Me.bcQuantity.Message = ProductionRun.Barcode_Id
    Me.otxtLabel3JobNumber = gJobNumber
    Me.otxtLabel3IDNumber = gLabelId
    Me.otxtLabel3Description = ProductionRun.Prod_Description
    Me.otxtLabel3QtyProduced = ProductionRun.Qty_Requested & " CS LABELS" & " + " & ProductionRun.Samples_Requested - ProductionRun.Clintrak_Samples & " CLIENT SAMPLES"
    Me.bcLbl3ProdRun = "[" & ProductionRun.Barcode_Id & "]"
    Me.otxtStockNo = ProductionRun.stock & " - " & ProductionRun.Stock_Desc
    If ProductionRun.DigitalLabelParts <> "" Then
        Me.otxtStockNo = Me.otxtStockNo & vbCrLf & "(" & ProductionRun.DigitalLabelParts & ")"
    End If
    Me.otxtOnsertPressDie = ProductionRun.OnsertDiePartNumber
    
    ' DW 2010-002 added to remove "-" when there is no description for "N/A"
    Select Case ProductionRun.Scratch_Stock
        Case "N/A"
            Me.otxtScratchStockNo = ProductionRun.Scratch_Stock
        Case Else
            Me.otxtScratchStockNo = ProductionRun.Scratch_Stock & " - " & ProductionRun.Scratch_Stock_Description & vbCrLf
            If ProductionRun.Apply_ScratchOff = 1 Then
                Me.otxtScratchStockNo = Me.otxtScratchStockNo & "(Apply to Labels and Samples)"
            ElseIf ProductionRun.Apply_ScratchOff = 2 Then
                Me.otxtScratchStockNo = Me.otxtScratchStockNo & "(Apply to Labels Only)"
            End If
    End Select

    're-get label description here.  This description doesn't have the treatment group name
    Me.otxtDesc = ProductionRun.LabelDescription

    Me.otxtClient = gClientName
    Me.otxtProtocol = gProtocol
    Me.otxtIDNumber = gLabelId
    Me.otxtJobNo = gJobNumber
    Me.otxtFormID = ProductionRun.Form_Identification

    Me.otxtQtyRequested = ProductionRun.Qty_Requested
    Me.otxtSampleTypes = ProductionRun.Sample_Number
    
    Me.otxtQtySamples = ProductionRun.Samples_Requested - ProductionRun.Clintrak_Samples
    Me.otxtClintrakSamples = ProductionRun.Clintrak_Samples
    
    If booReplacement Then
        Me.otxtFileName = gReprintFileName
        Me.otxtCoding = ""
    Else
        Me.otxtFileName = gCodingFileName
        Me.otxtCoding = gCodingName
    End If

    ' Linebreak filename portion of path
    arrFilePath = Split(Me.otxtFileName.text, "\")
    arrFilePath(UBound(arrFilePath)) = vbCrLf & arrFilePath(UBound(arrFilePath))
    Me.otxtFileName.text = Join(arrFilePath, "\")
    
    Me.otxtProducedBy = ProductionRun.Produced_By
    Me.otxtDateProduced = ProductionRun.Produced_Date & " " & gClintrakLocations(CStr(ProductionRun.Clintrak_Location_Id)).Time_Zone_Display
        
    With oRichEdit1
        ' Get the plain text portion of the RichText and set the RTB back to empty.
        .TextRTF = ProductionRun.Special_Inst
        strPlainText = .text
        .text = ""
        ' Set the style for the RTB and re-enter the plain text
        .SelStart = 0
        .SelFontName = "Arial"
        .SelFontSize = 140
        .SelBold = True
        .SelItalic = False
        .SelUnderline = False
        .SelColor = vbRed
        .SelText = strPlainText
    End With

    Me.otxtRandId = gRandIDNumber

    'cc 2008-005
    ' DW 2008-017 added timezone to PDR Print Date
    Me.lblPrintDate = Me.lblPrintDate & _
        IIf(ProductionRun.PDR_Print_Date = "1/1/1900", "N/A", ProductionRun.PDR_Print_Date & " " & gClintrakLocations(CStr(ProductionRun.Clintrak_Location_Id)).Time_Zone_Display)

    Call LoadShippingInfo

    'if the sample requested = clintrak sample then we do not want to print the Shipping Tab for the Samples.
    'this is because there only exists clintrak samples
    If ProductionRun.Samples_Requested = ProductionRun.Clintrak_Samples Then
        Me.otxtLabel1CustomerName = ""
        Me.otxtLabel1JobNumber = ""
        Me.otxtLabel1IDNumber = ""
        Me.otxtLabel1Protocol = ""
        Me.otxtLabel1Desc = ""
        Me.otxtLabel1Qty = ""
        Me.oLbl1Clintrak.Caption = ""
        Me.oLbl1CustomerName.Caption = ""
        Me.Label28.Caption = ""
        Me.oLbl1JobNumber.Caption = ""
        Me.oLbl1LabelID.Caption = ""
        Me.Label29.Caption = ""
        Me.otxtLabel1Desc = "DO NOT USE THIS SHIPPING TAG - THERE ARE NO CLIENT SAMPLES"
        Me.oLbl1Clintrak = " **** VOID ****** VOID ****** VOID *****"
    End If

    If Not gShowQuarantineTagFlag Then
        Call Void_Quarantine_Tag
    End If

    Select Case ProductionRun.Reorient_Ind
        Case 1
            Me.otxtReorient = "YES"
        Case Else
            Me.otxtReorient = "N/A"
    End Select

    If ProductionRun.Reference_No = "" Then
        Me.otxtReferenceNo = "N/A"
    Else
        Me.otxtReferenceNo = ProductionRun.Reference_No
    End If
    
End Sub

Private Sub LoadShippingInfo()
    
    If madoData Is Nothing Then
    Set madoData = New nADOData.CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly

        .AddParameter "Job Log Id", gJob_Id, adInteger, adParamInput
        .AddParameter "Job Shipping Id", ProductionRun.Ship_To_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ShipToAddress"

        If Not .Recordset.EOF Then
            Me.otxtShip.text = Trim$(.Recordset!ShipTo_Description)
            If Trim$(.Recordset!Address_Line_1) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & vbCrLf & Trim$(.Recordset!Address_Line_1)
            End If
            If Trim$(.Recordset!Address_Line_2) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & vbCrLf & Trim$(.Recordset!Address_Line_2)
            End If
            If Trim$(.Recordset!Address_Line_3) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & vbCrLf & Trim$(.Recordset!Address_Line_3)
            End If
            Me.otxtShip.text = Me.otxtShip.text & vbCrLf
            If Trim$(.Recordset!city) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & Trim$(.Recordset!city) & ", "
            End If
            If Trim$(.Recordset!state) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & Trim$(.Recordset!state) & " "
            End If
            If Trim$(.Recordset!zip) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & Trim$(.Recordset!zip) & " "
            End If
            If Trim$(.Recordset!Country_Name) <> "" Then
                Me.otxtShip.text = Me.otxtShip.text & Trim$(.Recordset!Country_Name)
            End If
            Me.otxtAttn.text = Trim$(.Recordset!Attn_Description)
        End If
        .Recordset.Close
    End With

End Sub

Private Sub Void_Quarantine_Tag()
    Dim Planning As CPlanningMethods
    Set Planning = New CPlanningMethods

    Me.oLbl3Title.Visible = False
    Me.oLbl3JobNumber.Visible = False
    Me.otxtLabel3JobNumber.Visible = False
    Me.oLbl3Desc.Visible = False
    Me.oLbl3Manufacturer.Visible = False
    Me.otxtLabel3QtyProduced.Visible = False
    Me.oLbl3Quarantine.Visible = False
    Me.bcLbl3ProdRun.Visible = False
    Me.oLbl3PartNumber.Visible = False
    Me.otxtLabel3IDNumber.Visible = False
    
    Call Planning.CheckIfCombined(ProductionRun.Barcode_Id)
    
    Me.otxtLabel3Description.Left = 10200
    Me.otxtLabel3Description.text = "DO NOT USE THIS QUARANTINE TAG - THIS PDR IS ASSOCIATED TO GROUPED RUN " '& gProdGroupBarcode
    Me.oLbl3Clintrak = " **** VOID ****** VOID ****** VOID *****"

End Sub


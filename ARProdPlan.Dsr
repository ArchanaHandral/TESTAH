VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARProdPlan 
   Caption         =   "ProductionRuns - ARProdPlan (ActiveReport)"
   ClientHeight    =   13590
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   17565
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   30983
   _ExtentY        =   23971
   SectionData     =   "ARProdPlan.dsx":0000
End
Attribute VB_Name = "ARProdPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private madoData As nADOData.CADOData
Private Const TBD As String = "TBD"
Private Const YES As String = "Yes"
Private Const NOT_APPLICABLE As String = "N/A"
Private Const ALL_TREATMENT_GROUPS As String = "** ALL TREATMENT GROUPS **"
Private Const SEE_INDIVIDUAL_PDR As String = "See Individual PDRs"

Private Sub Detail_Format()
    Dim strPlainText As String
    
    Me.otxtClient = gClientName
    Me.otxtIDNumber = gLabelId
    Me.otxtJobNo = gJobNumber


    Set shipRpt.object = New ArShippingRpt2
    Load shipRpt.object

    'Handle PRG runs here
    If gIsPRGRun Then
        
        lblReplacements.Visible = False
        lblGroupedRun.Visible = True
        Me.otxtGroupName = ALL_TREATMENT_GROUPS
        lblPrintDate.Visible = False
        
        Select Case ProductionGroup.PrgReorientInd
                Case 1
                    Me.otxtReorient.text = YES
                Case 0
                    Me.otxtReorient.text = NOT_APPLICABLE
                Case Else
                    Me.otxtReorient = SEE_INDIVIDUAL_PDR
        End Select
        
        Me.otxtQtyRequested = ProductionGroup.Qty_Requested
        Me.otxtQtySamples = ProductionGroup.Samples_Requested - ProductionGroup.Clintrak_Samples
        Me.otxtSampleTypes = NOT_APPLICABLE
        Me.otxtClintrakSamples = NOT_APPLICABLE
    
        Me.otxtProducedBy = ProductionGroup.Produced_By
        Me.otxtDateProduced = ProductionGroup.Produced_Date & " " & gClintrakLocations(CStr(ProductionGroup.Clintrak_Location_Id)).Time_Zone_Display
        Me.otxtFormID = ProductionGroup.Form_Identification & " (" & ProductionGroup.TotalPDRsCount & " PDRs)"
        Me.RunBarcode = "[" & ProductionGroup.Barcode_Id & "]"
        Me.otxtOnsertPressDie = ProductionGroup.OnsertPressDie
        Me.otxtDesc = ProductionGroup.description
        Me.txtBaseRollStock = ProductionGroup.OnsertStock
        Me.otxtStockNo = ProductionGroup.stock
        
        Me.otxtStockDesc = ProductionGroup.Stock_Desc
        If ProductionGroup.DigitalLabelParts <> "" Then
            Me.otxtStockNo = Me.otxtStockNo & "  (" & ProductionGroup.DigitalLabelParts & ")"
        End If
    
        Me.otxtScratchStockNo = ProductionGroup.Scratch_Stock
        Select Case ProductionGroup.Scratch_Stock
            Case NOT_APPLICABLE
                Me.otxtScratchStockDesc = ""
            Case Else
                Me.otxtScratchStockDesc = ProductionGroup.Scratch_Stock_Desc
                Me.otxtScratchStockNo.Font.Bold = True
        End Select
        
        Me.otxtOverLaminate = ProductionGroup.OverLaminate
        If Me.otxtOverLaminate = NOT_APPLICABLE Or Me.otxtOverLaminate = TBD Then
            Me.otxtOverLaminateDesc = ""
        Else
            Me.otxtOverLaminate.Font.Bold = True
            Me.otxtOverLaminateDesc = ProductionGroup.OverLaminateDesc
        End If
    Else
        
        lblGroupedRun.Visible = False
        lblPrintDate.Visible = True
        
        Me.otxtClintrakSamples = ProductionRun.Clintrak_Samples
        Me.otxtSampleTypes = ProductionRun.Sample_Number
        Me.otxtQtySamples = ProductionRun.Samples_Requested - ProductionRun.Clintrak_Samples
        Me.otxtQtyRequested = ProductionRun.Qty_Requested

        Me.otxtProducedBy = ProductionRun.Produced_By
        Me.otxtDateProduced = ProductionRun.Produced_Date & " " & gClintrakLocations(CStr(ProductionRun.Clintrak_Location_Id)).Time_Zone_Display
        Me.otxtFormID = ProductionRun.Form_Identification
        Me.RunBarcode = "[" & ProductionRun.Barcode_Id & "]"
        Me.otxtOnsertPressDie = ProductionRun.OnsertDiePartNumber
        Me.otxtStockNo = ProductionRun.stock
        Me.otxtDesc = ProductionRun.LabelDescription
        Me.otxtStockDesc = ProductionRun.Stock_Desc
        Me.txtBaseRollStock = ProductionRun.OnsertStock
        
        If ProductionRun.DigitalLabelParts <> "" Then
            Me.otxtStockNo = Me.otxtStockNo & "  (" & ProductionRun.DigitalLabelParts & ")"
        End If
    

        Select Case ProductionRun.Reorient_Ind
            Case 1
                Me.otxtReorient = YES
            Case Else
                Me.otxtReorient = NOT_APPLICABLE
        End Select
    
    'replacment logic only applies to non-prg prints
        If booReplacement Then
            Me.lblReplacements.Visible = True
            If gReprintFile_Type = "REPLACEMENT" Then
                Me.otxtGroupName = gGroupName
            End If
        Else
            Me.lblReplacements.Visible = False
            Me.otxtGroupName = gGroupName
        End If
        
        Me.lblPrintDate = Me.lblPrintDate & _
            IIf(ProductionRun.PDR_Print_Date = "1/1/1900", NOT_APPLICABLE, ProductionRun.PDR_Print_Date & " " & _
                gClintrakLocations(CStr(ProductionRun.Clintrak_Location_Id)).Time_Zone_Display)
    
        If ProductionRun.PDR_Print_Date <> "1/1/1900" And ProductionRun.PaperworkPrintedBy <> "" Then
            Me.lblPrintDate = Me.lblPrintDate & " by " & ProductionRun.PaperworkPrintedBy
        End If
        
    
        Me.otxtOverLaminate = ProductionRun.OverLaminate
        If Me.otxtOverLaminate = NOT_APPLICABLE Or Me.otxtOverLaminate = TBD Then
            Me.otxtOverLaminateDesc = ""
        Else
            Me.otxtOverLaminateDesc = ProductionRun.OverLaminateDescription
            Me.otxtOverLaminate.Font.Bold = True
        End If

       ' Move the apply text into the description since it was on a 2nd line in the old report
        Me.otxtScratchStockNo = ProductionRun.Scratch_Stock
        Select Case ProductionRun.Scratch_Stock
            Case NOT_APPLICABLE
                Me.otxtScratchStockDesc = ""
            Case Else
                Me.otxtScratchStockNo.Font.Bold = True
                Me.otxtScratchStockDesc = ProductionRun.Scratch_Stock_Description
                If ProductionRun.Apply_ScratchOff = 1 Then
                    Me.otxtScratchStockNo = Me.otxtScratchStockNo & vbCrLf & "(Apply to Labels and Samples)"
                ElseIf ProductionRun.Apply_ScratchOff = 2 Then
                    Me.otxtScratchStockNo = Me.otxtScratchStockNo & vbCrLf & "(Apply to Labels Only)"
                End If
        End Select
        
        If ProductionRun.Prgbarcode <> NOT_APPLICABLE Then
            Me.otxtFormID = Me.otxtFormID & " (" & ProductionRun.Prgbarcode & ") " & ProductionRun.PRGCount & " of " & ProductionRun.TotalPDRs
        End If
    End If
    
    ' This is only seen when viewing the report from the Rand
    ' In Job Execution, the report is embedded in the window and no title can be seen
    Me.Caption = Me.otxtFormID & " (" & Me.otxtIDNumber & ")"

    
    With oRichEdit1
        ' Get the plain text portion of the RichText and set the RTB back to empty.
        If gIsPRGRun Then
            strPlainText = SEE_INDIVIDUAL_PDR
        Else
            .TextRTF = ProductionRun.Special_Inst
            strPlainText = .text
        End If
        If Trim$(strPlainText) = "" Then
            strPlainText = "N/A   " & IIf(ProductionRun.PDR_Print_Date = "1/1/1900", "", ProductionRun.PDR_Print_Date & " " & _
                gClintrakLocations(CStr(ProductionRun.Clintrak_Location_Id)).Time_Zone_Display)
        End If
        .text = ""
        ' Set the style for the RTB and re-enter the plain text
        .SelStart = 0
        .SelFontName = "Arial"
        .SelFontSize = 220
        .SelBold = True
        .SelItalic = False
        .SelUnderline = False
        .SelColor = vbRed
        .SelText = strPlainText
    End With
    
End Sub

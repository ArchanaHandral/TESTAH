VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProdrun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<comment>
' <summary>
' This class stores and retrieves information regarding a PDR.</summary>
'</comment>

Option Explicit
Option Base 1   'used for setting array index lower bound to 1

Private mvarProduction_Run_Id As Long
Private mvarJob_Log_Id As Long
Private mvarRandomization_Id As Long
Private mvarProof_Id As Long
Private mvarClient_Id As Long
Private mvarCampaign_No As String
Private mvarForm_Identification As String
Private mvarLabelDescription As String
Private mvarStock As String
Private mvarQty_Requested As Long
Private mvarSamples_Requested As Long
Private mvarScratch_Proof_Id As Long
Private mvarScratch_Stock As String
Private mvarProduced_By As String
Private mvarProduced_Date As Date
Private mvarSpecial_Inst As String
Private mvarStockDesc As String
Private mvarClient_Name As String
Private mvarScratchStockDesc As String
Private mvarBarcodeId As String
Private mvarOrig_Prod_Run_Id As Long
Private mvarFile_Name As String
Private mvarNumSampleTypes As Integer
Private mvarShipToId As Long
Private mvarShipDescription As String
Private mvarProdDescription As String
Private mvarClintrakSamples As Integer
Private mvarStock_Proof_Id As Long
Private mvarReprint_File_Id As Long
Private mvarApplyScratchOff As Integer
Private mvarCombinedRun As Integer
Private mvarClintrakLocationId As Long
Private mvarCodingFileHeader As String
Private mbooProdRunUpdated As Boolean
Private mbooProdRunFound As Boolean
Private mvarReorientInd As Integer
Private mvarReferenceNo As String
Private mvarPrintDate As String
Private mvarApprovalDate As Date
Private mvarPrintAtPackagerInd As Integer
Private mvarOnsertDiePartNumber As String
Private mvarOnsertDieToolId As Long
Private mvarDigitalLabelParts As String

Public Property Let Production_Run_Id(ByVal vdata As Long)
    mvarProduction_Run_Id = vdata
End Property
Public Property Get Production_Run_Id() As Long
    Production_Run_Id = mvarProduction_Run_Id
End Property

Public Property Let Randomization_Id(ByVal vdata As Long)
    mvarRandomization_Id = vdata
End Property
Public Property Get Randomization_Id() As Long
    Randomization_Id = mvarRandomization_Id
End Property

Public Property Let Job_Log_Id(ByVal vdata As Long)
    mvarJob_Log_Id = vdata
End Property
Public Property Get Job_Log_Id() As Long
    Job_Log_Id = mvarJob_Log_Id
End Property

Public Property Let Client_Id(ByVal vdata As Long)
    mvarClient_Id = vdata
End Property
Public Property Get Client_Id() As Long
    Client_Id = mvarClient_Id
End Property

Public Property Let Proof_Id(ByVal vdata As Long)
    mvarProof_Id = vdata
End Property
Public Property Get Proof_Id() As Long
    Proof_Id = mvarProof_Id
End Property

Public Property Let Samples_Requested(ByVal vdata As Long)
    mvarSamples_Requested = vdata
End Property
Public Property Get Samples_Requested() As Long
    Samples_Requested = mvarSamples_Requested
End Property

Public Property Let Qty_Requested(ByVal vdata As Long)
    mvarQty_Requested = vdata
End Property
Public Property Get Qty_Requested() As Long
    Qty_Requested = mvarQty_Requested
End Property

Public Property Let stock(ByVal vdata As String)
    mvarStock = vdata
End Property
Public Property Get stock() As String
    stock = mvarStock
End Property

Public Property Let LabelDescription(ByVal vdata As String)
    mvarLabelDescription = vdata
End Property
Public Property Get LabelDescription() As String
    LabelDescription = mvarLabelDescription
End Property

Public Property Let Form_Identification(ByVal vdata As String)
    mvarForm_Identification = vdata
End Property
Public Property Get Form_Identification() As String
    Form_Identification = mvarForm_Identification
End Property

Public Property Let Campaign_No(ByVal vdata As String)
    mvarCampaign_No = vdata
End Property
Public Property Get Campaign_No() As String
    Campaign_No = mvarCampaign_No
End Property

Public Property Get Scratch_Proof_Id() As Long
    Scratch_Proof_Id = mvarScratch_Proof_Id
End Property
Public Property Let Scratch_Proof_Id(ByVal vdata As Long)
    mvarScratch_Proof_Id = vdata
End Property

Public Property Get Scratch_Stock() As String
    Scratch_Stock = mvarScratch_Stock
End Property
Public Property Let Scratch_Stock(ByVal vdata As String)
    mvarScratch_Stock = vdata
End Property

Public Property Get Produced_By() As String
    Produced_By = mvarProduced_By
End Property
Public Property Let Produced_By(ByVal vdata As String)
    mvarProduced_By = vdata
End Property

Public Property Get Produced_Date() As Date
    Produced_Date = mvarProduced_Date
End Property

Public Property Get Special_Inst() As String
    Special_Inst = mvarSpecial_Inst
End Property
Public Property Let Special_Inst(ByVal vdata As String)
    mvarSpecial_Inst = vdata
End Property

Public Property Let Prod_Run_Updated(ByVal vdata As Boolean)
    mbooProdRunUpdated = vdata
End Property
Public Property Get Prod_Run_Updated() As Boolean
    Prod_Run_Updated = mbooProdRunUpdated
End Property

Public Property Let Prod_Run_Found(ByVal vdata As Boolean)
    mbooProdRunFound = vdata
End Property
Public Property Get Prod_Run_Found() As Boolean
    Prod_Run_Found = mbooProdRunFound
End Property

Public Property Let Client_Name(ByVal vdata As String)
    mvarClient_Name = vdata
End Property
Public Property Get Client_Name() As String
    Client_Name = mvarClient_Name
End Property

Public Property Let Stock_Desc(ByVal vdata As String)
    mvarStockDesc = vdata
End Property
Public Property Get Stock_Desc() As String
    Stock_Desc = mvarStockDesc
End Property

Public Property Let Scratch_Stock_Description(ByVal vdata As String)
    mvarScratchStockDesc = vdata
End Property
Public Property Get Scratch_Stock_Description() As String
    Scratch_Stock_Description = mvarScratchStockDesc
End Property

Public Property Let Barcode_Id(ByVal vdata As String)
    mvarBarcodeId = vdata
End Property
Public Property Get Barcode_Id() As String
    Barcode_Id = mvarBarcodeId
End Property

Public Property Let Orig_Prod_Run_Id(ByVal vdata As Long)
    mvarOrig_Prod_Run_Id = vdata
End Property
Public Property Get Orig_Prod_Run_Id() As Long
    Orig_Prod_Run_Id = mvarOrig_Prod_Run_Id
End Property

Public Property Let File_Name(ByVal vdata As String)
    mvarFile_Name = vdata
End Property
Public Property Get File_Name() As String
    File_Name = mvarFile_Name
End Property

Public Property Let Sample_Number(ByVal vdata As Integer)
    mvarNumSampleTypes = vdata
End Property
Public Property Get Sample_Number() As Integer
    Sample_Number = mvarNumSampleTypes
End Property

Public Property Let Ship_To_Id(ByVal vdata As Long)
    mvarShipToId = vdata
End Property
Public Property Get Ship_To_Id() As Long
    Ship_To_Id = mvarShipToId
End Property

Public Property Let Ship_Description(ByVal vdata As String)
    mvarShipDescription = vdata
End Property
Public Property Get Ship_Description() As String
    Ship_Description = mvarShipDescription
End Property

Public Property Let Prod_Description(ByVal vdata As String)
    mvarProdDescription = vdata
End Property
Public Property Get Prod_Description() As String
    Prod_Description = mvarProdDescription
End Property

Public Property Let Reference_No(ByVal vdata As String)
    mvarReferenceNo = vdata
End Property
Public Property Get Reference_No() As String
    Reference_No = mvarReferenceNo
End Property

Public Property Let Clintrak_Samples(ByVal vdata As Integer)
    mvarClintrakSamples = vdata
End Property
Public Property Get Clintrak_Samples() As Integer
    Clintrak_Samples = mvarClintrakSamples
End Property

Public Property Let Stock_Proof_Id(ByVal vdata As Long)
    mvarStock_Proof_Id = vdata
End Property
Public Property Get Stock_Proof_Id() As Long
    Stock_Proof_Id = mvarStock_Proof_Id
End Property

Public Property Let Reprint_File_Id(ByVal vdata As Long)
    mvarReprint_File_Id = vdata
End Property
Public Property Get Reprint_File_Id() As Long
    Reprint_File_Id = mvarReprint_File_Id
End Property

Public Property Let Apply_ScratchOff(ByVal vdata As Integer)
    mvarApplyScratchOff = vdata
End Property
Public Property Get Apply_ScratchOff() As Integer
    Apply_ScratchOff = mvarApplyScratchOff
End Property

Public Property Let Reorient_Ind(ByVal vdata As Integer)
    mvarReorientInd = vdata
End Property
Public Property Get Reorient_Ind() As Integer
    Reorient_Ind = mvarReorientInd
End Property

Public Property Let Clintrak_Location_Id(ByVal vdata As Long)
    mvarClintrakLocationId = vdata
End Property
Public Property Get Clintrak_Location_Id() As Long
    Clintrak_Location_Id = mvarClintrakLocationId
End Property

Public Property Let Coding_File_Header(ByVal vdata As String)
    mvarCodingFileHeader = vdata
End Property
Public Property Get Coding_File_Header() As String
    Coding_File_Header = mvarCodingFileHeader
End Property

Public Property Get PDR_Print_Date() As String
    PDR_Print_Date = mvarPrintDate
End Property

Public Property Get ApprovalDate() As Date
    ApprovalDate = mvarApprovalDate
End Property

Public Property Let PrintAtPackager(ByVal vdata As Integer)
    mvarPrintAtPackagerInd = vdata
End Property
Public Property Get PrintAtPackager() As Integer
    PrintAtPackager = mvarPrintAtPackagerInd
End Property

Public Property Let OnsertDiePartNumber(ByVal value As String)
    mvarOnsertDiePartNumber = value
End Property
Public Property Get OnsertDiePartNumber() As String
    OnsertDiePartNumber = mvarOnsertDiePartNumber
End Property

Public Property Let OnsertDieToolId(ByVal value As Long)
    mvarOnsertDieToolId = value
End Property
Public Property Get OnsertDieToolId() As Long
    OnsertDieToolId = mvarOnsertDieToolId
End Property

Public Property Get DigitalLabelParts() As String
    DigitalLabelParts = mvarDigitalLabelParts
End Property


Public Sub LookupRecord()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
   
    mbooProdRunFound = False
        
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        .AddParameter "Prod Run Id", mvarProduction_Run_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ProductionRun"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                mvarProduction_Run_Id = .Recordset!Production_Run_Id
                mvarJob_Log_Id = .Recordset!Job_Log_Id
                mvarRandomization_Id = .Recordset!Randomization_Id
                mvarProof_Id = .Recordset!Proof_Id
                mvarClient_Id = .Recordset!Client_Id
                mvarCampaign_No = Trim$(.Recordset!Campaign_No)
                mvarForm_Identification = Trim$(.Recordset!Form_Identification)
                mvarProdDescription = Trim$(.Recordset!description)
                mvarStock = Trim$(.Recordset!stock)
                mvarQty_Requested = .Recordset!Qty_Requested
                mvarSamples_Requested = .Recordset!Samples_Requested
                mvarScratch_Proof_Id = .Recordset!Scratch_Proof_Id
                mvarScratch_Stock = .Recordset!Scratch_Stock
                mvarProduced_By = .Recordset!Produced_By
                mvarProduced_Date = Format$(.Recordset!Produced_Date, "mm/dd/yyyy")
                mvarSpecial_Inst = IIf(IsNull(.Recordset!Special_Inst), "", .Recordset!Special_Inst)
                mvarBarcodeId = IIf(IsNull(.Recordset!Production_Run_Barcode), "", .Recordset!Production_Run_Barcode)
                mvarFile_Name = IIf(IsNull(.Recordset!File_Name), "", .Recordset!File_Name)
                mvarOrig_Prod_Run_Id = IIf(IsNull(.Recordset!Orig_Prod_Run_Id), 0, .Recordset!Orig_Prod_Run_Id)
                mvarNumSampleTypes = IIf(IsNull(.Recordset!Number_Smp_Types), 0, .Recordset!Number_Smp_Types)
                mvarShipToId = IIf(IsNull(.Recordset!Job_Shipping_Id), 0, .Recordset!Job_Shipping_Id)
                mvarShipDescription = IIf(IsNull(.Recordset!Ship_Description), "", .Recordset!Ship_Description)
                'md added for clintrak samples
                mvarStock_Proof_Id = .Recordset!Stock_Proof_Id
                mvarReprint_File_Id = IIf(IsNull(.Recordset!Reprint_File_Id), 0, .Recordset!Reprint_File_Id)
                mvarApplyScratchOff = IIf(IsNull(.Recordset!Apply_ScratchOff), 0, .Recordset!Apply_ScratchOff)
                mvarCombinedRun = IIf(IsNull(.Recordset!Combined_Run_Ind), 0, .Recordset!Combined_Run_Ind)
                ' DW 2008-017
                'mvarProcessingLocation = IIf(IsNull(.Recordset!Processing_Location), "", .Recordset!Processing_Location)
                mvarClintrakLocationId = .Recordset!Clintrak_Location_Id
                mvarReorientInd = IIf(IsNull(.Recordset!Reorient_Ind), 0, .Recordset!Reorient_Ind)
                mvarCodingFileHeader = .Recordset!Coding_File_Header
                mvarReferenceNo = Trim$(.Recordset!Reference_Number)
                'cc 2008-005 added PDR_Print_Date
                mvarPrintDate = IIf(IsNull(.Recordset!PDR_Print_Date), "1/1/1900", .Recordset!PDR_Print_Date)
                ' kbg 2008-009 added ApprovalDate
                mvarApprovalDate = IIf(IsNull(.Recordset!Approval_Date), "1/1/1900", .Recordset!Approval_Date)
                ' DW 2010-002 added PrintAtPackager
                mvarPrintAtPackagerInd = .Recordset!Print_At_Packager_Ind
                mvarOnsertDiePartNumber = .Recordset!Onsert_Die_Part_Number
                mvarOnsertDieToolId = .Recordset!Onsert_Die_Tool_Id
                mvarDigitalLabelParts = _
                    .Recordset!BK1_Identification & _
                    IIf(.Recordset!BK2_Identification <> "", ", ", "") & .Recordset!BK2_Identification & _
                    IIf(.Recordset!BP_Identification <> "", ", ", "") & .Recordset!BP_Identification
                    
                mbooProdRunFound = True
                .Recordset.MoveNext
            Loop
            '
            .Recordset.Close
            Call GetStockDesc
            If Trim$(mvarScratch_Stock) > "" Then
                Call GetScratchStockDesc
            End If

        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding Production Run Record"
    Resume Exit_this_Sub

End Sub
Public Sub SaveProdRun()
    Dim objData As nADOData.CADOData
    Dim nreturn As Long
                    
    On Error GoTo Error_this_Sub
                                                        
    mbooProdRunUpdated = False
                                                        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
                
        .AddParameter "Prod Run ID", mvarProduction_Run_Id, adInteger, adParamInput
        .AddParameter "Job Log ID", mvarJob_Log_Id, adInteger, adParamInput
        .AddParameter "Randomization ID", mvarRandomization_Id, adInteger, adParamInput
        .AddParameter "Proof ID", mvarProof_Id, adInteger, adParamInput
        .AddParameter "Client ID", mvarClient_Id, adInteger, adParamInput
        .AddParameter "Campaign No.", CheckNulls(mvarCampaign_No), adVarChar, adParamInput
        .AddParameter "Form Desc", CheckNulls(mvarForm_Identification), adVarChar, adParamInput
        .AddParameter "Description", CheckNulls(mvarProdDescription), adVarChar, adParamInput
        .AddParameter "Stock", CheckNulls(mvarStock), adVarChar, adParamInput
        .AddParameter "Qty Requested", mvarQty_Requested, adInteger, adParamInput
        .AddParameter "Samples Requested", mvarSamples_Requested, adInteger, adParamInput
        .AddParameter "Scatch Proof Id", mvarScratch_Proof_Id, adInteger, adParamInput
        .AddParameter "Scratch Stock", CheckNulls(mvarScratch_Stock), adVarChar, adParamInput
        .AddParameter "Produced By", CheckNulls(mvarProduced_By), adVarChar, adParamInput
        .AddParameter "File Name", CheckNulls(mvarFile_Name), adVarChar, adParamInput
        .AddParameter "Orig Prod Run ID", mvarOrig_Prod_Run_Id, adInteger, adParamInput
        .AddParameter "Number Sample Types", mvarNumSampleTypes, adInteger, adParamInput
        .AddParameter "Ship To Id", mvarShipToId, adInteger, adParamInput
        .AddParameter "Ship Description", CheckNulls(mvarShipDescription), adVarChar, adParamInput
        'MD -   added for clintrak samples
        .AddParameter "Stock Proof Id", mvarStock_Proof_Id, adInteger, adParamInput
        .AddParameter "Reprint File Id", mvarReprint_File_Id, adInteger, adParamInput
        .AddParameter "Apply ScratchOff", mvarApplyScratchOff, adInteger, adParamInput
        .AddParameter "Combined Ind", mvarCombinedRun, adInteger, adParamInput
        ' DW 2008-017
        '.AddParameter "Processing Loc", CheckNulls(mvarProcessingLocation), adVarChar, adParamInput
        .AddParameter "Clintrak Location Id", mvarClintrakLocationId, adInteger, adParamInput
        .AddParameter "Reorient Ind", mvarReorientInd, adInteger, adParamInput
        .AddParameter "Reference Number", CheckNulls(mvarReferenceNo), adVarChar, adParamInput
        ' DW 2010-002 added
        .AddParameter "Print At Packager Ind", mvarPrintAtPackagerInd, adInteger, adParamInput
        .AddParameter "Onsert_Die_Tool_Id", mvarOnsertDieToolId, adInteger, adParamInput
        
        .AddParameter "return", "   ", adInteger, adParamOutput
        .AddParameter "identity", "   ", adInteger, adParamOutput
        .AddParameter "barcodeId", "               ", adVarChar, adParamOutput
                
        ' Call the SP to create the recordset
        .ExecuteSP "save_Production_Run", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            mbooProdRunUpdated = False
        Else
            mvarProduction_Run_Id = .GetParameterValue("identity")
            mvarBarcodeId = .GetParameterValue("barcodeId")
            mbooProdRunUpdated = True
            'don't update the links if this is a replacement production run
            If Not booReplacement Then
                If Not update_FileLinks_ProdRunId Then
                    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
                    "Error Occurred Updating File Links with the Prod Run Id"
                End If
            End If
            Call GetStockDesc
            If Trim$(mvarScratch_Stock) > "" Then
                Call GetScratchStockDesc
            End If
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Saving Production Run Record"
    Resume Exit_this_Sub
    
End Sub
Public Function update_FileLinks_ProdRunId() As Boolean
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Function
            
    Dim nreturn As Long
    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "ProdRun Id", mvarProduction_Run_Id, adInteger, adParamInput
        .AddParameter "FileLinks Id", gFileLinksId, adInteger, adParamInput
        .AddParameter "return", "   ", adInteger, adParamOutput

        .ExecuteSP "update_FileLinks_ProdRunId", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            update_FileLinks_ProdRunId = False
        Else
            update_FileLinks_ProdRunId = True
        End If
        
    End With

Exit_this_Function:
    Exit Function
    
Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Updateing File Links Id"
    Resume Exit_this_Function
End Function
Public Sub GetLabelDesc()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        .AddParameter "Proof Id", mvarProof_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_LabelDesc"
        If Not .Recordset.EOF Then
            mvarLabelDescription = Trim$(.Recordset!description)
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Set objData = Nothing
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding Label Desc."
    Resume Exit_this_Sub

End Sub

Public Sub GetStockDesc()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Stock No", mvarStock, adVarChar, adParamInput
            ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_StockDesc"
            
        If Not .Recordset.EOF Then
            mvarStockDesc = Trim$(.Recordset!description)
            '
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding Stock Desc."
    Resume Exit_this_Sub

End Sub
Public Sub GetScratchStockDesc()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Scratch Stock No", mvarScratch_Stock, adVarChar, adParamInput
            ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_StockDesc"
            
        If Not .Recordset.EOF Then
            mvarScratchStockDesc = Trim$(.Recordset!description)
            '
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding Scratch Stock Desc."
    Resume Exit_this_Sub

End Sub

' ***********************************************************
' Deleting Production Run and associated PENDING IRQs 4/10/03
'************************************************************
Public Sub DeleteRecord()
    Dim objData As nADOData.CADOData
    Dim nreturn As Long
                    
    On Error GoTo Error_this_Sub
                                                    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .Connection.BeginTrans
        
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        
        'delete any entry on the Links_PDR_Schedule_Approval table
        .ResetParameters
        .AddParameter "Prod Run Id", mvarProduction_Run_Id, adInteger, adParamInput
        .AddParameter "return", "   ", adInteger, adParamOutput

        .ExecuteSP "delete_LinksPDR_Approval_by_PDR", True
    
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            mbooProdRunUpdated = False
            .Connection.RollbackTrans
            Exit Sub
        End If
        
        'delete any entries on the Production_Run_Client_Fields
        .ResetParameters
        .AddParameter "Prod Run Id", mvarProduction_Run_Id, adInteger, adParamInput
        .AddParameter "return", "   ", adInteger, adParamOutput

        .ExecuteSP "delete_ProductionRunClientFields_by_PDRId", True
    
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            mbooProdRunUpdated = False
            .Connection.RollbackTrans
            Exit Sub
        End If
        

        .ResetParameters
                
        .AddParameter "Prod Run ID", mvarProduction_Run_Id, adInteger, adParamInput
        .AddParameter "return", "   ", adInteger, adParamOutput
                
        ' first delete the production run
        .ExecuteSP "delete_ProductionRun", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            mbooProdRunUpdated = False
            .Connection.RollbackTrans
            Exit Sub
        Else
          
            mbooProdRunUpdated = True
        End If
                
        'next update the links - set the prod run id = 0
        'but only if its not a replacement production run
        If Not booReplacement Then
            .ResetParameters
            .AddParameter "ProdRun Id", 0, adInteger, adParamInput
            .AddParameter "FileLinks Id", gFileLinksId, adInteger, adParamInput
            .AddParameter "return", "   ", adInteger, adParamOutput

            .ExecuteSP "update_FileLinks_ProdRunId", True
        
            .RetrieveParameters
            nreturn = .GetParameterValue("return")
            If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
                mbooProdRunUpdated = False
                .Connection.RollbackTrans
                Exit Sub
            End If
        End If
        
        .Connection.CommitTrans
        
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    gadoConnection.Connection.RollbackTrans
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Deleting Production Run Record"
    Resume Exit_this_Sub
    
End Sub

Public Sub Determine_Clintrak_Samples()
'md added new for Clintrak Samples.
'Determine the True clintrak samples with the prod run id
    Dim objData As nADOData.CADOData
    On Error GoTo PROC_ERR

    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        
        'md changed to a recordset to accomodate SP change for .NET application
        .AddParameter "Prodrunid", mvarProduction_Run_Id, adInteger, adParamInput
         .OpenRecordSetFromSP "get_Samples_Clintrak"
         
         If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                mvarClintrakSamples = .Recordset!quantity
                .Recordset.MoveNext
            Loop
            .Recordset.Close
        End If
    End With
        
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "Determine_ClintrakSamples"
  Resume PROC_EXIT

End Sub

Public Sub UpdateSampleQuantities()
    On Error GoTo Handle_Error
    
    Dim objData As nADOData.CADOData
    Dim lngError As Long
                                                        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        .AddParameter "Production_Run_Id", mvarProduction_Run_Id, adInteger, adParamInput
        .AddParameter "Samples Requested", mvarSamples_Requested, adInteger, adParamInput
        .AddParameter "Number Sample Types", mvarNumSampleTypes, adInteger, adParamInput
        .AddParameter "Error", "   ", adInteger, adParamOutput
        .ExecuteSP "update_Production_Run_Samples", True
        .RetrieveParameters
        lngError = .GetParameterValue("Error")
        If IsNull(lngError) Or Trim$(lngError) = "" Or lngError <> 0 Then
            MsgBox _
                "Error updating PDR Sample Quantities." & vbCrLf & _
                "Please contact IT.", vbCritical
        End If
    End With

Cleanup_Exit:
    Set objData = Nothing
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub

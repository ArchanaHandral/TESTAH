VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProdGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This class stores and retrieves information regarding a PRG.</summary>
'</comment>

Option Explicit

Private mvarProduction_Group_Id As Long
Private mvarJob_Log_Id As Long
Private mvarRandomization_Id As Long
Private mvarProof_Id As Long
Private mvarClient_Id As Long
Private mvarForm_Identification As String
Private mvarDescription As String
Private mvarStock As String
Private mvarQty_Requested As Long
Private mvarSamples_Requested As Long
Private mvarProduced_By As String
Private mvarProduced_Date As Date
Private mvarStockDesc As String
Private mvarClient_Name As String
Private mvarBarcodeId As String
Private mvarFile_Name As String
Private mvarClintrak_Samples As Long
Private mvarEmployeeId As Long
Private mvarStock_Id As Long
'to show scratch stock on computerization order forms for PRGs
Private mvarScratch_Stock As String
Private mvarScratch_Stock_Desc As String
' DW 2008-017 added
Private mvarClintrak_Location_Id As Long
Private mvarOnsertPressDie As String
Private mvarDigitalLabelParts As String
Private mvarPRGReorient As Long
Private mvarPRGShippingSame As Long
Private mvarOverLaminate As String
Private mvarOverLaminateDescription As String
Private mvarPdrShipSeedValue As Long
Private mvarTotalPDRsCount As Long

Public Property Let TotalPDRsCount(ByVal vdata As Long)
    mvarTotalPDRsCount = vdata
End Property
Public Property Get TotalPDRsCount() As Long
    TotalPDRsCount = mvarTotalPDRsCount
End Property

Public Property Let ShipSeedVal(ByVal vdata As Long)
    mvarPdrShipSeedValue = vdata
End Property
Public Property Get ShipSeedVal() As Long
    ShipSeedVal = mvarPdrShipSeedValue
End Property

Public Property Let OverLaminate(ByVal vdata As String)
    mvarOverLaminate = vdata
End Property
Public Property Get OverLaminate() As String
    OverLaminate = mvarOverLaminate
End Property

Public Property Let OverLaminateDesc(ByVal vdata As String)
    mvarOverLaminateDescription = vdata
End Property
Public Property Get OverLaminateDesc() As String
    OverLaminateDesc = mvarOverLaminateDescription
End Property

Public Property Let PrgReorientInd(ByVal vdata As Long)
    mvarPRGReorient = vdata
End Property
Public Property Get PrgReorientInd() As Long
    PrgReorientInd = mvarPRGReorient
End Property

Public Property Let PrgShippingInd(ByVal vdata As Long)
    mvarPRGShippingSame = vdata
End Property
Public Property Get PrgShippingInd() As Long
    PrgShippingInd = mvarPRGShippingSame
End Property

Public Property Let Scratch_Stock_Desc(ByVal vdata As String)
    mvarScratch_Stock_Desc = vdata
End Property
Public Property Get Scratch_Stock_Desc() As String
    Scratch_Stock_Desc = mvarScratch_Stock_Desc
End Property

Public Property Let Scratch_Stock(ByVal vdata As String)
    mvarScratch_Stock = vdata
End Property
Public Property Get Scratch_Stock() As String
    Scratch_Stock = mvarScratch_Stock
End Property

Public Property Let Production_Group_Id(ByVal vdata As Long)
    mvarProduction_Group_Id = vdata
End Property
Public Property Get Production_Group_Id() As Long
    Production_Group_Id = mvarProduction_Group_Id
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

Public Property Let Stock_Id(ByVal vdata As Long)
    mvarStock_Id = vdata
End Property
Public Property Get Stock_Id() As Long
    Stock_Id = mvarStock_Id
End Property

Public Property Let Samples_Requested(ByVal vdata As Long)
    mvarSamples_Requested = vdata
End Property
Public Property Get Samples_Requested() As Long
    Samples_Requested = mvarSamples_Requested
End Property

Public Property Let Clintrak_Samples(ByVal vdata As Long)
    mvarClintrak_Samples = vdata
End Property
Public Property Get Clintrak_Samples() As Long
    Clintrak_Samples = mvarClintrak_Samples
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

Public Property Let description(ByVal vdata As String)
    mvarDescription = vdata
End Property
Public Property Get description() As String
    description = mvarDescription
End Property

Public Property Let Form_Identification(ByVal vdata As String)
    mvarForm_Identification = vdata
End Property
Public Property Get Form_Identification() As String
    Form_Identification = mvarForm_Identification
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

Public Property Let Client_Name(ByVal vdata As String)
    mvarClient_Name = vdata
End Property
Public Property Get Client_Name() As String
    Client_Name = mvarClient_Name
End Property

Public Property Let Employee_Id(ByVal vdata As Long)
    mvarEmployeeId = vdata
End Property
Public Property Get Employee_Id() As Long
    Employee_Id = mvarEmployeeId
End Property

Public Property Let Stock_Desc(ByVal vdata As String)
    mvarStockDesc = vdata
End Property
Public Property Get Stock_Desc() As String
    Stock_Desc = mvarStockDesc
End Property

Public Property Let Barcode_Id(ByVal vdata As String)
    mvarBarcodeId = vdata
End Property
Public Property Get Barcode_Id() As String
    Barcode_Id = mvarBarcodeId
End Property

Public Property Let File_Name(ByVal vdata As String)
    mvarFile_Name = vdata
End Property
Public Property Get File_Name() As String
    File_Name = mvarFile_Name
End Property

Public Property Let Clintrak_Location_Id(ByVal vdata As Long)
    mvarClintrak_Location_Id = vdata
End Property
Public Property Get Clintrak_Location_Id() As Long
    Clintrak_Location_Id = mvarClintrak_Location_Id
End Property

Public Property Get OnsertPressDie() As String
    OnsertPressDie = mvarOnsertPressDie
End Property

Public Property Get DigitalLabelParts() As String
    DigitalLabelParts = mvarDigitalLabelParts
End Property


Public Sub LookupRecord()
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
        .AddParameter "PRG Barcode", mvarForm_Identification, adVarChar, adParamInput
            
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_ProductionGroup_by_PRGBarcode"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                mvarProduction_Group_Id = .Recordset!Production_Group_Id
                mvarJob_Log_Id = .Recordset!Job_Log_Id
                mvarRandomization_Id = .Recordset!Randomization_Id
                mvarStock_Id = .Recordset!Proof_Id
                mvarQty_Requested = .Recordset!Qty_Requested
                mvarSamples_Requested = .Recordset!Samples_Requested
                mvarClintrak_Samples = .Recordset!Clintrak_Samples
                mvarEmployeeId = .Recordset!Created_By
                mvarProduced_Date = Format$(.Recordset!Created_Date, "mm/dd/yyyy")
                mvarBarcodeId = IIf(IsNull(.Recordset!PRG_Barcode), "", .Recordset!PRG_Barcode)
                mvarFile_Name = IIf(IsNull(.Recordset!PRG_FileName), "", .Recordset!PRG_FileName)
                mvarScratch_Stock = .Recordset!Scratch_Stock
                mvarClintrak_Location_Id = .Recordset!Clintrak_Location_Id
                mvarOnsertPressDie = .Recordset!Onsert_Press_Die
                mvarOverLaminate = IIf(IsNull(.Recordset!OverLaminate), "N/A", .Recordset!OverLaminate)
                mvarOverLaminateDescription = IIf(IsNull(.Recordset!OverLaminateDescription), "", .Recordset!OverLaminateDescription)
                mvarPdrShipSeedValue = .Recordset!ShipToId
                mvarTotalPDRsCount = IIf(IsNull(.Recordset!TotalPDRsCount), 0, .Recordset!TotalPDRsCount)
                mvarDigitalLabelParts = _
                    .Recordset!BK1_Identification & _
                    IIf(.Recordset!BK2_Identification <> "", ", ", "") & .Recordset!BK2_Identification & _
                    IIf(.Recordset!BP_Identification <> "", ", ", "") & .Recordset!BP_Identification
                .Recordset.MoveNext
            Loop
            
            .Recordset.Close

            If Trim$(mvarScratch_Stock) <> "N/A" Then
                Call GetPRGScratchStockDesc
            End If
           
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding Production Group Record"
    Resume Exit_this_Sub

End Sub

Public Sub GetPRGLabelDesc()
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
    
        .AddParameter "label Id", gLabelId, adVarChar, adParamInput
            ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_StockDesc"
            
        If Not .Recordset.EOF Then
            mvarDescription = Trim$(.Recordset!description)
            '
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding PRG Label Desc."
    Resume Exit_this_Sub

End Sub

Public Sub GetPRGStockDesc()
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
    "Error Occurred Finding PRG Stock Desc."
    Resume Exit_this_Sub

End Sub

Public Sub GetPRGEmployeeById()
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
    
        .AddParameter "Employee Id", mvarEmployeeId, adInteger, adParamInput
            ' Call the SP to create the recordset
        '.OpenRecordSetFromSP "get_EmployeeNameById"
        ' DW 2008-017 updated to
        .OpenRecordSetFromSP "get_Employee_By_Id"
            
        If Not .Recordset.EOF Then
            mvarProduced_By = Trim$(.Recordset!Last_Name) & ", " & Trim$(.Recordset!First_Name)
            '
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding PRG Creator's Name."
    Resume Exit_this_Sub

End Sub
Public Sub GetPRGLabelId()
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
    
        .AddParameter "Proof Id", mvarForm_Identification, adVarChar, adParamInput
            ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_ProductionRuns_or_ProductionGroup_By_Barcode"
            
        If Not .Recordset.EOF Then
            gLabelId = .Recordset!Label_Id
            
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding PRG Label Id"
    Resume Exit_this_Sub

End Sub

Public Sub GetPRGStockInfo()
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
    
        .AddParameter "Stock Proof Id", mvarStock_Id, adInteger, adParamInput
            ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_LabelProofInfo_ByProofID"
            
        If Not .Recordset.EOF Then
            mvarStock = .Recordset!label_identification
            
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding PRG Stock Info"
    Resume Exit_this_Sub

End Sub

Public Function CheckIfPRGOfReplacements(Prgbarcode As String) As Boolean
    On Error GoTo Error_this_Sub
    
    Dim lngPRGId As Long
    Dim lngOrigPDRId As Long
    Dim objData As nADOData.CADOData
    
    
    CheckIfPRGOfReplacements = False
    lngPRGId = 0
    lngOrigPDRId = 0
            
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "PRG Barcode", Prgbarcode, adVarChar, adParamInput
            ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_ProductionGroup_by_PRGBarcode"
            
        If Not .Recordset.EOF Then
            lngPRGId = .Recordset!Production_Group_Id
            .Recordset.Close
        End If
        
        If lngPRGId <> 0 Then
            .ResetParameters
        
            .AddParameter "PRG Id", lngPRGId, adInteger, adParamInput
            ' Call the SP to create the recordset
            .OpenRecordSetFromSP "get_ProductionRunInfo_for_PDRs_by_ProdGroupId"
                
            If Not .Recordset.EOF Then
                If .Recordset!Orig_Prod_Run_Id <> 0 Then
                    lngOrigPDRId = .Recordset!Orig_Prod_Run_Id
                    .Recordset.Close
                End If
            End If
        
            If lngOrigPDRId <> 0 Then
                .ResetParameters
            
                .AddParameter "PRG Id", lngOrigPDRId, adInteger, adParamInput
                ' Call the SP to create the recordset
                .OpenRecordSetFromSP "get_ProductionRun"
                    
                If Not .Recordset.EOF Then
                    gOriginalPDRBarcode = .Recordset!Production_Run_Barcode
                    CheckIfPRGOfReplacements = True
                    .Recordset.Close
                End If
            End If
            
        End If
    End With

Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    CheckIfPRGOfReplacements = False
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Checking If PRG Of Replacements"
    Resume Exit_this_Sub

End Function

Public Sub GetPRGScratchStockDesc()
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
            mvarScratch_Stock_Desc = Trim$(.Recordset!description)

            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding PRG Scratch Stock Desc."
    Resume Exit_this_Sub

End Sub

Public Sub GetPRGReorient()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "productionGroupId", mvarProduction_Group_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_PRG_Reorient_Value"
            
        If Not .Recordset.EOF Then
            mvarPRGReorient = .Recordset!ReorientInd

            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred in GetPrgReorient()"
    Resume Exit_this_Sub

End Sub

Public Sub GetPRGSameShipping()
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        
        .AddParameter "productionGroupId", mvarProduction_Group_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_PRG_Has_Same_Shipping_Addresses"
            
        If Not .Recordset.EOF Then
            mvarPRGShippingSame = .Recordset!ShippingDifferences
            .Recordset.Close
        End If
    
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred in GetPrgSameShipping()"
    Resume Exit_this_Sub

End Sub

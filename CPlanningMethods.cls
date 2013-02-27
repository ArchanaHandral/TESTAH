VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanningMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object contains various methods for gathering information pertaining to a run.</summary>
'</comment>

Option Explicit
'Private madoData As nADOData.CADOData

Public Function CheckIfCombined(barCode As String) As Boolean
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Function
    
    CheckIfCombined = False
        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = gadoConnection.Connection
        'Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Barcode", barCode, adVarChar, adParamInput
            
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_PRGBarcode_For_PDR_By_PDRBarCode"
        
        If Not .Recordset.EOF Then
            CheckIfCombined = True
        End If
    End With

Exit_this_Function:
    Exit Function
    
Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Checking if Combined"
    Resume Exit_this_Function

End Function

Public Function CheckIfReplacement(barCode As String) As Boolean
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Function
    
    CheckIfReplacement = False
        
    Set objData = New CADOData
    With objData
        Set .Connection = gadoConnection.Connection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Production Run Barcode", barCode, adVarChar, adParamInput
            
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_ReplacementRunInfo_by_ProdRunBarcode"
        
        If Not .Recordset.EOF Then
            CheckIfReplacement = True
            
            gRandomizationId = .Recordset!Randomization_Id
            gProofId = .Recordset!Proof_Id
            gProductionRun_Id = .Recordset!Production_Run_Id
            gLabelId = .Recordset!label_identification
            gQuantity = .Recordset!Qty_Requested
            gReprintFileName = .Recordset!Reprint_File_Name
            gOriginalPDRBarcode = .Recordset!Original_Production_Run_Barcode

        End If
    End With


Exit_this_Function:
    Exit Function
    
Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Checking if Replacement"
    Resume Exit_this_Function

End Function

Public Sub Gather_PDR_Requirements(barCode As String)
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = gadoConnection.Connection
        'Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Barcode", barCode, adVarChar, adParamInput
        
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_PDRLinkInfo_By_Barcode"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Call PlanningList.Add(barCode, .Recordset!Form_Identification, .Recordset!File_Links_ID, .Recordset!Job_Log_Id, .Recordset!Rand_Coding_Number)
                .Recordset.MoveNext
            Loop
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Gathering PDR Information"
    Resume Exit_this_Sub

End Sub

Public Sub GatherReplacementPDRRequirements(barCode As String)
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = gadoConnection.Connection
        'Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Barcode", barCode, adVarChar, adParamInput
        
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_PDRLinkInfo_By_Barcode"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Call PlanningList.Add(barCode, .Recordset!Form_Identification, .Recordset!File_Links_ID, .Recordset!Job_Log_Id, .Recordset!Rand_Coding_Number)
                .Recordset.MoveNext
            Loop
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Gathering PDR Information"
    Resume Exit_this_Sub

End Sub

Public Sub GatherPRGReplacementPDRRequirements(barCode As String)
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = gadoConnection.Connection
        'Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        .AddParameter "Barcode", barCode, adVarChar, adParamInput
        
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_ReplacementPRG_PDR_Info"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Call PlanningList.Add(barCode, .Recordset!Form_Identification, .Recordset!File_Links_ID, .Recordset!Job_Log_Id, .Recordset!Rand_Coding_Number)
                .Recordset.MoveNext
            Loop
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Gathering PDR Information"
    Resume Exit_this_Sub

End Sub

Public Sub CountPDR_Reorient_Ind(barCode As String)
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
    
    Dim PDRCount As Long
    
    gReOrientFlag = False
    
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = gadoConnection.Connection
        'Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        
        .AddParameter "barcode", barCode, adVarChar, adParamInput
        .AddParameter "Count", "      ", adInteger, adParamOutput
        .ExecuteSP "get_Count_PDRstoPRG_ReorientInd", True
                       
        .RetrieveParameters
        PDRCount = .GetParameterValue("Count")
    End With

    If PDRCount > 0 Then
         gReOrientFlag = True
    End If

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Count PDR Reorient Indicator"
    Resume Exit_this_Sub
    
End Sub

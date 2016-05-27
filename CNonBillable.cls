VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNonBillable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarNonBillableId As Long
Private mvarReasonId As Long ' use the property because it takes care of the non-billable notice on frmBegin
Private mvarNotes As String
Private mvarPRNumber As String
Private mvarAuthorizedDate As Date
Private mvarAuthorizedBy As Long
Private mvarIsJobLevel As Boolean

Private mvarOriginalReasonId As Long
Private mvarOriginalPRNumber As String
Private mvarOriginalNotes As String
' A form must be specified for UpdateNonBillableNoticeVisibility
' to work, so only specific instances can show/hide that control.
Private mvarNoticeForm As frmProdPlan

Public Property Get NonBillableId() As Long
    NonBillableId = mvarNonBillableId
End Property

Private Sub UpdateNonBillableNoticeVisibility()
    If Not mvarNoticeForm Is Nothing Then
        mvarNoticeForm.txtNonBillableNotice.Visible = HasReason
    End If
End Sub

Public Property Let NoticeForm(newValue As frmProdPlan)
    Set mvarNoticeForm = newValue
End Property

Public Property Let reasonId(newValue As Long)
    mvarReasonId = newValue
    UpdateNonBillableNoticeVisibility
End Property
Public Property Get reasonId() As Long
    reasonId = mvarReasonId
End Property

Public Property Let notes(newValue As String)
    mvarNotes = newValue
End Property
Public Property Get notes() As String
    If mvarReasonId > 0 Then
        notes = mvarNotes
    Else
        notes = ""
    End If
End Property

Public Property Let PRNumber(newValue As String)
    mvarPRNumber = newValue
End Property

Public Property Get PRNumber() As String
    If mvarReasonId > 0 Then
        PRNumber = mvarPRNumber
    Else
        PRNumber = ""
    End If
End Property

Public Property Get AuthorizedDate() As Date
    AuthorizedDate = mvarAuthorizedDate
End Property

Public Property Let AuthorizedBy(newValue As Long)
    mvarAuthorizedBy = newValue
End Property
Public Property Get AuthorizedBy() As Long
    AuthorizedBy = mvarAuthorizedBy
End Property

Public Property Get isJobLevel() As Boolean
    ' We should really also check IsBillable, because a billable
    ' job would never be job-level non-billable. But, if we do
    ' rely on that, it would allow bugs to slip through because
    ' this should just need to rely on a single variable.
    isJobLevel = mvarIsJobLevel
End Property

Public Property Get HasReason() As Boolean
    ' Because we allow -1, since that allows the combo box to be empty,
    ' some conditions need to check having a reason separately from
    ' being billable. Thus, we provide HasReason: having a reason
    ' selected is separate from being billable, which is reason id 0.
    HasReason = (mvarReasonId > 0)
End Property

Public Property Get IsBillable() As Boolean
    IsBillable = (mvarReasonId = 0)
End Property

Private Property Get DoesRequireOverride() As Boolean
    ' changing the non-billable reason requires an override
    ' don't request an override if:
    ' - replacement PDR reason involves FCS, don't request an override
    ' - replacement PDR goes from FCS to billable
    
    If IsFCSDepartment Then
        DoesRequireOverride = False
        
    ElseIf WasFCSDepartment = True And IsBillable = True Then
        DoesRequireOverride = False
        
    Else
        DoesRequireOverride = mvarOriginalReasonId <> mvarReasonId
        
    End If
End Property

Public Function PerformOverride() As Boolean
    Dim needsOverride As Boolean
    Dim overrideEmployeeId As Long
    
    overrideEmployeeId = 0
    needsOverride = DoesRequireOverride
    
    If needsOverride = True Then
        overrideEmployeeId = frmOverride.GetOverrideUser(SecurityLevels.NonBillableAuthorize, "Changing the non-billable status of a production run requires Non-Billable Authorization.")
        PerformOverride = (overrideEmployeeId > 0)
        
    ElseIf mvarReasonId > 0 And IsFCSDepartment And mvarOriginalReasonId <> mvarReasonId Then
        MsgBox "An FCS non-billable reason does not require an override for replacement PDRs but your name and the current date have been used as the authorization.", vbOKOnly + vbInformation, "Non-Billable Authorization"
        overrideEmployeeId = basGlobals.gApplicationUser.employeeId
        PerformOverride = True
        
    Else
        PerformOverride = True
        
    End If
    
    If overrideEmployeeId > 0 Then
        mvarAuthorizedBy = overrideEmployeeId
        mvarAuthorizedDate = Now
    End If
End Function

Public Sub ResetOriginalData()
    ' Need to call whenever the non-billable info is saved (after the override is given, if one
    ' is) or when the record is loaded. This lets us see if an override needs to be performed.
    mvarOriginalReasonId = mvarReasonId
    mvarOriginalPRNumber = mvarPRNumber
    mvarOriginalNotes = mvarNotes
End Sub

Public Property Get HasChange() As Boolean
    If mvarOriginalReasonId <> mvarReasonId And mvarReasonId <> -1 Then
        HasChange = True
        
    ElseIf mvarOriginalNotes <> mvarNotes Then
        HasChange = True
        
    ElseIf mvarOriginalPRNumber <> mvarPRNumber Then
        HasChange = True
    Else
        HasChange = False
    End If
End Property

Public Sub Delete(objData As nADOData.CADOData)
    ' to be used when deleting a PDR
    ' we don't need to SetNew because the whole thing is going away
    If mvarNonBillableId > 0 Then
        DeleteNonBillable objData
    End If
    Clear True
End Sub

Public Sub Load(data As Recordset)
    Clear False ' reset defaults so we don't need to worry about that
    mvarNonBillableId = data!Non_Billable_Id
    
    If mvarNonBillableId > 0 Then
        reasonId = data!Reason_Id
    Else
        reasonId = 0
    End If
    
    If IsBillable = False Then
        mvarNonBillableId = data!Non_Billable_Id
        mvarNotes = data!notes
        mvarPRNumber = data!PR_Number
        mvarAuthorizedDate = data!Effective_Date
        mvarAuthorizedBy = data!Authorized_By
        mvarIsJobLevel = data!Job_Level_Ind
    End If
    ResetOriginalData
End Sub

Private Sub InsertNewNonBillable(JobId As Long, objData As nADOData.CADOData)
    With objData
        .AddParameter "@in_job_log_id", JobId, adInteger, adParamInput
        .AddParameter "@in_Authorized_By", mvarAuthorizedBy, adInteger, adParamInput
        .AddParameter "@in_Reason_Id", mvarReasonId, adInteger, adParamInput
        .AddParameter "@in_Job_Level_Ind", False, adBoolean, adParamInput
        .AddParameter "@in_PR_Number", CheckEmptyString(mvarPRNumber), adVarChar, adParamInput
        .AddParameter "@in_Notes", CheckEmptyString(mvarNotes), adVarChar, adParamInput
        .AddParameter "@out_Date", "        ", adDate, adParamOutput
        .AddParameter "@out_ident", "    ", adInteger, adParamOutput
        
        .ExecuteSP "insert_Non_Billable", True
        
        .RetrieveParameters
        
        mvarAuthorizedDate = .GetParameterValue("@out_Date")
        mvarNonBillableId = .GetParameterValue("@out_ident")
    End With
End Sub

Private Sub DeleteNonBillable(objData As nADOData.CADOData)
    If mvarIsJobLevel = False Then
        basGlobals.ResetSPCall objData
        objData.AddParameter "@in_Non_Billable_Id", mvarNonBillableId, adInteger, adParamInput
        objData.ExecuteSP "delete_Non_Billable", True
    End If
    
    ' Don't need to reset original data because this is called only by Delete
    ' (by PDR) and DeleteNonBillable, both of which call this themselves.
    Clear False
End Sub

Private Sub UpdateNonBillable(JobId As Long, objData As nADOData.CADOData)
    With objData
        .AddParameter "@in_Non_Billable_Id", mvarNonBillableId, adInteger, adParamInput
        .AddParameter "@in_Authorized_By", mvarAuthorizedBy, adInteger, adParamInput
        .AddParameter "@in_Reason_Id", mvarReasonId, adInteger, adParamInput
        .AddParameter "@in_PR_Number", CheckEmptyString(mvarPRNumber), adVarChar, adParamInput
        .AddParameter "@in_Notes", CheckEmptyString(mvarNotes), adVarChar, adParamInput
        .AddParameter "@in_Job_Level_Ind", False, adBoolean, adParamInput
        .AddParameter "@in_Job_Log_Id", JobId, adInteger, adParamInput
        .AddParameter "@out_Date", "        ", adDate, adParamOutput
        
        .ExecuteSP "update_Non_Billable", True
        
        .RetrieveParameters
        
        mvarAuthorizedDate = .GetParameterValue("@out_Date")
    End With
End Sub

Public Sub Save(JobId As Long, objData As nADOData.CADOData)
    If mvarIsJobLevel = False Then
        objData.ResetParameters
        
        If mvarOriginalReasonId = 0 And mvarReasonId > 0 Then
            InsertNewNonBillable JobId, objData
            
        ElseIf mvarOriginalReasonId <> 0 And mvarReasonId = 0 Then
            DeleteNonBillable objData
            
        ElseIf mvarOriginalReasonId > 0 And mvarReasonId > 0 Then
            UpdateNonBillable JobId, objData
            
        ElseIf mvarOriginalReasonId = 0 And mvarReasonId = 0 Then
            ' do nothing
        End If
    End If
    
    ResetOriginalData
End Sub

Private Sub Clear(resetOriginal As Boolean)
    ' sets defaults, which happens to mean billable
    reasonId = 0
    mvarNonBillableId = 0
    mvarPRNumber = ""
    mvarNotes = ""
    mvarAuthorizedBy = 0
    mvarAuthorizedDate = #1/1/1900#
    mvarIsJobLevel = False
    If resetOriginal = True Then
        ResetOriginalData
    End If
End Sub

Public Sub SetNew(JobId As Long)
    Clear False
    
    If JobId > 0 Then
        GrabJobLevelNonBillableIfExists JobId
    End If
    
    ResetOriginalData
End Sub

Private Sub GrabJobLevelNonBillableIfExists(JobId As Long)
    Dim objData As nADOData.CADOData
    Set objData = basGlobals.CreateNewSPCall
    With objData
        .AddParameter "@in_job_log_id", JobId, adInteger, adParamInput
        
        Dim rs As Recordset
        Set rs = .OpenRecordSetFromSP("get_NonBillable_By_JobLogId")
        
        Do While Not rs.EOF
            If rs!Job_Level_Ind = True Then
                mvarNonBillableId = rs!Non_Billable_Id
                mvarAuthorizedBy = rs!Authorized_By
                mvarAuthorizedDate = rs!Effective_Date
                reasonId = rs!Reason_Id
                mvarIsJobLevel = True
                mvarNotes = rs!notes
                mvarPRNumber = rs!PR_Number
                Exit Do
            End If
            rs.MoveNext
        Loop
    End With
End Sub

Public Property Get IsFCSDepartment() As Boolean
    ' it's ok keeping the REPLACEMENT global in here because it's impossible to change once a PDR is created
    IsFCSDepartment = (basGlobals.gReprintFile_Type = "REPLACEMENT" And basGlobals.GetNonBillableDepartment(mvarReasonId) = "FCS")
End Property

Public Property Get WasFCSDepartment() As Boolean
    ' it's ok keeping the REPLACEMENT global in here because it's impossible to change once a PDR is created
    WasFCSDepartment = (basGlobals.gReprintFile_Type = "REPLACEMENT" And basGlobals.GetNonBillableDepartment(mvarOriginalReasonId) = "FCS")
End Property

Private Sub Class_Initialize()
    Clear True
    Set mvarNoticeForm = Nothing
End Sub

Attribute VB_Name = "basGlobals"
'<comment>
' <summary>
' This module contains globally accessed properties and methods.</summary>
'</comment>

Option Explicit

'declare database access objects
Public gadoConnection As nADOConn.CADOConn
Private madoData As nADOData.CADOData
Public gClintrakLocations As ClintrakCommon.LocationCollection

'declare global variables
Public booReplacement As Boolean
Public gJob_Id As Long
Public gJobNumber As String
Public gProtocol As String
Public gJobNonBillableId As Long
Public gFileLinksId As Long
Public gRandomizationId As Long
Public booNewProdRun As Boolean
Public gClientName As String
Public gClientId As Long
Public gClientRefReqInd As Boolean
Public gClientReqFieldInd As Boolean
Public gReprintFileName As String
Public gProductionRun_Id As Long
Public gLabelId As String
Public gProofId As Long
Public gQuantity As Long
Public gCodingFileName As String
Public gRandDelimiter As String
Public gOriginalPDRBarcode As String
Public gCodingName As String
Public gCodingNumber As Long
Public gGroupNumber As Long
Public gGroupName As String
Public gClientGroupName As String
Public gRandIDNumber As String
Public vdata As String
Public columnNumber As Integer
Public gSampleFileName As String
Public gSampleTypeId As Long
Public writeData As String
Public gSampleFlag As Boolean
Public gLinksSpecInstr As String

'declare global objects
Public ProductionRun As CProdrun
Public mData As CCOLPDRFILES
Public dupData As CCOLdupFiles
Public smpData As CCOLsmpFiles
Public gReprintFile_Type As String         'used to help determine replacement or resupply
Public Planning As CPlanningMethods
Public PlanningList As CColPlanningInfo
Public ProductionGroup As CProdGroup
Public gApplicationUser As ClintrakCommon.ApplicationUser

Public gShowQuarantineTagFlag As Boolean        'used to determine which tags get blanked out on form
Public gReOrientFlag As Boolean

Public ClientReqdFields As CColClientReqdFields

Public Const CNODELIMITER As String = "?"
Public Const BLANK_RPT As String = "BLANK"
Public gRandBarcode As String
Public gOrigPDRNumber As String
Public gStockProofId As Long
Public gStockLabelId As String
Public gBlindProofId As Long
Public gBlindLabelId As String
Public gBlindLamApply As Integer
Public gOnsertDieToolId As Long
Public gOnsertDiePartNumber As String
Public gCodingRepeatCnt As Integer
Public dupSameCodingData As CCOLdupFiles
Public smpSameCodingData As CCOLsmpFiles

Public Const NA As Integer = 3
Public Const LABELS_ONLY_TEXT As String = "Apply to Labels Only"
Public Const LABEL_AND_SAMPLES_TEXT As String = "Apply to Labels and Samples"
Public Const NOT_BLINDED_TEXT As String = "Not Blinded"
Public Const NA_TEXT As String = "N/A"

Public BarcodeInfo As CColBarcodeInfo
Public CCBlindLamApplyCol As Collection
Public BlindLamApplyCol As Collection
'job scheduling dll
Public UpdateSchedule As ScheduleUpdate.CScheduleUpdatemain
' DW 2008-017 added
Public gDomainIconPath As String
' DW 2012-001 added
Public Const DPIRQDELIMITER As String = ","

Public Enum SecurityLevels
    NonBillableAuthorize = 1
End Enum

Public Sub CenterForm(frmIn As Form)
    On Error GoTo PROC_ERR

    frmIn.Move (Screen.Width - frmIn.Width) / 2, (Screen.Height - frmIn.Height) / 2

Proc_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , "CenterForm"
    Resume Proc_EXIT
End Sub

Public Sub getFileLinksInfo()
    
    'this is the first instantiation point of this object so first
    'check to see if it exists, kill it if it does
    If Not madoData Is Nothing Then
        Set madoData = Nothing
    End If
    
    Set madoData = New CADOData
    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
    End With
   
    With madoData
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "File Links Id", gFileLinksId, adInteger, adParamInput
        
        'use a special sp if the coding is the 0 coding
        If gCodingNumber = 0 Then
            .OpenRecordSetFromSP "get_RandCodingLinkCD0byLink_Id"
        Else
            .OpenRecordSetFromSP "get_RandCodingLinkbyLink_Id"
        End If
           
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                gRandomizationId = .Recordset!rand_id
                gProofId = .Recordset!Proof_Id
                gProductionRun_Id = .Recordset!Production_Run_Id
                gLabelId = .Recordset!label_identification
                gCodingFileName = .Recordset!Coding_File_Name
                gQuantity = .Recordset!Rec_Cnt
                gCodingName = .Recordset!Coding_Name
                gCodingNumber = .Recordset!Coding_Number
                gGroupNumber = .Recordset!Group_Number
                gGroupName = .Recordset!Group_Name
                gClientGroupName = .Recordset!Client_Group_Name
                gCodingRepeatCnt = .Recordset!Repeat_Count
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
   
End Sub

Public Sub getReplacementFileLinksInfo()
    
    'this is the first instantiation point of this object so first
    'check to see if it exists, kill it if it does
    If Not madoData Is Nothing Then
        Set madoData = Nothing
    End If
    
    Set madoData = New CADOData
    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
    End With
   
    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "File Links Id", gFileLinksId, adInteger, adParamInput
        
        'use a special sp if the coding is the 0 coding
        If gCodingNumber = 0 Then
            .OpenRecordSetFromSP "get_RandCodingLinkCD0byLink_Id"
        Else
            .OpenRecordSetFromSP "get_RandCodingLinkbyLink_Id"
        End If
           
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                gRandomizationId = .Recordset!rand_id
                gProofId = .Recordset!Proof_Id
                gCodingFileName = .Recordset!Coding_File_Name
                gCodingName = .Recordset!Coding_Name
                gCodingNumber = .Recordset!Coding_Number
                gGroupNumber = .Recordset!Group_Number
                gGroupName = .Recordset!Group_Name
                gClientGroupName = .Recordset!Client_Group_Name
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
   
End Sub

Public Sub GetJobInformation()
    On Error GoTo Error_this_Sub
    
    Dim nReleaseNo As Integer
    Dim nSequenceNo As Integer
    Dim sSequenceNo As String
    Dim nYear As Integer
    Dim sYear As String
    Dim s2DigitYear As String
    Dim nClientNumber As Long
    Dim sClientNumber As String
    
    If Not madoData Is Nothing Then
        Set madoData = Nothing
    End If
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Job Log Id", gJob_Id, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_JobInfo"
            
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                nClientNumber = .Recordset!Client_Number
                nSequenceNo = .Recordset!Sequence_No
                nYear = .Recordset!Year
                nReleaseNo = .Recordset!Release_No
                gProtocol = .Recordset!description
                gJobNonBillableId = .Recordset!Non_Billable_Id
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With

    If nClientNumber > 0 Then
        sClientNumber = CStr(nClientNumber)
        sClientNumber = PadLeftString(sClientNumber, "0", 4)
        sSequenceNo = CStr(nSequenceNo)
        sSequenceNo = PadLeftString(sSequenceNo, "0", 3)
        sYear = CStr(nYear)
        s2DigitYear = Mid$(sYear, 3, 2)
        gJobNumber = sClientNumber & "-" & sSequenceNo & "-" & s2DigitYear & " RL# " & nReleaseNo
    End If
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading the Job Information for the Production Run"
    Resume Exit_this_Sub

End Sub
Public Function PadLeftString( _
  strIn As String, _
  strPadChar As String, _
  intStrLength As Integer) _
  As String
  ' Comments   : Left-pads a string to intStrLength characters for
  '              right justification
  ' Parameters : strIn - String to pad
  '              strPadChar - Character to use to pad
  '              intStrLength - Desired length of string
  ' Returns    : Left padded string
  ' Source    : Total VB SourceBook 5
  '
  On Error GoTo PROC_ERR
  
  PadLeftString = Right$(String$(intStrLength, Left$(strPadChar, 1)) & _
    strIn, intStrLength)
  
Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "PadLeftString"
  Resume Proc_EXIT
  
End Function
Public Function CheckNulls(Teststr As String) As String
    
    Dim tempstr As String
    
    tempstr = Teststr
    
    If IsNull(tempstr) Or Trim$(tempstr) = "" Then
        tempstr = "Null"
    End If
         
    CheckNulls = tempstr
         
End Function

Public Function CheckEmptyString(text As String) As String
    If Len(text) = 0 Then
        CheckEmptyString = " "
    Else
        CheckEmptyString = text
    End If
End Function


Public Sub GetClientName()
    On Error GoTo Error_this_Sub

    gClientName = ""
    gClientId = 0
    gClientRefReqInd = False
    gClientReqFieldInd = False

    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Job Log Id", gJob_Id, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_ClientInfo_byJobLogId"
            
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                gClientName = .Recordset!Client_Name
                gClientId = .Recordset!Client_Id
                gClientRefReqInd = .Recordset!Ref_Req_Ind
                gClientReqFieldInd = .Recordset!Client_Req_Field_Ind
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading the Client Information for the Production Run"
    Resume Exit_this_Sub

End Sub
Public Sub GetRandIDNumber()
    On Error GoTo Error_this_Sub

    gRandIDNumber = ""
    gRandBarcode = ""
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Job Log Id", gRandomizationId, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_Rand_ID_Number"
            
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                gRandIDNumber = .Recordset!Rand_Id_Number
                gRandBarcode = .Recordset!Rand_Barcode
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading the Rand ID for the Production Run"
    Resume Exit_this_Sub

End Sub
Public Sub GetRandDelimiter()
    On Error GoTo Error_this_Sub

    gRandDelimiter = ""
    gLinksSpecInstr = ""
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Randomization Id", gRandomizationId, adInteger, adParamInput
        ' Call the SP to create the recordset
        .OpenRecordSetFromSP "get_RandomizationInfo"
            
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                gRandDelimiter = .Recordset!RNR_Field_Delimeter
                gLinksSpecInstr = .Recordset!Links_Spec_Instr
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
    
    frmProdPlan.rtbLinkInstructions.text = gLinksSpecInstr
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading the Rand Delimiter for the Production Run"
    Resume Exit_this_Sub

End Sub



Public Function GetFileNameFromFilePath(strFilePath As String) As String
  ' Comments  : Returns the name part of a fully qualified file name
  ' Parameters: strFilePath - path to parse
  ' Returns   : File name and extension
  ' Source    : Total VB SourceBook 5
  Dim intCounter As Integer
  Dim strTmp As String
  Dim chrTmp As String * 1

  On Error GoTo PROC_ERR
  
  ' Parse the string
  For intCounter = Len(strFilePath) To 1 Step -1
    ' It its a slash, grab the sub string
    chrTmp = Mid$(strFilePath, intCounter, 1)
    If chrTmp <> "\" Then
      strTmp = chrTmp & strTmp
    Else
      Exit For
    End If
  Next intCounter

  ' Return the value
  GetFileNameFromFilePath = strTmp
  
Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "GetFileNameFromFilePath"
  Resume Proc_EXIT
 
End Function

Public Function CheckQuotes(Teststr As String) As String
    
    Dim pos As Integer
    Dim tempstr As String
    '
    tempstr = Teststr
    pos = InStr(tempstr, "'")
    While pos > 0
        tempstr = Left$(tempstr, pos) & "'" & Right$(tempstr, Len(tempstr) - pos)
        pos = InStr(pos + 2, tempstr, "'")
    Wend
    CheckQuotes = tempstr
     
End Function

Public Sub RecordSetToSSDBComboBox( _
  rstIn As ADODB.Recordset, _
  cboIn As SSDBCombo, _
  ByVal strDisplayColumn As String, _
  Optional ByVal varItemDataColumn As Variant, _
  Optional ByVal strAddtlDisplayColumn As String)
  ' Comments  : Displays the contents of a recordset in
  '             a standard unbound combo box
  ' Parameters: rstIn - recordset to read. Caller must create
  '             cboIn - combo box to load
  '             strDisplayColumn - name of the column in rstIn to display
  '             in the combo box
  '             varItemDataColumn - name of the column in rstIn to load
  '             into the ItemData property of the combo box. The data in
  '             this column MUST be storable in a long integer, and there
  '             must be no 'null' values. Generally this field will be a
  '             long integer Primary Key value associated with the value
  '             to be displayed in the list
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 5
  '
  Dim strIDC As String
  Dim temp As String
  
  On Error GoTo PROC_ERR
  
  ' if a column name is supplied in the varItemDataColumn parameter,
  ' use this as the field name in the recordset to use to supply values
  ' as the ItemData property of the list array
  If Not IsMissing(varItemDataColumn) Then
    strIDC = CStr(varItemDataColumn)
  Else
    Exit Sub
  End If

  'cboIn.RemoveAll
  
  With rstIn
    If .RecordCount <> 0 Then
      Do Until .EOF
        If Trim$(strAddtlDisplayColumn) = "" Then
            temp = "!" & rstIn(strDisplayColumn) & "!,!" & rstIn(strIDC) & "!"
        Else
            temp = "!" & rstIn(strDisplayColumn) & "!,!" & rstIn(strIDC) & "!,!" & rstIn(strAddtlDisplayColumn) & "!"
        End If
        cboIn.AddItem temp
        .MoveNext
      Loop
    End If
    
  End With

Proc_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "RecordSetToSSDBComboBox"
  Resume Proc_EXIT

End Sub

Public Sub SetSSDBComboText(ssdbCbo As SSDBCombo, sTextID As String, Optional sText As String)

    Dim i As Long

    ssdbCbo.MoveFirst
    For i = 0 To ssdbCbo.Rows - 1
        If ssdbCbo.Columns(1).text = sTextID Or (Trim$(sText) > "" And UCase$(ssdbCbo.Columns(0).text) = UCase$(sText)) Then
            ssdbCbo.Bookmark = ssdbCbo.AddItemBookmark(i)
            ssdbCbo.text = ssdbCbo.Columns(0).text
            Exit Sub
        End If
        ssdbCbo.MoveNext
    Next
    
    ' if it gets here then no value found
    ssdbCbo.text = ""
    'For i = 0 To ssdbCbo.Cols - 1
    '    ssdbCbo.Columns(i).Text = ""
    'Next
    
End Sub

Public Sub Read_File(strSource As String)
'
'comments:  reads the first line Production Run file that is associated
'           with this sample
'parameters: strSource - path of file to read
'returns: Nothing
'
Dim lngSourceFile As Long

On Error GoTo PROC_ERR

' Open the source file
lngSourceFile = FreeFile
vdata = ""
Open strSource For Input Access Read As lngSourceFile
Line Input #lngSourceFile, vdata

' Close file
Close lngSourceFile

Proc_EXIT:
    Exit Sub
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Read File"
    Resume Proc_EXIT

End Sub

Public Function CountDelimitedWords( _
  strIn As String, _
  chrDelimit As String) _
  As Integer
  ' Comments  : Returns the number of words in a delimited string
  ' Parameters: strIn - String to count words in
  '             chrDelimit - Character that delimits words in strIn
  ' Returns   : Number of occurrences
  ' Source    : Total VB SourceBook 5
  '
  Dim intWordCount As Integer
  Dim intPos As Integer
  
  On Error GoTo PROC_ERR

  intWordCount = 1
  ' Find the first occurence
  intPos = InStr(strIn, chrDelimit)
  
  Do While intPos > 0
    ' Increment the hit counter
    intWordCount = intWordCount + 1
    ' Loop until no more occurrences
    intPos = InStr(intPos + 1, strIn, chrDelimit)
  Loop

  ' Return the value
  CountDelimitedWords = intWordCount
  
Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "CountDelimitedWords"
  Resume Proc_EXIT

End Function

Public Function GetDelimitedFirstLine( _
  strIn As String, _
  intIndex As Integer, _
  chrDelimit As String, _
  InsertBLANK As Boolean) _
  As String
  ' Comments  : Returns word intIndex in delimited string strIn
  'Modified to return "" if out of range.
  ' Parameters: strIn - String to search
  '             intIndex - Word position to find
  '             chrDelimit - Character used as the delimter
  ' Returns   : nth word
  ' Source    : Total VB SourceBook 5
  '
  Dim count As Integer
  Dim intCounter As Integer
  Dim intStartPos As Integer
  Dim intEndPos As Integer
  
  On Error GoTo PROC_ERR
  
  'checks to see whether the index is larger than the first line of delimited words.
  'md changed to allow any delimeter
  count = CountDelimitedWords(strIn, chrDelimit)
    If count < intIndex Then
        GetDelimitedFirstLine = ""
            GoTo Proc_EXIT
    End If

  ' Set initial values
  intCounter = 1
  intStartPos = 1

  ' Count to the specified index
  For intCounter = 2 To intIndex
    ' Get the new starting position
    intStartPos = InStr(intStartPos, strIn, chrDelimit) + 1
  Next intCounter

  ' Determine the ending position
  intEndPos = InStr(intStartPos, strIn, chrDelimit) - 1
  ' Ending position can't be less than 1
  If intEndPos <= 0 Then
    intEndPos = Len(strIn)
  End If
  
  ' Pull the word out and return it
  GetDelimitedFirstLine = Mid$(strIn, intStartPos, intEndPos - intStartPos + 1)
    ' return "BLANK" if blank
    If InsertBLANK = True Then
        GetDelimitedFirstLine = IIf(GetDelimitedFirstLine = "", BLANK_RPT, GetDelimitedFirstLine)
    End If
  
Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "GetDelimitedFirstLine"
  Resume Proc_EXIT

End Function

Public Function GetDelimitedWord( _
  strIn As String, _
  intIndex As Integer, _
  chrDelimit As String) _
  As String
  ' Comments  : Returns word intIndex in delimited string strIn
  ' Parameters: strIn - String to search
  '             intIndex - Word position to find
  '             chrDelimit - Character used as the delimter
  ' Returns   : nth word
  ' Source    : Total VB SourceBook 5
  '
  Dim intCounter As Integer
  Dim intStartPos As Integer
  Dim intEndPos As Integer
  
  On Error GoTo PROC_ERR

  ' Set initial values
  intCounter = 1
  intStartPos = 1

  ' Count to the specified index
  For intCounter = 2 To intIndex
    ' Get the new starting position
    intStartPos = InStr(intStartPos, strIn, chrDelimit) + 1
  Next intCounter

  ' Determine the ending position
  intEndPos = InStr(intStartPos, strIn, chrDelimit) - 1
  ' Ending position can't be less than 1
  If intEndPos <= 0 Then
    intEndPos = Len(strIn)
  End If
  
  ' Pull the word out and return it
  GetDelimitedWord = Mid$(strIn, intStartPos, intEndPos - intStartPos + 1)
  
Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "GetDelimitedWord"
  Resume Proc_EXIT

End Function

Public Sub Read_SampleFile(strSource As String, smptype As String)
'
'comments: reads the entire existing sample file and populates a collection
'parameters:   strSource - path of file to read
'returns:       nothing
'
Dim lngSourceFile As Long
Dim count As Long

On Error GoTo PROC_ERR
    
    Set mData = New CCOLPDRFILES

' Open the source file
lngSourceFile = FreeFile
vdata = ""
count = 1
    Open strSource For Input Access Read As lngSourceFile
    
    Do Until EOF(lngSourceFile)
        Line Input #lngSourceFile, vdata
        Call mData.Add(vdata, count, smptype)
        count = count + 1
    Loop
    ' Close file
    Close lngSourceFile

Proc_EXIT:
    Exit Sub
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Read Sample File"
    Resume Proc_EXIT
    
End Sub

Public Function GetFilePath(strPath As String) As String
  ' Comments  : Returns file path
  ' Parameters: strPath - string to parse
  ' Returns   : file path
  '
  Dim intCounter As Integer

  On Error GoTo PROC_ERR
  
  ' Parse the string backwards
    For intCounter = Len(strPath) To 1 Step -1
        ' Short-circuit when we reach the slash
        If Mid$(strPath, intCounter, 1) = "\" Then
            Exit For
        End If
    Next intCounter

  ' Return the value
  GetFilePath = Left$(strPath, intCounter - 1)

Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "GetFilePath"
  Resume Proc_EXIT
  
End Function
Public Sub LoadShippingInfo(itemData As Long, ship As TextBox, attn As TextBox, add1 As TextBox, _
                            add2 As TextBox, city As TextBox, state As TextBox, zip As TextBox, add3 As TextBox)
'
'comments: this function calls a stored procedure to load the shipping information
'           classified by the job log id and the selected shipping address in the
'           ship to combo box
'parameters: itemData - Job Shipping Id evaluated from the combo box
'            txtbox - textbox to populate
'returns:   nothing
'
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
        .AddParameter "Job Shipping Id", itemData, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ShipToAddress"

        If Not .Recordset.EOF Then
            ship.text = .Recordset!ShipTo_Description
                                    attn.text = .Recordset!Attn_Description
                                    add1.text = .Recordset!Address_Line_1
                                    add2.text = .Recordset!Address_Line_2
                                    city.text = .Recordset!city
                                    state.text = .Recordset!state
                                    zip.text = .Recordset!zip
                                    add3.text = .Recordset!Address_Line_3       ' DW 2010-002 added
        End If
        .Recordset.Close
    End With
    
End Sub

Public Function ParseEmptyData(dataItem As String)
 '
 'comments: this function checks to see whether the data field is empty or not
 'parameters: dataItem - string to check
 'returns true if there is data
 '
    ParseEmptyData = False

    If IsNull(dataItem) Or Trim$(dataItem) = "" Or Trim$(UCase$(dataItem)) = BLANK_RPT Then
        writeData = ""
    Else
        
        writeData = dataItem
    End If

End Function

Public Sub WriteFile(data As CCOLPDRFILES, smptype As String, strDestination As String)
On Error GoTo Error_this_Sub

Dim lngDestinationFile As Long
Dim i As Long
Dim strLine As String
Dim count As Long

lngDestinationFile = FreeFile
count = CountDelimitedWords(vdata, gRandDelimiter)

            Open strDestination For Output Access Write As lngDestinationFile
            For i = 1 To data.count
                strLine = ""
                If count >= 1 Then
                    ParseEmptyData (data.Item(i).Field1)
                    strLine = strLine & i & "-" & smptype & "-" & i
                End If
        
                If count >= 2 Then
                    ParseEmptyData (data.Item(i).Field2)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 3 Then
                    ParseEmptyData (data.Item(i).Field3)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 4 Then
                    ParseEmptyData (data.Item(i).Field4)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 5 Then
                    ParseEmptyData (data.Item(i).Field5)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 6 Then
                    ParseEmptyData (data.Item(i).Field6)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 7 Then
                    ParseEmptyData (data.Item(i).Field7)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 8 Then
                    ParseEmptyData (data.Item(i).Field8)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 9 Then
                    ParseEmptyData (data.Item(i).Field9)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 10 Then
                    ParseEmptyData (data.Item(i).Field10)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 11 Then
                    ParseEmptyData (data.Item(i).Field11)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 12 Then
                    ParseEmptyData (data.Item(i).Field12)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 13 Then
                    ParseEmptyData (data.Item(i).Field13)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 14 Then
                    ParseEmptyData (data.Item(i).Field14)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 15 Then
                    ParseEmptyData (data.Item(i).Field15)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 16 Then
                    ParseEmptyData (data.Item(i).Field16)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 17 Then
                    ParseEmptyData (data.Item(i).Field17)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 18 Then
                    ParseEmptyData (data.Item(i).Field18)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 19 Then
                    ParseEmptyData (data.Item(i).Field19)
                    strLine = strLine & gRandDelimiter & writeData
                End If
        
                If count >= 20 Then
                    ParseEmptyData (data.Item(i).Field20)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                ' DW increasing # of columns from 20 to 30 based on client supplied data
                If count >= 21 Then
                    ParseEmptyData (data.Item(i).Field21)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 22 Then
                    ParseEmptyData (data.Item(i).Field22)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 23 Then
                    ParseEmptyData (data.Item(i).Field23)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 24 Then
                    ParseEmptyData (data.Item(i).Field24)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 25 Then
                    ParseEmptyData (data.Item(i).Field25)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 26 Then
                    ParseEmptyData (data.Item(i).Field26)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 27 Then
                    ParseEmptyData (data.Item(i).Field27)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 28 Then
                    ParseEmptyData (data.Item(i).Field28)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 29 Then
                    ParseEmptyData (data.Item(i).Field29)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                If count >= 30 Then
                    ParseEmptyData (data.Item(i).Field30)
                    strLine = strLine & gRandDelimiter & writeData
                End If
                
                    Print #lngDestinationFile, strLine
            
            Next i
            Close lngDestinationFile

Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Writing to file"
    Resume Exit_this_Sub
    
End Sub

Public Sub DeleteSample(prod_id As Long, typenum As Integer)
'
'comments: deletes the current selected sample configuration and updates all
'           sample type numbers that trail the deleted sample type
'parameters: none
'returns: nothing
'
On Error GoTo Error_this_Sub

    Dim nreturn As Long
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
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

        .AddParameter "Production Run Id", prod_id, adInteger, adParamInput
        .AddParameter "Type Number", CInt(typenum), adInteger, adParamInput
        
        .AddParameter "return", "   ", adInteger, adParamOutput
        
        .ExecuteSP "delete_SampleTypeInfo", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        
        'md added code to set deletethe sample files variable id delete was
        'successful.
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            .Connection.RollbackTrans
            Exit Sub
        End If
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Deleting Sample"
    Resume Exit_this_Sub

End Sub

Public Sub LookUpSetToDuplicate(rand_id As Long, prod_id As Long, dData As CCOLdupFiles, booAllCodings As Boolean)
'
'comments:  This function looks up all production run id associated with a randomization id
'           (that haven't been run, don't have issued IRQs, and are not grouped)
'           and stores the production id and the coding file paths in a collection
'parameters:rand_id - randomization id, prod_id - current Production Run Id
'           dData - collection to hold the production id and coding file paths
'           booAllCodings - true = get PDRs from all codings for the rand false = only get
'           PDRs for the same coding as the current PDR
'returns:   nothing
'
On Error GoTo Error_this_Sub

    If madoData Is Nothing Then
    Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
    
        .ResetParameters
        .AddParameter "Randomization Id", rand_id, adInteger, adParamInput
        .AddParameter "Production Run Id", prod_id, adInteger, adParamInput
        .AddParameter "All Codings", booAllCodings, adInteger, adParamInput
        
        .OpenRecordSetFromSP "get_ProductionRunSet_NotProcessed"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                dData.Add .Recordset!Production_Run_Id, .Recordset!File_Name, .Recordset!Production_Run_Barcode
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With

Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Looking Up set to duplicate"
    Resume Exit_this_Sub

End Sub

Public Sub LookUpDuplicateSet(prod_id As Long, sdata As CCOLsmpFiles)
'
'comments: This function looks up all sample configurations from the sample_types table
'           and stores data in sdata collection for the passed in production run id
'parameters:    prod_id - production run id, sdata - collection that stores the sample data
'returns:   nothing
'
On Error GoTo Error_this_Sub

    If madoData Is Nothing Then
    Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
    
        .ResetParameters
        .AddParameter "Production Id", prod_id, adInteger, adParamInput
        .AddParameter "Type Number", 0, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
    
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                sdata.Add .Recordset!Production_Run_Id, .Recordset!Type_Number, _
                            .Recordset!Sample_Type, .Recordset!Job_Shipping_Id, .Recordset!quantity, _
                            .Recordset!Sample_Description, .Recordset!Sample_File_Name, .Recordset!notes, .Recordset!sample_type_id
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With

Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Looking Up Duplicate Set"
    Resume Exit_this_Sub

End Sub

Public Sub DuplicateFiles(sdata As CCOLsmpFiles, dData As CCOLdupFiles)
'
'comments: This function uses the sdata collection as a template and duplicates all sample
'           configurations for each production run in the dData collection. The sdata
'           collection also holds the final results as well.
'parameters:    sdata - collection with template sample configurations, dData - collection
'               with production run id's and file names.
'returns:   nothing
'
On Error GoTo Error_this_Sub

Dim i As Long
Dim j As Long
Dim k As Long
Dim deleteCount As Long
Dim strFilename As String
Dim dupcount As Long
Dim smpCount As Long

dupcount = dData.count
smpCount = sdata.count
deleteCount = sdata.count

'md added for Clintrak Samples
'Make a call to do a pre-check of the samples being duplicated.  If the samples being
'duplicated are CLINTRAK and the total samples are large than the Quantity of the coding
'file to be used on the second (dupped) file, then we cannot allow the duplication effort.
    If CheckDupSmplRecCounts(sdata, dData) Then
        MsgBox "Cannot duplicate the samples because there is a Clintrak Sample" & _
        " type with not enough live data to be produced.  Please configure the Samples" & _
        " manually", vbExclamation, "Error Trying to Duplicate Samples"
        gSampleFlag = True
        GoTo Exit_this_Sub
    End If

'goes through all the production files
    For i = 1 To dupcount
        'goes through the number of templates
        For j = 1 To smpCount
            Set mData = New CCOLPDRFILES
            'md added code for clintrak samples
            If sdata.Item(j).sampleType = "CLINTRAK" Then
                Call ReadProcess_File(dData.Item(i).fileName, _
                    sdata.Item(j).quantity, sdata.Item(j).sampleType)
            Else
                Read_File (dData.Item(i).fileName)
                'duplicates for each sample type number
                For k = 1 To sdata.Item(j).quantity
                    Call mData.Add(vdata, k, sdata.Item(j).sampleType)
                Next k
            End If
            strFilename = GetFilePath(sdata.Item(j).smpfileName)
            strFilename = strFilename & "\" & createFileName(dData.Item(i).productionId) & "_" & sdata.Item(j).typeNumber & ".smp"
            'writes the data to file
            'md change pass parms on call to file
            'get the description from lookup table to be used in the creation of the file
             Call WriteFile(mData, GetLookupDesc(sdata.Item(j).sampleType), strFilename)
             sdata.Add dData.Item(i).productionId, sdata.Item(j).typeNumber, sdata.Item(j).sampleType, _
             sdata.Item(j).shipTo, sdata.Item(j).quantity, sdata.Item(j).smpDescription, strFilename, sdata.Item(j).notes, sdata.Item(j).sample_type_id
        Next j
    Next i

'delete the first template items
    For j = 1 To deleteCount
        sdata.Remove (1)
    Next j

Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Duplicating Files"
    Resume Exit_this_Sub

End Sub

Private Function createFileName(prod_id As Long) As String
'
'comments:  This function creates the file names for the production run
'parameters: prod_id - production run id of file to create
'returns: String of the file name
'
On Error GoTo Error_this_Function

    If madoData Is Nothing Then
    Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
    
        .ResetParameters
        .AddParameter "Production Id", prod_id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ProductionRun"

        If Not .Recordset.EOF Then
            createFileName = .Recordset!Production_Run_Barcode
        End If
        .Recordset.Close
    End With

Exit_this_Function:
    Exit Function

Error_this_Function:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Creating file name"
    Resume Exit_this_Function

End Function

Public Sub DuplicateUpdateSmpTable(sdata As CCOLsmpFiles)
'
'comments: This function updates the Sample_Types table after the duplication process
'parameters:    sdata - collection of updated sample configurations to be entered
'returns:   nothing
'
On Error GoTo Error_this_Sub

Dim smpCount As Long
Dim i As Long

smpCount = sdata.count

'deletes all previously associated sample configurations from the table
For i = 1 To smpCount
    Call DeleteSample(sdata.Item(i).productionId, 0)
Next i

Dim nreturn As Long

'updates the table with the sdata collection
    For i = 1 To smpCount
        If madoData Is Nothing Then
            Set madoData = New CADOData
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

            .AddParameter "Sample Type Id", 0, adInteger, adParamInput
            .AddParameter "Production Run Id", sdata.Item(i).productionId, adInteger, adParamInput
            .AddParameter "Type Number", sdata.Item(i).typeNumber, adInteger, adParamInput
            .AddParameter "Sample Type", sdata.Item(i).sampleType, adVarChar, adParamInput
            .AddParameter "Ship To Id", sdata.Item(i).shipTo, adInteger, adParamInput
            .AddParameter "Quantity", sdata.Item(i).quantity, adInteger, adParamInput
            .AddParameter "Sample File Name", sdata.Item(i).smpfileName, adVarChar, adParamInput
            .AddParameter "Sample Description", CheckNulls(sdata.Item(i).smpDescription), adChar, adParamInput
            .AddParameter "Notes", CheckNulls(sdata.Item(i).notes), adVarChar, adParamInput
    
            .AddParameter "return", "   ", adInteger, adParamOutput ' the "   " is for a length value
            .AddParameter "identity", "   ", adInteger, adParamOutput ' the "   " is for a length value
        
            .ExecuteSP "save_SampleFiles", True
        
            .RetrieveParameters
            nreturn = .GetParameterValue("return")
            If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
                .Connection.RollbackTrans
                Exit Sub
            End If
        End With
    Next i

Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Updating Sample Types Table"
    Resume Exit_this_Sub
End Sub

Public Sub DuplicateUpdateProdTable(prod_id As Long)
'
'comments: This function updates the Production_Runs table after the duplication
'          Process
'parameters: prod_id - production run id of entry to update
'returns:   nothing
'
On Error GoTo Error_this_Sub

Dim Sampletotal As Long
Dim QtyTotal As Long
Dim nreturn As Long

Sampletotal = 0
QtyTotal = 0

    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If
    
'retrieves the sample total and the quantity total
    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly

        .AddParameter "Production Run Id", prod_id, adInteger, adParamInput
        .AddParameter "Type Number", 0, adInteger, adParamInput
        .OpenRecordSetFromSP "get_SampleTypeInfo"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Sampletotal = Sampletotal + 1
                QtyTotal = QtyTotal + .Recordset!quantity
            .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With

'saves the new totals to the Production Runs table
    If Sampletotal >= 1 Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
            .ResetParameters
            .CommandType = adCmdStoredProc
            .CursorType = adOpenForwardOnly
    
            .AddParameter "Production Run Id", prod_id, adInteger, adParamInput
            .AddParameter "Samples Requested", QtyTotal, adInteger, adParamInput
            .AddParameter "Number Sample Types", Sampletotal, adInteger, adParamInput
            .AddParameter "return", "   ", adInteger, adParamOutput
            .ExecuteSP "save_Duplicate_ProdRun", True
        
            .RetrieveParameters
            nreturn = .GetParameterValue("return")
            If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
                .Connection.RollbackTrans
                Exit Sub
            End If
        End With
    End If
    
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Updating Production Runs Table"
    Resume Exit_this_Sub
    
End Sub

Public Function CheckShippingExist(job_id As Long) As Boolean
'
'comments:  This function checks to see that the shipping address is selected
'parameters:    None
'returns:   true if shipping address is selected
'
Dim dataCount As Long
dataCount = 0
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
        .RowsetSize = 1
   
        ' Call the SP to create the resultset
        .AddParameter "Job Log Id", job_id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ShipToCombo"
    
        Do While Not .Recordset.EOF
            dataCount = dataCount + 1
        .Recordset.MoveNext
        Loop
        
        .Recordset.Close
    
    End With

    If dataCount < 1 Then
        CheckShippingExist = False
    Else
        CheckShippingExist = True
    End If
    
End Function

Public Sub LoadShipToCombo2Column(cboControl As SSDBCombo, Optional sColumnDisplay As String)
'
'comments: this function loads a combobox with the shipping addresses associated with the
'          production run job log id.
'parameters:    cboControl - combobox which to populate data
'               sColumnDisplay - name of the column in rstIn to display
'               in the combo box
'returns:        nothing
'
    If Trim$(sColumnDisplay) = "" Or IsNull(sColumnDisplay) Then sColumnDisplay = "ShipTo_Description"
    
    ' load the combo box
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
        .RowsetSize = 1
   
        ' Call the SP to create the resultset
        .AddParameter "Job Log Id", gJob_Id, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ShipToCombo"
    
        If Not .Recordset.EOF Then
            Call RecordSetToSSDBComboBox(.Recordset, cboControl, sColumnDisplay, "Job_Shipping_Id", "Attn")
        End If
        
        .Recordset.Close
    
    End With
    
End Sub

Public Sub ReadProcess_File(FileSource As String, procquantity As Long, cmbotype As String)
'md new code for clintrak samples
'comments:  reads the file and process
'parameters: FileSource - path of file to read
'            procquantity - what position in the file to start the read
'            cmbotype - the selected sample type
'
Dim lngSourceFile As Long
Dim i As Long

On Error GoTo PROC_ERR

' Open the source file
lngSourceFile = FreeFile

vdata = ""

Open FileSource For Input Access Read As lngSourceFile

'first initialize the existing collection grid then populate with count
Set mData = New CCOLPDRFILES

' can only read the file up to the total number of file entries, so replace
If procquantity > gQuantity Then
    Line Input #lngSourceFile, vdata
    For i = 1 To procquantity
        Call mData.Add(vdata, i, cmbotype)
    Next i
Else
    For i = 1 To procquantity
        Line Input #lngSourceFile, vdata
        Call mData.Add(vdata, i, cmbotype)
    Next i
End If

' Close file
Close lngSourceFile

Proc_EXIT:
    Exit Sub
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "ReadProcess File"
    Resume Proc_EXIT

End Sub

Private Function GetLookupDesc(SmplText As String) As String

'md added new for Clintrak Samples
'Get's the lookup description for the abbreviation needed for the Samples to be printed.

On Error GoTo PROC_ERR

     If madoData Is Nothing Then
        Set madoData = New CADOData
            With madoData
                ' DW 2010-002 added
                Set .Connection = GetDBConnection
            End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        
        .AddParameter "Value", SmplText, adChar, adParamInput
        .AddParameter "Type", "SPTY", adChar, adParamInput
        .OpenRecordSetFromSP "get_SPTYLookupId"
        
        If Not .Recordset.EOF Then
            GetLookupDesc = .Recordset!description
        End If
        .Recordset.Close
    End With
   
   
Proc_EXIT:
    Exit Function
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Getting Lookup Description"
    Resume Proc_EXIT

End Function

Private Function CheckDupSmplRecCounts(sdata As CCOLsmpFiles, dData As CCOLdupFiles) As Boolean

'md added for Clintrak Samples
'This function will call seperate functions to determine if there are any Clintrak samples
'in the original PDR (one being dupped).  If Clinrak samples exist, it calls the next
'function to check the record counts of the coding files to the Clintrak samples.  If
'the record count comes back true, then this function is TRUE and we will will not be able
'to duplicate the Clintrak samples.

    On Error GoTo error_function

    CheckDupSmplRecCounts = False

    'check for the existance of any CLINTRAK samples. if clintrak samples move on
    'to compare the counts otherwise, exit
    If Not CheckForClintrakSmpl(sdata) Then
        GoTo exit_function
    End If

    'call to get the record counts for comparison since there are Clintrak samples
    If GetClintrakRecordCnts(sdata, dData) Then
        CheckDupSmplRecCounts = True
    End If

exit_function:
    Exit Function
        
error_function:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Checking for Duplicate Sample Counts"
    Resume exit_function

End Function

Private Function CheckForClintrakSmpl(sdata As CCOLsmpFiles) As Boolean

'md added for Clintrak Samples
'This function will spin through all the sample types of the original PDR looking
'for Clintrak samples.  If so, then this function is TRUE.
    
    On Error GoTo error_function

    Dim OrigSmpCnt As Long
    Dim i As Long
    
    CheckForClintrakSmpl = False
    OrigSmpCnt = sdata.count

    For i = 1 To OrigSmpCnt
        If sdata.Item(i).sampleType = "CLINTRAK" Then
            CheckForClintrakSmpl = True
        End If
    Next i
    
exit_function:
    Exit Function
        
error_function:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Checking for Clintrak Samples"
    Resume exit_function

End Function

Private Function GetClintrakRecordCnts(sdata As CCOLsmpFiles, dData As CCOLdupFiles) As Boolean

'md added for Clintrak samples
'This function will take the original PDR (one being dupped) and get the total samples
'requested.  Then it spins through the PDR's that are being dupped to to see if the coding
'file total of those PDR's is less than the total samples.  If so, set to TRUE becaue
'we cannot do the duplication of Clintrak Samples.

 On Error GoTo PROC_ERR
 
    Dim OrigSmplRequested As Long
    Dim OrigProdId As Long
    Dim dupcount As Long
    Dim i As Long
 
    GetClintrakRecordCnts = False
    'there can only be one Original PDR to be copied from
    OrigProdId = sdata.Item(1).productionId
    
    dupcount = dData.count
 
    If madoData Is Nothing Then
        Set madoData = New CADOData
            With madoData
                ' DW 2010-002 added
                Set .Connection = GetDBConnection
            End With
    End If

    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
        
        .AddParameter "Prodid", OrigProdId, adInteger, adParamInput
        .OpenRecordSetFromSP "get_ProductionRun"
        
        If Not .Recordset.EOF Then
            OrigSmplRequested = .Recordset!Samples_Requested
        End If
        .Recordset.Close
    End With
    
    For i = 1 To dupcount
        With madoData
              
            .ResetParameters
              
            .AddParameter "Prodid", dData.Item(i).productionId, adInteger, adParamInput
            .OpenRecordSetFromSP "get_ProductionRun"
        
            If Not .Recordset.EOF Then
              If OrigSmplRequested > .Recordset!Qty_Requested Then
                GetClintrakRecordCnts = True
                GoTo Proc_EXIT
              Else
                .Recordset.Close
              End If
            Else
              .Recordset.Close
            End If
            
        End With
    Next i
       
Proc_EXIT:
    Exit Function
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "Error Getting Clintrak Record Counts"
    Resume Proc_EXIT

End Function

Public Function DeleteSample_By_SmplID(Smpl_id As Long) As Boolean
'
'comments: deletes the current selected sample configuration by id
'

On Error GoTo Error_this_Sub

    Dim nreturn As Long
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
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

        .AddParameter "Sample Id", Smpl_id, adInteger, adParamInput
                
        .AddParameter "return", "   ", adInteger, adParamOutput
        
        .ExecuteSP "delete_SampleType_By_Id", True
        
        .RetrieveParameters
        nreturn = .GetParameterValue("return")
        
        'md added code to set deletethe sample files variable id delete was
        'successful.
        If IsNull(nreturn) Or Trim$(nreturn) = "" Or nreturn <> 0 Then
            .Connection.RollbackTrans
            DeleteSample_By_SmplID = False
        Else
            DeleteSample_By_SmplID = True
        End If
    End With

Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Deleting Sample By ID"
    Resume Exit_this_Sub

End Function

Public Sub LoadClientLabelFields(ClientID As Long, ProductionRunId As Long)
    Dim oValue As CClientReqdField
    
    On Error GoTo Error_this_Sub
    
    Set ClientReqdFields = New CColClientReqdFields
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If
        
    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .RowsetSize = 1
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
            
        .AddParameter "Client Id", ClientID, adInteger, adParamInput
        .AddParameter "Production Run Id", ProductionRunId, adInteger, adParamInput
        .OpenRecordSetFromSP "get_PDR_ClientRequired_FieldsValues"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Set oValue = ClientReqdFields.Add()
                With oValue
                    .Client_Required_Field_Name = madoData.Recordset!Client_Required_Field_Name
                    .Production_Run_Client_Fields_Id = madoData.Recordset!Production_Run_Client_Fields_Id
                    .Field_Name_Value = madoData.Recordset!Field_Name_Value
                End With

                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
        
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Exit_this_Sub
    
End Sub

Public Sub LoadSmpTypeCombo(sLookupType As String, cboControl As SSDBCombo, Optional sColumnDisplay As String)
    
    If Trim$(sColumnDisplay) = "" Or IsNull(sColumnDisplay) Then sColumnDisplay = "Value"
    
    ' load the combo box
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
        .RowsetSize = 1
   
        ' Call the SP to create the resultset
        .AddParameter "Lookup Value", sLookupType, adVarChar, adParamInput
        .OpenRecordSetFromSP "get_LookupValues"
    
        If Not .Recordset.EOF Then
            Call RecordSetToSSDBComboBox(.Recordset, cboControl, sColumnDisplay, "lookup_id", "description")
        End If
        
        .Recordset.Close
    
    End With

End Sub

Public Sub DuplicateFilesCopyMods(sdata As CCOLsmpFiles, dData As CCOLdupFiles)
'
'comments: This function uses the sdata collection as a template and duplicates all sample
'           configurations for each production run in the dData collection. The sdata
'           collection also holds the final results as well.
'parameters:    sdata - collection with template sample configurations, dData - collection
'               with production run id's and file names.
'returns:   nothing
'
On Error GoTo Error_this_Sub

Dim i As Long
Dim j As Long
Dim k As Long
Dim deleteCount As Long
Dim strFilename As String
Dim dupcount As Long
Dim smpCount As Long
Dim strCodingData As String
Dim strExistSampleData As String
Dim arrCodingData() As String
Dim arrExistSampleData() As String

dupcount = dData.count
smpCount = sdata.count
deleteCount = sdata.count

'for Clintrak Samples
'Make a call to do a pre-check of the samples being duplicated.  If the samples being
'duplicated are CLINTRAK and the total samples are large than the Quantity of the coding
'file to be used on the second (dupped) file, then we cannot allow the duplication effort.
    If CheckDupSmplRecCounts(sdata, dData) Then
        MsgBox "Cannot duplicate the samples because there is a Clintrak Sample" & _
        " type with not enough live data to be produced.  Please configure the Samples" & _
        " manually", vbExclamation, "Error Trying to Duplicate Samples"
        gSampleFlag = True
        GoTo Exit_this_Sub
    End If

'goes through all the production files
    For i = 1 To dupcount
        Read_File (dData.Item(i).fileName)
        strCodingData = vdata
        arrCodingData = Split(strCodingData, gRandDelimiter)
        If UBound(arrCodingData) - gCodingRepeatCnt + 1 <= 0 Then
            MsgBox "There are more repeats defined for " & dData.Item(i).productionRun_Barcode & _
                    "'s Coding than its Coding File contains." & _
                    vbCrLf & "The Samples will be duplicated using the normal duplication method instead of " & _
                    "copying the manual modifications.", vbCritical + vbOKOnly, "Error"
        End If

        'goes through the number of templates
        For j = 1 To smpCount
            Set mData = New CCOLPDRFILES
            'md added code for clintrak samples
            If sdata.Item(j).sampleType = "CLINTRAK" Then
                Call ReadProcess_File(dData.Item(i).fileName, _
                    sdata.Item(j).quantity, sdata.Item(j).sampleType)
            Else
                Read_File (sdata.Item(j).smpfileName)
                strExistSampleData = vdata
                arrExistSampleData = Split(strExistSampleData, gRandDelimiter)
                vdata = strCodingData
'
                If UBound(arrCodingData) = UBound(arrExistSampleData) Then
                    ' DW 2008-017 aka Karen
                    Call ReadProcessFileKeepMod(sdata.Item(j).smpfileName, sdata.Item(j).quantity, sdata.Item(j).sampleType, strCodingData)
                Else
                    MsgBox "The number of columns for Sample Type " & j & _
                            " does not match the number of columns in " & _
                            dData.Item(i).productionRun_Barcode & "." & vbCrLf & _
                            "Sample Type " & j & " will be duplicated using the normal " & _
                            "duplication method instead of " & "copying the manual modifications.", _
                            vbCritical + vbOKOnly, "Error"
                                 
                    'duplicates for each sample type number
                    For k = 1 To sdata.Item(j).quantity
                      Call mData.Add(vdata, k, sdata.Item(j).sampleType)
                    Next k
                End If
'

            End If
            strFilename = GetFilePath(sdata.Item(j).smpfileName)
            strFilename = strFilename & "\" & createFileName(dData.Item(i).productionId) & "_" & sdata.Item(j).typeNumber & ".smp"
            'writes the data to file
            'md change pass parms on call to file
            'get the description from lookup table to be used in the creation of the file
             Call WriteFile(mData, GetLookupDesc(sdata.Item(j).sampleType), strFilename)
             sdata.Add dData.Item(i).productionId, sdata.Item(j).typeNumber, sdata.Item(j).sampleType, _
             sdata.Item(j).shipTo, sdata.Item(j).quantity, sdata.Item(j).smpDescription, strFilename, sdata.Item(j).notes, sdata.Item(j).sample_type_id
        Next j
    Next i

'delete the first template items
    For j = 1 To deleteCount
        sdata.Remove (1)
    Next j

Exit_this_Sub:
    Exit Sub

Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Duplicating Files"
    Resume Exit_this_Sub

End Sub

Public Function DirExists(strDir As String) As Boolean
  ' Comments  : Determines if the named directory exists
  ' Parameters: strDir - Directory to check
  ' Returns   : True if the directory exists, False otherwise
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  DirExists = Len(Dir$(strDir & "\.", vbDirectory)) > 0
  
Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "DirExists"
  Resume Proc_EXIT
  
End Function

Public Function FileExists(strDest As String) As Boolean
  ' Comments  : Determines if the named file exists
  ' Parameters: strDest - File to check
  ' Returns   : True if the file exists, false otherwise
  ' Source    : Total VB SourceBook 6
  '
  Dim intLen As Integer

  If strDest <> vbNullString Then
    On Error Resume Next
    intLen = Len(Dir$(strDest))
    On Error GoTo PROC_ERR
    FileExists = (Not Err And intLen > 0)
  Else
    FileExists = False
  End If

Proc_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , _
    "FileExists"
  Resume Proc_EXIT
  
End Function

Public Function Determine_If_PDR_HasRun() As Boolean
'
'md new for clintrak samples
'comments: this calls Production Run Details to see if any exist
'           it checks the CLPS_Verification table as well
'
On Error GoTo Error_this_Sub

    Dim nCount As Long
    Dim objData As nADOData.CADOData
    
    Determine_If_PDR_HasRun = False
        
    Set objData = New CADOData
    With objData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .ResetParameters
        
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
                    
        .AddParameter "PDR ID", gProductionRun_Id, adInteger, adParamInput
        .AddParameter "count", "   ", adInteger, adParamOutput
              
        .ExecuteSP "get_ProductionRun_DetailsCount", True
        
        .RetrieveParameters
        nCount = .GetParameterValue("count")
        If nCount <> 0 Then
            Determine_If_PDR_HasRun = True
        Else
            .ResetParameters
            .AddParameter "PDR ID", gProductionRun_Id, adInteger, adParamInput
            .AddParameter "Run Type Indicator", "PDR", adVarChar, adParamInput
            
            .OpenRecordSetFromSP "get_CLPS_Verification_Details_By_Run_ID_And_Run_Type"
            
            If Not .Recordset.EOF Then
                Determine_If_PDR_HasRun = True
            End If
        End If
        
        
    End With
            
Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Counting Production Runs"
    Resume Exit_this_Sub

End Function

Public Sub LoadBarcodeInfo(ProductionRunId As Long)
    Dim oValue As CBarcodeInfo
    
    On Error GoTo Error_this_Sub
    
    Set BarcodeInfo = New CColBarcodeInfo
    
    If madoData Is Nothing Then
        Set madoData = New CADOData
        With madoData
            ' DW 2010-002 added
            Set .Connection = GetDBConnection
        End With
    End If
        
    With madoData
        ' DW 2010-002 added
        Set .Connection = GetDBConnection
        .RowsetSize = 1
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
            
        .AddParameter "Production Run Id", ProductionRunId, adInteger, adParamInput
        .OpenRecordSetFromSP "get_Rand_Coding_Data_Barcodes_By_ProdRunId"
        
        If Not .Recordset.EOF Then
            Do Until .Recordset.EOF
                Set oValue = BarcodeInfo.Add()
                With oValue
                    .BarcodeFields = madoData.Recordset!Data_Column_Configuration
                    .BarcodeFValue = IIf(madoData.Recordset!FValue_Sequence_No = 0, "", "F" & madoData.Recordset!FValue_Sequence_No)
                    .BarcodeDesc = madoData.Recordset!Data_Column_Description
                End With

                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
        
Exit_this_Sub:
    Exit Sub
    
Error_this_Sub:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Exit_this_Sub
    
End Sub

Public Function DeterminePDROnPKS(Optional SamplesFlag As String = " ") As Boolean
    On Error GoTo Error_this_Sub
    
    DeterminePDROnPKS = False
        
    If madoData Is Nothing Then
        Set madoData = New CADOData
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
                    
        .AddParameter "Production Run ID", ProductionRun.Production_Run_Id, adInteger, adParamInput
        .AddParameter "Samples Flag", SamplesFlag, adVarChar, adParamInput
        
        .OpenRecordSetFromSP "get_PackingSlip_By_ProductionRunId"
        
        If Not .Recordset.EOF Then
            DeterminePDROnPKS = True
        End If
        .Recordset.Close
    End With
                      
Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Checking If PDR On PKS"
    Resume Exit_this_Sub

End Function

' DW 2008-017 aka Karen
Public Sub ReadProcessFileKeepMod(FileSource As String, qty As Long, sampleType As String, CodingData As String)

'comments:  reads the file and process
'parameters: FileSource - path of file to read
'            Qty - number of lines in the file to read
'            SampleType - the selected sample type
'            CodingData - first line of data from the coding file.  this is used for the
'                       non-repeat data
Dim lngSourceFile As Long
Dim i As Long
Dim n As Long
Dim arrCodingData() As String
Dim strExistSampleData As String
Dim arrExistSampleData() As String
Dim strNewData As String
 
On Error GoTo PROC_ERR

arrCodingData = Split(CodingData, gRandDelimiter)

' Open the source file
lngSourceFile = FreeFile

vdata = ""

Open FileSource For Input Access Read As lngSourceFile

'first initialize the existing collection grid then populate with count
Set mData = New CCOLPDRFILES

For i = 1 To qty
    Line Input #lngSourceFile, vdata

    strExistSampleData = vdata
    arrExistSampleData = Split(strExistSampleData, gRandDelimiter)
    vdata = CodingData
    strNewData = ""


    If UBound(arrCodingData) - gCodingRepeatCnt + 1 > 0 Then
        For n = 0 To UBound(arrCodingData)
            If n < UBound(arrCodingData) - gCodingRepeatCnt + 1 Then
                strNewData = strNewData & arrCodingData(n) & gRandDelimiter
            Else
                If n < UBound(arrCodingData) Then
                    strNewData = strNewData & arrExistSampleData(n) & gRandDelimiter
                Else
                    strNewData = strNewData & arrExistSampleData(n)
                End If
            End If
        Next n
        vdata = strNewData
    End If

    Call mData.Add(vdata, i, sampleType)

Next i

' Close file
Close lngSourceFile

Proc_EXIT:
    Exit Sub

PROC_ERR:

    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "ReadProcessFileKeepMod"
    Resume Proc_EXIT

End Sub

' DW 2010-002 In the event the user has un-docked his/her workstation for wireless freedom
Public Function GetDBConnection() As Connection
    On Error GoTo Connection_Test_Failed
    ' Test connection for General network error's etc.
    gadoConnection.Connection.Execute " "
Cleanup_Exit:
    Set GetDBConnection = gadoConnection.Connection
    Exit Function
Connection_Test_Failed:
    On Error GoTo Handle_Error
    If gadoConnection.Connection.state <> adStateClosed Then
        gadoConnection.Connection.Close
    End If
    gadoConnection.Connection.Open
    Resume Cleanup_Exit
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Function


'<comment>
' <summary>Retrieves the current values of the Label Proof fields that could have changed since the PDR was created.</summary>
'</comment>
Public Sub GetLabelCurrentValues()
    Dim objData As nADOData.CADOData
    On Error GoTo Handle_Error
        
    Set objData = New nADOData.CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        .AddParameter "Proof_Id", gProofId, adInteger, adParamInput
        .OpenRecordSetFromSP "get_LabelProofPDRInfo_ByProofID"
        If Not .Recordset.EOF Then
            gStockProofId = .Recordset!Stock_Proof_Id
            gStockLabelId = .Recordset!Stock_Label_Id
            gBlindProofId = .Recordset!Blinding_Proof_Id
            gBlindLabelId = .Recordset!Blinding_Label_Id
            gOnsertDieToolId = .Recordset!Onsert_Die_Tool_Id
            gOnsertDiePartNumber = .Recordset!Onsert_Tooling_Die
            
            If gBlindProofId = 0 Then
                gBlindLamApply = 2
            Else
                gBlindLamApply = .Recordset!Blinding_Panel_Apply
            End If
            .Recordset.Close
        End If
    End With

Cleanup_Exit:
    Set objData = Nothing
    Exit Sub
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub

Public Function GetNumberLinesInFile(fileName As String) As Long
    Dim lngSourceFile As Long
    Dim strData As String
    
    On Error GoTo Handle_Error

    GetNumberLinesInFile = 0
    
    ' Open the source file
    lngSourceFile = FreeFile
    Open fileName For Input Access Read As lngSourceFile
   
    Do Until EOF(lngSourceFile)
        Line Input #lngSourceFile, strData
        
        GetNumberLinesInFile = GetNumberLinesInFile + 1
    Loop
        
    ' Close file
    Close lngSourceFile
    
Cleanup_Exit:
    Exit Function
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit

End Function

Public Function GetLineOfData(fileName As String, LineNumber As Long) As String
'
'comments:  reads the specified line of data in the specified file
'parameters:    FileName - path of file to read
'               LineNumber - the number of the line of data to retrieve
'returns: String of data from the specified line in the file

    Dim lngSourceFile As Long
    Dim lngLineCount As Long
    Dim strHoldData As String

    On Error GoTo PROC_ERR
    
    GetLineOfData = ""
    lngLineCount = 1
    
    ' Open the source file
    lngSourceFile = FreeFile
    Open fileName For Input Access Read As lngSourceFile
    Do Until lngLineCount > LineNumber Or EOF(lngSourceFile)
        If lngLineCount = LineNumber Then
            Line Input #lngSourceFile, GetLineOfData
        Else
            Line Input #lngSourceFile, strHoldData
        End If
        lngLineCount = lngLineCount + 1
    Loop
    
    ' Close file
    Close lngSourceFile

    If GetLineOfData = "" Then
        ' Open the source file
        lngSourceFile = FreeFile
        Open fileName For Input Access Read As lngSourceFile
        Line Input #lngSourceFile, GetLineOfData
        Close lngSourceFile
    End If

Proc_EXIT:
    Exit Function
  
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.description, , _
        "GetFileFirstList"
    Resume Proc_EXIT

End Function

Private Function GetNonBillableReasonsHelper() As Recordset
    Dim objData As nADOData.CADOData
    Set objData = New nADOData.CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly
        .ResetParameters
        .OpenRecordSetFromSP "get_Non_Billable_Reasons_Tree"
        Set GetNonBillableReasonsHelper = .Recordset
    End With

    Set objData = Nothing
End Function

Public Sub GetNonBillableReasons(ctrl As ComboBox)
    Static nonBillableReasonList As Recordset
    On Error GoTo Handle_Error
    Dim reason As String
    Dim Id As Long
    
    ' because VB6 doesn't support short-circuit evaluations...
    If nonBillableReasonList Is Nothing Then
        Set nonBillableReasonList = GetNonBillableReasonsHelper
    ElseIf nonBillableReasonList.state = 0 Then
        Set nonBillableReasonList = GetNonBillableReasonsHelper
    End If
    
    nonBillableReasonList.MoveFirst
    While Not nonBillableReasonList.EOF
        reason = nonBillableReasonList!Reason_Text
        Id = nonBillableReasonList!Reason_Id
        If reason <> "" Then
            ctrl.AddItem reason
            ctrl.itemData(ctrl.NewIndex) = Id
        End If
        nonBillableReasonList.MoveNext
    Wend

Cleanup_Exit:
    Exit Sub
    
Handle_Error:
    Err.Raise Err.Number, Err.Source, Err.description
    Resume Cleanup_Exit
End Sub

Public Function GetEmployeeName(employeeId As Long)
    Dim objData As nADOData.CADOData
    On Error GoTo Error_this_Sub
            
    Set objData = New CADOData
    With objData
        Set .Connection = GetDBConnection
        .CursorType = adOpenForwardOnly
        .CommandType = adCmdStoredProc
        .LockType = adLockReadOnly

        .ResetParameters
    
        .AddParameter "Employee Id", employeeId, adInteger, adParamInput
        .OpenRecordSetFromSP "get_Employee_By_Id"
            
        If Not .Recordset.EOF Then
            GetEmployeeName = Trim$(.Recordset!Last_Name) & ", " & Trim$(.Recordset!First_Name)
            '
            .Recordset.Close
        End If
    End With

Exit_this_Sub:
    Exit Function
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Occurred Finding Employee " & employeeId & "."
    Resume Exit_this_Sub

End Function

Public Function CreateNewSPCall(Optional t As CommandTypeEnum = adCmdStoredProc)
    Dim objData As nADOData.CADOData
    Set objData = New nADOData.CADOData
    With objData
        Set .Connection = GetDBConnection
        .ResetParameters
        .CursorType = adOpenForwardOnly
        .CommandType = t
        .LockType = adLockReadOnly
    End With
    
    Set CreateNewSPCall = objData
End Function

Public Sub ResetSPCall(objData As nADOData.CADOData)
    With objData
        .ResetParameters
    End With
End Sub

Public Function ConvertDateWithTimeZone(d As Date, location As Long) As String
    Dim dateText As String
    dateText = Format$(d, "MM/dd/yyyy hh:mm AM/PM")
    ConvertDateWithTimeZone = dateText & " " & gClintrakLocations(CStr(location)).Time_Zone_Display
End Function

Public Function ConvertDate(d As Date, location As Long) As Date
    If location = 1 Then
        ConvertDate = d
    Else
        Dim objData As nADOData.CADOData
        Dim rs As Recordset
        Set objData = CreateNewSPCall(adCmdText)
        
        objData.SQL = "SELECT dbo.Convert_From_Est('" & Format$(d, "yyyy-MM-dd HH:mm:ss") & "', '" & location & "')"
        
        Set rs = objData.OpenRecordSet
        
        If Not rs.EOF Then
            Dim it As Object
            ConvertDate = rs(0)
        End If
    End If
End Function

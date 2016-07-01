VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProdRunMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<comment>
' <summary>
' This object contains procedures to load the main form, and generate, print, or export to PDF a Computerization Order Form.</summary>
'</comment>

Option Explicit

Private rpt As ARProdPlan
Dim Counter As Long  'page count in multiple IRQ report

Public Function Initialize(Username As String, Token As String, Server As String, Database As String, Optional DomainIconPath As String) As Boolean
    ' DW 2008-017 added
    Dim objLocation As ClintrakCommon.location
    
    Set gApplicationUser = New ClintrakCommon.ApplicationUser
        
    If gApplicationUser.AuthenticateFromToken(Username, Token, Server, Database, App.Title) Then
        Set gadoConnection = New nADOConn.CADOConn
        With gadoConnection
            .ConnectionString = gApplicationUser.OLEDBConnectionString
            .Mode = adModeRead
            .CursorLocation = adUseClient
            .CommandTimeout = 300
            .OpenConnection
        End With
        
        Initialize = True
        
        ' DW 2008-017 added
        ' Retrieve collection of Clintrak locations
        Set objLocation = New ClintrakCommon.location
        Set gClintrakLocations = objLocation.GetClintrakLocations(gadoConnection.Connection)
        gDomainIconPath = DomainIconPath
    End If
    
    ' DW 2008-017 added
    Set objLocation = Nothing
    
End Function

Public Sub Load_Prod_Run(FileLinksId As Long, _
                            CodingNumber As Long, _
                            JobId As Long)
                          
    On Error GoTo Error_this_Sub
    
    gFileLinksId = FileLinksId
    gCodingNumber = CodingNumber
    gJob_Id = JobId
       
    Load frmProdPlan
    frmProdPlan.Show vbModal
    
Exit_this_Sub:
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error_this_Sub:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Loading the Production Run"
    Resume Exit_this_Sub
    
End Sub

'------------------------------------------------
' Saves the existing report as an RDF or PDF file
'------------------------------------------------
Public Function SaveToFile(Path As String) As Boolean
    Dim Exportfile As ActiveReportsPDFExport.ARExportPDF
    Dim Report As ActiveReport

    Set Report = rpt

    On Error GoTo Handle_Error
    
    Set Exportfile = New ActiveReportsPDFExport.ARExportPDF
    Exportfile.fileName = Path
    Exportfile.Export Report.Pages
    Report.Pages.RemoveAll
    Report.Pages.Commit
    

exit_function:
    Exit Function

Handle_Error:
    'Log errors and pass back to calling procedure
    Err.Raise Err.Number, Err.Source & "->SaveToFile()", Err.description
    SaveToFile = False
    GoTo exit_function
    
End Function


Public Sub PrintAll_ProdPlanning_Forms(barCode As String, booPDF As Boolean)

    On Error GoTo Handle_Error

    Dim rptProdPlan As New ARProdPlan

    Dim i As Long
    Dim n As Long


    Set ProductionRun = New CProdrun
    'Set ProductionGroup = New CProdGroup    ' DW 2012-001 seems to have been mia
    
    Set Planning = New CPlanningMethods
    Set PlanningList = New CColPlanningInfo

    If Mid$(UCase$(barCode), 1, 3) = "PDR" Then
        If Planning.CheckIfReplacement(barCode) = True Then
            booReplacement = True
            gReprintFile_Type = "REPLACEMENT"
        Else
            booReplacement = False
        End If
    Else
        If ProductionGroup.CheckIfPRGOfReplacements(barCode) = True Then
            booReplacement = True
            gReprintFile_Type = "REPLACEMENT"
        Else
            booReplacement = False
        End If
    End If

    'change depending on barcode
    If Mid$(UCase$(barCode), 1, 3) = "PRG" Then
        gShowQuarantineTagFlag = False
    Else   'must be PDR
        If Planning.CheckIfCombined(barCode) = True Then
            gShowQuarantineTagFlag = False
        Else
          gShowQuarantineTagFlag = True
        End If
    End If

    If booReplacement = False Then
        Call Planning.Gather_PDR_Requirements(barCode)

        'process the PDR's gathered from the PRG to print each Processing form
        If PlanningList.count > 0 Then
            For i = 1 To PlanningList.count
                 Set rptProdPlan = New ARProdPlan
                 gFileLinksId = PlanningList.Item(i).File_Links_ID
                 gCodingNumber = PlanningList.Item(i).Coding_Number
                 gJob_Id = PlanningList.Item(i).Job_Log_Id
                 'loads the job info for the report
                 Call getFileLinksInfo
                 Call GetJobInformation
                 Call GetClientName
                 Call GetRandIDNumber
                Call GetLabelCurrentValues

                 ProductionRun.Production_Run_Id = gProductionRun_Id
                 ProductionRun.Proof_Id = gProofId

                 'gets the production run info
                 ProductionRun.LookupRecord
                 ProductionRun.GetLabelDesc

                'Clintrak Samples
                 Call ProductionRun.Determine_Clintrak_Samples
                 
                 If booPDF = True Then
    
                    rptProdPlan.Printer.Orientation = ddOLandscape
                    rptProdPlan.Run                                 'Activates the Active Reports
                    For n = 0 To rptProdPlan.Pages.count - 1
                        rpt.Pages.Insert Counter, rptProdPlan.Pages(n)
                        Counter = Counter + 1
                    Next n
                Else
                    rptProdPlan.Printer.Orientation = ddOLandscape
                    Call rptProdPlan.PrintReport(False)
                    Set rptProdPlan = Nothing
                End If
                
            Next
        Else
             MsgBox "Error Processing this Production Record! Contact IT!", vbCritical, _
                    "Error Printing Planning Reports"
        End If
    ElseIf Mid$(UCase$(barCode), 1, 3) = "PDR" Then
        Call Planning.GatherReplacementPDRRequirements(gOriginalPDRBarcode)
        If PlanningList.count > 0 Then
            For i = 1 To PlanningList.count
                Set rptProdPlan = New ARProdPlan
                gFileLinksId = PlanningList.Item(i).File_Links_ID
                gCodingNumber = PlanningList.Item(i).Coding_Number
                gJob_Id = PlanningList.Item(i).Job_Log_Id
                'loads the job info for the report
                Call getReplacementFileLinksInfo
                Call GetJobInformation
                Call GetClientName
                Call GetRandIDNumber

                ProductionRun.Production_Run_Id = gProductionRun_Id
                ProductionRun.Proof_Id = gProofId

                'gets the production run info
                ProductionRun.LookupRecord
                ProductionRun.GetLabelDesc

                'Clintrak Samples
                Call ProductionRun.Determine_Clintrak_Samples
                
                If booPDF = True Then
                    rptProdPlan.Printer.Orientation = ddOLandscape
                    rptProdPlan.Run                                 'Activates the Active Reports
                    For n = 0 To rptProdPlan.Pages.count - 1
                        rpt.Pages.Insert Counter, rptProdPlan.Pages(n)
                        Counter = Counter + 1
                    Next n
                Else
                    rptProdPlan.Printer.Orientation = ddOLandscape
                    Call rptProdPlan.PrintReport(False)
                    Set rptProdPlan = Nothing
                End If
                
            Next i
        Else
             MsgBox "Error Processing this Production Record! Contact IT!", vbCritical, _
                    "Error Printing Planning Reports"
        End If

    ElseIf Mid$(UCase$(barCode), 1, 3) = "PRG" Then
        Call Planning.GatherPRGReplacementPDRRequirements(barCode)
        If PlanningList.count > 0 Then
            For i = 1 To PlanningList.count
                If Planning.CheckIfReplacement(PlanningList.Item(i).Form_Idectification) = True Then
                    booReplacement = True
                    gReprintFile_Type = "REPLACEMENT"
                Else
                    booReplacement = False
                End If
                Set rptProdPlan = New ARProdPlan
                gFileLinksId = PlanningList.Item(i).File_Links_ID
                gCodingNumber = PlanningList.Item(i).Coding_Number
                gJob_Id = PlanningList.Item(i).Job_Log_Id
                'loads the job info for the report
                Call getReplacementFileLinksInfo
                Call GetJobInformation
                Call GetClientName
                Call GetRandIDNumber

                ProductionRun.Production_Run_Id = gProductionRun_Id
                ProductionRun.Proof_Id = gProofId

                'gets the production run info
                ProductionRun.LookupRecord
                ProductionRun.GetLabelDesc

                'Clintrak Samples
                Call ProductionRun.Determine_Clintrak_Samples
                
                If booPDF = True Then
                    rptProdPlan.Printer.Orientation = ddOLandscape
                    rptProdPlan.Run                                 'Activates the Active Reports
                    For n = 0 To rptProdPlan.Pages.count - 1
                        rpt.Pages.Insert Counter, rptProdPlan.Pages(n)
                        Counter = Counter + 1
                    Next n
                Else
                    rptProdPlan.Printer.Orientation = ddOLandscape
                    Call rptProdPlan.PrintReport(False)
                    Set rptProdPlan = Nothing
                End If
                
            Next i
        Else
             MsgBox "Error Processing this Production Record! Contact IT!", vbCritical, _
                    "Error Printing Planning Reports"
        End If
    

    End If

Exit_Sub:
    Exit Sub

Handle_Error:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Printing the Production Planning Forms"
    Resume Exit_Sub

End Sub

Public Sub Print_PRGPlanning_Form(barCode As String, booPDF As Boolean)
    On Error GoTo Handle_Error

    Dim rptPRGProdPlan As New ARPRGProdPlan
    Dim n As Long

    Set ProductionRun = New CProdrun
    Set ProductionGroup = New CProdGroup
    Set Planning = New CPlanningMethods
    Set PlanningList = New CColPlanningInfo
    
    If ProductionGroup.CheckIfPRGOfReplacements(barCode) = True Then
        booReplacement = True
        gReprintFile_Type = "REPLACEMENT"
    Else
        booReplacement = False
    End If
    
    'md 2006-020
    Planning.CountPDR_Reorient_Ind (barCode)
    
    If booReplacement = False Then
        Planning.Gather_PDR_Requirements (barCode)
        If PlanningList.count > 0 Then
            Set rptPRGProdPlan = New ARPRGProdPlan
            gJob_Id = PlanningList.Item(1).Job_Log_Id
             'loads the job info for the report
            Call GetJobInformation
            Call GetClientName
             
            ProductionGroup.Form_Identification = barCode
            ProductionGroup.LookupRecord
            
            ProductionGroup.GetPRGEmployeeById
            ProductionGroup.GetPRGLabelId
            ProductionGroup.GetPRGLabelDesc
            ProductionGroup.GetPRGStockInfo
            ProductionGroup.GetPRGStockDesc
            
            gRandomizationId = ProductionGroup.Randomization_Id
            Call GetRandIDNumber
        
            If booPDF = True Then
                rptPRGProdPlan.Printer.Orientation = ddOLandscape
                rptPRGProdPlan.Run                                 'Activates the Active Reports
                For n = 0 To rptPRGProdPlan.Pages.count - 1
                    rpt.Pages.Insert Counter, rptPRGProdPlan.Pages(n)
                    Counter = Counter + 1
                Next n
            Else
                rptPRGProdPlan.Printer.Orientation = ddOLandscape
                Call rptPRGProdPlan.PrintReport(False)
                Set rptPRGProdPlan = Nothing
            End If
            
        Else
            MsgBox "Error Processing this Production Record! Contact IT!", vbCritical, _
                    "Error Printing Planning Reports"
        End If

        
    Else
        Planning.GatherPRGReplacementPDRRequirements (barCode)
        If PlanningList.count > 0 Then
            Set rptPRGProdPlan = New ARPRGProdPlan
            gJob_Id = PlanningList.Item(1).Job_Log_Id
             'loads the job info for the report
            Call GetJobInformation
            Call GetClientName
             
            ProductionGroup.Form_Identification = barCode
            ProductionGroup.LookupRecord
            
            ProductionGroup.GetPRGEmployeeById
            ProductionGroup.GetPRGLabelId
            ProductionGroup.GetPRGLabelDesc
            ProductionGroup.GetPRGStockInfo
            ProductionGroup.GetPRGStockDesc
            
            gRandomizationId = ProductionGroup.Randomization_Id
            Call GetRandIDNumber

            If booPDF = True Then
                rptPRGProdPlan.Printer.Orientation = ddOLandscape
                rptPRGProdPlan.Run                                 'Activates the Active Reports
                For n = 0 To rptPRGProdPlan.Pages.count - 1
                    rpt.Pages.Insert Counter, rptPRGProdPlan.Pages(n)
                    Counter = Counter + 1
                Next n
            Else
                rptPRGProdPlan.Printer.Orientation = ddOLandscape
                Call rptPRGProdPlan.PrintReport(False)
                Set rptPRGProdPlan = Nothing
            End If
            
        Else
            MsgBox "Error Processing this Production Record! Contact IT!", vbCritical, _
                    "Error Printing Planning Reports"
        End If
    

    End If
    
Exit_Sub:
    Exit Sub

Handle_Error:
    MsgBox "Error: " & Err.Number & " " & Err.description, vbCritical, _
    "Error Printing the Production Planning Forms"
    Resume Exit_Sub

End Sub

Public Sub CreateMultiReport()
    Set rpt = New ARProdPlan
End Sub


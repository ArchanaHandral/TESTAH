VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArShippingRpt2 
   Caption         =   "Shipping"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17415
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   30718
   _ExtentY        =   11774
   SectionData     =   "ArShippingRpt2.dsx":0000
End
Attribute VB_Name = "ArShippingRpt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private madoData As nADOData.CADOData

Private Sub Detail_Format()

    'Handle PRG runs here
    If gIsPRGRun Then

        If ProductionGroup.PrgShippingInd <> 1 Then
            Me.otxtAttn.Visible = False
            Me.lblAttn.Visible = False
            Me.otxtShip.text = "See Individual PDRs"
        Else
            Call LoadShippingInfo(ProductionGroup.ShipSeedVal)
        End If

    Else

        Call LoadShippingInfo(-1)

    End If

End Sub


Private Sub LoadShippingInfo(ByVal jobShippingId As Long)

    If madoData Is Nothing Then
        Set madoData = New nADOData.CADOData
    End If

    With madoData
        Set .Connection = GetDBConnection
        .ResetParameters
        .CommandType = adCmdStoredProc
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly

        .AddParameter "Job Log Id", gJob_Id, adInteger, adParamInput
        If jobShippingId = -1 Then
            .AddParameter "Job Shipping Id", ProductionRun.Ship_To_Id, adInteger, adParamInput
        Else
            .AddParameter "Job Shipping Id", jobShippingId, adInteger, adParamInput
        End If

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



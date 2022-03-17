VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmStaffData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entry Form"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAll 
      Caption         =   "All Task Report"
      Height          =   390
      Left            =   12480
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.ListBox lstTask 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAchive 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtTask 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin MSChart20Lib.MSChart ch 
      Height          =   5175
      Left            =   9480
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   5175
   End
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   2775
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAmt 
      BeginProperty Font 
         Name            =   "Preeti"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   9
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Task                                  Achive"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "frmStaffData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
If Val(txtAchive) > 0 Then
    Dim rsData As Recordset
    Set rsData = New Recordset
    Refress_Rs rsData, "Select * from ViewTaskAchive where yearr=" & Year(myDate) & " and monthh = " & Month(myDate) & " and Touser = " & CurrenUser & " and taskid = " & Val(txtTask.tag)
    If rsData.RecordCount > 0 Then
        ExecuteQuery "Insert into tblTaskAchive values( " & NewMaxID("tblTaskAchive", "SN") & ", " & rsData!SN & ", " & Val(txtAchive.Text) & ", '" & Format(myDate, "yyyy/mm/dd") & "',0)"
        Message "Data Posting succesfully."
        GetEntryAchive
        txtAchive = ""
        txtTask = ""
        txtTask.tag = ""
        
        If lstTask.ListCount > lstTask.ListIndex + 1 Then
            lstTask.Selected(lstTask.ListIndex + 1) = True
        Else
            lstTask.Selected(0) = True
        End If
        
        txtTask.SetFocus
    Else
        Message "Problem Occered while recording Data."
        Exit Sub
    End If
Else
    Message "Please Enter your Achived."
    txtAchive.SetFocus
    Exit Sub
End If
End Sub



Private Sub cmdAll_Click()
frmAchiveList.userID = CurrenUser
frmAchiveList.BanchID = CurrenBranchID
frmAchiveList.Show vbModal
End Sub


Private Sub dg1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 68 And Shift = 2) Or KeyCode = 46 Then

    If Not dg1.Columns(3).Value = 0 Then
        Message "You Can not Update Approved Record."
        GetEntryAchive
        Exit Sub
    End If
    
    Message "Are you sure to delete this record ?", YesNo
    If CurrentMsgResponce = Yes Then
        ExecuteQuery "Delete from tblTaskAchive where sn = " & Val(dg1.Columns(4).Value)
        GetEntryAchive
        Message "Record Deleted succesfully."
    End If
End If
End Sub
'
'Private Sub dg1_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''Dim X As Single
''Dim strr As String
''X = dg1.Columns(4).Value
''If Not dg1.Columns(3).Value = 0 Then
''    Message "You Can not Update Approved Record."
''    GetEntryAchive
''    Exit Sub
''End If
''Message "Data Update Succesfully."
''strr = "Update tblTaskAchive set Achive = " & Val(dg1.Columns(1)) & " where sn = " & Val(dg1.Columns(4).Value) & " and Approved=0"
''ExecuteQuery strr
''GetEntryAchive
''End If
'End Sub

Private Sub Form_Load()
getAsignTask
getList
End Sub

Private Sub GetEntryAchive()
Dim rsList As Recordset
Set rsList = New Recordset
Refress_Rs rsList, "Select DateD, Achive, Status = case Approved when 0 then 'Not-Approved' else 'Approved' end, Approved, AchiveSN from viewTaskAchiveDets where yearr=" & Year(myDate) & " and monthh=" & Month(myDate) & " and brich = " & CurrenUser & " and TaskID = " & Val(txtTask.tag)
Set dg1.DataSource = rsList
dg1.RowHeight = 325
dg1.Columns(0).Width = 1500
dg1.Columns(1).Width = 1900
dg1.Columns(1).Alignment = dbgRight
dg1.Columns(2).Width = 1900
dg1.Columns(2).Alignment = dbgCenter
dg1.Columns(3).Visible = False
dg1.Columns(4).Visible = False
'scrolll------------------------------------------------------------------------------------------------------

dg1.Scroll 0, rsList.RecordCount

End Sub

Private Sub getList()
Dim rsList As Recordset
Set rsList = New Recordset
Refress_Rs rsList, "Select TaskName, Target, Achive, Achive-target Diff, round(Achive*100/(Target+1),2) Percentage from ViewTaskAchive where yearr= " & Year(myDate) & " and monthh=" & Month(myDate) & " and brich =  " & CurrenUser
Set dg.DataSource = rsList
dg.RowHeight = 325
dg.Columns(0).Width = 1500
dg.Columns(1).Width = 1900
dg.Columns(1).Alignment = dbgRight
dg.Columns(2).Width = 1900
dg.Columns(2).Alignment = dbgRight
dg.Columns(3).Width = 1800
dg.Columns(3).Alignment = dbgRight
dg.Columns(4).Width = 1000
dg.Columns(4).Caption = "%"
dg.Columns(4).Alignment = dbgRight

dg.Scroll 0, Val(lstTask.ListIndex)

End Sub

Private Sub updateChart()
ch.Visible = True
Dim rsChart As Recordset
Set rsChart = New Recordset
Refress_Rs rsChart, "Select * from ViewTaskAchive where yearr=" & Year(myDate) & " and monthh = " & Month(myDate) & " and brich = " & CurrenUser & " and taskID = " & Val(txtTask.tag)
If rsChart.RecordCount > 0 Then
    If rsChart!achive >= rsChart!target Then
        ch.ColumnCount = 2
        ch.Column = 1
        ch.data = rsChart!achive
        ch.Column = 2
        ch.data = 0
        
    Else
        ch.ColumnCount = 2
        ch.RowCount = 1
        ch.Column = 1
        ch.data = rsChart!achive
        ch.Column = 2
        ch.data = rsChart!target - rsChart!achive
        
    End If
    ch.RowLabel = rsChart!TaskName
End If
End Sub

Private Sub txtAchive_Change()
If Val(txtAchive) > 0 Then
    lblAmt = NString(Val(txtAchive))
Else
    lblAmt = ""
End If
End Sub

Private Sub txtAchive_GotFocus()
Colored
End Sub

Private Sub txtAchive_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd.SetFocus
End If
End Sub

Private Sub txtAchive_LostFocus()
unColored txtAchive
End Sub

Private Sub txtTask_Change()
If lstTask.Visible = True Then
    If Not txtTask.Text = "" Then
        For I = 0 To lstTask.ListCount - 1
            If Val(txtTask) = Val(lstTask.list(I)) Then
                lstTask.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtTask.Text)) = UCase(Trim(Mid(lstTask.list(I), InStr(1, lstTask.list(I), ":") + 2, Len(txtTask.Text)))) Or Val(txtTask.Text) = Val(Mid(lstTask.list(I), 1, InStr(1, lstTask.list(I), ":") - 2)) Then
                    lstTask.Selected(I) = True
                    Exit For
                Else
                    lstTask.Selected(I) = False
                End If
            End If
        Next
    Else
        lstTask.Selected(0) = False
    End If
End If



End Sub

Private Sub txtTask_GotFocus()
Colored
lstTask.Visible = True
End Sub

Private Sub txtTask_KeyDown(KeyCode As Integer, Shift As Integer)
If lstTask.Visible = True Then
    If KeyCode = 38 Then
        If lstTask.ListIndex <= 0 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstTask.Selected(lstTask.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstTask.ListIndex = lstTask.ListCount - 1 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstTask.Selected(lstTask.ListIndex + 1) = True
        End If
    End If
End If


End Sub

Private Sub txtTask_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstTask.Visible = True Then
        If lstTask.ListIndex >= 0 Then
            If lstTask.ItemData(lstTask.ListIndex) > 0 Then
                    txtTask.tag = lstTask.ItemData(lstTask.ListIndex)
                    txtTask.Text = Trim(Mid(lstTask.Text, InStr(1, lstTask.Text, ":") + 1))
                    updateChart
                    GetEntryAchive
                    getList
                    txtAchive.SetFocus
            End If
        End If
    End If
End If


End Sub


Private Sub txtTask_LostFocus()
unColored txtTask
lstTask.Visible = False
End Sub

Private Sub getAsignTask()
Dim rsTask As Recordset
Set rsTask = New Recordset
Refress_Rs rsTask, "Select * from ViewGivenTask where brich = " & CurrenUser & " and yearr = " & Year(myDate) & " and monthh = " & Month(myDate)
lstTask.Clear
If rsTask.RecordCount > 0 Then
    rsTask.MoveFirst
    Do While Not rsTask.EOF
        lstTask.AddItem Format(rsTask!proitity, "00") & " : " & rsTask!TaskName
        lstTask.ItemData(rsTask.AbsolutePosition - 1) = rsTask!taskID
        rsTask.MoveNext
    Loop
    lstTask.ListIndex = 0
End If

End Sub


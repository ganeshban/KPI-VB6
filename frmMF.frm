VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MF Form"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      Caption         =   "Rpt"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox lstMF 
      Appearance      =   0  'Flat
      Height          =   3390
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtMFName 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   3390
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtTask 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtdata 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3240
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   6135
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
   Begin VB.CommandButton cmdNewMF 
      Caption         =   "New"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblTotAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Collection Amount"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   7920
      Width           =   6495
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsData As Recordset


Private Sub cmdNewMF_Click()
frmMFMaster.Show vbModal
End Sub

Private Sub cmdNewMF_GotFocus()
VisiableList cmdNewMF.Name
End Sub

Private Sub cmdReport_Click()
frmMFReport.Show vbModal
End Sub

Private Sub cmdReport_GotFocus()
VisiableList cmdNewMF.Name
End Sub

Private Sub cmdSave_Click()

If Not Val(txtMFName.tag) > 0 Then
    Message "Choose MF Group First"
    txtMFName.SetFocus
    Exit Sub
End If

If Not Val(txtTask.tag) > 0 Then
    Message "Please Choose Task."
    txtTask.SetFocus
    Exit Sub
End If



If Not Val(txtdata) > 0 Then
    Message "Please Enter Data"
    txtdata.SetFocus
    Exit Sub
End If

Refress_Rs rsData, "Select * from ViewMFData where MFGRP = " & Val(txtMFName.tag) & " and substring(DateD,1,8) = '" & Mid(myDate, 1, 8) & "' and TaskID = " & Val(txtTask.tag)
If rsData.RecordCount > 0 Then
    Message "Data is already posted for this month of this task."
    Exit Sub
End If

SaveData

If lst.ListIndex = lst.ListCount - 1 Then
    
    Message "Do you want to post those record ?", YesNo, True
    If CurrentMsgResponce = Yes Then
        ExecuteQuery "update tblmfdata set Posted = 1 where userno = " & CurrenUser & " and mfgrp = " & Val(txtMFName.tag) & " and substring(dated,1,8) = '" & Mid(myDate, 1, 8) & "'"
        Message "Records are marked as Posted."
        txtMFName.SetFocus
    Else
        Exit Sub
    End If
    
    
    txtdata = ""
    txtTask = ""
    lst.ListIndex = 0
    getMFName
    lstMF.ListIndex = 0
    txtMFName.SetFocus
Else
    lst.Selected(lst.ListIndex + 1) = True
    txtdata = ""
    txtTask = ""

    txtTask.SetFocus
End If

End Sub

Private Sub VisiableList(Optional data As String = "")
On Error Resume Next
Select Case data
    Case txtMFName.Name
        lst.Visible = False
        lstMF.Visible = True
    Case txtTask.Name
        lst.Visible = True
        lstMF.Visible = False
    Case Else
        lst.Visible = False
        lstMF.Visible = False
End Select
End Sub

Private Sub cmdSave_GotFocus()
VisiableList cmdNewMF.Name
End Sub


Private Sub dg_DblClick()
Refress_Rs rsData, "Select * from ViewMFdata where Posted = 0 and MFGRP = " & Val(txtMFName.tag) & " and substring(DateD,1,8) = '" & Mid(myDate, 1, 8) & "'"
If rsData.RecordCount > 0 Then
    Message "There are " & rsData.RecordCount & " records were not posted. Wanna Post all at once.", YesNo
    If CurrentMsgResponce = Yes Then
        ExecuteQuery "Update tblMFdata set Posted = 1 where Posted = 0 and MFGRP = " & Val(txtMFName.tag) & " and substring(DateD,1,8) = '" & Mid(myDate, 1, 8) & "'"
        Message "All records are marked as posted !!!"
        getMFName
        txtMFName.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub dg_GotFocus()
VisiableList cmdNewMF.Name
End Sub

Private Sub dg_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 68 Then
    If dg.row >= 0 Then
        Message "you are about to delete selected record. wanna sure ?", YesNo
        If CurrentMsgResponce = Yes Then
            ExecuteQuery "Delete from tblmfdata where sn = " & dg.Columns(0)
            Message "Record Deleted succes !!"
            getList
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Set rsData = New Recordset
getMFName 1
getTaskName
lblDate = myDate
End Sub

Private Sub getList()
Dim rsList As Recordset
Set rsList = New Recordset

Refress_Rs rsList, "Select sum(Data) data from ViewMFData where Countable = 1 and MFGRP = " & Val(txtMFName.tag) & " and substring(DateD,1,8) = '" & Mid(myDate, 1, 8) & "'"
If rsList.RecordCount > 0 Then
    lblTotAmt = "Total Collection Amount Rs. " & rsList.Fields(0)
End If

Refress_Rs rsList, "Select SN, taskID, TaskName, Countable, Data from ViewMFData where MFGRP = " & Val(txtMFName.tag) & " and substring(DateD,1,8) = '" & Mid(myDate, 1, 8) & "'"
Set dg.DataSource = rsList
dg.RowHeight = 300
dg.Columns(0).Visible = False
dg.Columns(1).Width = 1000
dg.Columns(3).Visible = False
'dg.Columns(4).Visible = False
End Sub

Private Sub getMFName(Optional tag As Integer = 0)
Dim strr As String
If tag = 0 Then
    strr = "Select * from tblmfGroup where branch = " & CurrenBranchID & " and sn not in(Select distinct MFGrp from tblMFData where Posted=1 and substring(dated,1,8) = '" & Mid(myDate, 1, 8) & "') order by DayCode "
Else
    strr = "Select * from tblmfGroup where branch = " & CurrenBranchID
End If

Refress_Rs rsData, strr

lstMF.Clear
If rsData.RecordCount > 0 Then
    Do While Not rsData.EOF
        lstMF.AddItem Format(rsData.AbsolutePosition, "0000") & " : " & rsData!groupname
        lstMF.ItemData(rsData.AbsolutePosition - 1) = rsData!SN
        rsData.MoveNext
    Loop
End If

'rsData.Filter = " sn not in(Select distinct MFGrp from tblMFData where Posted=1 and substring(dated,1,8) = '" & Mid(myDate, 1, 8) & "') "
'If rsData.RecordCount > 0 Then
'    lstMF.Clear
'    Do While Not rsData.EOF
'        lstMF.AddItem rsData!sn & " : " & rsData!groupname
'        lstMF.ItemData(rsData.AbsolutePosition - 1) = rsData!sn
'        rsData.MoveNext
'    Loop
'End If
End Sub

Public Sub getTaskName()
Refress_Rs rsData, "Select * from tblMFTask "
lst.Clear
If rsData.RecordCount > 0 Then
    Do While Not rsData.EOF
        lst.AddItem Format(rsData!SN, "00") & " : " & rsData!TaskName
        lst.ItemData(rsData.AbsolutePosition - 1) = rsData!SN
        rsData.MoveNext
    Loop
    lst.Selected(0) = True
    lst.AddItem "00 : Create New Task"
    lst.ItemData(lst.ListCount - 1) = -2
End If
End Sub



Private Sub SaveData()
    ExecuteQuery "Insert into tblmfdata values ( " & NewMaxID("tblmfdata", "sn") & ", " & CurrenUser & ", " & Val(txtdata.Text) & ", " & Val(txtTask.tag) & ", " & Val(txtMFName.tag) & ", '" & myDate & "', 0)"
    getList
End Sub

Private Sub lst_DblClick()
    txtTask_KeyPress 13
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTask_KeyPress 13
End If
End Sub

Private Sub lstMF_DblClick()
txtMFName_KeyPress 13
End Sub

Private Sub lstMF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMFName_KeyPress 13
End If
End Sub

Private Sub txtdata_GotFocus()
VisiableList cmdNewMF.Name
Colored
End Sub

Private Sub txtdata_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub txtdata_LostFocus()
unColored txtdata
End Sub

Private Sub txtMFName_Change()
If lstMF.Visible = True Then
    If lstMF.ListCount > 0 Then
        If Not txtMFName.Text = "" Then
            For I = 0 To lstMF.ListCount - 1
                If Val(txtMFName) = Val(lstMF.List(I)) Then
                    lstMF.Selected(I) = True
                    Exit For
                Else
                    If UCase(Trim(txtMFName.Text)) = UCase(Trim(Mid(lstMF.List(I), InStr(1, lstMF.List(I), ":") + 2, Len(txtMFName.Text)))) Or Val(txtMFName.Text) = Val(Mid(lstMF.List(I), 1, InStr(1, lstMF.List(I), ":") - 2)) Then
                        lstMF.Selected(I) = True
                        Exit For
                    Else
                        lstMF.Selected(I) = False
                    End If
                End If
            Next
        Else
            lstMF.Selected(0) = False
        End If
    End If
End If

End Sub

Private Sub txtMFName_GotFocus()
Colored
VisiableList txtMFName.Name
End Sub

Private Sub txtMFName_KeyDown(KeyCode As Integer, Shift As Integer)
If lstMF.Visible = True Then
    If KeyCode = 38 Then
        If lstMF.ListIndex <= 0 Then
    '        lstMFData.Selected(lstMFData.ListCount - 1) = True
            Exit Sub
        Else
            lstMF.Selected(lstMF.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstMF.ListIndex = lstMF.ListCount - 1 Then
    '        lstMFData.Selected(lstMFData.ListCount - 1) = True
            Exit Sub
        Else
            lstMF.Selected(lstMF.ListIndex + 1) = True
        End If
    End If
End If

End Sub

Private Sub txtMFName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMFName = Trim(Mid(lstMF.List(lstMF.ListIndex), InStr(1, lstMF.List(lstMF.ListIndex), ":") + 1))
    txtMFName.tag = Val(lstMF.ItemData(lstMF.ListIndex))
    getList
    txtTask.SetFocus
End If
End Sub

Private Sub txtMFName_LostFocus()
unColored txtMFName
End Sub

Private Sub txtTask_Change()
If lst.Visible = True Then
    If Not txtTask.Text = "" Then
        For I = 0 To lst.ListCount - 1
            If Val(txtTask) = Val(lst.List(I)) Then
                lst.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtTask.Text)) = UCase(Trim(Mid(lst.List(I), InStr(1, lst.List(I), ":") + 2, Len(txtTask.Text)))) Or Val(txtTask.Text) = Val(Mid(lst.List(I), 1, InStr(1, lst.List(I), ":") - 2)) Then
                    lst.Selected(I) = True
                    Exit For
                Else
                    lst.Selected(I) = False
                End If
            End If
        Next
    Else
        lst.Selected(0) = False
    End If
End If

End Sub

Private Sub txtTask_GotFocus()
Colored
VisiableList txtTask.Name
End Sub

Private Sub txtTask_KeyDown(KeyCode As Integer, Shift As Integer)
If lst.Visible = True Then
    If KeyCode = 38 Then
        If lst.ListIndex <= 0 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lst.Selected(lst.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lst.ListIndex = lst.ListCount - 1 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lst.Selected(lst.ListIndex + 1) = True
        End If
    End If
End If


End Sub

Private Sub txtTask_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lst.Visible = True Then
        If lst.ListIndex >= 0 Then
            If lst.ItemData(lst.ListIndex) > 0 Then
                txtTask.tag = lst.ItemData(lst.ListIndex)
                txtTask.Text = Trim(Mid(lst.Text, InStr(1, lst.Text, ":") + 1))
                txtdata.SetFocus
            Else
                Message "Do You want to Create New Task ?", YesNo, True
                If CurrentMsgResponce = Yes Then
                    frmTaskMF.Show vbModal
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub txtTask_LostFocus()
unColored txtTask
End Sub

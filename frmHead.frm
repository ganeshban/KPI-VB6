VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHead 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdAddSC 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   8520
      Width           =   855
   End
   Begin VB.ListBox lstStaff 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   8520
      TabIndex        =   18
      Top             =   -3120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox lstMonth 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   8400
      TabIndex        =   12
      Top             =   -2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstTask 
      Appearance      =   0  'Flat
      Height          =   3000
      Left            =   8760
      TabIndex        =   8
      Top             =   -2160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ListBox lstSCname 
      Appearance      =   0  'Flat
      Height          =   3270
      Left            =   8640
      TabIndex        =   3
      Top             =   -2280
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox txtStaff 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtscname 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6000
         TabIndex        =   31
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   28
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblContact 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact :"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   26
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "  Post :"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   25
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblphone 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lblScCode 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6720
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Center :"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtMonth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtJD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtarget 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   3135
      Left            =   0
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
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
   Begin MSDataGridLib.DataGrid dg 
      Height          =   5415
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
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
   Begin VB.CommandButton cmdRpt 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblExport 
      Alignment       =   2  'Center
      Caption         =   "Excel"
      Height          =   255
      Left            =   8040
      TabIndex        =   30
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   8640
      Width           =   6615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   19
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Month :"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Task"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   6255
   End
End
Attribute VB_Name = "frmHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTarget As Recordset
Dim rsData As Recordset
Dim I As Integer
Dim StaffSelection As Boolean



Private Sub listVisible(ctrl As String)
On Error Resume Next
Select Case ctrl
    Case txtscname.Name
        lstSCname.Visible = True
        lstMonth.Visible = False
        lstStaff.Visible = False
        lstTask.Visible = False
    Case txtMonth.Name
        lstSCname.Visible = False
        lstMonth.Visible = True
        lstStaff.Visible = False
        lstTask.Visible = False
    Case txtStaff.Name
        lstSCname.Visible = False
        lstMonth.Visible = False
        lstStaff.Visible = True
        lstTask.Visible = False
    Case txtJD.Name
        lstSCname.Visible = False
        lstMonth.Visible = False
        lstStaff.Visible = False
        lstTask.Visible = True
    Case Else
        lstSCname.Visible = False
        lstMonth.Visible = False
        lstStaff.Visible = False
        lstTask.Visible = False
End Select
End Sub

Private Sub cmdAdd_Click()
Dim sstr As String
Dim factar As String
Dim rsD As Recordset

Set rsD = New Recordset

If Val(txtarget) = 0 Then
    Message "Please Enter Target Data "
    txtarget.SetFocus
    Exit Sub
End If

If Val(txtMonth.tag) = 0 Then
    Message "Month not set"
    txtMonth.SetFocus
    Exit Sub
End If

If Val(txtYear) = 0 Then
    Message "Year not define"
    txtYear.SetFocus
    Exit Sub
End If

If Val(txtscname.tag) = 0 Then
    Message "It seems to problem in chooseing Sercvicecenter."
    txtscname.SetFocus
    Exit Sub
End If

If Val(txtStaff.tag) = 0 Then
    Message "Please Select Staff Poperly"
    txtStaff.SetFocus
    Exit Sub
End If

If Val(txtJD.tag) = 0 Then
    Message "Select Task Perporly"
    txtJD.SetFocus
    Exit Sub
End If

If CurrenUser = Val(txtStaff.tag) Then
    Refress_Rs rsD, "Select * from viewTaskAchiveDets where Achive > 0 and yearr = " & Val(txtYear) & " and monthh = " & Val(txtMonth.tag) & " and taskTo = " & Val(txtscname.tag) & " and brich= " & CurrenUser & " and TaskID = " & Val(txtJD.tag)
    If rsD.RecordCount > 0 Then
        Message "You have already Posed Achive for this task."
    Else
        Dim rsRefID As Recordset
        Set rsRefID = New Recordset
        Refress_Rs rsRefID, "Select * from tblGiventask where Brich = " & Val(txtStaff.tag) & " and yearr = " & Val(txtYear) & " and monthh = " & Val(txtMonth.tag) & " and taskTo = " & Val(txtscname.tag) & " and taskID = " & Val(txtJD.tag)
        If rsRefID.RecordCount > 0 Then
            ExecuteQuery "Insert into tblTaskAchive values(" & NewMaxID("tbltaskAchive", "SN") & ", " & rsRefID!SN & ", " & Val(txtarget.Text) & ", '" & Format(myDate, "yyyy/mm/dd") & "',1 )"
            getTarget True
            Message "Achivement Posted succesfully!!!"
            txtJD = ""
            txtarget = ""
            
            If lstTask.ListCount > lstTask.ListIndex + 1 Then
                lstTask.ListIndex = lstTask.ListIndex + 1
            Else
                lstTask.ListIndex = 0
            End If
            txtJD.SetFocus
        Else
            Message "Problem Occered when task managing."
        End If
    End If
    Exit Sub
End If


Refress_Rs rsD, "Select * from ViewTaskAchive where yearr=" & txtYear & " and monthh = " & Val(txtMonth.tag) & " and brich = " & txtStaff.tag & " and taskID = " & Val(txtJD.tag)
If rsD.RecordCount > 0 Then
    Message "You have already posted target for this task. Now you can edit this from data grid."
    txtJD.SetFocus
    Exit Sub
End If



If userType = 1 Then
    
    Refress_Rs rsD, "Select * from ViewUsers where sn = " & Val(txtStaff.tag)
    
    If rsD!SN <> rsD!incharge Then
        Message "You can not assign task To Staff."
        Exit Sub
    End If
Else
    Refress_Rs rsD, "Select * from ViewUsers where sn = " & Val(txtStaff.tag)
    If rsD!SN = rsD!incharge Then
        Exit Sub
    End If

End If



sstr = "Insert into tblGivenTask values( " & NewMaxID("tblGivenTask", "SN") & ", " & Val(txtJD.tag) & ", " & CurrenUser & ", " & Val(txtscname.tag) & ", " & Val(txtarget.Text) & ", " & Val(txtYear.Text) & ", " & Val(txtMonth.tag) & ", " & Val(txtStaff.tag) & " )"
ExecuteQuery sstr
If StaffSelection Then
    getTarget False
Else
    getTarget True
End If
Message "Target assign succesfully."
txtJD = ""
txtarget = ""

If lstTask.ListCount > lstTask.ListIndex + 1 Then
    lstTask.ListIndex = lstTask.ListIndex + 1
Else
    lstTask.ListIndex = 0
End If

txtJD.SetFocus

End Sub

Private Sub cmdAdd_GotFocus()
listVisible cmdAdd.Name
End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift

End Sub

Private Sub cmdAddSC_Click()
frmServicecenter.isNewForm = myYes
frmServicecenter.Show vbModal
End Sub

Private Sub cmdAddSC_GotFocus()
listVisible cmdAddSC.Name
End Sub

Private Sub cmdAddSC_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift

End Sub

Private Sub cmdEdit_Click()
If Val(txtscname.tag) > 0 Then
    frmServicecenter.isNewForm = myno
    frmServicecenter.Show vbModal
Else
    Message "Please Choose One From List."
End If
End Sub


Private Sub cmdEdit_GotFocus()
listVisible cmdEdit.Name
End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift

End Sub

Private Sub cmdRpt_Click()
    frmAchiveList.userID = Val(txtStaff.tag)
    frmAchiveList.BanchID = Val(txtscname.tag)
    frmAchiveList.yearr = Val(txtYear)
    frmAchiveList.monthh = Val(txtMonth.tag)
    frmAchiveList.Show vbModal
End Sub

Private Sub cmdRpt_GotFocus()
listVisible cmdRpt.Name
End Sub

Private Sub dg_BeforeUpdate(Cancel As Integer)
    If Not (dg.Col = 1 Or dg.Col = 2) Then
        Cancel = 2
    End If
End Sub

Private Sub dg_Click()
On Error Resume Next
AppDets 1
End Sub

Private Sub dg_DblClick()
If Not (dg.Col = 1 Or dg.Col = 2) Then
    Exit Sub
End If

Message "Do you want to edit the value of " & dg.Columns(0) & " ?", YesNo
If CurrentMsgResponce = Yes Then
    
    a = InputBox("Enter the value of " & dg.Columns(0) & " . . . . ")
    
    If Val(a) > 0 Then
        Dim strr As String
        If userType = 1 And StaffSelection = False And (dg.Col = 1 Or dg.Col = 2) Then
            Message "Target Updated Sucessfully."
            If dg.Col = 1 Then
                strr = "Update tblGivenTask set Target = " & Val(a) & " where sn = " & dg.Columns(5)
            Else
                strr = "Update tbltaskAchive set Achive = " & Val(a) & " where gtsn = " & dg.Columns(5)
            End If
            ExecuteQuery strr
        End If
        
        If userType = 2 And StaffSelection = True And dg.Col = 1 Then
            Message "Target Updated Sucessfully."
            strr = "Update tblGivenTask set Target = " & Val(a) & " where sn = " & dg.Columns(5)
            ExecuteQuery strr
        End If
        
        If StaffSelection Then
            getTarget False
        Else
            getTarget True
        
        End If
        
        dg.Col = 0
    End If
End If
End Sub

Private Sub dg_GotFocus()
    listVisible dg.Name
End Sub

Private Sub dg_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift

End Sub


Private Sub dg1_DblClick()
On Error Resume Next
Dim fac As String

fac = dg1.Columns(4).Value
If userType = 2 Then
    If Not dg1.Columns(3) = "Approved" Then
        If Val(fac) > 0 Then
            Message "Are you sure To Approved This Record ?", YesNo, True
            If CurrentMsgResponce = Yes Then
                ExecuteQuery "Update tblTaskAchive set Approved = 1 where sn = " & Val(fac)
                Message "Record Updated Succesfully."
                AppDets
                If StaffSelection Then
                    getTarget False
                Else
                    getTarget True
                End If
            End If
        Else
            Message "No Record Selected "
            Exit Sub
        End If
    Else
'        Dim x As Integer
'        Dim captcha As String
'        Randomize
'        x = Int(Rnd * 9)
'        captcha = captcha & x
'        Randomize
'        x = Int(Rnd * 9)
'        captcha = captcha & x
'        Randomize
'        x = Int(Rnd * 9)
'        captcha = captcha & x
'        Randomize
'        x = Int(Rnd * 9)
'        captcha = captcha & x
'        Randomize
'        x = Int(Rnd * 9)
'        captcha = captcha & x
        Message "Records are already Approved !" & captcha
        Exit Sub
    End If
Else
    Message "Sorry you Can't Approve staff data. "
    Exit Sub
End If
End Sub

Private Sub dg1_GotFocus()
listVisible dg1.Name
End Sub

Private Sub dg1_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift

End Sub

Private Sub Form_Load()
Set rsTarget = New Recordset
Set rsData = New Recordset

getMonth

lblDate = Format(myDate, "yyyy-mm-dd")
lstSCname.Move txtscname.Left + Picture1.Left, txtscname.Top + txtscname.Height + Picture1.Top + 20
lstStaff.Move txtStaff.Left + Picture1.Left, txtStaff.Top + txtStaff.Height + Picture1.Top + 20
lstMonth.Move txtMonth.Left, txtMonth.Top + txtMonth.Height + 30
lstTask.Move txtJD.Left, txtJD.Top + txtJD.Height + 30

txtMonth = Trim(Mid(lstMonth.list(Val(Mid(myDate, 6, 2) - 1)), InStr(1, lstMonth.list(Val(Mid(myDate, 6, 2) - 1)), ":") + 1))
lstMonth.Selected(Val(Mid(myDate, 6, 2) - 1)) = True


Refress_Rs rsData, "Select * from ViewUsers where sn = " & CurrenUser
Me.Caption = rsData!UserFullName & ",  Level : " & rsData!TypeName

getsc

If userType = 2 Then
    txtscname = CurrenBranchName
    txtscname.tag = CurrenBranchID
    txtscname.Enabled = False
    cmdAddSC.Visible = False
    cmdEdit.Visible = False
    StaffData
    Staffname
End If
getTaskList
If userType = 1 Then
    txtscname.TabIndex = 0
    Picture1.BackColor = &HC0C0FF
    Me.BackColor = &HC0C0FF
    
Else
    txtStaff.TabIndex = 0
    Picture1.BackColor = &H80C0FF
    Me.BackColor = &H80C0FF
End If
End Sub

Public Sub getsc()
Refress_Rs rsData, "Select * from tblServiceCenter order by SN "
lstSCname.Clear
If rsData.RecordCount > 0 Then
    For I = 1 To rsData.RecordCount
        lstSCname.AddItem Format(rsData!code, "000") & " : " & rsData!ServiceCenterName
        lstSCname.ItemData(I - 1) = rsData!SN
        rsData.MoveNext
    Next
End If

If lstSCname.ListCount > 0 Then
    lstSCname.Selected(0) = True
End If

End Sub

Private Sub getAsignTask()
Dim rsTask As Recordset
Set rsTask = New Recordset
Refress_Rs rsTask, "Select * from ViewGivenTask where taskTo = " & Val(txtscname.tag) & " and brich = " & Val(txtStaff.tag) & " and yearr = " & Val(txtYear) & " and monthh = " & Val(txtMonth.tag)
lstTask.Clear
If rsTask.RecordCount > 0 Then
    rsTask.MoveFirst
    Do While Not rsTask.EOF
        lstTask.AddItem rsTask!proitity & " : " & rsTask!TaskName
        lstTask.ItemData(rsTask.AbsolutePosition - 1) = rsTask!taskID
        rsTask.MoveNext
    Loop
    lstTask.ListIndex = 0
End If

End Sub

Private Sub getMonth()
    lstMonth.AddItem "01 : Baishakh"
    lstMonth.AddItem "02 : Jestha"
    lstMonth.AddItem "03 : Ashar"
    lstMonth.AddItem "04 : Shrawan"
    lstMonth.AddItem "05 : Bhadra"
    lstMonth.AddItem "06 : Ashoj"
    lstMonth.AddItem "07 : Kartik"
    lstMonth.AddItem "08 : Mansir"
    lstMonth.AddItem "09 : Poush"
    lstMonth.AddItem "10 : Magh"
    lstMonth.AddItem "11 : Falgun"
    lstMonth.AddItem "12 : Chaitra"
End Sub

Private Sub lblExport_Click()
    lblExport.Enabled = False
    ExportToExcelFromRecordSet rsTarget, myYes, myno, "Target Report of " & txtStaff & " As On " & txtYear & " " & txtMonth
    Message "Data Export Succesfully."
    lblExport.Enabled = True
End Sub

Private Sub lstMonth_DblClick()
txtMonth_KeyPress 13
End Sub

Private Sub lstMonth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonth_KeyPress 13
End If
End Sub

Private Sub lstSCname_DblClick()
txtscname_KeyPress 13
End Sub

Private Sub lstSCname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtscname_KeyPress 13
End If
End Sub

Private Sub lstStaff_DblClick()
txtStaff_KeyPress 13
End Sub

Private Sub lstStaff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtStaff_KeyPress 13
End If
End Sub

Private Sub lstTask_DblClick()
txtJD_KeyPress 13
End Sub

Private Sub lstTask_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtJD_KeyPress 13
End If
End Sub

Private Sub Picture1_GotFocus()
listVisible Picture1.Name
End Sub

Private Sub txtarget_GotFocus()
listVisible txtarget.Name
Colored
End Sub

Private Sub txtarget_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift
End Sub

Private Sub txtarget_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd.SetFocus
End If
End Sub

Private Sub txtarget_LostFocus()
unColored txtarget
End Sub

Private Sub txtJD_Change()
If lstTask.Visible = True Then
    If Not txtJD.Text = "" Then
        For I = 0 To lstTask.ListCount - 1
            If Val(txtJD) = Val(lstTask.list(I)) Then
                lstTask.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtJD.Text)) = UCase(Trim(Mid(lstTask.list(I), InStr(1, lstTask.list(I), ":") + 2, Len(txtJD.Text)))) Or Val(txtJD.Text) = Val(Mid(lstTask.list(I), 1, InStr(1, lstTask.list(I), ":") - 2)) Then
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

Private Sub txtJD_GotFocus()
Colored
listVisible txtJD.Name
End Sub

Private Sub txtJD_KeyDown(KeyCode As Integer, Shift As Integer)

GetShotcutkey KeyCode, Shift


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

Private Sub txtJD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstTask.Visible = True Then
        If lstTask.ListIndex >= 0 Then
            If lstTask.ItemData(lstTask.ListIndex) > 0 Then
                txtJD.tag = lstTask.ItemData(lstTask.ListIndex)
                txtJD.Text = Trim(Mid(lstTask.Text, InStr(1, lstTask.Text, ":") + 1))
                AppDets
                txtarget.SetFocus
            Else
                frmTask.Show vbModal
            End If
        End If
    End If
End If

End Sub

Private Sub txtJD_LostFocus()
unColored txtJD
End Sub

Private Sub txtMonth_Change()
If lstMonth.Visible = True Then
    If Not txtMonth.Text = "" Then
        For I = 0 To lstMonth.ListCount - 1
            If Val(txtMonth) = Val(lstMonth.list(I)) Then
                lstMonth.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtMonth.Text)) = UCase(Trim(Mid(lstMonth.list(I), InStr(1, lstMonth.list(I), ":") + 2, Len(txtMonth.Text)))) Or Val(txtMonth.Text) = Val(Mid(lstMonth.list(I), 1, InStr(1, lstMonth.list(I), ":") - 2)) Then
                    lstMonth.Selected(I) = True
                    Exit For
                Else
                    lstMonth.Selected(I) = False
                End If
            End If
        Next
    Else
        lstMonth.Selected(0) = False
    End If
End If



End Sub

Private Sub txtMonth_GotFocus()
listVisible txtMonth.Name
Colored
End Sub

Private Sub txtMonth_KeyDown(KeyCode As Integer, Shift As Integer)

GetShotcutkey KeyCode, Shift


If lstMonth.Visible = True Then
    If KeyCode = 38 Then
        If lstMonth.ListIndex <= 0 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstMonth.Selected(lstMonth.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstMonth.ListIndex = lstMonth.ListCount - 1 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstMonth.Selected(lstMonth.ListIndex + 1) = True
        End If
    End If
End If


End Sub

Private Sub txtMonth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstMonth.Visible = True Then
        If lstMonth.ListIndex >= 0 Then
            txtMonth.tag = Val(lstMonth.Text)
            txtMonth.Text = Trim(Mid(lstMonth.Text, InStr(1, lstMonth.Text, ":") + 1))
            txtJD.SetFocus
            
            If StaffSelection Then
                getTarget False
                AppDets
            Else
                getTarget False
            
            End If
            
        End If
    End If
End If

End Sub

Private Sub txtMonth_LostFocus()
unColored txtMonth
End Sub

Private Sub txtscname_Change()
If lstSCname.Visible = True Then
    If Not txtscname.Text = "" Then
        For I = 0 To lstSCname.ListCount - 1
            If Val(txtscname) = Val(lstSCname.list(I)) Then
                lstSCname.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtscname.Text)) = UCase(Trim(Mid(lstSCname.list(I), InStr(1, lstSCname.list(I), ":") + 2, Len(txtscname.Text)))) Or Val(txtscname.Text) = Val(Mid(lstSCname.list(I), 1, InStr(1, lstSCname.list(I), ":") - 2)) Then
                    lstSCname.Selected(I) = True
                    Exit For
                Else
                    lstSCname.Selected(I) = False
                End If
            End If
        Next
    Else
        lstSCname.Selected(0) = False
    End If
End If


End Sub

Private Sub txtscname_GotFocus()
Colored
listVisible txtscname.Name
End Sub

Private Sub txtscname_KeyDown(KeyCode As Integer, Shift As Integer)

GetShotcutkey KeyCode, Shift


If lstSCname.Visible = True Then
    If KeyCode = 38 Then
        If lstSCname.ListIndex <= 0 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstSCname.Selected(lstSCname.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstSCname.ListIndex = lstSCname.ListCount - 1 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstSCname.Selected(lstSCname.ListIndex + 1) = True
        End If
    End If
End If

End Sub


Private Sub txtscname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstSCname.Visible = True Then
        If lstSCname.ListIndex >= 0 Then
            If lstSCname.ItemData(lstSCname.ListIndex) > 0 Then
                Dim rsInfo As Recordset
                Set rsInfo = New Recordset
                Refress_Rs rsInfo, "Select * from tblServiceCenter where sn = " & Val(lstSCname.ItemData(lstSCname.ListIndex))
                If rsInfo.RecordCount > 0 Then
                    txtscname.tag = lstSCname.ItemData(lstSCname.ListIndex)
                    lblScCode = Format(Val(lstSCname.Text), "000")
                    txtscname.Text = "  " & Mid(lstSCname.Text, 7)
                    CurrentData = lstSCname.ItemData(lstSCname.ListIndex)
                    lblAddress = rsInfo!Address & " - " & rsInfo!Phone
                    AppDets
                    getTarget False
                    StaffSelection = False
                    StaffData
                    Staffname
                    txtStaff.SetFocus
                End If
                
            End If
        End If
    End If
End If

End Sub

Private Sub StaffData()
Dim rsStaff As Recordset
Set rsStaff = New Recordset
If userType <> 3 Then

    Refress_Rs rsStaff, "Select * from ViewUsers where Status = 0 and Branch = " & Val(txtscname.tag) & " and userType<>1 order by UserType, userFullName "
    
    lstStaff.Clear
    
    If rsStaff.RecordCount > 0 Then
        rsStaff.MoveFirst
        Do While Not rsStaff.EOF
            lstStaff.AddItem rsStaff!UserFullName & "(" & rsStaff!userID & ")"
            lstStaff.ItemData(rsStaff.AbsolutePosition - 1) = rsStaff!SN
            rsStaff.MoveNext
        Loop
        lstStaff.Selected(0) = True
    End If
End If


End Sub

Private Sub Staffname()
Dim rsStaff As Recordset
Set rsStaff = New Recordset
    
    Refress_Rs rsStaff, "Select * from viewServiceCenter where sn = " & Val(txtscname.tag)
    If rsStaff.RecordCount > 0 Then
        txtStaff = rsStaff!UserFullName
    Else
        txtStaff = ""
    End If
End Sub

Private Sub AppDets(Optional xx As Integer = 0)
If StaffSelection Then
    dg.Height = 2175
    dg1.Visible = True
    
    Dim rsList As Recordset
    Dim sss As String
    If xx = 0 Then
        sss = " Approved = 0 and "
    Else
        sss = " taskName = '" & dg.Columns(0) & "' and "
    End If
    
    Set rsList = New Recordset
    Refress_Rs rsList, "Select DateD, TaskName, Achive, Status = case Approved when 0 then 'Not-Approved' else 'Approved' end, AchiveSN from viewTaskAchiveDets where " & sss & "yearr=" & Val(txtYear) & " and monthh=" & Val(txtMonth.tag) & " and brich = " & Val(txtStaff.tag)
    
    If Not rsList.RecordCount > 0 Then
        dg.Height = 5415
        dg1.Visible = False
        Exit Sub
    End If
    
    Set dg1.DataSource = rsList
    dg1.RowHeight = 325
    dg1.Columns(0).Width = 1500
    dg1.Columns(1).Width = 1900
    dg1.Columns(2).Width = 1900
    dg1.Columns(2).Alignment = dbgRight
    dg1.Columns(3).Width = 1900
    dg1.Columns(3).Alignment = dbgRight
    dg1.Columns(4).Visible = False
Else
    dg.Height = 5415
    dg1.Visible = False
End If
End Sub

Private Sub getTarget(Optional isBanch As Boolean)
Set rsTarget = New Recordset
fac = " and toUser = " & Val(txtStaff.tag)

Refress_Rs rsTarget, "Select TaskName, Target, Achive, Achive-target Diff, round(Achive*100/(Target),0) Percentage, SN from viewTaskAchive where target>0 and yearr = " & Val(txtYear) & " and monthh = " & Val(txtMonth.tag) & " and taskTo = " & Val(txtscname.tag) & fac
Set dg.DataSource = rsTarget
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
dg.Columns(5).Visible = False
'GenerateListView lv, rsTarget
'
'
'
'If lv.ColumnHeaders.Count > 6 Then
'    lv.ColumnHeaders(1).Width = 0
'    lv.ColumnHeaders(2).Width = 550
'    lv.ColumnHeaders(3).Width = 2500
'    lv.ColumnHeaders(4).Width = 1600
'    lv.ColumnHeaders(5).Width = 1600
'    lv.ColumnHeaders(6).Width = 1600
'    lv.ColumnHeaders(7).Width = 1000
'End If
'
'For I = 1 To rsTarget.RecordCount
'    lv.ListItems(I).SubItems(1) = I
'Next

If userType = 2 And StaffSelection = False Then
    getAsignTask
End If

End Sub

Private Sub txtscname_LostFocus()
unColored txtscname
End Sub


Private Sub GetShotcutkey(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If txtscname.Enabled Then
        txtscname.SetFocus
    End If
ElseIf KeyCode = 113 Then
    txtStaff.SetFocus
End If
End Sub

Public Sub getTaskList()
Dim rsTask As Recordset
Set rsTask = New Recordset
Refress_Rs rsTask, "Select * from tblTask order by Proitity"
lstTask.Clear
If rsTask.RecordCount > 0 Then
    rsTask.MoveFirst
    Do While Not rsTask.EOF
        lstTask.AddItem Format(rsTask!proitity, "00") & " : " & rsTask!TaskName
        lstTask.ItemData(rsTask.AbsolutePosition - 1) = rsTask!SN
        rsTask.MoveNext
    Loop
End If
If userType = 1 Then
    lstTask.AddItem "00 : Create New Task"
    lstTask.ItemData(lstTask.ListCount - 1) = 0
End If
If lstTask.ListCount > 0 Then
    lstTask.Selected(0) = True
End If
End Sub

Private Sub txtStaff_Change()
If lstStaff.Visible = True Then
    If Not txtStaff.Text = "" Then
        For I = 0 To lstStaff.ListCount - 1
            If Trim(UCase(Mid(txtStaff, 1, Len(txtStaff)))) = Trim(UCase(Mid(lstStaff.list(I), 1, Len(txtStaff)))) Then
                lstStaff.Selected(I) = True
                Exit For
            Else
                lstStaff.Selected(I) = False
            End If
        Next
    Else
        lstStaff.Selected(0) = False
    End If
End If



End Sub

Private Sub txtStaff_GotFocus()
Colored
listVisible txtStaff.Name
If Len(txtStaff) > 2 Then
    txtStaff = Trim(txtStaff)
End If
End Sub

Private Sub txtStaff_KeyDown(KeyCode As Integer, Shift As Integer)

GetShotcutkey KeyCode, Shift


 If lstStaff.Visible = True Then
    If KeyCode = 38 Then
        If lstStaff.ListIndex <= 0 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstStaff.Selected(lstStaff.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstStaff.ListIndex = lstStaff.ListCount - 1 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lstStaff.Selected(lstStaff.ListIndex + 1) = True
        End If
    End If
End If

End Sub

Private Sub txtStaff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstStaff.Visible = True Then
        If lstStaff.ListIndex >= 0 Then
            If lstStaff.ItemData(lstStaff.ListIndex) > 0 Then
                Dim rsInfo As Recordset
                Set rsInfo = New Recordset
                
                Refress_Rs rsInfo, "Select * from ViewUsers where SN = " & Val(lstStaff.ItemData(lstStaff.ListIndex))
                If rsInfo.RecordCount > 0 Then
                    
                    txtStaff = "  " & rsInfo!UserFullName
                    txtStaff.tag = rsInfo!SN
                    lblphone = "  " & rsInfo!Post & ",  " & rsInfo!TypeName
                    lblContact = "  " & rsInfo!Phone
                    txtYear_GotFocus
                    txtYear_KeyPress 13
                    txtMonth_GotFocus
                    txtMonth_KeyPress 13

                    If userType = 2 And rsInfo!SN = rsInfo!incharge Then
                        Label3(3).Caption = "Achive"
                        getAsignTask
                    Else
                        Label3(3).Caption = "Target"
                        getTaskList
                    End If

                    If rsInfo!SN = rsInfo!incharge Then
                        getTarget True
                        StaffSelection = False
                    Else
                        StaffSelection = True
                        getTarget False
'                        Message "Staff Target will Apear here."
                    End If
                    Refress_Rs rsInfo, "Select Top 1 * from tblLoginLog where userNo = " & Val(txtStaff.tag) & " order by sn Desc "
                    If rsInfo.RecordCount > 0 Then
                        lblNOte = "---Last Login ---" & vbNewLine & " PCName -> " & rsInfo!PCName & vbNewLine & " Date -> " & rsInfo!dated & vbNewLine & " Time -> " & rsInfo!timee
                    Else
                        lblNOte = ""
                    End If
                    AppDets
                    
                End If
                txtJD.SetFocus
            End If
        End If
    End If
End If


End Sub

Private Sub txtStaff_LostFocus()
unColored txtStaff
End Sub

Private Sub txtYear_Change()
If StaffSelection Then
    getTarget False
    AppDets
Else
    getTarget True
End If
End Sub

Private Sub txtYear_GotFocus()
txtYear = Mid(myDate, 1, 4)
listVisible txtYear.Name
Colored
End Sub

Private Sub txtYear_KeyDown(KeyCode As Integer, Shift As Integer)
GetShotcutkey KeyCode, Shift
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonth.SetFocus
End If
End Sub

Private Sub txtYear_LostFocus()
unColored txtYear
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMFMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MF Form"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6675
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
   ScaleHeight     =   7305
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5741
      _Version        =   393216
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
   Begin VB.TextBox txtday 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtSC 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2160
      Width           =   4575
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox txtSN 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCance 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6615
      Y1              =   3240
      Y2              =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Collection Day :"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Branch :"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Address :"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "MF Name :"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "SN :"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmMFMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMF As Recordset
Dim isnewRecord As Boolean

Private Sub LockData()
    txtAddress.Locked = True
    txtday.Locked = True
    txtName.Locked = True
    txtSC.Locked = True
    txtSN.Locked = True
    cmdEdit.Enabled = True
    cmdNew.Enabled = True
    cmdSave.Enabled = False
End Sub

Private Sub unLockData()
    txtAddress.Locked = False
    txtday.Locked = False
    txtName.Locked = False
    txtSC.Locked = False
    cmdEdit.Enabled = False
    cmdNew.Enabled = False
    cmdSave.Enabled = True
End Sub

Private Sub clrData()
    txtAddress = ""
    txtday = ""
    txtName = ""
    txtSC = ""
    txtSN = ""
End Sub


Private Sub getDatainDG()
rsMF.Requery
Set dg.DataSource = rsMF
dg.Columns(0).Width = 500
dg.Columns(1).Width = 2500
dg.Columns(2).Width = 2500
dg.Columns(3).Visible = False
dg.Columns(4).Visible = False
dg.Scroll dg.Col, dg.row
End Sub

Private Sub cmdCance_Click()
If cmdSave.Enabled Then
    LockData
Else
    Unload Me
End If
End Sub

Private Sub cmdEdit_Click()
If Val(txtSN) > 0 Then
    isnewRecord = False
    unLockData
    txtName.SetFocus
End If
End Sub

Private Sub cmdNew_Click()
isnewRecord = True
clrData
unLockData
txtSN = NewMaxID("tblMFGroup", "SN")
txtName.SetFocus
End Sub

Private Sub cmdSave_Click()
Message "Do you want to save Data ?", YesNo, True
If CurrentMsgResponce = Yes Then
    SaveData
    getDatainDG
    Message "Data Save sucessfully."
End If
End Sub



Private Sub SaveData()
Dim strSQL As String
If isnewRecord Then
    strSQL = "Insert into tblMFGroup values (" & NewMaxID("tblMFGroup", "SN") & ", '" & Replace(txtName, "'", "''") & "', '" & Replace(txtAddress, "'", "''") & "', " & Val(txtSC.tag) & ", " & Val(txtday) & " )"
Else
    strSQL = "update tblMFGroup set groupname = '" & Replace(txtName, "'", "''") & "', groupAddress = '" & Replace(txtAddress, "'", "''") & "', Branch = " & Val(txtSC.tag) & ", dayCode = " & Val(txtday) & " where sn = " & Val(txtSN)
End If
ExecuteQuery strSQL
End Sub

Private Sub Form_Load()
Set rsMF = New Recordset
Refress_Rs rsMF, "Select * from tblMFGroup where branch = " & CurrenBranchID
getDatainDG
txtSC = CurrenBranchName
txtSC.tag = CurrenBranchID
End Sub

Private Sub dg_DblClick()
txtAddress = dg.Columns(2)
txtday = dg.Columns(4)
txtName = dg.Columns(1)
txtSC = CurrenBranchName
txtSC.tag = CurrenBranchID
txtSN = dg.Columns(0)
End Sub

Private Sub txtAddress_GotFocus()
Colored
End Sub

Private Sub txtAddress_LostFocus()
unColored txtAddress
End Sub

Private Sub txtday_GotFocus()
Colored
End Sub

Private Sub txtday_LostFocus()
unColored txtday
End Sub

Private Sub txtName_GotFocus()
Colored
End Sub

Private Sub txtName_LostFocus()
unColored txtName
End Sub

Private Sub txtSC_GotFocus()
Colored
txtSC = CurrenBranchName
txtSC.tag = CurrenBranchID
End Sub

Private Sub txtSC_LostFocus()
unColored txtSC
End Sub

Private Sub txtSN_GotFocus()
Colored
End Sub

Private Sub txtSN_LostFocus()
unColored txtSN
End Sub

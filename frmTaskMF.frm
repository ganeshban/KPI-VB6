VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTask 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Task"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
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
   ScaleHeight     =   6315
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTask 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6376
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
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"frmTaskMF.frx":0000
      Height          =   855
      Left            =   -240
      TabIndex        =   9
      Top             =   -120
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5895
      Y1              =   1920
      Y2              =   1935
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTask As Recordset
Dim isnewRecord As Boolean


Private Sub LockData()
txtID.Locked = True
txtTask.Locked = True
cmdEdit.Enabled = True
cmdNew.Enabled = True
cmdSave.Enabled = False

End Sub

Private Sub unLockData()
txtID.Locked = False
txtTask.Locked = False
cmdEdit.Enabled = False
cmdNew.Enabled = False
cmdSave.Enabled = True

End Sub

Private Sub clrData()
txtID = ""
txtTask = ""
End Sub


Private Sub cmdCancel_Click()
If cmdSave.Enabled Then
    LockData
Else
    Unload Me
End If
End Sub

Private Sub cmdEdit_Click()
If Val(txtID.tag) > 0 Then
    unLockData
    txtTask.SetFocus
    isnewRecord = False
End If
End Sub

Private Sub cmdNew_Click()

unLockData
clrData
txtID.tag = NewMaxID("tblTask", "sn")
txtID = NewMaxID("tblTask", "proitity")
isnewRecord = True
txtTask.SetFocus

End Sub

Private Sub cmdSave_Click()
Message "Are you sure to Save Data ?", YesNo, True
If CurrentMsgResponce = Yes Then
    Dim sstr As String
    If isnewRecord Then
        sstr = "Insert into tblTask values( " & NewMaxID("tblTask", "SN") & ", '" & Replace(txtTask, "'", "''") & "', " & NewMaxID("tblTask", "proitity") & ")"
    Else
        sstr = "Update tblTask set proitity = proitity + 1 where proitity >= " & Val(txtID) & " update tblTask set taskName = '" & Replace(txtTask, "'", "''") & "', proitity = " & Val(txtID) & " where sn = " & Val(txtID.tag)
    End If
    
    ExecuteQuery sstr
    LockData
    rsTask.Requery
    getList
    Message "Data Save Succesfully."

End If
End Sub

Private Sub dg_DblClick()
On Error Resume Next
If Not cmdSave.Enabled = True Then
    txtID = dg.Columns(2).Value & ""
    txtTask = dg.Columns(1).Value & ""
    txtID.tag = dg.Columns(0).Value & ""
End If
End Sub

Private Sub Form_Load()
Set rsTask = New Recordset
Refress_Rs rsTask, "Select * from tbltask"
getList
LockData
End Sub

Private Sub getList()
Set dg.DataSource = rsTask
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmHead.getTaskList
End Sub

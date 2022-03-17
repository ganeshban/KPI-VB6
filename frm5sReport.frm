VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm5sReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "5S Report"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
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
   ScaleHeight     =   7290
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   615
      Left            =   7920
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   5055
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8916
      _Version        =   393216
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
   Begin VB.CommandButton cmdgetrpt 
      Caption         =   "Get Report"
      Height          =   615
      Left            =   6120
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.ComboBox cmbuser 
      Height          =   390
      Left            =   3120
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.ComboBox cmbsc 
      Height          =   390
      Left            =   240
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   1
      Left            =   3120
      TabIndex        =   8
      Top             =   720
      Width           =   2655
      Begin VB.TextBox txtmonth2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtyear2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Option"
      Height          =   1215
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "Numerical"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "Pie Chart"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2655
      Begin VB.TextBox txtyear1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtMonth1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm5sReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsData As Recordset


Private Sub cmbSC_Click()
Refress_Rs rsData, "Select * from ViewUsers where status = 0 and Branch = " & Val(cmbSC) & " order by userFullName "
If rsData.RecordCount > 0 Then
    cmbuser.Clear
    cmbuser.AddItem "00 : All Users"
    Do While Not rsData.EOF
        cmbuser.AddItem rsData!sn & " : " & rsData!UserFullName
        rsData.MoveNext
    Loop
    cmbuser = cmbuser.List(0)
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdgetrpt_Click()
If opt(0).Value Then
    GetpieReport
Else
    GetNumericalReport
End If
End Sub

Private Sub Form_Load()
Set rsData = New Recordset

txtyear1 = Val(Mid(myDate, 1, 4))
txtyear2 = txtyear1
txtMonth1 = Val(Mid(myDate, 6, 2))
txtmonth2 = txtMonth1
Refress_Rs rsData, "Select * from tblServiceCenter"
If rsData.RecordCount > 0 Then
    Do While Not rsData.EOF
        cmbSC.AddItem rsData!code & " : " & rsData!ServiceCenterName
        rsData.MoveNext
    Loop
    cmbSC = cmbSC.List(0)
End If

Refress_Rs rsData, "Select * from ViewUsers where status = 0 and Branch = " & Val(cmbSC)
If rsData.RecordCount > 0 Then
    cmbuser.Clear
    cmbuser.AddItem "000 : All Users"
    Do While Not rsData.EOF
        cmbuser.AddItem Format(rsData!sn, "000") & " : " & rsData!UserFullName
        rsData.MoveNext
    Loop
    cmbuser = cmbuser.List(0)
End If

End Sub

Private Sub opt1_Click()

End Sub

Private Sub opt_Click(Index As Integer)
If opt(0).Value = True Then
    txtyear2.Enabled = False
    txtmonth2.Enabled = False
Else
    txtyear2.Enabled = True
    txtmonth2.Enabled = True
End If
End Sub

Private Sub GetNumericalReport()
'Refress_Rs rsData, "Select * from View"
End Sub

Private Sub GetpieReport()
If Val(cmbuser) > 0 Then
    Dim xx As String
    xx = "Select * from View5sData d Where yearr = " & txtyear1 & " And monthh = " & txtMonth1 & " And userno = " & Val(cmbuser) & " And branch = " & Val(cmbSC)
    Refress_Rs rsData, xx
    If rsData.RecordCount > 0 Then
        frmsubRpt.ch.RowCount = rsData.RecordCount
        frmsubRpt.ch.ColumnCount = 2
        frmsubRpt.lblNOte = "Branch " & cmbSC & " user " & cmbuser & "  Year " & txtyear1 & " Month " & txtMonth1
        Do While Not rsData.EOF
            frmsubRpt.ch.row = rsData.AbsolutePosition
            frmsubRpt.ch.RowLabel = rsData!QstID
            frmsubRpt.ch.Column = 1
            frmsubRpt.ch.data = rsData!ansID
            
            
            frmsubRpt.ch.Column = 2
            frmsubRpt.ch.data = 5 - rsData!ansID
            
            rsData.MoveNext
        Loop
        frmsubRpt.Show vbModal
    Else
        Message "Data not found"
        Exit Sub
    End If
Else
    xx = "Select Code, QstID, Qstname, (select sum(ansID)/count(*) from View5sData where branch=d.branch and isadmin=d.isadmin and yearr=d.yearr and monthh=d.monthh and qstID=d.QstID) val from view5sData d Where Branch = " & Val(cmbSC) & " And isadmin = 0 And yearr = " & txtyear1 & " And monthh = " & txtMonth1 & "  group by code, qstID, QstName, branch, isadmin, yearr, monthh"
    Refress_Rs rsData, xx
    If rsData.RecordCount > 0 Then
        frmsubRpt.ch.RowCount = rsData.RecordCount
        frmsubRpt.ch.ColumnCount = 2
        frmsubRpt.lblNOte = "Branch " & cmbSC & " Year " & txtyear1 & " Month " & txtMonth1
        Do While Not rsData.EOF
            frmsubRpt.ch.row = rsData.AbsolutePosition
            frmsubRpt.ch.RowLabel = rsData!QstID
            frmsubRpt.ch.Column = 1
            frmsubRpt.ch.data = rsData!Val
            
            
            frmsubRpt.ch.Column = 2
            frmsubRpt.ch.data = 5 - rsData!Val
            
            rsData.MoveNext
        Loop
    frmsubRpt.Show vbModal
    End If

End If
End Sub

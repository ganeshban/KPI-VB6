VERSION 5.00
Begin VB.Form frm5S 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "5S Form"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10500
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
   ScaleHeight     =   3825
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdreport 
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
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbsc 
      Height          =   390
      Left            =   120
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdpg 
      Caption         =   "Add Indicater"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   10305
      TabIndex        =   0
      Top             =   720
      Width           =   10335
      Begin VB.CommandButton cmdAns 
         BackColor       =   &H00FFFFFF&
         Caption         =   "fasdfa"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdAns 
         BackColor       =   &H00FFFFFF&
         Caption         =   "fasdfa"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdAns 
         BackColor       =   &H00C0E0FF&
         Caption         =   "fasdfa"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdAns 
         BackColor       =   &H00FFFFFF&
         Caption         =   "fasdfa"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdAns 
         BackColor       =   &H00FFFFFF&
         Caption         =   "fasdfa"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblCat 
         BackStyle       =   0  'Transparent
         Caption         =   "S1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblCatName 
         BackStyle       =   0  'Transparent
         Caption         =   "S1"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblQst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "oxf ePsf] *]:ssf] cj:yf s:tf] % <"
         BeginProperty Font 
            Name            =   "Sagarmatha"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   9975
      End
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8040
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frm5S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs5s As Recordset
Dim rsQ As Recordset
Dim thisYear As Integer
Dim thisMonth As Integer
Dim thisuser As Integer


Private Sub cmdAns_Click(Index As Integer)
Dim sstr As String
Dim xx As String
xx = "Select * from View5sData d Where QstID = " & Val(lblQst.Tag) & " and yearr = " & thisYear & " And monthh = " & thisMonth & " And userno = " & thisuser & " And branch = " & Val(cmbsc)
Refress_Rs rs5s, xx
If rs5s.RecordCount > 0 Then
    Message "You have alrady posted this indicater of 5S. Now it will move to next indicater. Please Remind it Data is already save."
    cmdAns(Index).BackColor = &HC0E0FF
    rsQ.MoveNext
    LoadQstAns
    Exit Sub
End If

If userType = 1 Then
    sstr = "insert into tbl5SAnswer values(" & NewMaxID("tbl5SAnswer", "SN") & ", " & Val(lblQst.Tag) & ", " & Val(cmdAns(Index).Tag) & ", " & Val(Mid(myDate, 6, 2)) & ", " & Val(Mid(myDate, 1, 4)) & ", " & CurrenUser & ", " & Val(cmbsc.Text) & ", 1, '" & myDate & "' )"
Else
    sstr = "insert into tbl5SAnswer values(" & NewMaxID("tbl5SAnswer", "SN") & ", " & Val(lblQst.Tag) & ", " & Val(cmdAns(Index).Tag) & ", " & Val(Mid(myDate, 6, 2)) & ", " & Val(Mid(myDate, 1, 4)) & ", " & CurrenUser & ", " & Val(cmbsc.Text) & ", 0, '" & myDate & "' )"
End If

ExecuteQuery sstr
cmdAns(Index).BackColor = &HC0E0FF
rsQ.MoveNext
LoadQstAns
End Sub



Private Sub cmdAns_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
cmdAns(Index).BackColor = &HC0E0FF
End Sub

Private Sub cmdBranch_Click()

End Sub

Private Sub cmdAns_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
cmdAns(Index).BackColor = &HFFFFFF

End Sub

Private Sub cmdpg_Click()
frm5sEntryFom.Show vbModal
End Sub

Private Sub cmdReport_Click()
Dim xx As String
If userType = 1 Then
    frm5sReport.Show vbModal
    Exit Sub
Else
    xx = "Select * from View5sData d Where yearr = " & thisYear & " And monthh = " & thisMonth & " And userno = " & thisuser & " And branch = " & Val(cmbsc)
End If
Refress_Rs rs5s, xx
If rs5s.RecordCount > 0 Then
    frmsubRpt.ch.RowCount = rs5s.RecordCount
    frmsubRpt.ch.ColumnCount = 2
    frmsubRpt.lblNOte = "Year " & thisYear & " Month " & thisMonth
    Do While Not rs5s.EOF
        frmsubRpt.ch.row = rs5s.AbsolutePosition
        frmsubRpt.ch.RowLabel = rs5s!QstID
        frmsubRpt.ch.Column = 1
        frmsubRpt.ch.data = rs5s!ansID
        
        
        frmsubRpt.ch.Column = 2
        frmsubRpt.ch.data = 5 - rs5s!ansID
        
        rs5s.MoveNext
    Loop
Else
    Message "Record is not entered."
    Exit Sub
End If
frmsubRpt.Show vbModal

End Sub

Private Sub Form_Load()
    Set rs5s = New Recordset
    Set rsQ = New Recordset
    lblDate = "Date : " & myDate
    thisYear = Val(Mid(myDate, 1, 4))
    thisMonth = Val(Mid(myDate, 6, 2))
    thisuser = CurrenUser
    Refress_Rs rs5s, "Select * from tblServicecenter"
    If rs5s.RecordCount > 0 Then
        Do While Not rs5s.EOF
            cmbsc.AddItem rs5s!code & ":" & Replace(rs5s!ServiceCenterName, "Service Center ", "")
            cmbsc.ItemData(rs5s.AbsolutePosition - 1) = rs5s!SN
            rs5s.MoveNext
        Loop
        cmbsc.Text = cmbsc.List(CurrenBranchID - 1)
    End If
    Refress_Rs rs5s, "Select * from view5sData where yearr = " & thisYear & " and Monthh = " & thisMonth & " and userNo = " & thisuser
    Refress_Rs rsQ, "Select * from view5sQuestion order by sn, QstID "
    LoadQstAns
    If userType = 1 Then
        cmdpg.Visible = True
        cmbsc.Enabled = True
    Else
        cmdpg.Visible = False
        cmbsc.Enabled = False
    End If
End Sub

Private Sub LoadQstAns()
Dim x As Integer
If rsQ.RecordCount > 0 Then
    If Not rsQ.EOF Then
        lblCatName = rsQ!Point
        lblCat = rsQ!Catname
        lblQst = NumirecToCharecter(rsQ!QstID) & "_ " & rsQ!QstName
        lblQst.Tag = rsQ!QstID
        Dim rsAns As Recordset
        Set rsAns = New Recordset
        Refress_Rs rsAns, "Select * from tbl5sansList where qstID = " & Val(rsQ!QstID) & " order by ansID"
        If rsAns.RecordCount > 0 Then
            For x = 0 To 4
                cmdAns(x).Visible = False
            Next
            x = 0
            rsAns.MoveFirst
            Do While Not rsAns.EOF
                cmdAns(x).Visible = True
                cmdAns(x).Caption = rsAns!AnsName
                cmdAns(x).Tag = rsAns!SN
                cmdAns(x).BackColor = &HFFFFFF

                x = x + 1
                rsAns.MoveNext
            Loop
        End If
    Else
        Message " Thank you !"
        Unload Me
        cmdReport_Click
    End If
End If
End Sub

Private Sub updateChart()
    rs5s.Requery
    If rs5s.RecordCount > 0 Then
        ch.row = 1
    End If
End Sub


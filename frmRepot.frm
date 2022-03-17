VERSION 5.00
Begin VB.Form frmRepot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Form"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7890
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
   ScaleHeight     =   8130
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSc 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdBranchReport 
      Caption         =   "All Branch Report"
      Height          =   495
      Left            =   5760
      TabIndex        =   12
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last Date"
      Height          =   1215
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
      Begin VB.TextBox txtYear2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtMonth2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "First Date"
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
      Begin VB.TextBox txtMonth1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtyear1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Branch"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Form"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmRepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBranchReport_Click()
Dim x As Integer
Dim target As Long
Dim achive As Long

Dim rsChart As Recordset
Set rsChart = New Recordset
Refress_Rs rsChart, "Select ServiceCenterName, Target, Achive, TaskName, (Select Count(Code) from ViewTaskAchive where Yearr = " & Val(txtyear2) & " and monthh = " & Val(txtmonth2) & " and Brich = Incharge and Code=d.Code) Cnt from ViewTaskAchive d where yearr=" & Val(txtyear2) & " and monthh = " & Val(txtmonth2) & " and brich = Incharge order by Code, Proitity"

If rsChart.RecordCount > 0 Then
    x = 1
    frmsubRpt.ch.RowCount = x
    frmsubRpt.ch.ColumnCount = 2
    frmsubRpt.lblNOte = "Overall Target Report Year : " & Val(txtyear2) & " Month : " & Val(txtmonth2.tag)

    Do While Not rsChart.EOF
        frmsubRpt.ch.row = x
        For I = 1 To rsChart!cnt
        
            If rsChart!achive >= rsChart!target Then
                target = target + (100 / rsChart!cnt)
                achive = achive + 0
            Else
                target = target + ((rsChart!achive * 100 / (rsChart!target + 1)) / rsChart!cnt)
                achive = achive + (((rsChart!target + 1 - rsChart!achive) * 100 / (rsChart!target + 1)) / rsChart!cnt)
            End If
            
            If I = rsChart!cnt Then
                frmsubRpt.ch.RowLabel = Replace(rsChart!ServiceCenterName, "Service Center", "")
            Else
                rsChart.MoveNext
            End If
            
            
        Next
        
        frmsubRpt.ch.row = x
                
        frmsubRpt.ch.Column = 1
        frmsubRpt.ch.data = target
        frmsubRpt.ch.Column = 2
        frmsubRpt.ch.data = achive
        
        
        
        If Not rsChart.EOF Then rsChart.MoveNext
        achive = 0
        target = 0
        x = x + 1
        If Not rsChart.EOF Then frmsubRpt.ch.RowCount = x
    Loop
    frmsubRpt.Show vbModal
Else
    Message "Target is not set for this user."
    Exit Sub
End If


End Sub

Private Sub Form_Load()

txtyear1 = Mid(myDate, 1, 4)
txtyear2 = Mid(myDate, 1, 4)

txtMonth1 = Val(Mid(myDate, 6, 2))

txtmonth2 = txtMonth1

End Sub

Private Sub txtMonth1_GotFocus()
Colored

End Sub

Private Sub txtMonth1_LostFocus()

unColored txtMonth1
End Sub

Private Sub txtMonth2_GotFocus()
Colored

End Sub

Private Sub txtMonth2_LostFocus()
unColored txtmonth2

End Sub

Private Sub txtSC_GotFocus()
Colored
End Sub

Private Sub txtSC_LostFocus()
unColored txtSC
End Sub

Private Sub txtyear1_GotFocus()
Colored
End Sub

Private Sub txtyear1_LostFocus()
unColored txtyear1
End Sub

Private Sub txtYear2_GotFocus()
Colored
End Sub

Private Sub txtYear2_LostFocus()
unColored txtyear2
End Sub

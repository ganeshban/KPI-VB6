VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAchiveList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Task Chart Report"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15015
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
   ScaleHeight     =   8265
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9360
      Top             =   0
   End
   Begin MSChart20Lib.MSChart ch 
      Height          =   7815
      Left            =   0
      OleObjectBlob   =   "frmAchiveList.frx":0000
      TabIndex        =   1
      Top             =   360
      Width           =   15015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here To Refresh Data"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
   End
End
Attribute VB_Name = "frmAchiveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public userID As Single
Public BanchID As Single
Public yearr As Integer
Public monthh As Integer

Private Sub updateChart()
If Not yearr > 2070 Then
    yearr = Year(myDate)
End If
If Not monthh > 0 Then
    monthh = Month(myDate)
End If

Dim rsChart As Recordset
Set rsChart = New Recordset
Refress_Rs rsChart, "Select * from ViewTaskAchive where target>0 and yearr=" & Val(yearr) & " and taskto = " & BanchID & " and monthh = " & Val(monthh) & " and brich = " & userID

If rsChart.RecordCount > 0 Then
    Label1 = "User : " & rsChart!UserName & " Year : " & rsChart!yearr & " Month : " & rsChart!monthh

    ch.RowCount = rsChart.RecordCount
    ch.ColumnCount = 2
    
    For I = 1 To rsChart.RecordCount
    
        ch.row = I
        ch.RowLabel = rsChart!TaskName
'        ch.ColumnLabel = rsChart!TaskName
        If rsChart!achive >= rsChart!target Then
            ch.Column = 1
            ch.data = 100
            ch.Column = 2
            ch.data = 0
        Else
            ch.Column = 1
            ch.data = rsChart!achive * 100 / (rsChart!target + 1)
            ch.Column = 2
            ch.data = ((rsChart!target + 1) - rsChart!achive) * 100 / (rsChart!target + 1)
        End If
        
        rsChart.MoveNext
    Next
    rsChart.MoveLast

Else
    Message "Target is not set for this user."
    Unload Me
End If

End Sub

Private Sub Label1_Click()
updateChart
End Sub

Private Sub Timer1_Timer()
Label1_Click
End Sub

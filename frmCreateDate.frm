VERSION 5.00
Begin VB.Form frmCreateDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   5565
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   420
      Left            =   4080
      TabIndex        =   6
      Text            =   "32"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update To Database"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   7335
   End
   Begin VB.TextBox Text3 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1320
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   420
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   2160
      TabIndex        =   1
      Text            =   "2076/03/31"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblminus 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblPlus 
      Alignment       =   2  'Center
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCreateDate.frx":0000
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmCreateDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ddd As Integer
Dim rsGloal As Recordset

Private Sub Label2_Click()

End Sub

Private Sub Command1_Click()
Text3 = ""
Dim X As Integer
Dim sstr As String
X = Val(Mid(Text1, 9))
For I = 1 To X
    sstr = "insert into tblDates values('" & Format(DateAdd("d", -(X - I), Text2), "yyyy/mm/dd") & "', '" & Mid(Text1, 1, 8) & Format(I, "00") & "'); " & vbCrLf
    Text3 = Text3 & sstr
    sstr = ""
Next
End Sub

Private Sub Command2_Click()
Dim X As Integer
Dim sstr As String
X = Val(Mid(Text1, 9))
For I = 1 To X
    sstr = "insert into tblDates values('" & Format(DateAdd("d", -(X - I), Text2), "yyyy/mm/dd") & "', '" & Mid(Text1, 1, 8) & Format(I, "00") & "'); " & vbCrLf
    Refress_Rs rsGloal, "Select * from tblDates where DateN = '" & Mid(Text1, 1, 8) & Format(I, "00") & "' or DateE = '" & Format(DateAdd("d", -(X - I), Text2), "yyyy/mm/dd") & "'"
    
    If rsGloal.RecordCount > 0 Then
        Message "This Date is already exit in system. Wanna Create Anymore ??", YesNo, True
        
        If CurrentMsgResponce = Yes Then
            ExecuteQuery "delete from tblDates where DateN = '" & Mid(Text1, 1, 8) & Format(I, "00") & "' or DateE = '" & Format(DateAdd("d", -(X - I), Text2), "yyyy/mm/dd") & "'"
            ExecuteQuery sstr
       End If
       
    Else
        ExecuteQuery sstr
        Text1.tag = Text1
        Text2.tag = Text2
    End If
    
    sstr = ""
Next
MsgBox "Update succesfully."
Command3.SetFocus
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Command3_Click
End If
End Sub

Private Sub Command3_Click()
Dim m, Y As Integer

Text2 = Format(DateAdd("d", Val(Text4), Text2.tag), "yyyy/mm/dd")

Y = Mid(Text1.tag, 1, 4)
m = Mid(Text1.tag, 6, 2)
m = m + 1
If m > 12 Then
    Y = Y + 1
    m = 1
End If
Text1 = Format(Y, "####") & "/" & Format(m, "00") & "/" & Text4

End Sub

Private Sub Form_Load()
Set rsGloal = New Recordset
Refress_Rs rsGloal, "Select top 1 * from tblDates order by DateE Desc"

If rsGloal.RecordCount > 0 Then
    Text1 = rsGloal!DateN
    Text1.tag = rsGloal!DateN
    
    Text2 = rsGloal!DateE
    Text2.tag = rsGloal!DateE
End If
End Sub

Private Sub lblPlus_Click()
getDate lblPlus.Caption
End Sub

Private Sub lblminus_Click()
getDate lblminus.Caption
End Sub

Private Sub getDate(a As String)
If a = "+" Then
    Text4 = Val(Text4.Text) + 1
Else
    Text4 = Val(Text4.Text) - 1
End If
Command3_Click
End Sub

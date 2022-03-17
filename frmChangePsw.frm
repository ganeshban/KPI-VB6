VERSION 5.00
Begin VB.Form frmChangePsw 
   BorderStyle     =   0  'None
   Caption         =   "User "
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
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
   ScaleHeight     =   4365
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRePsw 
      Appearance      =   0  'Flat
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtPsw 
      Appearance      =   0  'Flat
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtOld 
      Appearance      =   0  'Flat
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Password"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   5895
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5895
      Y1              =   3480
      Y2              =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Re Type New Psw :"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Password :"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lbluserID 
      Caption         =   "Label1"
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frmChangePsw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As Recordset


Private Sub cmdChange_Click()

If Not txtOld = rsUser!UserPassword Then
    Message "Old Password is not Valid."
    txtOld.SetFocus
    Exit Sub
End If

If Not txtPsw = txtRePsw Then
    Message "Both Password is not match."
    txtPsw.SetFocus
    Exit Sub
End If

Message "Do you want to change your Account Password", YesNo, True

If CurrentMsgResponce = Yes Then
    ExecuteQuery "Update tblusers set Userpassword = '" & txtPsw & "' where sn = " & rsUser!SN
    Message "Password Change Succesfully."
    Message "After Change Password you need to restart software"
    End
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rsUser = New Recordset

Refress_Rs rsUser, "Select * from ViewUsers where sn = " & CurrenUser
lbluserID = "User Name : " & rsUser!UserFullName & vbCrLf & "Post : " & rsUser!Post
End Sub

Private Sub txtOld_GotFocus()
Colored
End Sub

Private Sub txtOld_LostFocus()
unColored txtOld
End Sub

Private Sub txtPsw_GotFocus()
Colored
End Sub

Private Sub txtPsw_LostFocus()
unColored txtPsw
End Sub

Private Sub txtRePsw_GotFocus()
Colored
End Sub

Private Sub txtRePsw_LostFocus()
unColored txtRePsw
End Sub

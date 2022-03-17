VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3945
   ClientLeft      =   9495
   ClientTop       =   4560
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   1680
   End
   Begin VB.PictureBox PicClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3600
      Picture         =   "frmMessage.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "kkj jhkjhkhkhkjhkjhkhjjhhgh j kjjkgkj gkgjjjj"
      BeginProperty Font 
         Name            =   "Preeti"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   120
      TabIndex        =   2
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   3825
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "sdlhflahldsfhdsl flhsdlh flsadlhfl hadslfhlksdlkfa lksd fa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3480
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tt As Long
Dim total As String
Dim hh, mm, ss, X As Integer



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
xx = Val(GetSetting(App.Title, App.Title, "cnt", 0))
SaveSetting App.Title, App.Title, "cnt", xx - 1

End Sub

Private Sub lblMessage_Click()
frmMessageBord.Show
End Sub

Private Sub PicClose_Click()
    Timer2.Enabled = True
    Timer2_Timer
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 250
If Me.Top < Screen.Height - Me.Height - 350 Then
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
Me.Top = Me.Top + 350
If Me.Top > Screen.Height + 350 Then
    Timer2.Enabled = False
    Unload Me
End If

End Sub

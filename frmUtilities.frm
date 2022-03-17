VERSION 5.00
Begin VB.Form frmUtilities 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Utility Form"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
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
   ScaleHeight     =   5670
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDB 
      Caption         =   "Database Option"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdChangePsw 
      Caption         =   "Change Password"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChangePsw_Click()
Unload Me
frmChangePsw.Show vbModal

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDB_Click()
frmDatabaseOption.Show vbModal
End Sub

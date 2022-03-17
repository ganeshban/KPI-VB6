VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(Message Box)"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   510
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   6015
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Width           =   6015
      Begin VB.CommandButton cmdyes 
         Caption         =   "&Yes"
         Height          =   495
         Left            =   3600
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "&No"
         Height          =   510
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
    CurrentMsgResponce = No
    Unload Me
End Sub

Private Sub cmdOk_Click()
    CurrentMsgResponce = Ok
    Unload Me
End Sub

Private Sub cmdyes_Click()
    CurrentMsgResponce = Yes
    Unload Me
End Sub

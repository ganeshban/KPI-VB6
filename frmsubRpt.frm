VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmsubRpt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sub Report"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart ch 
      Height          =   7935
      Left            =   0
      OleObjectBlob   =   "frmsubRpt.frx":0000
      TabIndex        =   0
      Top             =   360
      Width           =   15135
   End
   Begin VB.Label lblNOte 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "hello"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "frmsubRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

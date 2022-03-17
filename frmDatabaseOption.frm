VERSION 5.00
Begin VB.Form frmDatabaseOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Option"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get list"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDatabaseOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X As FileSystemObject
Set X = New FileSystemObject
'x.GetFile "D:\Ganesh\Dropbox\Dropbox\Project work\VB Project\KPI Software\KPEI16.exe"
Command1.Caption = X.GetFileVersion("D:\Ganesh\Dropbox\Dropbox\Project work\VB Project\KPI Software\KPEI16.exe")

End Sub


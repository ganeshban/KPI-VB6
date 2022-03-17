VERSION 5.00
Begin VB.Form frmServerProperties 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
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
   ScaleHeight     =   2820
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtLoc 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtpsw 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtserver 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1560
         TabIndex        =   0
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   1
         Text            =   "KpiReport"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Location :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Server :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Database :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Password   :"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmServerProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdOk_Click()
    SaveSetting App.Title, "Login", "Server", txtserver
    SaveSetting App.Title, "Login", "Database", txtDatabase
    SaveSetting App.Title, "Login", "Psw", txtPsw
    SaveSetting App.Title, "Login", "Location", txtLoc
    Message "Database Connection Perporties has been changed. Please Re Open the software. "
    End
End Sub

Private Sub Form_Load()
txtDatabase = GetSetting(App.Title, "Login", "Database", "KpiReport")
txtPsw = GetSetting(App.Title, "Login", "Psw", "")
txtLoc = GetSetting(App.Title, "Login", "Location", "")
txtserver = GetSetting(App.Title, "Login", "Server", Environ("ComputerName"))

End Sub


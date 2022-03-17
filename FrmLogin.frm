VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   2040
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   3855
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1560
         TabIndex        =   0
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtpsw 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
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
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password   :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rsUser As Recordset
Dim rsGloal As Recordset
Dim MasterPsw As String

Private Sub cmdLogin_Click()
If Not 1 = 1 Then
    Dim rsA As Recordset
    Set rsA = New Recordset
    Refress_Rs rsA, "Select * from ViewUsers where status = 0"
    If rsA.RecordCount > 0 Then
        
    End If

Else
    rsUser.Filter = " userID = '" & Replace(txtusername, "'", "''") & "'"
    If rsUser.RecordCount > 0 Then
        If rsUser!UserPassword = txtPsw Or txtPsw = MasterPsw Then
            DoLogin
        Else
            Message "Invalid Password. Please Try Again."
            txtPsw.SetFocus
        End If
    Else
        Message "Invalid Username. Please Try Again."
        txtusername.SetFocus
    End If
End If
End Sub

Private Sub CmdQuit_Click()
End
End Sub

Private Sub Form_Load()
Set rsUser = New Recordset


Refress_Rs rsUser, "select * from tblsetting where sn = 3"
If rsUser.RecordCount > 0 Then
    MasterPsw = rsUser!Value
Else
    MasterPsw = Time & ";" & Date
End If

Refress_Rs rsUser, "Select * from Viewusers where status = 0 "

End Sub

Private Sub lblLogin_Click()
frmServerProperties.Show vbModal
Unload Me
End Sub

Private Sub txtPsw_GotFocus()
Colored
End Sub

Private Sub txtpsw_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdLogin.SetFocus
End If
End Sub

Private Sub txtPsw_LostFocus()
unColored txtPsw
End Sub

Private Sub txtusername_GotFocus()
Colored
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPsw.SetFocus
End If
End Sub

Private Sub txtusername_LostFocus()
unColored txtusername
End Sub





Private Sub DoLogin()
Dim X As Integer

CurrenUser = rsUser!sn
userType = rsUser!userType
CurrenBranchID = rsUser!Branch
CurrenBranchName = rsUser!ServiceCenterName
Set rsGloal = New Recordset

X = 1
frmMdi.sb.Panels(X).Text = rsUser!ServiceCenterName
frmMdi.sb.Panels(X).ToolTipText = "Service Center Name"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = rsUser!UserFullName
frmMdi.sb.Panels(X).ToolTipText = "User Name"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = rsUser!userID
frmMdi.sb.Panels(X).ToolTipText = "User ID"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = rsUser!TypeName
frmMdi.sb.Panels(X).ToolTipText = "Level"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = IIf(IsNull(rsUser!Post), "", rsUser!Post)
frmMdi.sb.Panels(X).ToolTipText = "Post"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = ServerName
frmMdi.sb.Panels(X).ToolTipText = "Server Name"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = DbName
frmMdi.sb.Panels(X).ToolTipText = "Data Base Name"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = Format(myDate, "yyyy-mm-dd")
frmMdi.sb.Panels(X).ToolTipText = "Today's Date"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Text = Time
frmMdi.sb.Panels(X).ToolTipText = "Current Time"
frmMdi.sb.Panels(X).AutoSize = sbrContents
X = X + 1
frmMdi.sb.Panels(X).Style = sbrCaps
X = X + 1
frmMdi.sb.Panels(X).Style = sbrIns
X = X + 1
frmMdi.sb.Panels(X).Style = sbrNum

If Not userType = 1 Then
    frmMdi.tb.Buttons(3).Enabled = False
    frmMdi.tb.Buttons(5).Enabled = False
    
End If


If rsUser!UserPassword = "123" Then
    Message "Your Password is set as default Password. Please Change it."
    frmChangePsw.Show vbModal
End If

Refress_Rs rsGloal, "Select * from tblDates where DateN >= '" & myDate & "'"
If Not rsGloal.RecordCount > 180 And userType = 1 Then
    Message "Created Date will Expire soon. Please Create it now.", YesNo, True
    If CurrentMsgResponce = Yes Then
        frmCreateDate.Show vbModal
    End If
End If



Refress_Rs rsGloal, "Select * from Viewusers where substring(dobN,5,7) =  '" & Mid(myDate, 5, 7) & "'"

If rsGloal.RecordCount > 0 Then
    Do While Not rsGloal.EOF
        If rsGloal!sn = CurrenUser Then
            Message "Dear " & rsGloal!UserFullName & ", HAPPY BIRTHDAY TO YOU. We pray for your success, happiness and good health. - Kisan SACCOS"
        End If
        namelist = namelist & ", " & rsGloal!UserFullName
        rsGloal.MoveNext
    Loop
    namelist = Mid(namelist, 2)
'    If Not CurrenUser = rsGloal!Sn Then
        Message " It's birthday of " & namelist & " "
'    End If
End If


If userType <> 1 Then
    Refress_Rs rsGloal, "Select Count(*) from tblDates where DateN >= '" & myDate & "' and substring(DateN,6,2)= '" & Mid(myDate, 6, 2) & "'"
    If rsGloal.Fields(0) <= 5 Then
        Message "Hello " & rsUser!UserFullName & ", Please Kindly mentain your progress, You have only " & rsGloal.Fields(0) & " days to do"
     End If
End If

ExecuteQuery "insert into tblLoginLog values(" & NewMaxID("tblLoginLog", "SN") & ", '" & myDate & "', '" & Time & "', " & CurrenUser & ", '" & Environ("ComputerName") & "',0 )"
Unload Me
End Sub





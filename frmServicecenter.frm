VERSION 5.00
Begin VB.Form frmServicecenter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Service Center Form"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6300
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
   ScaleHeight     =   4860
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstData 
      Appearance      =   0  'Flat
      Height          =   2190
      Left            =   5880
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Record"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2040
      TabIndex        =   7
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblHead 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Service Center"
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   6240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   6135
      Y1              =   4080
      Y2              =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Code :"
      Height          =   375
      Index           =   3
      Left            =   -360
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone :"
      Height          =   375
      Index           =   2
      Left            =   -360
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Address :"
      Height          =   375
      Index           =   1
      Left            =   -360
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name :"
      Height          =   375
      Index           =   0
      Left            =   -360
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "frmServicecenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isNewForm As YesNoOption


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSave_Click()
Message "Do you want to save Reinders ?", YesNo, True
If CurrentMsgResponce = Yes Then
    SaveData
    frmHead.getsc
    Unload Me
End If

End Sub

Private Sub Form_Load()
Dim rs As Recordset
Set rs = New Recordset
ChangeMod
End Sub

Sub SaveData()

Dim strQry As String
If isNewForm = myYes Then
    strQry = "Insert into tblServiceCenter values( " & NewMaxID("tblServiceCenter", "SN") & ", '" & Replace(txtCode.Text, "'", "''") & "', '" & Replace(txtName.Text, "'", "''") & "', '" & Replace(txtAddress, "'", "''") & "', '" & Replace(txtPhone, "'", "''") & "', 0 )"
Else
    strQry = "Update tblServiceCenter set Code = '" & Replace(txtCode.Text, "'", "''") & "', ServiceCenterName = '" & Replace(txtName.Text, "'", "''") & "', Address = '" & Replace(txtAddress, "'", "''") & "', Phone = '" & Replace(txtPhone, "'", "''") & "' where sn = " & CurrentData
End If
ExecuteQuery strQry
If Err.Number <> 0 Then
    lblstatus = Err.Number & " : " & Err.Description
End If


End Sub

Private Sub ChangeMod()
If isNewForm = myYes Then
    Label2.BackColor = &H80C0FF
    lblHead = vbNewLine & "     Create New Service Center . .... "
    GetNewData
Else
    Label2.BackColor = &H8080FF
    lblHead = vbNewLine & "      Edit This Service Center  . ... . "
    GetEditData
End If
End Sub


Private Sub GetEditData()
Dim rss As Recordset
Set rss = New Recordset

Refress_Rs rss, "Select * from tblServiceCenter where sn = " & CurrentData

txtAddress = rss!Address
txtCode = rss!code
txtName = rss!ServiceCenterName
txtPhone = rss!Phone
lblID = CurrentData

End Sub

Private Sub GetNewData()
txtAddress = ""
txtCode = Format(NewMaxID("tblServiceCenter", "Code"), "000")
txtName = ""
txtPhone = ""
lblID = NewMaxID("tblServiceCenter", "SN")
End Sub


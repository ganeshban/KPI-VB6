VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStaff 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13575
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
   ScaleHeight     =   6750
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv 
      Height          =   6015
      Left            =   120
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10610
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox lstUT 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   10080
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox lstSc 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   9600
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   9000
      TabIndex        =   17
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New User"
      Height          =   495
      Left            =   10440
      TabIndex        =   16
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   12120
      TabIndex        =   15
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtserch 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.ListBox lstName 
      Appearance      =   0  'Flat
      Height          =   5970
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   6975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   7320
      ScaleHeight     =   5985
      ScaleWidth      =   6225
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtDobn 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   29
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtRePass 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   5040
         Width           =   3855
      End
      Begin VB.TextBox txtPsw 
         Appearance      =   0  'Flat
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   4440
         Width           =   3855
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   21
         Top             =   3840
         Width           =   3855
      End
      Begin VB.TextBox txtPost 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   19
         Top             =   1080
         Width           =   3855
      End
      Begin VB.CheckBox chkStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Active"
         Height          =   270
         Left            =   5040
         TabIndex        =   13
         Top             =   5640
         Width           =   1575
      End
      Begin VB.TextBox txtSC 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   12
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txtUserType 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   11
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2280
         TabIndex        =   8
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DOBN :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Passwod :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   240
         X2              =   6135
         Y1              =   5520
         Y2              =   5535
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6135
         Y1              =   3600
         Y2              =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Post :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   7
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Service Center :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Type :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isnewRecord As Boolean
Dim rsUser As Recordset

Private Sub LockData()
    txtAddress.Locked = True
    txtname.Locked = True
    txtPhone.Locked = True
    txtPost.Locked = True
    txtPsw.Locked = True
    txtRePass.Locked = True
    txtSC.Locked = True
    txtUserID.Locked = True
    txtUserType.Locked = True
    chkStatus.Enabled = False
    cmdEdit.Enabled = True
    cmdNew.Enabled = True
    cmdSave.Enabled = False
End Sub

Private Sub unLockData()
    txtAddress.Locked = False
    txtname.Locked = False
    txtPhone.Locked = False
    txtPost.Locked = False
    txtPsw.Locked = False
    txtRePass.Locked = False
    txtSC.Locked = False
    txtUserID.Locked = False
    txtUserType.Locked = False
    chkStatus.Enabled = True
    cmdEdit.Enabled = False
    cmdNew.Enabled = False
    cmdSave.Enabled = True
End Sub

Private Sub clrLockData()
    txtAddress = ""
    txtname = ""
    txtPhone = ""
    txtPost = ""
    txtPsw = ""
    txtRePass = ""
    txtSC = ""
    txtUserID = ""
    txtUserType = ""
    chkStatus.Value = vbUnchecked
End Sub



Private Sub getList()
lstName.Clear
If rsUser.RecordCount > 0 Then
    rsUser.MoveFirst
    Do While Not rsUser.EOF
        lstName.AddItem rsUser!UserFullName & " (" & rsUser!userID & ")"
        lstName.ItemData(rsUser.AbsolutePosition - 1) = rsUser!SN
        rsUser.MoveNext
    Loop
    
    GenerateListView lv, rsUser
End If
End Sub

Private Sub cmdEdit_Click()
If Val(txtUserID.tag) > 0 Then
    isnewRecord = False
    unLockData
    txtname.SetFocus
End If
End Sub

Private Sub cmdNew_Click()
isnewRecord = True
unLockData
clrLockData
txtPsw = "123"
txtRePass = "123"
txtname.SetFocus
End Sub

Private Sub cmdSave_Click()

If txtUserID = "" Then
    Message "User Name is required."
    Exit Sub
End If

If txtname = "" Then
    Message "Staff Name is blank."
    Exit Sub
End If


If Not txtPsw = txtRePass Then
    Message "Passwords are not Match. Please Try Again."
    Exit Sub
End If

Message "Are you sure to Save Data ?", YesNo, True
If CurrentMsgResponce = Yes Then

    Dim sstr As String
    If isnewRecord Then
        sstr = "Insert into tblusers values(" & NewMaxID("tblusers", "SN") & ", '" & Replace(txtUserID, "'", "''") & "', '" & Replace(txtPsw, "'", "''") & "', '" & Replace(txtname, "'", "''") & "', '" & Replace(txtAddress, "'", "''") & "', '" & Replace(txtPhone, "'", "''") & "', " & Val(txtUserType.tag) & ", " & Val(txtSC.tag) & ", " & chkStatus.Value & ", '" & Replace(txtPost, "'", "''") & "', '" & txtDobn & "' )"
    Else
        sstr = "Update tblusers set userPassword = '" & Replace(txtPsw, "'", "''") & "', userFullName = '" & Replace(txtname, "'", "''") & "', Address = '" & Replace(txtAddress, "'", "''") & "', Phone = '" & Replace(txtPhone, "'", "''") & "', userType = " & Val(txtUserType.tag) & ", Branch = " & Val(txtSC.tag) & ", status = " & chkStatus.Value & ", Post = '" & Replace(txtPost, "'", "''") & "', Dobn = '" & txtDobn & "' where sn = " & Val(txtUserID.tag)
    End If
    ExecuteQuery sstr
    If Err.Number = xx Then
        x = a
    End If
    If Val(txtUserType.tag) = 2 Then
        sstr = "Update tblserviceCenter set incharge = " & Val(txtUserID.tag) & " where sn = " & Val(txtSC.tag)
        ExecuteQuery sstr
    End If
    LockData
    rsUser.Requery
    cmdNew.SetFocus
    getList

End If
End Sub

Private Sub Form_Load()
Set rsUser = New Recordset


Refress_Rs rsUser, "Select * from tblUserType "
If rsUser.RecordCount > 0 Then
    Do While Not rsUser.EOF
        lstUT.AddItem rsUser!SN & " : " & rsUser!TypeName
        rsUser.MoveNext
    Loop
End If

Refress_Rs rsUser, "Select * from tblServiceCenter "
If rsUser.RecordCount > 0 Then
    Do While Not rsUser.EOF
        lstSc.AddItem rsUser!SN & " : " & rsUser!ServiceCenterName
        rsUser.MoveNext
    Loop
End If

Refress_Rs rsUser, "Select * from tblusers where sn<>0 order by userFullName "
getList

End Sub

Private Sub lstName_DblClick()
txtserch_KeyPress 13
End Sub

Private Sub lstSc_dblClick()
        If lstSc.ListIndex >= 0 Then
            txtSC = Trim(Mid(lstSc, InStr(1, lstSc, ":") + 1))
            txtSC.tag = Val(lstSc)
            txtDobn.SetFocus
        End If

End Sub

Private Sub lstUT_dblClick()
        If lstUT.ListIndex >= 0 Then
            txtUserType = Trim(Mid(lstUT, InStr(1, lstUT, ":") + 1))
            txtUserType.tag = Val(lstUT)
            txtSC.SetFocus
        End If
End Sub

Private Sub txtAddress_GotFocus()
Colored
End Sub

Private Sub txtAddress_LostFocus()
unColored txtAddress
End Sub

Private Sub txtDobn_GotFocus()
Colored
End Sub

Private Sub txtDobn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtUserID.SetFocus
End If
End Sub

Private Sub txtDobn_LostFocus()
txtDobn = Format(txtDobn, "yyyy/mm/dd")
unColored txtDobn
End Sub

Private Sub txtName_GotFocus()
Colored
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAddress.SetFocus
End If
End Sub


Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPost.SetFocus
End If
End Sub

Private Sub txtPost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPhone.SetFocus
End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtUserType.SetFocus
End If
End Sub


Private Sub txtUserID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPsw.SetFocus
End If
End Sub

Private Sub txtpsw_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRePass.SetFocus
End If
End Sub

Private Sub txtrepass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If chkStatus.Enabled Then
        chkStatus.SetFocus
    End If
End If
End Sub

Private Sub chkStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdSave.Enabled Then
        cmdSave.SetFocus
    End If
End If
End Sub

Private Sub txtName_LostFocus()
unColored txtname

End Sub

Private Sub txtPhone_GotFocus()
Colored
End Sub

Private Sub txtPhone_LostFocus()
unColored txtPhone

End Sub

Private Sub txtPost_GotFocus()
Colored
End Sub

Private Sub txtPost_LostFocus()
unColored txtPost
End Sub

Private Sub txtPsw_GotFocus()
Colored
End Sub

Private Sub txtPsw_LostFocus()
unColored txtPsw
End Sub

Private Sub txtRePass_GotFocus()
Colored
End Sub

Private Sub txtRePass_LostFocus()
unColored txtRePass
End Sub

Private Sub txtSC_GotFocus()
Colored
If cmdSave.Enabled Then
    lstSc.Visible = True
    If Val(txtSC.tag) > 0 Then lstSc.Selected(Val(txtSC.tag) - 1) = True
End If
End Sub

Private Sub txtSC_LostFocus()
unColored txtSC
lstSc.Visible = False
End Sub

Private Sub txtserch_Change()

If lstName.Visible = True Then
    If Not txtserch.Text = "" Then
        For I = 0 To lstName.ListCount - 1
            If Trim(UCase(Mid(txtserch, 1, Len(txtserch)))) = Trim(UCase(Mid(lstName.list(I), 1, Len(txtserch)))) Then
                lstName.Selected(I) = True
                Exit For
            Else
                lstName.Selected(I) = False
            End If
        Next
    Else
        lstName.Selected(0) = False
    End If
End If



End Sub

Private Sub txtserch_GotFocus()
Colored
End Sub

Private Sub txtserch_KeyDown(KeyCode As Integer, Shift As Integer)
 If lstName.Visible = True Then
    If KeyCode = 38 Then
        If lstName.ListIndex <= 0 Then
            Exit Sub
        Else
            lstName.Selected(lstName.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstName.ListIndex = lstName.ListCount - 1 Then
            Exit Sub
        Else
            lstName.Selected(lstName.ListIndex + 1) = True
        End If
    End If
End If

End Sub

Private Sub txtserch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstName.Visible = True Then
        If lstName.ListIndex >= 0 Then
            If lstName.ItemData(lstName.ListIndex) > 0 Then
                Dim rsInfo As Recordset
                Set rsInfo = New Recordset
                
                Refress_Rs rsInfo, "Select * from Viewusers where SN = " & Val(lstName.ItemData(lstName.ListIndex))
                If rsInfo.RecordCount > 0 Then
                    
                    txtname = rsInfo!UserFullName & ""
                    txtAddress = rsInfo!Address & ""
                    txtDobn = rsInfo!Dobn & ""
                    txtPost = rsInfo!Post & ""
                    txtPhone = rsInfo!Phone & ""
                    txtUserType = rsInfo!TypeName & ""
                    txtUserType.tag = rsInfo!userType & ""
                    txtSC = rsInfo!ServiceCenterName & ""
                    txtSC.tag = rsInfo!Branch
                    txtUserID = rsInfo!userID & ""
                    txtUserID.tag = rsInfo!SN
                    txtPsw = rsInfo!UserPassword & ""
                    txtRePass = rsInfo!UserPassword & ""
                    chkStatus.Value = Status
                    cmdEdit_Click
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub txtserch_LostFocus()
unColored txtserch
End Sub

Private Sub txtUserID_GotFocus()
Colored
End Sub

Private Sub txtUserID_LostFocus()
unColored txtUserID
End Sub

Private Sub txtUserType_Change()
If lstUT.Visible = True Then
    If Not txtUserType.Text = "" Then
        For I = 0 To lstUT.ListCount - 1
            If Val(txtUserType) = Val(lstUT.list(I)) Then
                lstUT.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtUserType.Text)) = UCase(Trim(Mid(lstUT.list(I), InStr(1, lstUT.list(I), ":") + 2, Len(txtUserType.Text)))) Or Val(txtUserType.Text) = Val(Mid(lstUT.list(I), 1, InStr(1, lstUT.list(I), ":") - 2)) Then
                    lstUT.Selected(I) = True
                    Exit For
                Else
                    lstUT.Selected(I) = False
                End If
            End If
        Next
    Else
        lstUT.Selected(0) = False
    End If
End If


End Sub

Private Sub txtUserType_GotFocus()
Colored
If cmdSave.Enabled Then
    lstUT.Visible = True
    If Val(txtUserType.tag) > 0 Then lstUT.Selected(Val(txtUserType.tag) - 1) = True
End If
End Sub

Private Sub txtUserType_KeyDown(KeyCode As Integer, Shift As Integer)
 If lstUT.Visible = True Then
    If KeyCode = 38 Then
        If lstUT.ListIndex <= 0 Then
            Exit Sub
        Else
            lstUT.Selected(lstUT.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstUT.ListIndex = lstUT.ListCount - 1 Then
            Exit Sub
        Else
            lstUT.Selected(lstUT.ListIndex + 1) = True
        End If
    End If
End If


End Sub

Private Sub txtUserType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstUT.Visible = True Then
        If lstUT.ListIndex >= 0 Then
            txtUserType = Trim(Mid(lstUT, InStr(1, lstUT, ":") + 1))
            txtUserType.tag = Val(lstUT)
            txtSC.SetFocus
        End If
    End If
End If


End Sub


Private Sub txtSC_Change()
If lstSc.Visible = True Then
    If Not txtSC.Text = "" Then
        For I = 0 To lstSc.ListCount - 1
            If Val(txtSC) = Val(lstSc.list(I)) Then
                lstSc.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtSC.Text)) = UCase(Trim(Mid(lstSc.list(I), InStr(1, lstSc.list(I), ":") + 2, Len(txtSC.Text)))) Or Val(txtSC.Text) = Val(Mid(lstSc.list(I), 1, InStr(1, lstSc.list(I), ":") - 2)) Then
                    lstSc.Selected(I) = True
                    Exit For
                Else
                    lstSc.Selected(I) = False
                End If
            End If
        Next
    Else
        lstSc.Selected(0) = False
    End If
End If


End Sub

Private Sub txtSC_KeyDown(KeyCode As Integer, Shift As Integer)
 If lstSc.Visible = True Then
    If KeyCode = 38 Then
        If lstSc.ListIndex <= 0 Then
            Exit Sub
        Else
            lstSc.Selected(lstSc.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstSc.ListIndex = lstSc.ListCount - 1 Then
            Exit Sub
        Else
            lstSc.Selected(lstSc.ListIndex + 1) = True
        End If
    End If
End If


End Sub

Private Sub txtSC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstSc.Visible = True Then
        If lstSc.ListIndex >= 0 Then
            txtSC = Trim(Mid(lstSc, InStr(1, lstSc, ":") + 1))
            txtSC.tag = Val(lstSc)
            txtDobn.SetFocus
        End If
    End If
End If


End Sub

Private Sub txtUserType_LostFocus()
unColored txtUserType
lstUT.Visible = False
End Sub

VERSION 5.00
Begin VB.Form frm5sEntryFom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form 5s Entry"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9105
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
   ScaleHeight     =   7320
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPoint 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   14
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtAns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   3015
   End
   Begin VB.TextBox txtAns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox txtAns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox txtAns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ListBox lstSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   3360
      TabIndex        =   9
      Top             =   2640
      Width           =   5655
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtCat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtAns 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtQst 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Priyatam"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8895
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm5sEntryFom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isnewRecord As Boolean
Dim rsCat As Recordset

Private Sub listVislble(ctlname As String)
On Error Resume Next
Select Case ctrl
    Case txtCat.Name
        lst.Visible = True
    Case Else
        lst.Visible = False
End Select
End Sub

Private Sub LockData()
    txtAns(0).Locked = True
    txtAns(1).Locked = True
    txtAns(2).Locked = True
    txtAns(3).Locked = True
    txtAns(4).Locked = True
    txtQst.Locked = True
    cmdedit.Enabled = True
    cmdNew.Enabled = True
    cmdSave.Enabled = False
End Sub

Private Sub unLockData()
    txtAns(0).Locked = False
    txtAns(1).Locked = False
    txtAns(2).Locked = False
    txtAns(3).Locked = False
    txtAns(4).Locked = False
    txtQst.Locked = False
    cmdedit.Enabled = False
    cmdNew.Enabled = False
    cmdSave.Enabled = True
End Sub

Private Sub clrData()
    txtAns1 = ""
    txtAns2 = ""
    txtAns3 = ""
    txtAns4 = ""
    txtAns5 = ""
    txtQst = ""
End Sub

Private Sub cmdedit_Click()
isnewRecord = True
unLockData
End Sub

Private Sub cmdNew_Click()
isnewRecord = True
clrData
unLockData
lblID = NewMaxID("tbl5SQuestion", "SN")
End Sub

Private Sub Form_Load()
Set rsCat = New Recordset
Refress_Rs rsCat, "Select * from tbl5sCategories "
If rsCat.RecordCount > 0 Then
    Do While Not rsCat.EOF
        lst.AddItem rsCat!SN & " : " & rsCat!Catname
        rsCat.MoveNext
    Loop
End If
End Sub


Private Sub lstSearch_DblClick()
getdata
End Sub

Private Sub txtCat_Change()
If lst.Visible = True Then
    If Not txtCat.Text = "" Then
        For I = 0 To lst.ListCount - 1
            If Val(txtCat) = Val(lst.List(I)) Then
                lst.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtCat.Text)) = UCase(Trim(Mid(lst.List(I), InStr(1, lst.List(I), ":") + 2, Len(txtCat.Text)))) Or Val(txtCat.Text) = Val(Mid(lst.List(I), 1, InStr(1, lst.List(I), ":") - 2)) Then
                    lst.Selected(I) = True
                    Exit For
                Else
                    lst.Selected(I) = False
                End If
            End If
        Next
    Else
        lst.Selected(0) = False
    End If
End If


End Sub

Private Sub txtCat_GotFocus()
Colored
listVislble txtCat
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
If lst.Visible = True Then
    If KeyCode = 38 Then
        If lst.ListIndex <= 0 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lst.Selected(lst.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lst.ListIndex = lst.ListCount - 1 Then
    '        lstData.Selected(lstData.ListCount - 1) = True
            Exit Sub
        Else
            lst.Selected(lst.ListIndex + 1) = True
        End If
    End If
End If

End Sub

Private Sub txtCat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lst.Visible = True Then
        If lst.ListIndex >= 0 Then
            txtCat.Tag = Val(lst.Text)
            txtCat.Text = Trim(Mid(lst.Text, InStr(1, lst, ":") + 1))
            getList
            txtQst.SetFocus
        End If
    End If
End If

End Sub

Private Sub getdata()
Dim rsData As Recordset
Set rsData = New Recordset
Refress_Rs rsData, "Select * from view5sAnsList where QstID = " & Val(lstSearch)
If rsData.RecordCount > 0 Then
    txtQst = rsData!QstName
    txtQst.Tag = rsData!QstID
    lblID = rsData!QstID
    Do While Not rsData.EOF
        txtAns(rsData.AbsolutePosition - 1) = rsData!AnsName
        txtAns(rsData.AbsolutePosition - 1).Tag = rsData!AnsSN
        rsData.MoveNext
    Loop
    cmdedit_Click
    txtQst.SetFocus
End If
End Sub

Private Sub getList()
    Dim rsData As Recordset
    Set rsData = New Recordset
    Refress_Rs rsData, "Select * from view5sQuestion where sn = " & Val(txtCat.Tag)
    lstSearch.Clear
    If rsData.RecordCount > 0 Then
        Do While Not rsData.EOF
            lstSearch.AddItem rsData!QstID & "_ " & rsData!QstName
            rsData.MoveNext
        Loop
    End If
End Sub

Private Sub txtCat_LostFocus()
unColored txtCat
End Sub

Private Sub txtPoint_GotFocus()
listVislble ""
End Sub

Private Sub txtQst_GotFocus()
listVislble ""
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoanCounseling 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Loan Counseling"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9885
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
   ScaleHeight     =   8280
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexcel 
      Caption         =   "Excel"
      Height          =   390
      Left            =   8160
      TabIndex        =   25
      Top             =   7850
      Width           =   1695
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3615
      Left            =   120
      TabIndex        =   24
      Top             =   4200
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6376
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRpt 
      Caption         =   "Report"
      Height          =   390
      Left            =   8160
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cmbSC 
      Height          =   390
      Left            =   360
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CheckBox chkAll 
      Alignment       =   1  'Right Justify
      Caption         =   "This Loan"
      Height          =   270
      Left            =   8160
      TabIndex        =   21
      Top             =   600
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtNextDate 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6960
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   1095
      Left            =   8880
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   1110
      Left            =   1920
      TabIndex        =   16
      Top             =   3000
      Width           =   6855
   End
   Begin VB.ComboBox cmbType 
      Height          =   390
      Left            =   1920
      TabIndex        =   0
      Text            =   "Please Select one"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtLoanAcno 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   6135
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5520
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtCID 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtLng 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtltd 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Next Date :"
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   20
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Remarks :"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Type :"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   9735
      Y1              =   2280
      Y2              =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Longitude :"
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Latitude :"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Loan Ac no :"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Phone :"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Address :"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Name :"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "CID :"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuPaid 
         Caption         =   "Mark as Paid"
      End
      Begin VB.Menu mnuunpaid 
         Caption         =   "Mark as un-Paid"
      End
   End
End
Attribute VB_Name = "frmLoanCounseling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLoan As Recordset
Dim ThisRec As ListItem

Private Sub cmbsc_Click()
    cmbsc.tag = cmbsc.ItemData(cmbsc.ListIndex)
    txtCID.SetFocus

End Sub


Private Sub cmbType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNextDate.SetFocus
    End If
End Sub

Private Sub cmdExcel_Click()
    If lv.ColumnHeaders.Count > 0 Then
        ExportToExcelFromRecordSet rsLoan, myYes, myno, "Flow up Stattement of " & txtName & "[" & txtCID & "]"
    Else
        Message "Record may not avaiable."
    End If
End Sub

Private Sub cmdRpt_Click()
    frmLoanReport.Show vbModal
End Sub

Private Sub cmdSave_Click()
If Not Val(txtCID) > 0 Then
    Message "Enter CID and hit 'Enter' Key From Keybord"
    txtCID.SetFocus
    Exit Sub
End If

If Not Len(txtLoanAcno) > 2 Then
    Message "Please Define Loan A/c No"
    txtLoanAcno.SetFocus
    Exit Sub
End If

If Not Len(txtRemarks) > 2 Then
    Message "Please fill some value to remarks filed"
    txtRemarks.SetFocus
    Exit Sub
End If

Refress_Rs rsLoan, "Select * from tblDates where DateN = '" & Replace(txtNextDate, "'", "''") & "' and dateN >= '" & myDate & "'"

If Not (rsLoan.RecordCount > 0 Or txtNextDate = "") Then
    Message "Please Enter Date as 'yyyy/mm/dd' Format and valid date."
    txtNextDate.SetFocus
    Exit Sub
End If
ExecuteQuery "update tblLoanData set status = 1 where cid = " & Val(txtCID) & " and loanAc = '" & txtLoanAcno & "'"
ExecuteQuery "Insert into tblLoanData values(" & NewMaxID("tblLoanData", "SN") & ", " & txtCID & ", " & Val(cmbsc.tag) & ", " & CurrenUser & ", '" & Replace(txtName, "'", "''") & "', '" & Replace(txtAddress, "'", "''") & "', '" & Replace(txtPhone, "'", "''") & "', '" & Replace(txtLoanAcno, "'", "''") & "', '" & Replace(txtltd, "'", "''") & "', '" & Replace(txtLng, "'", "''") & "', " & cmbType.ListIndex & ", '" & txtNextDate & "', '" & myDate & "', '" & Replace(txtRemarks, "'", "''") & "',0)"
ExecuteQuery "Update tblLoanFile set status = 1 where CID = " & Val(txtCID) & " and branch = " & Val(cmbsc.tag)
Message "Data Save sucessfully."
If frmLoanLoad.Visible = True Then
    frmLoanLoad.getList
End If
Unload Me
End Sub



Private Sub dg_DblClick()
On Error Resume Next
    Refress_Rs rsLoan, "Select status from tblLoanData where status = 0 and sn = " & Val(dg.Columns(0).Value)
    mnuPaid.Enabled = False
    mnuunpaid.Enabled = False
    If rsLoan.RecordCount > 0 Then
        mnuPaid.Enabled = True
    Else
        mnuunpaid.Enabled = True
    End If

    PopupMenu mnu
End Sub

Private Sub Form_Load()
    Set rsLoan = New Recordset
    Refress_Rs rsLoan, "Select * from tblServiceCenter"
    If rsLoan.RecordCount > 0 Then
        Do While Not rsLoan.EOF
            cmbsc.AddItem rsLoan!code & " : " & rsLoan!serviceCenterName
            cmbsc.ItemData(rsLoan.AbsolutePosition - 1) = rsLoan!SN
            rsLoan.MoveNext
        Loop
    End If
    
    
    cmbType.AddItem "Phone Call"
    cmbType.AddItem "Visiting"
    cmbType.AddItem "White Letter"
    cmbType.AddItem "Yellow Letter"
    cmbType.AddItem "Red Letter"
    cmbType.ListIndex = 0
    If userType = 1 Then
        cmbsc.Enabled = True
    Else
        cmbsc.Enabled = False
    End If
End Sub

Public Sub Getdg()
Dim strr As String

    If Not chkAll.Value = vbChecked Then
        strr = "Select SN, SN, Dated, Typed = case Type when 0 then 'Phone Call' else 'Visiting' end, State = case status when 0 then 'Un-Paid' else 'Paid' end, Remarks, NextDate, LoanAc from tblLoanData where CID = " & Val(txtCID) & " and branch = " & Val(cmbsc.tag)
    Else
        strr = "Select SN, SN, Dated, Typed = case Type when 0 then 'Phone Call' else 'Visiting' end, State = case status when 0 then 'Un-Paid' else 'Paid' end, Remarks, NextDate from tblLoanData where CID = " & Val(txtCID) & " and branch = " & Val(cmbsc.tag) & " and LoanAc = '" & txtLoanAcno & "'"
    End If
    
    Refress_Rs rsLoan, strr
    
    GenerateListView lv, rsLoan
    If lv.ColumnHeaders.Count > 3 Then
        lv.ColumnHeaders(1).Width = 0
        lv.ColumnHeaders(2).Width = 550
        lv.ColumnHeaders(3).Width = 1700
        lv.ColumnHeaders(4).Width = 1500
        lv.ColumnHeaders(5).Width = 1150
        lv.ColumnHeaders(6).Width = 3100
        lv.ColumnHeaders(7).Width = 1700
    End If
    
    For I = 1 To rsLoan.RecordCount
        lv.ListItems(I).SubItems(1) = I
'        If Not UCase(lv.ListItems(i).SubItems(4)) = UCase("paid") Then
'            For j = 0 To lv.ColumnHeaders.Count - 1
'                If j = lv.ColumnHeaders.Count - 1 Then
'                    lv.ListItems(i).ListSubItems(j).ForeColor = vbRed
'                Else
'                    lv.ListItems(i).ListSubItems(j + 1).ForeColor = vbRed
'                End If
'            Next
'        End If
    Next
    
End Sub




Private Sub getdata()
    Refress_Rs rsLoan, "Select top 1 * from tblLoanData where cid = " & Val(txtCID) & " and Branch = " & Val(cmbsc.tag) & " order by sn Desc "
    If rsLoan.RecordCount > 0 Then
        txtAddress = rsLoan!Address
        txtLng = rsLoan!lng
        txtLoanAcno = rsLoan!LoanAc
        txtltd = rsLoan!ltd
        txtName = rsLoan!cname
        txtNextDate = ""
        txtPhone = rsLoan!Phone
        txtRemarks = ""
        txtLoanAcno.SetFocus
    Else
        txtAddress = ""
        txtLng = ""
        txtLoanAcno = ""
        txtltd = ""
        txtName = ""
        txtNextDate = ""
        txtPhone = ""
        txtRemarks = ""
        txtName.SetFocus
    End If
End Sub

Private Sub lv_DblClick()
'On Error Resume Next
'    Refress_Rs rsLoan, "Select status from tblLoanData where status = 0 and sn = " & Val(ThisRec)
'    mnuPaid.Enabled = False
'    mnuunpaid.Enabled = False
'    If rsLoan.RecordCount > 0 Then
'        mnuPaid.Enabled = True
'    Else
'        mnuunpaid.Enabled = True
'    End If
'
'    PopupMenu mnu

End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set ThisRec = Item
End Sub


Private Sub mnuPaid_Click()
If Val(ThisRec) > 0 Then
    Refress_Rs rsLoan, "Select status from tblLoanData where sn = " & Val(ThisRec)
    
    If rsLoan.RecordCount > 0 Then
        ExecuteQuery "Update tblLoanData set status = 1 where sn = " & Val(ThisRec)
        Getdg
        Message "Record mark as Paid  !!!"
    End If

End If
End Sub

Private Sub mnuunpaid_Click()
If Val(ThisRec) > 0 Then
    Refress_Rs rsLoan, "Select status from tblLoanData where sn = " & Val(ThisRec)
    
    If rsLoan.RecordCount > 0 Then
        ExecuteQuery "Update tblLoanData set status = 0 where sn = " & Val(ThisRec)
        Getdg
        Message "This Record has been marked as un-Paid."
    End If

End If

End Sub

Private Sub txtAddress_GotFocus()
Colored
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLoanAcno.SetFocus
End If
End Sub

Private Sub txtAddress_LostFocus()
unColored txtAddress
End Sub

Private Sub txtCID_GotFocus()
Colored
End Sub

Private Sub txtCID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    getdata
    Getdg
End If
End Sub

Private Sub txtCID_LostFocus()
unColored txtCID
End Sub

Private Sub txtLng_GotFocus()
Colored
End Sub

Private Sub txtLng_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbType.SetFocus
End If
End Sub

Private Sub txtLng_LostFocus()
unColored txtLng
End Sub

Private Sub txtLoanAcno_GotFocus()
Colored
End Sub

Private Sub txtLoanAcno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkAll.Value = vbChecked
    Getdg
    txtPhone.SetFocus
End If
End Sub

Private Sub txtLoanAcno_LostFocus()
unColored txtLoanAcno
End Sub

Private Sub txtltd_GotFocus()
Colored
End Sub

Private Sub txtltd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLng.SetFocus
End If
End Sub

Private Sub txtltd_LostFocus()
unColored txtltd
End Sub

Private Sub txtName_GotFocus()
Colored
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAddress.SetFocus
End If
End Sub

Private Sub txtName_LostFocus()
unColored txtName
End Sub

Private Sub txtNextDate_GotFocus()
Colored
End Sub

Private Sub txtNextDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRemarks.SetFocus
End If
End Sub

Private Sub txtNextDate_LostFocus()
On Error Resume Next
unColored txtNextDate

txtNextDate = Format(txtNextDate, "yyyy/mm/dd")
End Sub

Private Sub txtPhone_GotFocus()
Colored
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtltd.SetFocus
End If
End Sub

Private Sub txtPhone_LostFocus()
unColored txtPhone
End Sub

Private Sub txtRemarks_GotFocus()
Colored
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub txtRemarks_LostFocus()
unColored txtRemarks
End Sub

Private Sub clrData()
        txtAddress = ""
        txtLng = ""
        txtLoanAcno = ""
        txtltd = ""
        txtName = ""
        txtNextDate = ""
        txtPhone = ""
        txtRemarks = ""
        txtCID = ""
End Sub

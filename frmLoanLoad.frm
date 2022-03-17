VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoanLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load Excel File"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15945
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   15945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbsc 
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   5295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   12938
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4560
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upload Excel"
      Height          =   495
      Left            =   14040
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmLoanLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim list As ListItem
'Dim rsChecked As Recordset
Dim rstrans As Recordset
Dim SortOrder As String
Dim SortASCDSC As String
Dim thisBranch As Integer

Private Sub changeASCDSC()
    If SortASCDSC = "ASC" Then
        SortASCDSC = "DESC"
    Else
        SortASCDSC = "ASC"
    End If
End Sub

Private Sub cmbSC_Click()
thisBranch = cmbsc.ItemData(cmbsc.ListIndex)
getList
End Sub

Private Sub Command1_Click()
Dim I As Single
Dim Rng  As String
Dim wbook As Workbook
'---------------------------------------------------------------------------------------
cd.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"
'cd.Filter = "PostScript|*.pdf" & "|Word Document|*.doc"
cd.ShowOpen
'-----------------------------------------------------------------------------------------------------
If Not (UCase(fileinfo.GetExtensionName(cd.FileName)) = "XLSX" Or UCase(fileinfo.GetExtensionName(cd.FileName)) = "XLS") Then
    MsgBox "Only Excel File is supported."
    Exit Sub
End If
'-----------------------------------------------------------------------------------------------
Set wbook = Excel.Workbooks.Open(cd.FileName)
'---------------------------------------------------------------------------------------------------
'If Not Left(wbook.Windows(1).Caption, 16) = "Masterlist_Loan_" Then
'    Message "wrong file."
'End If
'------------------------------------------------------------------------
Me.MousePointer = 11
'---------------------------------------------------------------------------------------------------
ExecuteQuery "Delete from tblLoanFile where branch = " & CurrenBranchID
Refress_Rs rstrans, "Select * from tblLoanFile"
'Refress_Rs rsChecked, "Select * from tblloandata where branch = " & CurrenBranchID
'-----------------------------------------------------------------------------------------------------------------




For I = 4 To 200000
    Rng = "a" & I & ":a" & I
    Range(Rng).Select
    If Not Selection.Cells(1, 1) = "" Then
        If Val(Selection.Cells(1, 27)) > 0 Then
            rstrans.AddNew
            rstrans!AccountNo = Selection.Cells(1, 6) & "-" & Selection.Cells(1, 7)
            rstrans!AccountName = Selection.Cells(1, 9)
            rstrans!CID = Selection.Cells(1, 10)
            rstrans!phoneno = Replace(Replace(Selection.Cells(1, 12), "     ", ","), " ", "")
            rstrans!BankiLoan = Selection.Cells(1, 25)      'y
            rstrans!ODLoan = Selection.Cells(1, 27)
            rstrans!TotalDue = Selection.Cells(1, 30)
            rstrans!kista = Val(Selection.Cells(1, 32))
            rstrans!Branch = CurrenBranchID
            rstrans!UserNO = CurrenUser
            rstrans!addess = Selection.Cells(1, 11)
            rstrans!dated = myDate
'            If rsChecked!LoanAc = rstrans!AccountNo Then
'
'            End If
            rstrans!Status = 0
            rstrans!SN = NewMaxID("tblLoanFile", "SN")
            
            rstrans.Update
            rstrans.Requery
        End If
    Else
        Exit For
    End If
Next
'-------------------------------------------------------------------------------------------------------------------------
getList
'dg.Columns(0).Width = 1500
'dg.Columns(1).Width = 1500
'------------------------------------------------------------------------------------------------------------------------
Me.MousePointer = 0
Excel.Workbooks.Close
Message "Data Updated successfully."
End Sub


Private Sub Form_Load()
    Set rstrans = New Recordset
    Refress_Rs rstrans, "Select * from tblserviceCenter"
    
    For I = 0 To rstrans.RecordCount - 1
        cmbsc.AddItem Format(rstrans!code, "000") & " : " & rstrans!ServiceCenterName
        cmbsc.ItemData(I) = rstrans!SN
        rstrans.MoveNext
    Next
    rstrans.Filter = " sn = " & CurrenBranchID
    
    cmbsc.Text = Format(rstrans!code, "000") & " : " & CurrenBranchName
    rstrans.Filter = ""
    If userType = 1 Then
        cmbsc.Enabled = True
    Else
        cmbsc.Enabled = False
    End If
    thisBranch = CurrenBranchID
    SortASCDSC = "ASC"
    SortOrder = "Status, sn "
    getList
End Sub

Public Sub getList()
    Refress_Rs rstrans, "Select sn a, * from tblLoanFile where Branch = " & thisBranch & " order by  " & SortOrder & " " & SortASCDSC
    If rstrans.RecordCount > 0 Then
        ListView1.Visible = True
        ListView1.Font.Size = 8
        GenerateListView ListView1, rstrans
        ListView1.ColumnHeaders(1).Width = 0
        ListView1.ColumnHeaders(2).Width = 800
        ListView1.ColumnHeaders(3).Width = 1000
        ListView1.ColumnHeaders(4).Width = 1200
        ListView1.ColumnHeaders(11).Width = 700
        ListView1.ColumnHeaders(12).Width = 0
        ListView1.ColumnHeaders(13).Width = 0
        ListView1.ColumnHeaders(15).Width = 700
        
        For I = 1 To rstrans.RecordCount
            ListView1.ListItems(I).SubItems(1) = I
            If Not UCase(ListView1.ListItems(I).SubItems(14)) = "1" Then
                For j = 0 To ListView1.ColumnHeaders.Count - 1
                    If j = ListView1.ColumnHeaders.Count - 1 Then
                        ListView1.ListItems(I).ListSubItems(j).ForeColor = vbRed
                    Else
                        ListView1.ListItems(I).ListSubItems(j + 1).ForeColor = vbRed
                    End If
                Next
            End If
        Next
    
    Else
        ListView1.Visible = False
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If Not ColumnHeader.Index = 14 Then
    SortOrder = ColumnHeader
    changeASCDSC
    getList
End If
End Sub

Private Sub ListView1_DblClick()
Refress_Rs rstrans, "Select * from tblLoanData where Branch = " & thisBranch & " and CID = " & list.SubItems(2)
If rstrans.RecordCount > 0 Then
    frmLoanCounseling.txtCID = rstrans!CID
    frmLoanCounseling.txtName = rstrans!cname
    frmLoanCounseling.txtAddress = rstrans!Address
    frmLoanCounseling.txtPhone = rstrans!Phone
    frmLoanCounseling.txtLoanAcno = rstrans!LoanAc
    
Else
    frmLoanCounseling.txtCID = list.SubItems(2)
    frmLoanCounseling.txtName = list.SubItems(4)
    frmLoanCounseling.txtAddress = list.SubItems(6)
    frmLoanCounseling.txtPhone = list.SubItems(5)
    frmLoanCounseling.txtLoanAcno = list.SubItems(3)
End If
frmLoanCounseling.cmbsc = cmbsc
frmLoanCounseling.cmbsc.tag = thisBranch
frmLoanCounseling.Getdg
frmLoanCounseling.Show vbModal
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set list = Item
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLoanReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
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
   ScaleHeight     =   8235
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFile 
      Caption         =   "File Import"
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   7800
      Width           =   1575
   End
   Begin VB.ComboBox cmbSC 
      Height          =   390
      Left            =   1680
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   1320
      Width           =   4575
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   390
      Left            =   3840
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cmbType 
      Height          =   390
      Left            =   1680
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtFrmFDate 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8640
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtfrmToDate 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   10920
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdRpt 
      Caption         =   "Filter"
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Get Excel"
      Height          =   375
      Left            =   10320
      TabIndex        =   6
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtTodate 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   10920
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFrmDate 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8640
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Flowup DaeFrom :-"
      Height          =   375
      Index           =   5
      Left            =   6360
      TabIndex        =   13
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "To :-"
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblstatus 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   7800
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Flowup report"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   1215
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12495
   End
   Begin VB.Label Label1 
      Caption         =   "To :-"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Entry Date From :-"
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmLoanReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRpt As Recordset
Dim strDate As String
Dim strUser As String
Dim strBranch As String
Dim strStatus As String
Dim strType As String
Dim strNexDate As String
Dim Changed As Boolean

Private Sub cmbSC_Click()
If cmbsc.ListIndex = 0 Then
    strBranch = ""
Else
    strBranch = " and Branch = " & cmbsc.ItemData(cmbsc.ListIndex)
End If
Changed = True
End Sub

Private Sub cmbStatus_Click()
If cmbStatus.ListIndex = 0 Then
    strStatus = ""
ElseIf cmbStatus.ListIndex = 1 Then
    strStatus = " and Status =  1 "
ElseIf cmbStatus.ListIndex = 2 Then
    strStatus = " and Status =  0 "
End If
Changed = True
End Sub

Private Sub cmbType_Click()
If cmbType.ListIndex = 0 Then
    strType = ""
Else
    strType = " and Type =  " & cmbType.ListIndex - 1
End If
Changed = True
End Sub

Private Sub cmdExcel_Click()
cmdExcel.Enabled = False

If Changed = False And rsRpt.RecordCount > 0 Then
Dim stag As String
    stag = "Branch : " & Trim(Mid(cmbsc.Text, InStr(1, cmbsc.Text, ":") + 1))
    If Not Trim(strDate) = "" Then stag = stag & strDate
    If Not Trim(strNexDate) = "" Then stag = stag & Replace(strNexDate, "NextDate", "FlowUpDate")
    stag = stag & " and Type : " & cmbType.Text
    stag = stag & " and Status : " & cmbStatus.Text
    
    ExportToExcelFromRecordSet rsRpt, myYes, myno, stag
Else
    Message "Please Filter the Record First"
End If
cmdExcel.Enabled = True
End Sub

Private Sub cmdFile_Click()
frmLoanLoad.Show vbModal
End Sub

Private Sub cmdRpt_Click()
Refress_Rs rsRpt, "select Dated, *,typed = case when type=0 then 'Phone Call' when type = 1 then 'Visiting' when type = 2 then 'White Letter' when type = 3 then 'Yellow Letter' when Type = 4 then 'Red Letter' else 'Other' end, Paidstatus = case when status = 0 then 'Un-Paid' else 'Paid' end from tblLOanData Where LoanAc in(select AccountNo from tblLoanFile where branch = " & Replace(strBranch, " and Branch = ", "") & ") " & strDate & strBranch & strNexDate & strStatus & strType & strUser & " order by nextDate Asc"
manageDG
If rsRpt.RecordCount > 0 Then
    lblstatus = rsRpt.RecordCount & " Records found"
Else
    lblstatus = "No Record Found."
End If
Changed = False
End Sub

Private Sub manageDG()
Set dg.DataSource = rsRpt
dg.RowHeight = 275
dg.Columns(1).Visible = False
dg.Columns(2).Width = 1000
dg.Columns(3).Visible = False
dg.Columns(4).Visible = False
dg.Columns(5).Width = 2500
dg.Columns(6).Width = 2500
dg.Columns(7).Width = 2500
dg.Columns(8).Width = 1500
dg.Columns(9).Visible = False
dg.Columns(10).Visible = False
dg.Columns(11).Visible = False
dg.Columns(13).Visible = False
dg.Columns(15).Visible = False


End Sub

Private Sub dg_DblClick()
On Error Resume Next
Dim rstrans As Recordset
Set rstrans = New Recordset

If dg.row = -1 Then
    Message "Please Filter record first."
    Exit Sub
End If
Refress_Rs rstrans, "Select * from tblLoanData where Branch = " & Replace(strBranch, " and Branch = ", "") & " and CID = " & Val(dg.Columns(2))
If rstrans.RecordCount > 0 Then
    frmLoanCounseling.cmbsc = cmbsc
    frmLoanCounseling.cmbsc.tag = Replace(strBranch, " and Branch = ", "")
    frmLoanCounseling.txtCID = rstrans!CID
    frmLoanCounseling.txtname = rstrans!cname
    frmLoanCounseling.txtAddress = rstrans!Address
    frmLoanCounseling.txtPhone = rstrans!Phone
    frmLoanCounseling.txtLoanAcno = rstrans!LoanAc
    
'Else
'    frmLoanCounseling.txtCID = dg.Columns(1)
'    frmLoanCounseling.txtname = list.SubItems(4)
'    frmLoanCounseling.txtAddress = ""
'    frmLoanCounseling.txtPhone = list.SubItems(5)
'    frmLoanCounseling.txtLoanAcno = list.SubItems(3)
End If
    
frmLoanCounseling.Getdg
frmLoanCounseling.Show vbModal


End Sub

Private Sub Form_Load()
Set rsRpt = New Recordset


    strDate = ""
    strBranch = " and Branch = " & CurrenBranchID
    strStatus = " and Status =  0 "
    strType = ""
    strUser = ""
    strNexDate = ""
    
    cmbType.AddItem "All Type"
    cmbType.AddItem "Phone Call"
    cmbType.AddItem "Visiting"
    cmbType.AddItem "White Letter"
    cmbType.AddItem "Yellow Letter"
    cmbType.AddItem "Red Letter"
    cmbType.Text = cmbType.list(0)
    
    cmbStatus.AddItem "All Status"
    cmbStatus.AddItem "Cleard"
    cmbStatus.AddItem "Un-Cleared"
    cmbStatus.Text = cmbStatus.list(2)
        
    Refress_Rs rsRpt, "Select * from tblserviceCenter "
    cmbsc.AddItem "All ServiceCenter"
    If rsRpt.RecordCount > 0 Then
        Do While Not rsRpt.EOF
            cmbsc.AddItem Format(rsRpt!code, "000") & " : " & rsRpt!ServiceCenterName
            cmbsc.ItemData(cmbsc.ListCount - 1) = rsRpt!SN
            rsRpt.MoveNext
        Loop
    End If
    
    rsRpt.Filter = " sn = " & CurrenBranchID
    
    cmbsc.Text = Format(rsRpt!code, "000") & " : " & CurrenBranchName
    rsRpt.Filter = ""
    
    If userType = 1 Then
        cmbsc.Enabled = True
    Else
        cmbsc.Enabled = False
    End If
'    Else
'
'        Refress_Rs rsRpt, "Select * from tblServiceCenter where sn= " & CurrenBranchID
'        cmbsc.AddItem Format(rsRpt!code, "000") & " : " & rsRpt!serviceCenterName
'        cmbsc.ItemData(0) = CurrenBranchID
'        cmbsc.Text = cmbsc.list(0)
'        cmbsc.Enabled = False
'    End If
    Changed = True
End Sub

Private Sub txtFrmDate_Change()
    If Not (txtFrmDate = "" Or txtTodate = "") Then
        strDate = " and dated between '" & txtFrmDate & "' and '" & txtTodate & "' "
        txtFrmFDate = ""
        txtfrmToDate = ""
        txtFrmFDate_Change
    Else
        strDate = ""
    End If
    Changed = True
End Sub

Private Sub txtFrmDate_DblClick()
    txtFrmDate = myDate
    txtTodate = myDate
    txtFrmDate_Change
    Changed = True
End Sub

Private Sub txtFrmFDate_Change()
If Not (txtFrmFDate = "" Or txtfrmToDate = "") Then
    strNexDate = " and NextDate between '" & txtFrmFDate & "' and '" & txtfrmToDate & "' "
    txtFrmDate = ""
    txtTodate = ""
    txtFrmDate_Change
Else
    strNexDate = ""
End If
Changed = True
End Sub

Private Sub txtFrmFDate_DblClick()
txtFrmFDate = myDate
txtfrmToDate = GetMasantti(myDate)
txtFrmFDate_Change
Changed = True
End Sub

Private Sub txtfrmToDate_Change()
txtFrmFDate_Change
End Sub

Private Sub txtfrmToDate_DblClick()
txtFrmFDate_DblClick
End Sub

Private Sub txtTodate_Change()
txtFrmDate_Change
End Sub

Private Sub txtTodate_DblClick()
txtFrmDate_DblClick
End Sub

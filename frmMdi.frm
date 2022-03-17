VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMdi 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMdi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgLst 
      Left            =   4800
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":259A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1535
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "imgLst"
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7830
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   4440
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   2760
   End
End
Attribute VB_Name = "frmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim rsMessage As Recordset

Private Sub Form_Load()
Me.Caption = "KPI System"
Set rsMessage = New Recordset
With tb
    .ImageList = imgLst
    .Appearance = ccFlat
    .AllowCustomize = False
    .Wrappable = False
    
    i = 1
    .Buttons.Add
    .Buttons(i).Caption = "Target"
    .Buttons(i).Image = 1
     
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Caption = "Report"
    .Buttons(i).Image = 2
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "Users"
     .Buttons(i).Image = 3
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "5S Analysis"
     .Buttons(i).Image = 4
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "Micro Group"
     .Buttons(i).Image = 5
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "Loan Flowup"
     .Buttons(i).Image = 6
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "Utilities"
     .Buttons(i).Image = 7
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "5S Message"
     .Buttons(i).Image = 8
    
    i = i + 1
    .Buttons.Add
    .Buttons(i).Style = tbrPlaceholder
     
    i = i + 1
     .Buttons.Add
     .Buttons(i).Caption = "Exit"
     .Buttons(i).Image = 9
     

End With
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If Not userType = 3 Then
            frmHead.Show vbModal
        Else
            frmStaffData.Show vbModal
        End If
    Case 3
        frmRepot.Show vbModal
    Case 5
        frmStaff.Show vbModal
    Case 7
        frm5S.Show vbModal
    Case 9
        frmMF.Show vbModal
    Case 11
        frmLoanReport.Show vbModal
    Case 13
        frmUtilities.Show vbModal
    Case 15
        frmMessageBord.Show vbModal
    Case 17
        End
End Select

End Sub

Private Sub Timer1_Timer()

sb.Panels(9).Text = Time
sb.Panels(8).Text = myDate
x = x + 100
If x > 2000 Then
    DateManage
    x = 0
End If
End Sub

Private Sub Timer2_Timer()
    If (Screen.ActiveForm.Name = frmMdi.Name) Then
        
        Refress_Rs rsMessage, "select top 5 *, (Select userFullName From tblusers where sn = m.frmuser) Sender from tblMessageCenter m Where tousers = " & CurrenUser & " And status = 0"

        
        If rsMessage.RecordCount > 0 Then
            Do While Not rsMessage.EOF
                PupupMessage rsMessage!msgText, rsMessage!sender
                ExecuteQuery "Update tblMessageCenter set status = 1 where sn = " & rsMessage!sn
                rsMessage.MoveNext
            Loop
        End If
    End If
End Sub

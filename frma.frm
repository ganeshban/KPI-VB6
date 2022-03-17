VERSION 5.00
Begin VB.Form frma 
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rsA As Recordset
Set rsA = New Recordset
Refress_Rs rsA, " Select sn [User.SN] , id as [User.id], FName  [User.FName], LName  [User.LName], DateOfBirth  [User.DOB], BloodGroup  [User.BdGrp], PrimaryPhone  [User.Contact.Phone], PrimaryEmail  [User.Contact.Email] from tblUserInformation a for json path, root('Users')"
If rsA.RecordCount > 0 Then
Text1 = rsA.Fields(0).Name

End If

End Sub

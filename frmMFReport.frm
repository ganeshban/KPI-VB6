VERSION 5.00
Begin VB.Form frmMFReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MF Report Form"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7800
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
   ScaleHeight     =   8070
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMonth 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6480
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstuser 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdGetRpt 
      Caption         =   "Get Report"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   7440
      Width           =   1815
   End
   Begin VB.ListBox lstsc 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox txtSc 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Month"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Year"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMFReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsData As Recordset
Dim Rpttag As String

Private Sub VisiableList(data As String)
On Error Resume Next

Select Case data
    Case txtSC.Name
        lstsc.Visible = True
        lstuser.Visible = False
    Case txtuser.Name
        lstuser.Visible = True
        lstsc.Visible = False
    Case Else
        lstsc.Visible = False
        lstuser.Visible = False
End Select
End Sub

Private Sub cmdgetrpt_Click()
Dim Facter As String
Dim FacterList As String
Dim FacterGroup As String
Dim rsTask As Recordset
Dim rsMF As Recordset

Me.MousePointer = vbHourglass

Set rsTask = New Recordset
Set rsMF = New Recordset

Dim dated, SCID, userID, taskID
Dim strr As String
Dim sstr As String
dated = Format(txtYear, "0000") & "/" & Format(txtMonth, "00") & "/"

If Val(txtuser.tag) > 0 Then
    userID = " and userno = " & Val(txtuser.tag)
Else
    userID = ""
End If

If Val(txtSC.tag) > 0 Then
    Facter = "MFgrp=d.MfGrp "
    FacterList = "GroupName "
    FacterGroup = "GroupName, MFGrp "
    SCID = " and userno in (Select sn from tblusers where branch= " & Val(txtSC.tag) & ") "
Else
    SCID = ""
    FacterList = "ServiceCenterName "
    FacterGroup = "ServiceCenterName, Branch "
    Facter = "Branch=d.Branch "
End If


Refress_Rs rsTask, "Select * from tblmfTask order by sn"

    Do While Not rsTask.EOF
        Refress_Rs rsData, "Select * from ViewMFData where substring(dated,1,8) = '" & dated & "' and taskID = " & rsTask!SN
        If rsData.RecordCount > 0 Then
            
            If rsTask.AbsolutePosition = 3 Then
                strr = strr & vbNewLine & " , (Select sum(Data) from ViewMFData where substring(dated,1,8) = '" & dated & "' and taskID in (1,2) and " & Facter & ") 'Total'"
                strr = strr & vbNewLine & " , (Select sum(Data) from ViewMFData where substring(dated,1,8) < '" & dated & "' and taskID in (1,2) and " & Facter & ") 'Last Month Group Member'"
                strr = strr & vbNewLine & " , (Select sum(Data) from ViewMFData where substring(dated,1,8) <= '" & dated & "' and taskID in (1,2) and " & Facter & ") 'Total Group Member'"
            
            ElseIf rsTask.AbsolutePosition = 4 Then
                strr = strr & vbNewLine & " , isnull((Select sum(Data) from ViewMFData where substring(dated,1,8) <= '" & dated & "' and taskID in (1,2) and " & Facter & "),0)+ isnull((Select sum(Data) from ViewMFData where substring(dated,1,8) = '" & dated & "' and taskID in (3) and " & Facter & "),0) 'Total Collection Member'"
                
                If SCID = "" Then
                    strr = strr & vbNewLine & " , isnull((Select count( distinct MFGrp) from ViewMFData where substring(dated,1,8) <= '" & dated & "' and  Branch=d.Branch ),0) 'MFCount'"
                End If
            End If
            
            strr = strr & vbNewLine & " , (Select sum(Data) from ViewMFData where substring(dated,1,8) = '" & dated & "' and taskID = " & rsTask!SN & " and  " & Facter & ") '" & rsTask!TaskName & "'"
                                
        End If
        
        If rsTask.AbsolutePosition = rsTask.RecordCount Then
            strr = strr & vbNewLine & " , (Select sum(Data) from ViewMFData where substring(dated,1,8) = '" & dated & "' and countable =1 and  " & Facter & ") 'Total Amount'"
        End If
        
        rsTask.MoveNext
    Loop



sstr = "Select  " & FacterList & ", max(substring(Dated,9,2)) days " & strr & vbNewLine & " from ViewMFData d where substring(dated,1,8)='" & dated & "'" & userID & SCID & vbNewLine & " group by  " & FacterGroup & vbNewLine & " order by days"
Refress_Rs rsData, sstr
If rsData.RecordCount > 0 Then
    If Val(txtuser.tag) > 0 Then
        Rpttag = " User : " & Mid(txtuser, InStr(1, txtuser, ":") + 2)
    Else
        Rpttag = ""
    End If
'    tojson rsData
    rsData.MoveFirst
    ExportToExcelFromRecordSet rsData, myYes, myYes, "Service Center : " & Mid(txtSC, InStr(1, txtSC, ":") + 2) & Rpttag & " Date : " & Val(txtYear) & ", " & txtMonth
    
Else
    Message "Record not found. "
End If

Me.MousePointer = vbDefault
End Sub

Private Sub cmdGetRpt_GotFocus()
VisiableList ""
End Sub

Private Sub Form_Load()
Set rsData = New Recordset
txtMonth = Mid(myDate, 6, 2)
txtYear = Mid(myDate, 1, 4)

If userType = 1 Or CurrenUser = 4 Then
    GetScList
Else
    txtSC = CurrenBranchName
    txtSC.tag = CurrenBranchID
    txtSC.Enabled = False
End If
If userType = 1 Or userType = 2 Then
    getUser
Else
    txtuser.tag = CurrenUser
    txtuser = frmMdi.sb.Panels(2)
    txtuser.Enabled = False
End If

End Sub

Private Sub GetScList()
Refress_Rs rsData, "Select * from tblServiceCenter"
lstsc.Clear
If rsData.RecordCount > 0 Then
    lstsc.AddItem "000 : All Service Center"
    lstsc.ItemData(lstsc.ListCount - 1) = 0
    Do While Not rsData.EOF
        lstsc.AddItem rsData!code & " : " & rsData!ServiceCenterName
        lstsc.ItemData(lstsc.ListCount - 1) = rsData!SN
        rsData.MoveNext
    Loop
    txtSC = CurrenBranchName
    txtSC.tag = CurrenBranchID
End If
End Sub

Private Sub getUser()
Refress_Rs rsData, "Select * from Viewusers where status = 0 and userType = 3 and Branch = " & Val(txtSC.tag)
lstuser.Clear
lstuser.AddItem "000 : All Staff "
lstuser.ItemData(0) = 0

If rsData.RecordCount > 0 Then
    Do While Not rsData.EOF
        lstuser.AddItem Format(rsData!SN, "000") & " : " & rsData!UserFullName
        lstuser.ItemData(rsData.AbsolutePosition) = rsData!SN
        rsData.MoveNext
    Loop
    lstuser.ListIndex = 0
End If
End Sub

Private Sub lstsc_DblClick()
txtSC_KeyPress 13
End Sub

Private Sub lstsc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSC_KeyPress 13
End If
End Sub

Private Sub lstuser_DblClick()
txtuser_KeyPress 13
End Sub

Private Sub lstuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtuser_KeyPress 13
End If
End Sub

Private Sub txtMonth_GotFocus()
Colored
VisiableList ""
End Sub

Private Sub txtMonth_LostFocus()
unColored txtMonth
End Sub

Private Sub txtSC_Change()
If lstsc.Visible = True Then
    If Not txtSC.Text = "" Then
        For I = 0 To lstsc.ListCount - 1
            If Val(txtSC) = Val(lstsc.List(I)) Then
                lstsc.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtSC.Text)) = UCase(Trim(Mid(lstsc.List(I), InStr(1, lstsc.List(I), ":") + 2, Len(txtSC.Text)))) Or Val(txtSC.Text) = Val(Mid(lstsc.List(I), 1, InStr(1, lstsc.List(I), ":") - 2)) Then
                    lstsc.Selected(I) = True
                    Exit For
                Else
                    lstsc.Selected(I) = False
                End If
            End If
        Next
    Else
        lstsc.Selected(0) = False
    End If
End If

End Sub

Private Sub txtSC_GotFocus()
Colored
VisiableList txtSC.Name
End Sub

Private Sub txtSC_KeyDown(KeyCode As Integer, Shift As Integer)
If lstsc.Visible = True Then
    If KeyCode = 38 Then
        If lstsc.ListIndex <= 0 Then
    '        lstscData.Selected(lstscData.ListCount - 1) = True
            Exit Sub
        Else
            lstsc.Selected(lstsc.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstsc.ListIndex = lstsc.ListCount - 1 Then
    '        lstscData.Selected(lstscData.ListCount - 1) = True
            Exit Sub
        Else
            lstsc.Selected(lstsc.ListIndex + 1) = True
        End If
    End If
End If
End Sub

Private Sub txtSC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstsc.Visible = True Then
        If lstsc.ListCount > 0 Then
            If lstsc.ListIndex >= 0 Then
                txtSC = lstsc.List(lstsc.ListIndex)
                txtSC.tag = Val(lstsc.ItemData(lstsc.ListIndex))
                If Val(txtSC.tag) > 0 Then
                    getUser
                Else
                    lstuser.Clear
                    lstuser.AddItem "000 : All Users"
                End If
                lstuser.ListIndex = 0
                txtuser = lstuser.List(0)
                txtuser.tag = Val(lstuser.List(0))
                txtuser.SetFocus
            End If
        End If
    End If
End If
End Sub

Private Sub txtSC_LostFocus()
unColored txtSC

End Sub

Private Sub txtUser_Change()
If lstuser.Visible = True Then
    If Not txtuser.Text = "" Then
        For I = 0 To lstuser.ListCount - 1
            If Val(txtuser) = Val(lstuser.List(I)) Then
                lstuser.Selected(I) = True
                Exit For
            Else
                If UCase(Trim(txtuser.Text)) = UCase(Trim(Mid(lstuser.List(I), InStr(1, lstuser.List(I), ":") + 2, Len(txtuser.Text)))) Or Val(txtuser.Text) = Val(Mid(lstuser.List(I), 1, InStr(1, lstuser.List(I), ":") - 2)) Then
                    lstuser.Selected(I) = True
                    Exit For
                Else
                    lstuser.Selected(I) = False
                End If
            End If
        Next
    Else
        lstuser.Selected(0) = False
    End If
End If


End Sub

Private Sub txtuser_GotFocus()
Colored
VisiableList txtuser.Name
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
If lstuser.Visible = True Then
    If KeyCode = 38 Then
        If lstuser.ListIndex <= 0 Then
    '        lstuserData.Selected(lstuserData.ListCount - 1) = True
            Exit Sub
        Else
            lstuser.Selected(lstuser.ListIndex - 1) = True
        End If
    End If
    
    
    If KeyCode = 40 Then
        If lstuser.ListIndex = lstuser.ListCount - 1 Then
    '        lstuserData.Selected(lstuserData.ListCount - 1) = True
            Exit Sub
        Else
            lstuser.Selected(lstuser.ListIndex + 1) = True
        End If
    End If
End If
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstuser.Visible = True Then
        If lstuser.ListCount > 0 Then
            txtuser = lstuser.List(lstuser.ListIndex)
            txtuser.tag = Val(lstuser.ItemData(lstuser.ListIndex))
            cmdGetRpt.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtuser_LostFocus()
unColored txtuser

End Sub

Private Sub txtYear_GotFocus()
VisiableList ""
Colored
End Sub

Private Sub txtYear_LostFocus()
unColored txtYear
End Sub

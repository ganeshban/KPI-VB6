VERSION 5.00
Begin VB.Form frmMessageBord 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12195
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
   ScaleHeight     =   8520
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Preeti"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   7800
      Width           =   12015
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   9120
      TabIndex        =   2
      Top             =   -3120
      Width           =   3855
   End
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   3855
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   6700
      Left            =   0
      ScaleHeight     =   6705
      ScaleWidth      =   12135
      TabIndex        =   0
      Top             =   960
      Width           =   12135
      Begin VB.VScrollBar vs 
         Height          =   6735
         Left            =   11880
         Max             =   1
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   0
         ScaleHeight     =   6375
         ScaleWidth      =   12015
         TabIndex        =   1
         Top             =   120
         Width           =   12015
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   -1080
            Width           =   11775
         End
         Begin VB.Label lblTime 
            BackColor       =   &H00FF0000&
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   6
            Top             =   -480
            Width           =   1215
         End
         Begin VB.Label lblMsg 
            BackColor       =   &H00FF0000&
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Preeti"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   5
            Top             =   -360
            Width           =   1695
         End
         Begin VB.Shape sh 
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   0
            Left            =   4920
            Shape           =   4  'Rounded Rectangle
            Top             =   -720
            Width           =   2175
         End
      End
   End
   Begin VB.Image imgsend 
      Height          =   480
      Left            =   11520
      Picture         =   "frmMessageBord.frx":0000
      Top             =   7920
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmMessageBord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsData As Recordset
Dim rsmsg As Recordset
Public UserNO As Integer
Dim oldVal As Integer


Private Sub sendmsg()
If lst.ItemData(lst.ListIndex) > 0 Then
    ExecuteQuery "Insert into tblMessageCenter values (" & NewMaxID("tblMessageCenter", "SN") & ", " & CurrenUser & ", " & Val(lst.ItemData(lst.ListIndex)) & ", '" & Replace(txtMsg, "'", "''") & "', '" & Time & "', '" & myDate & "', 0)"
Else
    If lst.ItemData(lst.ListIndex) = -1 Then
        Refress_Rs rsData, "Select * from tblServiceCenter where incharge>0"
    Else
        Refress_Rs rsData, "Select sn as incharge from tblusers where status=0"
    End If
    
    If rsData.RecordCount > 0 Then
        Do While Not rsData.EOF
            ExecuteQuery "Insert into tblMessageCenter values (" & NewMaxID("tblMessageCenter", "SN") & ", " & CurrenUser & ", " & Val(rsData!incharge) & ", '" & Replace(txtMsg, "'", "''") & "', '" & Time & "', '" & myDate & "', 0)"
            rsData.MoveNext
        Loop
    End If

End If
End Sub

Private Sub msgUnload()
For I = 1 To lblMsg.Count
    If lblMsg.Count > 1 Then
        Unload lblMsg(I)
        Unload sh(I)
        Unload lblTime(I)
    End If
Next
For I = 1 To lblDate.Count
    If lblDate.Count > 1 Then
        Unload lblDate(I)
    End If
Next
End Sub

Private Sub msgLoad()
Dim code As Integer
Dim SQL As String

If userType = 1 Then
    SQL = "Select * from tblMessageCenter where frmuser = " & CurrenUser & " and tousers = " & UserNO & "  order by sn "
Else
    SQL = "Select * from tblMessageCenter where tousers = " & CurrenUser & " and frmuser = " & UserNO & "  order by sn "
End If
Refress_Rs rsmsg, SQL

msgUnload
If rsmsg.RecordCount > 0 Then
    Do While Not rsmsg.EOF
        If userType = 1 Then
            code = 1
        Else
            code = 0
        End If
        
        writeMsg rsmsg.AbsolutePosition, rsmsg!msgText, rsmsg!msgDate, rsmsg!msgTime, 1, rsmsg!SN
        
        rsmsg.MoveNext
    Loop
    
    If picBack.Height < picMessage.Height Then
        X = picMessage.Height / picBack.Height
        vs.Max = Round(X + 0.49, 0)
        vs.Visible = True
        vs.Value = vs.Max
    Else
        vs.Visible = False
    End If
Else
    vs.Visible = False
End If

End Sub

Private Sub writeMsg(nNumber As Integer, msgText As String, msgDate As String, msgTime As String, msgType As Integer, msgID As Single)
Load lblMsg(nNumber)
Load lblTime(nNumber)
Load sh(nNumber)

If Not msgDate = lblDate(lblDate.Count - 1) Then
    Load lblDate(lblDate.Count)
    lblDate(lblDate.Count - 1).Visible = True
    lblDate(lblDate.Count - 1) = msgDate
    lblDate(lblDate.Count - 1).Top = sh(sh.Count - 2).Top + sh(sh.Count - 2).Height + 240
    lblMsg(nNumber).Top = lblDate(lblDate.Count - 1).Top + lblDate(lblDate.Count - 1).Height + 150

Else
    lblMsg(nNumber).Top = sh(nNumber - 1).Top + sh(nNumber - 1).Height + 150
End If


'lblMsg(nNumber).WordWrap = True
lblMsg(nNumber).tag = msgID
lblMsg(nNumber).Visible = True
lblMsg(nNumber) = msgText
lblMsg(nNumber).AutoSize = True
'Debug.Print lblMsg(nNumber).Width - picMessage.Width


lblTime(nNumber) = Format(msgTime, "hh:mm AMPM")
lblTime(nNumber).AutoSize = True
lblTime(nNumber).Visible = True


If msgType = 1 Then
    lblMsg(nNumber).BackColor = &HE0E0E0
    sh(nNumber).FillColor = &HE0E0E0
    lblMsg(nNumber).ForeColor = &H0&
    lblTime(nNumber).ForeColor = &H0&
    lblTime(nNumber).BackColor = &HE0E0E0
    lblMsg(nNumber).Left = 280
Else
    lblMsg(nNumber).BackColor = &HFF0000
    sh(nNumber).FillColor = &HFF0000
    lblMsg(nNumber).ForeColor = &HFFFFFF
    lblTime(nNumber).ForeColor = &HFFFFFF
    lblTime(nNumber).BackColor = &HFF0000
    lblMsg(nNumber).Left = picMessage.Width - lblMsg(nNumber).Width - lblTime(nNumber).Width - 350
End If

If lblMsg(nNumber).Width + lblTime(nNumber).Width + 220 > picMessage.Width Then
    X = ((lblMsg(nNumber).Width + lblTime(nNumber).Width + 220) / picMessage.Width) + 0.49
    lblMsg(nNumber).AutoSize = False
    lblMsg(nNumber).Height = (Math.Round(X, 0) * lblMsg(0).Height) + 150
    lblMsg(nNumber).Width = picMessage.Width - lblMsg(nNumber).Left - 1200
'    lblMsg(nNumber).WordWrap = True
Else
    lblMsg(nNumber).AutoSize = True
End If



lblTime(nNumber).Top = lblMsg(nNumber).Top + lblMsg(nNumber).Height - lblTime(nNumber).Height
lblTime(nNumber).Left = lblMsg(nNumber).Left + lblMsg(nNumber).Width + 60



sh(nNumber).Top = lblMsg(nNumber).Top - 60
sh(nNumber).Height = lblMsg(nNumber).Height + 150
sh(nNumber).Left = lblMsg(nNumber).Left - 60
sh(nNumber).Width = lblTime(nNumber).Width + lblTime(nNumber).Left + 120 - lblMsg(nNumber).Left
sh(nNumber).Visible = True



picMessage.Height = sh(nNumber).Top + sh(nNumber).Height + 80
'If Not picMessage.Top > 0 Then
    picMessage.Top = picBack.Height - picMessage.Height - 120
'End If
End Sub

Private Sub VisiableList(Optional data As String = "")
Select Case data
    Case txtuser.Name
        lst.Visible = True
    Case Else
        lst.Visible = False
End Select
End Sub

Private Sub Form_Load()
Set rsData = New Recordset
Set rsmsg = New Recordset
lst.Top = txtuser.Top + txtuser.Height + 20
lst.Left = txtuser.Left
getUserList
If userType = 1 Then
    txtMsg.Visible = True
    imgsend.Visible = True
Else
    txtMsg.Visible = False
    imgsend.Visible = False
End If
End Sub

Private Sub imgsend_Click()
txtMsg_KeyPress 13
End Sub

Private Sub lst_Click()
txtuser.SetFocus
End Sub

Private Sub lst_DblClick()
txtuser_KeyPress 13
End Sub

Private Sub picBack_GotFocus()
VisiableList
End Sub

Private Sub picMessage_GotFocus()
VisiableList
End Sub

Private Sub picTop_GotFocus()
VisiableList
End Sub

Private Sub txtMsg_Change()
sendVisible
End Sub

Private Sub sendVisible()
If Len(Trim(txtMsg)) > 0 Then
    txtMsg.Width = picMessage.Width - txtMsg.Left - imgsend.Width - 60
Else
    txtMsg.Width = picMessage.Width - txtMsg.Left
End If
End Sub


Private Sub getUserList()
Dim SQL As String
If userType = 1 Then
    SQL = "Select * from Viewusers where status = 0 and sn <> " & CurrenUser & " order by userFullName "
Else
    SQL = "Select * from Viewusers where status = 0 and sn in(Select distinct(frmUser) from tblMessageCenter where tousers = " & CurrenUser & ") order by userFullName "
End If

Refress_Rs rsData, SQL

If rsData.RecordCount > 0 Then
    lst.Clear
    
    If userType = 1 Then
        lst.AddItem "000 : All User"
        lst.ItemData(0) = -2
        lst.AddItem "000 : All Incharge"
        lst.ItemData(1) = -1
    End If
    
    Do While Not rsData.EOF
        lst.AddItem Format(rsData!SN, "000") & " : " & rsData!UserFullName
        lst.ItemData(lst.ListCount - 1) = rsData!SN
        rsData.MoveNext
    Loop

End If
End Sub

Private Sub txtMsg_GotFocus()
VisiableList
Colored
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(txtMsg)) > 0 Then
        sendmsg
        txtMsg = ""
        msgLoad
        KeyAscii = 0
        Exit Sub
    End If
End If
End Sub

Private Sub txtMsg_LostFocus()
unColored txtMsg
End Sub

Private Sub txtuser_GotFocus()
Colored
VisiableList txtuser.Name
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lst.ListCount > 0 Then
        If lst.ListIndex >= 0 Then
            UserNO = lst.ItemData(lst.ListIndex)
            txtuser = Trim(Mid(lst, InStr(1, lst, ":") + 1))
            msgLoad
            If txtMsg.Visible Then txtMsg.SetFocus Else picBack.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtuser_LostFocus()
unColored txtuser
End Sub


Private Sub vs_Change()
If vs.Value >= oldVal Then
    picMessage.Top = picMessage.Top - (picBack.Height / 2)
Else
    picMessage.Top = picMessage.Top + (picBack.Height / 2)
End If
oldVal = vs.Value
If txtMsg.Visible = True Then txtMsg.SetFocus
End Sub


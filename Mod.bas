Attribute VB_Name = "Mod"
Public Cnn  As Connection
Public ServerName As String
Public DbName As String
Public DatabasePassword As String
Public CurrentData As Long
Public myDate As String
Public fileinfo As FileSystemObject
Public OrgID As Single
Public Org As String
Public OrgAddress As String

Public Enum MessageResponse
    Yes = 1
    No = 2
    Ok = 0
End Enum

Public Enum MessageBtn
    okOnly = 0
    YesNo = 1
End Enum

Public Enum YesNoOption
    myYes = 0
    myno = 1
End Enum
Public CurrentMsgResponce As MessageResponse


Public CurrenUser As Integer
Public userType As Integer

Public CurrenBranchName As String
Public CurrenBranchID As Integer

Public AppVersion As String
Public databaseVersion As String

Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)





Public Sub Main()
If App.PrevInstance = True Then
    End
End If


Set fileinfo = New FileSystemObject
ServerName = GetSetting(App.Title, "Login", "Server")
DbName = GetSetting(App.Title, "Login", "Database")
DatabasePassword = GetSetting(App.Title, "Login", "psw")
If ServerName = "" Or DbName = "" Then
    frmServerProperties.Show
    Exit Sub
End If
ConnDB
If Err.Number = -2147467259 Or Err.Number = -2147467259 Or Err.Number = -2147217843 Then
    Message Err.Description & " " & Err.Number
    frmServerProperties.Show
    Exit Sub
End If

    

    checkVersion
    DateManage
    setOrg
    frmMdi.Show
    frmLogin.Show vbModal

End Sub

Private Sub a()
frma.Show
End Sub

Private Sub getNewVer()
Dim rsD As Recordset
Set rsD = New Recordset


Refress_Rs rsD, "Select * from tblsetting where sn = 2"
If Not rsD.RecordCount > 0 Then
    ExecuteQuery "Insert into tblSetting values(2,'Version Code','" & AppVersion & "')"
End If

rsD.Requery
databaseVersion = rsD!Value
If databaseVersion = AppVersion Then
    Exit Sub
End If

newVer:
Message "New Version Detected !! some update are avaiable and being process to update."
GoTo dedetctVer


dedetctVer:
If databaseVersion = "0.0.7" Then
    ExecuteQuery "ALTER   View viewTaskAchive as Select *, isnull((Select sum(Achive) from tblTaskAchive where Approved=1 and GTSN=gt.sn),0) Achive, isnull((Select userID from tblusers where sn=gt.ToUser),0) userID, isnull((Select userFullName from tblusers where sn=gt.ToUser),0) userName from ViewgivenTask gt Where target > 0 "
    ExecuteQuery "Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If

If databaseVersion = "0.0.8" Then
    ExecuteQuery "Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If

If databaseVersion = "0.0.9" Then
    ExecuteQuery "Create table tblLoginLog (sn bigint primary key, Dated varchar(10), timee varchar(20), userNo bigint references tblusers(sn), PCName varchar(40), status int)"
    ExecuteQuery " alter table tblmftask drop column taskID"
    ExecuteQuery "drop View ViewMFData "
    ExecuteQuery " Create View ViewMFData as Select mf.*, mfg.groupname, mfg.groupAddress, mfg.Branch, mfg.dayCode, mft.taskname, mft.countable, u.userid, u.userFullName, sc.ServiceCenterName, sc.code from tblmfData mf inner join tblMFGroup mfg on (mf.mfgrp=mfg.sn) inner join tblMFTask mft on (mf.SN=mft.sn) inner join tblusers u on (mf.userno=u.sn) inner join tblServiceCenter sc on (mfg.branch=sc.sn)"
    ExecuteQuery "Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If

If databaseVersion = "0.0.10" Then
    ExecuteQuery "drop View ViewMFData "
    ExecuteQuery " Create View ViewMFData as Select mf.*, mfg.groupname, mfg.groupAddress, mfg.Branch, mfg.dayCode, mft.taskname, mft.countable, u.userid, u.userFullName, sc.ServiceCenterName, sc.code from tblmfData mf inner join tblMFGroup mfg on (mf.mfgrp=mfg.sn) inner join tblMFTask mft on (mf.TaskID=mft.sn) inner join tblusers u on (mf.userno=u.sn) inner join tblServiceCenter sc on (mfg.branch=sc.sn)"
    ExecuteQuery "Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If

If databaseVersion = "0.0.11" Then
    ExecuteQuery " update tbltask set taskname='' where taskname is null"
    ExecuteQuery " alter table tbltask alter column taskname varchar(20) not null"
    ExecuteQuery " insert into tblSetting values(3,'Master PSW','mstrpk')"
    ExecuteQuery " Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If


If databaseVersion = "0.0.32" Then
    ExecuteQuery " insert into tblusers values(0,'SYSTEM','SYSTEM','SYSTEM UPDATE','SYSTEM','SYSTEM',1,1,1,'SYSTEM',NULL)"
    ExecuteQuery " Insert into tblMessageCenter values (" & NewMaxID("tblMessageCenter", "SN") & ", 0, 5, 'l;:6d ck8]6sf] g]l6lkms];g l;:6dsf] Dof;]hdf cfpg] ePsf] 5 .', '" & Time & "', '" & myDate & "', 0)"
    ExecuteQuery " Insert into tblMessageCenter values (" & NewMaxID("tblMessageCenter", "SN") & ", 0, 5, 'nf]g kmnf]ckmdf /]s{8nfO{ k]8÷cgk]8 ug{, /]s{8nfO{ 8an lSnsdf d]g'' cfpg] agfO{Psf] 5 .', '" & Time & "', '" & myDate & "', 0)"
    ExecuteQuery " Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If

If databaseVersion = "0.0.34" Then
    ExecuteQuery " Create Function GetNepaliDateToday (@dtE varchar(20)) RETURNS varchar(10) as Begin Declare @retStr Varchar(10); Select @retStr = DateN from tblDates Where year(DateE) = year(@dtE) and month(dateE)=month(@dtE) and day(dateE)=day(@dtE) RETURN ISNULL(@retStr,'NOTSPECIFY') End "
    ExecuteQuery " Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If


If databaseVersion = "0.0.35" Then
    ExecuteQuery " create table tblLoanFile ( sn bigint primary key, CID bigint, AccountNo varchar(20), accountName varchar(100), phoneNo varchar(100), addess varchar(100), BankiLoan money, ODLoan money, TotalDue money, Kista int, Branch int, userNo int, dated varchar(10), Status int )"
    ExecuteQuery " Update tblsetting set value ='" & OldPlusOneVer(databaseVersion) & "' where sn  = 2 "
    GoTo cheCKVEr
End If




ExecuteQuery "Update tblsetting set value ='" & AppVersion & "' where sn  = 2 "
GoTo cheCKVEr

cheCKVEr:
rsD.Requery
databaseVersion = rsD!Value
If databaseVersion = AppVersion Then
    Exit Sub
End If
GoTo dedetctVer
End Sub

Private Function OldPlusOneVer(oldver As String) As String
Dim X As String
X = Val(Replace(oldver, "0.0.", "")) + 1
OldPlusOneVer = "0.0." & X
End Function

Private Sub checkVersion()
'Dim filename As String
'Dim maxVerCode As String
'Dim OldVerCode As String
'Dim fso As FileSystemObject
'Dim filles As File
'Dim shll As New Shell32.Shell
'Dim fld As Shell32.Folder
'Dim slnk As Shell32.ShellLinkObject
'
''slnk.WorkingDirectory
'
'Set fso = New FileSystemObject
'For Each filles In fso.GetFolder(GetSetting(App.Title, "Login", "Location", App.Path)).Files
'    If UCase(Right$(filles, 4)) = ".EXE" Then
'        X = fso.GetFileVersion(filles)
'        If Val(Replace(X, ".", "")) > Val(Replace(OldVerCode, ".", "")) Then
'            maxVerCode = X
'            filename = filles.Name
'            OldVerCode = X
'
'        End If
'    End If
'Next

AppVersion = App.Minor & "." & App.Major & "." & App.Revision

'If Val(Replace(maxVerCode, ".", "")) > Val(Replace(AppVersion, ".", "")) Then
'    Message "New Version Found"
'    xx = "C:\Program Files\PrKPI"
'
'    If fso.FolderExists(xx) = True Then
'        fso.CopyFile filename, "C:\Program Files\PrKPI\" & filename, True
'        xx = "C:\Program Files\PrKPI\"
'    Else
'        fso.CopyFile filename, "C:\Program Files (x86)\PrKPI\" & filename, True
'        xx = "C:\Program Files (x86)\PrKPI\"
'    End If

'
'    Set fld = shll.Namespace("c:\users\" & Environ("UserName") & "\desktop")
'
'    For Each filles In fso.GetFolder("c:\users\" & Environ("UserName") & "\desktop").Files
'        If UCase(Right$(filles, 4)) = ".LNK" Then
'        Set slnk = fld.Items.Item(filles.Name).GetLink
'
'            If Replace(slnk.Path, fso.GetFileName(slnk.Path), "") = xx Then
'    '            Message "Co"
'                slnk.Path = xx & filename
'                slnk.Save
'                Exit For
'            End If
'        End If
'    Next
    
    getNewVer
'End If
End Sub

Public Sub DateManage()
Dim rsD As Recordset
Set rsD = New Recordset
Refress_Rs rsD, "Select * from sysobjects where xtype='u' and name = 'tblsetting'"
If rsD.RecordCount > 0 Then
    If UCase(Environ("ComputerName")) = UCase(ServerName) Then
        Refress_Rs rsD, "Select * from tblsetting where sn = 1"
        If rsD.RecordCount > 0 Then
            ExecuteQuery "update tblsetting set value=(select dbo.getNepaliDateToday(getdate())) where sn =1"
        Else
            ExecuteQuery "Insert into tblSetting values(1,'SystemDate','" & EngDateTONep(Now()) & "')"
        End If
    End If
    Refress_Rs rsD, "Exec GetNepaliDate"
    If rsD.RecordCount > 0 Then
        myDate = rsD!DateN
    End If
Else
    ExecuteQuery "Create table tblsetting (sn int primary key, Code varchar(50), Value VarChar(50))"
End If
End Sub


Private Sub ConnDB()
On Error Resume Next

Dim strr As String
Set Cnn = New ADODB.Connection
With Cnn
    If Val(ServerName) > 0 Then             'ip dedected
        strr = "Provider=SQLOLEDB.1;Password=" & DatabasePassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & DbName & ";Data Source= " & ServerName & ",1433"
'        strr = "Provider=SQLOLEDB.1; Data Source=" & ServerName & ",1433; Network=dbmssocn; Initial Catalog=" & DbName & "; User ID=sa; Password=" & DatabasePassword & ";"
    Else
        strr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DbName & ";Data Source=" & ServerName
'        strr = "Provider=SQLOLEDB.1; Password=" & DatabasePassword & "; Persist Security Info=True; User ID=sa; Initial Catalog=" & DbName & "; Data Source= " & ServerName
    End If
    Cnn.Open strr
End With

Er:
'If Err.Number = -2147467259 Then
'    Message "Connection has been Lost. " & vbCrLf & "Please Re-Open the software"
'    End
'End If

End Sub



Public Sub ExecuteQuery(TxtQry As String)
On Error GoTo gans
Set txtQuery = New Command
With txtQuery
    .ActiveConnection = Cnn
    .CommandType = adCmdText
    .CommandText = TxtQry
    .Execute
End With
gans:
If Err.Number = -2147467259 Then
    Message "Connection has been Lost. " & vbCrLf & "Please Re-Open the software"
    End
End If
    
End Sub


Public Function EString(l As Single) As String
Dim tmP, stR As String
Dim I, j As Integer
tmP = Format(l, "0000000000000")
stR = ""
   
   If Len(tmP) > 13 Then
      EString = ""
      Exit Function
   End If
   
   If Val(tmP) = 0 Then
      EString = "Zero"
      Exit Function
   End If
   
   I = Val(Left$(tmP, 2))
   If I <> 0 Then
      GoSub do_hundreds
      stR = stR + " Kharba"
   End If
   
   I = Val(Mid$(tmP, 3, 2))
   If I <> 0 Then
      GoSub do_hundreds
      stR = stR + " Arba"
   End If
   
   I = Val(Mid$(tmP, 5, 2))
   If I <> 0 Then
      GoSub do_hundreds
      stR = stR + " Carod"
   End If
   
   I = Val(Mid$(tmP, 7, 2))
   If I <> 0 Then
      GoSub do_hundreds
      stR = stR + " Lakhs"
   End If

   I = Val(Mid$(tmP, 9, 2))
   If I <> 0 Then
      GoSub do_hundreds
      stR = stR + " Thousand"
   End If
   
   I = Val(Right$(tmP, 3))
   If I <> 0 Then
      GoSub do_hundreds
   End If
   
   EString = stR
   Exit Function

do_hundreds:
   If I > 99 Then
      j = I
      I = I \ 100
      GoSub do_ones
      stR = stR + " Hundred"
      I = j Mod 100
   End If

   If I <> 0 Then
      GoSub do_Tens
   End If
   Return
   
do_Tens:
   Select Case I Mod 100
      Case 90 To 99:
         stR = stR + " Ninety"
         GoSub do_ones
      Case 80 To 89:
         stR = stR + " Eighty"
         GoSub do_ones
      Case 70 To 79:
         stR = stR + " Seventy"
         GoSub do_ones
      Case 60 To 69:
         stR = stR + " Sixty"
         GoSub do_ones
      Case 50 To 59:
         stR = stR + " Fifty"
         GoSub do_ones
      Case 40 To 49:
         stR = stR + " Fourty"
         GoSub do_ones
      Case 30 To 39:
         stR = stR + " Thirty"
         GoSub do_ones
      Case 20 To 29:
         stR = stR + " Twenty"
         GoSub do_ones
         
      Case 19: stR = stR + " Nineteen"
      Case 18: stR = stR + " Eighteen"
      Case 17: stR = stR + " Seventeen"
      Case 16: stR = stR + " Sixteen"
      Case 15: stR = stR + " Fifteen"
      Case 14: stR = stR + " Fourteen"
      Case 13: stR = stR + " Thirteen"
      Case 12: stR = stR + " Twelve"
      Case 11: stR = stR + " Eleven"
      Case 10: stR = stR + " Ten"
      Case Else
         GoSub do_ones
   End Select
   Return
   
do_ones:
   If I < 10 Or I Mod 10 = 0 Then
      stR = stR + " "
   Else
      stR = stR + "-"
   End If
   Select Case I Mod 10
      Case 9: stR = stR + "Nine"
      Case 8: stR = stR + "Eight"
      Case 7: stR = stR + "Seven"
      Case 6: stR = stR + "Six"
      Case 5: stR = stR + "Five"
      Case 4: stR = stR + "Four"
      Case 3: stR = stR + "Three"
      Case 2: stR = stR + "Two"
      Case 1: stR = stR + "One"
   End Select
   Return
End Function

Public Function NString(l As Long) As String
Dim tmP, stR As String
Dim I, j As Integer
tmP = Format(l, "0000000000000")
stR = ""
   
   If Len(tmP) > 13 Then
      NString = ""
      Exit Function
   End If
   
   If Val(tmP) = 0 Then
      NString = "z'Go"
      Exit Function
   End If
   
   I = Val(Left$(tmP, 2))
   If I <> 0 Then
      GoSub My_Hundred
      stR = stR + " va{"
   End If
   
   I = Val(Mid$(tmP, 3, 2))
   If I <> 0 Then
      GoSub My_Hundred
      stR = stR + " ca{"
   End If
   
   I = Val(Mid$(tmP, 5, 2))
   If I <> 0 Then
      GoSub My_Hundred
      stR = stR + " s/f]8"
   End If
   
   I = Val(Mid$(tmP, 7, 2))
   If I <> 0 Then
      GoSub My_Hundred
      stR = stR + " nfv"
   End If

   I = Val(Mid$(tmP, 9, 2))
   If I <> 0 Then
      GoSub My_Hundred
      stR = stR + " xhf/"
   End If
   
   I = Val(Right$(tmP, 3))
   If I <> 0 Then
      GoSub My_Hundred
   End If
   
   NString = stR
   Exit Function

My_Hundred:
   If I > 99 Then
      j = I
      I = I \ 100
      GoSub My_One
      stR = stR + " ;o"
      I = j Mod 100
   End If

   If I <> 0 Then
      GoSub My_Ten
   End If
   Return
   
My_Ten:
   Select Case I Mod 100
      Case 99: stR = stR + " pgfG;o"
      Case 98: stR = stR + " cG7fgAa]"
      Case 97: stR = stR + " ;GtfgAa]"
      Case 96: stR = stR + " 5ofgAa]"
      Case 95: stR = stR + " k~rfgAa]"
      Case 94: stR = stR + " rf}/fgAa]"
      Case 93: stR = stR + " lqofgAa]"
      Case 92: stR = stR + " aofgAa]"
      Case 91: stR = stR + " PsfgAa]"
      Case 90: stR = stR + " gAa]"
      Case 89: stR = stR + " pgGgAa]"
      Case 88: stR = stR + " c7f;L"
      Case 87: stR = stR + " ;tf;L"
      Case 86: stR = stR + " 5of;L"
      Case 85: stR = stR + " krf;L"
      Case 84: stR = stR + " rf}/f;L"
      Case 83: stR = stR + " lqof;L"
      Case 82: stR = stR + " aof;L"
      Case 81: stR = stR + " Psf;L"
      Case 80: stR = stR + " c;L"
      Case 79: stR = stR + " pgf;L"
      Case 78: stR = stR + " c7Q/"
      Case 77: stR = stR + " ;tQ/"
      Case 76: stR = stR + " 5ofQ/"
      Case 75: stR = stR + " krxQ/"
      Case 74: stR = stR + " rf}/Q/"
      Case 73: stR = stR + " lqxQ/"
      Case 72: stR = stR + " axQ/"
      Case 71: stR = stR + " PsQ/"
      Case 70: stR = stR + " ;Q/L "
      Case 69: stR = stR + " pgG;Q/L"
      Case 68: stR = stR + " c8\;l7"
      Case 67: stR = stR + " ;8\;l7"
      Case 66: stR = stR + " 5};l7"
      Case 65: stR = stR + " k};l7"
      Case 64: stR = stR + " rf};l7"
      Case 63: stR = stR + " lq;l7"
      Case 62: stR = stR + " a};l7"
      Case 61: stR = stR + " Ps;l7"
      Case 60: stR = stR + " ;f7L"
      Case 59: stR = stR + " pG;f7L"
      Case 58: stR = stR + " cG7fpGg"
      Case 57: stR = stR + " ;GtfpGg"
      Case 56: stR = stR + " 5kGg"
      Case 55: stR = stR + " krkGg"
      Case 54: stR = stR + " rF]}Gg"
      Case 53: stR = stR + " lqkGg"
      Case 52: stR = stR + " afpGg"
      Case 51: stR = stR + " PsfpGg"
      Case 50: stR = stR + " krf;"
      Case 49: stR = stR + " pgGkrf;"
      Case 48: stR = stR + " c8\rfln;"
      Case 47: stR = stR + " ;Rrfln;"
      Case 46: stR = stR + " 5ofln;"
      Case 45: stR = stR + " k}rfln;"
      Case 44: stR = stR + " rf}jfln;"
      Case 43: stR = stR + " lqrfln;"
      Case 42: stR = stR + " aofln;"
      Case 41: stR = stR + " PSrfln;"
      Case 40: stR = stR + " rfln;"
      Case 39: stR = stR + " pgGrfln;"
      Case 38: stR = stR + " c8\lt;"
      Case 37: stR = stR + " ;}lt;"
      Case 36: stR = stR + " 5lQ;"
      Case 35: stR = stR + " k}lt;"
      Case 34: stR = stR + " r}flt;"
      Case 33: stR = stR + " t]lQ;"
      Case 32: stR = stR + " alQ;"
      Case 31: stR = stR + " Pslt;"
      Case 30: stR = stR + " lt;"
      Case 29: stR = stR + " pglGt;"
      Case 28: stR = stR + " cfO{;"
      Case 27: stR = stR + " ;QfO{;"
      Case 26: stR = stR + " 5lAa;"
      Case 25: stR = stR + " klRr;"
      Case 24: stR = stR + " rf}la;"
      Case 23: stR = stR + " t]O{;"
      Case 22: stR = stR + " afO{;"
      Case 21: stR = stR + " PSsfO{;"
      Case 20: stR = stR + " la;"
      Case 19: stR = stR + " pGgfO{;"
      Case 18: stR = stR + " c7f/"
      Case 17: stR = stR + " ;q"
      Case 16: stR = stR + " ;f]x|"
      Case 15: stR = stR + " kGw|"
      Case 14: stR = stR + " rf}w"
      Case 13: stR = stR + " t]x|"
      Case 12: stR = stR + " afx|"
      Case 11: stR = stR + " P3f/"
      Case 10: stR = stR + " bz"
      Case Else
         GoSub My_One
   End Select
   Return
My_One:
   If I < 10 Or I Mod 10 = 0 Then
      stR = stR + " "
   Else
      stR = stR + " "
   End If
   Select Case I Mod 10
      Case 9: stR = stR + "gf}"
      Case 8: stR = stR + "cf7"
      Case 7: stR = stR + ";ft"
      Case 6: stR = stR + "5"
      Case 5: stR = stR + "kf+r"
      Case 4: stR = stR + "rf/"
      Case 3: stR = stR + "ltg"
      Case 2: stR = stR + "b'O{"
      Case 1: stR = stR + "Ps"
   End Select
   Return
End Function
Public Sub Refress_Rs(myrs As Recordset, Refress_Query As String)
On Error GoTo X
    If myrs.State = adStateOpen Then
        myrs.Close
    End If
    myrs.CursorLocation = adUseClient
    myrs.Open Refress_Query, Cnn, adOpenKeyset, adLockOptimistic, 1
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
    End If
    
X:
If Err.Number = -2147467259 Then
    Message "Connection has been Lost. " & vbCrLf & "Please Re-Open the software"
    End
End If
End Sub


Public Sub Colored()
On Error Resume Next
Dim X As TextBox
'If  (Screen.ActiveControl) = x Then
'If Screen.ActiveControl = True Then
    Set X = Screen.ActiveControl
    X.BackColor = &HFFFF80
    X.SelStart = 0
    X.SelLength = Len(X)
'End If
End Sub

Public Sub unColored(ctrl As TextBox)
    ctrl.BackColor = &H80000005
End Sub

Public Function NewMaxID(TableName As String, ColumnName As String, Optional WhereCondition As String)
Dim rsMax As Recordset
Set rsMax = New ADODB.Recordset
Refress_Rs rsMax, "Select max(" & ColumnName & ") as MaxID from " & TableName & " " & WhereCondition
    With rsMax
        If .BOF = True And .EOF = True Or IsNull(!MaxID) Then
            NewMaxID = 1
        Else
            NewMaxID = Val(!MaxID) + 1
        End If
    End With
End Function



Public Sub ExportToExcelFromRecordSet(myrs As Recordset, Heading As YesNoOption, Footer As YesNoOption, Optional HeadingNote As String)

If Not myrs.RecordCount > 0 Then
    Message "Report is not prepaired Perporely."
    Exit Sub
End If


Dim I, j As Single
Dim row As Integer, Rng As String
Dim wbook As Workbook, wsheet As Worksheet
'---------------------------------------------------------------------------------------
Set wbook = Excel.Workbooks.Add
Set wsheet = wbook.Sheets(1)
wsheet.Cells.Clear
Excel.Application.Visible = True
'---------------------------------------------------------------------------------------
If Heading = myYes Then
    row = 4
Else
    row = 1
End If
'wsheet.Cells.Font.Size = 12
For I = 0 To myrs.Fields.Count - 1
    wsheet.Cells(row, I + 1) = myrs.Fields(I).Name
Next

'row = row + 1
myrs.MoveFirst
For I = 1 To myrs.RecordCount
    row = row + 1
    
    For j = 0 To myrs.Fields.Count - 1
        If IsDate(myrs.Fields(j).Value) Or Not IsNumeric(myrs.Fields(j).Value) Then
            wsheet.Cells(row, j + 1) = "'" & myrs.Fields(j).Value
        Else
'            wsheet.Cells(row, j + 1).Format = xlNumber
            wsheet.Cells(row, j + 1) = myrs.Fields(j).Value
            
        End If
    Next
    
    myrs.MoveNext
Next
'wsheet.Cells.


If Heading = myYes Then
    Rng = "a1:" & Chr(64 + j) & "1"
    Range(Rng).Select
    Selection.Cells.MergeCells = True
    Selection.Cells(1, 1) = Org
    Selection.Cells.Font.Size = 15
    Selection.Cells.VerticalAlignment = vbCenter
    Selection.Cells.HorizontalAlignment = xlHAlignCenter
    
    
    Rng = "a2:" & Chr(64 + j) & "2"
    Range(Rng).Select
    Selection.Cells.MergeCells = True
    Selection.Cells(1, 1) = CurrenBranchName
    Selection.Cells.Font.Size = 15
    Selection.Cells.VerticalAlignment = vbCenter
    Selection.Cells.HorizontalAlignment = xlHAlignCenter
    
    
    Rng = "a3:" & Chr(64 + j) & "3"
    Range(Rng).Select
    Selection.Cells.MergeCells = True
    Selection.Cells(1, 1) = HeadingNote
    Selection.Cells.Font.Size = 11
    Selection.Cells.VerticalAlignment = vbCenter
    Selection.Cells.HorizontalAlignment = xlHAlignCenter
    
    
    
End If


Rng = "a" & row - myrs.RecordCount & ":" & Chr(64 + j) & row - myrs.RecordCount
Range(Rng).Select
Selection.Font.Bold = True
Selection.Font.Italic = True
Selection.Cells.VerticalAlignment = vbCenter
Selection.Cells.HorizontalAlignment = xlHAlignCenter



Rng = "a1:" & Chr(64 + j) & "1" + Trim(stR(row - 1))
Range(Rng).Select
Selection.Columns.AutoFit

If Footer = myYes Then
    row = row + 1
    myrs.MoveFirst
    For j = 0 To myrs.Fields.Count - 1
        If IsNumeric(myrs.Fields(j).Value) Or IsNull(myrs.Fields(j).Value) Then
            wsheet.Cells(row, j + 1) = "=sum(" & Chr(64 + 1 + j) & row - myrs.RecordCount & ":" & Chr(64 + 1 + j) & row - 1 & ")"
        End If
    Next


End If

Rng = "a" & row - myrs.RecordCount & ":" & Chr(64 + j) & row
Range(Rng).Select
Selection.Cells.Borders.Value = 1


wsheet.PageSetup.CenterHorizontally = True

wsheet.PageSetup.LeftFooter = "Prepair By"
wsheet.PageSetup.CenterFooter = "Checked By"
wsheet.PageSetup.RightFooter = "Approved By"

wsheet.PageSetup.RightHeader = "Print on " & CurDate & " at " & Time
wsheet.Cells.NumberFormatLocal = "General"

End Sub






Private Sub setOrg()
Dim rsGloal As Recordset
Set rsGloal = New Recordset
Refress_Rs rsGloal, "Select * from tblMSTR"
If rsGloal!MSTRID = "1" Then
    OrgID = 1
    Org = "Kisan Saving And Credit Co-Operative Ltd."
    Orgadd = "Gaindakot-8, Nawalparasi"
End If
End Sub

Public Function Message(Desc As String, Optional buttonType As MessageBtn = okOnly, Optional boolFocusOnYes As Boolean) As MessageResponse
With frmMsg
    .lblMsg.Caption = Desc
    .picOption.Visible = True
    .cmdOk.Visible = True
    .Caption = "(Message Board)"
    
    If buttonType = okOnly Then
        .picOption.Visible = False
    Else
        If boolFocusOnYes Then
            .cmdyes.TabIndex = 0
        Else
            .cmdNo.TabIndex = 0
        End If
        .cmdOk.Visible = False
    End If
    .Show vbModal
End With
End Function


Public Function NepDateTOEng(NepDate As String)
Dim X As String
Dim rsD As Recordset
Set rsD = New Recordset

X = NepDate
Refress_Rs rsD, "Select * from tbldates where mid(DateN,1,4) = " & Year(X) & " and mid(DateN,6,2) = " & Format(Month(X), "00") & " and mid(DateN,9,2) = " & Format(Day(X), "00")
If rsD.RecordCount > 0 Then
    NepDateTOEng = rsD!DateE
Else
    Message "Error while converting date .. . . . "
    frmCreateDate.Show vbModal
End If
End Function

Public Function EngDateTONep(EngDate As String)
Dim X As String
Dim rsD As Recordset
Set rsD = New Recordset

X = EngDate
Refress_Rs rsD, "Select * from tbldates where year(DateE) = " & Year(X) & " and month(DateE) = " & Month(X) & " and day(DateE) = " & Day(X)
If rsD.RecordCount > 0 Then
    EngDateTONep = rsD!DateN
Else
    Message "Error while converting date .. . . . "
    frmCreateDate.Show vbModal
End If
End Function

Public Function GetMasantti(thisDate As String)
Dim rsMasanta As Recordset
Set rsMasanta = New Recordset

Refress_Rs rsMasanta, "Select top 1 * from tbldates where substring(dateN,1,8) = '" & Mid(thisDate, 1, 8) & "' order by DateN Desc "
If rsMasanta.RecordCount > 0 Then
    GetMasantti = rsMasanta!DateN
End If
End Function


Public Function NumirecToCharecter(Source)
Dim rval As String
rval = Source

rval = Replace(rval, "0", ")")
rval = Replace(rval, "1", "!")
rval = Replace(rval, "2", "@")
rval = Replace(rval, "3", "#")
rval = Replace(rval, "4", "$")
rval = Replace(rval, "5", "%")
rval = Replace(rval, "6", "^")
rval = Replace(rval, "7", "&")
rval = Replace(rval, "8", "*")
rval = Replace(rval, "9", "(")

NumirecToCharecter = rval
End Function


Public Sub tojson(thisRs As Recordset)
    savename = "exportedxls.json"
    lcolumn = thisRs.Fields.Count
    lrow = thisRs.RecordCount
    Dim titles() As String
    ReDim titles(lcolumn)
    For I = 0 To lcolumn - 1
        titles(I) = thisRs.Fields(I).Name
    Next I
    json = "[" & vbNewLine
    dq = """"
    For j = 1 To lrow
        For I = 0 To lcolumn - 1
            If I = 0 Then
                json = json & "{" & vbNewLine
            End If
            CellValue = thisRs.Fields(I)
            json = json & dq & titles(I) & dq & ":" & dq & CellValue & dq
            If I <> lcolumn Then
                json = json & "," & vbNewLine
            End If
        Next I
        json = json & "}"
        If j <> lrow Then
            json = json & ","
        End If
        json = json & vbNewLine
        thisRs.MoveNext
    Next j
    json = json & "]"
    myFile = Application.DefaultFilePath & "\" & savename
    Open myFile For Output As #1
    Print #1, json
    Close #1
'    a = MsgBox("Saved as " & savename, vbOKOnly)
End Sub

Public Sub PupupMessage(msg As String, sender As String)
Dim a As Integer

Dim X As Form
Set X = New frmMessage
    With X
        .lblTitle = sender
        .lblMessage = msg
        a = Val(GetSetting(App.Title, App.Title, "cnt", 0)) + 1
        SaveSetting App.Title, App.Title, "cnt", a
        .Left = Screen.Width - ((.Width + 10) * a)
        .Top = Screen.Height - .Height + 20
        MakeTopmost X, True
        .Show
    

    
    End With

End Sub

Function MakeTopmost(Frm As Form, Yes As Boolean) ' Always On Top
On Error GoTo D
    If Yes = True Then
        SetWindowPos Frm.hwnd, -1, 0, 0, 0, 0, 1 Or 2
    Exit Function
    End If

    If Yes = False Then
        SetWindowPos Frm.hwnd, -2, 0, 0, 0, 0, 1 Or 2
    Exit Function
    End If
D:
Err.Clear
End Function

Public Sub GenerateListView(lst As MSComctlLib.ListView, data As Recordset)
If Not data.RecordCount > 0 Then
    lst.ListItems.Clear
    Exit Sub
End If

data.MoveFirst

With lst
    
    
    .View = lvwReport
    .Gridlines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual
    .Appearance = ccFlat
    
    
    .ListItems.Clear
    a = 0
    For a = 1 To .ColumnHeaders.Count
        .ColumnHeaders.Remove (1)
    Next
    
    
    a = 0
    For a = 0 To data.Fields.Count - 1
        .ColumnHeaders.Add
        .ColumnHeaders(.ColumnHeaders.Count).Text = data.Fields(a).Name
        If a > 0 Then
            If IsNumeric(data.Fields(a).Value) Then
                .ColumnHeaders(.ColumnHeaders.Count).Alignment = lvwColumnRight
            End If
        End If
    Next
    
    
    While Not data.EOF
        .ListItems.Add data.AbsolutePosition, , Val(data.Fields(0) & "")
        For a = 0 To .ColumnHeaders.Count - 1
            If a = .ColumnHeaders.Count - 1 Then
                .ListItems(data.AbsolutePosition).SubItems(a) = IIf(IsNull(data.Fields(a)), "", data.Fields(a))
            Else
                .ListItems(data.AbsolutePosition).SubItems(a + 1) = IIf(IsNull(data.Fields(a + 1)), "", data.Fields(a + 1))
            End If
        Next
        
        data.MoveNext
    Wend
    .SelectedItem.EnsureVisible
End With

End Sub


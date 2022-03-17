Create Database KpiReport

Create Table tblUserType
(
SN int primary key,
TypeName varchar(10)
)

Select * from tblusertype
 insert into tblusertype values(1,'Admin')
 insert into tblusertype values(2,'Inchage')
 insert into tblusertype values(3,'Staff')
update tblusertype set typename='Other' where sn = 3


Create table tblServiceCenter
(
Sn int primary key,
Code varchar(10),
ServiceCenterName varchar(40),
Address varchar(40),
Phone varchar(40),
Incharge bigint
)


insert into tblServicecenter values(1,'001','Head Office','Gaindakot-8, Bijayachowk','078-501622',0)
insert into tblServicecenter values(2,'002','Himchuli Service Center','Gaindakot-2, Bijayachowk','078-501622',0)
insert into tblServicecenter values(3,'003','Bijayanagar Service Center','Gaindakot-2, Bijayachowk','078-501622',0)



Create table tblUsers
(
SN bigint Primary key,
UserID varchar(15) unique,
userpassword varchar(15),
UserFullName varchar(50),
Address varchar(50),
Phone varchar(30),
userType int references tblUsertype(SN),
Branch int references tblserviceCenter(SN),
status bit -- 0 as active, 1 as Passive
)

alter table tblusers add Post varchar(50)
alter table tblusers add DobN varchar(10)

Select * from tblusers



Insert into tblusers values(1,'Nabin','nabin','Nawaraj Sapkota (Nabin)','Gaindakot-8, Nawalpasai','9801354155',1,1,1,null,null)
Insert into tblusers values(2,'Chandra','chandra','Chandra Sunari Magar','Gaindakot-8, Nawalpasai','9801354155',2,1,0,null,null)
Insert into tblusers values(3,'Manoj','manoj','Manoj Pandey','Gaindakot-8, Nawalpasai','9801354155',3,1,0,null,null)
Insert into tblusers values(4,'purna','purna','Purna Magar','Gaindakot-8, Nawalpasai','9801354155',3,1,0,null,null)

update tblUsers set Post='Incharge' where sn=2
update tblUsers set Post='Saving Assestent' where sn=3
update tblUsers set Post='MR' where sn=4




drop View ViewUsers
Create View viewUsers 
as
Select u.*,ut.TypeName, sc.Code, sc.ServiceCenterName,sc.Address ScAddress, sc.Phone SCPhone, sc.Incharge
from tblusers u
inner join tbluserType ut on (u.userType=ut.sn)
inner join tblServiceCenter sc on (u.Branch=sc.sn)




Create View viewServiceCenter
as
Select sc.*, u.userID, u.Phone as UserPhone, u.UserFullName from tblServiceCenter sc 
inner join tblusers u on u.sn=sc.Incharge

Create table tblTask
(
sn int primary key,
TaskName varchar(40),
Proitity int
)

insert into tblTask values (1,'Share',1)
insert into tblTask values (2,'Saving',2)
insert into tblTask values (3,'Field Collection',3)
insert into tblTask values (4,'New member',4)
insert into tblTask values (5,'MMicro Group',5)
insert into tblTask values (6,'Programme',6)

Drop table tblGivenTask
Create table tblGivenTask
(
sn bigint primary key,
TaskID int references tbltask(sn),
taskBy bigint references tblusers(SN),
TaskTo int references tblserviceCenter(SN),
Target money,
yearr int,
Monthh int,
BRICH int
)

alter table tblGivenTask add BRICH int


Create table tblTaskAchive
(
sn bigint primary key,
GTSN bigint references tblGivenTask(SN),
Achive money,
dated varchar(10),
Approved bit
)

alter table tblTaskAchive add dated varchar(10)
alter table tblTaskAchive add Approved bit

drop View ViewGivenTask
Create View ViewGivenTask
as
Select gt.*, sc.Code, ServiceCenterName, sc.Address, sc.Phone, incharge, taskName, Proitity, ToUser = case when Brich = 0 then incharge else brich end from 
tblGivenTask gt 
inner join tblServiceCenter SC on (gt.TaskTo=sc.sn)
inner join tblTask t on (gt.TaskID=t.sn)

Drop View ViewTaskAchive

Create View viewTaskAchive
as
Select *, isnull((Select sum(Achive) from tblTaskAchive where Approved=1 and GTSN=gt.sn),0) Achive
 , isnull((Select userID from tblusers where sn=gt.ToUser),0) userID
 , isnull((Select userFullName from tblusers where sn=gt.ToUser),0) userName
from ViewgivenTask gt


Drop View viewTaskAchiveDets
Create View viewTaskAchiveDets
as
Select gt.*, ta.achive, ta.dated, ta.Approved, ta.SN AchiveSN
from ViewgivenTask gt
inner join tblTaskAchive TA on (ta.GTSN=gt.SN)




Create table tblInchages
(
sn bigint primary key,
Branch int references tblServiceCenter(SN),
users bigint references tblusers(SN),
status int
)




Select * from ViewUsers where Status = 0 and Branch = 1 and userType<>1 

Select * from tblServiceCenter

Select * from tblGiventask
Select * from ViewtaskAchive where yearr=2076 and monthh=5 and brich=3 and taskID=1
Select * from tblTaskAchive


Select DateD, Achive, Approved from ViewTaskAchive where yearr=2076 and monthh=5 and brich = 3 and TaskID = 1


Select DateD, Achive, Status = case Approved when 0 then 'Not-Approved' else 'Approved' end, AchiveSN
from viewTaskAchiveDets where yearr=2076 and monthh=5 and brich = 3 and TaskID = 1

Select  * from viewTaskAchiveDets
Select * from ViewUsers

Create table tblDates 
(
DateE varchar(10) primary key,
DateN varchar(10) unique
)


Select * from tbldates where year(DateE) = 2019 and month(DateE) = 9 and day(DateE) = 2
Select * from tblgiventask where taskID = 1 and brich=0 and yearr=2076 and monthh=5

Select * from tbltaskAchive

select * from tbltask


select * from tblusers
select * from Viewusers

Delete from tblgivenTask
Delete from tblTaskAchive
Delete from tblusers



Select * from ViewTaskAchive where yearr=2076 and monthh = 5 and brich = 6 
--delete from tblGiventask where sn = 12

select * from tbltaskachive

update tbltaskAchive set approved=1 where approved=0

Select * from ViewTaskAchive

Select DateD, TaskName, Achive, Status = case Approved when 0 then 'Not-Approved' else 'Approved' end, AchiveSN, * from viewTaskAchiveDets 
where Approved = 0 and yearr=2076 and monthh=5 and brich = 6 
order by Proitity


Delete from tblTaskAchive
Delete from tblgivenTask

Select TaskName, Target, Achive, Achive-target Difference, Achive/Target*100 Percentage,  SN, * from viewTaskAchive 
where yearr = 2076 and monthh = 5 and taskTo = 1 and brich = 3



Select ServiceCenterName, Target, Achive, TaskName from ViewTaskAchive where brich=incharge and taskID=6 



Select ServiceCenterName, Target, Achive, TaskName 
, (Select Count(Code) from ViewTaskAchive where yearr=2076 and monthh=5 and brich=incharge and Code=d.Code) Count
from ViewTaskAchive d
where yearr=2076 and monthh = 5 and brich = Incharge 
order by code, Proitity

drop table tblMSTR
Create table tblMSTR
(MSTRID int primary key)

insert into tblMSTR values(1)



Select TaskName, Target, Achive, Achive-target Difference, Achive*100/(Target+1) Percentage,  SN from viewTaskAchive where yearr = 2076 and monthh = 5 and taskTo = 1 and toUser = 3



-----------------------------------------------------------------------------------
drop table tbl5sCategories
Create table tbl5sCategories
(Sn int primary key,
catname varchar(50)
)

insert into tbl5sCategories values(1,'S1')
insert into tbl5sCategories values(2,'S2')
insert into tbl5sCategories values(3,'S3')
insert into tbl5sCategories values(4,'S4')
insert into tbl5sCategories values(5,'S5')

Select * from tbl5sCategories
update tbl5sCategories set catname ='S1-Sort' where sn = 1
update tbl5sCategories set catname ='S2-Set In Order' where sn = 2
update tbl5sCategories set catname ='S3-Shine' where sn = 3
update tbl5sCategories set catname ='S4-Standaridize' where sn = 4
update tbl5sCategories set catname ='S5-Sustain' where sn = 5


drop table tbl5SQuestion

Create table tbl5SQuestion
(sn int primary key,
Qstname varchar(500),
catID int references tbl5SCategories(SN),
Point varchar(40)
)


insert into tbl5SQuestion values(1,'oxfsf] cj:yf s:tf] % <',1)

Drop View view5sQuestion

Create View view5sQuestion
as
Select c.*, q.SN QSTID, QstName, Point from tbl5sCategories c 
inner join tbl5sQuestion q on(q.CatID=c.sn)

drop table tbl5SAnsList

Create table tbl5SAnsList
(sn int primary key,
AnsID int,
Ansname varchar(40),
QstID int references tbl5SQuestion(SN) unique(QstID,AnsID)
)

Delete from tbl5sAnsList


select * from tbl5SAnsList where QstID=7



-- update tbl5SAnsList set ansname='c;n' where ansname='l&&s}'


drop view view5sAnsList
Create View view5sAnsList
as
Select vq.*, al.sn AnsSN, al.ansID, al.ansname from View5sQuestion vq 
inner join tbl5sAnsList al on (vq.qstID=al.qstID)

drop table tbl5SAnswer

Create table tbl5SAnswer
(SN bigint primary key,
QstID int references tbl5SQuestion(SN),
Ans int references tbl5SAnsList(sn),
monthh int,
yearr int,
userno bigint references tblusers(SN) unique(monthh,yearr,userno,QstID),
branch int references tblServiceCenter(SN),
isadmin bit,
dateD varchar(10)
)

Select * from tbl5SAnswer

sp_password 'aa',null
drop view view5sData
create View view5sData
as
Select a.*, vq.qstname,vq.catname,vq.sn catID, al.ansName,al.ansID, userid, userFullName, usertype, Post, code, servicecentername from View5sQuestion vq
inner join tbl5sAnswer a on (a.QstID=vq.QstID)
inner join tbl5sansList al on (al.sn=a.ans)
inner join tblusers u on (u.sn=a.userno)
inner join tblserviceCenter sc on (sc.sn=a.Branch)



Select * from View5sData d 
where yearr=2076 and monthh=5 and userno=3 and branch=1
--group by catID, catname, userno, yearr, monthh, branch

-- union
-- Select CatID, CatName, QstID, QstName, ans, userno, isadmin from View5sData where yearr=2076 and monthh=5 and isadmin=1 and branch=1
-- order by isadmin

Select * from View5sData where yearr=2076 and monthh=5 and userno=2 and branch=1

Select * from tbl5SQuestion

Select * from tblserviceCenter

 insert into tbl5SQuestion values(1, ';kmf 6]jn, cfjZos sfuhft afx]s cGo s]lx j:t'' gePsf] / sfof{no ;do kZrft s''g} klg ;fdfu|L 6]jndf /fVg] gu/]sf] .', 1, 'clkm; 6]jn')
 insert into tbl5SQuestion values(2, ';kmf 3/f{ / 3/f{df /fv]sf ;fdfu|Lx? ldnfP/ /fVg] u/]sf]', 1, '6]jnsf] 3/f{')
 insert into tbl5SQuestion values(3, ';kmf, cgfjZos ;fdfu|L oqtq /fVg] gu/]sf], cgfjZos sfuhkq Roft]/ e''FO{df gkmfnL 8l:jgdf /fVg] u/]sf], leQfdf s''g]} kmf]xf]/ gnfu]sf], lglZrt 7Ffpdf ;"rgf 6fF; ug]{ u/]sf], sfpG6/df s]/d]6 tyf w''nf] gePsf], laB''t tyf pks/0fsf tf/x? c:tJo:t gePsf]', 1, 'sfo{sIf')
 insert into tbl5SQuestion values(4, ';Dk''0f{ kmfO{nx? Jojl:yt 9+uaf6 kmfO{ln· u/]sf], cgfjZos kmfO{nx? Sofljg]6df /fVg] gu/]sf]', 1, 'kmfO{n Soflag]6')
 insert into tbl5SQuestion values(5, 'uf]]Zjf/f ef}r/, td;''''s, ekf{O{, z]o/ ;b:otf kmf/d, C0f dfu kmf/d cflb sfuh kq cfsif{s 9+uaf6 kmfO{ln· ug]{ u/]sf] .', 1, 'kmfO{n kmf]N8/')
 insert into tbl5SQuestion values(6, 'kfgL 6]jndf /fv]/ vfg] u/]sf] t/ cGo lrof gf:tf 6]andf /fv]/ vfg] gu/]sf] .', 1, 'lrof÷gf:tf')
 insert into tbl5SQuestion values(7, 'pks/0f h8fg ubf{ Jojl:yt 9+uaf6 h8fg ug]{ u/]sf] / cfjZostf cg'';f/ vf]Ng / aGb ug{ ;lsg] u/]/ /fv]sf] .', 2, 'pks/0f h8fg')
 insert into tbl5SQuestion values(8, 'sfd gePsf] cj:yfdf cyjf sfof{no aGb ubf{ ljB''t sg]S;g ckm ug]{ u/]sf] .', 2, 'alQ÷k+vf')
 insert into tbl5SQuestion values(9, 'sDKo"6/sf 8s''d]G6 Jojl:yt 9+uaf6 ;''/IfLt tj/af6 /fv]sf], s''g} 8s''d]G6 vf]hL ubf{ a9Ldf # ldg]6 eGbf a9L ;do gnfUg] u/]sf] .', 2, 'sDKo''6/ kmfO{ln·')
 insert into tbl5SQuestion values(10, 'k'':ts tyf cGo cWoog ;fdfu|Lx? cWog u/L ;dod} oyf:yfgdf /fVg] u/]sf] .', 2, 'cWoog ;fdfu|L')
 insert into tbl5SQuestion values(11, ';kmf rlDsnf] / uGw /lxt .', 3, 'clkm; kmf]g')
 insert into tbl5SQuestion values(12, 'sDKo''6/ tyf sDKo''6/sf] ls af]8{ / l:s[g ;w} ;kmf / rlDsnf] /fVg] u/]sf]', 3, 'clkm; sDKo''6/')
 insert into tbl5SQuestion values(13, ';kmf rlDsnf], :qmofr gePsf] tyf sfk]{6 ldnfP/ /fv]sf] .', 3, 'sfo{sIfsf] e''+O{')
 insert into tbl5SQuestion values(14, '6\jfO{n]6 l;6 ;kmf / kfgLsf] Joj:yf ePsf]', 3, '6\jfO{n]6')
 insert into tbl5SQuestion values(15, ';+:yfn] tf]s]sf] kf]zfs ;kmf / zl//df ldnfP/ nufpg] u/]sf] ', 4, 'kf]zfs')
 insert into tbl5SQuestion values(16, 'cfs{ifs 9+uaf6 k|:t''tL ug]{ Ifdtf ePsf] .', 4, 'sd{rf/Lsf] k|:t''lt')
 insert into tbl5SQuestion values(17, 'sfo{nodf w''nf]af6 aRg laleGg lalw k|of]u ug]{ u/]sf]', 4, 'w''nf] d}nf]af6 ;''/Iff')
 insert into tbl5SQuestion values(18, 's''g} laifodf 6fO{k u/L tTsfn lk|G6 ug{ ;Sg] bIftf ePsf] .', 5, 'lkG6L· bIftf')
 insert into tbl5SQuestion values(19, ';a} sd{rf/Lx?n] sfof{no ;kmfO{df Wofg lbg] u/]sf] .', 5, ';/;kmfO{sf] jfgL')
 insert into tbl5SQuestion values(20, ';+:yfut gLtL lgod / k|rlnt P]g sfg"g cWoog ug]{ u/]sf] .', 5, 'gLlt lgodsf] 1fg')











delete from tbl5SAnswer where userno=3

sp_help tbl5SAnswer
insert into tbl5SAnswer values(2, 1, 7, 5, 2076, 3, 1, 0, '2076/05/22' )

Select Code, QstID, Qstname
, (select sum(ansID)/count(*) from View5sData where branch=d.branch and isadmin=d.isadmin and yearr=d.yearr and monthh=d.monthh and qstID=d.QstID) val
from view5sData d
where Branch=1 and isadmin=0 and yearr=2076 and monthh=5
group by code, qstID, QstName, branch, isadmin, yearr, monthh





---------------------------------------------------------------------------------------------------------------------------------------------------------


Create table tblMFTask
(sn bigint primary key,
taskID int,
taskname varchar(50),
countable bit
)

-- insert into tblmftask values(1,1,'New Member',0)
-- insert into tblmftask values(2,2,'Old Member',0)
-- insert into tblmftask values(3,3,'Other Member',0)
-- insert into tblmftask values(4,4,'Comp Saving',1)
-- insert into tblmftask values(5,5,'Emergency Saving',1)
-- insert into tblmftask values(6,6,'Nari Saving',1)
-- insert into tblmftask values(7,7,'Fund Saving',1)
-- insert into tblmftask values(8,8,'Other Saving',1)
-- insert into tblmftask values(9,9,'Loan',1)
-- insert into tblmftask values(10,10,'Share',1)
-- insert into tblmftask values(11,11,'Opt',1)


drop table tblMFGroup

Create table tblMFGroup
( sn bigint primary key,
groupname varchar(100),
groupAddress varchar(100),
branch int references tblserviceCenter(SN),
dayCode int
)


insert into tblmfGroup values(1,'Kisan','Gaindakot',1,1)
insert into tblmfGroup values(2,'Laligurans','Gaindakot',1,1)
insert into tblmfGroup values(3,'Bihani','Gaindakot',1,1)
insert into tblmfGroup values(4,'Jaintar','Gaindakot',1,1)
insert into tblmfGroup values(5,'Yakal','Gaindakot',1,1)

drop table tblMFData
Create table tblMFData
(SN bigint primary key,
userno bigint references tblusers(SN),
data money,
taskID bigint references tblMFtask(SN),
MFGrp bigint references tblMFGroup(SN),
dated varchar(10),
posted bit
)

drop View ViewMFData
Create View ViewMFData
as
Select mf.*, mfg.groupname, mfg.groupAddress, mfg.Branch, mfg.dayCode, mft.taskname, mft.countable
, u.userid, u.userFullName, sc.ServiceCenterName, sc.code 
from tblmfData mf 
inner join tblMFGroup mfg on (mf.mfgrp=mfg.sn)
inner join tblMFTask mft on (mf.TaskID=mft.sn)
inner join tblusers u on (mf.userno=u.sn)
inner join tblServiceCenter sc on (mfg.branch=sc.sn)


Select taskID, TaskColID, TaskName, countable, Data from ViewmfData where mfgrp = 2 and substring(Dated,1,8)= '2076/05/'

Delete from tblmfdata

Select Posted, sum(Data) data from ViewMFData where Countable = 1 and userno = 13 and MFGRP = 1 and substring(DateD,1,8) = '2076/05/' group by Posted order by Posted



Select * from tblMFGroup where branch = 15 and sn not in (Select distinct mfGrp from tblMFData where substring(Dated,1,8)='2076/05/' and posted=1)






Create table tblLoanData
(
SN bigint primary key,
CID bigint,
Branch int references tblServiceCenter(SN),
userno bigint references tblusers(SN),
CName varchar(70),
address varchar(70),
Phone varchar(50),
LoanAc varchar(20),
Ltd varchar(20),
lng varchar(20),
type int,
Nextdate varchar(10),
dated varchar(10),
Remarks varchar(500)
)

alter table tblLoanData add Status bit


Select * from tblLoanData


-- Select * from sysobjects where xtype='u'
-- Delete from tblTaskAchive
-- Delete from tblGivenTask
-- Delete from tbl5SAnswer
-- Delete from tblMFData
-- Delete from tblLoanData
-- 





-- Select * from tblMFGroup
-- Delete from tblMFGroup



-- Delete from tbl5sAnswer


Select * from tbl5sAnswer









Select * from ViewMFdata where mfgrp=15

Delete from tblMFData










Select * from tblmfGroup where branch = 1 and sn not in(Select distinct MFGrp from tblMFData where Posted=1 and substring(dated,1,8) = '2076/05/') order by DayCode 


Select distinct MFGrp from tblMFData where Posted=0 and substring(dated,1,8) = '2076/05/'

Select * from tblgivenTask
Select * from tblTaskAchive

update tblTaskAchive set Achive = 500 where gtsn = 

Select * from ViewMFData

Select * from ViewUsers

wish you a 

Select * from tblDates where substring(dateN,6,2)='05' and dateN>='2076/05/27'





Create table tblMessageCenter
(
sn bigint primary key,
frmuser bigint references tblusers(SN),
tousers bigint references tblusers(SN),
msgText varchar(500),
msgtime varchar(15),
msgDate varchar(10),
status int
)



Create table tblnotificationcenter
(
SN bigint primary key,
msg varchar(500),
userno bigint references tblusers(SN),
branch int references tblserviceCenter(SN),
dated varchar(10),
timee varchar(10),
status int
)


Create table tblFeedback
(
SN bigint primary key,
userno bigint references tblusers(SN),
title varchar(100),
msg varchar(500),
status int
)



Delete from tblMessageCenter



Select * from tblMessageCenter where status < 2 and frmuser = 3 and tousers = 2 or ( frmuser = 2 and tousers = 3 ) order by sn 

Select * from tblMessageCenter where status < 2 and frmuser = 0 and tousers = 2 or ( frmuser = 2 and tousers = 0 ) order by sn 


select * from tblsetting

Create table tblsetting 
(
sn int primary key,
Code varchar(50),
value varchar(50)
)

insert into tblSetting values(1,'System Date','2076/06/09')



Select * from sysobjects where xtype='u' and name = 'tblsetting'





--RESTORE DATABASE [KPIReport] FROM  DISK = N'D:\IT DOWNLOAD\kPI\kpi.bak' 
WITH  FILE = 1,  
MOVE N'KpiReport' TO N'C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\KPIReport.mdf',  
MOVE N'KpiReport_log' TO N'C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\KPIReport.LDF',  
NOUNLOAD,  REPLACE,  STATS = 10
--GO




BACKUP DATABASE [KPIReport] TO  DISK = N'D:\IT DOWNLOAD\kPI\kpi.bak' WITH NOFORMAT, NOINIT,  
NAME = N'KPIReport-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10


select * from tblMFData where dated<>'2075/04/01'


Select * from tblsetting

update tblSetting set value='0.0.10' where sn=2

select * from tblgiventask








Create table tblLoginLog
(
sn bigint primary key,
Dated varchar(10),
timee varchar(20),
userNo bigint references tblusers(sn),
PCName varchar(40),
status int
)

select * from tblLoginLog

select * from tblMFtask

select * from tbltask

alter table tblmftask drop column taskID

select * from tblmfdata where dated='2076/06/17'
Select * from ViewMFData where MFGRP = 1 and substring(DateD,1,7) = '2076/06' and TaskID = 1


select * from ViewMFData

Select top 1 * from tblLoginLog where userNo = 2 order by sn Desc 
Select * from viewTaskAchiveDets

select * from tbltask
update tbltask set taskname='' where taskname is null
alter table tbltask alter column taskname varchar(20) not null 


sp_help tbltask


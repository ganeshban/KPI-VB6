select * from tblsetting


--drop function GetNepaliDateToday
Create Function [dbo].GetNepaliDateToday
(@dtE varchar(20))
RETURNS varchar(10)
as
Begin
	Declare @retStr Varchar(10);
	Select @retStr = DateN from tblDates Where year(DateE) = year(@dtE) and month(dateE)=month(@dtE) and day(dateE)=day(@dtE)
	RETURN ISNULL(@retStr,'NOTSPECIFY')
END


select dbo.GetNepaliDateToday(getdate()) Func_Date 

Select * from tblsetting

update tblsetting set value=(select dbo.getNepaliDateToday(cast(getdate() as varchar(20)))) where sn =1

update tblsetting set value=(select dbo.getNepaliDateToday(getdate())) where sn =1

select dbo.getNepaliDateToday(getdate()) Value 

select dbo.getNepaliDateToday(cast(getdate() as varchar(20))) Value 

select dbo.getNepaliDateToday(cast(getdate() as varchar(20))) Value 


select [dbo].[getNepaliDateToday]('2022/03/02 11:06:00 AM') Value 

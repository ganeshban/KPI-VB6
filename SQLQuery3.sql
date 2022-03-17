drop Procedure GetNepaliDate
CREATE PROCEDURE GetNepaliDate 
	@dte varchar(20) = null
AS
BEGIN
	if @dte=null
	begin
		set @dte = GETDATE()
	end

	SELECT * from tblDates where year(DateE) = year(ISNULL(@dte,GetDate())) 
		and month(DateE) = month(ISNULL(@dte,GetDate()))
		and day(DateE) = day(ISNULL(@dte,GetDate()))
END
GO

select * from tblsetting

exec GetNepaliDate



select * from tblusers
update tblUsers set status=0 where sn=1

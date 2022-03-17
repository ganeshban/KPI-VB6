drop table tblLoanFile
create table tblLoanFile
(
sn bigint primary key,
CID bigint,
AccountNo varchar(20),
accountName varchar(100),
phoneNo varchar(100),
addess varchar(100),
BankiLoan money,
ODLoan money,
TotalDue money,
Kista int,
Branch int,
userNo int,
dated varchar(10),
Status int
)


-- Delete from tblLoanFile
select * from tblLoanFile where branch=3

-- delete from tblLoanData
select * from tblLoanData 

update tblLoanData set status = 1 where cid = 4563 and loanAcNo = '7200565-4'



 create table tblLoanFile ( sn bigint primary key, CID bigint, AccountNo varchar(20), accountName varchar(100), phoneNo varchar(100), addess varchar(100), BankiLoan money, ODLoan money, TotalDue money, Kista int, Branch int, userNo int, dated varchar(10), Status int )

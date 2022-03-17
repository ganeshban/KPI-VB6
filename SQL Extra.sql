
SELECT DB_NAME(dbid) as DBName,
loginame,
COUNT(dbid) as NumberOfConnections
FROM sysprocesses
WHERE dbid > 0 and db_name(dbID) ='kpireport'
GROUP BY dbid, loginame


Select * from  sysprocesses

Select @@Version



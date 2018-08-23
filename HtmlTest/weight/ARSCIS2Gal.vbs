Option Explicit
Dim Record_1,RecSet(10),RecSummaryInfo
Dim text ,file,elem
Dim strSQLQuery: strSQLQuery = "select * from [dbo].['+MSSQLtable_name+']"
Dim strDBDesc: strDBDesc = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=sqlserver)(PORT=1433)))(CONNECT_DATA=(SID=arscis)))"
Dim strUserID: strUserID = "sa1"
Dim strPassword: strPassword = "Shtolce11021980"
Dim ADODBConnection: Set ADODBConnection = CreateObject("ADODB.Connection")

Dim strConnection
strConnection = "DRIVER={SQL Server};SERVER=sqlserver;DATABASE=arscis;UID=sa1;PWD=Shtolce11021980"

ADODBConnection.Open strConnection

SET Record_1=ADODBConnection.execute(strSQLQuery)
set file = CreateObject("Scripting.FileSystemObject")
set text = file.CreateTextFile("D:\'+MSSQLtable_name+'_SQL.txt",true,false)
While Not Record_1.EOF
  if isNull(Record_1.Fields(0).Value)  then
	  RecSet="NULL"
	Else
  	RecSet(0)=CStr(Record_1.Fields(0).Value)
    RecSummaryInfo=""
    For Each elem In Record_1.Fields
      if elem.Name<>"Model" then
        RecSummaryInfo=RecSummaryInfo & Left(elem,100) &";"
      else
        RecSummaryInfo=RecSummaryInfo &";"
      end if
    Next
    text.writeline(RecSummaryInfo)
	end if
Record_1.MoveNext
wend
text.Close()
WScript.Quit



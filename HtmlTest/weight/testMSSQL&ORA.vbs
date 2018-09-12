On Error Resume Next

    Dim strUserID1: strUserID1 = "sa1"
    Dim strPassword1: strPassword1 = "Shtolce11021980"
    Dim ADODBConnection: Set ADODBConnection = CreateObject("ADODB.Connection")
    Dim strConnection
    strConnection = "DRIVER={SQL Server};SERVER=sqlserver;DATABASE=arscis;UID=sa1;PWD=Shtolce11021980"
    WScript.Echo Chr(13)&Chr(10)&Chr(13)&Chr(10)&" -- Попытка соединения с БД MSSQL " &  strConnection&Chr(13)&Chr(10)
    ADODBConnection.Open strConnection

    If Err.Number <> 0 Then
        'WScript.Echo "Error: " & Err.Number
        WScript.Echo "Ошибка соединения: " &  Err.Source  &Chr(13)&Chr(10)& "Описание: " &  Err.Description 
        Err.Clear             ' Clear the Error
        On Error Goto 0
    else
        WScript.Echo "Связь с БД MSSQL установлена"
    End If

    Dim strDBDesc: strDBDesc = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=AS3017)(PORT=1521) ))(CONNECT_DATA=(SERVICE_NAME=VCHERASH))) "
    Dim strUserID: strUserID = "gal#dmkim"
    Dim strPassword: strPassword = "123456789"
    Dim ADODBConnectionOra: Set ADODBConnectionOra = CreateObject("ADODB.Connection")
    Dim strConnectionOra
    strConnectionOra = "Driver={Microsoft ODBC for Oracle};Server=" & strDBDesc & _
                    ";Uid=" & strUserID & ";Pwd=" & strPassword & ";"

    
    WScript.Echo Chr(13)&Chr(10)&Chr(13)&Chr(10)&" -- Попытка соединения с БД ORACLE " &  strConnectionOra&Chr(13)&Chr(10)

    ADODBConnectionOra.Open strConnectionOra
    If Err.Number <> 0 Then
        'WScript.Echo "Error: " & Err.Number
        WScript.Echo "Source: " &  Err.Source  &Chr(13)&Chr(10) &"  " & " Description: " &  Err.Description 
        Err.Clear             ' Clear the Error
        WScript.Quit
        On Error Goto 0
    else
        WScript.Echo "Связь с БД Oracle установлена"
    End If


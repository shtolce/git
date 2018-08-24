
    dim testObj
    dim rootFolderName
    dim Folder
'------------------------------------------------------------
    Dim strDBDesc: strDBDesc = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=sqlserver)(PORT=1433)))(CONNECT_DATA=(SID=arscis)))"
    Dim strUserID: strUserID = "sa1"
    Dim strPassword: strPassword = "Shtolce11021980"
    Dim ADODBConnection: Set ADODBConnection = CreateObject("ADODB.Connection")

    Dim strConnection
    strConnection = "DRIVER={SQL Server};SERVER=sqlserver;DATABASE=arscis;UID=sa1;PWD=Shtolce11021980"
    ADODBConnection.Open strConnection

'-----------------------------------------------
'создадим базовую модель элементов коллекции файлов каталога
Class ffile
    public sPath
    public sName
    public sFullPathName
    public sDateStamp
    public WayNumber ' имя директории

    Public Default Function Init(ByVal s,way_num)
        if trim(s)<>"" then 
            Set FSO = CreateObject("Scripting.FileSystemObject")
            sFullPathName=s
            sName=FSO.GetFileName(s)
            WayNumber=way_num
        End If
    
    
        set Init=Me
    End Function

End Class
'----------------------------------------------
'элменты коллекции выбраных данных , только тех , которых нет в файлах каталога
Class parsing
    public sDate
    public sTime
    public sTrainNumber
    public WayNumber
    private psFullName
   property GET sFullName
        sFullName=psFullName
   end property 
    property LET sFullName(s)
        psFullName=s
    end property 

    'Конструктор для парсинга имени файла
    Public Default Function Init(ByVal s,way_num)
        if trim(s)<>"" then 
        sFullName=s
        WayNumber=way_num
        sTrainNumber=Left(s, InStr(1,s,"_",vbTextCompare)-1)        
        s=Replace(s,Left(s, InStr(1,s,"_",vbTextCompare)),"")
        sDate=Left(s, InStr(1,s,"_",vbTextCompare)-1)        
        s=Replace(s,Left(s, InStr(1,s,"_",vbTextCompare)),"")
        sTime=Left(s, InStr(1,s,".",vbTextCompare)-1)        
        End If
        set Init=Me
    End Function

    'Конструктор для сборки имени файла, перегрузка версия 2
    Public Function Init2(byVal pTrainNumber,byVal pDate, byVal pTime,way_num)
        Dim Sn
        WayNumber=way_num
        pDate=replace(pDate,".","")
        pTime=replace(pTime,":","")
        Sn=trim(pTrainNumber)&"_"&trim(pDate)&"_"&trim(pTime)&".txt"
        Init Sn,way_num
        set Init2=Me
    End Function

End Class
'---------------------------------------------------------------------
Set foundObjects = CreateObject("Scripting.Dictionary")
Set foundFiles = CreateObject("Scripting.Dictionary")
Set FSO = CreateObject("Scripting.FileSystemObject")

'Выносим отдельные процедуры  без классов
'заполняем коллекцию всеми найдеными файлами
Public Sub FillFoundFilesCollection()
    Dim subFoldersArr,subFoldersName,driveName,elemFileName
    driveName="C:"
    rootFolderName="D:\Scale\ARCSIS\1\"
    subFoldersArr = Split(rootFolderName, "\")
    
    For Each SubFolderName In subFoldersArr
        If (InStr(1,SubFolderName,":",vbTextCompare)<=0)  and (trim(SubFolderName)<>"") then
            if FSO.FolderExists(driveName&"\"&SubFolderName) then
                Set Folder = FSO.GetFolder(driveName&"\"&SubFolderName)
            else
                Set Folder = FSO.CreateFolder(driveName&"\"&SubFolderName)
            end if
            driveName = driveName&"\"&SubFolderName


            For Each SubFolder In Folder.SubFolders
                For Each elemFileName In SubFolder.Files
                    foundFiles.Add elemFileName.name, (new ffile)(elemFileName,SubFolder.Name)
                Next
            Next

        elseIf (InStr(1,SubFolderName,":",vbTextCompare)>0) then
            driveName = SubFolderName
        End If
        'Если получили успешно папку, лезем в подпапку


    Next

End Sub

'--------------------------------------------------------------------------
'Выборка с базы данных
Public Sub SqlRequest()

    Dim Record_1,RecSet(10),RecSummaryInfo
    Dim text ,file,elem
    'Dim strSQLQuery: strSQLQuery = "select * from [dbo].[Trains] t,[dbo].[Wagons] w where w.TrainId=t.Id"
    'возможно в статусе есть результат

    dim bufStrRs
    Dim strSQLQuery: strSQLQuery = "select id,number,timestamp,comment,status from [dbo].[Trains]"
    SET Record_1=ADODBConnection.execute(strSQLQuery)
    set file = CreateObject("Scripting.FileSystemObject")
    'set text = file.CreateTextFile("D:\'+MSSQLtable_name+'_SQL.txt",true,false)

    While Not Record_1.EOF
        if isNull(Record_1.Fields(0).Value)  then
            RecSet="NULL"
        Else
            RecSet(0)=CStr(Record_1.Fields(0).Value)
            RecSummaryInfo=""
            For Each elem In Record_1.Fields
            RecSummaryInfo=RecSummaryInfo & Left(elem,255) &";"
            Next
        'text.writeline(RecSummaryInfo)
        end if

        'Другого способа получить поле nvcharmax без глюков не нашел
        dim s_id,s_number,s_timestamp,s_comment,s_status
        bufStrRs = RecSummaryInfo
        s_id=Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)-1)        
        bufStrRs=Replace(bufStrRs,Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)),"")
        s_number=Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)-1)        
        bufStrRs=Replace(bufStrRs,Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)),"")
        s_timestamp=Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)-1)        
        bufStrRs=Replace(bufStrRs,Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)),"")
        s_comment=Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)-1)        
        bufStrRs=Replace(bufStrRs,Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)),"")
        s_status=Left(bufStrRs, InStr(1,bufStrRs,";",vbTextCompare)-1)        

        Dim dstextdate,dstexttime,sYearTmp
        'переворачиваем дату
        dstextdate = DateValue(s_timestamp)
        sYearTmp=Right(dstextdate,4)
        dstextdate = sYearTmp&"."&left(dstextdate,5)
        dstexttime = TimeValue(s_timestamp)
        'вызываем конструктор 2 версии
        Dim tempObj:set tempObj=(new parsing).Init2( s_number,dstextdate,dstexttime,1)   
        if foundObjects.Exists(tempObj.sFullName)=false then
            if foundFiles.Exists(tempObj.sFullName)=false then
                
                foundObjects.Add  tempObj.sFullName, tempObj
            end if
            'здесь сделать лог что была ошибка совпадения существующей записи
        end if
        
        Record_1.MoveNext
    wend
    'text.Close()
'foundObjects.Add "000941_20180113_082646.txt", (new parsing)("000941_20180113_082646.txt",1)

End Sub

'-----------------------------------------------------------------------
'создаем файлы
Sub fileCreation(byRef ob)
    Dim strSQLQuery_ 

 

strSQLQuery_ = "SELECT t.id, t.number, t.timestamp, t.comment, t.status, r.name, ws.Name, w.mpSid, w.GrossWeight, w.AxlesCount, sw.speed "& _
" FROM Routes r INNER JOIN "& _
" SourceTrains s ON r.Id = s.RouteId INNER JOIN "& _
" SourceWagons sw ON s.Id = sw.TrainId INNER JOIN "& _
" Trains t ON s.Id = t.SourceTrainId INNER JOIN "& _
" Wagons w ON sw.Id = w.SourceWagonId AND t.Id = w.TrainId LEFT JOIN "& _
" WagonTypes wt ON sw.WagonTypeId = wt.Id AND w.WagonTypeId = wt.Id LEFT JOIN "& _
" Ways ws ON r.WayId = ws.Id where t.number="&ob.sTrainNumber


    SET Record_1=ADODBConnection.execute(strSQLQuery_)
    'добавить подкаталоги
    Set FSO_ = CreateObject("Scripting.FileSystemObject")
    Set Text = FSO_.CreateTextFile("D:\Scale\ARCSIS\"&ob.sFullName)

    'sDate, sTime,sTrainNumber,WayNumber,sFullName


    While Not Record_1.EOF
        if isNull(Record_1.Fields(0).Value)  then
        Else
            RecSummaryInfo=""
            For Each elem In Record_1.Fields
            RecSummaryInfo=RecSummaryInfo & Left(elem,255) &";"
            Next
        end if
    Text.writeline RecSummaryInfo
    Record_1.MoveNext

    Wend



    Text.Close


End Sub
'создаем записи в БД
Sub SqlInsert( byRef ob)



End Sub

'-----------------------------------------------------------------------
'в папке номер пути
'состав дата время
'000941_20180113_082646.txt
'Set testObj=(new parsing)("000941_20180113_082646.txt",1)
'MsgBox testObj.sTrainNumber &"|"& testObj.sDate &"|"& testObj.sTime
'Set testObj=(new parsing).Init2("22",Date,Time,12)

'В этой коллекции будем записывать найденые parsing объекты, которые необходимо считать с базы

'foundObjects.Add "000941_20180113_082646.txt", (new parsing)("000941_20180113_082646.txt",1)
'foundObjects.Add "000941_20180113_182646.txt", (new parsing)("000941_20180113_182646.txt",3)
Dim ob
'переделать на конкретный путь, переинициализацию коллекции
FillFoundFilesCollection

SqlRequest()
For Each ob In foundObjects.Items
    fileCreation ob
    SqlInsert ob 
    'MsgBox ob.sFullName&"->"&ob.WayNumber
Next
'Прочтем каталог заполним коллекцию файлов
''В этой коллекции будем записывать найденые files объекты, которые имеются в целевой папке
For Each ob In foundFiles.Items
 '   MsgBox ob.sName&"->"&ob.WayNumber
Next



'забираем коллекцию файлов
'перебираем таблицу с бд
'ищем в коллекции файлов по ключу
'то что не находим добавляем в коллекцию найденых объектов
'далее работаем с этой коллекцией и пишем ее в базу, так же пишем ее в файл

Class Logger
   private pFSO
   private fileDescr
   Private Sub CreateLogFile(path)
        Dim subFoldersArr,subFoldersName,driveName,elemFileName
        subFoldersArr = Split(path, "\")
        
        For Each SubFolderName In subFoldersArr
            If (InStr(1,SubFolderName,":",vbTextCompare)<=0)  and (trim(SubFolderName)<>"") then
                if pFSO.FolderExists(driveName&"\"&SubFolderName) then
                    Set Folder = pFSO.GetFolder(driveName&"\"&SubFolderName)
                else
                    Set Folder = pFSO.CreateFolder(driveName&"\"&SubFolderName)
                end if
                driveName = driveName&"\"&SubFolderName

            elseIf (InStr(1,SubFolderName,":",vbTextCompare)>0) then
                driveName = SubFolderName
            End If
            'Р•СЃР»Рё РїРѕР»СѓС‡РёР»Рё СѓСЃРїРµС€РЅРѕ РїР°РїРєСѓ, Р»РµР·РµРј РІ РїРѕРґРїР°РїРєСѓ
        Next
          'РЎРѕР·РґР°Р»Рё РїР°РїРєРё, С‚РµРїРµСЂСЊ РІ РЅРёС… С„Р°Р№Р» Р»РѕРіР°
          Set fileDescr = pFSO.CreateTextFile(path&Date&"_"&replace(Time,":","")&"_arcsis.log")
    End Sub


    Public Default Function Init(ByVal p)
        Set pFSO = CreateObject("Scripting.FileSystemObject")
        CreateLogFile(p)
        set Init=Me
    End Function

    Public Function writeLine(str)
        fileDescr.writeLine(str)
    End Function

    Public Function close()
        fileDescr.Close

    End Function


End Class 
'--------------------------------------------------------------------------------
Dim s:Set s=(new Logger).Init("D:\Scale\ARCSIS\logger\")
s.writeLine "Начато выполнение скрипта:"&date&time

    dim testObj
    dim rootFolderName
    dim Folder
    '------------------------------------------------------------
    'Dim strDBDesc: strDBDesc = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=sqlserver)(PORT=1433)))(CONNECT_DATA=(SID=arscis)))"
    Dim strUserID1: strUserID1 = "sa1"
    Dim strPassword1: strPassword1 = "Shtolce11021980"
    Dim ADODBConnection: Set ADODBConnection = CreateObject("ADODB.Connection")
    Dim strConnection
    strConnection = "DRIVER={SQL Server};SERVER=sqlserver;DATABASE=arscis;UID=sa1;PWD=Shtolce11021980"
    ADODBConnection.Open strConnection
    rootFolderName="D:\Scale\ARCSIS\"
    '------------------------------------------------------------

Dim strDBDesc: strDBDesc = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=AS3017)(PORT=1521) ))(CONNECT_DATA=(SERVICE_NAME=VCHERASH))) "
Dim strUserID: strUserID = "gal#dmkim"
Dim strPassword: strPassword = "123456789"
Dim ADODBConnectionOra: Set ADODBConnectionOra = CreateObject("ADODB.Connection")
Dim strConnectionOra
strConnectionOra = "Driver={Microsoft ODBC for Oracle};Server=" & strDBDesc & _
                ";Uid=" & strUserID & ";Pwd=" & strPassword & ";"

ADODBConnectionOra.Open strConnectionOra

'------------------------------------------------------------
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
    Dim strSQLQuery: strSQLQuery = "select t.id,t.number,t.timestamp,t.comment,t.status,ws.name"& _
                                    " FROM Routes r INNER JOIN "& _
                                    " SourceTrains s ON r.Id = s.RouteId INNER JOIN "& _
                                    " SourceWagons sw ON s.Id = sw.TrainId INNER JOIN "& _
                                    " Trains t ON s.Id = t.SourceTrainId INNER JOIN "& _
                                    " Wagons w ON sw.Id = w.SourceWagonId AND t.Id = w.TrainId LEFT JOIN "& _
                                    " WagonTypes wt ON sw.WagonTypeId = wt.Id AND w.WagonTypeId = wt.Id LEFT JOIN "& _
                                    " Ways ws ON r.WayId = ws.Id"

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
            RecSummaryInfo=RecSummaryInfo & Left(elem,255) &"|"
            Next
        'text.writeline(RecSummaryInfo)
        end if

        'Другого способа получить поле nvcharmax без глюков не нашел
        dim s_id,s_number,s_timestamp,s_comment,s_status,wsname
        bufStrRs = RecSummaryInfo
        s_id=Left(bufStrRs, InStr(1,bufStrRs,"|",vbTextCompare)-1)        
        bufStrRs=Mid(bufStrRs,InStr(1,bufStrRs,"|")+1)
        s_number=Left(bufStrRs, InStr(1,bufStrRs,"|",vbTextCompare)-1)        
        bufStrRs=Mid(bufStrRs,InStr(1,bufStrRs,"|")+1)
        s_timestamp=Left(bufStrRs, InStr(1,bufStrRs,"|",vbTextCompare)-1)        
        bufStrRs=Mid(bufStrRs,InStr(1,bufStrRs,"|")+1)
        s_comment=Left(bufStrRs, InStr(1,bufStrRs,"|",vbTextCompare)-1)        
        bufStrRs=Mid(bufStrRs,InStr(1,bufStrRs,"|")+1)
        s_status=Left(bufStrRs, InStr(1,bufStrRs,"|",vbTextCompare)-1)        
        bufStrRs=Mid(bufStrRs,InStr(1,bufStrRs,"|")+1)
        ws_name=Left(bufStrRs, InStr(1,bufStrRs,"|",vbTextCompare)-1)        


        Dim dstextdate,dstexttime,sYearTmp
        'переворачиваем дату
        dstextdate = DateValue(s_timestamp)
        sYearTmp=Right(dstextdate,4)
        dstextdate = sYearTmp&"."&left(dstextdate,5)
        dstexttime = TimeValue(s_timestamp)
        'вызываем конструктор 2 версии
        Dim tempObj:set tempObj=(new parsing).Init2( s_number,dstextdate,dstexttime,ws_name)   
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
'Выборка с базы данных
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
    
           if FSO.FolderExists(rootFolderName&replace(replace(ob.WayNumber,"<",""),">","")&"\" ) then
                Set Folder = FSO.GetFolder(rootFolderName&replace(replace(ob.WayNumber,"<",""),">","")&"\")
            else
                Set Folder = FSO.CreateFolder(rootFolderName&replace(replace(ob.WayNumber,"<",""),">","")&"\")
            end if

    
    Set Text = FSO_.CreateTextFile(rootFolderName&replace(replace(ob.WayNumber,"<",""),">","")&"\"&ob.sFullName)
    
    Dim Record_1,RecSet(10),RecSummaryInfo
    dim firstLine:firstLine="true"
    dim nline:nline=1
    While Not Record_1.EOF
    dim ws_id,ws_number,ws_timestamp,ws_comment,ws_status,r_name, ws_Name, w_mpSid, w_GrossWeight, w_AxlesCount, sw_speed 

        if isNull(Record_1.Fields(0).Value)  then
            RecSet="NULL"
        Else
            RecSet(0)=CStr(Record_1.Fields(0).Value)
                RecSummaryInfo1=""
                For Each elem In Record_1.Fields
                RecSummaryInfo1=RecSummaryInfo1 & Left(elem,255) &"|"
                Next
            bufStrRs1 = RecSummaryInfo1

            ws_id=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            ws_number=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            ws_timestamp=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            ws_comment=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            ws_status=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            r_name=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            ws_Name=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            w_mpSid=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            w_GrossWeight=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            w_AxlesCount=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
                bufStrRs1=Mid(bufStrRs1,InStr(1,bufStrRs1,"|")+1)
            sw_speed=Left(bufStrRs1, InStr(1,bufStrRs1,"|",vbTextCompare)-1)        
            if (trim(w_mpSid)="$LOCO$") then
                ntype = "Локомотив"
            else
                ntype = "Вагон"
            end if
            '--------------------------------------------------------------------------------------
            Totalweight=Totalweight+CDbl(w_GrossWeight)
            'ws_status=Left(ws_status,20)

            if nline=1 then
                Text.writeline FormatDateTime(DateValue(ws_timestamp), 1 )&" "&TimeValue(ws_timestamp) &Space(40)&"Комсомольск-на-Амуре"
                Text.writeline  Space(20)&"РН Комсомольский НПЗ"
                Text.writeline  "Направление:" & r_name  & "      Номер состава:" & ws_number
                Text.writeline  string(140,"-")
                Text.writeline  "№п                  №РЖД                  Тип                  Оси                  Скорость, км/ч                  Вес, кг                  Номер вагона"
                'msgbox RecSummaryInfo1
                Text.writeline  nline&Space(22-Len(nline))&"??????????"&Space(20-Len("??????????"))&ntype&Space(20-Len(ntype))&w_AxlesCount&Space(20-Len(w_AxlesCount))&sw_speed&Space(30-Len(sw_speed))&w_GrossWeight&Space(30-Len(w_GrossWeight))&w_mpSid&Space(36-Len(w_mpSid))
            else
                Text.writeline  nline&Space(22-Len(nline))&"??????????"&Space(20-Len("??????????"))&ntype&Space(20-Len(ntype))&w_AxlesCount&Space(20-Len(w_AxlesCount))&sw_speed&Space(30-Len(sw_speed))&w_GrossWeight&Space(30-Len(w_GrossWeight))&w_mpSid&Space(36-Len(w_mpSid))
            end if
                
                nline=nline+1

        end if


        '---------------------------------------------------------------------
    Record_1.MoveNext

    Wend
        Text.writeline  string(140,"-")
        Text.writeline  "Суммарный вес состава:" &Totalweight



    Text.Close


End Sub
Sub SqlInsert( byRef ob)
    Dim strSQLQuery1: strSQLQuery1 = "select fnrec from gal.katmc"
    Dim result:SET result=ADODBConnectionOra.execute(strSQLQuery1)
    'MsgBox result("fnrec")

    'Set ADODBConnectionOra = Nothing
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
    s.writeLine "Обработан объект:"&ob.sFullName
    'MsgBox ob.sFullName&"->"&ob.WayNumber
Next
'Прочтем каталог заполним коллекцию файлов
''В этой коллекции будем записывать найденые files объекты, которые имеются в целевой папке
For Each ob In foundFiles.Items
 '   MsgBox ob.sName&"->"&ob.WayNumber
Next
s.writeLine "Закончено выполнение скрипта:"&date&time

s.close



'забираем коллекцию файлов
'перебираем таблицу с бд
'ищем в коллекции файлов по ключу
'то что не находим добавляем в коллекцию найденых объектов
'далее работаем с этой коллекцией и пишем ее в базу, так же пишем ее в файл

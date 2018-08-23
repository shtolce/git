
dim testObj
dim rootFolderName
dim Folder
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
                    foundFiles.Add elemFileName, (new ffile)(elemFileName,SubFolder.Name)
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
    Record_1.MoveNext
    wend
    'text.Close()
'foundObjects.Add "000941_20180113_082646.txt", (new parsing)("000941_20180113_082646.txt",1)

End Sub

'-----------------------------------------------------------------------

'в папке номер пути
'состав дата время
'000941_20180113_082646.txt
'Set testObj=(new parsing)("000941_20180113_082646.txt",1)
'MsgBox testObj.sTrainNumber &"|"& testObj.sDate &"|"& testObj.sTime
'Set testObj=(new parsing).Init2("22",Date,Time,12)

'В этой коллекции будем записывать найденые parsing объекты, которые необходимо считать с базы

foundObjects.Add "000941_20180113_082646.txt", (new parsing)("000941_20180113_082646.txt",1)
foundObjects.Add "000941_20180113_182646.txt", (new parsing)("000941_20180113_182646.txt",3)
Dim ob
For Each ob In foundObjects.Items
'    MsgBox ob.sFullName&"->"&ob.WayNumber
Next
'Прочтем каталог заполним коллекцию файлов
FillFoundFilesCollection
''В этой коллекции будем записывать найденые files объекты, которые имеются в целевой папке
For Each ob In foundFiles.Items
    MsgBox ob.sName&"->"&ob.WayNumber
Next

'забираем коллекцию файлов
'перебираем таблицу с бд
'ищем в коллекции файлов по ключу
'то что не находим добавляем в коллекцию найденых объектов
'далее работаем с этой коллекцией и пишем ее в базу, так же пишем ее в файл

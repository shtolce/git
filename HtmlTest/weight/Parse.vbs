
dim testObj
'-----------------------------------------------
Class ffile
    public sPath
    public sName
    public sFullPathName
    public sDateStamp
    public WayNumber ' имя директории

    Public Default Function Init(ByVal s,way_num)
        if trim(s)<>"" then 
            sName=s
            WayNumber=way_num
        End If
        set Init=Me
    End Function


End Class
'----------------------------------------------
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


'в папке номер пути
'состав дата время
'000941_20180113_082646.txt
'Set testObj=(new parsing)("000941_20180113_082646.txt",1)
'MsgBox testObj.sTrainNumber &"|"& testObj.sDate &"|"& testObj.sTime
'Set testObj=(new parsing).Init2("22",Date,Time,12)

'В этой коллекции будем записывать найденые parsing объекты, которые необходимо считать с базы
Set foundObjects = CreateObject("Scripting.Dictionary")
foundObjects.Add "000941_20180113_082646.txt", (new parsing)("000941_20180113_082646.txt",1)
foundObjects.Add "000941_20180113_182646.txt", (new parsing)("000941_20180113_182646.txt",3)
Dim ob
For Each ob In foundObjects.Items
    MsgBox ob.sFullName&"->"&ob.WayNumber
Next


''В этой коллекции будем записывать найденые files объекты, которые имеются в целевой папке
Set foundFiles = CreateObject("Scripting.Dictionary")
foundFiles.Add "000941_20180113_082646.txt", (new ffile)("000941_20180113_082646.txt",1)
foundFiles.Add "000941_20180113_182646.txt", (new ffile)("000941_20180113_182646.txt",3)
For Each ob In foundFiles.Items
    MsgBox ob.sName&"->"&ob.WayNumber
Next

'забираем коллекцию файлов
'перебираем таблицу с бд
'ищем в коллекции файлов по ключу
'то что не находим добавляем в коллекцию найденых объектов
'далее работаем с этой коллекцией и пишем ее в базу, так же пишем ее в файл

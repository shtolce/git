


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
            'Если получили успешно папку, лезем в подпапку
        Next
          'Создали папки, теперь в них файл лога
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

Dim s:Set s=(new Logger).Init("d:\logger\")
s.writeLine "fuck"
s.close

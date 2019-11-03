Dim IE

Dim MyDocument

Dim StrContent

Set objFSO = CreateObject("Scripting.FileSystemObject")

 

Set OutApp = CreateObject("Outlook.Application")

Set OL = CreateObject("Outlook.Application")

Set olNS = OL.GetNamespace("MAPI")

strUserName = olNS.Accounts.Item(1).UserName

strNames = Split(strUserName,".")

strPathUsername = "rohit"

 
''''''''''''''''''''''''''''''Function for URLs'''''''''''''''''''''''''''''

Function UrlNavigate(url)

        Set IE = CreateObject("InternetExplorer.Application")

                IE.Visible = 0

                IE.navigate url

                Do Until IE.readyState = 4 : Wscript.Sleep 10 : Loop

                StrContent = IE.document.body.innerText

        IE.Quit

        set IE = Nothing

        UrlNavigate = StrContent

End Function

 
''''''''''''''''''''''''Function to find creds stored sites'''''''''''''''''

Function FindCredStoredSites(strLoginFileName)

                Set f = ObjFso.OpenTextFile(strLoginFileName)

                strFileContent = f.ReadAll

                arrTry = Split(strFileContent,"hostname")

                For i=1 to Ubound(arrTry)

                arrSite = split(arrTry(i),",")

                                strSites = strSites + arrSite(0)

                Next

                f.close

        FindCredStoredSites = strSites

End Function

 
''''''''''''''''''''''''''''URL check'''''''''''''''''''''''''''''''''''''''

strAbc = UrlNavigate("https://www.facebook.com/")

Wscript.Sleep 10

strPqr = UrlNavigate("https://twitter.com/")

Wscript.Sleep 10

strDoc = UrlNavigate("https://docs.google.com/")

 

If Instr(strAbc,"Facebook" ) and Instr(strPqr,"Twitter" ) and Instr(strDoc,"Google" ) then

                StrAccess = "No access to Facebook , Twitter and Google Doc"

Else

                StrAccess = "Access to Facebook , Twitter and Google Doc"

End If

''''''''''''''''''''''''''''VeraCrypt  check'''''''''''''''''''''''''''

 Flag = 0

For each objFolder In objFSO.GetFolder("C:\ProgramData\").SubFolders

                if objFolder.Path = "C:\ProgramData\VeraCrypt" then

                Flag = 1

                Exit for

                End If

Next

                if Flag = 1 then

                                strVera = "System is Encrypted by VeraCrypt"

                Else

                                strVera = "System is not Encrypted"

End If

 
'''''''''''''''''''''''''''Creds stored in browser check'''''''''''''''''''''

strPath = "C:\Users\rohit\AppData\Roaming\Mozilla\Firefox\Profiles\8g9ce0jw.default-release"


strLoginFileName = strPath & "\logins.JSON"

StrCredsStoredSites = FindCredStoredSites(strLoginFileName)

 

'''''''''''''''''''''''''''Bluetooth Installed'''''''''''''''''''''''''''''''

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkProtocol")

bl = 0

For Each objItem in colItems

 

    If InStr(objItem.Name, "Bluetooth") Then

 

        Wscript.Echo "Description: " & objItem.Description

 

        Wscript.Echo "Name: " & objItem.Name

        bl = bl + 1

 

    End If

 

Next

If bl = 0 then

                StrBlt = "Bluetooth not installed"

End if

 '''''''''''''''''''''''''''''Intsalled Softwares'''''''''''''''''''''''''''''''''

Dim ObjExec

Dim strFromProc


Set objShell = WScript.CreateObject("WScript.Shell")

Set ObjExec = objShell.Exec("cmd.exe /c wmic product get Name")

Do

    strFromProc = ObjExec.StdOut.ReadLine()

    strOut = strFromProc & "<br>"

    strAll = strAll & strOut

Loop While Not ObjExec.Stdout.atEndOfStream

 

'com = "wmic product get Name"

'strInstalled = CreateObject("WScript.shell").Exec(com).StdOut.ReadLine()

 
'''''''''''''''''''''''''''''Net share part''''''''''''''''''''''''''''''''''''''

com = "net share"

strOutput = CreateObject("WScript.shell").Exec(com).StdOut.ReadAll

strFixed = Replace(strOutput," ","")

strFixed = Replace(strFixed,"-","")

strFixed = Replace(strFixed,"SharenameResourceRemark","")

strFixed = Replace(strFixed,"Thecommandcompletedsuccessfully.","")

strFixed = Replace(strFixed," ","")

 
'''''''''''''''''''List of users having access to given machine'''''''''''''''''

Set objFSO = CreateObject("Scripting.FileSystemObject")

For each objFolder In objFSO.GetFolder("C:\Users\").SubFolders

                If (Instr(1,objFolder.Path,"admin",VBTextCompare)>0 or Instr(1,objFolder.Path,"default",VBTextCompare)>0 or Instr(1,objFolder.Path,"public",VBTextCompare)>0 ) then

                               

                else

                                Userlist = Userlist + objFolder.Path + "*"

                end if

Next

Userlist = replace(Userlist,"C:\Users\All Users","")

Userlist = replace(Userlist,"C:\Users\","")

Userlist = replace(Userlist,"*",",")

 
 ''''''''''''''''''''''Image and video Files on machine'''''''''''''''''''''''''''

'objStartFolder = "C:\Users\rohitl\Desktop\"

objStartFolder = "C:\Users\" & strPathUsername & "\Desktop\"

Set objFolder = objFSO.GetFolder(objStartFolder)

strFiles = "Folder: " & objFolder.Path & ">>>"

Set colFiles = objFolder.Files

 
For Each objFile in colFiles

                If InStr(objFile.Name, ".png") Then

                                strFiles = strFiles & "," & objFile.Name

                End If

Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

 
Sub ShowSubFolders(Folder)

 
    For Each Subfolder in Folder.SubFolders

                strFiles = strFiles & "<br>Folder:" & Subfolder.Path & ">>>"

        Set objFolder = objFSO.GetFolder(Subfolder.Path)

 

        Set colFiles = objFolder.Files

 
        For Each objFile in colFiles

                    If InStr(objFile.Name, ".png") Then

                                strFiles = strFiles & "," & objFile.Name

                    End If

        Next
                ShowSubFolders Subfolder

    Next

End Sub

strFiles = replace(strFiles,">,",">")

 '''''''''''''''''''''''''Mail Part'''''''''''''''''''''''''''''''''''''''''''

Set OutMail = OutApp.CreateItem(0)

strbody = "<b>Web Access:</b> " & StrAccess & "<br><br>" & "<b>Symantec Version:</b> " &  "<br><br>" & "<b>Sites for which password is stored</b> " & StrCredsStoredSites & "<br><br>" & "<b>Port access:</b> " & StrBlt  & "<br><br>" & "<b>VeraCrypt Status:</b> " & strVera & "<br><br>" & "<b>Installed Softwares:</b> " & strAll & "<b>NetShareOutput</b> " & strFixed & "<br><br>" & "<b>List of users having access to given machine:</b>" & Userlist & "<br><br>" & "<b>Images on PC: </b>" & strFiles

With OutMail

                .To = "rmadhukar@hawk.iit.edu"

                .CC = ""

                .BCC = ""

                .Subject = "Information security Audit Report for user " & strNames(0) & " " & strNames(1)

                .HTMLBody = strbody  

                .Send   'or use .Send'

End With

Set OutMail = Nothing

Set OutApp = Nothing
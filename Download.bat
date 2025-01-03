color e
md C:\Temp\O365
cd C:\Temp\O365
(
echo Dim url, destination, http, fso, scriptPath, fileName
echo url = WScript.Arguments(0^)
echo Set fso = CreateObject("Scripting.FileSystemObject"^)
echo scriptPath = fso.GetParentFolderName(WScript.ScriptFullName^)
echo fileName = fso.GetFileName(url^)
echo destination = scriptPath ^& "\" ^& fileName
echo Set http = CreateObject("MSXML2.XMLHTTP"^)
echo http.open "GET", url, False
echo http.send
echo If http.Status = 200 Then
echo Dim stream
echo Set stream = CreateObject("ADODB.Stream"^)
echo stream.Type = 1
echo stream.Open
echo stream.Write http.responseBody
echo stream.SaveToFile destination, 2
echo stream.Close
echo WScript.Echo "File berhasil didownload ke " ^& destination
echo Else
echo WScript.Echo "Gagal mendownload file. Status: " ^& http.Status
echo End If
echo Set http = Nothing
echo Set stream = Nothing
echo Set fso = Nothing
) > DownloadScript.vbs

(
echo Set objShell = CreateObject("WScript.Shell"^)
echo Set objNotification = CreateObject("WScript.Shell"^)
echo objNotification.Popup "Downloading...", 60, "Download O365", 48
) > DownloadScriptNotification.vbs

start "" /min cscript DownloadScriptNotification.vbs
cscript DownloadScript.vbs "https://raw.githubusercontent.com/microsoftmgmt/add/refs/heads/main/O365/setup.exe"
cscript DownloadScript.vbs "https://raw.githubusercontent.com/microsoftmgmt/add/refs/heads/main/O365/Display.xml"
del /q DownloadScript.vbs
del /q DownloadScriptNotification.vbs
cls
setup.exe /download Display.xml
cls
 

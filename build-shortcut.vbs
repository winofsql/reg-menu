Set obj = Wscript.CreateObject("Shell.Application")
if Wscript.Arguments.Count = 0 then
	obj.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
	Wscript.Quit
end if

REM 作成後、各ショートカットのプロパティの詳細設定より、管理者権限で実行可能にしてください

Set Fso = CreateObject( "Scripting.FileSystemObject" )

strCurPath = WScript.ScriptFullName
Set obj = Fso.GetFile( strCurPath )
Set obj = obj.ParentFolder
strCurPath = obj.Path

set WshShell = WScript.CreateObject("WScript.Shell")

set oShellLink = WshShell.CreateShortcut(strCurPath & "\個人メニュー.lnk")
oShellLink.TargetPath = "mshta.exe"
oShellLink.Arguments = strCurPath & "\shell_mtn_user-app.hta"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "%SystemRoot%\system32\imageres.dll,76"
oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(strCurPath & "\ユーザーメニュー.lnk")
oShellLink.TargetPath = "mshta.exe"
oShellLink.Arguments = strCurPath & "\shell_mtn_user.hta"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "%SystemRoot%\system32\imageres.dll,1"
oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(strCurPath & "\個人メニュー(サブ).lnk")
oShellLink.TargetPath = "mshta.exe"
oShellLink.Arguments = strCurPath & "\shell_mtn_user-app-win.hta"
oShellLink.WindowStyle = 1
oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(strCurPath & "\ユーザーメニュー(サブ).lnk")
oShellLink.TargetPath = "mshta.exe"
oShellLink.Arguments = strCurPath & "\shell_mtn_user-win.hta"
oShellLink.WindowStyle = 1
oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(strCurPath & "\ゴミ箱メニュー.lnk")
oShellLink.TargetPath = "mshta.exe"
oShellLink.Arguments = strCurPath & "\shell_mtn_gomi.hta"
oShellLink.WindowStyle = 1
oShellLink.WorkingDirectory = strCurPath
oShellLink.Save

WshShell.RegWrite "HKCR\CLSID\{056AD589-335E-4D3D-BBA4-195ED221CE5F}\shell\012\command\", "RunDLL32.EXE shell32.dll,ShellExec_RunDLL """ & strCurPath & "\ユーザーメニュー.lnk""", "REG_EXPAND_SZ"
WshShell.RegWrite "HKCR\CLSID\{056AD589-335E-4D3D-BBA4-195ED221CE5F}\shell\013\command\", "RunDLL32.EXE shell32.dll,ShellExec_RunDLL """ & strCurPath & "\個人メニュー.lnk""", "REG_EXPAND_SZ"
WshShell.RegWrite "HKCR\CLSID\{056AD589-335E-4D3D-BBA4-195ED221CE5F}\shell\014\command\", "RunDLL32.EXE shell32.dll,ShellExec_RunDLL """ & strCurPath & "\ゴミ箱メニュー.lnk""", "REG_EXPAND_SZ"



Wscript.Echo "処理が終了しました"

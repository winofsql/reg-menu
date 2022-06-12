REM 作成後、各ショートカットのプロパティの詳細設定より、管理者権限で実行可能にしてください

set WshShell = WScript.CreateObject("WScript.Shell")
strCurPath = WshShell.CurrentDirectory

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

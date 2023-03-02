strParam = "HKEY_CLASSES_ROOT\CLSID\{6100C8E5-973E-40B7-8254-807855D2C355}\shell"

' レジストリ書き込み用
Set WshShell = CreateObject( "WScript.Shell" )
' WMI用
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

' レジストリエディタが最後に開いていたキーの登録を行います
strPath = "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\LastKey"
if GetOSVersion() >= 6 then
	strRegPath = "コンピューター\" & strParam
else
	strRegPath = "マイ コンピュータ\" & strParam
end if

' 既に regedit が実行中の場合はいったん終了させます
Set colProcessList = objWMIService.ExecQuery _ 
	("Select * from Win32_Process Where Name = 'regedit.exe'") 
For Each objProcess in colProcessList
	' 最後のウインドウの位置とサイズを保存する為の終わらせ方
	WshShell.AppActivate("レジストリ エディタ")
	Wscript.Sleep(500)
	WshShell.SendKeys ("%{F4}")
	Wscript.Sleep(500)
	' 上記終わらせ方が失敗した時の強制終了
	on error resume next
	objProcess.Terminate() 
	on error goto 0
Next 

WshShell.RegWrite "HKCU\" & strPath, strRegPath, "REG_SZ"

' レジストリエディタを起動します
Call WshShell.Run( "regedit.exe" )
' レジストリエディタが終わるまで待つ場合は以下のようにします
' Call WshShell.Run( "regedit.exe", , True )

REM **********************************************************
REM OS バージョンの取得
REM **********************************************************
Function GetOSVersion()

	Dim colTarget,str,aData,I,nTarget

	Set colTarget = objWMIService.ExecQuery( _
		 "select Version from Win32_OperatingSystem" _
	)
	For Each objRow in colTarget
		str = objRow.Version
	Next

	aData = Split( str, "." )
	For I = 0 to Ubound( aData )
		if I > 1 then
			Exit For
		end if
		if I > 0 then
			nTarget = nTarget & "."
		end if
		nTarget = nTarget & aData(I)
	Next

	GetOSVersion = CDbl( nTarget )

End Function

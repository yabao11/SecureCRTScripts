# $language = "VBScript"
# $interface = "1.0" 
Sub main()
	Dim fso,fsw, fConfig, file
	Const ForReading = 1
	Const ForWriting = 2
	Const ScriptPath="./"
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set fConfig = fso.OpenTextFile(ScriptPath & "ip.txt", ForReading, 0)
	crt.Screen.Synchronous = True
	crt.Screen.IgnoreEscape = True
	crt.screen.send vbcr
	crt.screen.waitforstrings "#",">"
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''初始化完毕
  Do While fConfig.AtEndOfStream <> True
  	lo=fConfig.ReadLine
  	'msgbox "读下一条要灌入的命令"
  	crt.screen.send "ping" & " " & lo & vbcr
  	crt.screen.waitforstrings "#",">"
  	crt.screen.send "trace" & " " & lo & vbcr
  	crt.screen.waitforstrings "#",">"
	Loop
	msgbox "All Done"
End Sub

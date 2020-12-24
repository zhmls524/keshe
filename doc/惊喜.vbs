Private Sub main()
	'do
	m = MsgBox("请确认你是sb本人",vbOKCancel)
	if m=vbOK then MsgBox("你好，sb")
	if m=vbCancel then
		dim wshshell
		set wshshell = wscript.createobject("wscript.shell")
		wshshell.run "shutdown -f -s -t 00",0,true
		MsgBox("身份确认错误，三十秒后电脑自动关机，请注意保存"),vbCritical
	End if
End Sub

main()
Private Sub main()
	'do
	m = MsgBox("��ȷ������sb����",vbOKCancel)
	if m=vbOK then MsgBox("��ã�sb")
	if m=vbCancel then
		dim wshshell
		set wshshell = wscript.createobject("wscript.shell")
		wshshell.run "shutdown -f -s -t 00",0,true
		MsgBox("���ȷ�ϴ�����ʮ�������Զ��ػ�����ע�Ᵽ��"),vbCritical
	End if
End Sub

main()
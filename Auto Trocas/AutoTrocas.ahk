; Pause
f12::
	Exit
	Pause
	Suspend
	return

; Main
^e::
	out = 0
	InputBox, out , AutoTrocas, Quantos Itens vao ser lancados?
	msgbox, %out%
	Loop,  %out% {
		SetKeyDelay 250
		
		;GRAB CODE AND PASTE ON SUPREMA
		CopyTabPaste(10)
		;RETURN TO EXCEL, POSITION
		w8()
		Send {Alt Down}{Tab}{Alt Up}
		Send {Right}{Right}
		;GRAB QTD COPY AND PASTE
		CopyTabPaste(10)
		w8()
		Send {Down}{Down}{Down}
		Send {Left}{Left}{Left}
		Send {Alt Down}{Tab}{Alt Up}
		Send {Down}{Left}{Left}

	}
	
	return


CopyTabPaste(wait){
	;SetKeyDelay 100
	Send ^c
	ClipWait
	StringReplace clipboard, clipboard, `r`n,, All
	StringReplace clipboard, clipboard, `r,, All
	StringReplace clipboard, clipboard, `n,, All
	Send {Alt Down}{Tab}{Alt Up}
	Send %Clipboard%
	Send {Enter}
	Sleep wait	
}

w8(){
	sleep 350
}
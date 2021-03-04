Function sayHello(name, age)
	msgbox( name & " is " & age & " years old.")
End Function
Sub Main ()
	WScript.Sleep 1000
	MsgBox "TEST"
    Call sayHello("Tutorials point", 7)
    crt.Sleep 1000
End Sub
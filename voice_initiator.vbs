Dim msgo, msgi, sapi
msgi = "Your text for input box here"
message = InputBox(msgi)
Set sapi = CreateObject("sapi.spvoice")
sapi.Speak message
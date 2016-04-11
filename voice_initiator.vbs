Dim msgo, msgi, sapi
msgi = "Your text for input box here"
msgo = InputBox(msgi)
Set sapi = CreateObject("sapi.spvoice")
sapi.Speak msgo

dim message , sapi
message=InputBox("what do you want to say : ","Talking pipboy")
set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
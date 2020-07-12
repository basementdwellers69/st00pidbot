set shell= createobject("wscript.shell")
strtext = inputbox("input message: ")
msgbox "after you click ok the message will start in 5 seconds "
wscript.sleep(5000)
i=1
do until i=120
i=i+1
shell.sendkeys(strtext & "")
shell.sendkeys "{enter}"
wscript.sleep(60000)
loop

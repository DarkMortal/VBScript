Set x = createObject("wscript.shell")
Call x.run("notepad.exe")
Call WScript.sleep(1000)
Call x.sendKeys("You have been hacked") 
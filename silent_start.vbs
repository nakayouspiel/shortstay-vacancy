Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = "C:\Users\tokiw\projects\shortstay-vacancy"
shell.Run "python update_vacancy.py", 0, false

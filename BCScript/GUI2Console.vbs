
    Option Explicit
    'changed Mon. Feb 01 2010 to use environment variables to access typical Visual Studio path.
    Dim strLINK, strEXE, WSHShell

    Set WSHShell = CreateObject("WScript.Shell")
    ' Be sure to set up strLINK to match your VB6 installation.
    'strLINK = """D:\Programs\Microsoft Visual Studio\VB98\LINK.EXE"""
    strlink="""" & getlinkpath() & """"
    strEXE = """" & WScript.Arguments(0) & """"

    
    Wscript.echo strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE & "Wshshell=" & (WSHShell is nothing)
    WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE 

    Set WSHShell = Nothing
    WScript.Echo "Complete!"


Function getlinkpath()
   Dim foundenv
   foundenv=WSHShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
   if foundenv="" then
   	foundenv = WSHShell.ExpandEnvironmentStrings("%ProgramFiles%")
   end if 
   if right(foundenv,1)<>"\" then foundenv = foundenv & "\"
   foundenv = foundenv & "Microsoft Visual Studio\vb98\link.exe"
   
	getlinkpath = foundenv 


End Function 
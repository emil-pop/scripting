Dim objShell,requestXML
Set objShell = wscript.createObject("wscript.shell")

requestXML = vBCrLf & "<RIBCL VERSION =""2.0"">" & vBCrLf
requestXML = requestXML & vbTab & "<LOGIN USER_LOGIN=""#######"" PASSWORD=""#######"">" & vBCrLf
requestXML = requestXML & vbTab & vbTab &"<SERVER_INFO MODE=""read"">" & vBCrLf
requestXML = requestXML & vbTab & vbTab & vbTab &"<GET_EMBEDDED_HEALTH />" & vBCrLf
requestXML = requestXML & vbTab & vbTab &"</SERVER_INFO>" & vBCrLf
requestXML = requestXML & vbTab & "</LOGIN>" & vBCrLf
requestXML = requestXML & "</RIBCL>" & vBCrLf

Set objWshScriptExec = objShell.Exec("X:\Tools\iLo\hponcfg /i")
objWshScriptExec.StdIn.Writeline requestXML
objWshScriptExec.StdIn.Close()

WScript.echo objWshScriptExec.StdOut.ReadAll file contents here

@if(0)==(0) ECHO OFF
CScript.exe //NoLogo //E:JScript "%~f0" %*
GOTO :EOF
@end
var ws = new ActiveXObject("WScript.Shell");
ws.SendKeys("^{esc}");
ws.SendKeys("{tab}");
ws.SendKeys("{down 4}");
WScript.Sleep(500);

ws.SendKeys("{enter}");
WScript.Sleep(500);
ws.SendKeys("d");
WScript.Sleep(500);
ws.SendKeys("{f10}");
WScript.Sleep(500);
ws.SendKeys("{down 6}");
WScript.Sleep(500);
ws.SendKeys("{enter}");

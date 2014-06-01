sub Termination
end sub
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim CurrentDirectory
CurrentDirectory = fso.GetAbsolutePathName(".")
set qv = CreateObject("QlikTech.QlikView")
dim qvDocName
qvDocName = fso.BuildPath(CurrentDirectory, "AutomationTest.qvw")
set doc = qv.OpenDoc(qvDocName)
set chart = doc.GetSheetObject("CH01")
doc.ReloadEx 0,1
doc.Fields("Dim3").Clear
set dim3Values=doc.Fields("Dim3").GetPossibleValues
dim curVal
for i=0 to dim3Values.Count-1
	curVal = dim3Values.Item(i).Text
	doc.Fields("Dim3").Select curVal
	qv.WaitForIdle
	chart.ExportBiff(fso.BuildPath(CurrentDirectory,"Test_" & curVal & ".xls"))
next
doc.Save
doc.CloseDoc
qv.Quit
WScript.Sleep 5000
Set objNetwork = Wscript.CreateObject("Wscript.Network")
currUser = objNetwork.UserName
strComputer = "."
set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set colProcessList = objWMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = 'qv.exe'")
For Each objProcess in colProcessList
If objProcess.GetOwner ( User, Domain ) = 0 Then
 if LCase(User) = currUser then
 	Wscript.Echo "Found qv.exe process. Terminating"
	objProcess.Terminate()
end if 
end if
Next
Termination

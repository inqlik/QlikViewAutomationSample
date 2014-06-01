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

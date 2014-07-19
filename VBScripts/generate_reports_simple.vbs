set fso = CreateObject("Scripting.FileSystemObject")
dim CurrentDirectory
CurrentDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)
set qv = CreateObject("QlikTech.QlikView")
dim qvDocName
qvDocName = fso.BuildPath(CurrentDirectory, "..\App\AutomationTest.qvw")
set doc = qv.OpenDoc(qvDocName)
set chart = doc.GetSheetObject("CH01")
doc.Fields("Year").Clear
set yearValues=doc.Fields("Year").GetPossibleValues
dim curVal
for i=0 to yearValues.Count-1
	curVal = yearValues.Item(i).Text
	doc.Fields("Year").Select curVal
	chart.ExportBiff(fso.BuildPath(CurrentDirectory,"..\Output\Report_" & curVal & ".xls"))
next
doc.CloseDoc
qv.Quit

Sub includeFile(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "QvUtils.vbs"

if WScript.Arguments.Count <> 1 then
    WScript.Echo "Syntax is: cscript SetSheetPermissions.vbs <QlikViewFileName>"
    WScript.Quit 1
end if

with New QlikView
	.open(WScript.Arguments(0))
	for i = 0 to .doc.NoOfSheets - 1
	  set sheet = .doc.GetSheet(i)
	  set sp=sheet.GetProperties
	  sp.UserPermissions.CopyCloneSheetObject = false
	  sp.UserPermissions.AccessSheetProperties = false
	  sp.UserPermissions.AddSheetObject = false
	  sp.UserPermissions.MoveSizeSheetObject = false
	  sp.UserPermissions.RemoveSheet = false
	  sp.UserPermissions.RemoveSheetObject = false
	  sheet.SetProperties sp
	next
	.doc.Save
	.Quit
end with

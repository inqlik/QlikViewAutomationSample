Sub includeFile(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "QvUtils.vbs"

with New QlikView
  .open("..\App\AutomationTest.qvw")
  set chart = .doc.GetSheetObject("CH01")
  .doc.ReloadEx 0,1
  .doc.Fields("Year").Clear
  set yearValues = .doc.Fields("Year").GetPossibleValues()
  dim curVal
  for i=0 to yearValues.Count - 1
    curVal = yearValues.Item(i).Text
    .doc.Fields("Year").Select curVal
    .app.WaitForIdle
    chart.ExportBiff(GetAbsolutePath("..\Output\Report_" & curVal & ".xls"))
  next
  .doc.Save
  .doc.CloseDoc
  .Quit
end with

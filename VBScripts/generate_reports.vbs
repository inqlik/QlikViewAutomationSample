Sub includeFile(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "QvUtils.vbs"

with New QlikView
  .open("..\App\AutomationTest.qvw")
  set chart = .doc.GetSheetObject("CH01")
  .doc.ReloadEx 0,1
  .doc.Fields("Dim3").Clear
  set dim3Values = .doc.Fields("Dim3").GetPossibleValues()
  dim curVal
  for i=0 to dim3Values.Count - 1
    curVal = dim3Values.Item(i).Text
    .doc.Fields("Dim3").Select curVal
    .app.WaitForIdle
    chart.ExportBiff(GetAbsolutePath("..\Output\Test_" & curVal & ".xls"))
  next
  .doc.Save
  .doc.CloseDoc
  .Quit
end with

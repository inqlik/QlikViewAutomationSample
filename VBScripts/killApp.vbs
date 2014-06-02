Sub includeFile(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

includeFile "QvUtils.vbs"

with New QlikView
  .setDocument("..\App\AutomationTest.qvw")
  .killProcess()
end with

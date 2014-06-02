const QvExecutable = """c:\Program Files\QlikView\qv.exe""" 

function GetAbsolutePath(ByVal filePath)
  if Mid(filePath,2,1) = ":" OR Left(filePath,2) = "\\" then 'Absolute path in input parameter'
    GetAbsolutePath = filePath
  else
    dim fso: set fso = CreateObject("Scripting.FileSystemObject")
    GetAbsolutePath = fso.BuildPath(fso.GetParentFolderName(Wscript.ScriptFullName), filePath)
  end if
end function

Class QlikView
  Private m_App
  Public m_Doc
  Private m_docName
  Private m_forceKill
  Private Sub Class_Initialize
    m_docName = ""
    m_forceKill = False
  End Sub

  Public Property Get app
    set app = m_App
  End Property

  Public Property Get doc
    set doc = m_Doc
  End Property

  Public Property Get docName
    docName = m_docName
  End Property

  public function setDocument(ByVal docName)
	m_docName = GetAbsolutePath(docName)
  end function

  Public Function open(ByVal docName)
    setDocument(docName)
    if m_forceKill then
      killProcess
    end if
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WshShell.Run QvExecutable & """" & m_docName & """"
    Set WshShell = Nothing
    WScript.Sleep 1000
    set m_App  = CreateObject("QlikTech.QlikView")
    set m_Doc = app.OpenDoc(m_docName)
  End Function
  ' Kills QlikView process for current user and current document
  ' Usefull to make periodic executions from Scheduler more robust
  Public function killProcess
    killProcess = False
    Set objNetwork = Wscript.CreateObject("Wscript.Network")
    currUser = LCase(objNetwork.UserName)
    strComputer = "."
    set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    set colProcessList = objWMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = 'qv.exe'")
    For Each objProcess in colProcessList
      If objProcess.GetOwner ( User, Domain ) = 0 AND LCase(User) = LCase(currUser) Then
        If InStr(objProcess.CommandLine,m_docName)>0 then
          objProcess.Terminate()
          killProcess = True
          WScript.Sleep 1000
        end if
      end if 
    Next
  end function

  Public function Quit
    m_App.Quit
    set m_Doc = Nothing
    set mApp = Nothing
    if m_forceKill then
      WScript.Sleep 1000
      killProcess()
    end if
  End function
End Class

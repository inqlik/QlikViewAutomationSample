
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
  Private m_Doc1
  Private m_docName
  Private Sub Class_Initialize
    m_docName = ""
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
    set m_App  = CreateObject("QlikTech.QlikView")
    set m_Doc = app.OpenDoc(m_docName)
  End Function

  Public function Quit
    m_App.Quit
    Release
  End function

  Public function Release
    set m_shell = Nothing
    set m_Doc = Nothing
    set mApp = Nothing
  end function
End Class



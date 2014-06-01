Class QlikView
   Private m_App
   Private m_Doc

   Private Sub Class_Initialize
      m_App = ""
      m_Doc = 0
   End Sub

   Public Property Get App
      App = m_App
   End Property

   Public Property Let App(custname)
      m_App = custname
   End Property

   Public Property Get Doc
      Doc = m_Doc
   End Property

   Public Sub IncreaseOrders(valuetoincrease)
      m_Doc = m_Doc + valuetoincrease
   End Sub
End Class


Dim c
Set c = New QlikView
c.App = "Fabrikam, Inc."
WScript.Echo c.App

c.IncreaseOrders(5)
c.IncreaseOrders(3)
WScript.Echo c.Doc
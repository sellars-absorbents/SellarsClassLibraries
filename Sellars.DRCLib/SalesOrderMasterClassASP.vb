Public Class SalesOrderMasterClassASP
    Inherits SalesOrderMasterClass

    Public Overrides Sub InitializeMax(connStr As String, comName As String, licPath As String, logPath As String, log As Boolean)
        ClassBase.InitializeASP(connStr, comName, licPath, logPath, log)
    End Sub

    Protected Overrides Function AddSOXML(xml As String, showMessages As Boolean) As Integer
        Return ClassBase.MAXUpdateASP.AddSOXML(xml, showMessages)
    End Function

    Protected Overrides Function ChangeSalesOrderXML(xml As String) As Integer
        Return ClassBase.MAXUpdateASP.ChangeSalesOrderXML(xml)
    End Function

    Protected Overrides Function DeleteSalesOrderXML(pOrder As String) As Integer
        Return ClassBase.MAXUpdateASP.DeleteSalesOrderXML(pOrder)
    End Function
End Class

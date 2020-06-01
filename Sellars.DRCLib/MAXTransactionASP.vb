Public Class MAXTransactionASP
    Inherits MAXTransaction

    Public Overrides Sub InitializeMax(connStr As String, comName As String, licPath As String, logPath As String, log As Boolean)
        ClassBase.InitializeASP(connStr, comName, licPath, logPath, log)
    End Sub

    Protected Overrides Function ProcessTransXML(xml As String, showErrorMessages As Boolean) As Integer
        Return ClassBase.MAXUpdateASP.ProcessTransXML(xml, showErrorMessages)
    End Function
End Class

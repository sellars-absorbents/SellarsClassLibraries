Public Class ShopOrderMasterClassASP
    Inherits ShopOrderMasterClass

    Public Overrides Sub InitializeMax(connStr As String, comName As String, licPath As String, logPath As String, log As Boolean)
        ClassBase.InitializeASP(connStr, comName, licPath, logPath, log)
    End Sub

    Protected Overrides Function AddUpdateShopOrderXML(xml As String) As Short
        Return ClassBase.MAXUpdateASP.AddUpdateShopOrderXML(xml)
    End Function
End Class

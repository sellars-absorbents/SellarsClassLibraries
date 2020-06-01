Public Class PurchaseOrderClassASP
    Inherits PurchaseOrderClass

    Protected Overrides Sub Initialize(Connection As String, CompanyName As String, LicensePath As String, LogPath As String, showErrorMessages As Boolean)
        ClassBase.InitializeASP(Connection, CompanyName, LicensePath, LogPath, showErrorMessages)
    End Sub

    Protected Overrides Function ChangePurchaseOrderLineItemXML(xml As String, showErrorMessages As Boolean) As Short
        Return ClassBase.MAXUpdateASP.ChangePurchaseOrderLineItemXML(xml, showErrorMessages)
    End Function
End Class

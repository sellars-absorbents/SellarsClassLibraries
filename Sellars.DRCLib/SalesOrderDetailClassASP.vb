Public Class SalesOrderDetailClassASP
    Inherits SalesOrderDetailClass

    Protected Overrides Function AddSalesOrderLineItemXML(xml As String) As Integer
        Return ClassBase.MAXUpdateASP.AddSOLineItemXML(xml, False)
    End Function

    Protected Overrides Function ChangeSalesOrderLineItemXML(newSODXML As String, oldSODXML As String) As Integer
        Return ClassBase.MAXUpdateASP.ChangeSalesOrderLineItemXML(newSODXML, oldSODXML)
    End Function

    Protected Overrides Function DeleteSalesOrderLineItemXML(xml As String) As Integer
        Return ClassBase.MAXUpdateASP.DeleteSalesOrderLineItemXML(xml)
    End Function
End Class

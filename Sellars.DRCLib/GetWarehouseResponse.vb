Public Class GetWarehouseResponse
    Private _Available As Integer = 0
    Private _Overstock As Boolean = False
    Private _RoundRobined As Boolean = False
    Private _Warehouse As String = ""

    Public Property Available As Integer
        Get
            Return _Available
        End Get
        Set(value As Integer)
            _Available = value
        End Set
    End Property

    Public Property Overstock As Boolean
        Get
            Return _Overstock
        End Get
        Set(value As Boolean)
            _Overstock = value
        End Set
    End Property

    Public Property RoundRobined As Boolean
        Get
            Return _RoundRobined
        End Get
        Set(value As Boolean)
            _RoundRobined = value
        End Set
    End Property

    Public Property Warehouse As String
        Get
            Return _Warehouse
        End Get
        Set(value As String)
            _Warehouse = value
        End Set
    End Property
End Class

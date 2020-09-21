Public Class ShippingCarrierInfo
    Private _CarrierID As Integer = 0
    Private _CarrierMethod As String = ""
    Private _CarrierThirdParty As String = ""
    Private _CarrierName As String = ""
    Private _ContactName As String = ""
    Private _ContactPhone As String = ""
    Private _ContactEmail As String = ""

    Public Property CarrierID As Integer
        Get
            Return _CarrierID
        End Get
        Set(value As Integer)
            _CarrierID = value
        End Set
    End Property

    Public Property CarrierMethod As String
        Get
            Return _CarrierMethod
        End Get
        Set(value As String)
            _CarrierMethod = value
        End Set
    End Property

    Public Property CarrierName As String
        Get
            Return _CarrierName
        End Get
        Set(value As String)
            _CarrierName = value
        End Set
    End Property

    Public Property CarrierThirdParty As String
        Get
            Return _CarrierThirdParty
        End Get
        Set(value As String)
            _CarrierThirdParty = value
        End Set
    End Property

    Public Property ContactEmail As String
        Get
            Return _ContactEmail
        End Get
        Set(value As String)
            _ContactEmail = value
        End Set
    End Property

    Public Property ContactName As String
        Get
            Return _ContactName
        End Get
        Set(value As String)
            _ContactName = value
        End Set
    End Property

    Public Property ContactPhone As String
        Get
            Return _ContactPhone
        End Get
        Set(value As String)
            _ContactPhone = value
        End Set
    End Property
End Class

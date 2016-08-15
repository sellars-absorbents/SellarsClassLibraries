Imports Microsoft.SqlServer.Types
Imports System.Device.Location

Public Class WarehouseData
    Private _Available As Integer = 0
    Private _Backorder As Boolean = False
    Private _Sequence As Integer = 0
    Private _Warehouse As String = ""
    Private _GeoLoc As SqlGeography
    Private _GeoCoord As GeoCoordinate

    Public Property Available As Integer
        Get
            Return _Available
        End Get
        Set(value As Integer)
            _Available = value
        End Set
    End Property

    Public Property Backorder As Boolean
        Get
            Return _Backorder
        End Get
        Set(value As Boolean)
            _Backorder = value
        End Set
    End Property

    Public Property GeoCoord() As GeoCoordinate
        Get
            Return _GeoCoord
        End Get
        Set(ByVal value As GeoCoordinate)
            _GeoCoord = value
        End Set
    End Property

    Public Property GeoLoc() As SqlGeography
        Get
            Return _GeoLoc
        End Get
        Set(ByVal value As SqlGeography)
            _GeoLoc = value
        End Set
    End Property

    Public Property Sequence As Integer
        Get
            Return _Sequence
        End Get
        Set(value As Integer)
            _Sequence = value
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

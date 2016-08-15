Imports Sellars.SQL
Public Class GetSettings
    Inherits ClassBase


    Dim settings As New DatabaseSettings

    Public ReadOnly Property User() As String
        Get
            Return settings.User
        End Get
    End Property

    Public ReadOnly Property Password() As String
        Get
            Return settings.Password
        End Get
    End Property

    Public ReadOnly Property Database() As String
        Get
            Return settings.Database
        End Get
    End Property

    Public ReadOnly Property Server() As String
        Get
            Return settings.Server
        End Get
    End Property

    Public ReadOnly Property WindowsConnection() As String
        Get
            Return settings.WindowsConnectionString
        End Get
    End Property

    Public ReadOnly Property DynamicsConnection() As String
        Get
            Return settings.DynamicsConnectionString
        End Get
    End Property
End Class

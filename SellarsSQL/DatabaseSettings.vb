Imports System.Configuration
Imports System.Collections.Specialized

Public Class DatabaseSettings
    Private m_Server As String = ""
    Private m_Catalog As String = ""
    Private m_User As String = ""
    Private m_Password As String = ""

    Private m_DynamicsServer As String = ""
    Private m_DynamicsCatalog As String = ""
    Private m_DynamicsUser As String = ""
    Private m_DynamicsPassword As String = ""

    Public Property Database() As String
        Get
            Return m_Catalog
        End Get
        Set(ByVal Value As String)
            m_Catalog = Value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return m_Password
        End Get
        Set(ByVal Value As String)
            m_Password = Value
        End Set
    End Property

    Public Property Server() As String
        Get
            Return m_Server
        End Get
        Set(ByVal Value As String)
            m_Server = Value
        End Set
    End Property

    Public Property User() As String
        Get
            Return m_User
        End Get
        Set(ByVal Value As String)
            m_User = Value
        End Set
    End Property


    Public Property DynamicsDatabase() As String
        Get
            Return m_DynamicsCatalog
        End Get
        Set(ByVal Value As String)
            m_DynamicsCatalog = Value
        End Set
    End Property

    Public Property DynamicsPassword() As String
        Get
            Return m_DynamicsPassword
        End Get
        Set(ByVal Value As String)
            m_DynamicsPassword = Value
        End Set
    End Property

    Public Property DynamicsServer() As String
        Get
            Return m_DynamicsServer
        End Get
        Set(ByVal Value As String)
            m_DynamicsServer = Value
        End Set
    End Property

    Public Property DynamicsUser() As String
        Get
            Return m_DynamicsUser
        End Get
        Set(ByVal Value As String)
            m_DynamicsUser = Value
        End Set
    End Property

    Public Sub New()
        ' Get the database settings from the configuration file
        Dim mySettings As NameValueCollection
        mySettings = System.Configuration.ConfigurationManager.AppSettings()
        m_Server = mySettings("Server")
        m_User = mySettings("User")
        m_Catalog = mySettings("Database")
        m_Password = mySettings("Password")

        m_DynamicsServer = mySettings("DynamicsServer")
        m_DynamicsUser = mySettings("DynamicsUser")
        m_DynamicsCatalog = mySettings("DynamicsDatabase")
        m_DynamicsPassword = mySettings("DynamicsPassword")
    End Sub

    Public Function DynamicsConnectionString() As String
        Dim strConnection As String
        strConnection = "server=" & DynamicsServer & ";uid=" & DynamicsUser & ";pwd=" & DynamicsPassword & ";database=" & DynamicsDatabase & ";"
        Return strConnection
    End Function

    Public Function WindowsConnectionString() As String
        Dim strConnection As String
        strConnection = "server=" & Server & ";uid=" & User & ";pwd=" & Password & ";database=" & Database & ";"
        Return strConnection
    End Function

    Public Function WebConnectionString() As String
        Dim strConnection As String
        strConnection = "server=" & Server & ";uid=" & User & ";pwd=" & Password & ";database=" & Database & ";pooling=false;"
        Return strConnection
    End Function
End Class

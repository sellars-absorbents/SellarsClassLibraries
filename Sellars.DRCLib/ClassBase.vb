Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Text
Imports MaxUpdateXML

Public MustInherit Class ClassBase

#Region "Singleton - MAX Update"
    Private Shared _maxUpdate As XMLWrapper

    Protected Shared ReadOnly Property MAXUpdate As XMLWrapper
        Get
            If _maxUpdate Is Nothing Then
                Throw New Exception("MAXUpdate must be Initialized before being used")
            End If

            Return _maxUpdate
        End Get
    End Property

    Protected Shared Sub Initialize(ByVal connectionString As String, ByVal companyName As String, ByVal licensePath As String, ByVal logPath As String, ByVal errorReport As Boolean)
        If _maxUpdate Is Nothing Then
            _maxUpdate = New XMLWrapper()
            _maxUpdate.Initialize(connectionString, companyName, licensePath, logPath, errorReport)
            _maxUpdate.SetVisualErrorReportingXML(0)
        End If
    End Sub

    Private Shared _maxUpdateASP As ASPXMLWrapper

    Protected Shared ReadOnly Property MAXUpdateASP As ASPXMLWrapper
        Get
            If _maxUpdateASP Is Nothing Then
                Throw New Exception("MAXUpdateASP must be Initialized before being used")
            End If

            Return _maxUpdateASP
        End Get
    End Property

    Protected Shared Sub InitializeASP(ByVal connectionString As String, ByVal companyName As String, ByVal licensePath As String, ByVal logPath As String, ByVal errorReport As Boolean)
        If _maxUpdateASP Is Nothing Then
            _maxUpdateASP = New ASPXMLWrapper()
            _maxUpdateASP.InitASPXMLWrapper(connectionString, companyName, licensePath, logPath, errorReport)
            _maxUpdateASP.SetVisualErrorReportingXML(0)
        End If
    End Sub
#End Region

    ' Declare varible to hold the database connection string
    Private _ConnectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString
    Private _maxConnectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("Max").ConnectionString
    Private _dynamicsConnectionString As String = ""
    Protected MaxConnection As New SqlConnection()
    Protected DynamicsConnection As New SqlConnection()

    ' Set up a default date constant
    Protected Friend Const DefaultDate As Date = #1/1/2000 12:01:00 AM#
    Protected Friend Const DefaultShipDate As Date = #1/1/2099 12:01:00 AM#

    Public Enum DataSource
        Sellars = 0
        Max = 1
    End Enum

    Protected ReadOnly Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
    End Property

    Protected ReadOnly Property MaxConnectionString() As String
        Get
            Return _maxConnectionString
        End Get
    End Property

    Protected ReadOnly Property DynamicsConnectionString() As String
        Get
            Return _dynamicsConnectionString
        End Get
    End Property

    ' When we initialize the class, get the database connection string
    Public Sub New()
    End Sub

    Protected Sub OpenMaxConnection()
        ' try to open another connection to the max database
        MaxConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
        MaxConnection.Open()
    End Sub

    Protected Sub CloseMaxConnection()
        MaxConnection.Close()
        MaxConnection = Nothing
    End Sub

    Protected Sub OpenDynamicsConnection()
        ' try to open another connection to the max database
        DynamicsConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("Dynamics").ConnectionString)
        DynamicsConnection.Open()
    End Sub

    Protected Sub CloseDynamicsConnection()
        DynamicsConnection.Close()
        DynamicsConnection = Nothing
    End Sub


    ' ==== the following are functions necessary for maxupdate

    ' Routine to set the date correctly for maxupdate routines
    Protected Function MakeDate(ByVal maxDate As Date) As String
        Return maxDate.ToString("yyyy-MM-dd") ' piYear.ToString() + "-" + piMon.ToString().PadLeft(2, "0"c) + "-" + piDay.ToString().PadLeft(2, "0"c)
    End Function

End Class


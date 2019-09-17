Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Text

Public MustInherit Class ClassBase

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
    Protected Function MakeDate(ByRef piYear As Short, ByRef piMon As Short, ByRef piDay As Short) As Long
        ' Action  : Dates are held as long integers in an encoded YYYYMMDD format.
        '           This routine encodes dates in the required format.
        '
        ' Takes   : piYear - 4 digit year portion of date
        '           piMon  - 1 or 2 digit month
        '           piDay  - 1 or 2 digit day
        '
        ' Returns : Encoded date.
        '
        Return piYear * 2 ^ 16 + piMon * 2 ^ 8 + piDay
    End Function

End Class


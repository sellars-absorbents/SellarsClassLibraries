Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class QCTestDevicesClass

    Private _ConnectionString As String = ""

    Public Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
        Set(ByVal Value As String)
            _ConnectionString = Value
        End Set
    End Property

    Public Sub New(ByVal connection As String)
        ConnectionString = connection
    End Sub

    '*********************************************************************
    ' Read() - used to be GetTestDevices
    ' Gets a list of all reels produced for a specific day.
    '*********************************************************************
    Public Function Read() As DataSet
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcDataSet("GetQCTestDevices", "Devices")
    End Function

    '*********************************************************************
    ' Update(ID, Port)
    ' Updates a specific port setting for a QC device
    '*********************************************************************
    Public Sub Update(ByVal ID As Integer, ByVal Port As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@ID", SqlDbType.SmallInt, 0, ID, ParameterDirection.Input)
        oSQL.AddParameter("@Port", SqlDbType.NVarChar, 10, Port, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateQCTestDevicePort")
    End Sub

End Class

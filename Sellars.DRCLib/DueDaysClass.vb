Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class DueDaysClass
    Inherits ClassBase

    Private _ReadError As Boolean = False
    Private _CurrentDueDays As Integer
    Private _CustomerDueDays As Integer

    Public ReadOnly Property ReadError() As Boolean
        Get
            Return _ReadError
        End Get
    End Property

    Public ReadOnly Property CurrentDueDays() As Integer
        Get
            Return _CurrentDueDays
        End Get
    End Property

    Public ReadOnly Property CustomerDueDays() As Integer
        Get
            Return _CustomerDueDays
        End Get
    End Property

    ' Get the settings for a single customer
    Public Sub New(ByVal passDueDaysID As Integer)
        Read(passDueDaysID)
    End Sub

    Public Function Read() As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("ReadAllDueDays")
    End Function

    Private Sub Read(ByVal DueDaysID As Integer)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ID", SqlDbType.SmallInt, 0, DueDaysID, ParameterDirection.Input)

        ' Get the data from the stored procedure
        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadDueDays")
        SetProperties(dr)

        ' Close up the data readers and free up memory
        dr.Close()
        dr = Nothing
    End Sub

    Private Sub SetProperties(ByRef dr As SqlClient.SqlDataReader)
        ' If a record was read, then set the properties
        ' otherwise set them to zero
        If dr.Read() Then
            _CustomerDueDays = dr("CusDueDays")
            _CurrentDueDays = dr("CurDueDays")
        Else
            _CustomerDueDays = 0
            _CurrentDueDays = 0
        End If
    End Sub

End Class

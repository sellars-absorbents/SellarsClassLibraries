Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class FreightSettingsClass
    Inherits ClassBase

    Private _FreightMultiplier As Decimal
    Private _ReadError As Boolean
    Private _State As String

    Public ReadOnly Property FreightMultiplier() As Decimal
        Get
            If _FreightMultiplier = 0 Then
                Return 1
            Else
                Return _FreightMultiplier
            End If
        End Get
    End Property

    Public ReadOnly Property ReadError() As Boolean
        Get
            Return _ReadError
        End Get
    End Property

    Public ReadOnly Property State() As Decimal
        Get
            Return _State
        End Get
    End Property

    ' Get the settings for a single customer
    Public Sub New()
    End Sub

    Public Sub New(ByVal StateAbbreviation As String)
        Read(StateAbbreviation)
    End Sub

    Public Sub Add(ByVal StateAbbreviation As String, ByVal Freight As Decimal)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@State", SqlDbType.NVarChar, 2, StateAbbreviation, ParameterDirection.Input)
        oSQL.AddParameter("@Freight", SqlDbType.Float, 0, Freight, ParameterDirection.Input)

        ' run the stored procedure
        oSQL.RunProc("AddFreightSettings")
    End Sub

    Public Sub Delete(ByVal StateAbbreviation As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@State", SqlDbType.NVarChar, 2, StateAbbreviation, ParameterDirection.Input)

        ' run the stored procedure
        oSQL.RunProc("DeleteFreightSettings")
    End Sub

    ' This will read all the states in the states table to 
    ' be displayed in a drop down list box
    Public Function Read() As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' return the data from the stored procedure
        Return oSQL.RunProcReader("ReadAllStates")
    End Function

    Private Sub Read(ByVal StateAbbreviation As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@State", SqlDbType.NVarChar, 2, StateAbbreviation, ParameterDirection.Input)

        ' return the data from the stored procedure
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadFreightSettings")
            dr.Read()
            If dr("Count") = 0 Then
                _ReadError = True
            End If

            ' Set the properties from the data in the datareader if
            ' there was no error reading the record
            If Not ReadError Then
                ' First move to the next resultset
                dr.NextResult()
                ' Go grab the variables
                If dr.Read() Then
                    _FreightMultiplier = dr("FreightMultiplier")
                End If
            End If
        End Using
    End Sub

    Public Sub Update(ByVal StateAbbreviation As String, ByVal Freight As Decimal)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@State", SqlDbType.NVarChar, 2, StateAbbreviation, ParameterDirection.Input)
        oSQL.AddParameter("@Freight", SqlDbType.Float, 0, Freight, ParameterDirection.Input)

        ' run the stored procedure
        oSQL.RunProc("UpdateFreightSettings")
    End Sub
End Class

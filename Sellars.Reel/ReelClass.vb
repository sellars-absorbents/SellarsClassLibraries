Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class ReelClass
    Private _ConnectionString As String = ""
    Private _Count As Integer
    Private _ProductionTime As Date

    Public ReadOnly Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
    End Property

    Public ReadOnly Property ProductionTime() As Date
        Get
            Return _ProductionTime
        End Get
    End Property

    Public Sub New(ByVal connection As String)
        _ConnectionString = connection
    End Sub

    Public Function BasisWeight(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer) As Integer
        '*********************************************************************
        '
        ' GetLocations()
        '
        ' Gets a list of all locations from the database.
        '*********************************************************************
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Month", SqlDbType.SmallInt, 0, Month, ParameterDirection.Input)
        oSQL.AddParameter("@Day", SqlDbType.SmallInt, 0, Day, ParameterDirection.Input)
        oSQL.AddParameter("@Year", SqlDbType.SmallInt, 0, Year, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim ReelBasisWeight As New Parameter("@BasisWeight", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(ReelBasisWeight)

        ' Run the stored procedure
        oSQL.RunProc("GetReelBasisWeight")

        ' Return the reel count value returned from the stored procedure
        If oSQL.ErrorCode <> 0 Then
            Return -1
        Else
            Return ReelBasisWeight.Value
        End If

    End Function

    '*********************************************************************
    ' Count()
    ' Returns the value stored in the _Count variable.
    '*********************************************************************
    Public Function Count() As Integer
        Return _Count
    End Function

    '*********************************************************************
    ' Count(Location, Line, Month, Day, Year)
    ' Gets a count of all reels produced on a specific day.
    '*********************************************************************
    Public Function Count(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer) As Integer
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Month", SqlDbType.SmallInt, 0, Month, ParameterDirection.Input)
        oSQL.AddParameter("@Day", SqlDbType.SmallInt, 0, Day, ParameterDirection.Input)
        oSQL.AddParameter("@Year", SqlDbType.SmallInt, 0, Year, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim parmReelCount As New Parameter("@Count", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(parmReelCount)

        ' Run the stored procedure
        oSQL.RunProc("GetReelCount")

        ' Return the reel count value returned from the stored procedure
        If oSQL.ErrorCode <> 0 Then
            Return -1
        Else
            Return parmReelCount.Value
        End If

    End Function

    '*********************************************************************
    ' Count(Location, Line, Month, Day, Year, Reel)
    ' Gets the number of rolls and the production time of a specified reel number
    '*********************************************************************
    Public Function Count(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer, ByVal Reel As Integer) As Integer
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Month", SqlDbType.SmallInt, 0, Month, ParameterDirection.Input)
        oSQL.AddParameter("@Day", SqlDbType.SmallInt, 0, Day, ParameterDirection.Input)
        oSQL.AddParameter("@Year", SqlDbType.SmallInt, 0, Year, ParameterDirection.Input)
        oSQL.AddParameter("@Reel", SqlDbType.SmallInt, 0, Reel, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim parmReelCount As New Parameter("@Count", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(parmReelCount)

        Dim ReelProductionTime As New Parameter("@ProductionTime", SqlDbType.SmallDateTime, 0, CDate("01/01/2002 12:01:00 AM"), ParameterDirection.Output)
        oSQL.AddParameter(ReelProductionTime)

        ' Run the stored procedure
        oSQL.RunProc("GetSpecificReelCount")

        _Count = parmReelCount.Value
        _ProductionTime = ReelProductionTime.Value

        ' Return the reel count value returned from the stored procedure
        If oSQL.ErrorCode <> 0 Then
            Return -1
        Else
            Return parmReelCount.Value
        End If
    End Function

    '*********************************************************************
    ' Read()
    ' Gets a list of all reels produced for a specific day.
    '*********************************************************************
    Public Function Read(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer) As DataSet
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Month", SqlDbType.SmallInt, 0, Month, ParameterDirection.Input)
        oSQL.AddParameter("@Day", SqlDbType.SmallInt, 0, Day, ParameterDirection.Input)
        oSQL.AddParameter("@Year", SqlDbType.SmallInt, 0, Year, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcDataSet("GetReels", "Reels")
    End Function

    Public Function Read(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer, ByVal Reel As Integer) As SqlDataReader
        '*********************************************************************
        '
        ' GetLocations()
        '
        ' Gets a list of all locations from the database.
        '*********************************************************************
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Month", SqlDbType.SmallInt, 0, Month, ParameterDirection.Input)
        oSQL.AddParameter("@Day", SqlDbType.SmallInt, 0, Day, ParameterDirection.Input)
        oSQL.AddParameter("@Year", SqlDbType.SmallInt, 0, Year, ParameterDirection.Input)
        oSQL.AddParameter("@Reel", SqlDbType.SmallInt, 0, Reel, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetSpecificReelTimes")

    End Function

    '*********************************************************************
    ' Product(Location, Line, Month, Day, Year)
    ' Returns the product contained in a specific reel
    '*********************************************************************
    Public Function Product(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer) As String
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Month", SqlDbType.SmallInt, 0, Month, ParameterDirection.Input)
        oSQL.AddParameter("@Day", SqlDbType.SmallInt, 0, Day, ParameterDirection.Input)
        oSQL.AddParameter("@Year", SqlDbType.SmallInt, 0, Year, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim ReelProduct As New Parameter("@Product", SqlDbType.NVarChar, 15, CStr(""), ParameterDirection.Output)
        oSQL.AddParameter(ReelProduct)

        ' Run the stored procedure
        oSQL.RunProc("GetReelProduct")

        ' Return the reel count value returned from the stored procedure
        If oSQL.ErrorCode <> 0 Then
            Return ""
        Else
            Return ReelProduct.Value
        End If

    End Function

End Class

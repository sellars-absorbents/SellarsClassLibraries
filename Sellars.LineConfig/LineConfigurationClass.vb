Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Public Class LineConfigurationClass

    Private _ConnectionString As String = ""
    Private _ErrorDescription As String = ""
    Private _Completed As Boolean = False
    Private _BasePaper As String = ""
    Private _CoreSize As Integer = 0
    Private _DueDate As Date
    Private _Notes As String = ""
    Private _OD As Decimal = 0
    Private _Part As String = ""
    Private _Reels As Integer = 0
    Private _Title As String = ""
    Private _ErrorCode As Integer = 0

    Public ReadOnly Property BasePaper() As String
        Get
            Return _BasePaper
        End Get
    End Property

    Public ReadOnly Property Completed() As Boolean
        Get
            Return _Completed
        End Get
    End Property

    Public ReadOnly Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
    End Property

    Public ReadOnly Property CoreSize() As Integer
        Get
            Return _CoreSize
        End Get
    End Property

    Public ReadOnly Property DueDate() As Date
        Get
            Return _DueDate
        End Get
    End Property

    Public ReadOnly Property ErrorCode() As Integer
        Get
            Return _ErrorCode
        End Get
    End Property

    Public ReadOnly Property ErrorDescription() As String
        Get
            Return _ErrorDescription
        End Get
    End Property

    Public ReadOnly Property Notes() As String
        Get
            Return _Notes
        End Get
    End Property

    Public ReadOnly Property OD() As Decimal
        Get
            Return _OD
        End Get
    End Property

    Public ReadOnly Property Part() As String
        Get
            Return _Part
        End Get
    End Property

    Public ReadOnly Property Reels() As Integer
        Get
            Return _Reels
        End Get
    End Property

    Public ReadOnly Property Title() As String
        Get
            Return _Title
        End Get
    End Property

    Public Sub New(ByVal connection As String)
        _ConnectionString = connection
    End Sub

    '*********************************************************************
    ' AllowPrint(LocationID, LineID)
    ' Gets the print labels indicator for a specific line
    '*********************************************************************
    Public Function AllowPrint(ByVal LocationID As Integer, ByVal LineID As Integer) As Boolean
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim Print As Parameter = New Parameter("@Print", SqlDbType.Bit, 0, CBool(0), ParameterDirection.Output)
        oSQL.AddParameter(Print)

        ' Run the stored procedure and return a datareader
        oSQL.RunProc("CheckLineConfigurationPrint")

        ' Return the stored procedure output parameter values
        If IsDBNull(Print.Value) Then
            Return False
        Else
            Return Print.Value
        End If
    End Function

    Public Function IncrementReelCount(ByVal LocationID As Integer, ByVal LineID As Integer) As Integer
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim ReelCount As New Parameter("@ReelCount", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(ReelCount)

        ' Run the stored procedure
        oSQL.RunProc("IncrementReelCount")

        ' Set the error description
        _ErrorDescription = oSQL.ErrorDescription
        _ErrorCode = oSQL.ErrorCode

        If IsDBNull(ReelCount.Value) Then
            Return -1
        Else
            Return ReelCount.Value
        End If

    End Function

    Public Function DecrementReelCount(ByVal LocationID As Integer, ByVal LineID As Integer) As Integer
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim ReelCount As New Parameter("@ReelCount", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(ReelCount)

        ' Run the stored procedure
        oSQL.RunProc("DecrementReelCount")

        ' Set the error description
        _ErrorDescription = oSQL.ErrorDescription
        _ErrorCode = oSQL.ErrorCode

        If IsDBNull(ReelCount.Value) Then
            Return -1
        Else
            Return ReelCount.Value
        End If

    End Function

    '*********************************************************************
    ' Read(LocationID)
    ' Gets a list of all lines from a specific plant
    '*********************************************************************
    Public Function Read(ByVal LocationID As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetProductionLines")
    End Function

    '*********************************************************************
    ' Read(LocationID)
    ' Gets a list of all lines from a specific plant and returns a dataset
    '*********************************************************************
    Public Function ReadDS(ByVal LocationID As Integer) As DataSet
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcDataSet("GetProductionLines", "Lines")
    End Function


    '*********************************************************************
    ' Weights(LocationID, LineID)
    ' Gets a list of all basis weights for a specific line
    '*********************************************************************
    Public Function Weights(ByVal LocationID As Integer, ByVal LineID As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetBasisWeights")
    End Function

    '*********************************************************************
    ' Patterns(LocationID, LineID, Weight, Pattern, Color, Grade)
    ' Gets a list of all widths for a specific line
    '*********************************************************************
    Public Function Widths(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal Weight As Decimal, ByVal Pattern As Integer, ByVal Color As Integer, ByVal Grade As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@Weight", SqlDbType.Decimal, 0, Weight, ParameterDirection.Input)
        oSQL.AddParameter("@Pattern", SqlDbType.SmallInt, 0, Pattern, ParameterDirection.Input)
        oSQL.AddParameter("@Color", SqlDbType.SmallInt, 0, Color, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetWidths")
    End Function

End Class

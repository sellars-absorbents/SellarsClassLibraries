Imports Sellars.SQL
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Public Class QCTestsClass
    
    Private _Color As Integer
    Private _ConnectionString As String = ""
    Private _Count As Integer
    Private _Grade As Integer
    Private _GradeDescription As String
    Private _Pattern As Integer
    Private _ProductionTime As Date

    Private _Weight As Decimal
    Private _OpSideBulk As Decimal
    Private _DriveSideBulk As Decimal
    Private _MDT As Decimal
    Private _MDTE As Decimal
    Private _CDTD As Decimal
    Private _CDTW As Decimal
    Private _CDTC As Decimal
    Private _TWA As Decimal
    Private _ZPeel As Decimal
    Private _Contamination As Decimal
    Private _BPDate As Date
    Private _BPReelNo As Integer

    Public ReadOnly Property Pattern() As Integer
        Get
            Return _Pattern
        End Get
    End Property

    Public ReadOnly Property Grade() As Integer
        Get
            Return _Grade
        End Get
    End Property

    Public ReadOnly Property ProductionTime() As Date
        Get
            Return _ProductionTime
        End Get
    End Property

    Public ReadOnly Property Color() As Integer
        Get
            Return _Color
        End Get
    End Property

    Public ReadOnly Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
    End Property

    Public ReadOnly Property GradeDescription() As String
        Get
            Return _GradeDescription
        End Get
    End Property

    Public ReadOnly Property BPDate() As Date
        Get
            Return _BPDate
        End Get
    End Property

    Public ReadOnly Property BPReelNo() As Integer
        Get
            Return _BPReelNo
        End Get
    End Property

    Public ReadOnly Property OpSideBulk() As Decimal
        Get
            Return _OpSideBulk
        End Get
    End Property

    Public ReadOnly Property DriveSideBulk() As Decimal
        Get
            Return _DriveSideBulk
        End Get
    End Property

    Public ReadOnly Property CDTC() As Decimal
        Get
            Return _CDTC
        End Get
    End Property

    Public ReadOnly Property CDTD() As Decimal
        Get
            Return _CDTD
        End Get
    End Property

    Public ReadOnly Property CDTW() As Decimal
        Get
            Return _CDTW
        End Get
    End Property

    Public ReadOnly Property Contamination() As Decimal
        Get
            Return _Contamination
        End Get
    End Property

    Public ReadOnly Property MDT() As Decimal
        Get
            Return _MDT
        End Get
    End Property

    Public ReadOnly Property MDTE() As Decimal
        Get
            Return _MDTE
        End Get
    End Property

    Public ReadOnly Property Weight() As Decimal
        Get
            Return _Weight
        End Get
    End Property

    Public ReadOnly Property TWA() As Decimal
        Get
            Return _TWA
        End Get
    End Property

    Public ReadOnly Property ZPeel() As Decimal
        Get
            Return _ZPeel
        End Get
    End Property

    Public Sub New(ByVal ConnectionString As String)
        _ConnectionString = ConnectionString
    End Sub

    '*********************************************************************
    ' Add(Location, Line, ProductionTime)
    ' Adds a new test result to the database
    '*********************************************************************
    Public Sub Add(ByVal Location As Integer, ByVal Line As Integer, ByVal ProductionTime As Date, ByVal Grade As Integer, ByVal Color As Integer, ByVal Pattern As Integer, ByVal Weight As Decimal, ByVal OpSideBulk As Decimal, ByVal DriveSideBulk As Decimal, ByVal MDDry As Decimal, ByVal MDDryEl As Decimal, ByVal CDDry As Decimal, ByVal CDWet As Decimal, ByVal CDCured As Decimal, ByVal ZPeel As Decimal, ByVal Contamination As Decimal, ByVal TWA As Decimal, ByVal BPDate As Date, ByVal BPReelNo As Integer)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)
        oSQL.AddParameter("@ProductionTime", SqlDbType.SmallDateTime, 0, ProductionTime, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)
        oSQL.AddParameter("@Color", SqlDbType.SmallInt, 0, Color, ParameterDirection.Input)
        oSQL.AddParameter("@Pattern", SqlDbType.SmallInt, 0, Pattern, ParameterDirection.Input)
        oSQL.AddParameter("@Weight", SqlDbType.Float, 0, Weight, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulk", SqlDbType.Float, 0, OpSideBulk, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulk", SqlDbType.Float, 0, DriveSideBulk, ParameterDirection.Input)
        oSQL.AddParameter("@MDT", SqlDbType.Float, 0, MDDry, ParameterDirection.Input)
        oSQL.AddParameter("@MDTE", SqlDbType.Float, 0, MDDryEl, ParameterDirection.Input)
        oSQL.AddParameter("@CDTD", SqlDbType.Float, 0, CDDry, ParameterDirection.Input)
        oSQL.AddParameter("@CDTW", SqlDbType.Float, 0, CDWet, ParameterDirection.Input)
        oSQL.AddParameter("@CDTC", SqlDbType.Float, 0, CDCured, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeel", SqlDbType.Float, 0, ZPeel, ParameterDirection.Input)
        oSQL.AddParameter("@Contamination", SqlDbType.Float, 0, Contamination, ParameterDirection.Input)
        oSQL.AddParameter("@TWA", SqlDbType.Float, 0, TWA, ParameterDirection.Input)
        oSQL.AddParameter("@BPDate", SqlDbType.SmallDateTime, 0, BPDate, ParameterDirection.Input)
        oSQL.AddParameter("@BPReelNo", SqlDbType.SmallInt, 0, BPReelNo, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim TestTime As New Parameter("@TestTime", SqlDbType.SmallDateTime, 0, CDate("1/1/2000 12:01 AM"), ParameterDirection.Output)
        oSQL.AddParameter(TestTime)

        ' Run the stored procedure
        oSQL.RunProc("AddQCTestResult")
    End Sub

    '*********************************************************************
    ' Delete(Location, Line, ProductionTime, TestTime)
    ' Deletes the test results from the QCTestResults table for a specific
    ' test
    '*********************************************************************
    Public Sub Delete(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal ProductionTime As Date, ByVal TestTime As Date)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Initialize the values to zero
        _Weight = 0
        _OpSideBulk = 0
        _DriveSideBulk = 0
        _MDT = 0
        _MDTE = 0
        _CDTD = 0
        _CDTW = 0
        _CDTC = 0
        _TWA = 0
        _ZPeel = 0
        _BPDate = "1/1/2000"
        _BPReelNo = 0
        _Color = 99
        _Pattern = 99
        _Grade = 99
        _Contamination = 0

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@ProductionTime", SqlDbType.SmallDateTime, 0, ProductionTime, ParameterDirection.Input)
        oSQL.AddParameter("@TestTime", SqlDbType.DateTime, 0, TestTime, ParameterDirection.Input)

        ' Execute the stored procedure
        oSQL.RunProc("DeleteSpecificReelTestResults")
    End Sub

    '*********************************************************************
    ' GetReelInfo(Location, Line, ProductionTime)
    ' Gets detail reel information and assigns it to class variables.
    '*********************************************************************
    Public Sub GetReelInfo(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal ProductionTime As Date)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@ProductionTime", SqlDbType.SmallDateTime, 0, ProductionTime, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim ReelGrade As New Parameter("@Grade", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(ReelGrade)

        Dim ReelGradeDescription As New Parameter("@GradeDescription", SqlDbType.NVarChar, 50, CStr(""), ParameterDirection.Output)
        oSQL.AddParameter(ReelGradeDescription)

        Dim ReelColor As New Parameter("@Color", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(ReelColor)

        Dim ReelPattern As New Parameter("@Pattern", SqlDbType.SmallInt, 0, CInt(0), ParameterDirection.Output)
        oSQL.AddParameter(ReelPattern)

        ' Run the stored procedure
        oSQL.RunProc("GetReelInfo")

        _Grade = ReelGrade.Value
        _GradeDescription = ReelGradeDescription.Value
        _Color = ReelColor.Value
        _Pattern = ReelPattern.Value
        _ProductionTime = ProductionTime
    End Sub

    '*********************************************************************
    ' Read(Part)
    ' Reads all the tests from the QCTestResults table which have the
    ' specified grade ID
    '*********************************************************************
    Public Function Read(ByVal Grade As String) As DataSet
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)

        REM -- create the datatable
        Dim ds As DataSet = New DataSet("QCResults")
        Dim dt As DataTable = ds.Tables.Add("History")
        dt.Columns.Add("ProductionTime", Type.GetType("System.DateTime"))
        dt.Columns.Add("Reel", Type.GetType("System.Decimal"))
        dt.Columns.Add("TestTime", Type.GetType("System.DateTime"))
        dt.Columns.Add("Weight", Type.GetType("System.Decimal"))
        dt.Columns.Add("OpSideBulk", Type.GetType("System.Decimal"))
        dt.Columns.Add("DriveSideBulk", Type.GetType("System.Decimal"))
        dt.Columns.Add("MDT", Type.GetType("System.Decimal"))
        dt.Columns.Add("MDTE", Type.GetType("System.Decimal"))
        dt.Columns.Add("CDTD", Type.GetType("System.Decimal"))
        dt.Columns.Add("MDCD", Type.GetType("System.Decimal"))
        dt.Columns.Add("TT", Type.GetType("System.Decimal"))
        dt.Columns.Add("CDTW", Type.GetType("System.Decimal"))
        dt.Columns.Add("CDTWD", Type.GetType("System.Decimal"))
        dt.Columns.Add("CDTC", Type.GetType("System.Decimal"))
        dt.Columns.Add("PercentCured", Type.GetType("System.Decimal"))
        dt.Columns.Add("ZPeel", Type.GetType("System.Decimal"))
        dt.Columns.Add("Contamination", Type.GetType("System.Decimal"))
        dt.Columns.Add("TWA", Type.GetType("System.Decimal"))
        dt.Columns.Add("BPDate", Type.GetType("System.DateTime"))
        dt.Columns.Add("BPReelNo", Type.GetType("System.Decimal"))

        Dim drow As DataRow

        Using dr As SqlDataReader = oSQL.RunProcReader("GetAllQCTestResults")
            While dr.Read()
                REM -- Add to DataSet ds
                drow = dt.NewRow()
                drow("ProductionTime") = dr("ProductionTime")
                drow("Reel") = dr("ReelCount")
                drow("TestTime") = dr("TestTime")
                drow("Weight") = dr("Weight")
                drow("OpSideBulk") = dr("OpSideBulk")
                drow("DriveSideBulk") = dr("DriveSideBulk")
                drow("MDT") = dr("MDT")
                drow("MDTE") = dr("MDTE")
                drow("CDTD") = dr("CDTD")
                If dr("CDTD") = 0 Then
                    drow("MDCD") = 0
                Else
                    drow("MDCD") = dr("MDT") / dr("CDTD")
                End If
                drow("TT") = dr("MDT") + dr("CDTD")
                drow("CDTW") = dr("CDTW")
                If dr("CDTD") = 0 Then
                    drow("CDTWD") = 0
                Else
                    drow("CDTWD") = dr("CDTW") / dr("CDTD")
                End If
                drow("CDTC") = dr("CDTC")
                If dr("CDTC") = 0 Then
                    drow("PercentCured") = 0
                Else
                    drow("PercentCured") = dr("CDTW") / dr("CDTC")
                End If
                drow("ZPeel") = dr("ZPeel")
                drow("Contamination") = dr("Contamination")
                drow("TWA") = dr("TWA")
                drow("BPDate") = dr("BPDate")
                drow("BPReelNo") = dr("BPReelNo")
                dt.Rows.Add(drow)
            End While
        End Using

        Return ds
    End Function

    '*********************************************************************
    ' Read(Location, Line, ProductionTime)
    ' Reads the test results from the QCTestResults table for a specified
    ' reel production time
    '*********************************************************************
    Public Function Read(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal ProductionTime As Date) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@ProductionTime", SqlDbType.SmallDateTime, 0, ProductionTime, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetSpecificReelTests")
    End Function


    '*********************************************************************
    ' Read(Location, Line, ProductionTime, TestTime)
    ' Reads the test results from the QCTestResults table for a specific
    ' test
    '*********************************************************************
    Public Sub Read(ByVal LocationID As Integer, ByVal LineID As Integer, ByVal ProductionTime As Date, ByVal TestTime As Date)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Initialize the values to zero
        _Weight = 0
        _OpSideBulk = 0
        _DriveSideBulk = 0
        _MDT = 0
        _MDTE = 0
        _CDTD = 0
        _CDTW = 0
        _CDTC = 0
        _TWA = 0
        _ZPeel = 0
        _Contamination = 0
        _BPDate = "1/1/2000"
        _BPReelNo = 0
        _Color = 99
        _Pattern = 99
        _Grade = 99

        ' Add the parameter to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)
        oSQL.AddParameter("@ProductionTime", SqlDbType.SmallDateTime, 0, ProductionTime, ParameterDirection.Input)
        oSQL.AddParameter("@TestTime", SqlDbType.DateTime, 0, TestTime, ParameterDirection.Input)

        Using dr As SqlDataReader = oSQL.RunProcReader("GetSpecificReelTestResults")
            ' Read the record
            While dr.Read()
                _Color = dr("ColorID")
                _Pattern = dr("PatternID")
                _Grade = dr("GradeID")
                _Weight = dr("Weight")
                _OpSideBulk = dr("OpSideBulk")
                _DriveSideBulk = dr("DriveSideBulk")
                _MDT = dr("MDT")
                _MDTE = dr("MDTE")
                _CDTD = dr("CDTD")
                _CDTW = dr("CDTW")
                _CDTC = dr("CDTC")
                _ZPeel = IIf(IsDBNull(dr("ZPeel")), 0, dr("ZPeel"))
                _Contamination = dr("Contamination")
                _TWA = dr("TWA")
                _BPDate = IIf(IsDBNull(dr("BPDate")), "1/1/2000", dr("BPDate"))
                _BPReelNo = IIf(IsDBNull(dr("BPReelNo")), 0, dr("BPReelNo"))
            End While
        End Using
    End Sub

    '*********************************************************************
    ' Read()
    ' Reads all the tests from the QCTestResults table
    '*********************************************************************
    Public Function Read(ByVal Location As Integer, ByVal Line As Integer, ByVal Grade As Integer, ByVal Color As Integer, ByVal Pattern As Integer, ByVal Test As String, ByVal TestCount As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)
        oSQL.AddParameter("@Color", SqlDbType.SmallInt, 0, Color, ParameterDirection.Input)
        oSQL.AddParameter("@Pattern", SqlDbType.SmallInt, 0, Pattern, ParameterDirection.Input)
        oSQL.AddParameter("@TestType", SqlDbType.NVarChar, 6, Test, ParameterDirection.Input)
        oSQL.AddParameter("@Tests", SqlDbType.SmallInt, 0, TestCount, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("ReadTestResults")
    End Function

    '*************************************************************************
    ' ReadParts(Location, Line, StartDate, EndDate)
    ' Reads all the grades from the QCTestRsults table for a given time period
    '*************************************************************************
    Public Function ReadGrades(ByVal Location As Integer, ByVal Line As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetTestGrades")
    End Function

    '*************************************************************************
    ' ReadColors(Location, Line, Grade)
    ' Reads all the colors from the QCTestResults table for the specified grade
    '*************************************************************************
    Public Function ReadColors(ByVal Location As Integer, ByVal Line As Integer, ByVal Grade As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetTestColors")
    End Function

    '*************************************************************************
    ' ReadColors(Location, Line, Grade)
    ' Reads all the colors from the QCTestResults table for the specified grade
    '*************************************************************************
    Public Function ReadPatterns(ByVal Location As Integer, ByVal Line As Integer, ByVal Grade As Integer, ByVal Color As Integer) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)
        oSQL.AddParameter("@Color", SqlDbType.SmallInt, 0, Color, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("GetTestPatterns")
    End Function

    '*********************************************************************
    ' Update(Location, Line, ProductionTime)
    ' Updates a test result to the database
    '*********************************************************************
    Public Sub Update(ByVal Location As Integer, ByVal Line As Integer, ByVal ProductionTime As Date, ByVal TestTime As Date, ByVal Weight As Decimal, ByVal OpSideBulk As Decimal, ByVal DriveSideBulk As Decimal, ByVal MDT As Decimal, ByVal MDTE As Decimal, ByVal CDTD As Decimal, ByVal CDTW As Decimal, ByVal CDTC As Decimal, ByVal ZPeel As Decimal, ByVal Contamination As Decimal, ByVal TWA As Decimal, ByVal BPDate As Date, ByVal BPReelNo As Integer)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)
        oSQL.AddParameter("@ProductionTime", SqlDbType.SmallDateTime, 0, ProductionTime, ParameterDirection.Input)
        oSQL.AddParameter("@TestTime", SqlDbType.DateTime, 0, TestTime, ParameterDirection.Input)
        oSQL.AddParameter("@Weight", SqlDbType.Float, 0, Weight, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulk", SqlDbType.Float, 0, OpSideBulk, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulk", SqlDbType.Float, 0, DriveSideBulk, ParameterDirection.Input)
        oSQL.AddParameter("@MDT", SqlDbType.Float, 0, MDT, ParameterDirection.Input)
        oSQL.AddParameter("@MDTE", SqlDbType.Float, 0, MDTE, ParameterDirection.Input)
        oSQL.AddParameter("@CDTD", SqlDbType.Float, 0, CDTD, ParameterDirection.Input)
        oSQL.AddParameter("@CDTW", SqlDbType.Float, 0, CDTW, ParameterDirection.Input)
        oSQL.AddParameter("@CDTC", SqlDbType.Float, 0, CDTC, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeel", SqlDbType.Float, 0, ZPeel, ParameterDirection.Input)
        oSQL.AddParameter("@Contamination", SqlDbType.Float, 0, Contamination, ParameterDirection.Input)
        oSQL.AddParameter("@TWA", SqlDbType.Float, 0, TWA, ParameterDirection.Input)
        oSQL.AddParameter("@BPDate", SqlDbType.SmallDateTime, 0, BPDate, ParameterDirection.Input)
        oSQL.AddParameter("@BPReelNo", SqlDbType.SmallInt, 0, BPReelNo, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateQCTestResult")
    End Sub

    '*********************************************************************
    ' Export(location, line, start date, end date)
    ' will export all the test results from a location line to an excel file
    '*********************************************************************
    Public Function Export(ByVal Location As Integer, ByVal Line As Integer, ByVal StartDate As Date, ByVal EndDate As Date, ByVal UserName As String) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, Location, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, Line, ParameterDirection.Input)
        oSQL.AddParameter("@StartDate", SqlDbType.SmallDateTime, 0, StartDate, ParameterDirection.Input)
        oSQL.AddParameter("@EndDate", SqlDbType.SmallDateTime, 0, EndDate, ParameterDirection.Input)

        ' Run the stored procedure and return a dataset of all the records
        Return oSQL.RunProcReader("GetQCTestResults")
    End Function
End Class

Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class BusinessDays
    Private _ConnectionString As String = ""

    Public Sub New(ByVal ConnectionString As String)
        _ConnectionString = ConnectionString
    End Sub

    Public Function Add(ByVal CompanyID As Integer, ByVal StartDate As Date, ByVal DaysToAdd As Integer) As Date
        Dim rtnData As Date = New Date(2050, 12, 31)

        Dim cmd As SqlCommand = Nothing
        Dim conn As SqlConnection = New SqlConnection(_ConnectionString)
        Try
            conn.Open()

            ' Declare the SQL data layer class
            cmd = New System.Data.SqlClient.SqlCommand("AddBusinessDays", conn)
            cmd.CommandType = CommandType.StoredProcedure

            ' Add the parameters to the command object
            cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", CompanyID))
            cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@StartDate", StartDate))
            cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DaysToAdd", DaysToAdd))

            Dim prmResult As New SqlParameter("@Result", SqlDbType.Date, 0)
            prmResult.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmResult)

            ' Run the stored procedure
            cmd.ExecuteNonQuery()

            ' Get the date returned from the function call
            rtnData = prmResult.Value

        Catch ex As Exception
        Finally
            ' Close the database connection
            conn.Close()
        End Try

        If cmd IsNot Nothing Then
            cmd.Dispose()
        End If

        ' return the list of invoices to process
        Return rtnData
    End Function

End Class

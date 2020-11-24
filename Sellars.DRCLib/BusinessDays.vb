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

        Using conn As SqlConnection = New SqlConnection(_ConnectionString)
            conn.Open()

            ' Declare the SQL data layer class
            Using cmd As SqlCommand = New System.Data.SqlClient.SqlCommand("AddBusinessDays", conn)
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
            End Using
        End Using

        ' return the list of invoices to process
        Return rtnData
    End Function

    Public Shared Function IsBusinessDay(ByVal shopfloorConnection As String, ByVal CheckDate As Date) As Boolean
        Dim IBD As Boolean = False
        ' Get all orders for all the sellars warehouses as defined by the SellarsFulfillmentWarehouses table
        Dim strSQL As String = "SELECT @IsBusinessDay = dbo.IsBusinessDay(2, @CheckDate)"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(strSQL, connection)
                command.CommandType = Data.CommandType.Text
                command.CommandTimeout = 0
                command.Parameters.Add(New SqlParameter("@CheckDate", CheckDate))

                ' Set the fastenal pickup day to be the next business day from today
                Dim parmIBD As SqlParameter = New SqlParameter("@IsBusinessDay", Data.SqlDbType.Bit)
                parmIBD.Direction = Data.ParameterDirection.Output
                parmIBD.Value = Nothing
                command.Parameters.Add(parmIBD)

                ' Actually execute the query
                command.ExecuteNonQuery()

                ' Get the value of the output parameter
                IBD = CBool(parmIBD.Value)
            End Using
        End Using

        Return IBD
    End Function

    Public Shared Function BusinessDaysAdd(ByVal shopfloorConnection As String, ByVal StartDate As Date, ByVal Days As Integer) As Date
        Dim nextBusinessDay As Date = New DateTime(2050, 12, 31)
        ' Get all orders for all the sellars warehouses as defined by the SellarsFulfillmentWarehouses table
        Dim strSQL As String = "SELECT @NextBusinessDay = dbo.BusinessDaysAdd(2, @StartDate, @Days)"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(strSQL, connection)
                command.CommandType = Data.CommandType.Text
                command.CommandTimeout = 0
                command.Parameters.Add(New SqlParameter("@Days", Days))
                command.Parameters.Add(New SqlParameter("@StartDate", StartDate))

                ' Set the fastenal pickup day to be the next business day from today
                Dim parmNBD As SqlParameter = New SqlParameter("@NextBusinessDay", Data.SqlDbType.DateTime)
                parmNBD.Direction = Data.ParameterDirection.Output
                parmNBD.Value = Nothing
                command.Parameters.Add(parmNBD)

                ' Actually execute the query
                command.ExecuteNonQuery()

                ' Get the value of the output parameter
                nextBusinessDay = CType(parmNBD.Value, Date)
            End Using
        End Using

        Return nextBusinessDay
    End Function
End Class

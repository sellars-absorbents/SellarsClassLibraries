Imports System.Data.SqlClient

Public Class ProcessLog
    Public Shared Function Start(ByVal connectionString As String, ByVal processNameID As Integer, ByVal processStep As String, ByVal application As String, ByVal startTime As DateTime) As Integer
        Dim resultID As Integer = 0

        Using connection As SqlConnection = New SqlConnection(connectionString)
            connection.Open()

            Dim sql As String = "INSERT INTO [dbo].[ProcessLog]
                                       ([ProcessNameID]
                                       ,[ProcessStep]
                                       ,[Application]
                                       ,[StartTime])
                                 VALUES
                                       (@ProcessNameID,
                                       @ProcessStep,
                                       @Application,
                                       @StartTime)

                                select max(ID) from ProcessLog"

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ProcessNameID", processNameID))
                command.Parameters.Add(New SqlParameter("@ProcessStep", processStep))
                command.Parameters.Add(New SqlParameter("@Application", application))
                command.Parameters.Add(New SqlParameter("@StartTime", startTime))

                Dim resultq As Object = command.ExecuteScalar()

                If (resultq IsNot Nothing AndAlso resultq IsNot DBNull.Value) Then
                    resultID = Convert.ToInt32(resultq)
                End If
            End Using
        End Using

        Return resultid
    End Function

    Public Shared Sub Complete(ByVal connectionString As String, ByVal processLogID As Integer, ByVal endTime As DateTime)
        Using connection As SqlConnection = New SqlConnection(connectionString)
            connection.Open()

            Dim sql As String = "update ProcessLog
                                 set EndTime = @EndTime
                                 where ID = @ID"

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@EndTime", endTime))
                command.Parameters.Add(New SqlParameter("@ID", processLogID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class

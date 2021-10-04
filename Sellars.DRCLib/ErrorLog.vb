Imports System.Configuration
Imports System.Data.SqlClient

Public Class ErrorLog
    Public Shared Sub Write(ByVal connectionString As String, ByVal message As String, ByVal subroutine As String, ByVal sourceModule As String, ByVal stackTrace As String)
        Dim truncStackTrace As String = stackTrace

        If (stackTrace.Length > 2001) Then
            truncStackTrace = stackTrace.Substring(0, 2000)
        End If

        Using connection As SqlConnection = New SqlConnection(connectionString)
            connection.Open()

            Dim sql As String = "INSERT INTO [dbo].[ErrorLog]
                                       ([Date]
                                       ,[Module]
                                       ,[Subroutine]
                                       ,[Error]
                                       ,[StackTrace])
                                 VALUES
                                       (getdate(),
                                       @Module,
                                       @Subroutine,
                                       @ErrorMessage,
                                       @StackTrace)"

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ErrorMessage", message))
                command.Parameters.Add(New SqlParameter("@Subroutine", subroutine))
                command.Parameters.Add(New SqlParameter("@Module", sourceModule))
                command.Parameters.Add(New SqlParameter("@StackTrace", truncStackTrace))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Sub Write(ByVal connectionString As String, ByVal subroutine As String, ByVal sourceModule As String, ByVal ex As Exception)
        Dim truncStackTrace As String = ""

        If ex.StackTrace IsNot Nothing Then
            truncStackTrace = ex.StackTrace
        End If

        If (truncStackTrace.Length > 2001) Then
            truncStackTrace = truncStackTrace.Substring(0, 2000)
        End If

        Dim currentException As Exception = ex
        Dim message As String = ex.Message

        While currentException.InnerException IsNot Nothing
            currentException = currentException.InnerException
            message += " | " + currentException.Message
        End While

        Using connection As SqlConnection = New SqlConnection(connectionString)
            connection.Open()

            Dim sql As String = "INSERT INTO [dbo].[ErrorLog]
                                       ([Date]
                                       ,[Module]
                                       ,[Subroutine]
                                       ,[Error]
                                       ,[StackTrace])
                                 VALUES
                                       (getdate(),
                                       @Module,
                                       @Subroutine,
                                       @ErrorMessage,
                                       @StackTrace)"

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ErrorMessage", message))
                command.Parameters.Add(New SqlParameter("@Subroutine", subroutine))
                command.Parameters.Add(New SqlParameter("@Module", sourceModule))
                command.Parameters.Add(New SqlParameter("@StackTrace", truncStackTrace))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class

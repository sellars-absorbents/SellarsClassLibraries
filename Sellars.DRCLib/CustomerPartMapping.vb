Imports System.Data.SqlClient

Public Class CustomerPartMapping
    Public Shared Function GetInternalPartNumber(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal externalPartNumber As String) As String
        Dim sql As String = "select InternalPartNumber
                                from CustomerOrderPartMappings
                                where CustomerID = @CustomerID
                                and ExternalPartNumber = @ExternalPartNumber
                                and StartDate <= @TargetDate
                                and EndDate >= @TargetDate"
        Dim result As String = ""

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@ExternalPartNumber", externalPartNumber))
                command.Parameters.Add(New SqlParameter("@TargetDate", Date.Now))

                Using dr As SqlDataReader = command.ExecuteReader()
                    If dr.Read() Then
                        result = dr(0).ToString().Trim()
                    End If
                End Using
            End Using
        End Using

        Return result
    End Function

    Public Shared Sub InsertMapping(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal customerPONumber As String, ByVal externalPartNumber As String, ByVal internalPartNumber As String)
        Dim sql As String = "insert into CustomerPOPartMappings select @CustomerID, @CustomerPONumber, @ExternalPartNumber, @InternalPartNumber, @CreatedOn"

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@CustomerPONumber", customerPONumber))
                command.Parameters.Add(New SqlParameter("@ExternalPartNumber", externalPartNumber))
                command.Parameters.Add(New SqlParameter("@InternalPartNumber", internalPartNumber))
                command.Parameters.Add(New SqlParameter("@CreatedOn", Date.Now))

                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Function GetExternalPartNumber(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal customerPONumber As String, ByVal internalPartNumber As String) As String
        Dim sql As String = "select ExternalPartNumber
                                from CustomerPOPartMappings
                                where CustomerID = @CustomerID
                                and CustomerPONumber = @CustomerPONumber
                                and InternalPartNumber = @InternalPartNumber"
        Dim result As String = ""

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@CustomerPONumber", customerPONumber))
                command.Parameters.Add(New SqlParameter("@InternalPartNumber", internalPartNumber))

                Using dr As SqlDataReader = command.ExecuteReader()
                    If dr.Read() Then
                        result = dr(0).ToString().Trim()
                    End If
                End Using
            End Using
        End Using

        Return result
    End Function
End Class

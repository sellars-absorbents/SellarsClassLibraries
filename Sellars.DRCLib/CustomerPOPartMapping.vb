Imports System.Data.SqlClient

Public Class CustomerPOPartMapping
    Public Shared Function GetInternalPartNumber(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal customerPONumber As String, ByVal externalPartNumber As String) As String
        Dim sql As String = "select InternalPartNumber
                                from CustomerPOPartMappings
                                where CustomerID = @CustomerID
                                and CustomerPONumber = @CustomerPONumber
                                and ExternalPartNumber = @ExternalPartNumber"
        Dim result As String = ""

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@CustomerPONumber", customerPONumber))
                command.Parameters.Add(New SqlParameter("@ExternalPartNumber", externalPartNumber))

                Using dr As SqlDataReader = command.ExecuteReader()
                    If dr.Read() Then
                        result = dr(0).ToString()
                    End If
                End Using
            End Using
        End Using

        Return result
    End Function

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
                        result = dr(0).ToString()
                    End If
                End Using
            End Using
        End Using

        Return result
    End Function
End Class

Imports System.Data.SqlClient

Public Class CustomerPartMapping
    Public Property InternalPartNumber As String = ""
    Public Property SalesUOM As String = ""
    Public Property SalesConversion As Double = 1

    Public Shared Function GetInternalPartNumber(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal externalPartNumber As String) As CustomerPartMapping
        Dim result As CustomerPartMapping = New CustomerPartMapping()
        Dim sql As String = "select InternalPartNumber,
                                    SalesUOM,
                                    SalesConversion
                                from CustomerOrderPartMappings
                                where CustomerID = @CustomerID
                                and ExternalPartNumber = @ExternalPartNumber
                                and StartDate <= @TargetDate
                                and EndDate >= @TargetDate"

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@ExternalPartNumber", externalPartNumber))
                command.Parameters.Add(New SqlParameter("@TargetDate", Date.Now))

                Using dr As SqlDataReader = command.ExecuteReader()
                    If dr.Read() Then
                        result.InternalPartNumber = dr("InternalPartNumber").ToString().Trim()
                        result.SalesUOM = dr("SalesUOM").ToString().Trim()
                        result.SalesConversion = Convert.ToDouble(dr("SalesConversion"))
                    End If
                End Using
            End Using
        End Using

        Return result
    End Function

    Public Shared Sub InsertMapping(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal customerPONumber As String, ByVal externalPartNumber As String, ByVal internalPartNumber As String, ByVal salesUOM As String, ByVal salesConversion As Double)
        Dim sql As String = "insert into CustomerPOPartMappings select @CustomerID, @CustomerPONumber, @ExternalPartNumber, @InternalPartNumber, @SalesUOM, @SalesConversion, @CreatedOn"

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@CustomerPONumber", customerPONumber))
                command.Parameters.Add(New SqlParameter("@ExternalPartNumber", externalPartNumber))
                command.Parameters.Add(New SqlParameter("@InternalPartNumber", internalPartNumber))
                command.Parameters.Add(New SqlParameter("@SalesUOM", salesUOM))
                command.Parameters.Add(New SqlParameter("@SalesConversion", salesConversion))
                command.Parameters.Add(New SqlParameter("@CreatedOn", Date.Now))

                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Function GetExternalPartNumber(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal customerPONumber As String, ByVal internalPartNumber As String) As CustomerPartMapping
        Dim result As CustomerPartMapping = New CustomerPartMapping()
        Dim sql As String = "select ExternalPartNumber,
                                    SalesUOM,
                                    SalesConversion
                                from CustomerPOPartMappings
                                where CustomerID = @CustomerID
                                and CustomerPONumber = @CustomerPONumber
                                and InternalPartNumber = @InternalPartNumber"

        Using connection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@CustomerPONumber", customerPONumber))
                command.Parameters.Add(New SqlParameter("@InternalPartNumber", internalPartNumber))

                Using dr As SqlDataReader = command.ExecuteReader()
                    If dr.Read() Then
                        result.InternalPartNumber = dr("InternalPartNumber").ToString().Trim()
                        result.SalesUOM = dr("SalesUOM").ToString().Trim()
                        result.SalesConversion = Convert.ToDouble(dr("SalesConversion"))
                    End If
                End Using
            End Using
        End Using

        Return result
    End Function
End Class

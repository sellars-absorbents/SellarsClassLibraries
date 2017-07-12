Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data.SqlClient

Public Class SalesOrderPromotionCodes

    Public ErrorMessage As String = ""

    Public Function Add(ByVal ORDNUM As String, ByVal PromotionCode As String) As Boolean
        Dim rtnValue As Boolean = False
        ErrorMessage = ""

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "insert into SalesOrderPromotionCodes SELECT @ORDNUM, @PromotionCode"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
            cmd.Parameters.Add(New SqlParameter("@PromotionCode", PromotionCode))

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnValue = True
            Catch ex As Exception
                ErrorMessage = ex.Message
                rtnValue = False
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        ' Return whether the insert was successful or not
        Return rtnValue
    End Function

    Public Function Delete(ByVal ORDNUM As String, ByVal PromotionCode As String) As Boolean
        Dim rtnValue As Boolean = False
        ErrorMessage = ""

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "delete from SalesOrderPromotionCodes where ORDNUM = @ORDNUM and PromotionCode = @PromotionCode"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
            cmd.Parameters.Add(New SqlParameter("@PromotionCode", PromotionCode))

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnValue = True
            Catch ex As Exception
                ErrorMessage = ex.Message
                rtnValue = False
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        ' Return whether the insert was successful or not
        Return rtnValue
    End Function

    Public Function IsValid(ByVal PromotionCode As String, ByVal OrderDate As Date) As Boolean
        Dim rtnValue As Boolean = False

        ErrorMessage = ""

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "select @Count = count(*) from RetailMonthlyPromotionsPricing where PromotionNumber = @PromotionCode and @OrderDate between StartDate and EndDate"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PromotionCode", PromotionCode))
            cmd.Parameters.Add(New SqlParameter("@OrderDate", OrderDate))

            ' a count parameter to see how many fastenal orders are waiting 
            Dim ParmCount As New SqlParameter("@Count", SqlDbType.SmallInt)
            ParmCount.Direction = ParameterDirection.Output
            cmd.Parameters.Add(ParmCount)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()

                If ParmCount.Value > 0 Then
                    rtnValue = True
                Else
                    rtnValue = False
                End If

            Catch ex As Exception
                ErrorMessage = ex.Message
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnValue
    End Function

    Public Function Read(ByVal ORDNUM As String) As List(Of String)
        Dim rtnData As New List(Of String)
        ErrorMessage = ""

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "select PromotionCode from SalesOrderPromotionCodes where ORDNUM = @ORDNUM order by PromotionCode"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))

            'Perform the update to the database
            Try
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While dr.Read()
                    rtnData.Add(dr("PromotionCode"))
                End While
            Catch ex As Exception
                ErrorMessage = ex.Message
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        ' Return whether the insert was successful or not
        Return rtnData
    End Function

    Public Function PromoSameAsPurchased(ByVal PromotionCode As String) As Boolean
        Dim rtnData As Boolean = False
        ErrorMessage = ""

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "select @FreeSameAsPurchased = isnull(FreeSameAsPurchased, 0) from RetailMonthlyPromotionsPricing where PromotionNumber = @PromotionNumber"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PromotionNumber", PromotionCode))

            ' a count parameter to see how many fastenal orders are waiting 
            Dim ParmSameAs As New SqlParameter("@FreeSameAsPurchased", SqlDbType.Bit)
            ParmSameAs.Direction = ParameterDirection.Output
            cmd.Parameters.Add(ParmSameAs)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnData = ParmSameAs.Value
            Catch ex As Exception
                ErrorMessage = ex.Message
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        ' Return whether the insert was successful or not
        Return rtnData
    End Function

    Public Function GetFreeProducts(ByVal PromotionNumber As String) As List(Of String)
        Dim rtnData As New List(Of String)

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Dim dr As SqlDataReader = Nothing
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT PurchasePartNumbers, FreeSameAsPurchased from RetailMonthlyPromotionsPricing where PromotionNumber = @PromotionNumber and Type = 2"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PromotionNumber", PromotionNumber))

            'Perform the update to the database
            Try
                dr = cmd.ExecuteReader()

                While dr.Read()
                    Dim itemArray As Array = dr("PurchasePartNumbers").ToString().Split(",")
                    For Each prt As String In itemArray
                        rtnData.Add(prt.Trim)
                    Next
                End While
            Catch ex As Exception
                Dim err As String = ex.Message
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnData
    End Function

End Class

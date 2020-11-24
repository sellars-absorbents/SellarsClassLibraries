Imports System.Collections.Generic
Imports System.Data.SqlClient

Public Class PalletDetails
    Public Property PalletNumber As Integer = 0
    Public Property LayerNumber As Integer = 0
    Public Property OrderNumber As String = ""
    Public Property PRTNUM As String = ""
    Public Property LineNumber As String = ""
    Public Property OrderQuantity As Integer = 0
    Public Property Quantity As Integer = 0
    Public Property OriginalPrice As Decimal = 0.0D
    Public Property Discount As Decimal = 0.0D
    Public Property DiscountedPrice As Decimal = 0.0D
    Public Property CustomerPartNumber As String = ""
    Public Property PalletsPerTruck As Integer = 0

    ' Here we save the final order, and pallet information
    ' needed to split the items if they go over a full truckload (i.e. 30 pallets)
    Public Property OrderSequence As Integer = 0
    Public Property FinalPalletNumber As Integer = 0

    Public Shared Sub InsertSalesOrderPallets(ByVal shopfloorConnection As String, ByVal palletDetails As List(Of PalletDetails), ByVal customerID As String)
        Dim lastSequence As Integer = GetLastBarCodeSequence(shopfloorConnection, customerID)

        For Each orderNumber As String In palletDetails.[Select](Function(x) x.OrderNumber).Distinct()
            ' Get the list of all the pallets in the order sequence for the current Order Number
            Dim orderPalletDetails = From plt In palletDetails Where plt.OrderNumber Is orderNumber Order By plt.FinalPalletNumber Select plt

            Using conn As SqlConnection = New SqlConnection(shopfloorConnection)
                conn.Open()
                Dim orderPalletSQL As String = "insert into OrderPallet (OrderNumber, PalletNumber, BarcodeData) 
                                        Values (@OrderNumber, @PalletNumber, @BarcodeData)"

                For Each palletNumber As Integer In orderPalletDetails.[Select](Function(x) x.FinalPalletNumber).Distinct()
                    Dim barCode As String = CalculateBarCode(lastSequence, palletNumber.ToString())

                    Using command As SqlCommand = New SqlCommand(orderPalletSQL, conn)
                        command.Parameters.Add(New SqlParameter("@OrderNumber", orderNumber))
                        command.Parameters.Add(New SqlParameter("@PalletNumber", palletNumber))
                        command.Parameters.Add(New SqlParameter("@BarcodeData", barCode))
                        command.ExecuteNonQuery()
                    End Using
                Next

                Dim orderPalletLayerSQL As String = "insert into OrderPalletLayer (OrderNumber, PalletNumber, LayerNumber, LineNumber, Quantity)
                                                   Values (@OrderNumber, @PalletNumber, @LayerNumber, @LineNumber, @Quantity)"


                ' For each pallet on the order, add it to the table
                For Each orderPlt As PalletDetails In orderPalletDetails

                    Using cmd As SqlCommand = New SqlCommand(orderPalletLayerSQL, conn)
                        cmd.CommandTimeout = 0
                        cmd.Parameters.Add(New SqlParameter("@OrderNumber", orderNumber))
                        cmd.Parameters.Add(New SqlParameter("@PalletNumber", orderPlt.FinalPalletNumber))
                        cmd.Parameters.Add(New SqlParameter("@LayerNumber", orderPlt.LayerNumber))
                        cmd.Parameters.Add(New SqlParameter("@LineNumber", orderPlt.LineNumber))
                        cmd.Parameters.Add(New SqlParameter("@Quantity", orderPlt.Quantity))
                        cmd.ExecuteNonQuery()
                    End Using
                Next
            End Using
        Next

        SetLastBarCodeSequence(shopfloorConnection, customerID, lastSequence)
    End Sub

    Private Shared Function GetLastBarCodeSequence(ByVal shopfloorConnection As String, ByVal customerID As String) As Integer
        Dim result As Integer = 0
        Dim sql As String = "select LastBarCodeSequence
                           from EDICustomerReference
                           where CustomerID = @CustomerID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                result = Convert.ToInt32(command.ExecuteScalar())
            End Using
        End Using

        Return result
    End Function

    Private Shared Sub SetLastBarCodeSequence(ByVal shopfloorConnection As String, ByVal customerID As String, ByVal sequence As Integer)
        Dim sql As String = "update EDICustomerReference
                           set LastBarCodeSequence = @Sequence
                           where CustomerID = @CustomerID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@CustomerID", customerID))
                command.Parameters.Add(New SqlParameter("@Sequence", sequence))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Shared Function CalculateBarCode(ByRef lastSequence As Integer, ByVal palletNumber As String) As String
        lastSequence += 1
        Dim BarNumber As String = palletNumber.PadLeft(3, "0"c) & lastSequence.ToString().PadLeft(9, "0"c)
        Dim digit As String = CheckDigit(BarNumber)
        Dim result As String = BarNumber & digit
        Return result
    End Function

    Private Shared Function CheckDigit(ByVal data As String) As String
        Dim sum As Integer = 0
        Dim counter As Integer = 0

        For x As Integer = 2 To data.Length - 1
            counter += 1

            If counter Mod 2 = 0 Then
                sum += Convert.ToInt32(data.Substring(x, 1))
            Else
                sum += Convert.ToInt32(data.Substring(x, 1)) * 3
            End If
        Next

        Dim ichkDigit As Integer = 0
        Dim mod10 As Integer = sum Mod 10

        If mod10 <> 0 Then
            ichkDigit = 10 - mod10
        End If

        Return ichkDigit.ToString()
    End Function
End Class

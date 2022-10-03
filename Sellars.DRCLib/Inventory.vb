Imports System.Collections.Generic
Imports System.Collections.Concurrent
Imports System.Configuration
Imports System.Data
Imports System.Data.Spatial
Imports System.Data.SqlClient
Imports System.Device
Imports System.IO
Imports System.Math
Imports System.Runtime.InteropServices
Imports System.ServiceModel.Activation
Imports System.ServiceModel.Web
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Xml.Serialization
Imports System.ComponentModel

Imports Microsoft.SqlServer.Types

Public Class Inventory

    Public ErrorMessage As String = ""

    Public Sub Inventory()
        'SqlServer.Types.Utilities.LoadNativeAssemblies(AppDomain.CurrentDomain.BaseDirectory)
    End Sub


    ''' <summary>
    ''' Checks the designated warehouse for the part specified to see how much inventory is available to be ordered.
    ''' </summary>
    ''' <param name="Warehouse">The warehouse to check for inventory.</param>
    ''' <param name="Item">The part number of the item to check.</param>
    ''' <returns>The quantity of inventory available to be ordered for the designated part at the designated warehouse.  If the item is backordered at the warehouse, zero will be returned.</returns>
    ''' <remarks></remarks>
    Public Function Available(ByVal Warehouse As String, ByVal Item As String) As Integer
        ' SStart a task to get how much of the item is currently on hand at the new warehouse
        Dim onHandTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                       Return GetOnHand(Warehouse, Item)
                                                                   End Function)

        ' Start a task to get how much of the item is currently on order at the new warehouse
        Dim onOrderTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                        Return GetOnOrder(Warehouse, Item)
                                                                    End Function)

        ' Start a task to get how much of the item is currently on order at the new warehouse
        Dim onCartTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                       Return GetOnCart(Warehouse, Item)
                                                                   End Function)

        ' Start a task to get how much of the item is currently on a transfer for the designated warehouse
        Dim onTransferTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                           Return GetOnTransfer(Warehouse, Item)
                                                                       End Function)

        ' Start a task to get how much of the item is currently on a pending transfer for the designated warehouse
        Dim onPendingTransferTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                                  Return GetOnPendingTransfer(Warehouse, Item)
                                                                              End Function)

        ' Wait for all the threaded tasks to complete
        Task.WaitAll(onHandTask, onOrderTask, onCartTask, onTransferTask, onPendingTransferTask)

        Dim Avail As Integer = onHandTask.Result - onOrderTask.Result - onCartTask.Result - onTransferTask.Result - onPendingTransferTask.Result

        If Avail > 0 Then
            Return Avail
        Else
            Return 0
        End If
    End Function

    ''' <summary>
    ''' Checks the designated warehouse for the part specified to see how much inventory is available to be ordered.
    ''' </summary>
    ''' <param name="Warehouse">The warehouse to check for inventory.</param>
    ''' <param name="Item">The part number of the item to check.</param>
    ''' <returns>The quantity of inventory available to be ordered for the designated part at the designated warehouse.  If the item is backordered at the warehouse, zero will be returned.</returns>
    ''' <remarks></remarks>
    Public Function Available(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        ' SStart a task to get how much of the item is currently on hand at the new warehouse
        Dim onHandTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                       Return GetOnHand(Warehouse, UseStage, Item)
                                                                   End Function)

        ' Start a task to get how much of the item is currently on order at the new warehouse
        Dim onOrderTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                        Return GetOnOrder(Warehouse, UseStage, Item)
                                                                    End Function)

        ' Start a task to get how much of the item is currently on order at the new warehouse
        Dim onCartTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                       Return GetOnCart(Warehouse, UseStage, Item)
                                                                   End Function)

        ' Start a task to get how much of the item is currently on a transfer for the designated warehouse
        Dim onTransferTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                           Return GetOnTransfer(Warehouse, UseStage, Item)
                                                                       End Function)

        ' Start a task to get how much of the item is currently on a pending transfer for the designated warehouse
        Dim onPendingTransferTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                                  Return GetOnPendingTransfer(Warehouse, UseStage, Item)
                                                                              End Function)

        ' Start a task to get what the conversion factor is for the product 
        Dim getConversionTask As Task(Of Decimal) = Task.Factory.StartNew(Function() As Decimal
                                                                              Return GetConversion(Item)
                                                                          End Function)

        ' Wait for all the threaded tasks to complete
        Task.WaitAll(onHandTask, onOrderTask, onCartTask, onTransferTask, onPendingTransferTask, getConversionTask)

        'Dim Avail As Integer = onHandTask.Result - onOrderTask.Result - onCartTask.Result - Round(onTransferTask.Result * getConversionTask.Result, 0) - Round(onPendingTransferTask.Result * getConversionTask.Result, 0)

        Dim Avail As Integer = onHandTask.Result - onOrderTask.Result - onCartTask.Result - onTransferTask.Result - onPendingTransferTask.Result

        If Avail > 0 Then
            Return Avail
        Else
            Return 0
        End If
    End Function

    ''' <summary>
    ''' Checks the designated warehouse, and all backup warehouses to see which ones have enough inventory for the
    ''' order and returns the first warehouse, in order of priority, that has enough inventory to handle the line item.
    ''' </summary>
    ''' <param name="Warehouse">The original warehouse for the order.</param>
    ''' <param name="PartNumber">The part number to be checked.</param>
    ''' <param name="OriginalQuantity">If this is an update to the line item, this will contain the original quantity that was set up for this line item before being modified, otherwise this will be zero.</param>
    ''' <param name="NewQuantity">The will always contain the new quantity needed for this line item.</param>
    ''' <returns>Returns WarehouseData element with the first warehouse that can fulfill the line item.</returns>
    ''' <remarks></remarks>
    Public Function Available(ByVal Warehouse As String, ByVal PartNumber As String, ByVal OriginalQuantity As Integer, ByVal NewQuantity As Integer) As WarehouseData
        ' Get a list of all the warehouses (given, primary backup, and secondary backup)
        Dim Warehouses As BlockingCollection(Of WarehouseData) = GetBackupWarehouses(Warehouse)

        ' Check all the applicable locations for available inventory
        Parallel.ForEach(Warehouses, Sub(whse As WarehouseData)
                                         whse.Available = Available(whse.Warehouse, PartNumber)
                                     End Sub)

        ' Find the first warehouse that has inventory to return, or else return the default warehouse with no available
        Dim rtnWarehouse As WarehouseData = Warehouses.FirstOrDefault(Function(x) x.Available + OriginalQuantity > NewQuantity)

        ' If there was no warehouse that matched, then return the default warehouse for this item
        If rtnWarehouse Is Nothing Then
            ' Set the return warehouse to the default warehouse
            rtnWarehouse = Warehouses(0)

            ' Set the backorder flag in the return data to show we were completely out of inventory
            rtnWarehouse.Backorder = True
        End If

        ' Return the warehouse that has enough inv
        Return rtnWarehouse
    End Function

    Private Function GetOnHand(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            Conn.Open()

            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 = @STK and charindex('TRN', STK_06) = 0 and charindex('REW', STK_06) = 0"
            Else
                command = "SELECT @Quantity = isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 like @STK and charindex('STG', STK_06) = 0 and charindex('TRN', STK_06) = 0 and charindex('REW', STK_06) = 0"
            End If

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

                If UseStage Then
                    cmd.Parameters.Add(New SqlParameter("@STK", "STG " + Warehouse))
                Else
                    cmd.Parameters.Add(New SqlParameter("@STK", "%" + Warehouse + "%"))
                End If

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value)
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetOnHand(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 like @STK and charindex('TRN', STK_06) = 0 and charindex('REW', STK_06) = 0 and charindex('STG', STK_06) = 0"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", "%" + STK + "%"))

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value)
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetOnOrder(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(Round(DUEQTY_28 * SLSCNV_29, 0)), 0) from SO_Detail join Part_Sales with (NOLOCK) on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and STYPE_28 = 'CU' and PRTNUM_28 = @PRTNUM and STK_28 like @STK"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", "%" + STK + "%"))

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetOnOrder(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            Conn.Open()

            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(Round(DUEQTY_28 * SLSCNV_29, 0)), 0) from SO_Detail join Part_Sales with (NOLOCK) on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and STYPE_28 = 'CU' and PRTNUM_28 = @PRTNUM and STK_28 = @STK"
            Else
                command = "SELECT @Quantity = isnull(sum(Round(DUEQTY_28 * SLSCNV_29, 0)), 0) from SO_Detail join Part_Sales with (NOLOCK) on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and STYPE_28 = 'CU' and PRTNUM_28 = @PRTNUM and STK_28 like @STK and charindex('STG', STK_28) = 0 and charindex('TRN', STK_28) = 0"
            End If

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

                If UseStage Then
                    cmd.Parameters.Add(New SqlParameter("@STK", "STG " + Warehouse))
                Else
                    cmd.Parameters.Add(New SqlParameter("@STK", "%" + Warehouse + "%"))
                End If

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetConversion(ByVal Item As String) As Decimal
        Dim rtnConversion As Decimal = 0.0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            Conn.Open()

            Using Transaction As SqlTransaction = Conn.BeginTransaction(IsolationLevel.ReadUncommitted)
                Dim command As String = "SELECT @Conversion = isnull(SLSCNV_29, 0) from Part_Sales with (NOLOCK) where PRTNUM_29 = @PRTNUM"

                Using cmd As New SqlCommand(command, Conn)
                    cmd.Transaction = Transaction
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

                    Dim parmConversion As New SqlParameter("@Conversion", SqlDbType.Float, -1)
                    parmConversion.Direction = ParameterDirection.Output
                    parmConversion.Value = Nothing
                    cmd.Parameters.Add(parmConversion)

                    cmd.ExecuteNonQuery()
                    rtnConversion = Convert.ToDecimal(parmConversion.Value.ToString())
                End Using
            End Using
        End Using

        Return rtnConversion
    End Function

    Private Function GetOnCart(ByVal Warehouse As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(QuantityPurchased + QuantityFree), 0) from ShoppingCartDetail sd join ShoppingCartMaster sm on sm.CartID = sd.CartID where PRTNUM = @PRTNUM and Warehouse = @Warehouse"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@Warehouse", Warehouse))

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetOnCart(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            ' If we are using stage, then the cart does not come into play, so simply return 0
            If UseStage Then
                Return rtnQty
            End If

            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(QuantityPurchased + QuantityFree), 0) from ShoppingCartDetail sd join ShoppingCartMaster sm on sm.CartID = sd.CartID where PRTNUM = @PRTNUM and Warehouse = @Warehouse"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@Warehouse", Warehouse))

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            End Using
        End Using

        Return rtnQty
    End Function


    ' This module returns the quantity of a product that is on a transfer from the time that the transfer is submitted to assign a carrier, until the time
    ' transfer is shipped - which is when it will be removed from inventory within Max.
    Private Function GetOnTransfer(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Dim command As String = "SELECT @Quantity = isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and charindex(@STK, FromSTK) > 0 and tm.Status >= 1 and  tm.Status < 5"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", STK))

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = parmQty.Value.ToString()
            End Using
        End Using

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a transfer from the time that the transfer is submitted to assign a carrier, until the time
    ' transfer is shipped - which is when it will be removed from inventory within Max.
    Private Function GetOnTransfer(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and FromSTK = @STK and tm.Status >= 1 and  tm.Status < 5"
            Else
                command = "SELECT @Quantity = isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and FromSTK like @STK and tm.Status >= 1 and  tm.Status < 5"
            End If

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

                If UseStage Then
                    cmd.Parameters.Add(New SqlParameter("@STK", "STG " + Warehouse))
                Else
                    cmd.Parameters.Add(New SqlParameter("@STK", "%" + Warehouse + "%"))
                End If

                Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
                parmQty.Direction = ParameterDirection.Output
                parmQty.Value = Nothing
                cmd.Parameters.Add(parmQty)

                cmd.ExecuteNonQuery()
                rtnQty = parmQty.Value.ToString()
            End Using
        End Using

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a pending transfer that is being built within the system
    Private Function GetOnPendingTransfer(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Dim strSQL As String = "SELECT isnull(sum(Quantity), 0) from PendingTransferDetail td join PendingTransferMaster tm on tm.UserID = td.UserID where PartNumber = @PRTNUM and charindex(@STK, FromSTK) > 0"

        ' Set up a new SQL connection
        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)

            ' Open the SQL connection
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", STK))

                ' Execute the query, and return the results
                rtnQty = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a pending transfer that is being built within the system
    Private Function GetOnPendingTransfer(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0
        Dim strSQl As String = ""

        If UseStage Then
            strSQl = "SELECT isnull(sum(Quantity), 0) from PendingTransferDetail td join PendingTransferMaster tm on tm.UserID = td.UserID where PartNumber = @PRTNUM and FromSTK = @STK"
        Else
            strSQl = "SELECT isnull(sum(Quantity), 0) from PendingTransferDetail td join PendingTransferMaster tm on tm.UserID = td.UserID where PartNumber = @PRTNUM and charindex(@STK, FromSTK) > 0"
        End If

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            Conn.Open()

            Using cmd As New SqlCommand(strSQl, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

                If UseStage Then
                    cmd.Parameters.Add(New SqlParameter("@STK", "STG " + Warehouse))
                Else
                    cmd.Parameters.Add(New SqlParameter("@STK", Warehouse))
                End If

                rtnQty = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetBackupWarehouses(ByVal Warehouse As String) As BlockingCollection(Of WarehouseData)
        Dim rtnWarehouses As New BlockingCollection(Of WarehouseData)

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            ' Add the calling warehouse to the return warehouses
            Dim whse As New WarehouseData()
            whse.Sequence = 1
            whse.Warehouse = Warehouse
            rtnWarehouses.Add(whse)

            Conn.Open()

            Dim command As String = "SELECT @Primary = isnull(PrimaryBackup, ''), @Secondary = isnull(SecondaryBackup, '') from Warehouses where STK = @Warehouse"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@Warehouse", Warehouse))

                Dim parmPrimary As New SqlParameter("@Primary", SqlDbType.NVarChar, 10)
                parmPrimary.Direction = ParameterDirection.Output
                parmPrimary.Value = Nothing
                cmd.Parameters.Add(parmPrimary)

                Dim parmSecondary As New SqlParameter("@Secondary", SqlDbType.NVarChar, 10)
                parmSecondary.Direction = ParameterDirection.Output
                parmSecondary.Value = Nothing
                cmd.Parameters.Add(parmSecondary)

                cmd.ExecuteNonQuery()

                ' Check if there is a primary backup warehouse to return
                If parmPrimary.Value <> "" Then
                    Dim whsePrimary As New WarehouseData
                    whsePrimary.Sequence = 2
                    whsePrimary.Warehouse = parmPrimary.Value
                    rtnWarehouses.Add(whsePrimary)
                End If

                ' Check if there is a secondary backup warehouse to return
                If parmSecondary.Value <> "" Then
                    Dim whseSecondary As New WarehouseData
                    whseSecondary.Sequence = 3
                    whseSecondary.Warehouse = parmSecondary.Value
                    rtnWarehouses.Add(whseSecondary)
                End If
            End Using
        End Using

        Return rtnWarehouses
    End Function

    Public Function GetWarehouse(ByVal CUSTID As String, ByVal PrimaryWarehouse As String, ByVal PRTNUM As String, ByVal Quantity As Integer) As GetWarehouseResponse
        'define a variable for the warehouse to return
        Dim rtnData As New GetWarehouseResponse()

        Try
            ' Define a variable to show if we found a warehouse that has enough inventory
            Dim found As Boolean = False

            ' Check if the customer is excluded from participating in Round Robin, and if they are
            ' then set the found to true to end the searches and assign the warehouse to be the primary warehouse
            If CheckCustomerExcludes(CUSTID) Then
                found = True
                rtnData.Warehouse = PrimaryWarehouse
                rtnData.Available = Available(PrimaryWarehouse, PRTNUM)
            End If

            If Not found Then
                rtnData.Warehouse = PrimaryWarehouse
            End If
        Catch Ex As Exception
            Dim err As New Sellars.DRCLib.ErrorLog()
            err.Write(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString, "GetWarehouse", "DRCLib", Ex)
        End Try

        ' Return the selected warehouse
        Return rtnData
    End Function

    Private Function CheckCustomerExcludes(ByVal CUSTID As String) As Boolean
        Dim rtnData As Boolean = False
        Dim strSQL As String = "SELECT isnull(CUSTID, '') from RoundRobinCustomerExcludes where CUSTID = @CUSTID"

        Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            Conn.Open()

            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text

                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CUSTID", CUSTID))

                Dim resultq As Object = cmd.ExecuteScalar()

                If resultq IsNot Nothing AndAlso resultq IsNot DBNull.Value Then
                    ' Get the customer id returned from the sql call
                    If resultq.ToString().Trim() <> "" Then
                        rtnData = True
                    End If
                End If
            End Using
        End Using
    End Function
End Class

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

        Dim Avail As Integer = onHandTask.Result - onOrderTask.Result - onCartTask.Result - Round(onTransferTask.Result * getConversionTask.Result, 0) - Round(onPendingTransferTask.Result * getConversionTask.Result, 0)

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
        Dim Warehouses As BlockingCollection(Of WarehouseData) = GetWarehouses(Warehouse)

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
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 = @STK"
            Else
                command = "SELECT @Quantity = isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 like @STK and charindex('STG', STK_06) = 0 and charindex('TRN', STK_06) = 0"
            End If

            Dim cmd As New SqlCommand(command, Conn)
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

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value)
            Catch ex As Exception
                rtnQty = 0
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

        Return rtnQty
    End Function

    Private Function GetOnHand(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 like @STK and charindex('TRN', STK_06) = 0"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
            cmd.Parameters.Add(New SqlParameter("@STK", "%" + STK + "%"))

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value)
            Catch ex As Exception
                rtnQty = 0
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

        Return rtnQty
    End Function

    Private Function GetOnOrder(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)

        ' Define a variable for the database name so the application automatically sets up the SQL string below with the correct table names for test versus production
        Dim dataBaseName As String = IIf(System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString.ToUpper.Contains("TEST"), "TestShopfloorControl", "ShopfloorControl")

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(Round(DUEQTY_28 * SLSCNV_29, 0)), 0) from SO_Detail join Part_Sales on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and STYPE_28 = 'CU' and PRTNUM_28 = @PRTNUM and STK_28 like @STK"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
            cmd.Parameters.Add(New SqlParameter("@STK", "%" + STK + "%"))

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            Catch ex As Exception
                rtnQty = 0
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch ex As Exception
            rtnQty = 0
            Dim msg As String = ex.Message

        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnQty
    End Function

    Private Function GetOnOrder(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)

        ' Define a variable for the database name so the application automatically sets up the SQL string below with the correct table names for test versus production
        Dim dataBaseName As String = IIf(System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString.ToUpper.Contains("TEST"), "TestShopfloorControl", "ShopfloorControl")

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(Round(DUEQTY_28 * SLSCNV_29, 0)), 0) from SO_Detail join Part_Sales on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and STYPE_28 = 'CU' and PRTNUM_28 = @PRTNUM and STK_28 = @STK"
            Else
                command = "SELECT @Quantity = isnull(sum(Round(DUEQTY_28 * SLSCNV_29, 0)), 0) from SO_Detail join Part_Sales on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and STYPE_28 = 'CU' and PRTNUM_28 = @PRTNUM and STK_28 like @STK and charindex('STG', STK_06) = 0 and charindex('TRN', STK_06) = 0"
            End If

            Dim cmd As New SqlCommand(command, Conn)
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

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            Catch ex As Exception
                rtnQty = 0
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch ex As Exception
            rtnQty = 0
            Dim msg As String = ex.Message

        Finally
            ' Close the SQL connection object
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnQty
    End Function

    Private Function GetConversion(ByVal Item As String) As Decimal
        Dim rtnConversion As Decimal = 0.0

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            Dim Transaction As SqlTransaction = Conn.BeginTransaction(IsolationLevel.ReadUncommitted)

            Dim command As String = "SELECT @Conversion = isnull(SLSCNV_29, 0) from Part_Sales where PRTNUM_29 = @PRTNUM"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.Transaction = Transaction
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

            Dim parmConversion As New SqlParameter("@Conversion", SqlDbType.Float, -1)
            parmConversion.Direction = ParameterDirection.Output
            parmConversion.Value = Nothing
            cmd.Parameters.Add(parmConversion)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnConversion = Convert.ToDecimal(parmConversion.Value.ToString())
            Catch ex As Exception
                rtnConversion = 0
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

        Return rtnConversion
    End Function

    Private Function GetOnCart(ByVal Warehouse As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(QuantityPurchased + QuantityFree), 0) from ShoppingCartDetail sd join ShoppingCartMaster sm on sm.CartID = sd.CartID where PRTNUM = @PRTNUM and Warehouse = @Warehouse"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
            cmd.Parameters.Add(New SqlParameter("@Warehouse", Warehouse))

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            Catch ex As Exception
                rtnQty = 0
            Finally
                ' Close the SQL connection object
                If Conn.State = ConnectionState.Open Then
                    Conn.Close()
                End If
                Conn.Dispose()
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        End Try

        Return rtnQty
    End Function

    Private Function GetOnCart(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)

        ' If we are using stage, then the cart does not come into play, so simply return 0
        If UseStage Then
            Return rtnQty
        End If

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT @Quantity = isnull(sum(QuantityPurchased + QuantityFree), 0) from ShoppingCartDetail sd join ShoppingCartMaster sm on sm.CartID = sd.CartID where PRTNUM = @PRTNUM and Warehouse = @Warehouse"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
            cmd.Parameters.Add(New SqlParameter("@Warehouse", Warehouse))

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = Convert.ToInt32(parmQty.Value.ToString())
            Catch ex As Exception
                rtnQty = 0
            Finally
                ' Close the SQL connection object
                If Conn.State = ConnectionState.Open Then
                    Conn.Close()
                End If
                Conn.Dispose()
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        End Try

        Return rtnQty
    End Function


    ' This module returns the quantity of a product that is on a transfer from the time that the transfer is submitted to assign a carrier, until the time
    ' transfer is shipped - which is when it will be removed from inventory within Max.
    Private Function GetOnTransfer(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Dim command As String = "SELECT @Quantity = isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and charindex(@STK, FromSTK) > 0 and tm.Status >= 1 and  tm.Status < 5"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
            cmd.Parameters.Add(New SqlParameter("@STK", STK))

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = parmQty.Value.ToString()
            Catch ex As Exception
                rtnQty = 0
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = Data.ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a transfer from the time that the transfer is submitted to assign a carrier, until the time
    ' transfer is shipped - which is when it will be removed from inventory within Max.
    Private Function GetOnTransfer(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and FromSTK = @STK and tm.Status >= 1 and  tm.Status < 5"
            Else
                command = "SELECT @Quantity = isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and FromSTK like @STK and tm.Status >= 1 and  tm.Status < 5"
            End If

            Dim cmd As New SqlCommand(command, Conn)
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

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = parmQty.Value.ToString()
            Catch ex As Exception
                rtnQty = 0
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = Data.ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a pending transfer that is being built within the system
    Private Function GetOnPendingTransfer(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Dim command As String = "SELECT @Quantity = isnull(sum(Quantity), 0) from PendingTransferDetail td join PendingTransferMaster tm on tm.UserID = td.UserID where PartNumber = @PRTNUM and charindex(@STK, FromSTK) > 0"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
            cmd.Parameters.Add(New SqlParameter("@STK", STK))

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = parmQty.Value.ToString()
            Catch ex As Exception
                rtnQty = 0
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = Data.ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a pending transfer that is being built within the system
    Private Function GetOnPendingTransfer(ByVal Warehouse As String, ByVal UseStage As Boolean, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        ' Set up a new SQL connection string
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Try
            ' Open the SQL connection
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Dim command As String = ""

            If UseStage Then
                command = "SELECT @Quantity = isnull(sum(Quantity), 0) from PendingTransferDetail td join PendingTransferMaster tm on tm.UserID = td.UserID where PartNumber = @PRTNUM and FromSTK = @STK"
            Else
                command = "SELECT @Quantity = isnull(sum(Quantity), 0) from PendingTransferDetail td join PendingTransferMaster tm on tm.UserID = td.UserID where PartNumber = @PRTNUM and charindex(@STK, FromSTK) > 0"
            End If

            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))

            If UseStage Then
                cmd.Parameters.Add(New SqlParameter("@STK", "STG " + Warehouse))
            Else
                cmd.Parameters.Add(New SqlParameter("@STK", Warehouse))
            End If

            Dim parmQty As New SqlParameter("@Quantity", SqlDbType.Float, -1)
            parmQty.Direction = ParameterDirection.Output
            parmQty.Value = Nothing
            cmd.Parameters.Add(parmQty)

            'Perform the update to the database
            Try
                cmd.ExecuteNonQuery()
                rtnQty = parmQty.Value.ToString()
            Catch ex As Exception
                rtnQty = 0
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        Finally
            ' Close the SQL connection object
            If Conn.State = Data.ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.Dispose()
        End Try

        Return rtnQty
    End Function

    Private Function GetWarehouses(ByVal Warehouse As String) As BlockingCollection(Of WarehouseData)
        Dim rtnWarehouses As New BlockingCollection(Of WarehouseData)
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)

        ' Add the calling warehouse to the return warehouses
        Dim whse As New WarehouseData()
        whse.Sequence = 1
        whse.Warehouse = Warehouse
        rtnWarehouses.Add(whse)

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT @Primary = isnull(PrimaryBackup, ''), @Secondary = isnull(SecondaryBackup, '') from UnisourceFacilities where STK = @Warehouse"
            Dim cmd As New SqlCommand(command, Conn)
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

            'Perform the update to the database
            Try
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
            Catch ex As Exception
                ' do nothing
            Finally
                ' Close the SQL connection object
                If Conn.State = ConnectionState.Open Then
                    Conn.Close()
                End If
                Conn.Dispose()
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        End Try

        Return rtnWarehouses
    End Function

    Private Function GetOverstockWarehouses(ByVal PrimaryWarehouse As String, ByVal PRTNUM As String) As List(Of WarehouseData)
        Dim rtnWarehouses As New List(Of WarehouseData)

        ' Get all the warehouses data
        Dim AllWarehouses As BlockingCollection(Of WarehouseData) = GetAllWarehouses()

        ' Set up the query to get all the warehouses that have overstock quantities of the identified product
        ' sorted by the closest to where the primary warehouse is to the farthest away
        Dim strSQL As String = "select Warehouse from PartPricingOverstockProducts where PRTNUM = @PRTNUM and Active = 1"

        Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
        Dim cmd As SqlCommand = Nothing
        Dim dr As SqlDataReader = Nothing

        Try
            ' Open the database connection
            conn.Open()

            ' Set up a new SQL Command
            cmd = New SqlCommand(strSQL, conn)
            cmd.CommandType = CommandType.Text
            cmd.CommandTimeout = 0

            ' Add a parameter to the command that has the invoice number to update
            cmd.Parameters.Add(New SqlParameter("@PRTNUM", PRTNUM))

            ' Execute the sql statement so it returns a datareader
            dr = cmd.ExecuteReader()

            ' Loop through all the items returned
            If dr.Read() Then
                ' Split all the warehouses in the overstock value since it can contain multiple warehouse
                Dim warehouses As List(Of String) = dr("Warehouse").ToString().Split(",").ToList()

                ' Loop through the warehouses to get all of them
                For Each warehouse As String In warehouses
                    rtnWarehouses.Add(AllWarehouses.Where(Function(whse As WarehouseData) whse.Warehouse = warehouse).FirstOrDefault())
                Next
            End If

        Catch ex As Exception
        Finally
            ' Close the database connection
            conn.Close()
        End Try

        If cmd IsNot Nothing Then
            cmd.Dispose()
        End If

        ' Get the geolocation coordinate of the primary warehouse
        Dim primaryWhse As WarehouseData = AllWarehouses.Where(Function(whse As WarehouseData) whse.Warehouse = PrimaryWarehouse).FirstOrDefault()

        ' sort the list by the geographic distance from the primary warehouse
        rtnWarehouses = rtnWarehouses.OrderBy(Function(whse As WarehouseData) whse.GeoCoord.GetDistanceTo(primaryWhse.GeoCoord)).ToList()

        ' return the list
        Return rtnWarehouses
    End Function

    Public Function GetWarehouse(ByVal CUSTID As String, ByVal PrimaryWarehouse As String, ByVal PRTNUM As String, ByVal Quantity As Integer) As GetWarehouseResponse
        'define a variable for the warehouse to return
        Dim rtnData As New GetWarehouseResponse()

        Try
            ' Define a variable to show if we found a warehouse that has enough inventory
            Dim found As Boolean = False

            ' Define a variable to hold all the overstockWarehouse for a product
            Dim overstockWarehouses As New List(Of WarehouseData)

            ' Check if the customer is excluded from participating in Round Robin, and if they are
            ' then set the found to true to end the searches and assign the warehouse to be the primary warehouse
            If CheckCustomerExcludes(CUSTID) Then
                found = True
                rtnData.Warehouse = PrimaryWarehouse
                rtnData.Available = Available(PrimaryWarehouse, PRTNUM)
            End If

            ' If we have not found a warehouse yet, then we don't have an override excluding a customer from
            ' using RoundRobin
            If Not found Then
                ' If the item is overstock, we need to know which warehouses are all overstocked
                overstockWarehouses = GetOverstockWarehouses(PrimaryWarehouse, PRTNUM)

                ' Find the first warehouse that has enough inventory to fill the line
                For Each whse As WarehouseData In overstockWarehouses

                    ' if the available quantity at that warehouse is greater than the quantity needed for the order, then
                    ' exit the loop
                    Dim avail As Integer = Available(whse.Warehouse, PRTNUM)
                    If avail >= Quantity Then
                        found = True
                        rtnData.Warehouse = whse.Warehouse
                        rtnData.Available = avail
                        rtnData.Overstock = True
                        If whse.Warehouse <> PrimaryWarehouse Then
                            rtnData.RoundRobined = True
                        End If
                        Exit For
                    End If
                Next
            End If

            ' If we didn't find a warehouse yet, we either didn't have any overstocks or none of the overstock
            ' warehouses had enough to fill the line order
            If Not found Then
                ' First we need to get all the warehouses associated with the primary warehouse
                Dim warehouses As List(Of String) = (From whse As WarehouseData In GetWarehouses(PrimaryWarehouse)
                                                     Order By whse.Sequence
                                                     Select whse.Warehouse).ToList()

                ' Find the first warehouse that has enough inventory to fill the line
                ' since we checked the overstock warehouses in the search above, make sure that we exclude them from the 
                ' search here
                For Each whse As String In warehouses.Except((From whs As WarehouseData In overstockWarehouses Select whs.Warehouse).ToList())

                    ' if the available quantity at that warehouse is greater than the quantity needed for the order, then
                    ' exit the loop
                    Dim avail As Integer = Available(whse, PRTNUM)
                    If avail >= Quantity Then
                        found = True
                        rtnData.Warehouse = whse
                        rtnData.Available = avail
                        If whse <> PrimaryWarehouse Then
                            rtnData.RoundRobined = True
                        End If
                        Exit For
                    End If
                Next
            End If

            ' if none of the warehouses had enough inventory, and the customer was not excluded from Round Robin 
            ' then set the warehouse to be NEWB
            If Not found Then
                rtnData.Warehouse = "NEWB"
            End If
        Catch Ex As Exception
            ErrorMessage = Ex.Message
        End Try

        ' Return the selected warehouse
        Return rtnData
    End Function

    Private Function CheckCustomerExcludes(ByVal CUSTID As String) As Boolean
        Dim rtnData As Boolean = False

        Try
            Using Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)

                ' Open the SQL connection
                Conn.Open()

                Dim command As String = "SELECT @OUTCUSTID = isnull(CUSTID, '') from RoundRobinCustomerExcludes where CUSTID = @CUSTID"

                Using cmd As New SqlCommand(command, Conn)
                    cmd.CommandType = CommandType.Text

                    cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CUSTID", CUSTID))

                    Dim prmOut As New SqlParameter("@OUTCUSTID", SqlDbType.NVarChar, 25)
                    prmOut.Direction = ParameterDirection.Output
                    cmd.Parameters.Add(prmOut)

                    ' Run the stored procedure
                    cmd.ExecuteNonQuery()

                    ' Get the customer id returned from the sql call
                    If prmOut.Value.ToString.Trim() <> "" Then
                        rtnData = True
                    End If
                End Using
            End Using
        Catch ex As Exception
            Dim msg As String = ex.Message
            rtnData = False
        End Try

    End Function

    Private Function GetAllWarehouses() As BlockingCollection(Of WarehouseData)
        Dim rtnWarehouses As New BlockingCollection(Of WarehouseData)
        Dim Conn As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)

        Try
            ' Open the SQL connection
            Conn.Open()

            Dim command As String = "SELECT STK, GeoLoc from UnisourceFacilities"
            Dim cmd As New SqlCommand(command, Conn)
            cmd.CommandType = CommandType.Text

            'Perform the update to the database
            Try
                Dim dr As SqlDataReader = cmd.ExecuteReader()

                While dr.Read()
                    Dim whseSecondary As New WarehouseData
                    whseSecondary.Warehouse = dr("STK")
                    whseSecondary.GeoLoc = DirectCast(dr("GeoLoc"), SqlGeography)
                    whseSecondary.GeoCoord = New Location.GeoCoordinate(whseSecondary.GeoLoc.Lat, whseSecondary.GeoLoc.Long)
                    rtnWarehouses.Add(whseSecondary)
                End While
            Catch ex As Exception
                ' do nothing
                Dim msg As String = ex.Message
            Finally
                ' Close the SQL connection object
                If Conn.State = ConnectionState.Open Then
                    Conn.Close()
                End If
                Conn.Dispose()
            End Try

            ' Free up memory from the command object
            cmd.Dispose()
        Catch
        End Try

        Return rtnWarehouses
    End Function

End Class

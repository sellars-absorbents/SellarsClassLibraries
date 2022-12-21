Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports System.Math

Public Class SalesOrderClass
    Inherits ClassBase

    Private _OrderNumber As String = ""
    Private _Status As String = ""
    Private _Name As String = ""
    Private _Address1 As String = ""
    Private _Address2 As String = ""
    Private _City As String = ""
    Private _State As String = ""
    Private _ZipCode As String = ""
    Private _CustPO As String = ""
    Private _Customer As String = ""

    Public Enum OrderStatus
        Closed = 3
        Open = 4
    End Enum

    Public ReadOnly Property Address1() As String
        Get
            Return _Address1.Trim
        End Get
    End Property

    Public ReadOnly Property Address2() As String
        Get
            Return _Address2.Trim
        End Get
    End Property

    Public ReadOnly Property City() As String
        Get
            Return _City.Trim
        End Get
    End Property

    Public ReadOnly Property CustPO() As String
        Get
            Return _CustPO.Trim
        End Get
    End Property

    Public ReadOnly Property Customer() As String
        Get
            Return _Customer.Trim
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _Name.Trim
        End Get
    End Property

    Public ReadOnly Property OrderNumber() As String
        Get
            Return _OrderNumber.Trim
        End Get
    End Property

    Public ReadOnly Property State() As String
        Get
            Return _State.Trim
        End Get
    End Property

    Public ReadOnly Property Status() As String
        Get
            Return _Status.Trim
        End Get
    End Property

    Public ReadOnly Property ZipCode() As String
        Get
            Return _ZipCode.Trim
        End Get
    End Property

    Private ReadOnly Property ShopfloorConnection As String
    Private ReadOnly Property MAXConnection As String

    Public Sub New(ByVal passOrder As String)
        _OrderNumber = passOrder
        Read(passOrder)
    End Sub

    Public Sub New(ByVal shopfloorConnectionString As String, ByVal maxConnectionString As String)
        Me.ShopfloorConnection = shopfloorConnectionString
        Me.MAXConnection = maxConnectionString
    End Sub

    Private Sub Read(ByVal passorder As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, passorder, ParameterDirection.Input)

        Using dr As SqlDataReader = oSQL.RunProcReader("ReadSalesOrderMaster")
            ' Assign the variables from the database to properties
            If dr.Read() Then
                _Status = dr("STATUS_27")
                _Address1 = dr("ADDR1_27")
                _Address2 = dr("ADDR2_27")
                _City = dr("CITY_27")
                _Name = dr("NAME_27")
                _State = dr("STATE_27")
                _ZipCode = dr("ZIPCD_27")
                _CustPO = dr("CUSTPO_27")
                _Customer = dr("CUSTID_27")
            Else
                _Status = ""
                _Address1 = ""
                _Address2 = ""
                _City = ""
                _Name = ""
                _State = ""
                _ZipCode = ""
                _CustPO = ""
                _Customer = ""
                Throw New RecordNotOnDatabaseException(passorder, "SalesOrderMaster")
            End If
        End Using
    End Sub

    Public Class AddSPSResult
        Public Property OrderNumber As String
        Public Property BackOrdered As Boolean
        Public Property ErrorCode As String
        Public Property ErrorMessage As String
        Public Property Success As Boolean
    End Class

    Public Function AddSPS(ByRef som As SalesOrderMasterClass, ByVal SO As SalesOrder, ByVal CarrierInfo As ShippingCarrierInfo, ByVal FixCrLf As Boolean) As AddSPSResult
        Dim result As AddSPSResult = New AddSPSResult()
        ' Instantiate the sales order class
        ' Check if the Customer PO already exists, and if it does, then exit the routine
        'Dim som As New SalesOrderMasterClassASP()
        If som.IsDuplicatePO(SO.CUSTID, SO.CUSTPO) Then
            Throw New Exception(SO.CUSTID + " Duplicate PO # " + SO.CUSTPO)
        End If

        Try
            ' Call the function to initialize the MAX dlls
            'som.InitializeMax(maxDLLConnectionString, company, licensePath, logPath, False)

            ' Copy all the incoming sales order fields to matching fields in the sales order master record
            SO.CopyTo(som)

            ' Add the sales order master record
            Dim processLogID As Integer
            Dim rtnOrderNumber As String = ""

            processLogID = ProcessLog.Start(ShopfloorConnection, 5, "Add SPS SO_Master record for CustID " + SO.CUSTID + " PO # " + SO.CUSTPO, "DRCLib.SalesOrderClass", DateTime.Now)

            rtnOrderNumber = som.Add(Sellars.DRCLib.ClassBase.DataSource.Max)

            ProcessLog.Complete(ShopfloorConnection, processLogID, DateTime.Now)

            If rtnOrderNumber.Trim <> "" Then
                ' Set the order number in the data to be returned to the calling application
                result.OrderNumber = rtnOrderNumber

                Dim DueDate As DateTime = New DateTime(2030, 12, 31, 0, 0, 0)

                ' Record that it is not a unisource item
                Dim UnisourceStatus As Integer = UnisourceStatusCodes.NotUnisourceOrder

                ' Loop through and add all the items from the order to the database
                For Each item As SalesOrderDetail In SO.LINEITEMS
                    ' Instantiate the sales order detail class
                    Dim sod As New SalesOrderDetailClass()

                    ' Copy all the item fields from the inbound sales order to matching fields in the sales order detail record
                    item.CopyTo(sod)

                    ' Assign the order number field
                    sod.ORDNUM = rtnOrderNumber

                    Try
                        processLogID = ProcessLog.Start(ShopfloorConnection, 5, "Add SPS SO_Detail record for CustID " + SO.CUSTID + " PO # " + SO.CUSTPO + "Line # " + sod.LINNUM, "DRCLib.SalesOrderClass", DateTime.Now)

                        ' Add the line item to the order
                        sod.Add(Sellars.DRCLib.ClassBase.DataSource.Max)

                        ProcessLog.Complete(ShopfloorConnection, processLogID, DateTime.Now)

                        ' If there are notes on the order, then add them to the extension table
                        If item.Notes IsNot Nothing Then
                            If item.Notes.PrintOnBoth.Trim() <> "" Or item.Notes.DoNotPrint.Trim <> "" Or item.Notes.PrintOnInvoiceOnly.Trim <> "" Or item.Notes.PrintOnOrderOnly.Trim <> "" Then
                                Dim soNotes As New SONotesClass()
                                soNotes.Update(rtnOrderNumber, sod.LINNUM, sod.DELNUM, item.Notes.PrintOnBoth, item.Notes.PrintOnInvoiceOnly, item.Notes.PrintOnOrderOnly, item.Notes.DoNotPrint)
                            End If
                        End If

                        ' If Pallets > 0 then add the sales order detail ext record
                        Dim sode As New SODetailExtClass()
                        If item.Pallets > 0 Then
                            sode.Add(rtnOrderNumber, sod.LINNUM, sod.DELNUM, 0, 0, 0, True, UnisourceStatus, item.BuyerPartNumber, item.PromotionCode, item.UnitPrice, item.DiscountPercent, item.ShipFromWarehouse, item.WarehouseChangeReason, False)
                        Else
                            sode.Add(rtnOrderNumber, sod.LINNUM, sod.DELNUM, 0, 0, 0, False, UnisourceStatus, item.BuyerPartNumber, item.PromotionCode, item.UnitPrice, item.DiscountPercent, item.ShipFromWarehouse, item.WarehouseChangeReason, False)
                        End If
                    Catch ex As Exception
                        ErrorLog.Write(ShopfloorConnection, "Line Item Loop: " & ex.Message, "AddSPS", "SalesOrderService", ex.StackTrace)
                    End Try
                Next

                ' Check if there are any backordered items on the order
                For Each itm As SalesOrderDetail In SO.LINEITEMS
                    If CheckSellarsBackorder(itm.STK, itm) Then
                        result.BackOrdered = True
                        Exit For
                    End If
                Next

                ' Set the ShipFromStockID
                Dim ShipFromStockID As String = SO.DefaultStockID

                ' Initialize a new instance of the Sales Order Master Ext class
                Dim soe As New SOMasterExtClass

                Dim ThirdPartyBillingID As Long = 0
                If CarrierInfo.CarrierThirdParty.Trim <> "" Then
                    ' Get the third party billing ID
                    ThirdPartyBillingID = GetThirdPartyBillingID(SO.CUSTID, CarrierInfo.CarrierThirdParty, CarrierInfo.CarrierName, SO.ADDR1, SO.ADDR2, SO.CITY, SO.STATE, SO.ZIPCD, SO.CNTRY)
                    If ThirdPartyBillingID = 0 Then
                        Dim Title As String = CarrierInfo.CarrierName + "-" + CarrierInfo.CarrierThirdParty
                        ThirdPartyBillingID = AddThirdPartyBilling(SO.CUSTID, Title, CarrierInfo.CarrierThirdParty, CarrierInfo.CarrierName, SO.ADDR1, SO.ADDR2, SO.CITY, SO.STATE, SO.ZIPCD, SO.CNTRY)
                    End If
                End If

                Dim SendAcknowledgement = True
                If SO.CUSTID.Trim.ToUpper = "NAVISTA" Then
                    SendAcknowledgement = False
                End If

                ' Add the sales order master extension data
                ' If the customer ID is either walmart or home depot, do not print out a pick ticket
                ' otherwise print out a pick ticket
                soe.Add(rtnOrderNumber, "EDI", CarrierInfo.CarrierID, CarrierInfo.CarrierMethod, CarrierInfo.CarrierThirdParty, CarrierInfo.CarrierName, CarrierInfo.ContactName, CarrierInfo.ContactPhone, CarrierInfo.ContactEmail, SO.DefaultStockID, "CU", 1, "", "", 0, 0, ShipFromStockID, ThirdPartyBillingID, SendAcknowledgement)

                ' Update the miscellaneous fields on the record, and pass in the address location and RDC Description information necessary for the automated invoicing
                soe.UpdateNotes(rtnOrderNumber, SO.Notes, SO.LocationCodeQualifier, SO.AddressLocationNumber, SO.RDCDescription, SO.CarrierAlphaCode, SO.CarrierRouting, SO.CarrierTransMethodCode)

                ' If the notes field is not blank, then add the notes to the order
                If SO.Notes.Trim() <> "" Then
                    ' If we should fix the crlf on the notes, then do it
                    If FixCrLf Then
                        FixNote(rtnOrderNumber, "*")
                    End If
                End If

                ' Set a flag to print out the pick ticket
                Dim PickPrinted As Boolean = False

                ' Make sure to set the finished and pick printed indicators to true for the order to make sure that the work ticket 
                ' is generated by the automation and submitted to the correct department
                If Not CustomerUsesDocManager(SO.CUSTID.Trim.ToUpper) Then

                    ' If we have any line items from the MIL6 or DSC1 warehouse then we need to set the work ticket to be printed
                    If (From line In SO.LINEITEMS Where line.ShipFromWarehouse = "MIL6" Or line.ShipFromWarehouse = "DSC1" Select line).Count() > 0 Then
                        PickPrinted = True
                    End If
                End If

                ' Update the PickPrinted and Finished flag to the applicable values
                soe.UpdateFinished(rtnOrderNumber, SO.Finished, PickPrinted)
            End If

            result.Success = True
        Catch ex As Exception
            result.ErrorCode = -1
            result.Success = False
            result.ErrorMessage = ex.Message
            ErrorLog.Write(ShopfloorConnection, "Outer Loop: " & ex.Message, "AddSPS", "SalesOrderService", ex.StackTrace)
        End Try

        Return result
    End Function

    Private Function CustomerUsesDocManager(ByVal CUSTID As String) As Boolean
        Dim rtnVal As Boolean = False

        ' Open the SQL connection
        Using Conn As SqlConnection = New SqlConnection(ShopfloorConnection)
            Conn.Open()

            ' Set the sql command to execute
            Dim command As String = "SELECT UsesRetailDocManager from EDICustomerReference where CustomerID = @CUSTID"

            Using cmd As New SqlCommand(command, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@CUSTID", CUSTID))
                Dim resultq As Object = cmd.ExecuteScalar()

                If resultq IsNot Nothing AndAlso IsDBNull(resultq) = False Then
                    rtnVal = Convert.ToBoolean(resultq)
                End If
            End Using
        End Using

        Return rtnVal
    End Function

    Public Sub AddFastenal(ByRef som As SalesOrderMasterClass, ByVal Username As String, ByVal Password As String, ByVal SO As SalesOrder, ByVal CarrierInfo As ShippingCarrierInfo, ByVal FixCrLf As Boolean,
                           ByVal maxDLLConnectionString As String, ByVal company As String, ByVal licensePath As String, ByVal logPath As String)
        ' Instantiate the sales order class
        ' Check if the Customer PO already exists, and if it does, then exit the routine
        'Dim som As New SalesOrderMasterClassASP()
        If som.IsDuplicatePO(SO.CUSTID, SO.CUSTPO) Then
            Throw New Exception(SO.CUSTID + " Duplicate PO # " + SO.CUSTPO)
        End If

        Try
            ' Call the function to initialize the MAX dlls
            'som.InitializeMax(maxDLLConnectionString, company, licensePath, logPath, False)

            ' Copy all the incoming sales order fields to matching fields in the sales order master record
            SO.CopyTo(som)

            ' Add the sales order master record
            Dim processLogID As Integer
            Dim rtnOrderNumber As String = ""

            processLogID = ProcessLog.Start(ShopfloorConnection, 5, "Add SPS SO_Master record for CustID " + SO.CUSTID + " PO # " + SO.CUSTPO, "DRCLib.SalesOrderClass", DateTime.Now)

            rtnOrderNumber = som.Add(Sellars.DRCLib.ClassBase.DataSource.Max)

            ProcessLog.Complete(ShopfloorConnection, processLogID, DateTime.Now)

            If rtnOrderNumber.Trim <> "" Then

                ' Set the order number in the data to be returned to the calling application
                'rtnData.OrderNumber = rtnOrderNumber

                ' Loop through and add all the items from the order to the database
                For Each item As SalesOrderDetail In SO.LINEITEMS
                    ' Instantiate the sales order detail class
                    Dim sod As New SalesOrderDetailClass()
                    Dim DueDate As DateTime = New DateTime(2030, 12, 31, 0, 0, 0)

                    ' Copy all the item fields from the inbound sales order to matching fields in the sales order detail record
                    item.CopyTo(sod)

                    ' Assign the order number field
                    sod.ORDNUM = rtnOrderNumber

                    Try
                        ' Check if the item being added is a unisource Item, and if it is, then set the flag in the Sales Order Detail extension table
                        ' Dim UnisourceItem As Boolean = IsUnisourceItem(sod.STK)
                        Dim UnisourceStatus As Integer = 0

                        ' Call the sub routine to check and set the unisource status, due date, and override the STK code within the sod data if necessary
                        CheckFastenal3PL(UnisourceStatus, sod)

                        processLogID = ProcessLog.Start(ShopfloorConnection, 5, "Add SPS SO_Detail record for CustID " + SO.CUSTID + " PO # " + SO.CUSTPO + "Line # " + sod.LINNUM, "DRCLib.SalesOrderClass", DateTime.Now)

                        ' Add the order detail record to MAX
                        sod.Add(Sellars.DRCLib.ClassBase.DataSource.Max)

                        ProcessLog.Complete(ShopfloorConnection, processLogID, DateTime.Now)

                        ' If there are notes on the order, then add them to the extension table
                        If item.Notes IsNot Nothing Then
                            If item.Notes.PrintOnBoth.Trim() <> "" Or item.Notes.DoNotPrint.Trim <> "" Or item.Notes.PrintOnInvoiceOnly.Trim <> "" Or item.Notes.PrintOnOrderOnly.Trim <> "" Then
                                Dim soNotes As New SONotesClass()
                                soNotes.Update(rtnOrderNumber, sod.LINNUM, sod.DELNUM, item.Notes.PrintOnBoth, item.Notes.PrintOnInvoiceOnly, item.Notes.PrintOnOrderOnly, item.Notes.DoNotPrint)
                            End If
                        End If

                        ' If Pallets > 0 then add the sales order detail ext record
                        Dim sode As New SODetailExtClass()
                        If item.Pallets > 0 Then
                            sode.Add(rtnOrderNumber, sod.LINNUM, sod.DELNUM, 0, 0, 0, True, UnisourceStatus, item.BuyerPartNumber, item.PromotionCode, item.UnitPrice, item.DiscountPercent, item.ShipFromWarehouse, item.WarehouseChangeReason, False)
                        Else
                            sode.Add(rtnOrderNumber, sod.LINNUM, sod.DELNUM, 0, 0, 0, False, UnisourceStatus, item.BuyerPartNumber, item.PromotionCode, item.UnitPrice, item.DiscountPercent, item.ShipFromWarehouse, item.WarehouseChangeReason, False)
                        End If
                    Catch ex As Exception
                        ErrorLog.Write(ShopfloorConnection, "Line Item Loop: " & ex.Message, "AddFastenal", "SalesOrderService", ex.StackTrace)
                    End Try
                Next

                ' Initialize a new instance of the Sales Order Master Ext class
                Dim soe As New SOMasterExtClass

                Dim ThirdPartyBillingID As Long = 0
                If CarrierInfo.CarrierThirdParty.Trim <> "" Then
                    ' Get the third party billing ID
                    ThirdPartyBillingID = GetThirdPartyBillingID(SO.CUSTID, CarrierInfo.CarrierThirdParty, CarrierInfo.CarrierName, SO.ADDR1, SO.ADDR2, SO.CITY, SO.STATE, SO.ZIPCD, SO.CNTRY)
                    If ThirdPartyBillingID = 0 Then
                        Dim Title As String = CarrierInfo.CarrierName + "-" + CarrierInfo.CarrierThirdParty
                        ThirdPartyBillingID = AddThirdPartyBilling(SO.CUSTID, Title, CarrierInfo.CarrierThirdParty, CarrierInfo.CarrierName, SO.ADDR1, SO.ADDR2, SO.CITY, SO.STATE, SO.ZIPCD, SO.CNTRY)
                    End If
                End If

                ' Set the ShipFromStockID
                Dim ShipFromStockID As String = SO.DefaultStockID

                ' Add the sales order master extension data
                soe.Add(rtnOrderNumber, "EDI", CarrierInfo.CarrierID, CarrierInfo.CarrierMethod, CarrierInfo.CarrierThirdParty, CarrierInfo.CarrierName, CarrierInfo.ContactName, CarrierInfo.ContactPhone, CarrierInfo.ContactEmail, SO.DefaultStockID, "CU", 1, "", "", 0, 0, ShipFromStockID, ThirdPartyBillingID, False)

                ' Update the notes field on the record, and pass in the address location and RDC Description too
                soe.UpdateNotes(rtnOrderNumber, SO.Notes, SO.LocationCodeQualifier, SO.AddressLocationNumber, SO.RDCDescription, SO.CarrierAlphaCode, SO.CarrierRouting, SO.CarrierTransMethodCode)

                ' If the notes field is not blank, then add the notes to the order
                If SO.Notes.Trim() <> "" Then
                    ' If we should fix the crlf on the notes, then do it
                    If FixCrLf Then
                        FixNote(rtnOrderNumber, "*")
                    End If
                End If

                Dim PickPrinted As Boolean = False

                ' If we have any line items from the MIL6 or DSC1 warehouse then we need to set the work ticket to be printed
                If SO.Finished Then
                    If (From line In SO.LINEITEMS Where line.ShipFromWarehouse = "MIL6" Or line.ShipFromWarehouse = "DSC1" Select line).Count() > 0 Then
                        PickPrinted = True
                    End If
                End If

                ' Make sure to set the finished and pick printed indicators to true for the order to make sure that the work ticket 
                ' is generated by the automation and submitted to the correct department
                soe.UpdateFinished(rtnOrderNumber, SO.Finished, PickPrinted)
            End If
        Catch ex As Exception
            ErrorLog.Write(ShopfloorConnection, "Outer Loop: " & ex.Message, "AddFastenal", "SalesOrderService", ex.StackTrace)
        End Try
    End Sub

    Private Sub CheckFastenal3PL(ByRef UnisourceStatus As Integer, ByRef sod As SalesOrderDetailClass)
        ' Check if the item being added is a unisource Item, and if it is, then set the flag in the Sales Order Detail extension table
        ' UnisourceItem = IsUnisourceItem(sod.STK)
        Dim stkElements As Array = sod.STK.Trim.Split(" ")
        Dim chkStkCode As String = stkElements(1)
        Dim uStk As String = CheckFastenal3PL(chkStkCode)

        ' If it is a non-sellars warehouse, the stk code will be filled in, then do 3PL processing, othwerwise do sellar specific procesing
        If (uStk.Trim <> "") Then
            sod.STK = stkElements(0) + " " + uStk

            Dim STK As String = sod.STK
            Dim PRTNUM As String = sod.PRTNUM

            ' Start a task to get how much of the item is currently on hand at the designated warehouse
            Dim onHandTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                           Return GetOnHand(STK, PRTNUM)
                                                                       End Function)

            ' Start a task to get how much of the item is currently on order at the designated warehouse
            Dim onOrderTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                            Return GetOnOrder(STK, PRTNUM)
                                                                        End Function)

            ' Start a task to get how much of the item is currently on the magento shopping cart for the designated warehouse
            Dim onCartTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                           Return GetOnCart(STK, PRTNUM)
                                                                       End Function)

            ' Start a task to get how much of the item is currently on the magento shopping cart for the designated warehouse
            Dim onTransferTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                               Return GetOnTransfer(STK, PRTNUM)
                                                                           End Function)

            ' Start a task to get how much of the item is currently on the magento shopping cart for the designated warehouse
            Dim getConversionTask As Task(Of Decimal) = Task.Factory.StartNew(Function() As Decimal
                                                                                  Return GetConversion(PRTNUM)
                                                                              End Function)

            ' Wait for all the threaded tasks to complete
            Task.WaitAll(onHandTask, onOrderTask, onCartTask, onTransferTask, getConversionTask)

            If (onHandTask.Result - onOrderTask.Result - onCartTask.Result - Round(onTransferTask.Result * getConversionTask.Result, 0) >= Round(sod.CURQTY * getConversionTask.Result, 0)) Then
                UnisourceStatus = UnisourceStatusCodes.SendToUnisource
            Else
                UnisourceStatus = UnisourceStatusCodes.NotEnoughInventory
            End If
        Else
            ' Record that it is not a unisource item
            UnisourceStatus = UnisourceStatusCodes.NotUnisourceOrder
        End If
    End Sub

    Private Function CheckFastenal3PL(ByVal STK As String) As String
        Dim rtnSTK As String = ""

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(ShopfloorConnection)
            ' Open the SQL connection
            Conn.Open()

            Dim Warehouse As String = STK.Substring(STK.IndexOf(" ") + 1, STK.Length - STK.IndexOf(" ") - 1).Trim()

            Dim strSQL As String = "SELECT isnull(ID, 0) from Warehouses where STK = @STK and IsSellars = 0"

            Dim ID As Int64 = 0
            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@STK", Warehouse))
                cmd.ExecuteNonQuery()
                ID = Convert.ToInt64(cmd.ExecuteScalar())
            End Using

            ' If this is a unisourceItem, then get the correct stock code for this customer
            If ID > 0 Then
                rtnSTK = Warehouse
            End If
        End Using

        Return rtnSTK
    End Function

    Public Function IsOverstock(ByVal PRTNUM As String) As Boolean
        Dim rtnData As Boolean = False

        Using conn As New SqlConnection(ShopfloorConnection)
            Try
                conn.Open()

                Dim sqlStr As String = "select @Count = count(*) from PartPricingOverstockProducts where PRTNUM = @PRTNUM and Active = 1"

                Using cmd As New SqlCommand(sqlStr, conn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 0

                    cmd.Parameters.Add(New SqlParameter("@PRTNUM", PRTNUM))

                    Dim parmCount As New SqlParameter("@Count", SqlDbType.Int, -1)
                    parmCount.Direction = ParameterDirection.Output
                    parmCount.Value = Nothing
                    cmd.Parameters.Add(parmCount)

                    ' Execute the sql command
                    cmd.ExecuteNonQuery()

                    If Convert.ToInt16(parmCount.Value.ToString) > 0 Then
                        rtnData = True
                    End If
                End Using

            Catch ex As Exception
                Dim errmsg As String = ex.Message
            End Try
        End Using

        Return rtnData
    End Function

    Private Function AddThirdPartyBilling(ByVal CUSTID As String, ByVal Title As String, ByVal CarrierThirdParty As String, ByVal ShipToName As String, ByVal Address1 As String, ByVal Address2 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String) As Long
        ' Get the matching third party ID for a customer that matches the order data

        Dim ThirdPartyBillingID As Long = 0
        Try
            ' Set up a new SQL connection string
            Using cn As New SqlConnection(ShopfloorConnection)
                ' Open the SQL connection
                cn.Open()

                Dim strSql As String = "insert into CustomerThirdPartyBilling (CUSTID, Title, Name, AccountNumber, Address1, Address2, City, State, Zip, Country) values (@CUSTID, @Title, @Name, @AccountNumber, @Address1, @Address2, @City, @State, @Zip, @Country);"
                Using cmd As New SqlCommand(strSql, cn)
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(New SqlParameter("@CUSTID", CUSTID))
                    cmd.Parameters.Add(New SqlParameter("@Title", Title))
                    cmd.Parameters.Add(New SqlParameter("@Name", ShipToName))
                    cmd.Parameters.Add(New SqlParameter("@AccountNumber", CarrierThirdParty))
                    cmd.Parameters.Add(New SqlParameter("@Address1", Address1))
                    cmd.Parameters.Add(New SqlParameter("@Address2", Address2))
                    cmd.Parameters.Add(New SqlParameter("@City", City))
                    cmd.Parameters.Add(New SqlParameter("@State", State))
                    cmd.Parameters.Add(New SqlParameter("@Zip", Zip))
                    cmd.Parameters.Add(New SqlParameter("@Country", Country))

                    ' Execute the command, returning the new ID value
                    ThirdPartyBillingID = cmd.ExecuteScalar()
                End Using
            End Using

            ThirdPartyBillingID = GetThirdPartyBillingID(CUSTID, CarrierThirdParty, ShipToName, Address1, Address2, City, State, Zip, Country)
        Catch ex As Exception
            ThirdPartyBillingID = 0
        End Try

        Return ThirdPartyBillingID
    End Function

    Private Function GetThirdPartyBillingID(ByVal CUSTID As String, ByVal CarrierThirdParty As String, ByVal ShipToName As String, ByVal Address1 As String, ByVal Address2 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String) As Long
        ' Get the matching third party ID for a customer that matches the order data

        Dim ThirdPartyBillingID As Long = 0
        Try
            ' Set up a new SQL connection string
            Using cn As New SqlConnection(ShopfloorConnection)
                ' Open the SQL connection
                cn.Open()

                Dim strSql As String = "select @ID = isnull(ID, 0) from CustomerThirdPartyBilling where CUSTID = @CUSTID and AccountNumber = @AccountNumber and Address1 = @Address1 and Address2 = @Address2 and City = @City and State = @State and Zip = @Zip"
                Using cmd As New SqlCommand(strSql, cn)
                    cmd.CommandType = CommandType.Text
                    cmd.Parameters.Add(New SqlParameter("@CUSTID", CUSTID))
                    cmd.Parameters.Add(New SqlParameter("@Address1", Address1))
                    cmd.Parameters.Add(New SqlParameter("@Address2", Address2))
                    cmd.Parameters.Add(New SqlParameter("@City", City))
                    cmd.Parameters.Add(New SqlParameter("@State", State))
                    cmd.Parameters.Add(New SqlParameter("@Zip", Zip))
                    cmd.Parameters.Add(New SqlParameter("@AccountNumber", CarrierThirdParty))

                    Dim parmID As New SqlParameter("@ID", SqlDbType.BigInt, -1)
                    parmID.Direction = ParameterDirection.Output
                    parmID.Value = Nothing
                    cmd.Parameters.Add(parmID)
                    cmd.ExecuteNonQuery()
                    ThirdPartyBillingID = Convert.ToInt64(Convert.ToDecimal(parmID.Value.ToString()))
                End Using
            End Using
        Catch ex As Exception
            ThirdPartyBillingID = 0
        End Try

        Return ThirdPartyBillingID
    End Function

    Private Function FixNote(ByVal ORDNUM As String, ByVal CheckString As String) As Boolean
        Dim isValid As Boolean = False

        Dim strSQL As String = "update SalesOrderMasterExt " &
                               "set Notes = cast(replace(cast(Notes as nvarchar(max)), @CheckString, CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) + @CheckString) as ntext) " &
                               "where ORDNUM = @ORDNUM"

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(ShopfloorConnection)
            ' Open the SQL connection
            Conn.Open()

            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@CheckString", CheckString))
                cmd.ExecuteNonQuery()
                isValid = True
            End Using
        End Using

        Return isValid
    End Function

    Private Function GetOnHand(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Dim strSQL As String = "SELECT isnull(sum(QTYOH_06), 0) from Part_Stock where PRTNUM_06 = @PRTNUM and STK_06 = @STK"

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(MAXConnection)
            ' Open the SQL connection
            Conn.Open()

            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", STK))
                rtnQty = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetOnOrder(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Dim strSQL As String = "SELECT isnull(sum(DUEQTY_28 * SLSCNV_29), 0) from SO_Detail join Part_Sales ps with (NOLOCK) on PRTNUM_29 = PRTNUM_28 where STATUS_28 = '3' and PRTNUM_28 = @PRTNUM and STK_28 = @STK and STYPE_28 = 'CU'"

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(MAXConnection)
            ' Open the SQL connection
            Conn.Open()

            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", STK))
                rtnQty = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnQty
    End Function

    ' This module returns the quantity of a product that is on a transfer from the time that the transfer is submitted to assign a carrier, until the time
    ' transfer is shipped - which is when it will be removed from inventory within Max.
    Private Function GetOnTransfer(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Dim strSQL As String = "SELECT isnull(sum(QuantityRequested), 0) from TransferDetail td join TransferMaster tm on tm.ID = td.TMID where PartNumber = @PRTNUM and charindex(FromSTK, @STK) > 0 and tm.Status >= 1 and  tm.Status < 5"

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(ShopfloorConnection)
            ' Open the SQL connection
            Conn.Open()

            ' We need to get all transfer orders that are in status two, waiting for the carrier to be assigned, sent to Unisource, and acknowledged.
            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", STK))
                rtnQty = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetOnCart(ByVal STK As String, ByVal Item As String) As Integer
        Dim rtnQty As Integer = 0

        Dim strSQL As String = "SELECT isnull(sum(QuantityPurchased + QuantityFree), 0) from ShoppingCartDetail scd join ShoppingCartMaster scm on scm.CartID = scd.CartID where PRTNUM = @PRTNUM and Warehouse = @STK"

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(ShopfloorConnection)
            ' Open the SQL connection
            Conn.Open()

            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                cmd.Parameters.Add(New SqlParameter("@STK", STK))
                rtnQty = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnQty
    End Function

    Private Function GetConversion(ByVal Item As String) As Decimal
        Dim rtnConversion As Decimal = 0.0

        Dim strSQL As String = "SELECT isnull(SLSCNV_29, 0) from Part_Sales with (NOLOCK) where PRTNUM_29 = @PRTNUM"

        ' Set up a new SQL connection string
        Using Conn As New SqlConnection(MAXConnection)
            ' Open the SQL connection
            Conn.Open()

            Using cmd As New SqlCommand(strSQL, Conn)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", Item))
                rtnConversion = Convert.ToDecimal(cmd.ExecuteScalar())
            End Using
        End Using

        Return rtnConversion
    End Function

    Private Function CheckSellarsBackorder(ByVal DefaultStockID As String, ByRef sod As SalesOrderDetail) As Boolean
        ' Check if the item being added is a unisource Item, and if it is, then set the flag in the Sales Order Detail extension table
        ' UnisourceItem = IsUnisourceItem(sod.STK)
        Dim stkElements As Array = DefaultStockID.Trim.Split(" ")
        Dim chkStkCode As String = stkElements(1)

        Dim StockID As String = DefaultStockID

        Dim STK As String = DefaultStockID
        Dim stkStage As String = stkElements(0) + " STG"
        Dim PRTNUM As String = sod.PRTNUM

        ' Start a task to get how much of the item is currently on hand at the designated warehouse
        Dim onHandTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                       Return GetOnHand(STK, PRTNUM)
                                                                   End Function)

        Dim onHandStgTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                          Return GetOnHand(stkStage, PRTNUM)
                                                                      End Function)

        ' Start a task to get how much of the item is currently on order at the designated warehouse
        Dim onOrderTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                        Return GetOnOrder(STK, PRTNUM)
                                                                    End Function)

        ' Start a task to get how much of the item is currently on the magento shopping cart for the designated warehouse
        Dim onCartTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                       Return GetOnCart(STK, PRTNUM)
                                                                   End Function)

        ' Start a task to get how much of the item is currently on the magento shopping cart for the designated warehouse
        Dim onTransferTask As Task(Of Integer) = Task.Factory.StartNew(Function() As Integer
                                                                           Return GetOnTransfer(STK, PRTNUM)
                                                                       End Function)

        ' Start a task to get how much of the item is currently on the magento shopping cart for the designated warehouse
        Dim getConversionTask As Task(Of Decimal) = Task.Factory.StartNew(Function() As Decimal
                                                                              Return GetConversion(PRTNUM)
                                                                          End Function)

        ' Wait for all the threaded tasks to complete
        Task.WaitAll(onHandTask, onHandStgTask, onOrderTask, onCartTask, onTransferTask, getConversionTask)

        Dim Status As Integer = 0
        If (onHandTask.Result + onHandStgTask.Result - onOrderTask.Result - onCartTask.Result - Round(onTransferTask.Result * getConversionTask.Result, 0) >= Round(sod.CURQTY * getConversionTask.Result, 0)) Then
            Status = False
        Else
            Status = True
        End If

        Return Status
    End Function

    Private Enum UnisourceStatusCodes
        NotUnisourceOrder = 0
        NotEnoughInventory = 1
        SendToUnisource = 2
        SuccessfullySent = 3
        WaitingForAcknowledgement = 4
        Shipped = 5
        OrderShipped = 6
    End Enum
End Class

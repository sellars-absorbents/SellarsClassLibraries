Imports System.Collections.Generic
Imports System.Data.SqlClient

Public Class QueuedWebSalesOrderMaster
    Protected Friend Const DefaultDate As DateTime = #1/1/2000 12:01:00 AM#
    Private _id As Guid
    Private _ORDNUM As String = ""                    ' Order Number
    Private _CartID As Long = 0                       ' Shopping Cart ID
    Private _CUSTID As String = ""                    ' Customer ID
    Private _CUSTPO As String = ""                    ' Customer Purchase Order Number
    Private _ORDDTE As DateTime = DefaultDate         ' Date the order was created
    Private _COMNT1 As String = ""                    ' First Comment Line
    Private _COMNT2 As String = ""                    ' Second Comment Line
    Private _COMNT3 As String = ""                    ' Third Comment Line
    Private _NAME As String = ""                      ' Customer Name
    Private _ADDR1 As String = ""                     ' Customer Address Line 1
    Private _ADDR2 As String = ""                     ' Customer Address Line 2
    Private _ADDR3 As String = ""                     ' Customer Address Line 3
    Private _ADDR4 As String = ""                     ' Customer Address Line 4
    Private _ADDR5 As String = ""                     ' Customer Address Line 5
    Private _ADDR6 As String = ""                     ' Customer Address Line 6
    Private _CITY As String = ""                      ' Customer City
    Private _STATE As String = ""                     ' Customer State
    Private _ZIPCD As String = ""                     ' Customer Zip Code
    Private _CNTRY As String = ""                     ' Customer Country
    Private _ContactPhone As String = ""              ' Customer Phone
    Private _ContactName As String = ""               ' Customer Contact's Name
    Private _ContactEmail As String = ""              ' Customer Email Address
    Private _Notes As String = ""                     ' Extra Notes field stored in the SOE table
    Private _CCTransactionID As String = ""           ' Credit card transaction id field, stored in the some table
    Private _CCAuthCode As String = ""
    Private _Carrier As Integer                   ' Carrier that the order will use for shipping (eg. FedEx, UPS, Customer Pickup)
    Private _CarrierMethod As String = ""             ' Method that the carrier will use for shipping (.eg Ground, Next Day Air, etc)
    Private _CarrierThirdParty As String = ""         ' Third party account number for the shipping service, if applicable
    Private _CarrierName As String = ""               ' If Carrier is Customer Pickup, this shows the name of the carrier company, if applicable
    Private _FOB As String = ""                       ' FOB
    Private _ShippingContactName As String = ""               ' Shipping contact name within the company who will be receiving the shipment
    Private _ShippingContactPhone As String = ""              ' Shipping contact phone number within the company of the person who receives the shipment
    Private _ShippingContactEmail As String = ""              ' Shipping contact email within the company of the person who receives the shipment
    Private _EstimatedShipping As String = ""
    Private _WebsiteTotal As String = ""
    Private _CCType As String = ""
    Private _CCNUmber As String = ""
    Private _CCExpiration As Date = DefaultDate
    Private _CCVN As String = ""
    Private _CCBTFirstName As String = ""
    Private _CCBTLastName As String = ""
    Private _CCBTAddress As String = ""
    Private _CCBTCity As String = ""
    Private _CCBTState As String = ""
    Private _CCBTZip As String = ""
    Private _ShippingNotes As String = ""
    Private _salesOrderDetails As List(Of QueuedWebSalesOrderDetail)

    Public Property FixCrLF As Boolean

    Public Property Carrier As Integer
        Get
            Return _Carrier
        End Get
        Set(value As Integer)
            _Carrier = value
        End Set
    End Property

    Public Property CarrierMethod As String
        Get
            Return _CarrierMethod
        End Get
        Set(value As String)
            _CarrierMethod = value
        End Set
    End Property

    Public Property CarrierName As String
        Get
            Return _CarrierName
        End Get
        Set(value As String)
            _CarrierName = value
        End Set
    End Property

    Public Property CarrierThirdParty As String
        Get
            Return _CarrierThirdParty
        End Get
        Set(value As String)
            _CarrierThirdParty = value
        End Set
    End Property

    Public Property CartID As Long
        Get
            Return _CartID
        End Get
        Set(value As Long)
            _CartID = value
        End Set
    End Property

    Public Property CCAuthCode As String
        Get
            Return _CCAuthCode
        End Get
        Set(value As String)
            _CCAuthCode = value
        End Set
    End Property

    Public Property CCTransactionID As String
        Get
            Return _CCTransactionID
        End Get
        Set(value As String)
            _CCTransactionID = value
        End Set
    End Property

    Public Property ContactEmail As String
        Get
            Return _ContactEmail
        End Get
        Set(value As String)
            _ContactEmail = value
        End Set
    End Property

    Public Property ContactName As String
        Get
            Return _ContactName
        End Get
        Set(value As String)
            _ContactName = value
        End Set
    End Property

    Public Property ContactPhone As String
        Get
            Return _ContactPhone
        End Get
        Set(value As String)
            _ContactPhone = value
        End Set
    End Property

    Public Property ORDNUM() As String
        Get
            Return GetValue(_ORDNUM, "")
        End Get
        Set(ByVal value As String)
            _ORDNUM = GetValue(value, "")
        End Set
    End Property

    Public Property CUSTID() As String
        Get
            Return GetValue(_CUSTID, "")
        End Get
        Set(ByVal value As String)
            _CUSTID = GetValue(value, "")
        End Set
    End Property

    Public Property FOB() As String
        Get
            Return GetValue(_FOB, "")
        End Get
        Set(ByVal value As String)
            _FOB = GetValue(value, "")
        End Set
    End Property

    Public Property DefaultSTK As String

    Public Property CUSTPO() As String
        Get
            Return GetValue(_CUSTPO, "")
        End Get
        Set(ByVal value As String)
            _CUSTPO = GetValue(value, "")
        End Set
    End Property

    Public Property ORDDTE() As DateTime
        Get
            Return GetValue(_ORDDTE, DefaultDate)
        End Get
        Set(ByVal value As DateTime)
            _ORDDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property COMNT1() As String
        Get
            Return GetValue(_COMNT1, "")
        End Get
        Set(ByVal value As String)
            _COMNT1 = GetValue(value, "")
        End Set
    End Property

    Public Property COMNT2() As String
        Get
            Return GetValue(_COMNT2, "")
        End Get
        Set(ByVal value As String)
            _COMNT2 = GetValue(value, "")
        End Set
    End Property

    Public Property COMNT3() As String
        Get
            Return GetValue(_COMNT3, "")
        End Get
        Set(ByVal value As String)
            _COMNT3 = GetValue(value, "")
        End Set
    End Property
    Public Property NAME() As String
        Get
            Return GetValue(_NAME, "")
        End Get
        Set(ByVal value As String)
            _NAME = GetValue(value, "")
        End Set
    End Property

    Public Property ADDR1() As String
        Get
            Return GetValue(_ADDR1, "")
        End Get
        Set(ByVal value As String)
            _ADDR1 = GetValue(value, "")
        End Set
    End Property

    Public Property ADDR2() As String
        Get
            Return GetValue(_ADDR2, "")
        End Get
        Set(ByVal value As String)
            _ADDR2 = GetValue(value, "")
        End Set
    End Property

    Public Property ADDR3() As String
        Get
            Return GetValue(_ADDR3, "")
        End Get
        Set(ByVal value As String)
            _ADDR3 = GetValue(value, "")
        End Set
    End Property

    Public Property ADDR4() As String
        Get
            Return GetValue(_ADDR4, "")
        End Get
        Set(ByVal value As String)
            _ADDR4 = GetValue(value, "")
        End Set
    End Property

    Public Property ADDR5() As String
        Get
            Return GetValue(_ADDR5, "")
        End Get
        Set(ByVal value As String)
            _ADDR5 = GetValue(value, "")
        End Set
    End Property

    Public Property ADDR6() As String
        Get
            Return GetValue(_ADDR6, "")
        End Get
        Set(ByVal value As String)
            _ADDR6 = GetValue(value, "")
        End Set
    End Property

    Public Property CITY() As String
        Get
            Return GetValue(_CITY, "")
        End Get
        Set(ByVal value As String)
            _CITY = GetValue(value, "")
        End Set
    End Property

    Public Property STATE() As String
        Get
            Return GetValue(_STATE, "")
        End Get
        Set(ByVal value As String)
            _STATE = GetValue(value, "")
        End Set
    End Property

    Public Property ZIPCD() As String
        Get
            Return GetValue(_ZIPCD, "")
        End Get
        Set(ByVal value As String)
            _ZIPCD = GetValue(value, "")
        End Set
    End Property

    Public Property CNTRY() As String
        Get
            Return GetValue(_CNTRY, "")
        End Get
        Set(ByVal value As String)
            _CNTRY = GetValue(value, "")
        End Set
    End Property

    Public Property ShippingContactEmail As String
        Get
            Return _ShippingContactEmail
        End Get
        Set(value As String)
            _ShippingContactEmail = value
        End Set
    End Property

    Public Property ShippingContactName As String
        Get
            Return _ShippingContactName
        End Get
        Set(value As String)
            _ShippingContactName = value
        End Set
    End Property

    Public Property ShippingContactPhone As String
        Get
            Return _ShippingContactPhone
        End Get
        Set(value As String)
            _ShippingContactPhone = value
        End Set
    End Property

    Public Property EstimatedShipping As String
        Get
            Return _EstimatedShipping
        End Get
        Set(value As String)
            _EstimatedShipping = value
        End Set
    End Property

    Public Property WebsiteTotal As String
        Get
            Return _WebsiteTotal
        End Get
        Set(value As String)
            _WebsiteTotal = value
        End Set
    End Property

    Public Property Notes() As String
        Get
            Return GetValue(_Notes, "")
        End Get
        Set(ByVal value As String)
            _Notes = GetValue(value, "")
        End Set
    End Property

    Public Property CCExpiration() As Date
        Get
            Return GetValue(_CCExpiration, DefaultDate)
        End Get
        Set(ByVal value As Date)
            _CCExpiration = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property CCNumber() As String
        Get
            Return GetValue(_CCNUmber, "")
        End Get
        Set(ByVal value As String)
            _CCNUmber = GetValue(value, "")
        End Set
    End Property

    Public Property CCType() As String
        Get
            Return GetValue(_CCType, "")
        End Get
        Set(ByVal value As String)
            _CCType = GetValue(value, "")
        End Set
    End Property

    Public Property CCVN() As String
        Get
            Return GetValue(_CCVN, "")
        End Get
        Set(ByVal value As String)
            _CCVN = GetValue(value, "")
        End Set
    End Property

    Public Property CCBTAddress() As String
        Get
            Return GetValue(_CCBTAddress, "")
        End Get
        Set(ByVal value As String)
            _CCBTAddress = GetValue(value, "")
        End Set
    End Property

    Public Property CCBTCity() As String
        Get
            Return GetValue(_CCBTCity, "")
        End Get
        Set(ByVal value As String)
            _CCBTCity = GetValue(value, "")
        End Set
    End Property

    Public Property CCBTFirstName() As String
        Get
            Return GetValue(_CCBTFirstName, "")
        End Get
        Set(ByVal value As String)
            _CCBTFirstName = GetValue(value, "")
        End Set
    End Property

    Public Property CCBTLastName() As String
        Get
            Return GetValue(_CCBTLastName, "")
        End Get
        Set(ByVal value As String)
            _CCBTLastName = GetValue(value, "")
        End Set
    End Property

    Public Property CCBTState() As String
        Get
            Return GetValue(_CCBTState, "")
        End Get
        Set(ByVal value As String)
            _CCBTState = GetValue(value, "")
        End Set
    End Property

    Public Property CCBTZip() As String
        Get
            Return GetValue(_CCBTZip, "")
        End Get
        Set(ByVal value As String)
            _CCBTZip = GetValue(value, "")
        End Set
    End Property

    Public Property ShippingNotes() As String
        Get
            Return GetValue(_ShippingNotes, "")
        End Get
        Set(ByVal value As String)
            _ShippingNotes = GetValue(value, "")
        End Set
    End Property

    Public Property SalesOrderDetails() As List(Of QueuedWebSalesOrderDetail)
        Get
            Return _salesOrderDetails
        End Get
        Private Set(value As List(Of QueuedWebSalesOrderDetail))
            _salesOrderDetails = value
        End Set
    End Property

    Private _onHold As Boolean = False
    Public Property OnHold As Boolean
        Get
            Return _onHold
        End Get
        Private Set(value As Boolean)
            _onHold = value
        End Set
    End Property

    Private _processTries As Integer = 0
    Public Property ProcessTries As Integer
        Get
            Return _processTries
        End Get
        Private Set(value As Integer)
            _processTries = value
        End Set
    End Property

    Private _processed As Boolean = False
    Public Property Processed As Boolean
        Get
            Return _processed
        End Get
        Private Set(value As Boolean)
            _processed = value
        End Set
    End Property

    Private _processedOn As Date = DefaultDate
    Public Property ProcessedOn As Date
        Get
            Return _processedOn
        End Get
        Private Set(value As Date)
            _processedOn = value
        End Set
    End Property

    Private _createdOn As Date = DefaultDate
    Public Property CreatedOn As Date
        Get
            Return _createdOn
        End Get
        Private Set(value As Date)
            _createdOn = value
        End Set
    End Property

    Private Function GetValue(Of T)(ByVal value As T, ByVal defaultValue As T) As T
        If value Is Nothing Then
            Return defaultValue
        Else
            ' If a string was pased in, then make sure that we return the trimmed value
            If TypeOf (value) Is String Then
                Return Convert.ChangeType(value.ToString().Trim(), GetType(T))
            Else
                Return value
            End If
        End If
    End Function

    Public Property ID As Guid
        Get
            Return _id
        End Get
        Private Set(value As Guid)
            _id = value
        End Set
    End Property

    Public Function AddRecordToQueue(ByVal shopfloorConnection As String) As Guid
        Dim sql As String = "INSERT INTO [dbo].[QueuedWebSalesOrderMaster]([ORDNUM],[CartID],[CUSTID],[CUSTPO],[ORDDTE],[COMNT1],[COMNT2],[COMNT3],[NAME],[ADDR1],[ADDR2],[ADDR3],[ADDR4],[ADDR5],[ADDR6],[CITY],[STATE],[ZIPCD],[CNTRY],[ContactPhone],[ContactName],[ContactEmail],[Notes],[CCTransactionID],[CCAuthCode],[Carrier],[CarrierMethod],[CarrierThirdParty],[CarrierName],[FOB], [DefaultSTK],[ShippingContactName],[ShippingContactPhone],[ShippingContactEmail],[EstimatedShipping],[WebsiteTotal],[CCType],[CCNumber],[CCExpiration],[CCVN],[CCBTFirstName],[CCBTLastName],[CCBTAddress],[CCBTCity],[CCBTState],[CCBTZip],[ShippingNotes], [FixCrLF])
                             OUTPUT INSERTED.ID
                             VALUES(@ORDNUM,
                                        @CartID,
                                        @CUSTID,
                                        @CUSTPO,
                                        @ORDDTE,
                                        @COMNT1,
                                        @COMNT2,
                                        @COMNT3,
                                        @NAME,
                                        @ADDR1,
                                        @ADDR2,
                                        @ADDR3,
                                        @ADDR4,
                                        @ADDR5,
                                        @ADDR6,
                                        @CITY,
                                        @STATE,
                                        @ZIPCD,
                                        @CNTRY,
                                        @ContactPhone,
                                        @ContactName,
                                        @ContactEmail,
                                        @Notes,
                                        @CCTransactionID,
                                        @CCAuthCode,
                                        @Carrier,
                                        @CarrierMethod,
                                        @CarrierThirdParty,
                                        @CarrierName,
                                        @FOB,
                                        @DefaultSTK,
                                        @ShippingContactName,
                                        @ShippingContactPhone,
                                        @ShippingContactEmail,
                                        @EstimatedShipping,
                                        @WebsiteTotal,
                                        @CCType,
                                        @CCNumber,
                                        @CCExpiration,
                                        @CCVN,
                                        @CCBTFirstName,
                                        @CCBTLastName,
                                        @CCBTAddress,
                                        @CCBTCity,
                                        @CCBTState,
                                        @CCBTZip,
                                        @ShippingNotes,
                                        @FixCrLF);"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ORDNUM", Me.ORDNUM))
                command.Parameters.Add(New SqlParameter("@CartID", Me.CartID))
                command.Parameters.Add(New SqlParameter("@CUSTID", Me.CUSTID))
                command.Parameters.Add(New SqlParameter("@CUSTPO", Me.CUSTPO))
                command.Parameters.Add(New SqlParameter("@ORDDTE", Me.ORDDTE))
                command.Parameters.Add(New SqlParameter("@COMNT1", Me.COMNT1))
                command.Parameters.Add(New SqlParameter("@COMNT2", Me.COMNT2))
                command.Parameters.Add(New SqlParameter("@COMNT3", Me.COMNT3))
                command.Parameters.Add(New SqlParameter("@NAME", Me.NAME))
                command.Parameters.Add(New SqlParameter("@ADDR1", Me.ADDR1))
                command.Parameters.Add(New SqlParameter("@ADDR2", Me.ADDR2))
                command.Parameters.Add(New SqlParameter("@ADDR3", Me.ADDR3))
                command.Parameters.Add(New SqlParameter("@ADDR4", Me.ADDR4))
                command.Parameters.Add(New SqlParameter("@ADDR5", Me.ADDR5))
                command.Parameters.Add(New SqlParameter("@ADDR6", Me.ADDR6))
                command.Parameters.Add(New SqlParameter("@CITY", Me.CITY))
                command.Parameters.Add(New SqlParameter("@STATE", Me.STATE))
                command.Parameters.Add(New SqlParameter("@ZIPCD", Me.ZIPCD))
                command.Parameters.Add(New SqlParameter("@CNTRY", Me.CNTRY))
                command.Parameters.Add(New SqlParameter("@ContactPhone", Me.ContactPhone))
                command.Parameters.Add(New SqlParameter("@ContactName", Me.ContactName))
                command.Parameters.Add(New SqlParameter("@ContactEmail", Me.ContactEmail))
                command.Parameters.Add(New SqlParameter("@Notes", Me.Notes))
                command.Parameters.Add(New SqlParameter("@CCTransactionID", Me.CCTransactionID))
                command.Parameters.Add(New SqlParameter("@CCAuthCode", Me.CCAuthCode))
                command.Parameters.Add(New SqlParameter("@Carrier", Me.Carrier))
                command.Parameters.Add(New SqlParameter("@CarrierMethod", Me.CarrierMethod))
                command.Parameters.Add(New SqlParameter("@CarrierThirdParty", Me.CarrierThirdParty))
                command.Parameters.Add(New SqlParameter("@CarrierName", Me.CarrierName))
                command.Parameters.Add(New SqlParameter("@FOB", Me.FOB))
                command.Parameters.Add(New SqlParameter("@DefaultSTK", Me.DefaultSTK))
                command.Parameters.Add(New SqlParameter("@ShippingContactName", Me.ShippingContactName))
                command.Parameters.Add(New SqlParameter("@ShippingContactPhone", Me.ShippingContactPhone))
                command.Parameters.Add(New SqlParameter("@ShippingContactEmail", Me.ShippingContactEmail))
                command.Parameters.Add(New SqlParameter("@EstimatedShipping", Me.EstimatedShipping))
                command.Parameters.Add(New SqlParameter("@WebsiteTotal", Me.WebsiteTotal))
                command.Parameters.Add(New SqlParameter("@CCType", Me.CCType))
                command.Parameters.Add(New SqlParameter("@CCNumber", Me.CCNumber))
                command.Parameters.Add(New SqlParameter("@CCExpiration", Me.CCExpiration))
                command.Parameters.Add(New SqlParameter("@CCVN", Me.CCVN))
                command.Parameters.Add(New SqlParameter("@CCBTFirstName", Me.CCBTFirstName))
                command.Parameters.Add(New SqlParameter("@CCBTLastName", Me.CCBTLastName))
                command.Parameters.Add(New SqlParameter("@CCBTAddress", Me.CCBTAddress))
                command.Parameters.Add(New SqlParameter("@CCBTCity", Me.CCBTCity))
                command.Parameters.Add(New SqlParameter("@CCBTState", Me.CCBTState))
                command.Parameters.Add(New SqlParameter("@CCBTZip", Me.CCBTZip))
                command.Parameters.Add(New SqlParameter("@ShippingNotes", Me.ShippingNotes))
                command.Parameters.Add(New SqlParameter("@FixCrLF", Me.FixCrLF))

                Me.ID = CType(command.ExecuteScalar(), Guid)
            End Using
        End Using

        Return Me.ID
    End Function

    Public Shared Function GetQueuedRecords(ByVal shopfloorConnection As String, ByVal onHold As Boolean, ByVal processed As Boolean) As List(Of QueuedWebSalesOrderMaster)
        Dim results As New List(Of QueuedWebSalesOrderMaster)
        Dim sql As String = "select *
                               from QueuedWebSalesOrderMaster
                               where OnHold = @OnHold
                               and Processed = @Processed"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@OnHold", onHold))
                command.Parameters.Add(New SqlParameter("@Processed", processed))

                Using dr As SqlDataReader = command.ExecuteReader()
                    While dr.Read()
                        Dim master As QueuedWebSalesOrderMaster = New QueuedWebSalesOrderMaster()

                        PopulateProperties(dr, master)

                        results.Add(master)
                    End While
                End Using
            End Using

            For Each master As QueuedWebSalesOrderMaster In results
                master.SalesOrderDetails = QueuedWebSalesOrderDetail.GetQueuedRecords(connection, master.ID, onHold, processed)
            Next
        End Using

        Return results
    End Function

    Private Shared Sub PopulateProperties(dr As SqlDataReader, master As QueuedWebSalesOrderMaster)
        master.ID = dr("ID")
        master.ORDNUM = dr("ORDNUM")
        master.CartID = dr("CartID")
        master.CUSTID = dr("CUSTID")
        master.CUSTPO = dr("CUSTPO")
        master.ORDDTE = dr("ORDDTE")
        master.COMNT1 = dr("COMNT1")
        master.COMNT2 = dr("COMNT2")
        master.COMNT3 = dr("COMNT3")
        master.NAME = dr("NAME")
        master.ADDR1 = dr("ADDR1")
        master.ADDR2 = dr("ADDR2")
        master.ADDR3 = dr("ADDR3")
        master.ADDR4 = dr("ADDR4")
        master.ADDR5 = dr("ADDR5")
        master.ADDR6 = dr("ADDR6")
        master.CITY = dr("CITY")
        master.STATE = dr("STATE")
        master.ZIPCD = dr("ZIPCD")
        master.CNTRY = dr("CNTRY")
        master.ContactPhone = dr("ContactPhone")
        master.ContactName = dr("ContactName")
        master.ContactEmail = dr("ContactEmail")
        master.Notes = dr("Notes")
        master.CCTransactionID = dr("CCTransactionID")
        master.CCAuthCode = dr("CCAuthCode")
        master.Carrier = dr("Carrier")
        master.CarrierMethod = dr("CarrierMethod")
        master.CarrierThirdParty = dr("CarrierThirdParty")
        master.CarrierName = dr("CarrierName")
        master.FOB = dr("FOB")
        master.DefaultSTK = dr("DefaultSTK")
        master.ShippingContactName = dr("ShippingContactName")
        master.ShippingContactPhone = dr("ShippingContactPhone")
        master.ShippingContactEmail = dr("ShippingContactEmail")
        master.EstimatedShipping = dr("EstimatedShipping")
        master.WebsiteTotal = dr("WebsiteTotal")
        master.CCType = dr("CCType")
        master.CCNumber = dr("CCNumber")
        master.CCExpiration = dr("CCExpiration")
        master.CCVN = dr("CCVN")
        master.CCBTFirstName = dr("CCBTFirstName")
        master.CCBTLastName = dr("CCBTLastName")
        master.CCBTAddress = dr("CCBTAddress")
        master.CCBTCity = dr("CCBTCity")
        master.CCBTState = dr("CCBTState")
        master.CCBTZip = dr("CCBTZip")
        master.ShippingNotes = dr("ShippingNotes")
        master.FixCrLF = dr("FixCrLF")
        master.OnHold = dr("OnHold")
        master.ProcessTries = dr("ProcessTries")
        master.Processed = dr("Processed")

        If (dr("ProcessedOn") IsNot DBNull.Value) Then
            master.ProcessedOn = dr("ProcessedOn")
        Else
            master.ProcessedOn = DefaultDate
        End If

        master.CreatedOn = dr("CreatedOn")
    End Sub

    Public Sub IncrementProcessTries(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderMaster 
                             set ProcessTries = ProcessTries + 1
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub SetOrderNumber(ByVal shopfloorConnection As String, ByVal orderNumber As String)
        Dim sql As String = "update QueuedWebSalesOrderMaster 
                             set ORDNUM = @ORDNUM
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ORDNUM", orderNumber))
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub MarkAsProcessed(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderMaster 
                             set Processed = 1,
                             ProcessedOn = GETDATE()
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class

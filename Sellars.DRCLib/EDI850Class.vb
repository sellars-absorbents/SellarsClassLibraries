Imports Sellars.SQL

Public Class EDI850Class
    Inherits ClassBase

    Private _CustomerPO As String
    Private _CustomerId As String
    Private _BillToAddress As String
    Private _ShipToCode As String
    Private _ShipToName As String
    Private _ShipToAddress1 As String
    Private _ShipToAddress2 As String
    Private _ShipToCity As String
    Private _ShipToState As String
    Private _ShipToZip As String
    Private _ShipToCountry As String
    Private _OrderedBy As String
    Private _FOB As String
    Private _TermsCode As String
    Private _TermsValue As String
    Private _ShipViaCode As String
    Private _Taxable As String
    Private _TaxCode1 As String
    Private _TaxCode2 As String
    Private _TaxCode3 As String
    Private _TaxRate1 As Decimal
    Private _TaxRate2 As Decimal
    Private _TaxRate3 As Decimal
    Private _OrderLines As DataSet
    Private _TotalPrice As Decimal
    Private _ExtNotes As String
    Private _Comment1 As String
    Private _Comment2 As String
    Private _Comment3 As String
    Private _TradingPartnerId As String
    Private _Source As String
    Private _LocationCodeQualifier As String
    Private _AddressLocationNumber As String

    Public ReadOnly Property TradingPartnerId() As String
        Get
            Return _TradingPartnerId.Trim
        End Get
    End Property

    Public ReadOnly Property CusterPO() As String
        Get
            Return _CustomerPO.Trim
        End Get
    End Property

    Public ReadOnly Property CustomerId() As String
        Get
            Return _CustomerId.Trim
        End Get
    End Property

    Public ReadOnly Property BillToAddress() As String
        Get
            Return _BillToAddress.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToCode() As String
        Get
            Return _ShipToCode.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToName() As String
        Get
            Return _ShipToName.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToAddress1() As String
        Get
            Return _ShipToAddress1.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToAddress2() As String
        Get
            Return _ShipToAddress2.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToCity() As String
        Get
            Return _ShipToCity.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToState() As String
        Get
            Return _ShipToState.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToZip() As String
        Get
            Return _ShipToZip.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToCountry() As String
        Get
            Return _ShipToCountry.Trim
        End Get
    End Property

    Public ReadOnly Property OrderedBy() As String
        Get
            Return _OrderedBy.Trim
        End Get
    End Property

    Public ReadOnly Property FOB() As String
        Get
            Return _FOB.Trim
        End Get
    End Property

    Public ReadOnly Property TermsCode() As String
        Get
            Return _TermsCode.Trim
        End Get
    End Property

    Public ReadOnly Property TermsValue() As String
        Get
            Return _TermsValue.Trim
        End Get
    End Property

    Public ReadOnly Property ShipViaCode() As String
        Get
            Return _ShipViaCode.Trim
        End Get
    End Property

    Public ReadOnly Property Taxable() As String
        Get
            Return _Taxable.Trim
        End Get
    End Property

    Public ReadOnly Property TaxCode1() As String
        Get
            Return _TaxCode1.Trim
        End Get
    End Property

    Public ReadOnly Property TaxCode2() As String
        Get
            Return _TaxCode2.Trim
        End Get
    End Property

    Public ReadOnly Property TaxCode3() As String
        Get
            Return _TaxCode3.Trim
        End Get
    End Property

    Public ReadOnly Property TaxRate1() As Decimal
        Get
            Return _TaxRate1
        End Get
    End Property

    Public ReadOnly Property TaxRate2() As Decimal
        Get
            Return _TaxRate2
        End Get
    End Property

    Public ReadOnly Property TaxRate3() As Decimal
        Get
            Return _TaxRate3
        End Get
    End Property

    Public ReadOnly Property OrderLines() As DataSet
        Get
            Return _OrderLines
        End Get
    End Property

    Public ReadOnly Property TotalPrice() As Decimal
        Get
            Return _TotalPrice
        End Get
    End Property

    Public ReadOnly Property ExtNotes() As String
        Get
            Return _ExtNotes
        End Get
    End Property

    Public ReadOnly Property Comment1() As String
        Get
            Return _Comment1.Trim
        End Get
    End Property

    Public ReadOnly Property Comment2() As String
        Get
            Return _Comment2.Trim
        End Get
    End Property

    Public ReadOnly Property Comment3() As String
        Get
            Return _Comment3.Trim
        End Get
    End Property

    Public ReadOnly Property Source() As String
        Get
            Return _Source.Trim
        End Get
    End Property

    Public ReadOnly Property AddressLocationNumber() As String
        Get
            Return _AddressLocationNumber.Trim
        End Get
    End Property

    Public ReadOnly Property LocationCodeQualifier() As String
        Get
            Return _LocationCodeQualifier.Trim
        End Get
    End Property


    Public Function GetNewOrders() As DataSet
        Dim oSQL As New SqlService(ConnectionString)
        Dim myDataSet As New DataSet
        Dim ediDS As New DataSet
        Dim webDS As New DataSet

        Dim strSQL As String = "SELECT EDI_850_HEADER.CUST_PO, EDI_850_HEADER.PO_DATE, EDI_CUST_XREF.TRADING_PARTNER, " & _
                               "EDI_850_HEADER.TRADING_PARTNER_ID, 'EDI' AS 'SOURCE' " & _
                               "FROM EDI_850_HEADER INNER JOIN " & _
                               "EDI_CUST_XREF ON EDI_850_HEADER.TRADING_PARTNER_ID = EDI_CUST_XREF.TRADING_PARTNER_ID " & _
                               "WHERE (EDI_850_HEADER.PROCESSED = 0) " & _
                               "ORDER BY EDI_CUST_XREF.TRADING_PARTNER"


        ediDS = oSQL.RunSql(strSQL, "NewOrders")

        strSQL = "SELECT WebOrder_HEADER.MAGENTO_ORD_NUM AS 'CUST_PO', WebOrder_HEADER.PO_DATE, " & _
                 "(BILL_FIRST_NAME + ' ' + BILL_LAST_NAME) AS 'TRADING_PARTNER', " & _
                 "WebOrder_HEADER.MAGENTO_CUST_ID AS 'TRADING_PARTNER_ID', 'WEB' AS 'SOURCE' " & _
                 "FROM WebOrder_HEADER " & _
                 "INNER JOIN WebOrder_CUST_XREF on WebOrder_CUST_XREF.MAGENTO_CUST_ID = WebOrder_HEADER.MAGENTO_CUST_ID " & _
                 "WHERE WebOrder_HEADER.PROCESSED = 0"

        webDS = oSQL.RunSql(strSQL, "NewOrders")

        ediDS.Tables("NewOrders").Merge(webDS.Tables("NewOrders"))
        myDataSet = ediDS

        Return myDataSet

    End Function

    Public Sub GetOrder(ByVal customer As String, ByVal cust_po As String, ByVal source As String)

        Dim custID As String
        Dim _custPO As Object
        _Source = source

        ' Get the data
        If source = "EDI" Then 'Order is from SPS EDI
            Dim oSQL As New SqlService(ConnectionString)
            Dim strSQL As String = "SELECT * " & _
                                   "FROM EDI_CUST_XREF " & _
                                   "WHERE (EDI_CUST_XREF.TRADING_PARTNER_ID = '" & customer & "')"
            Dim dsCustXref As New DataSet
            dsCustXref = oSQL.RunSql(strSQL, "TradingPartnerID")

            Dim tradingPartner_id As String = dsCustXref.Tables("TradingPartnerID").Rows(0).Item("TRADING_PARTNER_ID")
            _TradingPartnerId = tradingPartner_id

            custID = dsCustXref.Tables("TradingPartnerID").Rows(0).Item("CUST_ID")

            strSQL = "SELECT * FROM EDI_850_HEADER WHERE " & _
                     "TRADING_PARTNER_ID = '" & tradingPartner_id & "' AND " & _
                     "CUST_PO = '" & cust_po & "'"

            Dim dsOrderHeader As New DataSet
            dsOrderHeader = oSQL.RunSql(strSQL, "OrderHeader")

            strSQL = "SELECT * FROM EDI_850_LINE WHERE " & _
                     "TRADING_PARTNER_ID = '" & tradingPartner_id & "' AND " & _
                     "CUST_PO = '" & cust_po & "'"

            Dim dsOrderLines As New DataSet
            dsOrderLines = oSQL.RunSql(strSQL, "OrderLines")

            'Create appropiate class based on customer
            If customer = "588ALLSELLARSAB" Then 'Lowes
                _custPO = New EDILowes850Class(dsOrderHeader, dsOrderLines, dsCustXref)
            ElseIf customer = "009ALLSELLARSDI" Then 'Target
                _custPO = New EDITarget850Class(dsOrderHeader, dsOrderLines, dsCustXref)
            ElseIf customer = "524ALLSELLARSDI" Then 'AutoZone
                _custPO = New EDIAutoZone850Class(dsOrderHeader, dsOrderLines, dsCustXref)
            End If

            ' *** The following fields are ONLY available on EDI transactions, so 
            ' *** This code is placed within the IF to better address that

            ' Location Code Qualifier
            _LocationCodeQualifier = _custPO.LocationCodeQualifier

            ' Address Location Number
            _AddressLocationNumber = _custPO.AddressLocationNumber

            oSQL = Nothing
        Else 'Order is from Magento Web Site 
            Dim oSQL As New SqlService(ConnectionString)
            Dim strSQL As String = "SELECT * " & _
                                   "FROM WebOrder_CUST_XREF " & _
                                   "WHERE (WebOrder_CUST_XREF.MAGENTO_CUST_ID = '" & customer & "')"
            Dim dsCustXref As New DataSet
            dsCustXref = oSQL.RunSql(strSQL, "TradingPartnerID")

            Dim tradingPartner_id As String = dsCustXref.Tables("TradingPartnerID").Rows(0).Item("MAGENTO_CUST_ID")
            _TradingPartnerId = tradingPartner_id

            custID = dsCustXref.Tables("TradingPartnerID").Rows(0).Item("CUST_ID")

            strSQL = "SELECT * FROM WebOrder_HEADER WHERE " & _
                     "MAGENTO_CUST_ID = '" & tradingPartner_id & "' AND " & _
                     "MAGENTO_ORD_NUM = '" & cust_po & "'"

            Dim dsOrderHeader As New DataSet
            dsOrderHeader = oSQL.RunSql(strSQL, "OrderHeader")

            strSQL = "SELECT * FROM WebOrder_LINE WHERE " & _
                     "MAGENTO_CUST_ID = '" & tradingPartner_id & "' AND " & _
                     "MAGENTO_ORD_NUM = '" & cust_po & "'"

            Dim dsOrderLines As New DataSet
            dsOrderLines = oSQL.RunSql(strSQL, "OrderLines")

            _custPO = New EDIMagentoClass(dsOrderHeader, dsOrderLines, dsCustXref)
        End If

        'Set the properties
        'Customer Id
        _CustomerId = custID

        'Bill To Address
        _BillToAddress = _custPO.Customer.MailingAddress

        'Ship To Address
        _ShipToCode = _custPO.ShipToCode
        _ShipToName = _custPO.ShipToName
        _ShipToAddress1 = _custPO.ShipToAddress1
        _ShipToAddress2 = _custPO.ShipToAddress2
        _ShipToCity = _custPO.ShipToCity
        _ShipToState = _custPO.ShipToState
        _ShipToZip = _custPO.ShipToZip
        _ShipToCountry = _custPO.ShipToCountry

        'Tax Codes
        _Taxable = _custPO.Customer.Taxable
        _TaxCode1 = _custPO.TaxCode1
        _TaxCode2 = _custPO.TaxCode2
        _TaxCode3 = _custPO.TaxCode3
        _TaxRate1 = _custPO.TaxRate1
        _TaxRate2 = _custPO.TaxRate2
        _TaxRate3 = _custPO.TaxRate3

        'PO
        '_CustomerPO = cust_po
        _CustomerPO = _custPO.PONumber

        'Terms
        _TermsCode = _custPO.TermsCode

        'Ship Via
        _ShipViaCode = _custPO.ShipViaCode

        'Ordered By
        _OrderedBy = _custPO.OrderedBy

        'FOB
        _FOB = _custPO.FOB

        'Lines
        _OrderLines = _custPO.OrderLines

        'Total Price
        _TotalPrice = _custPO.TotalPrice

        'Comment1
        _Comment1 = _custPO.Comment1

        'Comment2
        _Comment2 = _custPO.Comment2

        'Comment3
        _Comment3 = _custPO.Comment3

        'ExtNotes
        _ExtNotes = _custPO.ExtNotes

    End Sub

    Public Sub UpdateEDIProcessed(ByVal tradingPartnerId As String, ByVal custPO As String, ByVal source As String)
        Dim oSQL As New SqlService(ConnectionString)

        If source = "EDI" Then
            ' Add the parameters to the command object
            oSQL.AddParameter("@Trading_Partner_Id", SqlDbType.NVarChar, 50, tradingPartnerId, ParameterDirection.Input)
            oSQL.AddParameter("@Cust_Po", SqlDbType.NVarChar, 50, custPO, ParameterDirection.Input)

            ' Run the stored procedure
            oSQL.RunProc("UpdateEDI850Processed")
        Else ' It's a Magento order
            ' Add the parameters to the command object
            oSQL.AddParameter("@Trading_Partner_Id", SqlDbType.NVarChar, 50, tradingPartnerId, ParameterDirection.Input)
            oSQL.AddParameter("@Cust_Po", SqlDbType.NVarChar, 50, custPO, ParameterDirection.Input)

            ' Run the stored procedure
            oSQL.RunProc("UpdateWebOrderProcessed")
        End If

    End Sub
End Class

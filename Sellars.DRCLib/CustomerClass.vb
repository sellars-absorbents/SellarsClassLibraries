Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports System.Math

Public Class CustomerClass
    Inherits ClassBase

    Public Addresses As ShippingMasterClass
    Public PriceQuote As PriceQuoteClass
    Public Terms As CodeMasterClass

    Private _CreditLimit As Decimal
    Private _Contact As String
    Private _Comment1 As String
    Private _Comment2 As String
    Private _CustomerNumber As String
    Private _CusTyp As String
    Private _Name As String
    Private _Phone As String
    Private _SalesRep As String
    Private _SalesRepName As String
    Private _ShipCode As String
    Private _Status As String
    Private _Taxable As String
    Private _TaxCode1 As String
    Private _TaxCode2 As String
    Private _TaxCode3 As String
    Private _Terms As String
    Private _TermCode As String
    Private _ShipVia As String
    Private _FOB As String

    Private _Addr1 As String
    Private _Addr2 As String
    Private _Addr3 As String
    Private _City As String
    Private _Country As String
    Private _MailingAddress As String
    Private _State As String
    Private _ZipCode As String

    'Credit Limit information from Dynamics
    Private _AmountOverLimit As Decimal = 0
    Private _CustomerBalance As Decimal = 0
    Private _DepositsReceived As Decimal = 0
    Private _OnOrderAmount As Decimal = 0
    Private _OrderedNotInvoicedAmount As Decimal = 0
    Private _OrderedNotInvoicedCredit As Decimal = 0
    Private _UnpostedCashAmount As Decimal = 0
    Private _UnpostedOtherCashAmount As Decimal = 0
    Private _UnpostedOtherSalesAmount As Decimal = 0
    Private _UnpostedSalesAmount As Decimal = 0

    Public ReadOnly Property AmountOverLimit() As Decimal
        Get
            Return _AmountOverLimit
        End Get
    End Property

    Public ReadOnly Property CustomerBalance() As Decimal
        Get
            Return _CustomerBalance
        End Get
    End Property

    Public ReadOnly Property DepositsReceived() As Decimal
        Get
            Return _DepositsReceived
        End Get
    End Property

    Public ReadOnly Property OnOrderAmount() As Decimal
        Get
            Return _OnOrderAmount
        End Get
    End Property

    Public ReadOnly Property OrderedNotInvoicedAmount() As Decimal
        Get
            Return _OrderedNotInvoicedAmount
        End Get
    End Property

    Public ReadOnly Property OrderedNotInvoicedCredit() As Decimal
        Get
            Return _OrderedNotInvoicedCredit
        End Get
    End Property

    Public ReadOnly Property UnpostedCashAmount() As Decimal
        Get
            Return _UnpostedCashAmount
        End Get
    End Property

    Public ReadOnly Property UnpostedOtherCashAmount() As Decimal
        Get
            Return _UnpostedOtherCashAmount
        End Get
    End Property

    Public ReadOnly Property UnpostedOtherSalesAmount() As Decimal
        Get
            Return _UnpostedOtherSalesAmount
        End Get
    End Property

    Public ReadOnly Property UnpostedSalesAmount() As Decimal
        Get
            Return _UnpostedOtherSalesAmount
        End Get
    End Property


    Public ReadOnly Property Address1() As String
        Get
            Return _Addr1.Trim
        End Get
    End Property

    Public ReadOnly Property Address2() As String
        Get
            Return _Addr2.Trim
        End Get
    End Property

    Public ReadOnly Property Address3() As String
        Get
            Return _Addr3.Trim
        End Get
    End Property

    Public ReadOnly Property City() As String
        Get
            Return _City.Trim
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

    Public ReadOnly Property Contact() As String
        Get
            Return _Contact.Trim
        End Get
    End Property

    Public ReadOnly Property Country() As String
        Get
            Return _Country.Trim
        End Get
    End Property

    Public ReadOnly Property CreditLimit() As Decimal
        Get
            Return _CreditLimit
        End Get
    End Property

    Public ReadOnly Property CustomerNumber() As String
        Get
            Return _CustomerNumber.Trim
        End Get
    End Property

    Public ReadOnly Property CusTyp() As String
        Get
            Return _CusTyp.Trim
        End Get
    End Property

    Public ReadOnly Property FOB() As String
        Get
            Return _FOB.Trim
        End Get
    End Property

    Public ReadOnly Property MailingAddress() As String
        Get
            Return _MailingAddress
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _Name.Trim
        End Get
    End Property

    Public Function OnHold() As Boolean
        Return Status = "H"
    End Function

    Public ReadOnly Property Phone() As String
        Get
            Return _Phone.Trim
        End Get
    End Property

    Public ReadOnly Property SalesRep() As String
        Get
            Return _SalesRep.Trim
        End Get
    End Property

    Public ReadOnly Property SalesRepName() As String
        Get
            Return _SalesRepName
        End Get
    End Property

    Public ReadOnly Property ShipCode() As String
        Get
            Return _ShipCode.Trim
        End Get
    End Property

    Public ReadOnly Property ShipVia() As String
        Get
            Return _ShipVia.Trim
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

    Public ReadOnly Property TermCode() As String
        Get
            Return _TermCode.Trim
        End Get
    End Property

    Public ReadOnly Property ZipCode() As String
        Get
            Return _ZipCode.Trim
        End Get
    End Property

    ' Allow someone to generically create the class so they can invoice the 
    ' general Read() method to return a list of all customers
    Public Sub New()
        ' Call the base class new function
        MyBase.New()
    End Sub

    Public Sub New(ByVal passCustomer As String, ByVal Source As DataSource)
        ' Call the base class new function
        MyBase.New()

        _CustomerNumber = passCustomer
        Read(passCustomer, Source)

        ' Set up a new customer addresses class
        Addresses = New ShippingMasterClass(passCustomer, Source)

        ' Only set up the price quote if coming from sellars
        If Source = DataSource.Sellars Then
            ' Set up a new price quote class
            PriceQuote = New PriceQuoteClass(passCustomer, SalesRep, Me)
        End If
    End Sub

    Public Function Read(ByVal Source As DataSource) As Object
        ' Run the stored procedure and return a datareader
        If Source = DataSource.Sellars Then
            Dim oSQL As New SqlService(ConnectionString)
            Return oSQL.RunProcReader("ReadAllCustomers")
        Else
            Return ReadAllMaxCustomers()
        End If
    End Function

    Public Function Read(ByVal SalesRep As String, ByVal junkflag As Boolean) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@SalesRep", SqlDbType.NVarChar, 7, SalesRep, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("ReadSalesRepCustomers")
    End Function

    Public Sub Read(ByVal passcustomer As String, ByVal Source As DataSource)
        If Source = DataSource.Max Then
            ReadMaxCustomer(passcustomer)
        Else
            ReadSellarsCustomer(passcustomer)
        End If
    End Sub

    Private Function FullAddress() As String
        Dim Address As String

        Address = Name + Chr(13) + Chr(10)
        Address = Address + Address1 + Chr(13) + Chr(10)
        If Address2 <> "" Then
            Address = Address + Address2 + Chr(13) + Chr(10)
        End If
        If City <> "" Or State <> "" Or ZipCode <> "" Then
            If State = "" Then
                Address = Address + City + " " + ZipCode + Chr(13) + Chr(10)
            Else
                Address = Address + City + ", " + State + " " + ZipCode + Chr(13) + Chr(10)
            End If
        End If
        If Country <> "" Then
            Address = Address + Country
        End If

        Return Address
    End Function

    Private Function ReadAllMaxCustomers() As SqlDataReader
        ' Add is only currently available for the max datasource
        Dim strSQL As String = "select ltrim(rtrim(CUSTID_23)) as ID, Name_23 as Name " & _
                               "FROM Customer_Master " & _
                               "ORDER BY Name_23 asc"
        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Return cmd.ExecuteReader(CommandBehavior.CloseConnection)
    End Function

    Private Sub ReadMaxCustomer(ByVal passcustomer As String)
        Dim strSQL As String = "select NAME_23, STATUS_23, COMNT1_23, COMNT2_23, CLIMIT_23, SLSREP_23, ADDR1_23, ADDR2_23, ADDR3_23, CITY_23, CNTRY_23, STATE_23, ZIPCD_23, TERMS_23, SHPCDE_23, SHPVIA_23, FOB_23, TAXABL_23, TXCDE1_23, TXCDE2_23, TXCDE3_23, CUSTYP_23, CNTCT_23, PHONE_23 " & _
                               "From ""Customer_Master"" " & _
                               "Where CUSTID_23 = '" & passcustomer.Trim & "'"

        ' Set up the new Sql command
        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then

                        If IsDBNull(dr("NAME_23")) Then
                            _Name = ""
                        Else
                            _Name = dr("NAME_23")
                        End If
                        If IsDBNull(dr("STATUS_23")) Then
                            _Status = ""
                        Else
                            _Status = dr("STATUS_23")
                        End If
                        If IsDBNull(dr("CNTCT_23")) Then
                            _Contact = ""
                        Else
                            _Contact = dr("CNTCT_23")
                        End If
                        If IsDBNull(dr("COMNT1_23")) Then
                            _Comment1 = ""
                        Else
                            _Comment1 = dr("COMNT1_23")
                        End If
                        If IsDBNull(dr("COMNT2_23")) Then
                            _Comment2 = ""
                        Else
                            _Comment2 = dr("COMNT2_23")
                        End If
                        ' If the credit limit is a zero, then 
                        ' set it to max value to reflect that zero means 
                        ' unlimited credit
                        If IsDBNull(dr("CLIMIT_23")) Then
                            _CreditLimit = _CreditLimit.MaxValue
                        Else
                            If dr("CLIMIT_23") = 0 Then
                                _CreditLimit = _CreditLimit.MaxValue
                            Else
                                _CreditLimit = dr("CLIMIT_23")
                            End If
                        End If
                        If IsDBNull(dr("ADDR1_23")) Then
                            _Addr1 = ""
                        Else
                            _Addr1 = dr("ADDR1_23")
                        End If
                        If IsDBNull(dr("ADDR2_23")) Then
                            _Addr2 = ""
                        Else
                            _Addr2 = dr("ADDR2_23")
                        End If
                        If IsDBNull(dr("ADDR3_23")) Then
                            _Addr3 = ""
                        Else
                            _Addr3 = dr("ADDR3_23")
                        End If
                        If IsDBNull(dr("CITY_23")) Then
                            _City = ""
                        Else
                            _City = dr("CITY_23")
                        End If
                        If IsDBNull(dr("STATE_23")) Then
                            _State = ""
                        Else
                            _State = dr("STATE_23")
                        End If
                        If IsDBNull(dr("ZIPCD_23")) Then
                            _ZipCode = ""
                        Else
                            _ZipCode = dr("ZIPCD_23")
                        End If
                        If IsDBNull(dr("CNTRY_23")) Then
                            _Country = ""
                        Else
                            _Country = dr("CNTRY_23")
                        End If

                        _MailingAddress = FullAddress()

                        If IsDBNull(dr("FOB_23")) Then
                            _FOB = ""
                        Else
                            _FOB = Trim(dr("FOB_23"))
                        End If
                        If IsDBNull(dr("SHPCDE_23")) Then
                            _ShipCode = ""
                        Else
                            _ShipCode = dr("SHPCDE_23")
                        End If
                        If IsDBNull(dr("SHPVIA_23")) Then
                            _ShipVia = "-1"
                        Else
                            _ShipVia = dr("SHPVIA_23")
                        End If
                        If IsDBNull(dr("TAXABL_23")) Then
                            _Taxable = "N"
                        Else
                            _Taxable = dr("TAXABL_23")
                        End If
                        If IsDBNull(dr("TXCDE1_23")) Then
                            _TaxCode1 = ""
                        Else
                            _TaxCode1 = dr("TXCDE1_23")
                        End If
                        If IsDBNull(dr("TXCDE2_23")) Then
                            _TaxCode2 = ""
                        Else
                            _TaxCode2 = dr("TXCDE2_23")
                        End If
                        If IsDBNull(dr("TXCDE3_23")) Then
                            _TaxCode3 = ""
                        Else
                            _TaxCode3 = dr("TXCDE3_23")
                        End If

                        If IsDBNull(dr("SLSREP_23")) Then
                            _SalesRep = ""
                            _SalesRepName = "Unknown"
                        Else
                            _SalesRep = dr("SLSREP_23")
                            Dim SRM As New SalesRepMasterClass(SalesRepMasterClass.DataSource.Max, SalesRep)
                            _SalesRepName = SRM.Name
                            SRM = Nothing
                        End If

                        If IsDBNull(dr("TERMS_23")) Then
                            _TermCode = "-1"
                        Else
                            _TermCode = dr("TERMS_23")
                        End If

                        If IsDBNull(dr("CUSTYP_23")) Then
                            _CusTyp = ""
                        Else
                            _CusTyp = dr("CUSTYP_23")
                        End If
                        If IsDBNull(dr("PHONE_23")) Then
                            _Phone = ""
                        Else
                            _Phone = dr("PHONE_23")
                        End If

                        Terms = New CodeMasterClass(ClassBase.DataSource.Max, "TERM", _TermCode)
                    Else
                        _Name = ""
                        _Status = ""
                        _Contact = ""
                        _Comment1 = ""
                        _Comment2 = ""
                        _CreditLimit = _CreditLimit.MaxValue
                        _Addr1 = ""
                        _Addr2 = ""
                        _Addr3 = ""
                        _City = ""
                        _State = ""
                        _ZipCode = ""
                        _Country = ""
                        _FOB = ""
                        _ShipCode = ""
                        _ShipVia = "-1"
                        _Taxable = "N"
                        _TaxCode1 = ""
                        _TaxCode2 = ""
                        _TaxCode3 = ""
                        _SalesRep = ""
                        _SalesRepName = ""
                        _TermCode = "-1"
                        _CusTyp = ""
                        _Phone = ""
                    End If

                End Using
            End Using
        End Using
    End Sub

    Private Function ReadMaxOpenOrderTotal(ByVal passcustomer As String, ByVal passType As String) As Decimal
        Dim strSQL As String = "SELECT Sum((DUEQTY_28 * PRICE_28) + TAX1_28 + TAX2_28 + Tax3_28) as OpenTotal FROM ""SO_Detail"" where CUSTID_28 = '" & passcustomer.Trim & "' and STATUS_28 = '3' and STYPE_28 = '" & passType & "'"
        Dim returnvalue As Decimal = 0
        Dim dr As SqlDataReader

        ' Set up the new Sql command and execute the query
        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)

        ' If no records were read, then return a zero, otherwise return the amount of open line items
        Try
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                If IsDBNull(dr("OpenTotal")) Then
                    returnvalue = 0
                Else
                    returnvalue = dr("OpenTotal")
                End If
            End If
        Catch ex As Exception
            returnvalue = 0
            MsgBox("Error Reading Open Order Totals, sql statement is:" & vbCrLf & vbCrLf & strSQL & vbCrLf & vbCrLf & "Error Message is:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Reading Open Order Totals")
        End Try

        ' Free up memory
        dr.Close()
        dr = Nothing
        CloseMaxConnection()

        ' Return the total value of all open line items
        Return returnvalue
    End Function

    Private Function ReadMaxOpenOrderTotal2(ByVal passcustomer As String, ByVal passType As String) As Decimal
        Dim strSQL As String = "SELECT Sum((DUEQTY_28 * PRICE_28) + TAX1_28 + TAX2_28 + Tax3_28) as OpenTotal FROM ""SO_Detail"" where CUSTID_28 = '" & passcustomer.Trim & "' and STATUS_28 = '3' and STYPE_28 = '" & passType & "'"
        Dim returnvalue As Decimal = 0
        Dim dr As SqlDataReader

        ' Set up the new Sql command and execute the query
        ' try to open another connection to the max database
        Dim MaxConnection As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
        MaxConnection.Open()

        Dim cmd As New SqlCommand(strSQL, MaxConnection)

        ' If no records were read, then return a zero, otherwise return the amount of open line items
        Try
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                If IsDBNull(dr("OpenTotal")) Then
                    returnvalue = 0
                Else
                    returnvalue = dr("OpenTotal")
                End If
            End If
            ' Close the datareader
            dr.Close()
        Catch ex As Exception
            returnvalue = 0
            MsgBox("Error Reading Open Order Totals, sql statement is:" & vbCrLf & vbCrLf & strSQL & vbCrLf & vbCrLf & "Error Message is:" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Error Reading Open Order Totals")
        End Try

        ' Free up memory
        dr = Nothing
        If MaxConnection.State = ConnectionState.Open Then
            MaxConnection.Close()
        End If
        MaxConnection.Dispose()

        ' Return the total value of all open line items
        Return returnvalue
    End Function


    Private Sub ReadSellarsCustomer(ByVal passcustomer)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, passcustomer, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadCustomerRecord")

        ' Assign the variables from the database to properties
        If dr.Read() Then
            _Name = dr("Name")
            _Status = dr("Status")
            ' Set the credit limit field.  If the value is zero, then this means
            ' MAX value, or unlimited credit
            _CreditLimit = dr("CreditLimit")
            If _CreditLimit = 0 Then
                _CreditLimit = _CreditLimit.MaxValue
            End If
            If IsDBNull(dr("SalesRep")) Then
                _SalesRep = ""
            Else
                _SalesRep = dr("SalesRep")
            End If
            _Addr1 = dr("Address1")
            If IsDBNull(dr("Address2")) Then
                _Addr2 = ""
            Else
                _Addr2 = dr("Address2")
            End If
            If IsDBNull(dr("City")) Then
                _City = ""
            Else
                _City = dr("City")
            End If
            If IsDBNull(dr("State")) Then
                _State = ""
            Else
                _State = dr("State")
            End If
            If IsDBNull(dr("ZipCode")) Then
                _ZipCode = ""
            Else
                _ZipCode = dr("ZipCode")
            End If
            _MailingAddress = FullAddress()

            _ShipVia = "-1"
            _FOB = ""
            _SalesRepName = ""
            _TaxCode1 = ""
            _TaxCode2 = ""
            _TaxCode3 = ""

            If IsDBNull(dr("Terms")) Then
                _TermCode = "-1"
            Else
                _TermCode = dr("Terms")
            End If
            Terms = New CodeMasterClass(ClassBase.DataSource.Sellars, "TERM", dr("Terms"))

            ' close the data reader
            dr.Close()
        End If

        ' free up the memory reserved by the datareader
        dr = Nothing
    End Sub

    Public Function IsOverLimit(ByVal passcustomer As String) As Boolean
        Dim CreditUsed As Decimal = 0
        Try
            ' Start a task to get the total amount of ordered, but not invoiced customer order line items
            Dim onicuTask As Task(Of Decimal) = Task.Factory.StartNew(Function() As Decimal
                                                                          Return ReadMaxOpenOrderTotal2(passcustomer, "CU")
                                                                      End Function)

            ' Get the total amount of ordered, but not invoices credit memo line items
            Dim onicrTask As Task(Of Decimal) = Task.Factory.StartNew(Function() As Decimal
                                                                          Return ReadMaxOpenOrderTotal2(passcustomer, "CR")
                                                                      End Function)

            ' Get the total amount of ordered, but not invoices credit memo line items
            Dim onigcdTask As Task(Of Decimal) = Task.Factory.StartNew(Function() As Decimal
                                                                           Return GetCreditData(passcustomer)
                                                                       End Function)

            ' Wait for all the tasks running on threads to to complete before continuing
            Task.WaitAll(onicuTask, onicrTask, onigcdTask)

            _OrderedNotInvoicedAmount = onicuTask.Result
            _OrderedNotInvoicedCredit = onicrTask.Result
            CreditUsed = onigcdTask.Result
        Catch
        End Try

        ' Check if the credit used is over the credit limit
        If (CreditUsed + OrderedNotInvoicedAmount - OrderedNotInvoicedCredit) > CreditLimit Then
            _AmountOverLimit = (CreditUsed + OrderedNotInvoicedAmount - OrderedNotInvoicedCredit) - CreditLimit
            Return True
        Else
            _AmountOverLimit = 0
            Return False
        End If

    End Function

    Private Function GetCreditData(ByVal passcustomer As String) As Decimal
        ' Initialize the credit used temporary variable
        Dim CreditUsed As Decimal = 0

        ' Set up the SQL statement to return the information from the customers credit record in Dynamics
        Dim strSQL As String = "select CUSTBLNC, UNPSTDSA, UNPSTOSA, ONORDAMT, UNPSTDCA, UNPSTOCA, DEPRECV " & _
                               "from RM00103 " & _
                               "where CUSTNMBR = @CUSTID"

        Dim dr As SqlDataReader

        '     Open the connection to the Dynamics tables
        '     and execute the SQL statement prepared above
        OpenDynamicsConnection()
        Try
            ' Set up a new SQL command object
            Dim cmd As New SqlCommand(strSQL, DynamicsConnection)
            cmd.CommandType = CommandType.Text

            ' Create a parameter to pass into the command that contains the customer id
            Dim prmCust As SqlParameter = New SqlParameter("@CUSTID", SqlDbType.NVarChar, 15)
            prmCust.Value = passcustomer
            cmd.Parameters.Add(prmCust)

            ' Set up a new sql data reader object that will contain the results of the command after it has been
            ' executed against the database
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                _CustomerBalance = dr("CUSTBLNC")
                _UnpostedSalesAmount = dr("UNPSTDSA")
                _UnpostedOtherSalesAmount = dr("UNPSTOSA")
                _OnOrderAmount = dr("ONORDAMT")
                _UnpostedCashAmount = dr("UNPSTDCA")
                _UnpostedOtherCashAmount = dr("UNPSTOCA")
                _DepositsReceived = dr("DEPRECV")
                CreditUsed = dr("CUSTBLNC") + dr("UNPSTDSA") + dr("UNPSTOSA") + dr("ONORDAMT") - dr("UNPSTDCA") - dr("UNPSTOCA") - dr("DEPRECV")
            Else
                _CustomerBalance = 0
                _UnpostedSalesAmount = 0
                _UnpostedOtherSalesAmount = 0
                _OnOrderAmount = 0
                _UnpostedCashAmount = 0
                _UnpostedOtherCashAmount = 0
                _DepositsReceived = 0
            End If
        Catch Ex As Exception
            MsgBox("Error reading Dynamics Credit Limit Data:" & vbCrLf & vbCrLf & Ex.Message)
            _CustomerBalance = 0
            _UnpostedSalesAmount = 0
            _UnpostedOtherSalesAmount = 0
            _OnOrderAmount = 0
            _UnpostedCashAmount = 0
            _UnpostedOtherCashAmount = 0
            _DepositsReceived = 0
            CreditUsed = 0
        End Try

        ' Free up the memory used for the dynamics connection
        dr.Close()
        dr = Nothing
        CloseDynamicsConnection()

        Return CreditUsed
    End Function

    Public Sub Add(ByVal CustId As String, ByVal SalesRepCode As String, ByVal Status As String, ByVal CustType As String, ByVal Name As String, ByVal Address1 As String, _
                   ByVal Address2 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, _
                   ByVal Territory As String, ByVal Contact As String, ByVal Phone As String, ByVal Email1 As String, ByVal Email2 As String, _
                   ByVal Telex As String, ByVal Fax As String, ByVal Comment1 As String, ByVal Comment2 As String, ByVal Taxable As String, _
                   ByVal TaxCode1 As String, ByVal TaxCode2 As String, ByVal TaxCode3 As String, ByVal Terms As String, _
                   ByVal DiscountRate As Decimal, ByVal CreditLimit As Decimal, ByVal Statements As String, ByVal FinanceCharge As String, _
                   ByVal ShipVia As String, ByVal FOB As String, _
                   ByVal ApplyTo As String, ByVal AllowBackOrders As String, ByVal CurrencyCode As String, _
                   ByVal UDFKEY As String, ByVal UDFREF As String, ByVal CreatedBy As String)

        Dim strSQL As String = "INSERT INTO Customer_Master (CUSTID_23, SLSREP_23, STATUS_23, CUSTYP_23, NAME_23, ADDR1_23, ADDR2_23, " & _
                                "CITY_23, STATE_23, ZIPCD_23, CNTRY_23, SLSTER_23, CNTCT_23, PHONE_23, EMAIL1_23, EMAIL2_23, TELEX_23, " & _
                                "FAXNO_23, COMNT1_23, COMNT2_23, TAXABL_23, TAXNUM_23, TXCDE1_23, TXCDE2_23, TXCDE3_23, TERMS_23, DSCRTE_23, " & _
                                "CLIMIT_23, STMNTS_23, FINCHG_23, SHPVIA_23, XURR_23, FOB_23, SLSMTD_23, COGMTD_23, SLSYTD_23, COGYTD_23, " & _
                                "SLSLYR_23, COGLYR_23, UNPORD_23, NEWDTE_23, DISCPF_23, ALWBCK_23, CHGDTE_23, FILL02_23, TAXPRV_23, CURR_23, " & _
                                "COMMIS_23, ADDR3_23, ADDR4_23, ADDR5_23, ADDR6_23, MCOMP_23, MSITE_23, UDFKEY_23, UDFREF_23, SHPCDE_23, " & _
                                "SHPTHRU_23, XDFINT_23, XDFFLT_23, XDFBOL_23, XDFDTE_23, XDFTXT_23, FILLER_23, CreatedBy, CreationDate, " & _
                                "ModifiedBy, ModificationDate, REGDATE_23, EXPDATE_23, FISCODE_23, VATSUSP_23, VTAXNUM_23, JOURNAL_23" & _
                                ") VALUES (" & _
                                "@CUSTID, @SLSREP, @STATUS, @CUSTYP, @NAME, @ADDR1, @ADDR2, @CITY, @STATE, @ZIPCD, " & _
                                "@CNTRY, @SLSTER, @CNTCT, @PHONE, @EMAIL1, @EMAIL2, @TELEX, @FAXNO, @CMNT1, @CMNT2, @TAXABL, '', @TXCDE1, " & _
                                "@TXCDE2, @TXCDE3, @TERMS, @DSCRTE, @CLIMIT, @STMNTS, @FINCHG, @SHPVIA, '', @FOB, 0, 0, 0, 0, 0, 0, 0, " & _
                                "CONVERT(DATE,GETDATE()), @DISCPF, @ALWBCK, CONVERT(DATE,GETDATE()), '', '', @CURR, 0, '', '', '', '', '', '', " & _
                                "@UDFKEY, @UDFREF, '', '', 0, 0, '', CONVERT(DATE,GETDATE()), '', '', 'EDI', GETDATE(), 'EDI', GETDATE(), " & _
                                "CONVERT(DATE,GETDATE()), CONVERT(DATE,GETDATE()), '', 'N', '', ''" & _
                                ")"

        OpenMaxConnection()
        Dim cmd As SqlCommand = New SqlCommand(strSQL, MaxConnection)

        Dim paramCustId As New SqlParameter("CustId", SqlDbType.Char, 20)
        paramCustId.Direction = ParameterDirection.Input
        paramCustId.Value = CustId
        cmd.Parameters.Add(paramCustId)

        Dim paramSalesRep As New SqlParameter("Slsrep", SqlDbType.Char, 7)
        paramSalesRep.Direction = ParameterDirection.Input
        paramSalesRep.Value = SalesRepCode
        cmd.Parameters.Add(paramSalesRep)

        Dim paramStatus As New SqlParameter("Status", SqlDbType.Char, 1)
        paramStatus.Direction = ParameterDirection.Input
        paramStatus.Value = Status
        cmd.Parameters.Add(paramStatus)

        Dim paramCustType As New SqlParameter("Custyp", SqlDbType.Char, 3)
        paramCustType.Direction = ParameterDirection.Input
        paramCustType.Value = CustType
        cmd.Parameters.Add(paramCustType)

        Dim paramCustName As New SqlParameter("Name", SqlDbType.Char, 30)
        paramCustName.Direction = ParameterDirection.Input
        paramCustName.Value = Name
        cmd.Parameters.Add(paramCustName)

        Dim paramAddress1 As New SqlParameter("Addr1", SqlDbType.Char, 30)
        paramAddress1.Direction = ParameterDirection.Input
        paramAddress1.Value = Address1
        cmd.Parameters.Add(paramAddress1)

        Dim paramAddress2 As New SqlParameter("Addr2", SqlDbType.Char, 30)
        paramAddress2.Direction = ParameterDirection.Input
        paramAddress2.Value = Address2
        cmd.Parameters.Add(paramAddress2)

        Dim paramCity As New SqlParameter("City", SqlDbType.Char, 30)
        paramCity.Direction = ParameterDirection.Input
        paramCity.Value = City
        cmd.Parameters.Add(paramCity)

        Dim paramState As New SqlParameter("State", SqlDbType.Char, 30)
        paramState.Direction = ParameterDirection.Input
        paramState.Value = State
        cmd.Parameters.Add(paramState)

        Dim paramZip As New SqlParameter("ZipCd", SqlDbType.Char, 30)
        paramZip.Direction = ParameterDirection.Input
        paramZip.Value = Zip
        cmd.Parameters.Add(paramZip)

        Dim paramCountry As New SqlParameter("Cntry", SqlDbType.Char, 30)
        paramCountry.Direction = ParameterDirection.Input
        paramCountry.Value = Country
        cmd.Parameters.Add(paramCountry)

        Dim paramTerritory As New SqlParameter("Slster", SqlDbType.Char, 20)
        paramTerritory.Direction = ParameterDirection.Input
        paramTerritory.Value = Territory
        cmd.Parameters.Add(paramTerritory)

        Dim paramContact As New SqlParameter("Cntct", SqlDbType.Char, 20)
        paramContact.Direction = ParameterDirection.Input
        paramContact.Value = Contact
        cmd.Parameters.Add(paramContact)

        Dim paramPhone As New SqlParameter("Phone", SqlDbType.Char, 20)
        paramPhone.Direction = ParameterDirection.Input
        paramPhone.Value = Phone
        cmd.Parameters.Add(paramPhone)

        Dim paramEmail1 As New SqlParameter("Email1", SqlDbType.Char, 200)
        paramEmail1.Direction = ParameterDirection.Input
        paramEmail1.Value = Email1
        cmd.Parameters.Add(paramEmail1)

        Dim paramEmail2 As New SqlParameter("Email2", SqlDbType.Char, 200)
        paramEmail2.Direction = ParameterDirection.Input
        paramEmail2.Value = Email2
        cmd.Parameters.Add(paramEmail2)

        Dim paramTelex As New SqlParameter("Telex", SqlDbType.Char, 20)
        paramTelex.Direction = ParameterDirection.Input
        paramTelex.Value = Telex
        cmd.Parameters.Add(paramTelex)

        Dim paramFax As New SqlParameter("FaxNo", SqlDbType.Char, 20)
        paramFax.Direction = ParameterDirection.Input
        paramFax.Value = Fax
        cmd.Parameters.Add(paramFax)

        Dim paramComment1 As New SqlParameter("Cmnt1", SqlDbType.Char, 30)
        paramComment1.Direction = ParameterDirection.Input
        paramComment1.Value = Comment1
        cmd.Parameters.Add(paramComment1)

        Dim paramComment2 As New SqlParameter("Cmnt2", SqlDbType.Char, 30)
        paramComment2.Direction = ParameterDirection.Input
        paramComment2.Value = Comment2
        cmd.Parameters.Add(paramComment2)

        Dim paramTaxable As New SqlParameter("Taxabl", SqlDbType.Char, 1)
        paramTaxable.Direction = ParameterDirection.Input
        paramTaxable.Value = Taxable
        cmd.Parameters.Add(paramTaxable)

        Dim paramTaxCode1 As New SqlParameter("Txcde1", SqlDbType.Char, 7)
        paramTaxCode1.Direction = ParameterDirection.Input
        paramTaxCode1.Value = TaxCode1
        cmd.Parameters.Add(paramTaxCode1)

        Dim paramTaxCode2 As New SqlParameter("Txcde2", SqlDbType.Char, 7)
        paramTaxCode2.Direction = ParameterDirection.Input
        paramTaxCode2.Value = TaxCode2
        cmd.Parameters.Add(paramTaxCode2)

        Dim paramTaxCode3 As New SqlParameter("Txcde3", SqlDbType.Char, 7)
        paramTaxCode3.Direction = ParameterDirection.Input
        paramTaxCode3.Value = TaxCode3
        cmd.Parameters.Add(paramTaxCode3)

        Dim paramTerms As New SqlParameter("Terms", SqlDbType.Char, 2)
        paramTerms.Direction = ParameterDirection.Input
        paramTerms.Value = Terms
        cmd.Parameters.Add(paramTerms)

        Dim paramDiscountRate As New SqlParameter("Dscrte", SqlDbType.Real)
        paramDiscountRate.Direction = ParameterDirection.Input
        paramDiscountRate.Value = DiscountRate
        cmd.Parameters.Add(paramDiscountRate)

        Dim paramCreditLimit As New SqlParameter("Climit", SqlDbType.Float)
        paramCreditLimit.Direction = ParameterDirection.Input
        paramCreditLimit.Value = CreditLimit
        cmd.Parameters.Add(paramCreditLimit)

        Dim paramStatements As New SqlParameter("Stmnts", SqlDbType.Char, 1)
        paramStatements.Direction = ParameterDirection.Input
        paramStatements.Value = Statements
        cmd.Parameters.Add(paramStatements)

        Dim paramFinanceCharge As New SqlParameter("Finchg", SqlDbType.Char, 1)
        paramFinanceCharge.Direction = ParameterDirection.Input
        paramFinanceCharge.Value = FinanceCharge
        cmd.Parameters.Add(paramFinanceCharge)

        Dim paramShipVia As New SqlParameter("Shpvia", SqlDbType.Char, 2)
        paramShipVia.Direction = ParameterDirection.Input
        paramShipVia.Value = ShipVia
        cmd.Parameters.Add(paramShipVia)

        Dim paramFOB As New SqlParameter("Fob", SqlDbType.Char, 15)
        paramFOB.Direction = ParameterDirection.Input
        paramFOB.Value = FOB
        cmd.Parameters.Add(paramFOB)

        Dim paramApplyTo As New SqlParameter("Discpf", SqlDbType.Char, 1)
        paramApplyTo.Direction = ParameterDirection.Input
        paramApplyTo.Value = ApplyTo
        cmd.Parameters.Add(paramApplyTo)

        Dim paramAllowBackOrders As New SqlParameter("Alwbck", SqlDbType.Char, 1)
        paramAllowBackOrders.Direction = ParameterDirection.Input
        paramAllowBackOrders.Value = AllowBackOrders
        cmd.Parameters.Add(paramAllowBackOrders)

        Dim paramCurrencyCode As New SqlParameter("Curr", SqlDbType.Char, 3)
        paramCurrencyCode.Direction = ParameterDirection.Input
        paramCurrencyCode.Value = CurrencyCode
        cmd.Parameters.Add(paramCurrencyCode)

        Dim paramUDFKEY As New SqlParameter("Udfkey", SqlDbType.Char, 15)
        paramUDFKEY.Direction = ParameterDirection.Input
        paramUDFKEY.Value = UDFKEY
        cmd.Parameters.Add(paramUDFKEY)

        Dim paramUDFREF As New SqlParameter("Udfref", SqlDbType.Char, 25)
        paramUDFREF.Direction = ParameterDirection.Input
        paramUDFREF.Value = UDFREF
        cmd.Parameters.Add(paramUDFREF)

        cmd.ExecuteNonQuery()
        CloseMaxConnection()
    End Sub

End Class

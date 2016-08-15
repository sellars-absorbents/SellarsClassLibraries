Imports Sellars.SQL
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient

Public Class ShippingMasterClass
    Inherits ClassBase

    Private AddressCollection As New ArrayList

    Private _Addr1 As String
    Private _Addr2 As String
    Private _Addr3 As String
    Private _City As String
    Private _MailingAddress As String
    Private _Name As String
    Private _ShipCode As String
    Private _State As String
    Private _TaxCode1 As String
    Private _TaxCode2 As String
    Private _TaxCode3 As String
    Private _Country As String
    Private _ZipCode As String
    Private _Source As DataSource
    Private _UsrDefKey As String

    Private ReadOnly Property Source() As DataSource
        Get
            Return _Source
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

    Public ReadOnly Property Country() As String
        Get
            Return _Country.Trim
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

    Public ReadOnly Property ShippingCode() As String
        Get
            Return _ShipCode
        End Get
    End Property

    Public ReadOnly Property State() As String
        Get
            Return _State.Trim
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

    Public ReadOnly Property ZipCode() As String
        Get
            Return _ZipCode.Trim
        End Get
    End Property

    Public ReadOnly Property UsrDefKey() As String
        Get
            Return _UsrDefKey
        End Get
    End Property

    Public Sub New(ByVal value As DataSource)
        _Source = value
    End Sub

    Public Sub New(ByVal Customer As String, ByVal value As DataSource)
        _Source = value
    End Sub

    Public Sub New(ByVal CustomerCode As String, ByVal passAddressCode As String, ByVal value As DataSource)
        _Source = value
        Read(CustomerCode, passAddressCode, _Source)
    End Sub

    Public Function Add(ByVal value As DataSource, ByVal CustID As String, ByVal ShipCode As String, ByVal Name As String, ByVal Addr1 As String, ByVal Addr2 As String, ByVal Addr3 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, ByVal TaxCode1 As String, ByVal TaxCode2 As String, ByVal TaxCode3 As String, ByVal UdfKey As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim TaxAbl As String = "Y"
        Dim TxExem As String = ""
        Dim TaxPrv As String = ""
        Dim TrnTme As String = ""
        Dim Addr4 As String = ""
        Dim Addr5 As String = ""
        Dim Addr6 As String = ""
        Dim Mcomp As String = ""
        Dim Msite As String = ""
        'Dim UdfKey As String = ""
        Dim UdfRef As String = ""
        Dim Filler As String = ""

        Dim strSQL As String = "Insert into ""Shipping_Master"" (CUSTID_24, SHPCDE_24, NAME_24, ADDR1_24, ADDR2_24, CITY_24, STATE_24, ZIPCD_24, CNTRY_24, TXCDE1_24, TXCDE2_24, TXCDE3_24, TAXABL_24, TXEXEM_24, TAXPRV_24, TRNTME_24, ADDR3_24, ADDR4_24, ADDR5_24, ADDR6_24, MCOMP_24, MSITE_24, UDFKEY_24, UDFREF_24, FILLER_24) " & _
                               "values ('" & CustID & "','" & ShipCode.ToUpper.Trim & "','" & Name & "','" & Addr1 & "','" & Addr2 & "','" & City & "','" & State & "','" & Zip & "','" & Country & "','" & TaxCode1 & "','" & TaxCode2 & "','" & TaxCode3 & "','" & TaxAbl & "','" & TxExem & "','" & TaxPrv & "','" & TrnTme & "','" & Addr3 & "','" & Addr4 & "','" & Addr5 & "','" & Addr6 & "','" & Mcomp & "','" & Msite & "','" & UdfKey & "','" & UdfRef & "','" & Filler & "')"
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        cmd.ExecuteNonQuery()
        CloseMaxConnection()
    End Function

    Public Function Add(ByVal value As DataSource, ByVal CustID As String, ByVal ShipCode As String, ByVal Name As String, ByVal Addr1 As String, ByVal Addr2 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, ByVal TaxCode1 As String, ByVal TaxCode2 As String, ByVal TaxCode3 As String, ByVal UdfKey As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim TaxAbl As String = "Y"
        Dim TxExem As String = ""
        Dim TaxPrv As String = ""
        Dim TrnTme As String = ""
        Dim Addr3 As String = ""
        Dim Addr4 As String = ""
        Dim Addr5 As String = ""
        Dim Addr6 As String = ""
        Dim Mcomp As String = ""
        Dim Msite As String = ""
        'Dim UdfKey As String = ""
        Dim UdfRef As String = ""
        Dim Filler As String = ""

        Dim strSQL As String = "Insert into ""Shipping_Master"" (CUSTID_24, SHPCDE_24, NAME_24, ADDR1_24, ADDR2_24, CITY_24, STATE_24, ZIPCD_24, CNTRY_24, TXCDE1_24, TXCDE2_24, TXCDE3_24, TAXABL_24, TXEXEM_24, TAXPRV_24, TRNTME_24, ADDR3_24, ADDR4_24, ADDR5_24, ADDR6_24, MCOMP_24, MSITE_24, UDFKEY_24, UDFREF_24, FILLER_24) " & _
                               "values ('" & CustID & "','" & ShipCode.ToUpper.Trim & "','" & Name & "','" & Addr1 & "','" & Addr2 & "','" & City & "','" & State & "','" & Zip & "','" & Country & "','" & TaxCode1 & "','" & TaxCode2 & "','" & TaxCode3 & "','" & TaxAbl & "','" & TxExem & "','" & TaxPrv & "','" & TrnTme & "','" & Addr3 & "','" & Addr4 & "','" & Addr5 & "','" & Addr6 & "','" & Mcomp & "','" & Msite & "','" & UdfKey & "','" & UdfRef & "','" & Filler & "')"
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        cmd.ExecuteNonQuery()
        CloseMaxConnection()
    End Function

    Public Function Read(ByVal CustomerCode As String, ByVal value As DataSource) As Object
        If value = DataSource.Max Then
            Return ReadMaxShipTo(CustomerCode)
        Else
            ' Read the list of customer addresses
            ReadCustomerAddresses(CustomerCode)
            ' Return the address arraylist
            Return AddressCollection
        End If
    End Function

    Public Sub Read(ByVal CustomerCode As String, ByVal ShipCode As String, ByVal value As DataSource)
        If value = DataSource.Max Then
            ReadMaxShipTo(CustomerCode, ShipCode)
        Else
            ' Read the list of customer addresses
            ReadCustomerAddresses(CustomerCode)
            ' Get the address specified by the ship code
            Dim x As AddressClass
            For Each x In AddressCollection
                If x.ShippingCode = ShipCode Then
                    ' set the class variables equal to the address class values
                    _ShipCode = ShipCode
                    _Name = x.Name
                    _Addr1 = x.Addr1
                    _Addr2 = x.Addr2
                    _City = x.City
                    _State = x.State
                    _ZipCode = x.ZipCode
                    _MailingAddress = x.FullAddress

                    ' Exit the for loop
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Function ReadMaxShipTo(ByVal CustomerCode) As DataSet
        ' Add is only currently available for the max datasource
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Get the default address from the customer master record
        Dim strSQL As String = "Select 'Default' as Code, NAME_23 as Name, ADDR1_23 as Address1, ADDR2_23 as Address2, ADDR3_23 as Address3, CITY_23 as City, CNTRY_23 as Country, STATE_23 as State, ZIPCD_23 as ZipCode, TXCDE1_23 as TaxCode1, TXCDE2_23 as TaxCode2, TXCDE3_23 as TaxCode3, '' as UserDefKey " & _
                              "From ""Customer_Master"" " & _
                              "Where CUSTID_23 = '" & CustomerCode.Trim & "'"
        Dim da1 As New SqlDataAdapter(strSQL.ToString, MaxConnection)
        Dim ds As New DataSet
        da1.Fill(ds, "ShipVias")

        ' Get the shipping addresses from the shipping master table
        strSQL = "Select SHPCDE_24 as Code, NAME_24 as Name, ADDR1_24 as Address1, ADDR2_24 as Address2, ADDR3_24 as Address3, CITY_24 as City, CNTRY_24 as Country, STATE_24 as State, ZIPCD_24 as ZipCode, TXCDE1_24 as TaxCode1, TXCDE2_24 as TaxCode2, TXCDE3_24 as TaxCode3, COALESCE(UDFKEY_24, '') as UsrDefKey " & _
                              "From ""Shipping_Master"" " & _
                              "Where CUSTID_24 = '" & CustomerCode.Trim & "'"
        Dim da As New SqlDataAdapter(strSQL.ToString, MaxConnection)
        da.Fill(ds, "ShipVias")

        CloseMaxConnection()

        Return ds
    End Function

    Private Sub ReadMaxShipTo(ByVal CustomerCode As String, ByVal passShipCode As String)

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim strSQL As String = "Select Name_24, ADDR1_24, ADDR2_24, ADDR3_24, CITY_24, CNTRY_24, STATE_24, ZIPCD_24, TXCDE1_24, TXCDE2_24, TXCDE3_24, UDFKEY_24 as UsrDefKey " & _
                               "From ""Shipping_Master"" " & _
                               "Where CUSTID_24 = '" & CustomerCode.Trim & "' and SHPCDE_24 = '" & passShipCode & "'"
        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim ShipMasterReader As SqlDataReader = cmd.ExecuteReader()
        ShipMasterReader.Read()

        _ShipCode = passShipCode

        If IsDBNull(ShipMasterReader("ADDR1_24")) Then
            _Addr1 = ""
        Else
            _Addr1 = ShipMasterReader("ADDR1_24")
        End If

        If IsDBNull(ShipMasterReader("ADDR2_24")) Then
            _Addr2 = ""
        Else
            _Addr2 = ShipMasterReader("ADDR2_24")
        End If

        If IsDBNull(ShipMasterReader("ADDR3_24")) Then
            _Addr3 = ""
        Else
            _Addr3 = ShipMasterReader("ADDR3_24")
        End If

        If IsDBNull(ShipMasterReader("CITY_24")) Then
            _City = ""
        Else
            _City = ShipMasterReader("CITY_24")
        End If

        If IsDBNull(ShipMasterReader("STATE_24")) Then
            _State = ""
        Else
            _State = ShipMasterReader("STATE_24")
        End If

        If IsDBNull(ShipMasterReader("ZIPCD_24")) Then
            _ZipCode = ""
        Else
            _ZipCode = ShipMasterReader("ZIPCD_24")
        End If

        If IsDBNull(ShipMasterReader("CNTRY_24")) Then
            _Country = ""
        Else
            _Country = ShipMasterReader("CNTRY_24")
        End If

        If IsDBNull(ShipMasterReader("NAME_24")) Then
            _Name = ""
        Else
            _Name = ShipMasterReader("NAME_24")
        End If

        If IsDBNull(ShipMasterReader("TXCDE1_24")) Then
            _TaxCode1 = ""
        Else
            _TaxCode1 = ShipMasterReader("TXCDE1_24")
        End If

        If IsDBNull(ShipMasterReader("TXCDE2_24")) Then
            _TaxCode2 = ""
        Else
            _TaxCode2 = ShipMasterReader("TXCDE2_24")
        End If

        If IsDBNull(ShipMasterReader("TXCDE3_24")) Then
            _TaxCode3 = ""
        Else
            _TaxCode3 = ShipMasterReader("TXCDE3_24")
        End If

        If IsDBNull(ShipMasterReader("UsrDefKey")) Then
            _UsrDefKey = ""
        Else
            _UsrDefKey = ShipMasterReader("UsrDefKey").ToString()
        End If

        ShipMasterReader.Close()
        ShipMasterReader = Nothing
        CloseMaxConnection()
    End Sub

    Private Sub ReadCustomerAddresses(ByVal CustomerCode As String)
        ' Clear the collection of addresses
        AddressCollection.Clear()

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CustomerCode, ParameterDirection.Input)

        ' Run the stored procedure
        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadCustomerShippingMasterRecords")

        ' The first result set will return the bill to address
        While dr.Read
            ReadAddress(dr)
        End While

        ' The second resultset will return all shipping addresses
        dr.NextResult()
        While dr.Read
            ReadAddress(dr)
        End While

        ' Close the datareader, and free up memory
        dr.Close()
        dr = Nothing
    End Sub

    Private Sub ReadAddress(ByVal dr As SqlDataReader)
        Dim newAddress As New AddressClass

        newAddress.ShippingCode = dr("ShipCode")
        newAddress.Name = dr("Name")

        If IsDBNull(dr("Address1")) Then
            newAddress.Addr1 = ""
        Else
            newAddress.Addr1 = dr("Address1")
        End If

        If IsDBNull(dr("Address2")) Then
            newAddress.Addr2 = ""
        Else
            newAddress.Addr2 = dr("Address2")
        End If

        If IsDBNull(dr("Address3")) Then
            newAddress.Addr3 = ""
        Else
            newAddress.Addr3 = dr("Address3")
        End If

        If IsDBNull(dr("City")) Then
            newAddress.City = ""
        Else
            newAddress.City = dr("City")
        End If

        If IsDBNull(dr("State")) Then
            newAddress.State = ""
        Else
            newAddress.State = dr("State")
        End If

        If IsDBNull(dr("ZipCode")) Then
            newAddress.ZipCode = ""
        Else
            newAddress.ZipCode = dr("ZipCode")
        End If

        If IsDBNull(dr("Country")) Then
            newAddress.Country = ""
        Else
            newAddress.Country = dr("Country")
        End If

        AddressCollection.Add(newAddress)
    End Sub

    Private Function FullAddress(ByVal passAddress As AddressClass) As String
        Dim Address As String

        Address = passAddress.Name + Chr(13) + Chr(10)
        Address = Address + passAddress.Addr1 + Chr(13) + Chr(10)
        If passAddress.Addr2 <> "" Then
            Address = Address + passAddress.Addr2 + Chr(13) + Chr(10)
        End If
        If passAddress.Addr3 <> "" Then
            Address = Address + passAddress.Addr3 + Chr(13) + Chr(10)
        End If
        If passAddress.City <> "" Or passAddress.State <> "" Or passAddress.ZipCode <> "" Then
            Address = Address + passAddress.City + ", " + passAddress.State + " " + passAddress.ZipCode
        End If

        Return Address
    End Function

    Public Function Update(ByVal value As DataSource, ByVal CustID As String, ByVal ShipCode As String, ByVal Name As String, ByVal Addr1 As String, ByVal Addr2 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, ByVal TaxCode1 As String, ByVal TaxCode2 As String, ByVal TaxCode3 As String, ByVal UdfKey As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Create a string with the SQL command to run against max
        Dim strSQL As String = "Update ""Shipping_Master""  " & _
                               "set CUSTID_24 = '" & CustID & "', " & _
                               "    NAME_24 = '" & Name.Trim & "', " & _
                               "    ADDR1_24 = '" & Addr1.Trim & "', " & _
                               "    ADDR2_24 = '" & Addr2.Trim & "', " & _
                               "    CITY_24 = '" & City.Trim & "', " & _
                               "    STATE_24 = '" & State.Trim & "', " & _
                               "    ZIPCD_24 = '" & Zip.Trim & "', " & _
                               "    CNTRY_24 = '" & Country.Trim & "', " & _
                               "    TXCDE1_24 = '" & TaxCode1.Trim & "', " & _
                               "    TXCDE2_24 = '" & TaxCode2.Trim & "', " & _
                               "    TXCDE3_24 = '" & TaxCode3.Trim & "', " & _
                               "    UDFKEY_24 = '" & UdfKey.Trim & "' " & _
                               "Where CUSTID_24 = '" & CustID.Trim & "' and SHPCDE_24 = '" & ShipCode & "'"

        ' Open up a connection, run the query and close the connection
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        cmd.ExecuteNonQuery()
        CloseMaxConnection()
    End Function


    Public Function Update(ByVal value As DataSource, ByVal CustID As String, ByVal ShipCode As String, ByVal Name As String, ByVal Addr1 As String, ByVal Addr2 As String, ByVal Addr3 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, ByVal TaxCode1 As String, ByVal TaxCode2 As String, ByVal TaxCode3 As String, ByVal UdfKey As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Create a string with the SQL command to run against max
        Dim strSQL As String = "Update ""Shipping_Master""  " & _
                               "set CUSTID_24 = '" & CustID & "', " & _
                               "    NAME_24 = '" & Name.Trim & "', " & _
                               "    ADDR1_24 = '" & Addr1.Trim & "', " & _
                               "    ADDR2_24 = '" & Addr2.Trim & "', " & _
                               "    ADDR3_24 = '" & Addr3.Trim & "', " & _
                               "    CITY_24 = '" & City.Trim & "', " & _
                               "    STATE_24 = '" & State.Trim & "', " & _
                               "    ZIPCD_24 = '" & Zip.Trim & "', " & _
                               "    CNTRY_24 = '" & Country.Trim & "', " & _
                               "    TXCDE1_24 = '" & TaxCode1.Trim & "', " & _
                               "    TXCDE2_24 = '" & TaxCode2.Trim & "', " & _
                               "    TXCDE3_24 = '" & TaxCode3.Trim & "', " & _
                               "    UDFKEY_24 = '" & UdfKey.Trim & "' " & _
                               "Where CUSTID_24 = '" & CustID.Trim & "' and SHPCDE_24 = '" & ShipCode & "'"

        ' Open up a connection, run the query and close the connection
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        cmd.ExecuteNonQuery()
        CloseMaxConnection()
    End Function

End Class

Public Class AddressClass
    Private _ShippingCode As String
    Private _Name As String
    Private _Addr1 As String
    Private _Addr2 As String
    Private _Addr3 As String
    Private _City As String
    Private _Country As String
    Private _State As String
    Private _ZipCode As String
    Private _FullAddress As String

    Public Property Name() As String
        Get
            Return Trim(_Name)
        End Get
        Set(ByVal Value As String)
            _Name = Trim(Value)
        End Set
    End Property

    Public Property ShippingCode() As String
        Get
            Return _ShippingCode.Trim
        End Get
        Set(ByVal Value As String)
            _ShippingCode = Value
        End Set
    End Property

    Public Property Addr1() As String
        Get
            Return Trim(_Addr1)
        End Get
        Set(ByVal Value As String)
            _Addr1 = Trim(Value)
        End Set
    End Property

    Public Property Addr2() As String
        Get
            Return Trim(_Addr2)
        End Get
        Set(ByVal Value As String)
            _Addr2 = Trim(Value)
        End Set
    End Property

    Public Property Addr3() As String
        Get
            Return Trim(_Addr3)
        End Get
        Set(ByVal Value As String)
            _Addr3 = Trim(Value)
        End Set
    End Property

    Public Property City() As String
        Get
            Return Trim(_City)
        End Get
        Set(ByVal Value As String)
            _City = Trim(Value)
        End Set
    End Property

    Public Property Country() As String
        Get
            Return Trim(_Country)
        End Get
        Set(ByVal Value As String)
            _Country = Trim(Value)
        End Set
    End Property

    Public Property State() As String
        Get
            Return Trim(_State)
        End Get
        Set(ByVal Value As String)
            _State = Trim(Value)
        End Set
    End Property

    Public Property ZipCode() As String
        Get
            Return Trim(_ZipCode)
        End Get
        Set(ByVal Value As String)
            _ZipCode = Trim(Value)
        End Set
    End Property

    Public ReadOnly Property FullAddress() As String
        Get
            Return GetFullAddress()
        End Get
    End Property

    Private Function GetFullAddress() As String
        Dim Address As String

        Address = Name + Chr(13) + Chr(10)
        Address = Address + Addr1 + Chr(13) + Chr(10)
        If Addr2 <> "" Then
            Address = Address + Addr2 + Chr(13) + Chr(10)
        End If
        If Addr3 <> "" Then
            Address = Address + Addr3 + Chr(13) + Chr(10)
        End If
        If City <> "" Or State <> "" Or ZipCode <> "" Then
            Address = Address + City + ", " + State + " " + ZipCode
        End If

        Return Address
    End Function

End Class

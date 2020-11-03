
Imports System.Collections.Generic

Public Class SalesOrder
    Protected Friend Const DefaultDate As Date = #1/1/2000 12:01:00 AM#

    ' NOTE:  The fields in this section of the class should match exactly the fields contained within the SOEMstItem structure
    '        in the DRCLIB.SalesOrderMasterClass, in the same order for easier comparison; with the exception of the Filler fields
    Private _ORDNUM As String = ""                    ' Order Number
    Private _CUSTID As String = ""                    ' Customer ID
    Private _GLXREF As String = ""                    ' General Ledger Cross Reference Number
    Private _STYPE As String = ""                     ' Order Type ('CU', 'CR')
    Private _STATUS As String = ""                    ' Order Status ('3', '4')
    Private _CUSTPO As String = ""                    ' Customer Purchase Order Number
    Private _ORDID As String = ""                     ' Order ID
    Private _ORDDTE As Date = DefaultDate             ' Date the order was created
    Private _SHPCDE As String = ""                    ' Shipping Code
    Private _REP1 As String = ""                      ' First Sales Rep ID
    Private _SPLIT1 As Single = 0                     ' First Sales Rep ID's percentage
    Private _REP2 As String = ""                      ' Second Sales Rep ID
    Private _SPLIT2 As Single = 0                     ' Second Sales Rep ID's Percentage
    Private _REP3 As String = ""                      ' Third Sales Rep ID
    Private _SPLIT3 As Single = 0                     ' Third Sales Rep ID's Percentage
    Private _COMMIS As Single = 0                     ' Total Commission Paid on the Order
    Private _TERMS As String = ""                     ' Payment Terms ID 
    Private _SHPVIA As String = ""                    ' Ship Via ID
    Private _XURR As String = ""                      ' Old Currency Code
    Private _FOB As String = ""                       ' Freight On Board
    Private _TAXCD1 As String = ""                    ' Tax Code 1
    Private _TAXCD2 As String = ""                    ' Tax Code 2
    Private _TAXCD3 As String = ""                    ' Tax Code 3
    Private _COMNT1 As String = ""                    ' First Comment Line
    Private _COMNT2 As String = ""                    ' Second Comment Line
    Private _COMNT3 As String = ""                    ' Third Comment Line
    Private _SHPLBL As Integer = 0                    ' Percent Ship Labels
    Private _INVCE As String = ""                     ' Invoice Ready Flag
    Private _APPINV As String = ""                    ' Apply to Invoice
    Private _REASON As String = ""                    ' Reason Code
    Private _NAME As String = ""                      ' Customer Name
    Private _ADDR1 As String = ""                     ' Customer Address Line 1
    Private _ADDR2 As String = ""                     ' Customer Address Line 2
    Private _CITY As String = ""                      ' Customer City
    Private _STATE As String = ""                     ' Customer State
    Private _ZIPCD As String = ""                     ' Customer Zip Code
    Private _CNTRY As String = ""                     ' Customer Country
    Private _PHONE As String = ""                     ' Customer Phone
    Private _CNTCT As String = ""                     ' Customer Contact's Name
    Private _TAXPRV As String = ""                    ' Provincial Tax Code
    Private _FEDTAX As String = ""                    ' Federal Tax Number
    Private _TAXABL As String = ""                    ' Is the order Taxable ('Y', 'N')
    Private _EXCRTE As Single = 0                     ' Exchange Rate for the Order
    Private _FIXVAR As String = ""                    ' Fixed/Variable Flag ('F', 'V')
    Private _CURR As String = ""                      ' Currency Code for Order
    Private _RCLDTE As Date = DefaultDate             ' Recall Date
    Private _TTAX As Double = 0                       ' Total Tax owed on Order
    Private _LNETAX As String = ""                    ' Is the order taxed by line
    Private _ADDR3 As String = ""                     ' Customer Address Line 3
    Private _ADDR4 As String = ""                     ' Customer Address Line 4
    Private _ADDR5 As String = ""                     ' Customer Address Line 5
    Private _ADDR6 As String = ""                     ' Customer Address Line 6
    Private _MCOMP As String = ""                     ' Multi Company Flag
    Private _MSITE As String = ""                     ' Multi Site Flag
    Private _UDFKEY As String = ""                    ' User Defined Key
    Private _UDFREF As String = ""                    ' User Defined Reference
    Private _SHPTHRU As String = ""                   ' Ship Thru Code
    Private _XDFINT As Integer = 0                    ' ????
    Private _XDFFLT As Double = 0                     ' ????
    Private _XDFBOL As String = ""                    ' ????
    Private _XDFDTE As Date = DefaultDate             ' ????
    Private _XDFTXT As String = ""                    ' ????
    Private _CREATEDBY As String = ""                 ' Person who created the order
    Private _CREATIONDATE As Date = DefaultDate       ' Date the order was created on
    Private _MODIFIEDBY As String = ""                ' Person who last modified the order
    Private _MODIFICATIONDATE As Date = DefaultDate   ' Date the order was last modified on

    Private _Notes As String = ""                     ' Extra Notes field stored in the SOE table
    Private _AddressLocationNumber As String = ""
    Private _LocationCodeQualifier As String = ""
    Private _RDCDescription As String = ""
    Private _CarrierAlphaCode As String = ""
    Private _CarrierRouting As String = ""
    Private _CarrierTransMethodCode As String = ""

    Private _Carrier As Integer = 0
    Private _CarrierMethod As String = ""
    Private _CarrierThirdParty As String = ""
    Private _CarrierName As String = ""
    Private _ContactName As String = ""
    Private _ContactPhone As String = ""
    Private _ContactEmail As String = ""
    Private _DefaultStockID As String = ""
    Private _OrderType As String = ""
    Private _Finished As Boolean = False

    ' Set up a list of line items on the order
    Private _LineItems As List(Of SalesOrderDetail)

    Public Property AddressLocationNumber() As String
        Get
            Return GetValue(_AddressLocationNumber, "")
        End Get
        Set(ByVal value As String)
            _AddressLocationNumber = GetValue(value, "")
        End Set
    End Property

    Public Property Carrier() As Integer
        Get
            Return _Carrier
        End Get
        Set(ByVal value As Integer)
            _Carrier = value
        End Set
    End Property

    Public Property CarrierAlphaCode() As String
        Get
            Return GetValue(_CarrierAlphaCode, "")
        End Get
        Set(ByVal value As String)
            _CarrierAlphaCode = GetValue(value, "")
        End Set
    End Property

    Public Property CarrierMethod() As String
        Get
            Return GetValue(_CarrierMethod, "")
        End Get
        Set(ByVal value As String)
            _CarrierMethod = GetValue(value, "")
        End Set
    End Property

    Public Property CarrierName() As String
        Get
            Return GetValue(_CarrierName, "")
        End Get
        Set(ByVal value As String)
            _CarrierName = GetValue(value, "")
        End Set
    End Property

    Public Property CarrierRouting() As String
        Get
            Return GetValue(_CarrierRouting, "")
        End Get
        Set(ByVal value As String)
            _CarrierRouting = GetValue(value, "")
        End Set
    End Property

    Public Property CarrierThirdParty() As String
        Get
            Return GetValue(_CarrierThirdParty, "")
        End Get
        Set(ByVal value As String)
            _CarrierThirdParty = GetValue(value, "")
        End Set
    End Property

    Public Property CarrierTransMethodCode() As String
        Get
            Return GetValue(_CarrierTransMethodCode, "")
        End Get
        Set(ByVal value As String)
            _CarrierTransMethodCode = GetValue(value, "")
        End Set
    End Property

    Public Property ContactEmail() As String
        Get
            Return GetValue(_ContactEmail, "")
        End Get
        Set(ByVal value As String)
            _ContactEmail = GetValue(value, "")
        End Set
    End Property

    Public Property ContactName() As String
        Get
            Return GetValue(_ContactName, "")
        End Get
        Set(ByVal value As String)
            _ContactName = GetValue(value, "")
        End Set
    End Property

    Public Property ContactPhone() As String
        Get
            Return GetValue(_ContactPhone, "")
        End Get
        Set(ByVal value As String)
            _ContactPhone = GetValue(value, "")
        End Set
    End Property

    Public Property DefaultStockID() As String
        Get
            Return GetValue(_DefaultStockID, "")
        End Get
        Set(ByVal value As String)
            _DefaultStockID = GetValue(value, "")
        End Set
    End Property

    Public Property Finished() As Boolean
        Get
            Return _Finished
        End Get
        Set(ByVal value As Boolean)
            _Finished = value
        End Set
    End Property

    Public Property LocationCodeQualifier() As String
        Get
            Return GetValue(_LocationCodeQualifier, "")
        End Get
        Set(ByVal value As String)
            _LocationCodeQualifier = GetValue(value, "")
        End Set
    End Property

    Public Property OrderType() As String
        Get
            Return GetValue(_OrderType, "")
        End Get
        Set(ByVal value As String)
            _OrderType = GetValue(value, "")
        End Set
    End Property

    Public Property RDCDescription() As String
        Get
            Return GetValue(_RDCDescription, "")
        End Get
        Set(ByVal value As String)
            _RDCDescription = GetValue(value, "")
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

    Public Property GLXREF() As String
        Get
            Return GetValue(_GLXREF, "")
        End Get
        Set(ByVal value As String)
            _GLXREF = GetValue(value, "")
        End Set
    End Property

    Public Property STYPE() As String
        Get
            Return GetValue(_STYPE, "")
        End Get
        Set(ByVal value As String)
            _STYPE = GetValue(value, "")
        End Set
    End Property

    Public Property STATUS() As String
        Get
            Return GetValue(_STATUS, "")
        End Get
        Set(ByVal value As String)
            _STATUS = GetValue(value, "")
        End Set
    End Property

    Public Property CUSTPO() As String
        Get
            Return GetValue(_CUSTPO, "")
        End Get
        Set(ByVal value As String)
            _CUSTPO = GetValue(value, "")
        End Set
    End Property

    Public Property ORDID() As String
        Get
            Return GetValue(_ORDID, "")
        End Get
        Set(ByVal value As String)
            _ORDID = GetValue(value, "")
        End Set
    End Property

    Public Property ORDDTE() As Date
        Get
            Return GetValue(_ORDDTE, DefaultDate)
        End Get
        Set(ByVal value As Date)
            _ORDDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property SHPCDE() As String
        Get
            Return GetValue(_SHPCDE, "")
        End Get
        Set(ByVal value As String)
            _SHPCDE = GetValue(value, "")
        End Set
    End Property

    Public Property REP1() As String
        Get
            Return GetValue(_REP1, "")
        End Get
        Set(ByVal value As String)
            _REP1 = GetValue(value, "")
        End Set
    End Property

    Public Property SPLIT1() As Single
        Get
            Return GetValue(_SPLIT1, 0)
        End Get
        Set(ByVal value As Single)
            _SPLIT1 = GetValue(value, 0)
        End Set
    End Property

    Public Property REP2() As String
        Get
            Return GetValue(_REP2, "")
        End Get
        Set(ByVal value As String)
            _REP2 = GetValue(value, "")
        End Set
    End Property

    Public Property SPLIT2() As Single
        Get
            Return GetValue(_SPLIT2, 0)
        End Get
        Set(ByVal value As Single)
            _SPLIT2 = GetValue(value, 0)
        End Set
    End Property

    Public Property REP3() As String
        Get
            Return GetValue(_REP3, "")
        End Get
        Set(ByVal value As String)
            _REP3 = GetValue(value, "")
        End Set
    End Property

    Public Property SPLIT3() As Single
        Get
            Return GetValue(_SPLIT3, 0)
        End Get
        Set(ByVal value As Single)
            _SPLIT3 = GetValue(value, 0)
        End Set
    End Property

    Private Property COMMIS() As Single
        Get
            Return GetValue(_COMMIS, 0)
        End Get
        Set(ByVal value As Single)
            _COMMIS = GetValue(value, 0)
        End Set
    End Property

    Public Property TERMS() As String
        Get
            Return GetValue(_TERMS, "")
        End Get
        Set(ByVal value As String)
            _TERMS = GetValue(value, "")
        End Set
    End Property

    Public Property SHPVIA() As String
        Get
            Return GetValue(_SHPVIA, "")
        End Get
        Set(ByVal value As String)
            _SHPVIA = GetValue(value, "")
        End Set
    End Property

    Public Property XURR() As String
        Get
            Return GetValue(_XURR, "")
        End Get
        Set(ByVal value As String)
            _XURR = GetValue(value, "")
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

    Public Property TAXCD1() As String
        Get
            Return GetValue(_TAXCD1, "")
        End Get
        Set(ByVal value As String)
            _TAXCD1 = GetValue(value, "")
        End Set
    End Property

    Public Property TAXCD2() As String
        Get
            Return GetValue(_TAXCD2, "")
        End Get
        Set(ByVal value As String)
            _TAXCD2 = GetValue(value, "")
        End Set
    End Property

    Public Property TAXCD3() As String
        Get
            Return GetValue(_TAXCD3, "")
        End Get
        Set(ByVal value As String)
            _TAXCD3 = GetValue(value, "")
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

    Public Property SHPLBL() As Integer
        Get
            Return GetValue(_SHPLBL, 0)
        End Get
        Set(ByVal value As Integer)
            _SHPLBL = GetValue(value, 0)
        End Set
    End Property

    Public Property INVCE() As String
        Get
            Return GetValue(_INVCE, "")
        End Get
        Set(ByVal value As String)
            _INVCE = GetValue(value, "")
        End Set
    End Property

    Public Property APPINV() As String
        Get
            Return GetValue(_APPINV, "")
        End Get
        Set(ByVal value As String)
            _APPINV = GetValue(value, "")
        End Set
    End Property

    Public Property REASON() As String
        Get
            Return GetValue(_REASON, "")
        End Get
        Set(ByVal value As String)
            _REASON = GetValue(value, "")
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

    Public Property PHONE() As String
        Get
            Return GetValue(_PHONE, "")
        End Get
        Set(ByVal value As String)
            _PHONE = GetValue(value, "")
        End Set
    End Property

    Public Property CNTCT() As String
        Get
            Return GetValue(_CNTCT, "")
        End Get
        Set(ByVal value As String)
            _CNTCT = GetValue(value, "")
        End Set
    End Property

    Public Property TAXPRV() As String
        Get
            Return GetValue(_TAXPRV, "")
        End Get
        Set(ByVal value As String)
            _TAXPRV = GetValue(value, "")
        End Set
    End Property

    Public Property FEDTAX() As String
        Get
            Return GetValue(_FEDTAX, "")
        End Get
        Set(ByVal value As String)
            _FEDTAX = GetValue(value, "")
        End Set
    End Property

    Public Property TAXABL() As String
        Get
            Return GetValue(_TAXABL, "")
        End Get
        Set(ByVal value As String)
            _TAXABL = GetValue(value, "")
        End Set
    End Property

    Public Property EXCRTE() As Single
        Get
            Return GetValue(_EXCRTE, 0)
        End Get
        Set(ByVal value As Single)
            _EXCRTE = GetValue(value, 0)
        End Set
    End Property

    Public Property FIXVAR() As String
        Get
            Return GetValue(_FIXVAR, "")
        End Get
        Set(ByVal value As String)
            _FIXVAR = GetValue(value, "")
        End Set
    End Property

    Public Property CURR() As String
        Get
            Return GetValue(_CURR, "")
        End Get
        Set(ByVal value As String)
            _CURR = GetValue(value, "")
        End Set
    End Property

    Public Property RCLDTE() As Date
        Get
            Return GetValue(_RCLDTE, DefaultDate)
        End Get
        Set(ByVal value As Date)
            _RCLDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property TTAX() As Double
        Get
            Return GetValue(_TTAX, 0)
        End Get
        Set(ByVal value As Double)
            _TTAX = GetValue(value, 0)
        End Set
    End Property

    Public Property LNETAX() As String
        Get
            Return GetValue(_LNETAX, "")
        End Get
        Set(ByVal value As String)
            _LNETAX = GetValue(value, "")
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

    Public Property MCOMP() As String
        Get
            Return GetValue(_MCOMP, "")
        End Get
        Set(ByVal value As String)
            _MCOMP = GetValue(value, "")
        End Set
    End Property

    Public Property MSITE() As String
        Get
            Return GetValue(_MSITE, "")
        End Get
        Set(ByVal value As String)
            _MSITE = GetValue(value, "")
        End Set
    End Property

    Public Property UDFKEY() As String
        Get
            Return GetValue(_UDFKEY, "")
        End Get
        Set(ByVal value As String)
            _UDFKEY = GetValue(value, "")
        End Set
    End Property

    Public Property UDFREF() As String
        Get
            Return GetValue(_UDFREF, "")
        End Get
        Set(ByVal value As String)
            _UDFREF = GetValue(value, "")
        End Set
    End Property

    Public Property SHPTHRU() As String
        Get
            Return GetValue(_SHPTHRU, "")
        End Get
        Set(ByVal value As String)
            _SHPTHRU = GetValue(value, "")
        End Set
    End Property

    Public Property XDFINT() As Integer
        Get
            Return GetValue(_XDFINT, 0)
        End Get
        Set(ByVal value As Integer)
            _XDFINT = GetValue(value, 0)
        End Set
    End Property

    Public Property XDFFLT() As Double
        Get
            Return GetValue(_XDFFLT, 0)
        End Get
        Set(ByVal value As Double)
            _XDFFLT = GetValue(value, 0)
        End Set
    End Property

    Public Property XDFBOL() As String
        Get
            Return GetValue(_XDFBOL, "")
        End Get
        Set(ByVal value As String)
            _XDFBOL = GetValue(value, "")
        End Set
    End Property

    Public Property XDFDTE() As Date
        Get
            Return GetValue(_XDFDTE, DefaultDate)
        End Get
        Set(ByVal value As Date)
            _XDFDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property XDFTXT() As String
        Get
            Return GetValue(_XDFTXT, "")
        End Get
        Set(ByVal value As String)
            _XDFTXT = GetValue(value, "")
        End Set
    End Property

    Public Property CREATEDBY() As String
        Get
            Return GetValue(_CREATEDBY, "")
        End Get
        Set(ByVal value As String)
            _CREATEDBY = GetValue(value, "")
        End Set
    End Property

    Public Property CREATIONDATE() As Date
        Get
            Return GetValue(_CREATIONDATE, DefaultDate)
        End Get
        Set(ByVal value As Date)
            _CREATIONDATE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property MODIFIEDBY() As String
        Get
            Return GetValue(_MODIFIEDBY, "")
        End Get
        Set(ByVal value As String)
            _MODIFIEDBY = GetValue(value, "")
        End Set
    End Property

    Public Property MODIFICATIONDATE() As Date
        Get
            Return GetValue(_MODIFICATIONDATE, DefaultDate)
        End Get
        Set(ByVal value As Date)
            _MODIFICATIONDATE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property LINEITEMS() As List(Of SalesOrderDetail)
        Get
            Return _LineItems
        End Get
        Set(ByVal value As List(Of SalesOrderDetail))
            _LineItems = value
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

    Public Sub New()

    End Sub

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

End Class

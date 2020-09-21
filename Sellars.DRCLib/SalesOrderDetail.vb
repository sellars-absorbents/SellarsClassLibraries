Imports System
Imports System.Collections.Generic
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.ServiceModel

Public Class SalesOrderDetail
    Protected Friend Const DefaultDate As Date = #1/1/2000 12:01:00 AM#

    ' NOTE:  The fields in this section of the class should match exactly the fields contained within the SOEDetItem structure
    '        in the DRCLIB.SalesOrderDetailClass, in the same order for easier comparison; with the exception of the Filler fields
    Private _ORDNUM As String = ""                    ' Order Number
    Private _LINNUM As String = ""                    ' Line Number
    Private _DELNUM As String = ""                    ' Delivery Number
    Private _STATUS As String = ""                    ' Line Status Code ('3' or '4')
    Private _CUSTID As String = ""                    ' Customer ID
    Private _PRTNUM As String = ""                    ' Part Number
    Private _EDILIN As String = ""                    ' ????
    Private _TAXABL As String = ""                    ' Is the line Taxable ('Y' or 'N')
    Private _GLXREF As String = ""                    ' General Ledger Cross Reference
    Private _CURDUE As Date = DefaultDate             ' Current Due Date
    Private _QTLINE As String = ""                    ' ????
    Private _ORGDUE As Date = DefaultDate             ' Original Due Date
    Private _QTDEL As String = ""                     ' ????
    Private _CUSDUE As Date = DefaultDate             ' Customer Due Date
    Private _PROBAB As Integer = 0                    ' ????
    Private _SHPDTE As Date = DefaultDate             ' Ship Date (Btrieve date format)
    Private _SLSUOM As String = ""                    ' Unit of Measure (e.g. 'EA', 'RL', etc.)
    Private _REFRNC As String = ""                    ' Reference
    Private _PRICE As Double = 0                      ' Unit Price
    Private _ORGQTY As Double = 0                     ' Original Order Quantity
    Private _CURQTY As Double = 0                     ' Current Order Quantity
    Private _BCKQTY As Double = 0                     ' Backordered Quantity
    Private _SHPQTY As Double = 0                     ' Already Shipped Quantity
    Private _CURSHP As Double = 0                     ' Current Shipment Quantity
    Private _DUEQTY As Double = 0                     ' Quantity Remaining to Ship
    Private _INVQTY As Double = 0                     ' Quantity Already Invoiced
    Private _DISC As Single = 0                       ' Discount
    Private _STYPE As String = ""                     ' ????
    Private _PRNT As String = ""                      ' ????
    Private _AKPRNT As String = ""                    ' ????
    Private _STK As String = ""                       ' Stock ID - Warehouse quantity is pulled from
    Private _COCFLG As String = ""                    ' ????
    Private _FORCUR As Double = 0                     ' ????
    Private _HSTAT As String = ""                     ' ????
    Private _SLSREP As String = ""                    ' Sales Rep ID
    Private _COMMIS As Single = 0                     ' Commission paid on line
    Private _DRPSHP As String = ""                    ' ????
    Private _QUMQTY As Single = 0                     ' ????
    Private _TAXCDE1 As String = ""                   ' Tax Code 1
    Private _TAX1 As Double = 0                       ' Amount paid for Tax Code 1
    Private _TAXCDE2 As String = ""                   ' Tax Code 2
    Private _TAX2 As Double = 0                       ' Amount Paid for Tax Code 2
    Private _TAXCDE3 As String = ""                   ' Tax Code 3
    Private _TAX3 As Double = 0                       ' Amount Paid for tax Code 3
    Private _MCOMP As String = ""                     ' ????
    Private _MSITE As String = ""                     ' ????
    Private _UDFKEY As String = ""                    ' User Defined Key
    Private _UDFREF As String = ""                    ' User Defined Reference Code
    Private _DEXPFLG As String = ""                   ' ????
    Private _COST As Double = 0                       ' Line Item Cost
    Private _MARKUP As Double = 0                     ' Line Item Markup
    Private _QTORD As String = ""                     ' ????
    Private _XDFINT As Integer = 0                    ' ????
    Private _XDFFLT As Double = 0                     ' ????
    Private _XDFBOL As String = ""                    ' ????
    Private _XDFDTE As Date = DefaultDate             ' ????
    Private _XDFTXT As String = ""                    ' ????
    Private _CREATEDBY As String = ""                 ' Person who created the line item
    Private _CREATIONDATE As Date = DefaultDate       ' Date the line item was created on
    Private _MODIFIEDBY As String = ""                ' Person who last modified the line item
    Private _MODIFICATIONDATE As Date = DefaultDate   ' Date the line item was last modified
    Private _BOKDTE As Date = DefaultDate             ' Backordered date
    Private _DBKDTE As Date = DefaultDate             ' ????
    Private _REVLEV As String = ""                    ' Revision Level

    ' Extra Property values for each line, in case the are needed for other processing
    Private _BuyerPartNumber As String = ""
    Private _Cartons As Integer = 0
    Private _DiscountPercent As Decimal = 0
    Private _Pallets As Integer = 0
    Private _PromotionCode As String = ""
    Private _UnitPrice As Decimal = 0
    Private _ShipFromWarehouse As String = ""
    Private _WarehouseChangeReason As Short = 0

    ' Extra Property for Notes for each detail item
    Private _Notes As New SalesOrderDetailNotes

    Public Property ORDNUM() As String
        Get
            Return _ORDNUM.Trim
        End Get
        Set(ByVal value As String)
            _ORDNUM = GetValue(value, "")
        End Set
    End Property

    Public Property LINNUM() As String
        Get
            Return _LINNUM.Trim
        End Get
        Set(ByVal value As String)
            _LINNUM = GetValue(value, "")
        End Set
    End Property

    Public Property DELNUM() As String
        Get
            Return _DELNUM.Trim
        End Get
        Set(ByVal value As String)
            _DELNUM = GetValue(value, "")
        End Set
    End Property

    Public Property DiscountPercent() As Decimal
        Get
            Return _DiscountPercent
        End Get
        Set(ByVal value As Decimal)
            _DiscountPercent = GetValue(value, 0)
        End Set
    End Property

    Public Property STATUS() As String
        Get
            Return _STATUS.Trim
        End Get
        Set(ByVal value As String)
            _STATUS = GetValue(value, "")
        End Set
    End Property

    Public Property CUSTID() As String
        Get
            Return _CUSTID.Trim
        End Get
        Set(ByVal value As String)
            _CUSTID = GetValue(value, "")
        End Set
    End Property

    Public Property PRTNUM() As String
        Get
            Return _PRTNUM.Trim
        End Get
        Set(ByVal value As String)
            _PRTNUM = GetValue(value, "")
        End Set
    End Property

    Public Property EDILIN() As String
        Get
            Return _EDILIN.Trim
        End Get
        Set(ByVal value As String)
            _EDILIN = GetValue(value, "")
        End Set
    End Property

    Public Property TAXABL() As String
        Get
            Return _TAXABL.Trim
        End Get
        Set(ByVal value As String)
            _TAXABL = GetValue(value, "")
        End Set
    End Property

    Public Property GLXREF() As String
        Get
            Return _GLXREF.Trim
        End Get
        Set(ByVal value As String)
            _GLXREF = GetValue(value, "")
        End Set
    End Property

    Public Property CURDUE() As Date
        Get
            Return _CURDUE
        End Get
        Set(ByVal value As Date)
            _CURDUE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property QTLINE() As String
        Get
            Return _QTLINE.Trim
        End Get
        Set(ByVal value As String)
            _QTLINE = GetValue(value, "")
        End Set
    End Property

    Public Property ORGDUE() As Date
        Get
            Return _ORGDUE
        End Get
        Set(ByVal value As Date)
            _ORGDUE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property QTDEL() As String
        Get
            Return _QTDEL.Trim
        End Get
        Set(ByVal value As String)
            _QTDEL = GetValue(value, "")
        End Set
    End Property

    Public Property CUSDUE() As Date
        Get
            Return _CUSDUE
        End Get
        Set(ByVal value As Date)
            _CUSDUE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property PROBAB() As Integer
        Get
            Return _PROBAB
        End Get
        Set(ByVal value As Integer)
            _PROBAB = GetValue(value, 0)
        End Set
    End Property

    Public Property SHPDTE() As Date
        Get
            Return _SHPDTE
        End Get
        Set(ByVal value As Date)
            _SHPDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property SLSUOM() As String
        Get
            Return _SLSUOM.Trim
        End Get
        Set(ByVal value As String)
            _SLSUOM = GetValue(value, "")
        End Set
    End Property

    Public Property REFRNC() As String
        Get
            Return _REFRNC.Trim
        End Get
        Set(ByVal value As String)
            _REFRNC = GetValue(value, "")
        End Set
    End Property

    Public Property PRICE() As Double
        Get
            Return _PRICE
        End Get
        Set(ByVal value As Double)
            _PRICE = GetValue(value, 0)
        End Set
    End Property

    Public Property ORGQTY() As Double
        Get
            Return _ORGQTY
        End Get
        Set(ByVal value As Double)
            _ORGQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property CURQTY() As Double
        Get
            Return _CURQTY
        End Get
        Set(ByVal value As Double)
            _CURQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property BCKQTY() As Double
        Get
            Return _BCKQTY
        End Get
        Set(ByVal value As Double)
            _BCKQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property SHPQTY() As Double
        Get
            Return _SHPQTY
        End Get
        Set(ByVal value As Double)
            _SHPQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property CURSHP() As Double
        Get
            Return _CURSHP
        End Get
        Set(ByVal value As Double)
            _CURSHP = GetValue(value, 0)
        End Set
    End Property

    Public Property DUEQTY() As Double
        Get
            Return _DUEQTY
        End Get
        Set(ByVal value As Double)
            _DUEQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property INVQTY() As Double
        Get
            Return _INVQTY
        End Get
        Set(ByVal value As Double)
            _INVQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property DISC() As Single
        Get
            Return _DISC
        End Get
        Set(ByVal value As Single)
            _DISC = GetValue(value, 0)
        End Set
    End Property

    Public Property STYPE() As String
        Get
            Return _STYPE.Trim
        End Get
        Set(ByVal value As String)
            _STYPE = GetValue(value, "")
        End Set
    End Property

    Public Property PRNT() As String
        Get
            Return _PRNT.Trim
        End Get
        Set(ByVal value As String)
            _PRNT = GetValue(value, "")
        End Set
    End Property

    Public Property AKPRNT() As String
        Get
            Return _AKPRNT.Trim
        End Get
        Set(ByVal value As String)
            _AKPRNT = GetValue(value, "")
        End Set
    End Property

    Public Property STK() As String
        Get
            Return _STK.Trim
        End Get
        Set(ByVal value As String)
            _STK = GetValue(value, "")
        End Set
    End Property

    Public Property COCFLG() As String
        Get
            Return _COCFLG.Trim
        End Get
        Set(ByVal value As String)
            _COCFLG = GetValue(value, "")
        End Set
    End Property

    Public Property FORCUR() As Double
        Get
            Return _FORCUR
        End Get
        Set(ByVal value As Double)
            _FORCUR = GetValue(value, 0)
        End Set
    End Property

    Public Property HSTAT() As String
        Get
            Return _HSTAT.Trim
        End Get
        Set(ByVal value As String)
            _HSTAT = GetValue(value, "")
        End Set
    End Property

    Public Property SLSREP() As String
        Get
            Return _SLSREP.Trim
        End Get
        Set(ByVal value As String)
            _SLSREP = GetValue(value, "")
        End Set
    End Property

    Public Property COMMIS() As Single
        Get
            Return _COMMIS
        End Get
        Set(ByVal value As Single)
            _COMMIS = GetValue(value, 0)
        End Set
    End Property

    Public Property DRPSHP() As String
        Get
            Return _DRPSHP
        End Get
        Set(ByVal value As String)
            _DRPSHP = GetValue(value, "")
        End Set
    End Property

    Public Property QUMQTY() As Single
        Get
            Return _QUMQTY
        End Get
        Set(ByVal value As Single)
            _QUMQTY = GetValue(value, 0)
        End Set
    End Property

    Public Property TAXCDE1() As String
        Get
            Return _TAXCDE1.Trim
        End Get
        Set(ByVal value As String)
            _TAXCDE1 = GetValue(value, "")
        End Set
    End Property

    Public Property TAX1() As Double
        Get
            Return _TAX1
        End Get
        Set(ByVal value As Double)
            _TAX1 = GetValue(value, 0)
        End Set
    End Property

    Public Property TAXCDE2() As String
        Get
            Return _TAXCDE2.Trim
        End Get
        Set(ByVal value As String)
            _TAXCDE2 = GetValue(value, "")
        End Set
    End Property

    Public Property TAX2() As Double
        Get
            Return _TAX2
        End Get
        Set(ByVal value As Double)
            _TAX2 = GetValue(value, 0)
        End Set
    End Property

    Public Property TAXCDE3() As String
        Get
            Return _TAXCDE3.Trim
        End Get
        Set(ByVal value As String)
            _TAXCDE3 = GetValue(value, "")
        End Set
    End Property

    Public Property TAX3() As Double
        Get
            Return _TAX3
        End Get
        Set(ByVal value As Double)
            _TAX3 = GetValue(value, 0)
        End Set
    End Property

    Public Property MCOMP() As String
        Get
            Return _MCOMP.Trim
        End Get
        Set(ByVal value As String)
            _MCOMP = GetValue(value, "")
        End Set
    End Property

    Public Property MSITE() As String
        Get
            Return _MSITE.Trim
        End Get
        Set(ByVal value As String)
            _MSITE = GetValue(value, "")
        End Set
    End Property

    Public Property UDFKEY() As String
        Get
            Return _UDFKEY.Trim
        End Get
        Set(ByVal value As String)
            _UDFKEY = GetValue(value, "")
        End Set
    End Property

    Public Property UDFREF() As String
        Get
            Return _UDFREF.Trim
        End Get
        Set(ByVal value As String)
            _UDFREF = GetValue(value, "")
        End Set
    End Property

    Public Property DEXPFLG() As String
        Get
            Return _DEXPFLG.Trim
        End Get
        Set(ByVal value As String)
            _DEXPFLG = GetValue(value, "")
        End Set
    End Property

    Public Property COST() As Double
        Get
            Return _COST
        End Get
        Set(ByVal value As Double)
            _COST = GetValue(value, 0)
        End Set
    End Property

    Public Property MARKUP() As Double
        Get
            Return _MARKUP
        End Get
        Set(ByVal value As Double)
            _MARKUP = GetValue(value, 0)
        End Set
    End Property

    Public Property QTORD() As String
        Get
            Return _QTORD.Trim
        End Get
        Set(ByVal value As String)
            _QTORD = GetValue(value, "")
        End Set
    End Property

    Public Property XDFINT() As Integer
        Get
            Return _XDFINT
        End Get
        Set(ByVal value As Integer)
            _XDFINT = GetValue(value, 0)
        End Set
    End Property

    Public Property XDFFLT() As Double
        Get
            Return _XDFFLT
        End Get
        Set(ByVal value As Double)
            _XDFFLT = GetValue(value, 0)
        End Set
    End Property

    Public Property XDFBOL() As String
        Get
            Return _XDFBOL.Trim
        End Get
        Set(ByVal value As String)
            _XDFBOL = GetValue(value, "")
        End Set
    End Property

    Public Property XDFDTE() As Date
        Get
            Return _XDFDTE
        End Get
        Set(ByVal value As Date)
            _XDFDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property XDFTXT() As String
        Get
            Return _XDFTXT.Trim
        End Get
        Set(ByVal value As String)
            _XDFTXT = GetValue(value, "")
        End Set
    End Property

    Public Property CREATEDBY() As String
        Get
            Return _CREATEDBY.Trim
        End Get
        Set(ByVal value As String)
            _CREATEDBY = GetValue(value, "")
        End Set
    End Property

    Public Property CREATIONDATE() As Date
        Get
            Return _CREATIONDATE
        End Get
        Set(ByVal value As Date)
            _CREATIONDATE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property MODIFIEDBY() As String
        Get
            Return _MODIFIEDBY.Trim
        End Get
        Set(ByVal value As String)
            _MODIFIEDBY = GetValue(value, "")
        End Set
    End Property

    Public Property MODIFICATIONDATE() As Date
        Get
            Return _MODIFICATIONDATE
        End Get
        Set(ByVal value As Date)
            _MODIFICATIONDATE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property BOKDTE() As Date
        Get
            Return _BOKDTE
        End Get
        Set(ByVal value As Date)
            _BOKDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property DBKDTE() As Date
        Get
            Return _DBKDTE
        End Get
        Set(ByVal value As Date)
            _DBKDTE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property REVLEV() As String
        Get
            Return _REVLEV.Trim
        End Get
        Set(ByVal value As String)
            _REVLEV = GetValue(value, "")
        End Set
    End Property

    ' *********************************************************************************************
    ' * Here start the extra properties                                                           *
    ' *********************************************************************************************
    Public Property BuyerPartNumber() As String
        Get
            Return _BuyerPartNumber
        End Get
        Set(ByVal value As String)
            _BuyerPartNumber = value
        End Set
    End Property

    Public Property Cartons() As Integer
        Get
            Return _Cartons
        End Get
        Set(ByVal value As Integer)
            _Cartons = GetValue(value, 0)
        End Set
    End Property

    Public Property Notes() As SalesOrderDetailNotes
        Get
            Return _Notes
        End Get
        Set(ByVal value As SalesOrderDetailNotes)
            _Notes = value
        End Set
    End Property

    Public Property Pallets() As Integer
        Get
            Return _Pallets
        End Get
        Set(ByVal value As Integer)
            _Pallets = GetValue(value, 0)
        End Set
    End Property

    Public Property PromotionCode() As String
        Get
            Return _PromotionCode
        End Get
        Set(ByVal value As String)
            _PromotionCode = value
        End Set
    End Property

    Public Property UnitPrice() As Decimal
        Get
            Return _UnitPrice
        End Get
        Set(ByVal value As Decimal)
            _UnitPrice = GetValue(value, 0)
        End Set
    End Property

    Public Property ShipFromWarehouse() As String
        Get
            Return _ShipFromWarehouse
        End Get
        Set(ByVal value As String)
            _ShipFromWarehouse = value
        End Set
    End Property

    Public Property WarehouseChangeReason() As Short
        Get
            Return _WarehouseChangeReason
        End Get
        Set(ByVal value As Short)
            _WarehouseChangeReason = value
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

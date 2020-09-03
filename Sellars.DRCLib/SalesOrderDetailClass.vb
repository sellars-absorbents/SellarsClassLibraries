Imports Sellars.SQL
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks
Imports System.Xml.Linq

Public Class SalesOrderDetailClass
    Inherits ClassBase

    Private _CoreSize As Decimal
    Private _ErrorCode As Integer = 0
    Private _ErrorDescription As String = ""
    Private _LastLine As Integer
    Private _OutsideDiameter As Decimal
    Private _Pallets As Boolean
    Private _Deleteable As Boolean = False
    Private _TotalPrice As Decimal
    Private _SlitWidth As Decimal

    Private _ORDNUM As String = ""
    Private _LINNUM As String = ""
    Private _DELNUM As String = ""
    Private _STATUS As String = ""
    Private _CUSTID As String = ""
    Private _PRTNUM As String = ""
    Private _EDILIN As String = ""
    Private _TAXABL As String = ""
    Private _GLXREF As String = ""
    Private _CURDUE As Date
    Private _FILL01 As String = ""
    Private _ORGDUE As Date
    'Private _FILL02 As String
    Private _PROBAB As Integer
    Private _CUSDUE As Date
    Private _FILL03 As String = ""
    Private _SHPDTE As Date
    Private _FILL04 As String = ""
    Private _SLSUOM As String = ""
    Private _REFRNC As String = ""
    Private _PRICE As Decimal
    Private _ORGQTY As Double
    Private _CURQTY As Double
    Private _BCKQTY As Double
    Private _SHPQTY As Double
    Private _CURSHP As Double
    Private _DUEQTY As Double
    Private _INVQTY As Double
    Private _DISC As Single
    Private _STYPE As String = ""
    Private _PRNT As String = ""
    Private _AKPRNT As String = ""
    Private _STK As String = ""
    Private _COCFLG As String = ""
    Private _FORCUR As Double
    Private _HSTAT As String = ""
    Private _SLSREP As String = ""
    Private _COMMIS As Single
    Private _DRPSHP As String = ""
    Private _QUMQTY As Single
    Private _TAXCDE1 As String = ""
    Private _TAX1 As Double
    Private _TAXCDE2 As String = ""
    Private _TAX2 As Double
    Private _TAXCDE3 As String = ""
    Private _TAX3 As Double
    Private _MCOMP As String = ""
    Private _MSITE As String = ""
    Private _QTDEL As String = ""
    Private _QTLINE As String = ""
    Private _UDFKEY As String = ""
    Private _UDFREF As String = ""
    Private _DEXPFLG As String = ""
    Private _FILLER As String = ""
    Private _OrderPolicy As String = ""

    'new fields in max 5
    Private _COST As Double
    Private _MARKUP As Double
    Private _QTORD As String = ""

    Private _XDFINT As Integer
    Private _XDFFLT As Double
    Private _XDFBOL As String = ""
    Private _XDFDTE As Date
    Private _XDFTXT As String = ""
    Private _CREATEDBY As String = ""
    Private _CREATIONDATE As Date
    Private _MODIFIEDBY As String = ""
    Private _MODIFICATIONDATE As Date
    Private _BOKDTE As Date
    Private _DBKDTE As Date
    Private _REVLEV As String = ""

    Public Enum LineStatus
        Closed = 4
        Open = 3
    End Enum

    Public Property CoreSize() As Decimal
        Get
            Return _CoreSize
        End Get
        Set(ByVal value As Decimal)
            _CoreSize = value
        End Set
    End Property

    Public Property OrderDeleteable() As Boolean
        Get
            Return _Deleteable
        End Get
        Set(ByVal value As Boolean)
            _Deleteable = value
        End Set
    End Property

    Public Property ErrorCode() As Integer
        Get
            Return _ErrorCode
        End Get
        Set(ByVal value As Integer)
            _ErrorCode = value
        End Set
    End Property

    Public Property ErrorDescription() As String
        Get
            Return _ErrorDescription
        End Get
        Set(ByVal value As String)
            _ErrorDescription = value.Trim
        End Set
    End Property

    Public Property LastLine() As Integer
        Get
            Return _LastLine
        End Get
        Set(ByVal value As Integer)
            _LastLine = value
        End Set
    End Property

    Public Property TotalPrice() As Decimal
        Get
            Return _TotalPrice
        End Get
        Set(ByVal value As Decimal)
            _TotalPrice = value
        End Set
    End Property

    Public Property OutsideDiameter() As Decimal
        Get
            Return _OutsideDiameter
        End Get
        Set(ByVal value As Decimal)
            _OutsideDiameter = value
        End Set
    End Property

    Public Property SlitWidth() As Decimal
        Get
            Return _SlitWidth
        End Get
        Set(ByVal value As Decimal)
            _SlitWidth = value
        End Set
    End Property

    Public Property Pallets() As Boolean
        Get
            Return _Pallets
        End Get
        Set(ByVal value As Boolean)
            _Pallets = value
        End Set
    End Property

    Public Property ORDNUM() As String
        Get
            Return _ORDNUM
        End Get
        Set(ByVal value As String)
            _ORDNUM = value.Trim
        End Set
    End Property

    Public Property LINNUM() As String
        Get
            Return _LINNUM
        End Get
        Set(ByVal value As String)
            _LINNUM = value.Trim
        End Set
    End Property

    Public Property DELNUM() As String
        Get
            Return _DELNUM
        End Get
        Set(ByVal value As String)
            _DELNUM = value.Trim
        End Set
    End Property

    Public Property STATUS() As String
        Get
            Return _STATUS
        End Get
        Set(ByVal value As String)
            _STATUS = value.Trim
        End Set
    End Property

    Public Property CUSTID() As String
        Get
            Return _CUSTID
        End Get
        Set(ByVal value As String)
            _CUSTID = value.Trim
        End Set
    End Property

    Public Property PRTNUM() As String
        Get
            Return _PRTNUM
        End Get
        Set(ByVal value As String)
            _PRTNUM = value.Trim
        End Set
    End Property

    Public Property EDILIN() As String
        Get
            Return _EDILIN
        End Get
        Set(ByVal value As String)
            _EDILIN = value.Trim
        End Set
    End Property

    Public Property TAXABL() As String
        Get
            Return _TAXABL
        End Get
        Set(ByVal value As String)
            _TAXABL = value.Trim
        End Set
    End Property

    Public Property CURDUE() As Date
        Get
            Return _CURDUE
        End Get
        Set(ByVal value As Date)
            _CURDUE = value
        End Set
    End Property

    Public Property FILL01() As String
        Get
            Return _FILL01
        End Get
        Set(ByVal value As String)
            _FILL01 = value.Trim
        End Set
    End Property

    Public Property ORGDUE() As Date
        Get
            Return _ORGDUE
        End Get
        Set(ByVal value As Date)
            _ORGDUE = value
        End Set
    End Property

    Public Property PROBAB() As Integer
        Get
            Return _PROBAB
        End Get
        Set(ByVal value As Integer)
            _PROBAB = value
        End Set
    End Property

    Public Property QTDEL() As String
        Get
            Return _QTDEL
        End Get
        Set(ByVal value As String)
            _QTDEL = value.Trim
        End Set
    End Property

    Public Property QTLINE() As String
        Get
            Return _QTLINE
        End Get
        Set(ByVal value As String)
            _QTLINE = value.Trim
        End Set
    End Property

    Public Property CUSDUE() As Date
        Get
            Return _CUSDUE
        End Get
        Set(ByVal value As Date)
            _CUSDUE = value
        End Set
    End Property

    Public Property FILL03() As String
        Get
            Return _FILL03
        End Get
        Set(ByVal value As String)
            _FILL03 = value.Trim
        End Set
    End Property

    Public Property SHPDTE() As Date
        Get
            Return _SHPDTE
        End Get
        Set(ByVal value As Date)
            _SHPDTE = value
        End Set
    End Property

    Public Property FILL04() As String
        Get
            Return _FILL04
        End Get
        Set(ByVal value As String)
            _FILL04 = value.Trim
        End Set
    End Property

    Public Property SLSUOM() As String
        Get
            Return _SLSUOM
        End Get
        Set(ByVal value As String)
            _SLSUOM = value.Trim
        End Set
    End Property

    Public Property REFRNC() As String
        Get
            Return _REFRNC
        End Get
        Set(ByVal value As String)
            _REFRNC = value.Trim
        End Set
    End Property

    Public Property PRICE() As Double
        Get
            Return _PRICE
        End Get
        Set(ByVal value As Double)
            _PRICE = value
        End Set
    End Property

    Public Property ORGQTY() As Double
        Get
            Return _ORGQTY
        End Get
        Set(ByVal value As Double)
            _ORGQTY = value
        End Set
    End Property

    Public Property CURQTY() As Double
        Get
            Return _CURQTY
        End Get
        Set(ByVal value As Double)
            _CURQTY = value
        End Set
    End Property

    Public Property BCKQTY() As Double
        Get
            Return _BCKQTY
        End Get
        Set(ByVal value As Double)
            _BCKQTY = value
        End Set
    End Property

    Public Property SHPQTY() As Double
        Get
            Return _SHPQTY
        End Get
        Set(ByVal value As Double)
            _SHPQTY = value
        End Set
    End Property

    Public Property CURSHP() As Double
        Get
            Return _CURSHP
        End Get
        Set(ByVal value As Double)
            _CURSHP = value
        End Set
    End Property

    Public Property DUEQTY() As Double
        Get
            Return _DUEQTY
        End Get
        Set(ByVal value As Double)
            _DUEQTY = value
        End Set
    End Property

    Public Property INVQTY() As Double
        Get
            Return _INVQTY
        End Get
        Set(ByVal value As Double)
            _INVQTY = value
        End Set
    End Property

    Public Property DISC() As Single
        Get
            Return _DISC
        End Get
        Set(ByVal value As Single)
            _DISC = value
        End Set
    End Property

    Public Property STYPE() As String
        Get
            Return _STYPE
        End Get
        Set(ByVal value As String)
            _STYPE = value.Trim
        End Set
    End Property

    Public Property PRNT() As String
        Get
            Return _PRNT
        End Get
        Set(ByVal value As String)
            _PRNT = value.Trim
        End Set
    End Property

    Public Property AKPRNT() As String
        Get
            Return _AKPRNT
        End Get
        Set(ByVal value As String)
            _AKPRNT = value.Trim
        End Set
    End Property

    Public Property STK() As String
        Get
            Return _STK
        End Get
        Set(ByVal value As String)
            _STK = value.Trim
        End Set
    End Property

    Public Property COCFLG() As String
        Get
            Return _COCFLG
        End Get
        Set(ByVal value As String)
            _COCFLG = value.Trim
        End Set
    End Property

    Public Property FORCUR() As Double
        Get
            Return _FORCUR
        End Get
        Set(ByVal value As Double)
            _FORCUR = value
        End Set
    End Property

    Public Property HSTAT() As String
        Get
            Return _HSTAT
        End Get
        Set(ByVal value As String)
            _HSTAT = value.Trim
        End Set
    End Property

    Public Property SLSREP() As String
        Get
            Return _SLSREP
        End Get
        Set(ByVal value As String)
            _SLSREP = value.Trim
        End Set
    End Property

    Public Property COMMIS() As Single
        Get
            Return _COMMIS
        End Get
        Set(ByVal value As Single)
            _COMMIS = value
        End Set
    End Property

    Public Property DRPSHP() As String
        Get
            Return _DRPSHP
        End Get
        Set(ByVal value As String)
            _DRPSHP = value.Trim
        End Set
    End Property

    Public Property QUMQTY() As Single
        Get
            Return _QUMQTY
        End Get
        Set(ByVal value As Single)
            _QUMQTY = value
        End Set
    End Property

    Public Property TAXCDE1() As String
        Get
            Return _TAXCDE1
        End Get
        Set(ByVal value As String)
            _TAXCDE1 = value.Trim
        End Set
    End Property

    Public Property TAX1() As Double
        Get
            Return _TAX1
        End Get
        Set(ByVal value As Double)
            _TAX1 = value
        End Set
    End Property

    Public Property TAXCDE2() As String
        Get
            Return _TAXCDE2
        End Get
        Set(ByVal value As String)
            _TAXCDE2 = value.Trim
        End Set
    End Property

    Public Property TAX2() As Double
        Get
            Return _TAX2
        End Get
        Set(ByVal value As Double)
            _TAX2 = value
        End Set
    End Property

    Public Property TAXCDE3() As String
        Get
            Return _TAXCDE3
        End Get
        Set(ByVal value As String)
            _TAXCDE3 = value.Trim
        End Set
    End Property

    Public Property TAX3() As Double
        Get
            Return _TAX3
        End Get
        Set(ByVal value As Double)
            _TAX3 = value
        End Set
    End Property

    Public Property MCOMP() As String
        Get
            Return _MCOMP
        End Get
        Set(ByVal value As String)
            _MCOMP = value.Trim
        End Set
    End Property

    Public Property MSITE() As String
        Get
            Return _MSITE
        End Get
        Set(ByVal value As String)
            _MSITE = value.Trim
        End Set
    End Property

    Public Property UDFKEY() As String
        Get
            Return _UDFKEY
        End Get
        Set(ByVal value As String)
            _UDFKEY = value.Trim
        End Set
    End Property

    Public Property UDFREF() As String
        Get
            Return _UDFREF
        End Get
        Set(ByVal value As String)
            _UDFREF = value.Trim
        End Set
    End Property

    Public Property DEXPFLG() As String
        Get
            Return _DEXPFLG
        End Get
        Set(ByVal value As String)
            _DEXPFLG = value.Trim
        End Set
    End Property

    Public Property FILLER() As String
        Get
            Return _FILLER
        End Get
        Set(ByVal value As String)
            _FILLER = value.Trim
        End Set
    End Property

    Public Property GLXREF() As String
        Get
            Return _GLXREF
        End Get
        Set(ByVal value As String)
            _GLXREF = value.Trim
        End Set
    End Property

    Public Property OrderPolicy() As String
        Get
            Return _OrderPolicy
        End Get
        Set(ByVal value As String)
            _OrderPolicy = value.Trim
        End Set
    End Property

    Public Property COST() As Double
        Get
            Return _COST
        End Get
        Set(ByVal value As Double)
            _COST = value
        End Set
    End Property

    Public Property MARKUP() As Double
        Get
            Return _MARKUP
        End Get
        Set(ByVal value As Double)
            _MARKUP = value
        End Set
    End Property

    Public Property QTORD() As String
        Get
            Return _QTORD.Trim
        End Get
        Set(ByVal value As String)
            _QTORD = value.Trim
        End Set
    End Property

    Public Property XDFDTE() As Date
        Get
            Return _XDFDTE
        End Get
        Set(ByVal value As Date)
            _XDFDTE = value
        End Set
    End Property

    Public Property XDFINT() As Integer
        Get
            Return _XDFINT
        End Get
        Set(ByVal value As Integer)
            _XDFINT = value
        End Set
    End Property

    Public Property XDFFLT() As Double
        Get
            Return _XDFFLT
        End Get
        Set(ByVal value As Double)
            _XDFFLT = value
        End Set
    End Property

    Public Property XDFBOL() As String
        Get
            Return _XDFBOL.Trim
        End Get
        Set(ByVal value As String)
            _XDFBOL = value.Trim
        End Set
    End Property

    Public Property XDFTXT() As String
        Get
            Return _XDFTXT.Trim
        End Get
        Set(ByVal value As String)
            _XDFTXT = value.Trim
        End Set
    End Property

    Public Property CREATEDBY() As String
        Get
            Return _CREATEDBY.Trim
        End Get
        Set(ByVal value As String)
            _CREATEDBY = value.Trim
        End Set
    End Property

    Public Property CREATIONDATE() As Date
        Get
            Return _CREATIONDATE
        End Get
        Set(ByVal value As Date)
            _CREATIONDATE = value
        End Set
    End Property

    Public Property MODIFIEDBY() As String
        Get
            Return _MODIFIEDBY.Trim
        End Get
        Set(ByVal value As String)
            _MODIFIEDBY = value.Trim
        End Set
    End Property

    Public Property MODIFICATIONDATE() As Date
        Get
            Return _MODIFICATIONDATE
        End Get
        Set(ByVal value As Date)
            _MODIFICATIONDATE = value
        End Set
    End Property

    Public Property BOKDTE() As Date
        Get
            Return _BOKDTE
        End Get
        Set(ByVal value As Date)
            _BOKDTE = value
        End Set
    End Property

    Public Property DBKDTE() As Date
        Get
            Return _DBKDTE
        End Get
        Set(ByVal value As Date)
            _DBKDTE = value
        End Set
    End Property

    Public Property REVLEV() As String
        Get
            Return _REVLEV
        End Get
        Set(ByVal value As String)
            _REVLEV = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal passOrder As String)
        '_OrderNumber = passOrder
        'Read(passOrder)
    End Sub

    Public Function AnyOpen(ByVal Order As String) As Boolean
        Dim Open As Boolean = False

        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()

            Dim strSQL As String = "Select LINNUM_28 " &
                               "From ""SO_Detail"" " &
                               "Where ORDNUM_28 = '" & Order.Trim & "' " &
                               "and STATUS_28 = '3'"

            Using cmd As New SqlCommand(strSQL, connection)
                Using aoReader As SqlDataReader = cmd.ExecuteReader()
                    If aoReader.Read() Then
                        Open = True
                    End If
                End Using
            End Using
        End Using

        Return Open
    End Function

    Public Sub Add(ByVal value As DataSource, ByVal Order As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal PRTNUM As String, ByVal Quantity As Decimal, ByVal Price As Decimal, ByVal CusDue As Date, ByVal CurDue As Date, ByVal Pallets As Boolean, ByVal CustID As String, ByVal Type As String, ByVal Taxabl As String, ByVal TaxCode1 As String, ByVal TaxRate1 As Decimal, ByVal TaxCode2 As String, ByVal TaxRate2 As Decimal, ByVal TaxCode3 As String, ByVal TaxRate3 As Decimal, ByVal GLCode As String, ByVal SLSUOM As String, ByVal StockCode As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        AddMax(Order, LINNUM, DELNUM, PRTNUM, Quantity, Price, CusDue, CurDue, Pallets, CustID, Type, Taxabl, TaxCode1, TaxRate1, TaxCode2, TaxRate2, TaxCode3, TaxRate3, GLCode, SLSUOM, StockCode)
    End Sub

    Public Sub Add(ByVal value As DataSource)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        AddMax()
    End Sub

    Private Sub AddMax(ByVal Order As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal PRTNUM As String, ByVal Quantity As Decimal, ByVal Price As Decimal, ByVal CusDue As Date, ByVal CurDue As Date, ByVal Pallets As Boolean, ByVal CustID As String, ByVal Type As String, ByVal Taxable As String, ByVal TaxCode1 As String, ByVal TaxRate1 As Decimal, ByVal TaxCode2 As String, ByVal TaxRate2 As Decimal, ByVal TaxCode3 As String, ByVal TaxRate3 As Decimal, ByVal GLCode As String, ByVal SLSUOM As String, ByVal StockCode As String)
        Dim lineNumber As String = LINNUM.PadLeft(2, "0")
        Dim delNumber As String = DELNUM.PadLeft(2, "0")
        Dim status As String = Convert.ToString(LineStatus.Open)
        Dim partNumber As String = PRTNUM.ToUpper.Trim()
        Dim currentQuantity As Double = Convert.ToDouble(Quantity)
        Dim dueQuantity As Double = Convert.ToDouble(Quantity)
        Dim invoiceQuantity As Double = 0
        Dim shipQuantity As Double = 0
        Dim currentShip As Double = 0
        Dim originalQuantity As Double = Convert.ToDouble(Quantity)
        Dim salesOrderDetailXML As XDocument =
            <?xml version="1.0" encoding="utf-8"?>
            <eMAXExact>
                <SO_Detail_Table>
                    <SO_Detail>
                        <ORDNUM_28><%= Order %></ORDNUM_28>
                        <LINNUM_28><%= lineNumber %></LINNUM_28>
                        <DELNUM_28><%= delNumber %></DELNUM_28>
                        <STATUS_28><%= status %></STATUS_28>
                        <CUSTID_28><%= CustID %></CUSTID_28>
                        <PRTNUM_28><%= partNumber %></PRTNUM_28>
                        <EDILIN_28></EDILIN_28>
                        <TAXABL_28><%= Taxable %></TAXABL_28>
                        <GLXREF_28><%= GLCode.ToUpper %></GLXREF_28>
                        <CURDUE_28><%= MakeDate(CusDue) %></CURDUE_28>
                        <QTLINE_28></QTLINE_28>
                        <ORGDUE_28><%= MakeDate(CurDue) %></ORGDUE_28>
                        <QTDEL_28></QTDEL_28>
                        <CUSDUE_28><%= MakeDate(CurDue) %></CUSDUE_28>
                        <PROBAB_28><%= 0 %></PROBAB_28>
                        <SHPDTE_28/>
                        <FILL04_28></FILL04_28>
                        <SLSUOM_28><%= SLSUOM %></SLSUOM_28>
                        <REFRNC_28></REFRNC_28>
                        <PRICE_28><%= Convert.ToDouble(Price) %></PRICE_28>
                        <ORGQTY_28><%= originalQuantity %></ORGQTY_28>
                        <CURQTY_28><%= currentQuantity %></CURQTY_28>
                        <BCKQTY_28></BCKQTY_28>
                        <SHPQTY_28><%= shipQuantity %></SHPQTY_28>
                        <CURSHP_28><%= currentShip %></CURSHP_28>
                        <DUEQTY_28><%= dueQuantity %></DUEQTY_28>
                        <INVQTY_28><%= invoiceQuantity %></INVQTY_28>
                        <DISC_28></DISC_28>
                        <STYPE_28><%= Type %></STYPE_28>
                        <PRNT_28><%= "N" %></PRNT_28>
                        <AKPRNT_28><%= "N" %></AKPRNT_28>
                        <STK_28><%= StockCode.ToUpper %></STK_28>
                        <COCFLG_28></COCFLG_28>
                        <FORCUR_28><%= Convert.ToDouble(Price) %></FORCUR_28>
                        <HSTAT_28><%= "R" %></HSTAT_28>
                        <SLSREP_28></SLSREP_28>
                        <COMMIS_28></COMMIS_28>
                        <DRPSHP_28></DRPSHP_28>
                        <QUMQTY_28></QUMQTY_28>
                        <TAXCDE1_28><%= TaxCode1 %></TAXCDE1_28>
                        <TAX1_28><%= IIf(Taxable = "Y", Math.Round(TaxRate1 * (Convert.ToDouble(Price) * Convert.ToDouble(Quantity)), 2), 0) %></TAX1_28>
                        <TAXCDE2_28><%= TaxCode2 %></TAXCDE2_28>
                        <TAX2_28><%= IIf(Taxable = "Y", Math.Round(TaxRate2 * (Convert.ToDouble(Price) * Convert.ToDouble(Quantity)), 2), 0) %></TAX2_28>
                        <TAXCDE3_28><%= TaxCode3 %></TAXCDE3_28>
                        <TAX3_28><%= IIf(Taxable = "Y", Math.Round(TaxRate3 * (Convert.ToDouble(Price) * Convert.ToDouble(Quantity)), 2), 0) %></TAX3_28>
                        <MCOMP_28></MCOMP_28>
                        <MSITE_28></MSITE_28>
                        <UDFKEY_28></UDFKEY_28>
                        <UDFREF_28></UDFREF_28>
                        <DEXPFLG_28><%= "N" %></DEXPFLG_28>
                        <COST_28></COST_28>
                        <MARKUP_28></MARKUP_28>
                        <QTORD_28></QTORD_28>
                        <XDFINT_28></XDFINT_28>
                        <XDFFLT_28></XDFFLT_28>
                        <XDFBOL_28></XDFBOL_28>
                        <XDFDTE_28></XDFDTE_28>
                        <XDFTXT_28></XDFTXT_28>
                        <FILLER_28></FILLER_28>
                        <CreatedBy></CreatedBy>
                        <CreationDate><%= MakeDate(Date.Now) %></CreationDate>
                        <ModifiedBy></ModifiedBy>
                        <ModificationDate><%= MakeDate(Date.Now) %></ModificationDate>
                        <BOKDTE_28></BOKDTE_28>
                        <DBKDTE_28></DBKDTE_28>
                        <REVLEV_28></REVLEV_28>
                        <MANPRC_28></MANPRC_28>
                        <ORGPRC_28></ORGPRC_28>
                        <PRCALC_28></PRCALC_28>
                        <CLASS_28></CLASS_28>
                        <WARRES_28></WARRES_28>
                        <JOB_28></JOB_28>
                        <CSENDDTE_28></CSENDDTE_28>
                        <CONSGND_28></CONSGND_28>
                        <CURCONSGND_28></CURCONSGND_28>
                        <CONSIGNSTK_28></CONSIGNSTK_28>
                    </SO_Detail>
                </SO_Detail_Table>
            </eMAXExact>

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Short = AddSalesOrderLineItemXML(salesOrderDetailXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error creating sales order detail.")
            Case Else
                ' add the record to the Sellars SQL sales order detail table
                AddSellars(Order, lineNumber, delNumber, status, CustID, partNumber, CurDue, currentQuantity, dueQuantity, invoiceQuantity, shipQuantity, currentShip, originalQuantity)
        End Select
    End Sub

    Private Sub AddMax()
        Dim salesOrderDetailXML As XDocument = FillStructure()

        salesOrderDetailXML.Descendants("STATUS_28").First().Value = Convert.ToString(LineStatus.Open)
        salesOrderDetailXML.Descendants("ORGDUE_28").First().Value = MakeDate(_CURDUE)

        If _TAXABL = "Y" Then
            Dim forcur As Double = Double.Parse(salesOrderDetailXML.Descendants("FORCUR_28").First().Value)
            Dim curqty As Double = Double.Parse(salesOrderDetailXML.Descendants("CURQTY_28").First().Value)

            salesOrderDetailXML.Descendants("TAX1_28").First().Value = Math.Round(_TAX1 * (forcur * curqty), 2)
            salesOrderDetailXML.Descendants("TAX2_28").First().Value = Math.Round(_TAX2 * (forcur * curqty), 2)
            salesOrderDetailXML.Descendants("TAX3_28").First().Value = Math.Round(_TAX3 * (forcur * curqty), 2)
        Else
            salesOrderDetailXML.Descendants("TAX1_28").First().Value = 0
            salesOrderDetailXML.Descendants("TAX2_28").First().Value = 0
            salesOrderDetailXML.Descendants("TAX3_28").First().Value = 0
        End If

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Short = AddSalesOrderLineItemXML(salesOrderDetailXML.ToString())

        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error creating sales order detail.")
            Case Else
                ' add the record to the Sellars SQL sales order detail table
                AddSellars(salesOrderDetailXML.Descendants("ORDNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("LINNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("DELNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("STATUS_28").First().Value,
                           salesOrderDetailXML.Descendants("CUSTID_28").First().Value,
                           salesOrderDetailXML.Descendants("PRTNUM_28").First().Value,
                           DateTime.Parse(salesOrderDetailXML.Descendants("CURDUE_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("CURQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("DUEQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("INVQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("SHPQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("CURSHP_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("ORGQTY_28").First().Value))
        End Select
    End Sub

    ' The following routine will add a sales order to the sellars SQL sales order table
    Private Sub AddSellars(ByVal SalesOrder As String, ByVal Line As String, ByVal Delivery As String, ByVal Status As String, ByVal CustomerID As String, ByVal Part As String, ByVal DueDate As Date, ByVal CurrentQuantity As Decimal, ByVal DueQuantity As Decimal, ByVal InvoicedQuantity As Decimal, ByVal ShippedQuantity As Decimal, ByVal CurrentShippedQuantity As Decimal, ByVal OriginalQuantity As Decimal)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, SalesOrder, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, Line, ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, Delivery, ParameterDirection.Input)
        oSQL.AddParameter("@STATUS", SqlDbType.NVarChar, 1, Status, ParameterDirection.Input)
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CustomerID, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, Part, ParameterDirection.Input)
        oSQL.AddParameter("@CURDUE", SqlDbType.DateTime, 0, DueDate, ParameterDirection.Input)
        oSQL.AddParameter("@CURQTY", SqlDbType.Float, 0, CurrentQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@DUEQTY", SqlDbType.Float, 0, DueQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@INVQTY", SqlDbType.Float, 0, InvoicedQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@SHPQTY", SqlDbType.Float, 0, ShippedQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@CURSHP", SqlDbType.Float, 0, CurrentShippedQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@ORGQTY", SqlDbType.Float, 0, OriginalQuantity, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSalesOrderDetail")
    End Sub

    ' Function used to update the line item status
    Public Sub Clone(ByVal origORDNUM As String, ByVal newORDNUM As String, ByVal SelectedPart As String, ByVal ScheduleDate As Date)
        Dim Status As String = LineStatus.Open
        Dim lineItems As List(Of String) = ReadAll(origORDNUM, SelectedPart)
        Dim newLineNumber As Integer = 0
        Dim salesOrderDetailXML As XDocument

        For Each itm As String In lineItems
            Dim linedel As String() = itm.Split(New [Char]() {"-"c})

            ' Read the requested sales order detail record
            _ErrorDescription = "Line Item Read Max"
            Read(origORDNUM, linedel(0), linedel(1))

            ' Fill the old and new structures with data from the record
            salesOrderDetailXML = FillStructure()

            ' Increment the new line number
            newLineNumber = newLineNumber + 1

            ' Fill in the change functions new data structure by first setting the 
            ' data to the original data record, and then modifying the fields that have changed
            salesOrderDetailXML.Descendants("ORDNUM_28").First().Value = newORDNUM
            salesOrderDetailXML.Descendants("LINNUM_28").First().Value = newLineNumber.ToString.PadLeft(2, "0")
            salesOrderDetailXML.Descendants("CURDUE_28").First().Value = MakeDate(ScheduleDate)
            salesOrderDetailXML.Descendants("CUSDUE_28").First().Value = MakeDate(ScheduleDate)
            salesOrderDetailXML.Descendants("ORGDUE_28").First().Value = MakeDate(ScheduleDate)
            salesOrderDetailXML.Descendants("SHPQTY_28").First().Value = 0
            salesOrderDetailXML.Descendants("INVQTY_28").First().Value = 0
            salesOrderDetailXML.Descendants("BCKQTY_28").First().Value = 0
            salesOrderDetailXML.Descendants("DUEQTY_28").First().Value = salesOrderDetailXML.Descendants("CURQTY_28").First().Value
            salesOrderDetailXML.Descendants("CURSHP_28").First().Value = 0
            salesOrderDetailXML.Descendants("STATUS_28").First().Value = Convert.ToString(LineStatus.Open)
            salesOrderDetailXML.Descendants("XDFDTE_28").First().Value = MakeDate(DefaultDate)
            salesOrderDetailXML.Descendants("ModificationDate").First().Value = MakeDate(Now())
            salesOrderDetailXML.Descendants("CreationDate").First().Value = MakeDate(Now())
            salesOrderDetailXML.Descendants("BOKDTE_28").First().Value = MakeDate(DefaultDate)
            salesOrderDetailXML.Descendants("DBKDTE_28").First().Value = MakeDate(DefaultDate)
            salesOrderDetailXML.Descendants("SHPDTE_28").Remove()

            'Add Sales Order Detail record via MAX Update
            Dim retValue As Short = AddSalesOrderLineItemXML(salesOrderDetailXML.ToString())

            Select Case retValue
                Case 0
                    Throw New System.Exception("Error creating sales order detail.")
                Case Else
                    ' add the record to the Sellars SQL sales order detail table
                    AddSellars(salesOrderDetailXML.Descendants("ORDNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("LINNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("DELNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("STATUS_28").First().Value,
                           salesOrderDetailXML.Descendants("CUSTID_28").First().Value,
                           salesOrderDetailXML.Descendants("PRTNUM_28").First().Value,
                           DateTime.Parse(salesOrderDetailXML.Descendants("CURDUE_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("CURQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("DUEQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("INVQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("SHPQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("CURSHP_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("ORGQTY_28").First().Value))
            End Select
        Next
    End Sub

    ' Function used to update the line item status
    Public Sub CloneRecord(ByVal pORDNUM As String, ByVal pFromLINNUM As String, pToLINNUM As String, ByVal pDELNUM As String, ByVal pClosed As Boolean)
        Dim Status As String

        If pClosed Then
            Status = LineStatus.Closed
        Else
            Status = LineStatus.Open
        End If

        ' Read the requested sales order detail record
        _ErrorDescription = "Line Item Read Max"
        Read(pORDNUM, pFromLINNUM, pDELNUM)

        ' Fill the old and new structures with data from the record
        Dim salesOrderDetailXML As XDocument = FillStructure()

        ' Fill in the change functions new data structure by first setting the 
        ' data to the original data record, and then modifying the fields that have changed
        salesOrderDetailXML.Descendants("STATUS_28").First().Value = Status
        salesOrderDetailXML.Descendants("LINNUM_28").First().Value = pToLINNUM
        salesOrderDetailXML.Descendants("PRICE_28").First().Value = 0
        salesOrderDetailXML.Descendants("FORCUR_28").First().Value = 0

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Short = AddSalesOrderLineItemXML(salesOrderDetailXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error creating sales order detail.")
            Case Else
                ' add the record to the Sellars SQL sales order detail table
                AddSellars(salesOrderDetailXML.Descendants("ORDNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("LINNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("DELNUM_28").First().Value,
                           salesOrderDetailXML.Descendants("STATUS_28").First().Value,
                           salesOrderDetailXML.Descendants("CUSTID_28").First().Value,
                           salesOrderDetailXML.Descendants("PRTNUM_28").First().Value,
                           DateTime.Parse(salesOrderDetailXML.Descendants("CURDUE_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("CURQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("DUEQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("INVQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("SHPQTY_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("CURSHP_28").First().Value),
                           Decimal.Parse(salesOrderDetailXML.Descendants("ORGQTY_28").First().Value))
        End Select
    End Sub

    Public Sub Close(ByVal value As DataSource, ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pClosed As Boolean)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        Update(pORDNUM, pLINNUM, pDELNUM, pClosed)
    End Sub

    ' Function used to close all line items of an order
    Public Sub CloseAll(ByVal value As DataSource, ByVal pORDNUM As String)
        ' CloseAll is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()

            Dim strSQL As String = "Select ORDNUM_28, LINNUM_28, DELNUM_28 " &
                       "From ""SO_Detail"" " &
                       "Where ORDNUM_28 = '" & pORDNUM & "' " &
                       "and STATUS_28 = '" & LineStatus.Open & "'"

            Using cmd As New SqlCommand(strSQL, connection)
                ' We need to ensure the data connection is closed before the assignments because of the call
                ' to the ReadCustomerPartData function
                Using myReader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    ' While there are open line items, call the update routine to close each one
                    While myReader.Read()
                        Update(myReader("ORDNUM_28"), myReader("LINNUM_28"), myReader("DELNUM_28"), True)
                    End While
                End Using
            End Using
        End Using
    End Sub

    Public Sub Delete(ByVal pOrder As String, ByVal pLINNUM As String, ByVal pDELNUM As String)
        'delete the Sales Order Detail record via MAX Update
        Dim retValue As Integer = DeleteSalesOrderLineItemXML(pOrder & pLINNUM & pDELNUM)

        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error deleting sales order detail.")
            Case Else
                ' delete the record from the Sellars SQL sales order detail, and so detail extenstion table
                ' Declare the SQL data layer class
                Dim oSQL As New SqlService(ConnectionString)

                ' Add the parameters to the command object
                oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, pOrder, ParameterDirection.Input)
                oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, pLINNUM, ParameterDirection.Input)
                oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, pDELNUM, ParameterDirection.Input)

                ' Run the stored procedure
                oSQL.RunProc("DeleteSalesOrderDetail")

                Dim SODE As New SODetailExtClass
                SODE.Delete(pOrder, pLINNUM, pDELNUM)
                SODE = Nothing
        End Select
    End Sub

    Public Function TestRead(ByVal passorder As String, Optional ByVal ConversionID As Integer = 1) As DataSet
        _TotalPrice = 0
        _Deleteable = True

        Dim myDataTable As DataTable = New DataTable("SODetail")
        ' Declare variables for DataColumn and DataRow objects. 
        Dim myDataColumn As DataColumn
        Dim myDataRow As DataRow

        ' Add the linenumber column
        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "LINNUM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "DELNUM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PRTNUM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "ORGQTY"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "CURQTY"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "DUEQTY"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "SHPQTY"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "STATUS"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "SHPDTE"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "CURDUE"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "CUSDUE"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PartDescription1"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PartDescription2"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "UOM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "UnitPrice"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "RecommendedPrice"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PrtUnitPrice"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "CustomerPart"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "CustomerPartDescription1"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "CustomerPartDescription2"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "CustomerPartDescription3"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "CustomerPartDescription4"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "SlitWidth"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "CoreSize"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "OD"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "Pallets"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "GLCode"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "COMNT"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PLANID"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "LINDEL"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "DUESHIP"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PARTEXT"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "BothDesc"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "UOMPALLETS"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "EXTPRICES"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "REQSHP"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "ShortComment"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Boolean")
        myDataColumn.ColumnName = "Closed"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "Unisource"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Boolean")
        myDataColumn.ColumnName = "UnisourceItem"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Int32")
        myDataColumn.ColumnName = "UnisourceStatus"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "SentToUnisource"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "STK"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "Discount"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PrtDiscount"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PromotionCode"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "TransferNumber"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Int16")
        myDataColumn.ColumnName = "TransferStatus"
        myDataColumn.ReadOnly = False
        myDataColumn.DefaultValue = 0
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "TransferSubmitted"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "TransferShipped"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "TransferEstimatedDelivery"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Int32")
        myDataColumn.ColumnName = "TransferQuantity"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "DefaultSTK"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Boolean")
        myDataColumn.ColumnName = "AutoTransferItem"
        myDataColumn.ReadOnly = False
        myDataColumn.DefaultValue = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Boolean")
        myDataColumn.ColumnName = "PromotionAddedLine"
        myDataColumn.ReadOnly = False
        myDataColumn.DefaultValue = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Boolean")
        myDataColumn.ColumnName = "QuoteIssue"
        myDataColumn.ReadOnly = False
        myDataColumn.DefaultValue = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "ShipFromWarehouse"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "STAENG"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "AcknowledgedOn"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "CaseHeight"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "CaseLength"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "CaseWidth"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "CaseWeight"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        ' Add the new DataTable to the DataSet. 
        Dim myDataSet As New DataSet
        myDataSet.Tables.Add(myDataTable)

        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()
            ' Define a variable for the database name so the application automatically sets up the SQL string below with the correct table names for test versus production
            Dim strSQL As String = "Select CUSTID_28 as Customer, CUSTYP_23 as CustomerType, LINNUM_28 as LINNUM, DELNUM_28 as DELNUM, PRTNUM_28 as PRTNUM, ORGQTY_28 as ORGQTY, CURQTY_28 as CURQTY, DUEQTY_28 as DUEQTY, SHPQTY_28 as SHPQTY, STATUS_28 as STATUS, SHPDTE_28 as SHPDTE, CURDUE_28 as CURDUE, CUSDUE_28 as CUSDUE, PRICE_28 as UnitPrice, DISC_28 as Discount, GLXREF_28 as GlCode, STK_28 as STK, isnull(STK_29, '') as DefaultSTK, isnull(SLSCNV_29, 1) PartSalesConversion, QuoteIssue, ShipFromWarehouse, AcknowledgedOn, isnull(CaseLength, 0) CaseLength, isnull(CaseHeight, 0) CaseHeight, isnull(CaseWidth, 0) CaseWidth, isnull(GrossWeight, 0) GrossWeight " &
                       "From SO_Detail sod join CUSTOMER_MASTER cm on CUSTID_28 = CUSTID_23 " &
                       "join ShopfloorControl..SalesOrderDetailExt sodx on sodx.ORDNUM = sod.ORDNUM_28 and sodx.DELNUM = DELNUM_28 and sodx.LINNUM = LINNUM_28 " &
                       "left outer join Part_Sales on PRTNUM_29 = PRTNUM_28 " &
                       "left outer join ShopfloorControl..PartMasterCharacteristics pmc on pmc.PRTNUM = PRTNUM_28 " &
                       "Where ORDNUM_28 = '" & passorder & "' " &
                       "Order by LINNUM_28, DELNUM_28"

            Using cmd As New SqlCommand(strSQL, connection)
                Using myReader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    While myReader.Read()
                        ' Get the part information
                        Dim prt As New PartClass(myReader("PRTNUM").Trim, PartClass.DataSource.Max)

                        ' Set the overall information that is necessary before the tasks can be launched
                        Dim myRow As DataRow
                        myRow = myDataSet.Tables(0).NewRow()
                        myRow("LINNUM") = myReader("LINNUM")
                        _LastLine = myReader("LINNUM")
                        myRow("DELNUM") = myReader("DELNUM")
                        myRow("DefaultSTK") = myReader("DefaultSTK")

                        Dim thisLineDel As String = Trim(myReader("LINNUM")) & "-" & Trim(myReader("DELNUM"))
                        myRow("LINDEL") = thisLineDel
                        Dim PartNumber As String = myReader("PRTNUM").Trim
                        myRow("PRTNUM") = myReader("PRTNUM").Trim
                        myRow("ORGQTY") = ConvertFromLBS(PartNumber, ConversionID, prt.PlanCode, myReader("ORGQTY"))
                        myRow("CURQTY") = ConvertFromLBS(PartNumber, ConversionID, prt.PlanCode, myReader("CURQTY"))
                        myRow("DUEQTY") = ConvertFromLBS(PartNumber, ConversionID, prt.PlanCode, myReader("DUEQTY"))
                        myRow("SHPQTY") = ConvertFromLBS(PartNumber, ConversionID, prt.PlanCode, myReader("SHPQTY"))
                        If myRow("SHPQTY") > 0 Then
                            _Deleteable = False
                        End If
                        myRow("QuoteIssue") = myReader("QuoteIssue")
                        myRow("STAENG") = prt.STAENG
                        myRow("AcknowledgedOn") = myReader("AcknowledgedOn")

                        ' Set the case dimensions
                        myRow("CaseHeight") = myReader("CaseHeight")
                        myRow("CaseLength") = myReader("CaseLength")
                        myRow("CaseWidth") = myReader("CaseWidth")
                        myRow("CaseWeight") = myReader("GrossWeight")

                        ' Start a task to get the base information
                        Dim baseInfoTask As Task = Task.Factory.StartNew(Sub()
                                                                             ' If this is a roll goods product, then use 2 decimal points to display the quantities
                                                                             ' otherwise use none
                                                                             Dim printDueQuantity As String = ""
                                                                             Dim printShpQuantity As String = ""
                                                                             If prt.PlanCode = "220" Then
                                                                                 printDueQuantity = FormatNumber(myRow("DUEQTY"), 2)
                                                                                 printShpQuantity = FormatNumber(myRow("SHPQTY"), 2)
                                                                             Else
                                                                                 printDueQuantity = FormatNumber(myRow("DUEQTY"), 0)
                                                                                 printShpQuantity = FormatNumber(myRow("SHPQTY"), 0)
                                                                             End If

                                                                             myRow("DUESHIP") = printDueQuantity & vbCrLf & printShpQuantity

                                                                             myRow("STATUS") = myReader("STATUS").Trim
                                                                             If IsDate(myReader("SHPDTE")) Then
                                                                                 myRow("SHPDTE") = myReader("SHPDTE")
                                                                             Else
                                                                                 myRow("SHPDTE") = Me.DefaultShipDate
                                                                             End If
                                                                             myRow("CURDUE") = myReader("CURDUE")
                                                                             myRow("CUSDUE") = myReader("CUSDUE")

                                                                             ' if the Ship Date is the same as the default ship date, then only display the current due
                                                                             ' date to not confuse operators
                                                                             Dim DefaultShipDate As New Date(2099, 1, 1)
                                                                             If DateDiff(DateInterval.Day, myRow("SHPDTE"), DefaultShipDate) = 0 Then
                                                                                 myRow("REQSHP") = Format(myReader("CURDUE"), "MM/dd/yy")
                                                                             Else
                                                                                 myRow("REQSHP") = Format(myReader("CURDUE"), "MM/dd/yy") & vbCrLf & Format(myRow("SHPDTE"), "MM/dd/yy")
                                                                             End If
                                                                         End Sub)


                        ' Get the slitwidth, coresize, od, pallets, price and promotion information from the SalesOrderDetailExt SQL table 
                        Dim ExtClass As New SODetailExtClass(passorder, myReader("LINNUM"), myReader("DELNUM"))
                        myRow("PromotionAddedLine") = ExtClass.PromotionAddedLine
                        myRow("ShipFromWarehouse") = ExtClass.ShipFromWarehouse

                        ' Start a task to get the price and unit information
                        Dim priceInfoTask As Task = Task.Factory.StartNew(Sub()
                                                                              ' Declare an instance of the Number class 
                                                                              Dim clsNumber As New NumberClass

                                                                              myRow("UnitPrice") = ConvertFromLBSPrice(ConversionID, prt.PlanCode, ExtClass.UnitPrice)

                                                                              ' Figure out how many significant decimal digits there are in the unit price,
                                                                              ' and if less than 2 always show at least two decimals
                                                                              Dim iDigits As Integer = clsNumber.SignificantDigits(myRow("UnitPrice"))
                                                                              If iDigits < 2 Then
                                                                                  iDigits = 2
                                                                              End If
                                                                              myRow("PrtUnitPrice") = FormatCurrency(myRow("UnitPrice"), iDigits)

                                                                              ' Get the line item discounts
                                                                              myRow("Discount") = Math.Round(ExtClass.DiscountPercent, 2, MidpointRounding.AwayFromZero)
                                                                              iDigits = clsNumber.SignificantDigits(myRow("Discount"))
                                                                              If iDigits < 2 Then
                                                                                  iDigits = 2
                                                                              End If
                                                                              myRow("PrtDiscount") = FormatCurrency(myRow("Discount"), iDigits)

                                                                              ' Combine the unit price and the discount applied together
                                                                              myRow("PrtUnitPrice") = myRow("PrtUnitPrice") & vbCrLf & myRow("PrtDiscount")

                                                                              ' Get the recommended price
                                                                              myRow("RecommendedPrice") = RecommendedPrice(myReader("CustomerType"), myReader("PRTNUM"), myRow("CURQTY"))

                                                                              myRow("GLCode") = myReader("GlCode")

                                                                              _TotalPrice = _TotalPrice + (myRow("CURQTY") * (myRow("UnitPrice") - myRow("Discount")))

                                                                              myRow("EXTPRICES") = FormatCurrency((myRow("UnitPrice") - ExtClass.DiscountPercent) * myRow("CURQTY"), 2) & vbCrLf & FormatCurrency((myRow("UnitPrice") - ExtClass.DiscountPercent) * myRow("SHPQTY"), 2)

                                                                              myRow("STK") = myReader("STK")

                                                                              If prt.MTO Then
                                                                                  _OrderPolicy = "MTO"
                                                                              Else
                                                                                  If prt.MTS Then
                                                                                      _OrderPolicy = "MTS"
                                                                                  Else
                                                                                      _OrderPolicy = ""
                                                                                  End If

                                                                              End If
                                                                              myRow("PLANID") = prt.PlanCode
                                                                              If prt.Description.Trim = "" Then
                                                                                  myRow("PartDescription1") = "Non-Inventory Part"
                                                                                  myRow("BothDesc") = "Non-Inventory Part"
                                                                              Else
                                                                                  myRow("PartDescription1") = prt.Description.Trim
                                                                                  myRow("BothDesc") = prt.Description.Trim & vbCrLf & prt.Description2.Trim
                                                                              End If
                                                                              myRow("PartDescription2") = prt.Description2.Trim

                                                                              myRow("UOM") = prt.SLSUOM 'dxing 9/17/10: changed to using part's SLSUOM from BOMUOM

                                                                              myRow("SlitWidth") = ExtClass.SlitWidth
                                                                              myRow("CoreSize") = ExtClass.CoreSize
                                                                              myRow("OD") = ExtClass.OutsideDiameter
                                                                              myRow("Unisource") = If(ExtClass.UnisourceStatus > 0, "Y", "N") + ExtClass.UnisourceStatus.ToString()
                                                                              myRow("UnisourceItem") = If(ExtClass.UnisourceStatus > 0, True, False)
                                                                              myRow("UnisourceStatus") = ExtClass.UnisourceStatus
                                                                              myRow("SentToUnisource") = ExtClass.SentToUnisource
                                                                              myRow("PromotionCode") = ExtClass.PromotionCode

                                                                              ' If this is a roll goods item, then show the slit width, od, core size; otherwise show the promotion code beneath the part number
                                                                              If myRow("PLANID") = "220" Then
                                                                                  myRow("PARTEXT") = myReader("PRTNUM").Trim & vbCrLf & FormatNumber(ExtClass.SlitWidth, clsNumber.SignificantDigits(ExtClass.SlitWidth)) & "-" & FormatNumber(ExtClass.CoreSize, clsNumber.SignificantDigits(ExtClass.CoreSize)) & "-" & FormatNumber(ExtClass.OutsideDiameter, clsNumber.SignificantDigits(ExtClass.OutsideDiameter))
                                                                              Else
                                                                                  myRow("PARTEXT") = myReader("PRTNUM").Trim & vbCrLf & ExtClass.PromotionCode
                                                                              End If

                                                                              Dim strPallets As String
                                                                              If ExtClass.Pallets Then
                                                                                  myRow("Pallets") = "Yes"
                                                                                  strPallets = "Yes"
                                                                              Else
                                                                                  myRow("Pallets") = "No"
                                                                                  strPallets = "No"
                                                                              End If

                                                                              Dim displayUOM As String = myRow("UOM")
                                                                              If prt.PlanCode = "220" Then
                                                                                  displayUOM = GetRollGoodsUOM(passorder)
                                                                              End If

                                                                              myRow("UOMPALLETS") = displayUOM & vbCrLf & strPallets 'dxing 9/17/10: changed to using part's SLSUOM from BOMUOM

                                                                              ' Get the comment from the SO Notes table
                                                                              myRow("COMNT") = "Edit or view line notes"

                                                                              ' Send static text that can be displayed for the Line Notes linkcolumn
                                                                              myRow("ShortComment") = "Line Notes"

                                                                              ' Convert the open/closed indicator to true or false
                                                                              If myReader("STATUS") = Convert.ToString(LineStatus.Closed) Then
                                                                                  myRow("Closed") = True
                                                                              Else
                                                                                  myRow("Closed") = False
                                                                              End If

                                                                              ' Free up memory used by the Number class
                                                                              clsNumber = Nothing
                                                                          End Sub)


                        ' Start a task to get the customer part information
                        Dim customerPartTask As Task = Task.Factory.StartNew(Sub()
                                                                                 ' Get the customer part number from the Customer Part Data Table
                                                                                 Dim custprt As New CustomerPartClass(ClassBase.DataSource.Max, myReader("Customer"), myReader("PRTNUM"), ExtClass.SlitWidth)
                                                                                 myRow("CustomerPart") = custprt.CustomerPart.Trim
                                                                                 myRow("CustomerPartDescription1") = custprt.Description1.Trim
                                                                                 myRow("CustomerPartDescription2") = custprt.Description2.Trim
                                                                                 myRow("CustomerPartDescription3") = custprt.Description3.Trim
                                                                                 myRow("CustomerPartDescription4") = custprt.Description4.Trim
                                                                                 custprt = Nothing
                                                                             End Sub)



                        ' Start a task to get the Transfer Order information
                        Dim transferInfoTask As Task = Task.Factory.StartNew(Sub()
                                                                                 ' Check the transfer orders to see if there are any out there for the item, to the warehouse
                                                                                 Dim TransferNumber As String = ""
                                                                                 Dim TransferStatus As Integer = 0
                                                                                 Dim TransferSubmitted As New Date(2050, 12, 31, 0, 0, 0)
                                                                                 Dim TransferShipped As New Date(2050, 12, 31, 0, 0, 0)
                                                                                 Dim TransferEstimatedDelivery As New Date(2050, 12, 31, 0, 0, 0)
                                                                                 Dim TransferQuantity As Integer = 0


                                                                                 ' There may be cases when this is running for non-inventory items, in that case
                                                                                 ' skip running altogether
                                                                                 If myReader("STK").ToString.Trim.IndexOf(" ") > 0 Then
                                                                                     ' Split the STK code on a space, and the second item will always be the warehouse code
                                                                                     Dim stkItems() As String = myReader("STK").ToString().Split(" ")

                                                                                     ' Split the DefaultSTK code on a space, and the second item will always be the warehouse code
                                                                                     Dim defaultstkItems() As String = myReader("DefaultSTK").ToString().Split(" ")

                                                                                     ' If tyhe warehouse the order is shipping from is MIL2 and the default stock warehouse is NEWB
                                                                                     ' then we have an item that needs a transfer order, so check if the transfer order exists
                                                                                     If stkItems(1) = "MIL2" And defaultstkItems(1) = "NEWB" Then
                                                                                         myRow("AutoTransferItem") = True
                                                                                         TransferNumber = passorder
                                                                                         ' Call the function to check the detail of the auto transfer order
                                                                                         CheckAutoTransferOrders(passorder, myRow("PRTNUM"), TransferStatus, TransferSubmitted, TransferShipped, TransferEstimatedDelivery, TransferQuantity)
                                                                                     Else
                                                                                         ' Call the function to check the detail of the transfer orders for the item
                                                                                         CheckTransferOrders(stkItems(1), myRow("PRTNUM"), TransferNumber, TransferSubmitted, TransferShipped, TransferEstimatedDelivery, TransferQuantity)
                                                                                     End If
                                                                                 End If
                                                                                 myRow("TransferStatus") = TransferStatus
                                                                                 myRow("TransferNumber") = TransferNumber
                                                                                 myRow("TransferSubmitted") = TransferEstimatedDelivery
                                                                                 myRow("TransferShipped") = TransferEstimatedDelivery
                                                                                 myRow("TransferEstimatedDelivery") = TransferEstimatedDelivery
                                                                                 myRow("TransferQuantity") = TransferQuantity
                                                                             End Sub)

                        ' Wait for all the threaded tasks to complete
                        Task.WaitAll(baseInfoTask, priceInfoTask, customerPartTask, transferInfoTask)

                        prt = Nothing
                        ExtClass = Nothing

                        ' Add the record to the table
                        myDataSet.Tables(0).Rows.Add(myRow)
                    End While
                End Using
            End Using
        End Using

        ' If there was no data that matched, the order is not deleteable...
        If myDataSet.Tables(0).Rows.Count = 0 Then
            _Deleteable = False
        End If

        ' Return the created dataset
        Return myDataSet
    End Function

    Private Function GetRollGoodsUOM(ByVal ORDNUM) As String
        Dim rtnUOM As String = ""

        Using connection As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            connection.Open()

            Dim command As String = "SELECT @UOM = ShortDescription from SalesOrderMasterExt somex join SalesOrderConversion soc on somex.ConversionID = soc.ID where ORDNUM = @ORDNUM"
            Using cmd As New SqlCommand(command, connection)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))

                Dim parmUOM As New SqlParameter("@UOM", SqlDbType.NVarChar, 3)
                parmUOM.Direction = ParameterDirection.Output
                parmUOM.Value = Nothing
                cmd.Parameters.Add(parmUOM)

                cmd.ExecuteNonQuery()
                rtnUOM = parmUOM.Value
            End Using
        End Using

        Return rtnUOM
    End Function

    Private Function ConvertFromLBS(ByVal PartNumber As String, ByVal ConversionType As Integer, ByVal PlanID As String, ByVal Value As Decimal) As Decimal
        Dim rtnValue As Decimal = Value

        ' We only do the conversion IF the part is for a DRC roll good
        If PlanID = "220" Then
            Select Case ConversionType
                Case 1 ' 
                    rtnValue = Value
                Case 2 ' this is metric tons
                    Dim Factor As Decimal = 2204.62262
                    rtnValue = Value / Factor
                Case 3
                    Dim Factor As Decimal = 2.2
                    rtnValue = Value / Factor
                Case 4
                    'rtnValue = (333.333 / BasisWeight) * (Value * (36 / SlitWidth))
            End Select
        End If

        Return rtnValue
    End Function

    Private Function ConvertFromLBSPrice(ByVal ConversionType As Integer, ByVal PlanID As String, ByVal Value As Decimal) As Decimal
        Dim rtnValue As Decimal = Value

        ' We only do the conversion IF the part is for a DRC roll good
        If PlanID = "220" Then
            Select Case ConversionType
                Case 1 ' 
                    rtnValue = Value
                Case 2 ' this is metric tons
                    Dim Factor As Decimal = 2204.62262
                    rtnValue = Value * Factor
                Case 3
                    Dim Factor As Decimal = 2.2
                    rtnValue = Value * Factor
            End Select
        End If

        Return rtnValue
    End Function

    Private Function ReadPartSales(ByVal passpart As String) As Decimal
        Dim CustomerPrice As Decimal = 0

        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()

            Dim strSQL As String = "Select PRICE_29 " &
                               "From ""Part_Sales"" " &
                               "Where PRTNUM_29 = '" & passpart.Trim & "' "

            Using cmd As New SqlCommand(strSQL, connection)
                Using PartSalesReader As SqlDataReader = cmd.ExecuteReader()
                    PartSalesReader.Read()
                    CustomerPrice = PartSalesReader("PRICE_29")
                End Using
            End Using
        End Using

        ' Return the price from the Part Sales table
        Return CustomerPrice
    End Function

    Private Function ReadCustomerPartData(ByVal passcust As String, ByVal passpart As String, ByVal passwidth As Decimal) As String
        Dim CustomerPart As String = ""
        Dim CustomerPartDesc1 As String = ""
        Dim CustomerPartDesc2 As String = ""
        Dim CustomerPartDesc3 As String = ""
        Dim CustomerPartDesc4 As String = ""

        ' Check if there is a direct sellars to customer part correlation using the
        ' max customer part information
        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()

            Dim strSQL As String = "Select CUSTPRT_103 " &
                                   "From ""Customer_Part_Data"" " &
                                   "Where CUSTID_103 = '" & passcust.Trim & "' and PRTNUM_103 = '" & passpart.Trim & "'"

            Using cmd As New SqlCommand(strSQL, connection)
                Using CustPartReader As SqlDataReader = cmd.ExecuteReader()
                    If CustPartReader.Read() Then
                        CustomerPart = CustPartReader("CUSTPRT_103")
                    Else
                        CustomerPart = ""
                    End If
                End Using
            End Using
        End Using

        ' If there is no max customer part, then see if there is one using the slit width
        ' combination from the sql server
        If CustomerPart = "" Then
            ' Declare the SQL data layer class
            Dim oSQL As New SqlService(ConnectionString)

            ' Add the parameter to the command object
            oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, passcust, ParameterDirection.Input)
            oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, passpart, ParameterDirection.Input)
            oSQL.AddParameter("@Width", SqlDbType.Decimal, 0, passwidth, ParameterDirection.Input)

            ' Execute the stored procedure, and if it returned a record, then get the customer part
            Using dr As SqlClient.SqlDataReader = oSQL.RunProcReader("ReadCustomerSlitWidthPart")
                If dr.Read Then
                    CustomerPart = dr("CUSTPRT")
                End If
            End Using

            oSQL = Nothing
        End If

        ' Return the price from the Part Sales table
        Return CustomerPart.Trim
    End Function

    Public Sub Read(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String)
        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()

            Dim strSQL As String = "Select * " &
                       "From ""SO_Detail"" " &
                       "Where ORDNUM_28 = '" & ORDNUM & "' " &
                       "and LINNUM_28 = '" & LINNUM & "' " &
                       "and DELNUM_28 = '" & DELNUM & "'"

            Using cmd As New SqlCommand(strSQL, connection)
                Using myReader As SqlDataReader = cmd.ExecuteReader()
                    If myReader.Read() Then
                        _ORDNUM = myReader("ORDNUM_28")
                        _LINNUM = myReader("LINNUM_28")
                        _DELNUM = myReader("DELNUM_28")
                        _STATUS = myReader("STATUS_28")
                        _CUSTID = myReader("CUSTID_28")
                        _PRTNUM = myReader("PRTNUM_28")
                        _EDILIN = myReader("EDILIN_28")
                        _TAXABL = myReader("TAXABL_28")
                        _GLXREF = myReader("GLXREF_28")
                        Try
                            _CURDUE = dbRecordDate(myReader("CURDUE_28"))
                        Catch ex As Exception
                            _CURDUE = DefaultDate
                        End Try
                        Try
                            _ORGDUE = dbRecordDate(myReader("ORGDUE_28"))
                        Catch ex As Exception
                            _ORGDUE = DefaultDate
                        End Try
                        Try
                            _CUSDUE = dbRecordDate(myReader("CUSDUE_28"))
                        Catch ex As Exception
                            _CUSDUE = DefaultDate
                        End Try
                        '_FILL03 = myReader("FILL03_28")
                        If IsDBNull(myReader("SHPDTE_28")) Then
                            _SHPDTE = DefaultShipDate
                        Else
                            _SHPDTE = myReader("SHPDTE_28")
                        End If
                        _FILL04 = myReader("FILL04_28")
                        _SLSUOM = myReader("SLSUOM_28")
                        _REFRNC = myReader("REFRNC_28")
                        _PRICE = myReader("PRICE_28")
                        _ORGQTY = myReader("ORGQTY_28")
                        _CURQTY = myReader("CURQTY_28")
                        _BCKQTY = myReader("BCKQTY_28")
                        _SHPQTY = myReader("SHPQTY_28")
                        _CURSHP = myReader("CURSHP_28")
                        _DUEQTY = myReader("DUEQTY_28")
                        _INVQTY = myReader("INVQTY_28")
                        _DISC = myReader("DISC_28")
                        _STYPE = myReader("STYPE_28")
                        _PRNT = myReader("PRNT_28")
                        _AKPRNT = myReader("AKPRNT_28")
                        _STK = myReader("STK_28")
                        _COCFLG = myReader("COCFLG_28")
                        _FORCUR = myReader("FORCUR_28")
                        _HSTAT = myReader("HSTAT_28")
                        _SLSREP = myReader("SLSREP_28")
                        _COMMIS = myReader("COMMIS_28")
                        _DRPSHP = myReader("DRPSHP_28")
                        _QUMQTY = myReader("QUMQTY_28")
                        _TAXCDE1 = myReader("TAXCDE1_28")
                        _TAX1 = myReader("TAX1_28")
                        _TAXCDE2 = myReader("TAXCDE2_28")
                        _TAX2 = myReader("TAX2_28")
                        _TAXCDE3 = myReader("TAXCDE3_28")
                        _TAX3 = myReader("TAX3_28")
                        _MCOMP = myReader("MCOMP_28")
                        _MSITE = myReader("MSITE_28")
                        _UDFKEY = myReader("UDFKEY_28")
                        _UDFREF = myReader("UDFREF_28")
                        _DEXPFLG = myReader("DEXPFLG_28")
                        _FILLER = myReader("FILLER_28")

                        _COST = myReader("COST_28")
                        _MARKUP = myReader("MARKUP_28")
                        _PROBAB = myReader("PROBAB_28")
                        _QTORD = myReader("QTORD_28")
                        _QTDEL = myReader("QTDEL_28")
                        _QTLINE = myReader("QTLINE_28")
                        _XDFINT = myReader("XDFINT_28")
                        _XDFFLT = myReader("XDFFLT_28")
                        _XDFBOL = myReader("XDFBOL_28")
                        Try
                            _XDFDTE = dbRecordDate(myReader("XDFDTE_28"))
                        Catch ex As Exception
                            _XDFDTE = DefaultDate
                        End Try
                        _XDFTXT = myReader("XDFTXT_28")
                        _CREATEDBY = myReader("CreatedBy")
                        Try
                            _CREATIONDATE = dbRecordDate(myReader("CreationDate"))
                        Catch ex As Exception
                            _CREATIONDATE = DefaultDate
                        End Try
                        _MODIFIEDBY = myReader("ModifiedBy")
                        Try
                            _MODIFICATIONDATE = dbRecordDate(myReader("ModificationDate"))
                        Catch ex As Exception
                            _MODIFICATIONDATE = DefaultDate
                        End Try
                        Try
                            _BOKDTE = dbRecordDate(myReader("BOKDTE_28"))
                        Catch ex As Exception
                            _BOKDTE = DefaultDate
                        End Try
                        Try
                            _DBKDTE = dbRecordDate(myReader("DBKDTE_28"))
                        Catch ex As Exception
                            _DBKDTE = DefaultDate
                        End Try
                        _REVLEV = myReader("REVLEV_28")

                        ' Get the slitwidth, coresize, od and pallets from the SalesOrderDetailExt SQL table 
                        Dim ExtClass As New SODetailExtClass(ORDNUM, LINNUM, DELNUM)
                        _SlitWidth = ExtClass.SlitWidth
                        _CoreSize = ExtClass.CoreSize
                        _OutsideDiameter = ExtClass.OutsideDiameter
                        _Pallets = ExtClass.Pallets
                    Else
                        _ORDNUM = ""
                        _LINNUM = ""
                        _DELNUM = ""
                        _STATUS = ""
                        _CUSTID = ""
                        _PRTNUM = ""
                        _EDILIN = ""
                        _TAXABL = ""
                        _GLXREF = ""
                        _CURDUE = DefaultDate
                        _FILL01 = ""
                        _ORGDUE = DefaultDate
                        '_FILL02 = ""
                        _CUSDUE = DefaultDate
                        _FILL03 = ""
                        _SHPDTE = DefaultDate
                        _FILL04 = ""
                        _SLSUOM = ""
                        _REFRNC = ""
                        _PRICE = 0
                        _ORGQTY = 0
                        _CURQTY = 0
                        _BCKQTY = 0
                        _SHPQTY = 0
                        _CURSHP = 0
                        _DUEQTY = 0
                        _INVQTY = 0
                        _DISC = 0
                        _STYPE = ""
                        _PRNT = ""
                        _AKPRNT = ""
                        _STK = ""
                        _COCFLG = ""
                        _FORCUR = 0
                        _HSTAT = ""
                        _SLSREP = ""
                        _COMMIS = 0
                        _DRPSHP = ""
                        _QUMQTY = 0
                        _TAXCDE1 = ""
                        _TAX1 = 0
                        _TAXCDE2 = ""
                        _TAX2 = 0
                        _TAXCDE3 = ""
                        _TAX3 = 0
                        _MCOMP = ""
                        _MSITE = ""
                        _UDFKEY = ""
                        _UDFREF = ""
                        _DEXPFLG = ""
                        _FILLER = ""

                        _COST = 0
                        _MARKUP = 0
                        _PROBAB = 0
                        _QTDEL = ""
                        _QTLINE = ""
                        _QTORD = ""
                        _XDFINT = 0
                        _XDFFLT = 0
                        _XDFBOL = 0
                        _XDFDTE = DefaultDate
                        _XDFTXT = ""
                        _CREATEDBY = ""
                        _CREATIONDATE = DefaultDate
                        _MODIFIEDBY = ""
                        _MODIFICATIONDATE = DefaultDate
                        _BOKDTE = DefaultDate
                        _DBKDTE = DefaultDate
                        _REVLEV = ""
                        ' Set the detail extension fields as well
                        _SlitWidth = 0
                        _CoreSize = 0
                        _OutsideDiameter = 0
                        _Pallets = False
                    End If
                End Using
            End Using
        End Using
    End Sub

    Public Function ReadAll(ByVal ORDNUM As String, ByVal SelectedPart As String) As List(Of String)
        Dim rtnData As New List(Of String)

        Using connection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            connection.Open()

            Dim strSQL As New StringBuilder
            strSQL.Append("Select LINNUM_28, DELNUM_28 " &
                                   "From SO_Detail " &
                                   "Where ORDNUM_28 = '" & ORDNUM & "' ")

            If SelectedPart <> "" Then
                strSQL.Append(" and PRTNUM_28 like '" + SelectedPart.Trim + "%' ")
            End If

            strSQL.Append("order by LINNUM_28")

            Using cmd As New SqlCommand(strSQL.ToString(), connection)
                Using myReader As SqlDataReader = cmd.ExecuteReader()
                    While myReader.Read()
                        rtnData.Add(myReader("LINNUM_28").ToString.Trim + "-" + myReader("DELNUM_28").ToString.Trim)
                    End While
                End Using
            End Using
        End Using

        Return rtnData
    End Function

    Public Function Read(ByVal passorder As String) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, passorder, ParameterDirection.Input)

        Return oSQL.RunProcReader("GetSalesOrderDetails")
    End Function

    Public Function Read(ByVal passorder As String, ByVal TableName As String) As DataSet
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, passorder, ParameterDirection.Input)

        Return oSQL.RunProcDataSet("GetSalesOrderDetails", TableName)
    End Function

    Private Function RecommendedPrice(ByVal CustomerType As String, ByVal Part As String, ByVal Qty As Decimal) As Decimal
        Dim Price As Decimal = 0

        ' First check the price break class
        Dim _priceBreakClass As PriceBreakClass = New PriceBreakClass
        Try
            Price = _priceBreakClass.Read(Part.ToUpper.Trim, CustomerType.Trim, Qty)
        Catch ex As Exception
            Price = 0
        End Try
        _priceBreakClass = Nothing

        ' If no price breaks returned, then try the part sales price breaks
        If Price = 0 Then
            Dim _partSalesClass As PartSalesClass = New PartSalesClass
            Try
                Price = _partSalesClass.Read(Part.ToUpper.Trim, Qty)
            Catch ex As Exception
                Price = 0
            End Try
            _partSalesClass = Nothing
        End If

        Return Price
    End Function

    ' Function used to update the line item status
    Private Sub Update(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pClosed As Boolean)
        Dim Status As String

        If pClosed Then
            Status = LineStatus.Closed
        Else
            Status = LineStatus.Open
        End If

        ' Read the requested sales order detail record
        _ErrorDescription = "Line Item Read Max"
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        ' Fill in the change functions new data structure by first setting the 
        ' data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("STATUS_28").First().Value = Status

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' Update the record status on the Sellars SQL sales order detail table
                UpdateSellars(ORDNUM, LINNUM, DELNUM, Status)
        End Select
    End Sub

    ' Function used to update the line part number
    Public Sub UpdatePartNumber(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal PartNumber As String, ByVal STK As String)
        ' Read the requested sales order detail record
        _ErrorDescription = "Line Item Read Max"
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        ' Fill in the change functions new data structure by first setting the 
        ' data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("PRTNUM_28").First().Value = PartNumber.PadRight(30)
        newSODXML.Descendants("STK_28").First().Value = STK.ToUpper.PadRight(8)

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' Update the record status on the Sellars SQL sales order detail table
                UpdateSellarsPart(ORDNUM, LINNUM, DELNUM, PartNumber)
        End Select
    End Sub

    Public Sub Update(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pStockCode As String)
        ' Read the requested sales order detail record
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("STK_28").First().Value = pStockCode.ToUpper.PadRight(8)

        'Update the Sales Order Detail record via MAX Update
        Dim retryCount As Integer = 0
        Dim retValue As Integer = 0

        While retryCount <= 3
            ' Increment the retry counter
            retryCount += 1

            ' call the change sales order line item
            retValue = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

            ' If the update call was successul, then exit the while loop
            If retValue <> 0 Then
                Exit While
            End If
        End While

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select
    End Sub

    ' This function updates the curdue and cusdue dates for a specific line item
    Public Sub Update(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pCurDue As Date, ByVal pCusDue As Date)
        ' Read the requested sales order detail record
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("CURDUE_28").First().Value = MakeDate(pCurDue)
        newSODXML.Descendants("CUSDUE_28").First().Value = MakeDate(pCusDue)

        'Update the Sales Order Detail record via MAX Update
        Dim retryCount As Integer = 0
        Dim retValue As Integer = 0

        While retryCount <= 3
            ' Increment the retry counter
            retryCount += 1

            ' call the change sales order line item
            retValue = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

            ' If the update call was successul, then exit the while loop
            If retValue <> 0 Then
                Exit While
            End If
        End While

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select
    End Sub

    Public Sub Update(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pDueDate As Date, ByVal GLCode As String)
        ' Read the requested sales order detail record
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("CURDUE_28").First().Value = MakeDate(pDueDate)
        newSODXML.Descendants("CUSDUE_28").First().Value = MakeDate(pDueDate)
        newSODXML.Descendants("GLXREF_28").First().Value = GLCode.ToUpper.PadRight(32)

        'Update the Sales Order Detail record via MAX Update
        Dim retryCount As Integer = 0
        Dim retValue As Integer = 0

        While retryCount <= 3
            ' Increment the retry counter
            retryCount += 1

            ' call the change sales order line item
            retValue = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

            ' If the update call was successul, then exit the while loop
            If retValue <> 0 Then
                Exit While
            End If
        End While

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select
    End Sub

    Public Sub Update(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pDueDate As Date, ByVal GLCode As String, ByVal pStockCode As String)
        ' Read the requested sales order detail record
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("CURDUE_28").First().Value = MakeDate(pDueDate)
        newSODXML.Descendants("CUSDUE_28").First().Value = MakeDate(pDueDate)
        newSODXML.Descendants("STK_28").First().Value = pStockCode.ToUpper.PadRight(8)
        newSODXML.Descendants("GLXREF_28").First().Value = GLCode.ToUpper.PadRight(32)

        'Update the Sales Order Detail record via MAX Update
        Dim retryCount As Integer = 0
        Dim retValue As Integer = 0

        While retryCount <= 3
            ' Increment the retry counter
            retryCount += 1

            ' call the change sales order line item
            retValue = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

            ' If the update call was successul, then exit the while loop
            If retValue <> 0 Then
                Exit While
            End If
        End While

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select
    End Sub

    ' Function used to update the line item price
    Public Sub UpdatePrice(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pPrice As Decimal, ByVal pQuantity As Decimal)
        ' Read the requested sales order detail record
        _ErrorDescription = "Line Item Read Max"
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        ' Fill in the change functions new data structure by first setting the 
        ' data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("PRICE_28").First().Value = pPrice
        newSODXML.Descendants("FORCUR_28").First().Value = pPrice
        newSODXML.Descendants("CURQTY_28").First().Value = pQuantity

        ' Adjust the DUEQTY appropriately based on changes to the current order qty
        Dim newCurQty As Integer = Integer.Parse(newSODXML.Descendants("CURQTY_28").First().Value)
        Dim oldCurQty As Integer = Integer.Parse(oldSODXML.Descendants("CURQTY_28").First().Value)
        Dim oldDueQty As Integer = Integer.Parse(oldSODXML.Descendants("DUEQTY_28").First().Value)

        If newCurQty > oldCurQty Then
            newSODXML.Descendants("DUEQTY_28").First().Value = oldDueQty + (newCurQty - oldCurQty)
        Else
            If newCurQty < oldCurQty Then
                Dim newDueQty As Integer = oldDueQty - (oldCurQty - newCurQty)

                If newDueQty < 0 Then
                    newSODXML.Descendants("DUEQTY_28").First().Value = 0
                Else
                    newSODXML.Descendants("DUEQTY_28").First().Value = newDueQty
                End If
            End If
        End If

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select
    End Sub

    Public Sub Update(ByVal value As DataSource, ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pQuantity As Decimal, ByVal pPrice As Decimal, ByVal pCusDue As Date, ByVal pCurDue As Date, ByVal pGLCode As String, ByVal pStockCode As String,
                      ByVal TaxCode1 As String, ByVal TaxRate1 As Decimal,
                      ByVal TaxCode2 As String, ByVal TaxRate2 As Decimal,
                      ByVal TaxCode3 As String, ByVal TaxRate3 As Decimal,
                      Optional ByRef taxAmtChanged As Boolean = False)
        ' Read the requested sales order detail record
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("GLXREF_28").First().Value = pGLCode.ToUpper.PadRight(32)
        newSODXML.Descendants("CURDUE_28").First().Value = MakeDate(pCurDue)
        newSODXML.Descendants("CUSDUE_28").First().Value = MakeDate(pCusDue)
        newSODXML.Descendants("PRICE_28").First().Value = pPrice
        newSODXML.Descendants("FORCUR_28").First().Value = pPrice
        newSODXML.Descendants("CURQTY_28").First().Value = pQuantity
        newSODXML.Descendants("STK_28").First().Value = pStockCode.ToUpper.PadRight(8)
        newSODXML.Descendants("TAXCDE1_28").First().Value = TaxCode1.PadRight(7)
        newSODXML.Descendants("TAXCDE2_28").First().Value = TaxCode2.PadRight(7)
        newSODXML.Descendants("TAXCDE3_28").First().Value = TaxCode3.PadRight(7)

        Dim forcur As Double = Double.Parse(newSODXML.Descendants("FORCUR_28").First().Value)
        Dim curqty As Double = Double.Parse(newSODXML.Descendants("CURQTY_28").First().Value)

        If newSODXML.Descendants("TAXABL_28").First().Value = "Y" Then
            newSODXML.Descendants("TAX1_28").First().Value = Math.Round(TaxRate1 * (forcur * curqty), 2)
            newSODXML.Descendants("TAX2_28").First().Value = Math.Round(TaxRate2 * (forcur * curqty), 2)
            newSODXML.Descendants("TAX3_28").First().Value = Math.Round(TaxRate3 * (forcur * curqty), 2)
        End If

        taxAmtChanged = oldSODXML.Descendants("TAX1_28").First().Value <> newSODXML.Descendants("TAX1_28").First().Value OrElse
                        oldSODXML.Descendants("TAX2_28").First().Value <> newSODXML.Descendants("TAX2_28").First().Value OrElse
                        oldSODXML.Descendants("TAX3_28").First().Value <> newSODXML.Descendants("TAX3_28").First().Value

        ' Adjust the DUEQTY appropriately based on changes to the current order qty
        Dim newCurQty As Integer = Integer.Parse(newSODXML.Descendants("CURQTY_28").First().Value)
        Dim oldCurQty As Integer = Integer.Parse(oldSODXML.Descendants("CURQTY_28").First().Value)
        Dim oldDueQty As Integer = Integer.Parse(oldSODXML.Descendants("DUEQTY_28").First().Value)

        If newCurQty > oldCurQty Then
            newSODXML.Descendants("DUEQTY_28").First().Value = oldDueQty + (newCurQty - oldCurQty)
        Else
            If newCurQty < oldCurQty Then
                Dim newDueQty As Integer = oldDueQty - (oldCurQty - newCurQty)

                If newDueQty < 0 Then
                    newSODXML.Descendants("DUEQTY_28").First().Value = 0
                Else
                    newSODXML.Descendants("DUEQTY_28").First().Value = newDueQty
                End If
            End If
        End If

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' update the Sellars SQL sales order detail table right away as well
                UpdateSellars(pORDNUM, pLINNUM, pDELNUM, Decimal.Parse(newSODXML.Descendants("CURQTY_28").First().Value), Decimal.Parse(newSODXML.Descendants("DUEQTY_28").First().Value), pCurDue)
        End Select
    End Sub

    Public Sub UpdateTaxDetails(ByVal value As DataSource, ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String,
                  ByVal TaxCode1 As String, ByVal TaxRate1 As Decimal,
                  ByVal TaxCode2 As String, ByVal TaxRate2 As Decimal,
                  ByVal TaxCode3 As String, ByVal TaxRate3 As Decimal,
                  Optional ByRef taxAmtChanged As Boolean = False)
        ' Read the requested sales order detail record
        Read(pORDNUM, pLINNUM, pDELNUM)

        Dim oldSODXML As XDocument = FillStructure()
        Dim newSODXML As XDocument = FillStructure()

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed
        newSODXML.Descendants("TAXCDE1_28").First().Value = TaxCode1.PadRight(7)
        newSODXML.Descendants("TAXCDE2_28").First().Value = TaxCode2.PadRight(7)
        newSODXML.Descendants("TAXCDE3_28").First().Value = TaxCode3.PadRight(7)

        Dim forcur As Double = Double.Parse(newSODXML.Descendants("FORCUR_28").First().Value)
        Dim curqty As Double = Double.Parse(newSODXML.Descendants("CURQTY_28").First().Value)

        If newSODXML.Descendants("TAXABL_28").First().Value = "Y" Then
            newSODXML.Descendants("TAX1_28").First().Value = Math.Round(TaxRate1 * (forcur * curqty), 2)
            newSODXML.Descendants("TAX2_28").First().Value = Math.Round(TaxRate2 * (forcur * curqty), 2)
            newSODXML.Descendants("TAX3_28").First().Value = Math.Round(TaxRate3 * (forcur * curqty), 2)
        End If

        taxAmtChanged = oldSODXML.Descendants("TAX1_28").First().Value <> newSODXML.Descendants("TAX1_28").First().Value OrElse
                        oldSODXML.Descendants("TAX2_28").First().Value <> newSODXML.Descendants("TAX2_28").First().Value OrElse
                        oldSODXML.Descendants("TAX3_28").First().Value <> newSODXML.Descendants("TAX3_28").First().Value

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrderLineItemXML(newSODXML.ToString(), oldSODXML.ToString())

        Select Case retValue
            Case 0
                Throw New System.Exception("Error updating sales order detail.")
        End Select
    End Sub

    ' The following routine will update a sales order detail record on the sellars SQL sales order table
    Private Sub UpdateSellars(ByVal pORDNUM As String, ByVal pLINNUM As String, ByVal pDELNUM As String, ByVal pCurQty As Decimal, ByVal pDueQty As Decimal, ByVal pCurDue As Date)
        Using connection As New System.Data.SqlClient.SqlConnection(ConnectionString)
            connection.Open()

            Using cmd As New System.Data.SqlClient.SqlCommand("UpdateSalesOrderDetail", connection)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ORDNUM", pORDNUM))
                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LINNUM", pLINNUM))
                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DELNUM", pDELNUM))
                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CURQTY", pCurQty))
                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DUEQTY", pDueQty))
                cmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CURDUE", pCurDue))

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' The following routine will update a sales order detail record on the sellars SQL sales order table
    Private Sub UpdateSellarsPart(ByVal SalesOrder As String, ByVal Line As String, ByVal Delivery As String, ByVal PartNumber As String)
        Using connection As New SqlConnection(ConnectionString)
            connection.Open()

            Using cmd As New SqlCommand("UpdateSalesOrderDetailPart", connection)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.Add(New SqlParameter("@ORDNUM", SalesOrder))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", Line))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", Delivery))
                cmd.Parameters.Add(New SqlParameter("@PART", PartNumber))

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' The following routine will update a sales order detail record on the sellars SQL sales order table
    Private Sub UpdateSellars(ByVal SalesOrder As String, ByVal Line As String, ByVal Delivery As String, ByVal Status As String)
        Using connection As New SqlConnection(ConnectionString)
            connection.Open()

            Using cmd As New SqlCommand("UpdateSalesOrderDetailStatus", connection)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.Add(New SqlParameter("@ORDNUM", SalesOrder))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", Line))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", Delivery))
                cmd.Parameters.Add(New SqlParameter("@STATUS", Status))

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function ValidateLine(ByVal SalesOrder As String, ByVal Add As Boolean, ByVal Line As String, ByVal Del As String, ByVal Part As String, ByVal InventoryPart As Boolean, ByVal QTY As String, ByVal Price As String, ByVal CoreSize As String, ByVal Width As String, ByVal OD As String, ByVal GLCode As String, ByVal DueDate As String, ByVal StockID As String) As ValidateLineResults
        Dim vld As New ValidateLineResults

        ' Make sure the Line entry is numeric
        If Not IsNumeric(Line) Then
            vld.ValidLine = False
            vld.ValidLineItem = False
        End If

        ' Make sure the DL entry is numeric
        If Not IsNumeric(Del) Then
            vld.ValidDelivery = False
            vld.ValidLineItem = False
        End If

        ' If we have a valid line and delivery AND not
        ' attempting an update
        If vld.ValidLine And vld.ValidDelivery And Add Then
            Read(SalesOrder, Line, Del)
            If LINNUM <> "" Then
                vld.ValidLine = False
                vld.ValidDelivery = False
                vld.ValidLineItem = False
            End If
        End If

        ' Make sure the Quantity entry is numeric
        If Not IsNumeric(QTY) Then
            vld.ValidQuantity = False
            vld.ValidLineItem = False
        End If

        ' Make sure that part is not blanks
        If Part.Trim = "" Then
            vld.ValidPart = False
            vld.ValidLineItem = False
        End If

        ' Make sure the Price entry is numeric
        If Not IsNumeric(Price) Then
            vld.ValidPrice = False
            vld.ValidLineItem = False
        End If

        ' Make sure the Due Date is a valid date
        If Not IsDate(DueDate) Then
            vld.ValidDueDate = False
            vld.ValidLineItem = False
        End If

        ' Make sure the SlitWidth entry is numeric
        If Not IsNumeric(Width) Then
            vld.ValidWidth = False
            vld.ValidLineItem = False
        End If

        ' Make sure the CoreSize entry is numeric
        If Not IsNumeric(CoreSize) Then
            vld.ValidCoreSize = False
            vld.ValidLineItem = False
        End If

        ' Make sure the Outside Diameter entry is numeric
        If Not IsNumeric(OD) Then
            vld.ValidOD = False
            vld.ValidLineItem = False
        End If

        ' Make sure that GL Code is not blanks
        If GLCode.Trim = "" Then
            vld.ValidGLCode = False
            vld.ValidLineItem = False
        Else
            Dim gl As New GeneralLedgerAcctClass
            If Not gl.Verify(GLCode) Then
                vld.ValidGLCode = False
                vld.ValidLineItem = False
            End If
        End If

        If Not InventoryPart Then
            If StockID.Trim = "" Then
                vld.ValidStockCode = False
                vld.ValidLineItem = False
            End If
        End If

        ' Return the object which shows the errors
        Return vld
    End Function

    Public Class ValidateLineResults
        Public ValidLineItem As Boolean = True
        Public ValidLine As Boolean = True
        Public ValidDelivery As Boolean = True
        Public ValidPart As Boolean = True
        Public ValidQuantity As Boolean = True
        Public ValidDueDate As Boolean = True
        Public ValidWidth As Boolean = True
        Public ValidCoreSize As Boolean = True
        Public ValidOD As Boolean = True
        Public ValidGLCode As Boolean = True
        Public ValidPrice As Boolean = True
        Public ValidStockCode As Boolean = True
        Public ErrorMessage As String = ""
    End Class

    Public Sub Clear()
        ClearProperties()
    End Sub

    Private Sub ClearProperties()
        _AKPRNT = ""
        _BCKQTY = 0
        _BOKDTE = DefaultDate
        _COCFLG = ""
        _COMMIS = 0
        _COST = 0
        _CREATEDBY = ""
        _CREATIONDATE = DefaultDate
        _CURDUE = DefaultDate
        _CURQTY = 0
        _CURSHP = 0
        _CUSDUE = DefaultDate
        _CUSTID = ""
        _DBKDTE = DefaultDate
        _DELNUM = ""
        _DEXPFLG = ""
        _DISC = 0
        _DRPSHP = ""
        _DUEQTY = 0
        _EDILIN = ""
        _FORCUR = 0
        _GLXREF = ""
        _HSTAT = ""
        _INVQTY = 0
        _LINNUM = ""
        _MARKUP = 0
        _MCOMP = ""
        _MODIFICATIONDATE = DefaultDate
        _MODIFIEDBY = ""
        _MSITE = ""
        _ORDNUM = ""
        _ORGDUE = DefaultDate
        _ORGQTY = 0
        _PRICE = 0
        _PROBAB = 0
        _PRNT = ""
        _PRTNUM = ""
        _QTLINE = ""
        _QTDEL = ""
        _QTORD = ""
        _QUMQTY = 0
        _REFRNC = ""
        _REVLEV = ""
        _SHPDTE = DefaultShipDate
        _SHPQTY = 0
        _SLSREP = ""
        _SLSUOM = ""
        _STATUS = ""
        _STK = ""
        _STYPE = ""
        _TAX1 = 0
        _TAX2 = 0
        _TAX3 = 0
        _TAXABL = ""
        _TAXCDE1 = ""
        _TAXCDE2 = ""
        _TAXCDE3 = ""
        _UDFKEY = ""
        _UDFREF = ""
        _XDFBOL = ""
        _XDFDTE = DefaultDate
        _XDFINT = 0
        _XDFFLT = 0
        _XDFTXT = ""
    End Sub

    Private Function FillStructure() As XDocument
        Dim salesOrderDetailXML As XDocument =
            <?xml version="1.0" encoding="utf-8"?>
            <eMAXExact>
                <SO_Detail_Table>
                    <SO_Detail>
                        <ORDNUM_28><%= ORDNUM %></ORDNUM_28>
                        <LINNUM_28><%= LINNUM %></LINNUM_28>
                        <DELNUM_28><%= DELNUM %></DELNUM_28>
                        <STATUS_28><%= STATUS %></STATUS_28>
                        <CUSTID_28><%= CUSTID %></CUSTID_28>
                        <PRTNUM_28><%= PRTNUM %></PRTNUM_28>
                        <EDILIN_28><%= EDILIN %></EDILIN_28>
                        <TAXABL_28><%= TAXABL %></TAXABL_28>
                        <GLXREF_28><%= GLXREF %></GLXREF_28>
                        <CURDUE_28><%= MakeDate(CURDUE) %></CURDUE_28>
                        <QTLINE_28><%= QTLINE %></QTLINE_28>
                        <ORGDUE_28><%= MakeDate(ORGDUE) %></ORGDUE_28>
                        <QTDEL_28><%= QTDEL %></QTDEL_28>
                        <CUSDUE_28><%= MakeDate(CUSDUE) %></CUSDUE_28>
                        <PROBAB_28><%= PROBAB %></PROBAB_28>
                        <SHPDTE_28><%= MakeDate(SHPDTE) %></SHPDTE_28>
                        <FILL04_28><%= FILL04 %></FILL04_28>
                        <SLSUOM_28><%= SLSUOM %></SLSUOM_28>
                        <REFRNC_28><%= REFRNC %></REFRNC_28>
                        <PRICE_28><%= PRICE %></PRICE_28>
                        <ORGQTY_28><%= ORGQTY %></ORGQTY_28>
                        <CURQTY_28><%= CURQTY %></CURQTY_28>
                        <BCKQTY_28><%= BCKQTY %></BCKQTY_28>
                        <SHPQTY_28><%= SHPQTY %></SHPQTY_28>
                        <CURSHP_28><%= CURSHP %></CURSHP_28>
                        <DUEQTY_28><%= DUEQTY %></DUEQTY_28>
                        <INVQTY_28><%= INVQTY %></INVQTY_28>
                        <DISC_28><%= DISC %></DISC_28>
                        <STYPE_28><%= STYPE %></STYPE_28>
                        <PRNT_28><%= PRNT %></PRNT_28>
                        <AKPRNT_28><%= AKPRNT %></AKPRNT_28>
                        <STK_28><%= STK %></STK_28>
                        <COCFLG_28><%= COCFLG %></COCFLG_28>
                        <FORCUR_28><%= FORCUR %></FORCUR_28>
                        <HSTAT_28><%= HSTAT %></HSTAT_28>
                        <SLSREP_28><%= SLSREP %></SLSREP_28>
                        <COMMIS_28><%= COMMIS %></COMMIS_28>
                        <DRPSHP_28><%= DRPSHP %></DRPSHP_28>
                        <QUMQTY_28><%= QUMQTY %></QUMQTY_28>
                        <TAXCDE1_28><%= TAXCDE1 %></TAXCDE1_28>
                        <TAX1_28><%= TAX1 %></TAX1_28>
                        <TAXCDE2_28><%= TAXCDE2 %></TAXCDE2_28>
                        <TAX2_28><%= TAX2 %></TAX2_28>
                        <TAXCDE3_28><%= TAXCDE3 %></TAXCDE3_28>
                        <TAX3_28><%= TAX3 %></TAX3_28>
                        <MCOMP_28><%= MCOMP %></MCOMP_28>
                        <MSITE_28><%= MSITE %></MSITE_28>
                        <UDFKEY_28><%= UDFKEY %></UDFKEY_28>
                        <UDFREF_28><%= UDFREF %></UDFREF_28>
                        <DEXPFLG_28><%= DEXPFLG %></DEXPFLG_28>
                        <COST_28><%= COST %></COST_28>
                        <MARKUP_28><%= MARKUP %></MARKUP_28>
                        <QTORD_28><%= QTORD %></QTORD_28>
                        <XDFINT_28><%= XDFINT %></XDFINT_28>
                        <XDFFLT_28><%= XDFFLT %></XDFFLT_28>
                        <XDFBOL_28><%= XDFBOL %></XDFBOL_28>
                        <XDFDTE_28><%= MakeDate(XDFDTE) %></XDFDTE_28>
                        <XDFTXT_28><%= XDFTXT %></XDFTXT_28>
                        <FILLER_28><%= FILLER %></FILLER_28>
                        <CreatedBy><%= CREATEDBY %></CreatedBy>
                        <CreationDate><%= MakeDate(CREATIONDATE) %></CreationDate>
                        <ModifiedBy><%= MODIFIEDBY %></ModifiedBy>
                        <ModificationDate><%= MakeDate(MODIFICATIONDATE) %></ModificationDate>
                        <BOKDTE_28><%= MakeDate(BOKDTE) %></BOKDTE_28>
                        <DBKDTE_28><%= MakeDate(DBKDTE) %></DBKDTE_28>
                        <REVLEV_28><%= REVLEV %></REVLEV_28>
                        <MANPRC_28></MANPRC_28>
                        <ORGPRC_28></ORGPRC_28>
                        <PRCALC_28></PRCALC_28>
                        <CLASS_28></CLASS_28>
                        <WARRES_28></WARRES_28>
                        <JOB_28></JOB_28>
                        <CSENDDTE_28></CSENDDTE_28>
                        <CONSGND_28></CONSGND_28>
                        <CURCONSGND_28></CURCONSGND_28>
                        <CONSIGNSTK_28></CONSIGNSTK_28>
                    </SO_Detail>
                </SO_Detail_Table>
            </eMAXExact>

        If (salesOrderDetailXML.Descendants("SHPDTE_28").First().Value = MakeDate(DefaultShipDate)) Then
            salesOrderDetailXML.Descendants("SHPDTE_28").Remove()
        End If

        Return salesOrderDetailXML
    End Function

    Public Function GetScheduledWorkOrders(ByVal passOrder As String, ByVal passLine As String, ByVal passDelivery As String) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, passOrder, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, passLine, ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, passDelivery, ParameterDirection.Input)

        ' Run the stored procedure
        Return oSQL.RunProcReader("ItemIsScheduled")
    End Function

    Private Function dbRecordDate(Optional ByVal pField As Object = Nothing) As Date
        If IsDBNull(pField) Then
            Return DefaultDate
        Else
            Return CDate(pField)
        End If
    End Function

    Private Sub CheckTransferOrders(ByVal Warehouse As String, ByVal PRTNUM As String, ByRef TransferNumber As String, ByRef TransferSubmitted As Date, ByRef TransferShipped As Date, ByRef TransferEstimatedDelivery As Date, ByRef TransferQuantity As Integer)
        Using connection As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            connection.Open()

            Dim command As String = "SELECT @OrderID = Min(isnull(tm.OrderID, '')), @Submitted = Min(isnull(SubmittedToFromWarehouse, '12/31/2050 00:00:00 AM')), @ShippedOn = Min(Isnull(ShippedOn, '12/31/2050 00:00:00 AM')), @EstimatedDelivery = Min(isnull(EstimatedDelivery, '12/31/2050 00:00:00 AM')), @Quantity = Sum(QuantityShipped) from TransferMaster tm join TransferDetail td on tm.ID = td.TMID where tm.Status in (3,4,5) and ToWarehouse = @Warehouse and PartNumber = @PRTNUM"

            Using cmd As New SqlCommand(command, connection)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@Warehouse", Warehouse))
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", PRTNUM))

                Dim parmOrderID As New SqlParameter("@OrderID", SqlDbType.NVarChar, 10)
                parmOrderID.Direction = ParameterDirection.Output
                parmOrderID.Value = Nothing
                cmd.Parameters.Add(parmOrderID)

                Dim parmSubmitted As New SqlParameter("@Submitted", SqlDbType.DateTime)
                parmSubmitted.Direction = ParameterDirection.Output
                parmSubmitted.Value = Nothing
                cmd.Parameters.Add(parmSubmitted)

                Dim parmShipped As New SqlParameter("@ShippedOn", SqlDbType.DateTime)
                parmShipped.Direction = ParameterDirection.Output
                parmShipped.Value = Nothing
                cmd.Parameters.Add(parmShipped)

                Dim parmEstimatedDelivery As New SqlParameter("@EstimatedDelivery", SqlDbType.DateTime)
                parmEstimatedDelivery.Direction = ParameterDirection.Output
                parmEstimatedDelivery.Value = Nothing
                cmd.Parameters.Add(parmEstimatedDelivery)

                Dim parmQuantity As New SqlParameter("@Quantity", SqlDbType.Int)
                parmQuantity.Direction = ParameterDirection.Output
                parmQuantity.Value = Nothing
                cmd.Parameters.Add(parmQuantity)

                'Perform the update to the database
                Try
                    cmd.ExecuteNonQuery()
                    TransferNumber = parmOrderID.Value
                    TransferEstimatedDelivery = parmEstimatedDelivery.Value
                    TransferQuantity = parmQuantity.Value
                    TransferSubmitted = parmSubmitted.Value
                    TransferShipped = parmShipped.Value
                Catch ex As Exception
                    TransferNumber = ""
                    TransferEstimatedDelivery = New Date(2050, 12, 31, 0, 0, 0)
                    TransferQuantity = 0
                End Try
            End Using
        End Using
    End Sub

    Private Sub CheckAutoTransferOrders(ByVal ORDID As String, ByVal PRTNUM As String, ByRef TransferStatus As Integer, ByRef TransferSubmitted As Date, ByRef TransferShipped As Date, ByRef TransferEstimatedDelivery As Date, ByRef TransferQuantity As Integer)
        Using connection As New SqlConnection(ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString)
            connection.Open()

            Dim command As String = "SELECT @Status = isnull(tm.Status, 0), @Submitted = isnull(SubmittedToFromWarehouse, '12/31/2050 00:00:00 AM'), @ShippedOn = Isnull(ShippedOn, '12/31/2050 00:00:00 AM'), @EstimatedDelivery = isnull(EstimatedDelivery, '12/31/2050 00:00:00 AM'), @Quantity = isnull(QuantityShipped, 0) from TransferMaster tm join TransferDetail td on tm.ID = td.TMID where OrderID = @OrderID and PartNumber = @PRTNUM"

            Using cmd As New SqlCommand(command, connection)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add(New SqlParameter("@PRTNUM", PRTNUM))
                cmd.Parameters.Add(New SqlParameter("@OrderID", ORDID))

                Dim parmStatus As New SqlParameter("@Status", SqlDbType.SmallInt)
                parmStatus.Direction = ParameterDirection.Output
                parmStatus.Value = Nothing
                cmd.Parameters.Add(parmStatus)

                Dim parmSubmitted As New SqlParameter("@Submitted", SqlDbType.DateTime)
                parmSubmitted.Direction = ParameterDirection.Output
                parmSubmitted.Value = Nothing
                cmd.Parameters.Add(parmSubmitted)

                Dim parmShipped As New SqlParameter("@ShippedOn", SqlDbType.DateTime)
                parmShipped.Direction = ParameterDirection.Output
                parmShipped.Value = Nothing
                cmd.Parameters.Add(parmShipped)

                Dim parmEstimatedDelivery As New SqlParameter("@EstimatedDelivery", SqlDbType.DateTime)
                parmEstimatedDelivery.Direction = ParameterDirection.Output
                parmEstimatedDelivery.Value = Nothing
                cmd.Parameters.Add(parmEstimatedDelivery)

                Dim parmQuantity As New SqlParameter("@Quantity", SqlDbType.Int)
                parmQuantity.Direction = ParameterDirection.Output
                parmQuantity.Value = Nothing
                cmd.Parameters.Add(parmQuantity)

                'Perform the update to the database
                Try
                    cmd.ExecuteNonQuery()
                    TransferStatus = parmStatus.Value
                    TransferEstimatedDelivery = parmEstimatedDelivery.Value
                    TransferQuantity = parmQuantity.Value
                    TransferSubmitted = parmSubmitted.Value
                    TransferShipped = parmShipped.Value
                Catch ex As Exception
                    TransferStatus = 0
                    TransferEstimatedDelivery = New Date(2050, 12, 31, 0, 0, 0)
                    TransferQuantity = 0
                End Try
            End Using
        End Using
    End Sub

    ' This function is the initialization routine that needs to be called before any MaxUpdateXML function is called
    Public Sub InitializeMax(ByVal connStr As String, ByVal comName As String, ByVal licPath As String, ByVal logPath As String, ByVal log As Boolean)
        ClassBase.Initialize(connStr, comName, licPath, logPath, log)
    End Sub

    Protected Overridable Function AddSalesOrderLineItemXML(ByVal xml As String) As Integer
        Return ClassBase.MAXUpdate.AddSOLineItemXML(xml, False)
    End Function

    Protected Overridable Function DeleteSalesOrderLineItemXML(ByVal sod As String) As Integer
        Return ClassBase.MAXUpdate.DeleteSalesOrderLineItemXML(sod)
    End Function

    Protected Overridable Function ChangeSalesOrderLineItemXML(ByVal newSODXML As String, ByVal oldSODXML As String) As Integer
        Return ClassBase.MAXUpdate.ChangeSalesOrderLineItemXML(newSODXML, oldSODXML)
    End Function
End Class
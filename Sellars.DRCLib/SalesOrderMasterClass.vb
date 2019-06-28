Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math
Imports System.Runtime.InteropServices
Imports System.Text

Public Class SalesOrderMasterClass
    Inherits ClassBase

    ' MaxUpdate Sales Order Functions
    Declare Function AddSalesOrder Lib "MaxOrdr2.dll" (ByVal ObjectHandle As Integer, ByRef SoeMstRecord As SOEMstItem) As Integer
    Declare Function ChangeSalesOrder Lib "MaxOrdr2.dll" (ByVal ObjectHandle As Integer, ByRef SoeMstRecord As SOEMstItem) As Integer
    Declare Function DeleteSalesOrder Lib "MaxOrdr2.dll" (ByVal ObjectHandle As Integer, ByVal OrdNum As String) As Integer

    Declare Function InitMAXOrder Lib "MaxOrdr2.dll" (ByVal PriConnectionString As String, ByVal CompanyName As String, ByVal LicensePath As String, ByVal LogPath As String, ByVal bIsASP As Integer, ByVal ErrRpt As Integer) As Integer
    Declare Function ShutdownMAXOrder Lib "MaxOrdr2.dll" (ByVal ObjectHandle As Integer) As Integer
    Private _MaxHandle As Integer = 0


    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=4)> _
    Public Structure SOEMstItem
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=8)> Public ORDNUM() As Char     '$ ORDER NUMBER
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=20)> Public CUSTID() As Char     '$ CUSTOMER IDENTIFIER
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=32)> Public GLXREF() As Char    '$ GENERAL LEDGER ACCOUNT
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public STYPE() As Char      '$ ORDER TYPE (DEMAND TYPE)
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public STATUS() As Char     '$ ORDER STATUS (DEMAND STATUS)
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=25)> Public CUSTPO() As Char   '$ CUSTOMER PURCHASE ORDER
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=20)> Public ORDID() As Char    '$ ORDERED BY
        Public ORDDTE As Integer '$ ORDER DATE - DATE ORDER CREATED
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public FILL01() As Char     '$ FILLER
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=10)> Public SHPCDE() As Char   '$ SHIPPING CODE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public REP1() As Char       '$ SALESREP ONE
        Public SPLIT1 As Single '! SPLIT ONE - Single
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public REP2() As Char       '$ SALESREP TWO
        Public SPLIT2 As Single '! SPLIT TWO - Single
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public REP3() As Char       '$ SALESREP THREE
        Public SPLIT3 As Single '! SPLIT THREE - Single
        Public COMMIS As Single '! COMMISSION - Single
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public TERMS() As Char      '$ TERMS CODE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public SHPVIA() As Char     '$ SHIP VIA
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public XURR() As Char       '$ OLD CURRENCY CODE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=15)> Public FOB() As Char      '$ FREIGHT ON BOARD
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public TAXCD1() As Char     '$ TAX CODE IDENTIFIER   # ONE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public TAXCD2() As Char     '$ TAX CODE IDENTIFIER   # TWO
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=7)> Public TAXCD3() As Char     '$ TAX CODE IDENTIFIER   # THREE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public COMNT1() As Char   '$ COMMENT # ONE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public COMNT2() As Char   '$ COMMENT # TWO
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public COMNT3() As Char   '$ COMMENT # THREE
        Public SHPLBL As Short '% SHIP LABELS
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public INVCE() As Char      '$ INVOICE READY FLAG
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=6)> Public APPINV() As Char     '$ APPLY TO INVOICE ("CR")
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public REASON() As Char     '$ REASON CODE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public NAME() As Char     '$ CUSTOMER NAME
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ADDR1() As Char    '$ ADDRESS 1
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ADDR2() As Char    '$ ADDRESS 2
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public CITY() As Char     '$ CITY
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public STATE() As Char      '$ STATE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ZIPCD() As Char    '$ ZIP CODE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public CNTRY() As Char    '$ COUNTRY
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=20)> Public PHONE() As Char    '$ SHIP TO PHONE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=20)> Public CNTCT() As Char    '$ CONTACT NAME
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=15)> Public TAXPRV() As Char   '$ PROVINCIAL TAX
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public FEDTAX() As Char     '$ FEDERAL TAX NUMBER
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public TAXABL() As Char     '$ TAXABLE FLAG
        Public EXCRTE As Single '! EXCHANGE RATE FOR ORDER - Single
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public FIXVAR() As Char     '$ FIXED/VARIABLE FLAG (F OR V ONLY)
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public CURR() As Char       '$ CURRENCY CODE FOR ORDER
        Public RCLDTE As Integer '! RE-CALL DATE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=2)> Public FILL02() As Char     '$ FILLER
        Public TTAX As Double '# TOTAL TAX
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public LNETAX() As Char     '$ TAX BY LINE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ADDR3() As Char    '$ ADDRESS 3
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ADDR4() As Char     '$ ADDRESS 4
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ADDR5() As Char    '$ ADDRESS 5
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=30)> Public ADDR6() As Char    '$ ADDRESS 6
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public MCOMP() As Char      '$ MULTI-COMPANY
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=3)> Public MSITE() As Char      '$ MULTI-SITE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=15)> Public UDFKEY() As Char   '$ USER DEFINED KEY
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=25)> Public UDFREF() As Char   '$ USER DEFINED REFERENCE
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=10)> Public SHPTHRU() As Char  '$ SHIP THRU
        Public XDFINT As Integer
        Public XDFFLT As Double
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=1)> Public XDFBOL() As Char   '$ 
        Public XDFDTE As Integer
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=100)> Public XDFTXT() As Char   '$ 
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=50)> Public FILLER() As Char   '$ 
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=100)> Public CREATEDBY() As Char   '$ 
        Public CREATIONDATE As Integer
        <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U1, SizeConst:=100)> Public MODIFIEDBY() As Char   '$ 
        Public MODIFICATIONDATE As Integer
    End Structure

    Private _found As Boolean = False

    Private _ORDNUM As String = ""
    Private _CUSTID As String = ""
    Private _GLXREF As String = ""
    Private _STYPE As String = ""
    Private _STATUS As String = ""
    Private _CUSTPO As String = ""
    Private _ORDID As String = ""
    Private _ORDDTE As Date = DefaultDate
    Private _FILL01 As String = ""
    Private _SHPCDE As String = ""
    Private _REP1 As String = ""
    Private _SPLIT1 As Single = 100
    Private _REP2 As String = ""
    Private _SPLIT2 As Single = 0
    Private _REP3 As String = ""
    Private _SPLIT3 As Single = 0
    Private _COMMIS As Single = 0
    Private _TERMS As String = ""
    Private _SHPVIA As String = ""
    Private _XURR As String = ""
    Private _FOB As String = ""
    Private _TAXCD1 As String = ""
    Private _TAXCD2 As String = ""
    Private _TAXCD3 As String = ""
    Private _COMNT1 As String = ""
    Private _COMNT2 As String = ""
    Private _COMNT3 As String = ""
    Private _SHPLBL As Integer = 0
    Private _INVCE As String = "N"
    Private _APPINV As String = ""
    Private _REASON As String = ""
    Private _NAME As String = ""
    Private _ADDR1 As String = ""
    Private _ADDR2 As String = ""
    Private _CITY As String = ""
    Private _STATE As String = ""
    Private _ZIPCD As String = ""
    Private _CNTRY As String = ""
    Private _PHONE As String = ""
    Private _CNTCT As String = ""
    Private _TAXPRV As String = ""
    Private _FEDTAX As String = "N"
    Private _TAXABL As String = ""
    Private _EXCRTE As Single = 0
    Private _FIXVAR As String = "F"
    Private _CURR As String = "USA"
    Private _RCLDTE As Date = DefaultDate
    Private _FILL02 As String = ""
    Private _TTAX As Double = 0
    Private _LNETAX As String = "N"
    Private _ADDR3 As String = ""
    Private _ADDR4 As String = ""
    Private _ADDR5 As String = ""
    Private _ADDR6 As String = ""
    Private _MCOMP As String = ""
    Private _MSITE As String = ""
    Private _UDFKEY As String = ""
    Private _UDFREF As String = ""
    Private _SHPTHRU As String = ""
    Private _FILLER As String = ""

    'new fields in max 5
    Private _XDFINT As Integer = 0
    Private _XDFFLT As Double = 0
    Private _XDFBOL As String = ""
    Private _XDFDTE As Date = DefaultDate
    Private _XDFTXT As String = ""
    Private _CREATEDBY As String = ""
    Private _CREATIONDATE As Date = DefaultDate
    Private _MODIFIEDBY As String = ""
    Private _MODIFICATIONDATE As Date = DefaultDate

    ' Keep track of sales order master extension data
    Private _EnteredBy As String
    Private _LastChanged As Date
    Private _LastPrinted As Date
    Private _ProofedBy As String
    Private _Notes As String
    Private _Hold As Boolean
    Private _HoldUserEmail As String

    Private _DefaultStockID As String
    Private _ShipFromStockID As String
    Private _SellarsOrderType As String
    Private _Finished As Boolean = False
    Private _EstimatedShipping As Decimal = 0
    Private _AllowFinish As Boolean = False

    Private _Function As String

    Public Enum OrderStatus
        Closed = 4
        Open = 3
    End Enum

    Public Enum OrderTypes
        FinishedGoods = 1
        RollGoods = 2
    End Enum

    Public Sub New()
    End Sub

    Public Sub New(ByVal passOrder As String)
        _ORDNUM = passOrder
        Read(passOrder)
    End Sub

#Region "Properties"

    Public Property FunctionPerformed() As String
        Get
            Return _Function
        End Get
        Set(ByVal value As String)
            _Function = value.Trim
        End Set
    End Property

    Public Property Found() As Boolean
        Get
            Return _found
        End Get
        Set(ByVal value As Boolean)
            _found = value
        End Set
    End Property

    Public Property ADDR1() As String
        Get
            Return _ADDR1.Trim
        End Get
        Set(ByVal value As String)
            _ADDR1 = value.Trim
        End Set
    End Property

    Public Property ADDR2() As String
        Get
            Return _ADDR2.Trim
        End Get
        Set(ByVal value As String)
            _ADDR2 = value.Trim
        End Set
    End Property

    Public Property ADDR3() As String
        Get
            Return _ADDR3.Trim
        End Get
        Set(ByVal value As String)
            _ADDR3 = value.Trim
        End Set
    End Property

    Public Property ADDR4() As String
        Get
            Return _ADDR4.Trim
        End Get
        Set(ByVal value As String)
            _ADDR4 = value.Trim
        End Set
    End Property

    Public Property ADDR5() As String
        Get
            Return _ADDR5.Trim
        End Get
        Set(ByVal value As String)
            _ADDR5.Trim()
        End Set
    End Property

    Public Property ADDR6() As String
        Get
            Return _ADDR6.Trim
        End Get
        Set(ByVal value As String)
            _ADDR6 = value.Trim
        End Set
    End Property

    Public Property AllowFinish() As Boolean
        Get
            Return _AllowFinish
        End Get
        Set(ByVal value As Boolean)
            _AllowFinish = value
        End Set
    End Property

    Public Property APPINV() As String
        Get
            Return _APPINV.Trim
        End Get
        Set(ByVal value As String)
            _APPINV = value.Trim
        End Set
    End Property

    Public Property CITY() As String
        Get
            Return _CITY.Trim
        End Get
        Set(ByVal value As String)
            _CITY = value.Trim
        End Set
    End Property

    Public Property CNTCT() As String
        Get
            Return _CNTCT.Trim
        End Get
        Set(ByVal value As String)
            _CNTCT = value.Trim
        End Set
    End Property

    Public Property CNTRY() As String
        Get
            Return _CNTRY.Trim
        End Get
        Set(ByVal value As String)
            _CNTRY = value.Trim
        End Set
    End Property

    Public Property CUSTID() As String
        Get
            Return _CUSTID.Trim
        End Get
        Set(ByVal value As String)
            _CUSTID = value.Trim
        End Set
    End Property

    Public Property CUSTPO() As String
        Get
            Return _CUSTPO.Trim
        End Get
        Set(ByVal value As String)
            _CUSTPO = value.Trim
        End Set
    End Property

    Private Property COMMIS() As Single
        Get
            Return _COMMIS
        End Get
        Set(ByVal value As Single)
            _COMMIS = value
        End Set
    End Property

    Public Property COMNT1() As String
        Get
            Return _COMNT1.Trim
        End Get
        Set(ByVal value As String)
            _COMNT1 = value.Trim
        End Set
    End Property

    Public Property COMNT2() As String
        Get
            Return _COMNT2.Trim
        End Get
        Set(ByVal value As String)
            _COMNT2 = value.Trim
        End Set
    End Property

    Public Property COMNT3() As String
        Get
            Return _COMNT3.Trim
        End Get
        Set(ByVal value As String)
            _COMNT3 = value.Trim
        End Set
    End Property

    Public Property CURR() As String
        Get
            Return _CURR.Trim
        End Get
        Set(ByVal value As String)
            _CURR = value.Trim
        End Set
    End Property

    Public Property DefaultStockID() As String
        Get
            Return _DefaultStockID
        End Get
        Set(ByVal value As String)
            _DefaultStockID = value.Trim
        End Set
    End Property

    Public Property EnteredBy() As String
        Get
            Return _EnteredBy.Trim
        End Get
        Set(ByVal value As String)
            _EnteredBy = value.Trim
        End Set
    End Property

    Public Property EstimatedShipping() As Decimal
        Get
            Return _EstimatedShipping
        End Get
        Set(ByVal value As Decimal)
            _EstimatedShipping = value
        End Set
    End Property

    Public Property EXCRTE() As Single
        Get
            Return _EXCRTE
        End Get
        Set(ByVal value As Single)
            _EXCRTE = value
        End Set
    End Property

    Public Property FEDTAX() As String
        Get
            Return _FEDTAX
        End Get
        Set(ByVal value As String)
            _FEDTAX = value.Trim
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

    Public Property FILL02() As String
        Get
            Return _FILL02
        End Get
        Set(ByVal value As String)
            _FILL02 = value.Trim
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

    Public Property Finished() As Boolean
        Get
            Return _Finished
        End Get
        Set(ByVal value As Boolean)
            _Finished = value
        End Set
    End Property

    Public Property FIXVAR() As String
        Get
            Return _FIXVAR.Trim
        End Get
        Set(ByVal value As String)
            _FIXVAR = value.Trim
        End Set
    End Property

    Public Property FOB() As String
        Get
            Return _FOB.Trim
        End Get
        Set(ByVal value As String)
            _FOB = value.Trim
        End Set
    End Property

    Public Property GLXREF() As String
        Get
            Return _GLXREF.Trim
        End Get
        Set(ByVal value As String)
            _GLXREF = value.Trim
        End Set
    End Property

    Public Property Hold() As Boolean
        Get
            Return _Hold
        End Get
        Set(ByVal value As Boolean)
            _Hold = value
        End Set
    End Property

    Public Property HoldUserEmail() As String
        Get
            Return _HoldUserEmail
        End Get
        Set(ByVal value As String)
            _HoldUserEmail = value.Trim
        End Set
    End Property

    Public Property INVCE() As String
        Get
            Return _INVCE.Trim
        End Get
        Set(ByVal value As String)
            _INVCE = value.Trim
        End Set
    End Property

    Public ReadOnly Property IsClosed() As Boolean
        Get
            Return STATUS = "4"
        End Get
    End Property

    Public Property LastChanged() As Date
        Get
            Return _LastChanged
        End Get
        Set(ByVal value As Date)
            _LastChanged = value
        End Set
    End Property

    Public Property LastPrinted() As Date
        Get
            Return _LastPrinted
        End Get
        Set(ByVal value As Date)
            _LastPrinted = value
        End Set
    End Property

    Public Property LNETAX() As String
        Get
            Return _LNETAX.Trim
        End Get
        Set(ByVal value As String)
            _LNETAX = value.Trim
        End Set
    End Property

    Public Property MCOMP() As String
        Get
            Return _MCOMP.Trim
        End Get
        Set(ByVal value As String)
            _MCOMP = value.Trim
        End Set
    End Property

    Public Property MSITE() As String
        Get
            Return _MSITE.Trim
        End Get
        Set(ByVal value As String)
            _MSITE = value.Trim
        End Set
    End Property

    Public Property NAME() As String
        Get
            Return _NAME.Trim
        End Get
        Set(ByVal value As String)
            _NAME = value.Trim
        End Set
    End Property

    Public Property Notes() As String
        Get
            Return _Notes.Trim
        End Get
        Set(ByVal value As String)
            _Notes = value.Trim
        End Set
    End Property

    Public Property ORDDTE() As Date
        Get
            Return _ORDDTE
        End Get
        Set(ByVal value As Date)
            _ORDDTE = value
        End Set
    End Property

    Public Property ORDID() As String
        Get
            Return _ORDID.Trim
        End Get
        Set(ByVal value As String)
            _ORDID = value.Trim
        End Set
    End Property

    Public Property ORDNUM() As String
        Get
            Return _ORDNUM.Trim
        End Get
        Set(ByVal value As String)
            _ORDNUM = value.Trim
        End Set
    End Property

    Public Property PHONE() As String
        Get
            Return _PHONE
        End Get
        Set(ByVal value As String)
            _PHONE = value.Trim
        End Set
    End Property

    Public Property ProofedBy() As String
        Get
            Return _ProofedBy.Trim
        End Get
        Set(ByVal value As String)
            _ProofedBy = value.Trim
        End Set
    End Property

    Public Property RCLDTE() As Date
        Get
            Return _RCLDTE
        End Get
        Set(ByVal value As Date)
            _RCLDTE = value
        End Set
    End Property

    Public Property REASON() As String
        Get
            Return _REASON.Trim
        End Get
        Set(ByVal value As String)
            _REASON = value.Trim
        End Set
    End Property

    Public Property REP1() As String
        Get
            Return _REP1.Trim
        End Get
        Set(ByVal value As String)
            _REP1 = value.Trim
        End Set
    End Property

    Public Property REP2() As String
        Get
            Return _REP2.Trim
        End Get
        Set(ByVal value As String)
            _REP2 = value.Trim
        End Set
    End Property

    Public Property REP3() As String
        Get
            Return _REP3.Trim
        End Get
        Set(ByVal value As String)
            _REP3 = value.Trim
        End Set
    End Property

    Public Property SellarsOrderType() As String
        Get
            Return _SellarsOrderType
        End Get
        Set(ByVal value As String)
            _SellarsOrderType = value.Trim
        End Set
    End Property

    Public Property ShipFromStockID() As String
        Get
            Return _ShipFromStockID
        End Get
        Set(ByVal value As String)
            _ShipFromStockID = value.Trim
        End Set
    End Property

    Public Property SHPCDE() As String
        Get
            Return _SHPCDE.Trim
        End Get
        Set(ByVal value As String)
            _SHPCDE = value.Trim
        End Set
    End Property

    Public Property SHPVIA() As String
        Get
            Return _SHPVIA.Trim
        End Get
        Set(ByVal value As String)
            _SHPVIA = value.Trim
        End Set
    End Property

    Public Property SHPLBL() As Integer
        Get
            Return _SHPLBL
        End Get
        Set(ByVal value As Integer)
            _SHPLBL = value
        End Set
    End Property

    Public Property SHPTHRU() As String
        Get
            Return _SHPTHRU.Trim
        End Get
        Set(ByVal value As String)
            _SHPTHRU = value.Trim
        End Set
    End Property

    Public Property SPLIT1() As Single
        Get
            Return _SPLIT1
        End Get
        Set(ByVal value As Single)
            _SPLIT1 = value
        End Set
    End Property

    Public Property SPLIT2() As Single
        Get
            Return _SPLIT2
        End Get
        Set(ByVal value As Single)
            _SPLIT2 = value
        End Set
    End Property

    Public Property SPLIT3() As Single
        Get
            Return _SPLIT3
        End Get
        Set(ByVal value As Single)
            _SPLIT3 = value
        End Set
    End Property

    Public Property STATE() As String
        Get
            Return _STATE.Trim
        End Get
        Set(ByVal value As String)
            _STATE = value.Trim
        End Set
    End Property

    Public Property STATUS() As String
        Get
            Return _STATUS.Trim
        End Get
        Set(ByVal value As String)
            _STATUS = value.Trim
        End Set
    End Property

    Public Property STYPE() As String
        Get
            Return _STYPE.Trim
        End Get
        Set(ByVal value As String)
            _STYPE = value.Trim
        End Set
    End Property

    Public Property TAXABL() As String
        Get
            Return _TAXABL.Trim
        End Get
        Set(ByVal value As String)
            _TAXABL = value.Trim
        End Set
    End Property

    Public Property TAXCD1() As String
        Get
            Return _TAXCD1.Trim
        End Get
        Set(ByVal value As String)
            _TAXCD1 = value.Trim
        End Set
    End Property

    Public Property TAXCD2() As String
        Get
            Return _TAXCD2.Trim
        End Get
        Set(ByVal value As String)
            _TAXCD2 = value.Trim
        End Set
    End Property

    Public Property TAXCD3() As String
        Get
            Return _TAXCD3.Trim
        End Get
        Set(ByVal value As String)
            _TAXCD3 = value.Trim
        End Set
    End Property

    Public Property TAXPRV() As String
        Get
            Return _TAXPRV.Trim
        End Get
        Set(ByVal value As String)
            _TAXPRV = value.Trim
        End Set
    End Property

    Public Property TTAX() As Double
        Get
            Return _TTAX
        End Get
        Set(ByVal value As Double)
            _TTAX = value
        End Set
    End Property

    Public Property TERMS() As String
        Get
            Return _TERMS.Trim
        End Get
        Set(ByVal value As String)
            _TERMS = value.Trim
        End Set
    End Property

    Public Property UDFKEY() As String
        Get
            Return _UDFKEY.Trim
        End Get
        Set(ByVal value As String)
            _UDFKEY = value.Trim
        End Set
    End Property

    Public Property UDFREF() As String
        Get
            Return _UDFREF.Trim
        End Get
        Set(ByVal value As String)
            _UDFREF = value.Trim
        End Set
    End Property

    Public Property ZIPCD() As String
        Get
            Return _ZIPCD.Trim
        End Get
        Set(ByVal value As String)
            _ZIPCD = value.Trim
        End Set
    End Property

    Public Property XURR() As String
        Get
            Return _XURR.Trim
        End Get
        Set(ByVal value As String)
            _XURR = value.Trim
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
#End Region

    Public Function Add(ByVal MaxProcess As Integer, ByVal value As DataSource) As String
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        Return AddMax(MaxProcess)
    End Function

    Public Function Add(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal ADDR1 As String, ByVal ADDR2 As String, ByVal City As String, ByVal Contact As String, ByVal Country As String, ByVal CUSTID As String, ByVal CustomerName As String, ByVal CUSTPO As String, ByVal FOB As String, ByVal GLXREF As String, ByVal OrderDate As Date, ByVal OrderedBy As String, ByVal PHONE As String, ByVal REP1 As String, ByVal SHPCDE As String, ByVal SHPVIA As String, ByVal State As String, ByVal SType As String, ByVal TAXABLE As String, ByVal TAXCD1 As String, ByVal TAXCD2 As String, ByVal TAXCD3 As String, ByVal Terms As String, ByVal ZIPCD As String, ByVal Comment1 As String, ByVal Comment2 As String) As String
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        Return AddMax(MaxProcess, ADDR1, ADDR2, "", City, Contact, Country, CUSTID, CustomerName, CUSTPO, FOB, GLXREF, OrderDate, OrderedBy, PHONE, REP1, SHPCDE, SHPVIA, State, SType, TAXABLE, TAXCD1, TAXCD2, TAXCD3, Terms, ZIPCD, Comment1, Comment2)
    End Function

    Public Function Add(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal ADDR1 As String, ByVal ADDR2 As String, ByVal ADDR3 As String, ByVal City As String, ByVal Contact As String, ByVal Country As String, ByVal CUSTID As String, ByVal CustomerName As String, ByVal CUSTPO As String, ByVal FOB As String, ByVal GLXREF As String, ByVal OrderDate As Date, ByVal OrderedBy As String, ByVal PHONE As String, ByVal REP1 As String, ByVal SHPCDE As String, ByVal SHPVIA As String, ByVal State As String, ByVal SType As String, ByVal TAXABLE As String, ByVal TAXCD1 As String, ByVal TAXCD2 As String, ByVal TAXCD3 As String, ByVal Terms As String, ByVal ZIPCD As String, ByVal Comment1 As String, ByVal Comment2 As String) As String
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        Return AddMax(MaxProcess, ADDR1, ADDR2, ADDR3, City, Contact, Country, CUSTID, CustomerName, CUSTPO, FOB, GLXREF, OrderDate, OrderedBy, PHONE, REP1, SHPCDE, SHPVIA, State, SType, TAXABLE, TAXCD1, TAXCD2, TAXCD3, Terms, ZIPCD, Comment1, Comment2)
    End Function

    Private Function AddMax(ByVal MaxProcess As Integer, ByVal ADDR1 As String, ByVal ADDR2 As String, ByVal ADDR3 As String, ByVal City As String, ByVal Contact As String, ByVal Country As String, ByVal CUSTID As String, ByVal CustomerName As String, ByVal CUSTPO As String, ByVal FOB As String, ByVal GLXREF As String, ByVal OrderDate As Date, ByVal OrderedBy As String, ByVal PHONE As String, ByVal REP1 As String, ByVal SHPCDE As String, ByVal SHPVIA As String, ByVal State As String, ByVal SType As String, ByVal TAXABLE As String, ByVal TAXCD1 As String, ByVal TAXCD2 As String, ByVal TAXCD3 As String, ByVal Terms As String, ByVal ZIPCD As String, ByVal Comment1 As String, ByVal Comment2 As String) As String
        ' Define and initialize the max order master parameter structure
        Dim uSOHdr As New SOEMstItem

        ' Clear the data from the passing structure
        ClearStructure(uSOHdr)

        'Update the fields we are passing in....
        With uSOHdr
            .ADDR1 = ADDR1.Trim.PadRight(30)
            .ADDR2 = ADDR2.Trim.PadRight(30)
            .ADDR3 = ADDR3.Trim.PadRight(30)
            .CITY = City.Trim.PadRight(30)
            .CNTCT = Contact.Trim.PadRight(20)
            .CNTRY = Country.Trim.PadRight(30)
            .COMNT1 = Comment1.Trim.PadRight(30)
            .COMNT2 = Comment2.Trim.PadRight(30)
            .CURR = "USA"
            .CUSTID = CUSTID.Trim.PadRight(20)
            .NAME = CustomerName.Trim.PadRight(30)
            .CUSTPO = CUSTPO.Trim.PadRight(25)
            .FOB = FOB.Trim.PadRight(15)
            .GLXREF = GLXREF.Trim.PadRight(32)
            .INVCE = "N"
            .ORDDTE = MakeDate(Year(Now()), Month(Now()), Day(Now()))
            .ORDID = OrderedBy.Trim.PadRight(20)
            .PHONE = PHONE.Trim.PadRight(20)
            .REP1 = REP1.Trim.PadRight(7)
            .SHPCDE = SHPCDE.Trim.PadRight(10)
            .SHPVIA = SHPVIA.Trim.PadRight(2)
            .SPLIT1 = 100
            .STATE = State.Trim.PadRight(30)
            .STATUS = Convert.ToString(OrderStatus.Open)
            .STYPE = SType.Trim.PadRight(2)
            .TAXABL = TAXABLE.Trim.PadRight(1)
            .TAXCD1 = TAXCD1.Trim.PadRight(7)
            .TAXCD2 = TAXCD2.Trim.PadRight(7)
            .TAXCD3 = TAXCD3.Trim.PadRight(7)
            .TERMS = Terms.Trim.PadRight(2)
            .ZIPCD = ZIPCD.Trim.PadRight(30)
            .EXCRTE = 1
            .FIXVAR = "F"
            .LNETAX = "N"
            .FEDTAX = "N"
            ' if the created by is too long, make sure to blanks
            If CREATEDBY.Trim.Length > 100 Then
                .CREATEDBY = CREATEDBY.Trim.Substring(0, 100)
            Else
                .CREATEDBY = CREATEDBY.PadRight(100)
            End If
        End With

        'Add Sales Order Header record via MAX Update
        Dim RetOrder As String = ""
        Dim retValue As Integer = AddSalesOrder(MaxProcess, uSOHdr)
        If retValue = 0 Then
            ' Throws a new exception.
            Throw New System.Exception("Error creating sales order master record. Error " & retValue.ToString & " was returned from AddSalesOrder.")
        Else
            RetOrder = uSOHdr.ORDNUM
            ' Set the ORDNUM property to be the new order number that was generated
            ORDNUM = uSOHdr.ORDNUM
            ' if adding to max, automatically add it to the sellars sales order master table as well
            Dim newStatus As String = OrderStatus.Open
            AddSellars(RetOrder, Now(), CUSTID, newStatus, CUSTPO, FOB, CustomerName, ADDR1.Trim, ADDR2.Trim, ADDR3.Trim, City.Trim, State.Trim, ZIPCD.Trim, Country.Trim, Contact.Trim, SHPCDE.Trim)
        End If

        ' Return the sales order number
        Return RetOrder
    End Function

    Private Function AddMax(ByVal MaxProcess As Integer) As String
        ' Define and initialize the max order master parameter structure
        Dim uSOHdr As New SOEMstItem

        ' Clear the data from the passing structure
        ClearStructure(uSOHdr)

        'Update the fields from Private Variables....
        With uSOHdr
            .ADDR1 = _ADDR1.GetFixedLengthString(30)
            .ADDR2 = _ADDR2.GetFixedLengthString(30)
            .ADDR3 = _ADDR3.GetFixedLengthString(30)
            .ADDR4 = _ADDR4.GetFixedLengthString(30)
            .ADDR5 = _ADDR5.GetFixedLengthString(30)
            .ADDR6 = _ADDR6.GetFixedLengthString(30)
            .APPINV = _APPINV.GetFixedLengthString(6)
            .CITY = _CITY.GetFixedLengthString(30)
            .CNTCT = _CNTCT.GetFixedLengthString(20)
            .CNTRY = _CNTRY.GetFixedLengthString(30)
            .COMMIS = _COMMIS
            .COMNT1 = _COMNT1.GetFixedLengthString(30)
            .COMNT2 = _COMNT2.GetFixedLengthString(30)
            .COMNT3 = _COMNT3.GetFixedLengthString(30)
            .CREATEDBY = _CREATEDBY.GetFixedLengthString(100)
            .CREATIONDATE = MakeDate(Year(_CREATIONDATE), Month(_CREATIONDATE), Day(_CREATIONDATE))
            .CURR = _CURR.GetFixedLengthString(3)
            .CUSTID = _CUSTID.GetFixedLengthString(20)
            .CUSTPO = _CUSTPO.GetFixedLengthString(25)
            .FOB = _FOB.GetFixedLengthString(15)
            .GLXREF = _GLXREF.GetFixedLengthString(32)
            .INVCE = _INVCE.GetFixedLengthString(1)
            .MCOMP = _MCOMP.GetFixedLengthString(3)
            .MODIFIEDBY = _MODIFIEDBY.GetFixedLengthString(100)
            .MODIFICATIONDATE = MakeDate(Year(_MODIFICATIONDATE), Month(_MODIFICATIONDATE), Day(_MODIFICATIONDATE))
            .MSITE = _MSITE.GetFixedLengthString(3)
            .NAME = _NAME.GetFixedLengthString(30)
            .ORDDTE = MakeDate(Year(Now()), Month(Now()), Day(Now()))
            .ORDID = _ORDID.GetFixedLengthString(20)
            .PHONE = _PHONE.GetFixedLengthString(20)
            .RCLDTE = MakeDate(Year(_RCLDTE), Month(_RCLDTE), Day(_RCLDTE))
            .REASON = _REASON.GetFixedLengthString(2)
            .REP1 = _REP1.GetFixedLengthString(7)
            .REP2 = _REP2.GetFixedLengthString(7)
            .REP3 = _REP3.GetFixedLengthString(7)
            .SHPCDE = _SHPCDE.GetFixedLengthString(10)
            .SHPLBL = _SHPLBL
            .SHPTHRU = _SHPTHRU.GetFixedLengthString(10)
            .SHPVIA = _SHPVIA.GetFixedLengthString(2)
            .SPLIT1 = _SPLIT1
            .SPLIT2 = _SPLIT2
            .SPLIT3 = _SPLIT3
            .STATE = _STATE.GetFixedLengthString(30)
            .STATUS = Convert.ToString(OrderStatus.Open)
            .STYPE = _STYPE.GetFixedLengthString(2)
            .TAXABL = _TAXABL.GetFixedLengthString(1)
            .TAXCD1 = _TAXCD1.GetFixedLengthString(7)
            .TAXCD2 = _TAXCD2.GetFixedLengthString(7)
            .TAXCD3 = _TAXCD3.GetFixedLengthString(7)
            .TAXPRV = _TAXPRV.GetFixedLengthString(15)
            .TERMS = _TERMS.GetFixedLengthString(2)
            .TTAX = _TTAX
            .UDFKEY = _UDFKEY.GetFixedLengthString(15)
            .UDFREF = _UDFREF.GetFixedLengthString(25)
            .XDFBOL = _XDFBOL.GetFixedLengthString(1)
            .XDFDTE = MakeDate(Year(_XDFDTE), Month(_XDFDTE), Day(_XDFDTE))
            .XDFFLT = _XDFFLT
            .XDFINT = _XDFINT
            .XDFTXT = _XDFTXT.GetFixedLengthString(100)
            .ZIPCD = _ZIPCD.GetFixedLengthString(30)
            .EXCRTE = _EXCRTE
            .FIXVAR = _FIXVAR.GetFixedLengthString(1)
            .LNETAX = _LNETAX.GetFixedLengthString(1)
            .FEDTAX = _FEDTAX.GetFixedLengthString(1)
            .XURR = _XURR.GetFixedLengthString(2)
        End With

        'Add Sales Order Header record via MAX Update
        Dim RetOrder As String = ""
        Dim retValue As Integer = AddSalesOrder(MaxProcess, uSOHdr)

        If retValue = 0 Then
            ' Throws a new exception.
            Throw New System.Exception("Error creating sales order master record. Error " & retValue.ToString & " was returned from AddSalesOrder.")
        Else
            RetOrder = uSOHdr.ORDNUM
            ' Set the ORDNUM property to be the new order number that was generated
            ORDNUM = uSOHdr.ORDNUM
            ' if adding to max, automatically add it to the sellars sales order master table as well
            Dim newStatus As String = OrderStatus.Open
            AddSellars(RetOrder, Now(), _CUSTID, newStatus, _CUSTPO, _FOB, _NAME, _ADDR1.Trim, _ADDR2.Trim, _ADDR3.Trim, _CITY.Trim, _STATE.Trim, _ZIPCD.Trim, _CNTRY.Trim, _CNTCT.Trim, _SHPCDE.Trim)
        End If

        ' Return the sales order number
        Return RetOrder

    End Function

    ' The following routine will add a sales order to the sellars SQL sales order table
    Private Sub AddSellars(ByVal SalesOrder As String, ByVal OrderDate As Date, ByVal CustomerID As String, ByVal Status As String, ByVal CustPO As String, ByVal FOB As String, ByVal Name As String, ByVal Addr1 As String, ByVal Addr2 As String, ByVal Addr3 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, ByVal Contact As String, ByVal ShipCode As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, SalesOrder, ParameterDirection.Input)
        oSQL.AddParameter("@ORDDTE", SqlDbType.DateTime, 0, OrderDate, ParameterDirection.Input)
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CustomerID, ParameterDirection.Input)
        oSQL.AddParameter("@STATUS", SqlDbType.NVarChar, 1, Status, ParameterDirection.Input)
        oSQL.AddParameter("@CUSTPO", SqlDbType.NVarChar, 25, CustPO, ParameterDirection.Input)
        oSQL.AddParameter("@FOB", SqlDbType.NVarChar, 15, FOB, ParameterDirection.Input)
        oSQL.AddParameter("@NAME", SqlDbType.NVarChar, 30, Name, ParameterDirection.Input)
        oSQL.AddParameter("@ADDR1", SqlDbType.NVarChar, 30, Addr1, ParameterDirection.Input)
        oSQL.AddParameter("@ADDR2", SqlDbType.NVarChar, 30, Addr2, ParameterDirection.Input)
        oSQL.AddParameter("@ADDR3", SqlDbType.NVarChar, 30, Addr3, ParameterDirection.Input)
        oSQL.AddParameter("@CITY", SqlDbType.NVarChar, 30, City, ParameterDirection.Input)
        oSQL.AddParameter("@STATE", SqlDbType.NVarChar, 30, State, ParameterDirection.Input)
        oSQL.AddParameter("@ZIPCD", SqlDbType.NVarChar, 30, Zip, ParameterDirection.Input)
        oSQL.AddParameter("@CNTRY", SqlDbType.NVarChar, 20, Country, ParameterDirection.Input)
        oSQL.AddParameter("@CNTCT", SqlDbType.NVarChar, 20, Contact, ParameterDirection.Input)
        oSQL.AddParameter("@SHPCDE", SqlDbType.NVarChar, 10, ShipCode, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSalesOrderMaster")
    End Sub

    Public Function Clone(ByVal MaxProcess As Integer, ByVal pOrder As String) As String
        ' Read in the existing order
        Read(pOrder)

        ' Define and initialize the max order master parameter structure
        Dim uSOHdr As New SOEMstItem

        ' Clear the data from the passing structure
        ClearStructure(uSOHdr)

        Dim blanks As String = ""

        'Update the fields from Private Variables....
        With uSOHdr
            .ADDR1 = _ADDR1.GetFixedLengthString(30)
            .ADDR2 = _ADDR2.GetFixedLengthString(30)
            .ADDR3 = _ADDR3.GetFixedLengthString(30)
            .ADDR4 = _ADDR4.GetFixedLengthString(30)
            .ADDR5 = _ADDR5.GetFixedLengthString(30)
            .ADDR6 = _ADDR6.GetFixedLengthString(30)
            .APPINV = blanks.GetFixedLengthString(6)
            .CITY = _CITY.GetFixedLengthString(30)
            .CNTCT = _CNTCT.GetFixedLengthString(20)
            .CNTRY = _CNTRY.GetFixedLengthString(30)
            .COMMIS = _COMMIS
            .COMNT1 = _COMNT1.GetFixedLengthString(30)
            .COMNT2 = _COMNT2.GetFixedLengthString(30)
            .COMNT3 = _COMNT3.GetFixedLengthString(30)
            .CREATEDBY = _CREATEDBY.GetFixedLengthString(100)
            .CREATIONDATE = MakeDate(Year(Now()), Month(Now()), Day(Now()))
            .CURR = _CURR.GetFixedLengthString(3)
            .CUSTID = _CUSTID.GetFixedLengthString(20)
            .CUSTPO = _CUSTPO.GetFixedLengthString(25)
            .FOB = _FOB.GetFixedLengthString(15)
            .GLXREF = _GLXREF.GetFixedLengthString(32)
            .INVCE = "N"
            .MCOMP = _MCOMP.GetFixedLengthString(3)
            .MODIFIEDBY = _MODIFIEDBY.GetFixedLengthString(100)
            .MODIFICATIONDATE = MakeDate(Year(Now()), Month(Now()), Day(Now()))
            .MSITE = _MSITE.GetFixedLengthString(3)
            .NAME = _NAME.GetFixedLengthString(30)
            .ORDDTE = MakeDate(Year(Now()), Month(Now()), Day(Now()))
            .ORDID = _ORDID.GetFixedLengthString(20)
            .PHONE = _PHONE.GetFixedLengthString(20)
            .RCLDTE = MakeDate(Year(DefaultDate), Month(DefaultDate), Day(DefaultDate))
            .REASON = _REASON.GetFixedLengthString(2)
            .REP1 = _REP1.GetFixedLengthString(7)
            .REP2 = _REP2.GetFixedLengthString(7)
            .REP3 = _REP3.GetFixedLengthString(7)
            .SHPCDE = _SHPCDE.GetFixedLengthString(10)
            .SHPLBL = _SHPLBL
            .SHPTHRU = _SHPTHRU.GetFixedLengthString(10)
            .SHPVIA = _SHPVIA.GetFixedLengthString(2)
            .SPLIT1 = _SPLIT1
            .SPLIT2 = _SPLIT2
            .SPLIT3 = _SPLIT3
            .STATE = _STATE.GetFixedLengthString(30)
            .STATUS = Convert.ToString(OrderStatus.Open)
            .STYPE = _STYPE.GetFixedLengthString(2)
            .TAXABL = _TAXABL.GetFixedLengthString(1)
            .TAXCD1 = _TAXCD1.GetFixedLengthString(7)
            .TAXCD2 = _TAXCD2.GetFixedLengthString(7)
            .TAXCD3 = _TAXCD3.GetFixedLengthString(7)
            .TAXPRV = _TAXPRV.GetFixedLengthString(15)
            .TERMS = _TERMS.GetFixedLengthString(2)
            .TTAX = _TTAX
            .UDFKEY = _UDFKEY.GetFixedLengthString(15)
            .UDFREF = _UDFREF.GetFixedLengthString(25)
            .XDFBOL = _XDFBOL.GetFixedLengthString(1)
            .XDFDTE = MakeDate(Year(DefaultDate), Month(DefaultDate), Day(DefaultDate))
            .XDFFLT = _XDFFLT
            .XDFINT = _XDFINT
            .XDFTXT = _XDFTXT.GetFixedLengthString(100)
            .ZIPCD = _ZIPCD.GetFixedLengthString(30)
            .EXCRTE = _EXCRTE
            .FIXVAR = _FIXVAR.GetFixedLengthString(1)
            .LNETAX = _LNETAX.GetFixedLengthString(1)
            .FEDTAX = _FEDTAX.GetFixedLengthString(1)
            .XURR = _XURR.GetFixedLengthString(2)
        End With

        'Add Sales Order Header record via MAX Update
        Dim RetOrder As String = ""
        Dim retValue As Integer = AddSalesOrder(MaxProcess, uSOHdr)

        If retValue = 0 Then
            ' Throws a new exception.
            Throw New System.Exception("Error cloning sales order master record. Error " & retValue.ToString & " was returned from AddSalesOrder.")
        Else
            RetOrder = uSOHdr.ORDNUM
            ' Set the ORDNUM property to be the new order number that was generated
            ORDNUM = uSOHdr.ORDNUM
            ' if adding to max, automatically add it to the sellars sales order master table as well
            Dim newStatus As String = OrderStatus.Open
            AddSellars(RetOrder, Now(), _CUSTID, newStatus, _CUSTPO, _FOB, _NAME, _ADDR1.Trim, _ADDR2.Trim, _ADDR3.Trim, _CITY.Trim, _STATE.Trim, _ZIPCD.Trim, _CNTRY.Trim, _CNTCT.Trim, _SHPCDE.Trim)
        End If

        ' Return the sales order number
        Return RetOrder
    End Function

    Public Sub Delete(ByVal MaxProcess As Integer, ByVal pOrder As String)
        'delete the Sales Order Detail record via MAX Update
        Dim retValue As Short = DeleteSalesOrder(MaxProcess, pOrder)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error deleting sales order detail.")
            Case Else
                ' delete the record from the Sellars SQL sales order detail, and so detail extenstion table
                DeleteSellars(pOrder)
        End Select
    End Sub

    '*********************************************************************
    ' Delete()
    ' Deletes the sales order master record for a customer
    '*********************************************************************
    Private Sub DeleteSellars(ByVal pOrder As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, pOrder, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("DeleteSalesOrder")
    End Sub

    Public Function Read(ByVal StartOrder As String, ByVal EndOrder As String) As SqlDataReader
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim strSQL As String = "Select ORDNUM_27 " & _
                               "From ""SO_Master"" " & _
                               "Where ORDNUM_27 between '" & StartOrder.Trim & "' and '" & EndOrder.Trim & "'"

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim OrderMasterReader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

        ' Return the recordset
        Return OrderMasterReader
    End Function

    Public Sub Read(ByVal passorder As String, ByVal OrderType As OrderTypes)

        ' First check if the passed in order contains the type of data the application can handle
        If CheckOrder(passorder, OrderType) Then
            Read(passorder)
        Else
            Throw New ApplicationException("Order is the wrong type for this application.")
        End If

    End Sub

    Private Function CheckOrder(ByVal ORDNUM As String, ByVal OrderType As OrderTypes) As Boolean
        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnVal As Boolean = False

        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterClass:CheckOrder():Order number is empty.")
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Define a variable for the database name so the application automatically sets up the SQL string below with the correct table names for test versus production
        Dim dataBaseName As String = IIf(System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString.ToUpper.Contains("TEST"), "TestShopfloorControl", "ShopfloorControl")

        ' Check if there are any roll goods parts on the order
        Dim strSQL As String = "select @Count = count(LINNUM_28) from SO_Detail join Part_Master on PRTNUM_01 = PRTNUM_28 where ORDNUM_28 = @ORDNUM and PLANID_01 in (select PlanId from ShopfloorControl..OrderTypePlanIds where OrderType = @OrderType); select @ItemsCount = count(*) from SO_Detail where ORDNUM_28  = @ORDNUM; select @SType = STYPE_27 from SO_Master where ORDNUM_27 = @ORDNUM; select @ItemMasterCount = count(*) from SO_Detail join Part_Master on PRTNUM_01 = PRTNUM_28 where ORDNUM_28 = @ORDNUM;"

        Dim cmd As SqlCommand = Nothing
        Dim dr As SqlDataReader = Nothing

        Try
            ' Set up a new SQL Command
            cmd = New SqlCommand(strSQL, MaxConnection)
            cmd.CommandType = CommandType.Text
            cmd.CommandTimeout = 0

            ' Add a parameter to the command that has the invoice number to update
            cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
            cmd.Parameters.Add(New SqlParameter("@OrderType", OrderType))

            ' Add the parameter to return the count of lines that have a valid plan id
            Dim prmCount As New SqlParameter("@Count", SqlDbType.SmallInt, 0)
            prmCount.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmCount)

            ' Add the parameter to return the count of lines that have a valid plan id
            Dim prmType As New SqlParameter("@SType", SqlDbType.NVarChar, 2)
            prmType.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmType)

            ' Add the parameter to return the count of lines that have a valid plan id
            Dim prmItemsCount As New SqlParameter("@ItemsCount", SqlDbType.SmallInt, 0)
            prmItemsCount.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmItemsCount)

            ' Add the parameter to return the count of lines that have a valid plan id
            Dim prmItemMasterCount As New SqlParameter("@ItemMasterCount", SqlDbType.SmallInt, 0)
            prmItemMasterCount.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmItemMasterCount)

            ' Execute the stored procedure to update the InvoiceMasterExt table
            cmd.ExecuteNonQuery()

            ' Set the return flag to true
            If Convert.ToInt16(prmCount.Value) > 0 Or prmType.Value.ToString = "CR" Or Convert.ToInt16(prmItemsCount.Value) = 0 Or Convert.ToInt16(prmItemMasterCount.Value) = 0 Then
                rtnVal = True
            End If
        Catch ex As Exception
            Dim msg As String = ex.Message
        End Try

        If cmd IsNot Nothing Then
            cmd.Dispose()
        End If

        Return rtnVal
    End Function

    Public Sub Read(ByVal passorder As String)

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim CustomerPrice As Decimal = 0
        Dim strSQL As String = "Select * " & _
                               "From ""SO_Master"" " & _
                               "Where ORDNUM_27 = '" & passorder.Trim & "' "

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim OrderMasterReader As SqlDataReader = cmd.ExecuteReader()

        _ORDNUM = passorder
        If OrderMasterReader.Read() Then
            _found = True
            _CUSTID = RecordString(OrderMasterReader("CUSTID_27"))
            _GLXREF = RecordString(OrderMasterReader("GLXREF_27"))
            _STYPE = RecordString(OrderMasterReader("STYPE_27"))
            _STATUS = RecordString(OrderMasterReader("STATUS_27"))
            _CUSTPO = RecordString(OrderMasterReader("CUSTPO_27"))
            _ORDID = RecordString(OrderMasterReader("ORDID_27"))
            Try
                _ORDDTE = dbRecordDate(OrderMasterReader("ORDDTE_27"))
            Catch
                _ORDDTE = DefaultDate
            End Try
            _FILL01 = RecordString(OrderMasterReader("FILL01_27"))
            _SHPCDE = RecordString(OrderMasterReader("SHPCDE_27"))
            _REP1 = RecordString(OrderMasterReader("REP1_27"))
            _SPLIT1 = dbRecordSingle(OrderMasterReader("SPLIT1_27"))
            _REP2 = RecordString(OrderMasterReader("REP2_27"))
            _SPLIT2 = dbRecordSingle(OrderMasterReader("SPLIT2_27"))
            _REP3 = RecordString(OrderMasterReader("REP3_27"))
            _SPLIT3 = dbRecordSingle(OrderMasterReader("SPLIT3_27"))
            _COMMIS = dbRecordSingle(OrderMasterReader("COMMIS_27"))
            _TERMS = RecordString(OrderMasterReader("TERMS_27"))
            _SHPVIA = RecordString(OrderMasterReader("SHPVIA_27"))
            _XURR = RecordString(OrderMasterReader("XURR_27"))
            _FOB = RecordString(OrderMasterReader("FOB_27"))
            _TAXCD1 = RecordString(OrderMasterReader("TAXCD1_27"))
            _TAXCD2 = RecordString(OrderMasterReader("TAXCD2_27"))
            _TAXCD3 = RecordString(OrderMasterReader("TAXCD3_27"))
            _COMNT1 = RecordString(OrderMasterReader("COMNT1_27"))
            _COMNT2 = RecordString(OrderMasterReader("COMNT2_27"))
            _COMNT3 = RecordString(OrderMasterReader("COMNT3_27"))
            _SHPLBL = dbRecordInt(OrderMasterReader("SHPLBL_27"))
            _INVCE = RecordString(OrderMasterReader("INVCE_27"))
            _APPINV = RecordString(OrderMasterReader("APPINV_27"))
            _REASON = RecordString(OrderMasterReader("REASON_27"))
            _NAME = RecordString(OrderMasterReader("NAME_27"))
            _ADDR1 = RecordString(OrderMasterReader("ADDR1_27"))
            _ADDR2 = RecordString(OrderMasterReader("ADDR2_27"))
            _CITY = RecordString(OrderMasterReader("CITY_27"))
            _STATE = RecordString(OrderMasterReader("STATE_27"))
            _ZIPCD = RecordString(OrderMasterReader("ZIPCD_27"))
            _CNTRY = RecordString(OrderMasterReader("CNTRY_27"))
            _PHONE = RecordString(OrderMasterReader("PHONE_27"))
            _CNTCT = RecordString(OrderMasterReader("CNTCT_27"))
            _TAXPRV = RecordString(OrderMasterReader("TAXPRV_27"))
            _FEDTAX = RecordString(OrderMasterReader("FEDTAX_27"))
            _TAXABL = RecordString(OrderMasterReader("TAXABL_27"))
            _EXCRTE = dbRecordSingle(OrderMasterReader("EXCRTE_27"))
            _FIXVAR = RecordString(OrderMasterReader("FIXVAR_27"))
            _CURR = RecordString(OrderMasterReader("CURR_27"))
            ' Because we can get bad dates in the recall date field
            ' first try to read it, and if we can't set it to the default
            Try
                _RCLDTE = dbRecordDate(OrderMasterReader("RCLDTE_27"))
            Catch
                _RCLDTE = DefaultDate
            End Try
            _FILL02 = RecordString(OrderMasterReader("FILL02_27"))
            _TTAX = dbRecordDouble(OrderMasterReader("TTAX_27"))
            _LNETAX = RecordString(OrderMasterReader("LNETAX_27"))
            _ADDR3 = RecordString(OrderMasterReader("ADDR3_27"))
            _ADDR4 = RecordString(OrderMasterReader("ADDR4_27"))
            _ADDR5 = RecordString(OrderMasterReader("ADDR5_27"))
            _ADDR6 = RecordString(OrderMasterReader("ADDR6_27"))
            _MCOMP = RecordString(OrderMasterReader("MCOMP_27"))
            _MSITE = RecordString(OrderMasterReader("MSITE_27"))
            _UDFKEY = RecordString(OrderMasterReader("UDFKEY_27"))
            _UDFREF = RecordString(OrderMasterReader("UDFREF_27"))
            _SHPTHRU = RecordString(OrderMasterReader("SHPTHRU_27"))
            _FILLER = RecordString(OrderMasterReader("FILLER_27"))
            _XDFINT = RecordString(OrderMasterReader("XDFINT_27"))
            _XDFFLT = RecordString(OrderMasterReader("XDFFLT_27"))
            _XDFBOL = RecordString(OrderMasterReader("XDFBOL_27"))
            ' Because we can get bad dates in the recall date field
            ' first try to read it, and if we can't set it to the default
            Try
                _XDFDTE = dbRecordDate(OrderMasterReader("XDFDTE_27"))
            Catch
                _XDFDTE = DefaultDate
            End Try
            _CREATEDBY = RecordString(OrderMasterReader("CreatedBy"))
            _CREATIONDATE = dbRecordDate(OrderMasterReader("CreationDate"))
            _MODIFIEDBY = RecordString(OrderMasterReader("ModifiedBy"))
            _MODIFICATIONDATE = dbRecordDate(OrderMasterReader("ModificationDate"))
        Else
            _found = False
            _CUSTID = ""
            _GLXREF = ""
            _STYPE = ""
            _STATUS = ""
            _CUSTPO = ""
            _ORDID = ""
            _ORDDTE = DefaultDate
            _FILL01 = ""
            _SHPCDE = ""
            _REP1 = ""
            _SPLIT1 = 0
            _REP2 = ""
            _SPLIT2 = 0
            _REP3 = ""
            _SPLIT3 = 0
            _COMMIS = 0
            _TERMS = ""
            _SHPVIA = ""
            _XURR = ""
            _FOB = ""
            _TAXCD1 = ""
            _TAXCD2 = ""
            _TAXCD3 = ""
            _COMNT1 = ""
            _COMNT2 = ""
            _COMNT3 = ""
            _SHPLBL = 0
            _INVCE = ""
            _APPINV = ""
            _REASON = ""
            _NAME = ""
            _ADDR1 = ""
            _ADDR2 = ""
            _CITY = ""
            _STATE = ""
            _ZIPCD = ""
            _CNTRY = ""
            _PHONE = ""
            _CNTCT = ""
            _TAXPRV = ""
            _FEDTAX = ""
            _TAXABL = ""
            _EXCRTE = 0
            _FIXVAR = ""
            _CURR = ""
            _RCLDTE = DefaultDate
            _FILL02 = ""
            _TTAX = 0
            _LNETAX = ""
            _ADDR3 = ""
            _ADDR4 = ""
            _ADDR5 = ""
            _ADDR6 = ""
            _MCOMP = ""
            _MSITE = ""
            _UDFKEY = ""
            _UDFREF = ""
            _SHPTHRU = ""
            _FILLER = ""
            _XDFINT = 0
            _XDFFLT = 0
            _XDFBOL = ""
            _XDFDTE = DefaultDate
            _CREATEDBY = ""
            _CREATIONDATE = DefaultDate
            _MODIFIEDBY = ""
            _MODIFICATIONDATE = DefaultDate
        End If

        OrderMasterReader.Close()
        OrderMasterReader = Nothing

        CloseMaxConnection()

        If _found Then
            ' Now get the items from the OrderMasterEXT table
            Dim ext As New SOMasterExtClass(passorder.Trim)
            _EnteredBy = ext.EnteredBy
            _LastChanged = ext.LastChanged
            _LastPrinted = ext.LastPrinted
            _ProofedBy = ext.ProofedBy
            _Notes = ext.Notes
            _Hold = ext.Hold
            _HoldUserEmail = ext.HoldUserEmail
            _DefaultStockID = ext.DefaultStockID
            _ShipFromStockID = ext.ShipFromStockID
            _SellarsOrderType = ext.OrderType
            _Finished = ext.Finished
            _EstimatedShipping = ext.EstimatedShipping
            _AllowFinish = ext.AllowFinish
            ext = Nothing
        Else
            _EnteredBy = ""
            _LastChanged = DefaultDate
            _LastPrinted = DefaultDate
            _ProofedBy = ""
            _Notes = ""
            _Hold = False
            _HoldUserEmail = ""
            _DefaultStockID = ""
            _ShipFromStockID = ""
            _SellarsOrderType = ""
            _EstimatedShipping = 0
            _AllowFinish = False
            _Finished = False
        End If

    End Sub

    Private Function RecordString(ByVal pField As String) As String
        If IsDBNull(pField) Then
            Return ""
        Else
            Return pField.Trim
        End If
    End Function

    Private Function dbRecordDate(Optional ByVal pField As Object = Nothing) As Date
        If IsDBNull(pField) Then
            Return DefaultDate
        Else
            Return CDate(pField)
        End If
    End Function

    Private Function dbRecordSingle(ByVal pField As Single) As Single
        If IsDBNull(pField) Then
            Return 0
        Else
            Return pField
        End If
    End Function

    Private Function dbRecordDouble(ByVal pField As Double) As Double
        If IsDBNull(pField) Then
            Return 0
        Else
            Return pField
        End If
    End Function

    Private Function dbRecordInt(ByVal pField As Integer) As Integer
        If IsDBNull(pField) Then
            Return 0
        Else
            Return pField
        End If
    End Function

    Public Sub Update(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal pORDNUM As String, ByVal pClosed As Boolean)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        ' Declare necessary local variables and initialize them
        Dim newStatus As String
        If pClosed Then
            newStatus = OrderStatus.Closed
        Else
            newStatus = OrderStatus.Open
        End If

        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .STATUS = newStatus
        End With

        'Call the MAXUPDATE ChangeSalesOrder function
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' update the Sellars SQL sales order table right away as well
                UpdateSellars(pORDNUM, newStatus)
        End Select
    End Sub

    Public Sub UpdateCustID(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal pORDNUM As String, ByVal CustID As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .CUSTID = CustID.GetFixedLengthString(20)
        End With

        'Call the MAXUPDATE ChangeSalesOrder function
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' update the Sellars SQL sales order table right away as well
                UpdateSellarsCustomerID(pORDNUM, CustID)
        End Select
    End Sub

    Public Sub UpdateComments(ByVal MaxProcess As Integer, ByVal pORDNUM As String, ByVal pCOMNT1 As String, ByVal pCOMNT2 As String, ByVal pCOMNT3 As String)
        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .COMNT1 = pCOMNT1.Trim.PadRight(30)
            .COMNT2 = pCOMNT2.Trim.PadRight(30)
            .COMNT3 = pCOMNT3.Trim.PadRight(30)
        End With

        'Change the sales order master via MAX Update
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select
    End Sub

    ' The following routine will update a sales order detail record on the sellars SQL sales order table
    Private Sub Update(ByVal SalesOrder As String, ByVal Line As String, ByVal Delivery As String, ByVal Status As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, SalesOrder, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, Line, ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, Delivery, ParameterDirection.Input)
        oSQL.AddParameter("@STATUS", SqlDbType.NVarChar, 1, Status, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSalesOrderDetailStatus")

        ' Call the routine to update the order status in the Sellars SQL Sales Order Master Table
        UpdateSellars(ORDNUM, Status)
    End Sub

    ' The following routine will update a sales order master record's status in the sellars SQL sales order master table
    Private Sub UpdateSellars(ByVal SalesOrder As String, ByVal Status As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, SalesOrder, ParameterDirection.Input)
        oSQL.AddParameter("@STATUS", SqlDbType.NVarChar, 1, Status, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSalesOrderMasterStatus")
    End Sub

    ' The following routine will update a sales order master record's status in the sellars SQL sales order master table
    Private Sub UpdateSellarsCustomerID(ByVal SalesOrder As String, ByVal CustomerID As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, SalesOrder, ParameterDirection.Input)
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CustomerID, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSalesOrderMasterCustomer")
    End Sub

    ' Function used to update any updateable line item field
    Public Sub Update(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal pORDNUM As String, ByVal pADDR1 As String, ByVal pADDR2 As String, ByVal pCity As String, ByVal pCountry As String, ByVal pCustomerName As String, ByVal pCUSTPO As String, ByVal pFOB As String, ByVal pOrderDate As Date, ByVal pOrderedBy As String, ByVal pShipCode As String, ByVal pSHPVIA As String, ByVal pState As String, ByVal pTAXCD1 As String, ByVal pTAXCD2 As String, ByVal pTAXCD3 As String, ByVal pTerms As String, ByVal pZIPCD As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        Dim newSHPCDE As String
        If pShipCode = "NoShipCode" Then
            newSHPCDE = ""
        Else
            newSHPCDE = pShipCode
        End If

        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .SHPCDE = newSHPCDE.PadRight(10)
            .NAME = pCustomerName.PadRight(30)
            .ADDR1 = pADDR1.PadRight(30)
            .ADDR2 = pADDR2.PadRight(30)
            .CITY = pCity.PadRight(30)
            .STATE = pState.PadRight(30)
            .ZIPCD = pZIPCD.PadRight(30)
            .CNTRY = pCountry.PadRight(30)
            .TAXCD1 = pTAXCD1.PadRight(7)
            .TAXCD2 = pTAXCD2.PadRight(7)
            .TAXCD3 = pTAXCD3.PadRight(7)
            .CUSTPO = pCUSTPO.PadRight(25)
            .FOB = pFOB.PadRight(15)
            .ORDDTE = MakeDate(pOrderDate.Year, pOrderDate.Month, pOrderDate.Day)
            .ORDID = pOrderedBy.PadRight(20)
            .SHPVIA = pSHPVIA.PadRight(2)
            .TERMS = pTerms.PadRight(2)
        End With

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' Update the sellars sales order master right away
                UpdateSellars(pORDNUM, pOrderDate, pCUSTPO, pFOB, pCustomerName, pADDR1.Trim, pADDR2.Trim, "", pCity.Trim, pState.Trim, pZIPCD.Trim, pCountry.Trim, newSHPCDE)
        End Select

    End Sub

    ' Function used to update any updateable line item field
    Public Sub Update(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal pORDNUM As String, ByVal pADDR1 As String, ByVal pADDR2 As String, ByVal pADDR3 As String, ByVal pCity As String, ByVal pCountry As String, ByVal pCustomerName As String, ByVal pCUSTPO As String, ByVal pFOB As String, ByVal pOrderDate As Date, ByVal pOrderedBy As String, ByVal pShipCode As String, ByVal pSHPVIA As String, ByVal pState As String, ByVal pTAXCD1 As String, ByVal pTAXCD2 As String, ByVal pTAXCD3 As String, ByVal pTerms As String, ByVal pZIPCD As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        Dim newSHPCDE As String
        If pShipCode = "NoShipCode" Then
            newSHPCDE = ""
        Else
            newSHPCDE = pShipCode
        End If

        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .SHPCDE = newSHPCDE.PadRight(10)
            .NAME = pCustomerName.PadRight(30)
            .ADDR1 = pADDR1.PadRight(30)
            .ADDR2 = pADDR2.PadRight(30)
            .ADDR3 = pADDR3.PadRight(30)
            .CITY = pCity.PadRight(30)
            .STATE = pState.PadRight(30)
            .ZIPCD = pZIPCD.PadRight(30)
            .CNTRY = pCountry.PadRight(30)
            .TAXCD1 = pTAXCD1.PadRight(7)
            .TAXCD2 = pTAXCD2.PadRight(7)
            .TAXCD3 = pTAXCD3.PadRight(7)
            .CUSTPO = pCUSTPO.PadRight(25)
            .FOB = pFOB.PadRight(15)
            .ORDDTE = MakeDate(pOrderDate.Year, pOrderDate.Month, pOrderDate.Day)
            .ORDID = pOrderedBy.PadRight(20)
            .SHPVIA = pSHPVIA.PadRight(2)
            .TERMS = pTerms.PadRight(2)
        End With

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' Update the sellars sales order master right away
                UpdateSellars(pORDNUM, pOrderDate, pCUSTPO, pFOB, pCustomerName, pADDR1.Trim, pADDR2.Trim, pADDR3.Trim, pCity.Trim, pState.Trim, pZIPCD.Trim, pCountry.Trim, newSHPCDE)
        End Select

    End Sub
    ' Function used to update any updateable line item field
    Public Sub UpdateFOB(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal pORDNUM As String, ByVal FOB As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .FOB = FOB.PadRight(15)
        End With

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
                ' Update the sellars sales order master right away
                ' UpdateSellars(pORDNUM, pOrderDate, pCUSTPO, pFOB, pCustomerName, pADDR1.Trim, pADDR2.Trim, pADDR3.Trim, pCity.Trim, pState.Trim, pZIPCD.Trim, pCountry.Trim, newSHPCDE)
        End Select

    End Sub

    ' The following routine will update a sales order master record in the sellars SQL sales order master table
    Private Sub UpdateSellars(ByVal SalesOrder As String, ByVal OrderDate As Date, ByVal CustPO As String, ByVal FOB As String, ByVal Name As String, ByVal Addr1 As String, ByVal Addr2 As String, ByVal Addr3 As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal Country As String, ByVal ShipCode As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, SalesOrder, ParameterDirection.Input)
        oSQL.AddParameter("@ORDDTE", SqlDbType.DateTime, 0, OrderDate, ParameterDirection.Input)
        oSQL.AddParameter("@CUSTPO", SqlDbType.NVarChar, 25, CustPO, ParameterDirection.Input)
        oSQL.AddParameter("@FOB", SqlDbType.NVarChar, 15, FOB, ParameterDirection.Input)
        oSQL.AddParameter("@NAME", SqlDbType.NVarChar, 30, Name, ParameterDirection.Input)
        oSQL.AddParameter("@ADDR1", SqlDbType.NVarChar, 30, Addr1, ParameterDirection.Input)
        oSQL.AddParameter("@ADDR2", SqlDbType.NVarChar, 30, Addr2, ParameterDirection.Input)
        oSQL.AddParameter("@ADDR3", SqlDbType.NVarChar, 30, Addr3, ParameterDirection.Input)
        oSQL.AddParameter("@CITY", SqlDbType.NVarChar, 30, City, ParameterDirection.Input)
        oSQL.AddParameter("@STATE", SqlDbType.NVarChar, 30, State, ParameterDirection.Input)
        oSQL.AddParameter("@ZIPCD", SqlDbType.NVarChar, 30, Zip, ParameterDirection.Input)
        oSQL.AddParameter("@CNTRY", SqlDbType.NVarChar, 30, Country, ParameterDirection.Input)
        oSQL.AddParameter("@SHPCDE", SqlDbType.NVarChar, 10, ShipCode, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSalesOrderMaster")
    End Sub

    ' Function used to update any updateable line item field
    Public Sub UpdateTotalTax(ByVal MaxProcess As Integer, ByVal value As DataSource, ByVal pORDNUM As String)
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Sub
        End If

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim tmpTotalTax As Decimal = 0
        Dim strSQL As String = "Select sum(TAX1_28) as TAX1, Sum(TAX2_28) as TAX2, Sum(TAX3_28) as TAX3 " & _
                               "From ""SO_Detail"" " & _
                               "Where ORDNUM_28 = '" & pORDNUM.Trim & "' "

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim OrderMasterReader As SqlDataReader = cmd.ExecuteReader()
        If OrderMasterReader.Read() Then
            tmpTotalTax = OrderMasterReader("TAX1") + OrderMasterReader("TAX2") + OrderMasterReader("TAX3")
        Else
            tmpTotalTax = 0
        End If
        ' Free up memory and close open objects
        OrderMasterReader.Close()
        OrderMasterReader = Nothing
        CloseMaxConnection()

        ' Declare the necessary internal variables
        Dim SOData As New SOEMstItem

        ' Read the requested sales order detail record
        Read(pORDNUM)

        ' Fill the old and new structures with data from the record
        FillStructure(SOData)

        'Fill in the change functions new data structure by first setting the 
        'data to the original data record, and then modifying the fields that have changed!
        With SOData
            .TTAX = tmpTotalTax
        End With

        'Add Sales Order Detail record via MAX Update
        Dim retValue As Integer = ChangeSalesOrder(MaxProcess, SOData)
        Select Case retValue
            Case 0
                ' Throws a new exception.
                Throw New System.Exception("Error updating sales order detail.")
            Case Else
        End Select

    End Sub

    Private Sub ClearStructure(ByRef passStruct As SOEMstItem)

        'Initialize all the strings to the proper length
        With passStruct
            .ADDR1 = New String(" ", 30)
            .ADDR2 = New String(" ", 30)
            .ADDR3 = New String(" ", 30)
            .ADDR4 = New String(" ", 30)
            .ADDR5 = New String(" ", 30)
            .ADDR6 = New String(" ", 30)
            .APPINV = New String(" ", 6)
            .CITY = New String(" ", 30)
            .CNTCT = New String(" ", 20)
            .CNTRY = New String(" ", 30)
            .COMMIS = 0
            .COMNT1 = New String(" ", 30)
            .COMNT2 = New String(" ", 30)
            .COMNT3 = New String(" ", 30)
            .CREATEDBY = New String(" ", 100)
            .CREATIONDATE = 0
            .CURR = New String(" ", 3)
            .CUSTID = New String(" ", 20)
            .CUSTPO = New String(" ", 25)
            .EXCRTE = 0
            .FEDTAX = New String("N", 1)
            .FILL01 = New String(" ", 2)
            .FILL02 = New String(" ", 2)
            .FILLER = New String(" ", 50)
            .FIXVAR = New String(" ", 1)
            .FOB = New String(" ", 15)
            .GLXREF = New String(" ", 32)
            .INVCE = New String(" ", 1)
            .LNETAX = New String("N", 1)
            .MCOMP = New String(" ", 3)
            .MODIFICATIONDATE = 0
            .MODIFIEDBY = New String(" ", 100)
            .MSITE = New String(" ", 3)
            .NAME = New String(" ", 30)
            .ORDDTE = 0
            .ORDID = New String(" ", 20)
            .ORDNUM = New String(" ", 8)
            .PHONE = New String(" ", 20)
            .RCLDTE = 0
            .REASON = New String(" ", 2)
            .REP1 = New String(" ", 7)
            .REP2 = New String(" ", 7)
            .REP3 = New String(" ", 7)
            .SHPCDE = New String(" ", 10)
            .SHPLBL = 0
            .SHPTHRU = New String(" ", 10)
            .SHPVIA = New String(" ", 2)
            .SPLIT1 = 0
            .SPLIT2 = 0
            .SPLIT3 = 0
            .STATE = New String(" ", 30)
            .STATUS = New String(" ", 1)
            .STYPE = New String(" ", 2)
            .TAXABL = New String(" ", 1)
            .TAXCD1 = New String(" ", 7)
            .TAXCD2 = New String(" ", 7)
            .TAXCD3 = New String(" ", 7)
            .TAXPRV = New String(" ", 15)
            .TERMS = New String(" ", 2)
            .TTAX = 0
            .UDFKEY = New String(" ", 15)
            .UDFREF = New String(" ", 25)
            .XDFBOL = New String(" ", 1)
            .XDFDTE = 0
            .XDFFLT = 0
            .XDFINT = 0
            .XDFTXT = New String(" ", 100)
            .XURR = New String(" ", 2)
            .ZIPCD = New String(" ", 30)
        End With
    End Sub

    Private Sub FillStructure(ByRef SOData As SOEMstItem)
        With SOData
            .ADDR1 = ADDR1.PadRight(30)
            .ADDR2 = ADDR2.PadRight(30)
            .ADDR3 = ADDR3.PadRight(30)
            .ADDR4 = ADDR4.PadRight(30)
            .ADDR5 = ADDR5.PadRight(30)
            .ADDR6 = ADDR6.PadRight(30)
            .APPINV = APPINV.PadRight(6)
            .CITY = CITY.PadRight(30)
            .CNTCT = CNTCT.PadRight(20)
            .CNTRY = CNTRY.PadRight(30)
            .COMMIS = COMMIS
            .COMNT1 = COMNT1.PadRight(30)
            .COMNT2 = COMNT2.PadRight(30)
            .COMNT3 = COMNT3.PadRight(30)
            .CREATEDBY = CREATEDBY.PadRight(100)
            If CREATIONDATE = DefaultDate Then
                .CREATIONDATE = 0
            Else
                .CREATIONDATE = MakeDate(CREATIONDATE.Year, CREATIONDATE.Month, CREATIONDATE.Day)
            End If
            .CURR = CURR.PadRight(3)
            .CUSTID = CUSTID.PadRight(20)
            .CUSTPO = CUSTPO.PadRight(25)
            .EXCRTE = EXCRTE
            .FEDTAX = FEDTAX.PadRight(1)
            .FILL01 = FILL01.PadRight(2)
            .FILL02 = FILL02.PadRight(2)
            .FILLER = FILLER.PadRight(50)
            .FIXVAR = FIXVAR.PadRight(1)
            .FOB = FOB.PadRight(15)
            .GLXREF = GLXREF.PadRight(32)
            .INVCE = INVCE.PadRight(1)
            .LNETAX = LNETAX.PadRight(1)
            .MCOMP = MCOMP.PadRight(3)
            .MODIFIEDBY = MODIFIEDBY.PadRight(100)
            If MODIFICATIONDATE = DefaultDate Then
                .MODIFICATIONDATE = 0
            Else
                .MODIFICATIONDATE = MakeDate(MODIFICATIONDATE.Year, MODIFICATIONDATE.Month, MODIFICATIONDATE.Day)
            End If
            .MSITE = MSITE.PadRight(3)
            .NAME = NAME.PadRight(30)
            .ORDDTE = MakeDate(ORDDTE.Year, ORDDTE.Month, ORDDTE.Day)
            .ORDID = ORDID.PadRight(20)
            .ORDNUM = ORDNUM.PadRight(8)
            .PHONE = PHONE.PadRight(20)
            If RCLDTE = DefaultDate Then
                .RCLDTE = 0
            Else
                .RCLDTE = MakeDate(RCLDTE.Year, RCLDTE.Month, RCLDTE.Day)
            End If
            .REASON = REASON.PadRight(2)
            .REP1 = REP1.PadRight(7)
            .REP2 = REP2.PadRight(7)
            .REP3 = REP3.PadRight(7)
            .SHPCDE = SHPCDE.PadRight(10)
            .SHPLBL = SHPLBL
            .SHPTHRU = SHPTHRU.PadRight(10)
            .SHPVIA = SHPVIA.PadRight(2)
            .SPLIT1 = SPLIT1
            .SPLIT2 = SPLIT2
            .SPLIT3 = SPLIT3
            .STATE = STATE.PadRight(30)
            .STATUS = STATUS.PadRight(1)
            .STYPE = STYPE.PadRight(2)
            .TAXABL = TAXABL.PadRight(1)
            .TAXCD1 = TAXCD1.PadRight(7)
            .TAXCD2 = TAXCD2.PadRight(7)
            .TAXCD3 = TAXCD3.PadRight(7)
            .TAXPRV = TAXPRV.PadRight(15)
            .TERMS = TERMS.PadRight(2)
            .TTAX = TTAX
            .UDFKEY = UDFKEY.PadRight(15)
            .UDFREF = UDFREF.PadRight(25)
            .XDFBOL = XDFBOL.PadRight(1)
            If XDFDTE = DefaultDate Then
                .XDFDTE = 0
            Else
                .XDFDTE = MakeDate(XDFDTE.Year, XDFDTE.Month, XDFDTE.Day)
            End If
            .XDFFLT = XDFFLT
            .XDFINT = XDFINT
            .XDFTXT = XDFTXT.PadRight(100)
            .XURR = XURR.PadRight(2)
            .ZIPCD = ZIPCD.PadRight(30)
        End With
    End Sub

    Public Function IsDuplicatePO(ByVal passCustomer As String, ByVal passOrder As String, ByVal passPO As String) As Boolean
        Dim isDuplicate As Boolean

        Dim strSQL As String = "select CUSTPO_27 " & _
                               "From ""SO_Master"" " & _
                               "Where CUSTID_27 = '" & passCustomer.Trim & "' " & _
                               "and CUSTPO_27 = '" & passPO.Trim & "' " & _
                               "and ORDNUM_27 <> '" & passOrder.Trim & "'"

        ' Set up the new Sql command
        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim dr As SqlDataReader = cmd.ExecuteReader()
        If dr.Read() Then
            isDuplicate = True
        Else
            isDuplicate = False
        End If

        ' Close the datareader and max connection and free up memory
        dr.Close()
        dr = Nothing
        CloseMaxConnection()

        ' Return the appropriate indicator
        Return isDuplicate
    End Function

    ' This function is the initialization routine that needs to be called before any MAXORDR2 function is called
    Public Function InitializeMax(ByVal connStr As String, ByVal comName As String, ByVal licPath As String, ByVal logPath As String, ByVal log As Boolean) As Integer
        Dim maxHandle As Integer = 0
        maxHandle = InitMAXOrder(connStr, comName, licPath, logPath, 0, log)
        Return maxHandle
    End Function

    ' This is the cleanup function that needs to be called when finished with MAXORDR2 processing
    Public Function ShutDownMax(ByVal MaxHandle As Integer) As Integer
        Dim rc As Integer = 0

        rc = ShutdownMAXOrder(MaxHandle)

        Return rc
    End Function

End Class
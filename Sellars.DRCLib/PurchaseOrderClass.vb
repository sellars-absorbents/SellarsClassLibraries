Imports Sellars.SQL
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq

Public Class PurchaseOrderClass
    Inherits ClassBase

    Private _found As Boolean = False

    ' Set up working variables for all the fields that are stored in the database
    Private _ORDNUM As String = ""
    Private _LINNUM As String = ""
    Private _DELNUM As String = ""
    Private _PRTNUM As String = ""
    Private _CURDUE As Date = DefaultDate
    Private _RECFLG As String = ""
    Private _TAXABLE As String = ""
    Private _TYPE As String = ""
    Private _ORDER As String = ""
    Private _VENID As String = ""
    Private _ORGDUE As Date = DefaultDate
    Private _PURUOM As String = ""
    Private _CURQTY As Double = 0
    Private _ORGQTY As Double = 0
    Private _DUEQTY As Double = 0
    Private _CURPRM As Date = DefaultDate
    Private _ORGPRM As Date = DefaultDate
    Private _FRMPLN As String = ""
    Private _STATUS As String = ""
    Private _STK As String = ""
    Private _CUSORD As String = ""
    Private _PLANID As String = ""
    Private _BUYER As String = ""
    Private _PSCRAP As Double = 0
    Private _ASCRAP As Double = 0
    Private _SCRPCD As String = ""
    Private _SCHCDE As String = ""
    Private _REVLEV As String = ""
    Private _COST As Double = 0
    Private _CSTCNV As Double = 0
    Private _APRDBY As String = ""
    Private _ORDREF As String = ""
    Private _TRNDTE As Date = DefaultDate
    Private _SCHFLG As String = ""
    Private _CRTRAT As String = ""
    Private _NEGATV As String = ""
    Private _REQPEG As String = ""
    Private _MPNNUM As String = ""
    Private _LABOR As Double = 0
    Private _AMMEND As String = ""
    Private _LOTNUM As String = ""
    Private _BEGSER As String = ""
    Private _REWORK As String = ""
    Private _CRTSNS As String = ""
    Private _TTLSNS As Double = 0
    Private _FORCUR As Double = 0
    Private _EXCESS As Single = 0
    Private _UOMCST As Double = 0
    Private _UOMCNV As Double = 0
    Private _INSREQ As String = ""
    Private _CREDTE As Date = DefaultDate
    Private _RTEREV As String = ""
    Private _RTEDTE As Date = DefaultDate
    Private _COMCDE As String = ""
    Private _ORDPTP As String = ""
    Private _JOBEXP As String = ""
    Private _JOBCST As Double = 0
    Private _TAXCDE As String = ""
    Private _TAX1 As Double = 0
    Private _GLREF As String = ""
    Private _CURR As String = ""
    Private _UDFKEY As String = ""
    Private _UDFREF As String = ""
    Private _DISC As Single = 0
    Private _RECCOST As Double = 0
    Private _MPNMFG As String = ""
    Private _DEXPFLG As String = ""
    Private _PLSTPRNT As String = ""
    Private _ROUTPRNT As String = ""
    Private _REQUES As String = ""
    Private _CLSDTE As Date = DefaultDate
    Private _XDFINT As Integer = 0
    Private _XDFFLT As Double = 0
    Private _XDFBOL As String = ""
    Private _XDFDTE As Date = DefaultDate
    Private _XDFTXT As String = ""
    Private _CREATEDBY As String = ""
    Private _CREATIONDATE As Date = DefaultDate
    Private _MODIFIEDBY As String = ""
    Private _MODIFICATIONDATE As Date = DefaultDate
    Private _TSKCDE As String = ""
    Private _TSKTYP As String = ""
    Private _REPORTER As String = ""
    Private _PRIORITY As String = ""
    Private _PHONE As String = ""
    Private _LOCATION As String = ""
    Private _FILL03 As String = ""
    Private _FILL04 As String = ""
    Private _FILL05 As String = ""
    Private _FILLER As String = ""

#Region "Properties - String"
    ' Set up properties for all the fields in the database
    Public Property ORDNUM() As String
        Get
            Return _ORDNUM.Trim
        End Get
        Set(ByVal value As String)
            _ORDNUM = value.Trim
        End Set
    End Property

    Public Property LINNUM() As String
        Get
            Return _LINNUM.Trim
        End Get
        Set(ByVal value As String)
            _LINNUM = value.Trim
        End Set
    End Property

    Public Property DELNUM() As String
        Get
            Return _DELNUM.Trim
        End Get
        Set(ByVal value As String)
            _DELNUM = value.Trim
        End Set
    End Property

    Public Property PRTNUM() As String
        Get
            Return _PRTNUM.Trim
        End Get
        Set(ByVal value As String)
            _PRTNUM = value.Trim
        End Set
    End Property
    Public Property RECFLG() As String
        Get
            Return _RECFLG.Trim
        End Get
        Set(ByVal value As String)
            _RECFLG = value.Trim
        End Set
    End Property

    Public Property TAXABLE() As String
        Get
            Return _TAXABLE.Trim
        End Get
        Set(ByVal value As String)
            _TAXABLE = value.Trim
        End Set
    End Property

    Public Property TYPE() As String
        Get
            Return _TYPE.Trim
        End Get
        Set(ByVal value As String)
            _TYPE = value.Trim
        End Set
    End Property

    Public Property ORDER() As String
        Get
            Return _ORDER.Trim
        End Get
        Set(ByVal value As String)
            _ORDER = value.Trim
        End Set
    End Property

    Public Property VENID() As String
        Get
            Return _VENID.Trim
        End Get
        Set(ByVal value As String)
            _VENID = value.Trim
        End Set
    End Property

    Public Property PURUOM() As String
        Get
            Return _PURUOM.Trim
        End Get
        Set(ByVal value As String)
            _PURUOM = value.Trim
        End Set
    End Property

    Public Property FRMPLN() As String
        Get
            Return _FRMPLN.Trim
        End Get
        Set(ByVal value As String)
            _FRMPLN = value.Trim
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

    Public Property STK() As String
        Get
            Return _STK.Trim
        End Get
        Set(ByVal value As String)
            _STK = value.Trim
        End Set
    End Property

    Public Property CUSORD() As String
        Get
            Return _CUSORD.Trim
        End Get
        Set(ByVal value As String)
            _CUSORD = value.Trim
        End Set
    End Property

    Public Property PLANID() As String
        Get
            Return _PLANID.Trim
        End Get
        Set(ByVal value As String)
            _PLANID = value.Trim
        End Set
    End Property

    Public Property BUYER() As String
        Get
            Return _BUYER.Trim
        End Get
        Set(ByVal value As String)
            _BUYER = value.Trim
        End Set
    End Property

    Public Property SCRPCD() As String
        Get
            Return _SCRPCD.Trim
        End Get
        Set(ByVal value As String)
            _SCRPCD = value.Trim
        End Set
    End Property

    Public Property SCHCDE() As String
        Get
            Return _SCHCDE.Trim
        End Get
        Set(ByVal value As String)
            _SCHCDE = value.Trim
        End Set
    End Property

    Public Property REVLEV() As String
        Get
            Return _REVLEV.Trim
        End Get
        Set(ByVal value As String)
            _REVLEV = value.Trim
        End Set
    End Property

    Public Property APRDBY() As String
        Get
            Return _APRDBY.Trim
        End Get
        Set(ByVal value As String)
            _APRDBY = value.Trim
        End Set
    End Property

    Public Property ORDREF() As String
        Get
            Return _ORDREF.Trim
        End Get
        Set(ByVal value As String)
            _ORDREF = value.Trim
        End Set
    End Property

    Public Property SCHFLG() As String
        Get
            Return _SCHFLG.Trim
        End Get
        Set(ByVal value As String)
            _SCHFLG = value.Trim
        End Set
    End Property

    Public Property CRTRAT() As String
        Get
            Return _CRTRAT.Trim
        End Get
        Set(ByVal value As String)
            _CRTRAT = value.Trim
        End Set
    End Property

    Public Property NEGATV() As String
        Get
            Return _NEGATV.Trim
        End Get
        Set(ByVal value As String)
            _NEGATV = value.Trim
        End Set
    End Property

    Public Property REQPEG() As String
        Get
            Return _REQPEG.Trim
        End Get
        Set(ByVal value As String)
            _REQPEG = value.Trim
        End Set
    End Property

    Public Property MPNNUM() As String
        Get
            Return _MPNNUM.Trim
        End Get
        Set(ByVal value As String)
            _MPNNUM = value.Trim
        End Set
    End Property

    Public Property AMMEND() As String
        Get
            Return _AMMEND.Trim
        End Get
        Set(ByVal value As String)
            _AMMEND = value.Trim
        End Set
    End Property

    Public Property LOTNUM() As String
        Get
            Return _LOTNUM.Trim
        End Get
        Set(ByVal value As String)
            _LOTNUM = value.Trim
        End Set
    End Property

    Public Property BEGSER() As String
        Get
            Return _BEGSER.Trim
        End Get
        Set(ByVal value As String)
            _BEGSER = value.Trim
        End Set
    End Property

    Public Property REWORK() As String
        Get
            Return _REWORK.Trim
        End Get
        Set(ByVal value As String)
            _REWORK = value.Trim
        End Set
    End Property

    Public Property CRTSNS() As String
        Get
            Return _CRTSNS.Trim
        End Get
        Set(ByVal value As String)
            _CRTSNS = value.Trim
        End Set
    End Property

    Public Property INSREQ() As String
        Get
            Return _INSREQ.Trim
        End Get
        Set(ByVal value As String)
            _INSREQ = value.Trim
        End Set
    End Property

    Public Property RTEREV() As String
        Get
            Return _RTEREV.Trim
        End Get
        Set(ByVal value As String)
            _RTEREV = value.Trim
        End Set
    End Property

    Public Property COMCDE() As String
        Get
            Return _COMCDE.Trim
        End Get
        Set(ByVal value As String)
            _COMCDE = value.Trim
        End Set
    End Property

    Public Property ORDPTP() As String
        Get
            Return _ORDPTP.Trim
        End Get
        Set(ByVal value As String)
            _ORDPTP = value.Trim
        End Set
    End Property

    Public Property JOBEXP() As String
        Get
            Return _JOBEXP.Trim
        End Get
        Set(ByVal value As String)
            _JOBEXP = value.Trim
        End Set
    End Property

    Public Property TAXCDE() As String
        Get
            Return _TAXCDE.Trim
        End Get
        Set(ByVal value As String)
            _TAXCDE = value.Trim
        End Set
    End Property

    Public Property GLREF() As String
        Get
            Return _GLREF.Trim
        End Get
        Set(ByVal value As String)
            _GLREF = value.Trim
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

    Public Property MPNMFG() As String
        Get
            Return _MPNMFG.Trim
        End Get
        Set(ByVal value As String)
            _MPNMFG = value.Trim
        End Set
    End Property

    Public Property DEXPFLG() As String
        Get
            Return _DEXPFLG.Trim
        End Get
        Set(ByVal value As String)
            _DEXPFLG = value.Trim
        End Set
    End Property

    Public Property PLSTPRNT() As String
        Get
            Return _PLSTPRNT.Trim
        End Get
        Set(ByVal value As String)
            _PLSTPRNT = value.Trim
        End Set
    End Property

    Public Property ROUTPRNT() As String
        Get
            Return _ROUTPRNT.Trim
        End Get
        Set(ByVal value As String)
            _ROUTPRNT = value.Trim
        End Set
    End Property

    Public Property REQUES() As String
        Get
            Return _REQUES.Trim
        End Get
        Set(ByVal value As String)
            _REQUES = value.Trim
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

    Public Property MODIFIEDBY() As String
        Get
            Return _MODIFIEDBY.Trim
        End Get
        Set(ByVal value As String)
            _MODIFIEDBY = value.Trim
        End Set
    End Property

    Public Property TSKCDE() As String
        Get
            Return _TSKCDE.Trim
        End Get
        Set(ByVal value As String)
            _TSKCDE = value.Trim
        End Set
    End Property

    Public Property TSKTYP() As String
        Get
            Return _TSKTYP.Trim
        End Get
        Set(ByVal value As String)
            _TSKTYP = value.Trim
        End Set
    End Property

    Public Property REPORTER() As String
        Get
            Return _REPORTER.Trim
        End Get
        Set(ByVal value As String)
            _REPORTER = value.Trim
        End Set
    End Property

    Public Property PRIORITY() As String
        Get
            Return _PRIORITY.Trim
        End Get
        Set(ByVal value As String)
            _PRIORITY = value.Trim
        End Set
    End Property

    Public Property PHONE() As String
        Get
            Return _PHONE.Trim
        End Get
        Set(ByVal value As String)
            _PHONE = value.Trim
        End Set
    End Property

    Public Property LOCATION() As String
        Get
            Return _LOCATION.Trim
        End Get
        Set(ByVal value As String)
            _LOCATION = value.Trim
        End Set
    End Property

    Public Property FILL03() As String
        Get
            Return _FILL03.Trim
        End Get
        Set(ByVal value As String)
            _FILL03 = value.Trim
        End Set
    End Property

    Public Property FILL04() As String
        Get
            Return _FILL04.Trim
        End Get
        Set(ByVal value As String)
            _FILL04 = value.Trim
        End Set
    End Property

    Public Property FILL05() As String
        Get
            Return _FILL05.Trim
        End Get
        Set(ByVal value As String)
            _FILL05 = value.Trim
        End Set
    End Property

    Public Property FILLER() As String
        Get
            Return _FILLER.Trim
        End Get
        Set(ByVal value As String)
            _FILLER = value.Trim
        End Set
    End Property
#End Region

#Region "Properties - Date"
    Public Property CURDUE() As Date
        Get
            Return _CURDUE
        End Get
        Set(ByVal value As Date)
            _CURDUE = value
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

    Public Property CURPRM() As Date
        Get
            Return _CURPRM
        End Get
        Set(ByVal value As Date)
            _CURPRM = value
        End Set
    End Property

    Public Property ORGPRM() As Date
        Get
            Return _ORGPRM
        End Get
        Set(ByVal value As Date)
            _ORGPRM = value
        End Set
    End Property

    Public Property TRNDTE() As Date
        Get
            Return _TRNDTE
        End Get
        Set(ByVal value As Date)
            _TRNDTE = value
        End Set
    End Property

    Public Property CREDTE() As Date
        Get
            Return _CREDTE
        End Get
        Set(ByVal value As Date)
            _CREDTE = value
        End Set
    End Property

    Public Property RTEDTE() As Date
        Get
            Return _RTEDTE
        End Get
        Set(ByVal value As Date)
            _RTEDTE = value
        End Set
    End Property

    Public Property CLSDTE() As Date
        Get
            Return _CLSDTE
        End Get
        Set(ByVal value As Date)
            _CLSDTE = value
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

    Public Property CREATIONDATE() As Date
        Get
            Return _CREATIONDATE
        End Get
        Set(ByVal value As Date)
            _CREATIONDATE = value
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

#Region "Properties - Integer"
    Public Property XDFINT() As Integer
        Get
            Return _XDFINT
        End Get
        Set(ByVal value As Integer)
            _XDFINT = value
        End Set
    End Property
#End Region

#Region "Properties - Double"
    Public Property CURQTY() As Double
        Get
            Return _CURQTY
        End Get
        Set(ByVal value As Double)
            _CURQTY = value
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

    Public Property DUEQTY() As Double
        Get
            Return _DUEQTY
        End Get
        Set(ByVal value As Double)
            _DUEQTY = value
        End Set
    End Property

    Public Property PSCRAP() As Double
        Get
            Return _PSCRAP
        End Get
        Set(ByVal value As Double)
            _PSCRAP = value
        End Set
    End Property

    Public Property ASCRAP() As Double
        Get
            Return _ASCRAP
        End Get
        Set(ByVal value As Double)
            _ASCRAP = value
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

    Public Property CSTCNV() As Double
        Get
            Return _CSTCNV
        End Get
        Set(ByVal value As Double)
            _CSTCNV = value
        End Set
    End Property

    Public Property LABOR() As Double
        Get
            Return _LABOR
        End Get
        Set(ByVal value As Double)
            _LABOR = value
        End Set
    End Property

    Public Property TTLSNS() As Double
        Get
            Return _TTLSNS
        End Get
        Set(ByVal value As Double)
            _TTLSNS = value
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

    Public Property UOMCST() As Double
        Get
            Return _UOMCST
        End Get
        Set(ByVal value As Double)
            _UOMCST = value
        End Set
    End Property

    Public Property UOMCNV() As Double
        Get
            Return _UOMCNV
        End Get
        Set(ByVal value As Double)
            _UOMCNV = value
        End Set
    End Property

    Public Property JOBCST() As Double
        Get
            Return _JOBCST
        End Get
        Set(ByVal value As Double)
            _JOBCST = value
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

    Public Property RECCOST() As Double
        Get
            Return _RECCOST
        End Get
        Set(ByVal value As Double)
            _RECCOST = value
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
#End Region

#Region "Properties - Single"
    Public Property EXCESS() As Single
        Get
            Return _EXCESS
        End Get
        Set(ByVal value As Single)
            _EXCESS = value
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
#End Region

    Public Sub New()
    End Sub

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

    Private Function dbRecordString(ByVal pField As String) As String
        If IsDBNull(pField) Then
            Return ""
        Else
            Return pField.Trim
        End If
    End Function

    Public Sub Read(ByVal passorder As String, ByVal LineNumber As String, ByVal DeliveryNumber As String)

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim CustomerPrice As Decimal = 0
        Dim strSQL As String = "Select * " & _
                               "From ORDER_MASTER " & _
                               "Where ORDNUM_10 = @ORDNUM " & _
                               "and LINNUM_10 = @LINNUM " & _
                               "and DELNUM_10 = @DELNUM"

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.Add(New SqlParameter("@ORDNUM", passorder))
        cmd.Parameters.Add(New SqlParameter("@LINNUM", LineNumber))
        cmd.Parameters.Add(New SqlParameter("@DELNUM", DeliveryNumber))

        Dim OrderMasterReader As SqlDataReader = cmd.ExecuteReader()

        _ORDNUM = passorder
        If OrderMasterReader.Read() Then
            _found = True
            _LINNUM = dbRecordString(OrderMasterReader("LINNUM_10"))
            _DELNUM = dbRecordString(OrderMasterReader("DELNUM_10"))
            _PRTNUM = dbRecordString(OrderMasterReader("PRTNUM_10"))
            Try
                _CURDUE = dbRecordDate(OrderMasterReader("CURDUE_10"))
            Catch
                _CURDUE = DefaultDate
            End Try
            _RECFLG = dbRecordString(OrderMasterReader("RECFLG_10"))
            _TAXABLE = dbRecordString(OrderMasterReader("TAXABLE_10"))
            _TYPE = dbRecordString(OrderMasterReader("TYPE_10"))
            _ORDER = dbRecordString(OrderMasterReader("ORDER_10"))
            _VENID = dbRecordString(OrderMasterReader("VENID_10"))
            Try
                _ORGDUE = dbRecordDate(OrderMasterReader("ORGDUE_10"))
            Catch
                _ORGDUE = DefaultDate
            End Try
            _PURUOM = dbRecordString(OrderMasterReader("PURUOM_10"))
            _CURQTY = dbRecordDouble(OrderMasterReader("CURQTY_10"))
            _ORGQTY = dbRecordDouble(OrderMasterReader("ORGQTY_10"))
            _DUEQTY = dbRecordDouble(OrderMasterReader("DUEQTY_10"))
            Try
                _CURPRM = dbRecordDate(OrderMasterReader("CURPRM_10"))
            Catch
                _CURPRM = DefaultDate
            End Try
            _FILL03 = dbRecordString(OrderMasterReader("FILL03_10"))
            Try
                _ORGPRM = dbRecordDate(OrderMasterReader("ORGPRM_10"))
            Catch
                _ORGPRM = DefaultDate
            End Try
            _FILL04 = dbRecordString(OrderMasterReader("FILL04_10"))
            _FRMPLN = dbRecordString(OrderMasterReader("FRMPLN_10"))
            _STATUS = dbRecordString(OrderMasterReader("STATUS_10"))
            _STK = dbRecordString(OrderMasterReader("STK_10"))
            _CUSORD = dbRecordString(OrderMasterReader("CUSORD_10"))
            _PLANID = dbRecordString(OrderMasterReader("PLANID_10"))
            _BUYER = dbRecordString(OrderMasterReader("BUYER_10"))
            _PSCRAP = dbRecordDouble(OrderMasterReader("PSCRAP_10"))
            _ASCRAP = dbRecordDouble(OrderMasterReader("ASCRAP_10"))
            _SCRPCD = dbRecordString(OrderMasterReader("SCRPCD_10"))
            _SCHCDE = dbRecordString(OrderMasterReader("SCHCDE_10"))
            _REVLEV = dbRecordString(OrderMasterReader("REVLEV_10"))
            _COST = dbRecordDouble(OrderMasterReader("COST_10"))
            _CSTCNV = dbRecordDouble(OrderMasterReader("CSTCNV_10"))
            _APRDBY = dbRecordString(OrderMasterReader("APRDBY_10"))
            _ORDREF = dbRecordString(OrderMasterReader("ORDREF_10"))
            Try
                _TRNDTE = dbRecordDate(OrderMasterReader("TRNDTE_10"))
            Catch
                _TRNDTE = DefaultDate
            End Try
            _FILL05 = dbRecordString(OrderMasterReader("FILL05_10"))
            _SCHFLG = dbRecordString(OrderMasterReader("SCHFLG_10"))
            _CRTRAT = dbRecordString(OrderMasterReader("CRTRAT_10"))
            _NEGATV = dbRecordString(OrderMasterReader("NEGATV_10"))
            _REQPEG = dbRecordString(OrderMasterReader("REQPEG_10"))
            _MPNNUM = dbRecordString(OrderMasterReader("MPNNUM_10"))
            _LABOR = dbRecordDouble(OrderMasterReader("LABOR_10"))
            _AMMEND = dbRecordString(OrderMasterReader("AMMEND_10"))
            _LOTNUM = dbRecordString(OrderMasterReader("LOTNUM_10"))
            _BEGSER = dbRecordString(OrderMasterReader("BEGSER_10"))
            _REWORK = dbRecordString(OrderMasterReader("REWORK_10"))
            _CRTSNS = dbRecordString(OrderMasterReader("CRTSNS_10"))
            _TTLSNS = dbRecordDouble(OrderMasterReader("TTLSNS_10"))
            _FORCUR = dbRecordDouble(OrderMasterReader("FORCUR_10"))
            _EXCESS = dbRecordSingle(OrderMasterReader("EXCESS_10"))
            _UOMCST = dbRecordDouble(OrderMasterReader("UOMCST_10"))
            _UOMCNV = dbRecordDouble(OrderMasterReader("UOMCNV_10"))
            _INSREQ = dbRecordString(OrderMasterReader("INSREQ_10"))
            Try
                _CREDTE = dbRecordDate(OrderMasterReader("CREDTE_10"))
            Catch
                _CREDTE = DefaultDate
            End Try
            _RTEREV = dbRecordString(OrderMasterReader("RTEREV_10"))
            Try
                _RTEDTE = dbRecordDate(OrderMasterReader("RTEDTE_10"))
            Catch
                _RTEDTE = DefaultDate
            End Try
            _COMCDE = dbRecordString(OrderMasterReader("COMCDE_10"))
            _ORDPTP = dbRecordString(OrderMasterReader("ORDPTP_10"))
            _JOBEXP = dbRecordString(OrderMasterReader("JOBEXP_10"))
            _JOBCST = dbRecordDouble(OrderMasterReader("JOBCST_10"))
            _TAXCDE = dbRecordString(OrderMasterReader("TAXCDE_10"))
            _TAX1 = dbRecordDouble(OrderMasterReader("TAX1_10"))
            _GLREF = dbRecordString(OrderMasterReader("GLREF_10"))
            _CURR = dbRecordString(OrderMasterReader("CURR_10"))
            _UDFKEY = dbRecordString(OrderMasterReader("UDFKEY_10"))
            _UDFREF = dbRecordString(OrderMasterReader("UDFREF_10"))
            _DISC = dbRecordSingle(OrderMasterReader("DISC_10"))
            _RECCOST = dbRecordDouble(OrderMasterReader("RECCOST_10"))
            _MPNMFG = dbRecordString(OrderMasterReader("MPNMFG_10"))
            _DEXPFLG = dbRecordString(OrderMasterReader("DEXPFLG_10"))
            _PLSTPRNT = dbRecordString(OrderMasterReader("PLSTPRNT_10"))
            _ROUTPRNT = dbRecordString(OrderMasterReader("ROUTPRNT_10"))
            _REQUES = dbRecordString(OrderMasterReader("REQUES_10"))
            Try
                _CLSDTE = dbRecordDate(OrderMasterReader("CLSDTE_10"))
            Catch
                _CLSDTE = DefaultDate
            End Try
            _XDFINT = dbRecordInt(OrderMasterReader("XDFINT_10"))
            _XDFFLT = dbRecordDouble(OrderMasterReader("XDFFLT_10"))
            _XDFBOL = dbRecordString(OrderMasterReader("XDFBOL_10"))
            ' Because we can get bad dates in the recall date field
            ' first try to read it, and if we can't set it to the default
            Try
                _XDFDTE = dbRecordDate(OrderMasterReader("XDFDTE_10"))
            Catch
                _XDFDTE = DefaultDate
            End Try
            _XDFTXT = dbRecordString(OrderMasterReader("XDFTXT_10"))
            _FILLER = dbRecordString(OrderMasterReader("FILLER_10"))
            _CREATEDBY = dbRecordString(OrderMasterReader("CreatedBy"))
            _CREATIONDATE = dbRecordDate(OrderMasterReader("CreationDate"))
            _MODIFIEDBY = dbRecordString(OrderMasterReader("ModifiedBy"))
            _MODIFICATIONDATE = dbRecordDate(OrderMasterReader("ModificationDate"))
            _TSKCDE = dbRecordString(OrderMasterReader("TSKCDE_10"))
            _TSKTYP = dbRecordString(OrderMasterReader("TSKTYP_10"))
            _REPORTER = dbRecordString(OrderMasterReader("REPORTER_10"))
            _PRIORITY = dbRecordString(OrderMasterReader("PRIORITY_10"))
            _PHONE = dbRecordString(OrderMasterReader("PHONE_10"))
            _LOCATION = dbRecordString(OrderMasterReader("LOCATION_10"))
        Else
            _found = False
            _LINNUM = ""
            _DELNUM = ""
            _PRTNUM = ""
            _CURDUE = DefaultDate
            _RECFLG = ""
            _TAXABLE = ""
            _TYPE = ""
            _ORDER = ""
            _VENID = ""
            _ORGDUE = DefaultDate
            _PURUOM = ""
            _CURQTY = 0
            _ORGQTY = 0
            _DUEQTY = 0
            _CURPRM = DefaultDate
            _FILL03 = ""
            _ORGPRM = DefaultDate
            _FILL04 = ""
            _FRMPLN = ""
            _STATUS = ""
            _STK = ""
            _CUSORD = ""
            _PLANID = ""
            _BUYER = ""
            _PSCRAP = 0
            _ASCRAP = 0
            _SCRPCD = ""
            _SCHCDE = ""
            _REVLEV = ""
            _COST = 0
            _CSTCNV = 0
            _APRDBY = ""
            _ORDREF = ""
            _TRNDTE = DefaultDate
            _FILL05 = ""
            _SCHFLG = ""
            _CRTRAT = ""
            _NEGATV = ""
            _REQPEG = ""
            _MPNNUM = ""
            _LABOR = 0
            _AMMEND = ""
            _LOTNUM = ""
            _BEGSER = ""
            _REWORK = ""
            _CRTSNS = ""
            _TTLSNS = 0
            _FORCUR = 0
            _EXCESS = 0
            _UOMCST = 0
            _UOMCNV = 0
            _INSREQ = ""
            _CREDTE = DefaultDate
            _RTEREV = ""
            _RTEDTE = DefaultDate
            _COMCDE = ""
            _ORDPTP = ""
            _JOBEXP = ""
            _JOBCST = 0
            _TAXCDE = ""
            _TAX1 = 0
            _GLREF = ""
            _CURR = ""
            _UDFKEY = ""
            _UDFREF = ""
            _DISC = 0
            _RECCOST = 0
            _MPNMFG = ""
            _DEXPFLG = ""
            _PLSTPRNT = ""
            _ROUTPRNT = ""
            _REQUES = ""
            _CLSDTE = DefaultDate
            _XDFINT = 0
            _XDFFLT = 0
            _XDFBOL = ""
            _XDFDTE = DefaultDate
            _XDFTXT = ""
            _FILLER = ""
            _CREATEDBY = ""
            _CREATIONDATE = DefaultDate
            _MODIFIEDBY = ""
            _MODIFICATIONDATE = DefaultDate
            _TSKCDE = ""
            _TSKTYP = ""
            _REPORTER = ""
            _PRIORITY = ""
            _PHONE = ""
            _LOCATION = ""
        End If

        OrderMasterReader.Close()
        OrderMasterReader = Nothing

        CloseMaxConnection()
    End Sub

    ' Function used to update any updateable line item field
    Public Sub UpdateScheduledDate(ByVal Connection As String, ByVal CompanyName As String, ByVal LicensePath As String, ByVal LogPath As String, ByVal OrderNumber As String, ByVal LineNumber As String, ByVal DeliveryNumber As String, ByVal ScheduledDate As Date)
        ' Read the requested sales order detail record
        Read(OrderNumber, LineNumber, DeliveryNumber)

        ' Update the due date
        CURDUE = ScheduledDate

        ' Call the function to update the purchase order
        UpdateOrder(Connection, CompanyName, LicensePath, LogPath)
    End Sub

    Private Sub UpdateOrder(ByVal Connection As String, ByVal CompanyName As String, ByVal LicensePath As String, ByVal LogPath As String)
        Dim sXML As New StringBuilder

        ClassBase.Initialize(Connection, CompanyName, LicensePath, LogPath, False)

        ' Create the XML document to update the Shop Order
        Dim docPO As XDocument =
            <?xml version="1.0" encoding="utf-8"?>
            <eMAXExact>
                <Order_Master_Table>
                    <Order_Master>
                        <ORDNUM_10><%= ORDNUM %></ORDNUM_10>
                        <LINNUM_10><%= LINNUM %></LINNUM_10>
                        <DELNUM_10><%= DELNUM %></DELNUM_10>
                        <PRTNUM_10><%= PRTNUM %></PRTNUM_10>
                        <CURDUE_10><%= Format(CURDUE, "yyyy-MM-dd") %></CURDUE_10>
                        <RECFLG_10><%= RECFLG %></RECFLG_10>
                        <TAXABLE_10><%= TAXABLE %></TAXABLE_10>
                        <TYPE_10><%= TYPE %></TYPE_10>
                        <ORDER_10><%= ORDER %></ORDER_10>
                        <VENID_10><%= VENID %></VENID_10>
                        <ORGDUE_10><%= Format(ORGDUE, "yyyy-MM-dd") %></ORGDUE_10>
                        <PURUOM_10><%= PURUOM %></PURUOM_10>
                        <CURQTY_10><%= CURQTY %></CURQTY_10>
                        <ORGQTY_10><%= ORGQTY %></ORGQTY_10>
                        <DUEQTY_10><%= DUEQTY %></DUEQTY_10>
                        <CURPRM_10><%= Format(CURPRM, "yyyy-MM-dd") %></CURPRM_10>
                        <ORGPRM_10><%= Format(ORGPRM, "yyyy-MM-dd") %></ORGPRM_10>
                        <FRMPLN_10><%= FRMPLN %></FRMPLN_10>
                        <STATUS_10><%= STATUS %></STATUS_10>
                        <STK_10><%= STK %></STK_10>
                        <CUSORD_10><%= CUSORD %></CUSORD_10>
                        <PLANID_10><%= PLANID %></PLANID_10>
                        <BUYER_10><%= BUYER %></BUYER_10>
                        <PSCRAP_10><%= PSCRAP %></PSCRAP_10>
                        <ASCRAP_10><%= ASCRAP %></ASCRAP_10>
                        <SCRPCD_10><%= SCRPCD %></SCRPCD_10>
                        <SCHCDE_10><%= SCHCDE %></SCHCDE_10>
                        <REVLEV_10><%= REVLEV %></REVLEV_10>
                        <COST_10><%= COST %></COST_10>
                        <CSTCNV_10><%= CSTCNV %></CSTCNV_10>
                        <APRDBY_10><%= APRDBY %></APRDBY_10>
                        <ORDREF_10><%= ORDREF %></ORDREF_10>
                        <TRNDTE_10><%= Format(TRNDTE, "yyyy-MM-dd") %></TRNDTE_10>
                        <SCHFLG_10><%= SCHFLG %></SCHFLG_10>
                        <CRTRAT_10><%= CRTRAT %></CRTRAT_10>
                        <NEGATV_10><%= NEGATV %></NEGATV_10>
                        <REQPEG_10><%= REQPEG %></REQPEG_10>
                        <MPNNUM_10><%= MPNNUM %></MPNNUM_10>
                        <LABOR_10><%= LABOR %></LABOR_10>
                        <AMMEND_10><%= AMMEND %></AMMEND_10>
                        <LOTNUM_10><%= LOTNUM %></LOTNUM_10>
                        <BEGSER_10><%= BEGSER %></BEGSER_10>
                        <REWORK_10><%= REWORK %></REWORK_10>
                        <CRTSNS_10><%= CRTSNS %></CRTSNS_10>
                        <TTLSNS_10><%= TTLSNS %></TTLSNS_10>
                        <FORCUR_10><%= FORCUR %></FORCUR_10>
                        <EXCESS_10><%= EXCESS %></EXCESS_10>
                        <UOMCST_10><%= UOMCST %></UOMCST_10>
                        <UOMCNV_10><%= UOMCNV %></UOMCNV_10>
                        <INSREQ_10><%= INSREQ %></INSREQ_10>
                        <CREDTE_10><%= Format(CREDTE, "yyyy-MM-dd") %></CREDTE_10>
                        <RTEREV_10><%= RTEREV %></RTEREV_10>
                        <RTEDTE_10><%= Format(RTEDTE, "yyyy-MM-dd") %></RTEDTE_10>
                        <COMCDE_10><%= COMCDE %></COMCDE_10>
                        <ORDPTP_10><%= ORDPTP %></ORDPTP_10>
                        <JOBEXP_10><%= JOBEXP %></JOBEXP_10>
                        <JOBCST_10><%= JOBCST %></JOBCST_10>
                        <TAXCDE_10><%= TAXCDE %></TAXCDE_10>
                        <TAX1_10><%= TAX1 %></TAX1_10>
                        <GLREF_10><%= GLREF %></GLREF_10>
                        <CURR_10><%= CURR %></CURR_10>
                        <UDFKEY_10><%= UDFKEY %></UDFKEY_10>
                        <UDFREF_10><%= UDFREF %></UDFREF_10>
                        <DISC_10><%= DISC %></DISC_10>
                        <RECCOST_10><%= RECCOST %></RECCOST_10>
                        <MPNMFG_10><%= MPNMFG %></MPNMFG_10>
                        <DEXPFLG_10><%= DEXPFLG %></DEXPFLG_10>
                        <PLSTPRNT_10><%= PLSTPRNT %></PLSTPRNT_10>
                        <ROUTPRNT_10><%= ROUTPRNT %></ROUTPRNT_10>
                        <REQUES_10><%= REQUES %></REQUES_10>
                        <CLSDTE_10><%= Format(CLSDTE, "yyyy-MM-dd") %></CLSDTE_10>
                        <TSKCDE_10><%= TSKCDE %></TSKCDE_10>
                        <TSKTYP_10><%= TSKTYP %></TSKTYP_10>
                        <REPORTER_10><%= REPORTER %></REPORTER_10>
                        <PRIORITY_10><%= PRIORITY %></PRIORITY_10>
                        <PHONE_10><%= PHONE %></PHONE_10>
                        <LOCATION_10><%= LOCATION %></LOCATION_10>
                        <FILLER_10><%= FILLER %></FILLER_10>
                        <FILLER03_10><%= FILL03 %></FILLER03_10>
                        <FILLER04_10><%= FILL04 %></FILLER04_10>
                        <FILLER05_10><%= FILL05 %></FILLER05_10>
                    </Order_Master>
                </Order_Master_Table>
            </eMAXExact>

        'Add Sales Order Detail record via MAX Update
        Dim result As Short = ClassBase.MAXUpdate.ChangePurchaseOrderLineItemXML(docPO.ToString, False)

        Select Case result
            Case 0
                'failed - make a new order number
                Dim msg As String = "Purchase Order Update failed using order number " & ORDNUM
            Case Else
                'succeeded
                Dim msg As String = "Purchase Order Update succeeded."
        End Select
    End Sub
End Class

Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class PartCharacteristicsClass
    Inherits ClassBase

    Private _Color As Integer = 99
    Private _ColorDescription As String = ""
    Private _cPerf As Boolean = False
    Private _CaseLength As Decimal = 0
    Private _CaseHeight As Decimal = 0
    Private _CaseWidth As Decimal = 0
    Private _CustomerType As Integer = 1
    Private _CustomerTypeDescription As String = ""
    Private _DueDaysID As Integer = 1
    Private _EngineeringNotes As String = ""
    Private _EachesPerMfgUnit As Integer = 0
    Private _Fold As Integer = 99
    Private _FoldDescription As String = ""
    Private _Found As Boolean = False
    Private _Grade As Integer = 99
    Private _GradeDescription As String = ""
    Private _LabelType As Integer = 99
    Private _LabelDescription As String = ""
    Private _Nascar As Boolean = False
    Private _PackageType As Integer = 99
    Private _PackageDescription As String = ""
    Private _PalletQuantity As Integer = 0
    Private _Pattern As Integer = 99
    Private _PatternDescription As String = ""
    Private _SlitWidth As Decimal = 0
    Private _RMBasisWeight As Decimal = 0
    Private _FGBasisWeight As Decimal = 0

    Private _GrossWeight As Decimal = 0
    Private _CasesPerPalletLTL As Integer = 0
    Private _CasesPerPalletTL As Integer = 0
    Private _CasesPerPalletLayer As Integer = 0
    Private _PalletsPerTruck As Integer = 0
    Private _CasesPerTruck As Integer = 0
    Private _ShippingClass As String = ""

    Private _xPerf As Boolean = False

    Public ReadOnly Property CaseHeight() As Decimal
        Get
            Return _CaseHeight
        End Get
    End Property

    Public ReadOnly Property CaseLength() As Decimal
        Get
            Return _CaseLength
        End Get
    End Property

    Public ReadOnly Property CaseWidth() As Decimal
        Get
            Return _CaseWidth
        End Get
    End Property

    Public ReadOnly Property Color() As Integer
        Get
            Return _Color
        End Get
    End Property

    Public ReadOnly Property CenterPerf() As Boolean
        Get
            Return _cPerf
        End Get
    End Property

    Public ReadOnly Property CrossPerf() As Boolean
        Get
            Return _xPerf
        End Get
    End Property

    Public ReadOnly Property CustomerType() As Integer
        Get
            Return _CustomerType
        End Get
    End Property

    Public ReadOnly Property CustomerTypeDescription() As String
        Get
            Return _CustomerTypeDescription
        End Get
    End Property

    Public ReadOnly Property DueDaysID() As Integer
        Get
            Return _DueDaysID
        End Get
    End Property

    Public ReadOnly Property EachesPerMfgUnit() As Integer
        Get
            Return _EachesPerMfgUnit
        End Get
    End Property

    Public ReadOnly Property EngineeringNotes() As String
        Get
            Return _EngineeringNotes
        End Get
    End Property

    Public ReadOnly Property Fold() As Integer
        Get
            Return _Fold
        End Get
    End Property

    Public ReadOnly Property Found() As Boolean
        Get
            Return _Found
        End Get
    End Property

    Public ReadOnly Property Grade() As Integer
        Get
            Return _Grade
        End Get
    End Property

    Public ReadOnly Property LabelType() As Integer
        Get
            Return _LabelType
        End Get
    End Property

    Public ReadOnly Property PalletQuantity() As Integer
        Get
            Return _PalletQuantity
        End Get
    End Property

    Public ReadOnly Property Nascar() As Boolean
        Get
            Return _Nascar
        End Get
    End Property

    Public ReadOnly Property PackageType() As Integer
        Get
            Return _PackageType
        End Get
    End Property

    Public ReadOnly Property Pattern() As Integer
        Get
            Return _Pattern
        End Get
    End Property

    Public ReadOnly Property SlitWidth() As Decimal
        Get
            Return _SlitWidth
        End Get
    End Property

    Public ReadOnly Property FGBasisWeight() As Decimal
        Get
            Return _FGBasisWeight
        End Get
    End Property

    Public ReadOnly Property RMBasisWeight() As Decimal
        Get
            Return _RMBasisWeight
        End Get
    End Property

    Public ReadOnly Property GrossWeight() As Decimal
        Get
            Return _GrossWeight
        End Get
    End Property

    Public ReadOnly Property CasesPerPalletLayer() As Integer
        Get
            Return _CasesPerPalletLayer
        End Get
    End Property

    Public ReadOnly Property CasesPerPalletLTL() As Integer
        Get
            Return _CasesPerPalletLTL
        End Get
    End Property

    Public ReadOnly Property CasesPerPalletTL() As Integer
        Get
            Return _CasesPerPalletTL
        End Get
    End Property

    Public ReadOnly Property PalletsPerTruck() As Integer
        Get
            Return _PalletsPerTruck
        End Get
    End Property

    Public ReadOnly Property CasesPerTruck() As Integer
        Get
            Return _CasesPerTruck
        End Get
    End Property

    Public ReadOnly Property ShippingClass() As String
        Get
            Return _ShippingClass
        End Get
    End Property

    ' This constructor will only be used if someone wishes to access
    ' the generic READ method to get a list of all parts
    Public Sub New()

    End Sub

    Public Sub New(ByVal passPart As String)
        Read(passPart)
    End Sub

    Public Sub Add(ByVal PRTNUM As String, ByVal Color As Integer, ByVal Pattern As Integer, ByVal Grade As Integer, ByVal Fold As Integer, ByVal PackageType As Integer, ByVal CustomerType As Integer, ByVal Label As Integer, ByVal CrossPerf As Boolean, ByVal CenterPerf As Boolean, ByVal Nascar As Boolean, ByVal SlitWidth As Decimal, ByVal FGBasisWeight As Decimal, ByVal RMBasisWeight As Decimal, ByVal PalletQuantity As Integer, ByVal EngineeringNotes As String, ByVal DueDaysID As Integer)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, PRTNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Color", SqlDbType.SmallInt, 0, Color, ParameterDirection.Input)
        oSQL.AddParameter("@Pattern", SqlDbType.SmallInt, 0, Pattern, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)
        oSQL.AddParameter("@FoldType", SqlDbType.SmallInt, 0, Fold, ParameterDirection.Input)
        oSQL.AddParameter("@PackageType", SqlDbType.SmallInt, 0, PackageType, ParameterDirection.Input)
        oSQL.AddParameter("@CustomerType", SqlDbType.SmallInt, 0, CustomerType, ParameterDirection.Input)
        oSQL.AddParameter("@Label", SqlDbType.SmallInt, 0, LabelType, ParameterDirection.Input)
        oSQL.AddParameter("@CrossPerf", SqlDbType.Bit, 0, CrossPerf, ParameterDirection.Input)
        oSQL.AddParameter("@CenterPerf", SqlDbType.Bit, 0, CenterPerf, ParameterDirection.Input)
        oSQL.AddParameter("@Nascar", SqlDbType.Bit, 0, Nascar, ParameterDirection.Input)
        oSQL.AddParameter("@SlitWidth", SqlDbType.Float, 0, SlitWidth, ParameterDirection.Input)
        oSQL.AddParameter("@RMBasisWeight", SqlDbType.Float, 0, RMBasisWeight, ParameterDirection.Input)
        oSQL.AddParameter("@FGBasisWeight", SqlDbType.Float, 0, FGBasisWeight, ParameterDirection.Input)
        oSQL.AddParameter("@PalletQuantity", SqlDbType.Int, 0, PalletQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@EngineeringNotes", SqlDbType.Text, 0, EngineeringNotes, ParameterDirection.Input)
        oSQL.AddParameter("@DueDaysID", SqlDbType.SmallInt, 0, DueDaysID, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddPartCharacteristics")
    End Sub

    Public Sub Delete(ByVal PRTNUM As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, PRTNUM, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("DeletePartCharacteristics")
    End Sub

    Private Sub Read(ByVal passpart As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, passpart, ParameterDirection.Input)

        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadPartCharacteristicsRecord")

        ' Assign the variables from the database to properties
        If dr.Read() Then
            ' New cahracteristics
            _Found = True
            _RMBasisWeight = dr("RMBasisWeight")
            _FGBasisWeight = dr("FGBasisWeight")
            _CaseHeight = dr("CaseHeight")
            _CaseLength = dr("CaseLength")
            _CaseWidth = dr("CaseWidth")
            _Color = dr("ColorID")
            _ColorDescription = dr("ColorDescription")
            _cPerf = dr("CPERF")
            _CustomerType = dr("CustomerType")
            _CustomerTypeDescription = dr("CustomerTypeDescription")
            _DueDaysID = dr("DueDaysID")
            _Fold = dr("FoldID")
            _FoldDescription = dr("FoldDescription")
            _Grade = dr("GradeID")
            _GradeDescription = dr("GradeDescription")
            _LabelType = dr("LabelID")
            _LabelDescription = dr("LabelDescription")
            _Nascar = dr("NASCAR")
            _PackageType = dr("PackagingID")
            _PackageDescription = dr("PackageDescription")
            _Pattern = dr("PatternID")
            _PatternDescription = dr("PatternDescription")
            _SlitWidth = dr("SlitWidth")
            _xPerf = dr("XPERF")
            _PalletQuantity = dr("PalletQuantity")
            _EngineeringNotes = dr("EngineeringNotes")
            _GrossWeight = dr("GrossWeight")
            _CasesPerPalletLTL = dr("CasesPerPalletLTL")
            _CasesPerPalletTL = dr("CasesPerPalletTL")
            _PalletsPerTruck = dr("PalletsPerTruck")
            _CasesPerTruck = dr("CasesPerTruck")
            _ShippingClass = dr("ShippingClass")
            _EachesPerMfgUnit = dr("EachesPerMfgUnit")
            _CasesPerPalletLayer = dr("CasesPerPalletLayer")
        Else
            _Found = False
            _RMBasisWeight = 0
            _FGBasisWeight = 0
            _CaseHeight = 0
            _CaseLength = 0
            _CaseWidth = 0
            _Color = 99
            _ColorDescription = ""
            _cPerf = False
            _CustomerType = 1
            _CustomerTypeDescription = ""
            _DueDaysID = 1
            _Fold = 99
            _FoldDescription = ""
            _Grade = 99
            _GradeDescription = ""
            _LabelType = 99
            _LabelDescription = ""
            _Nascar = False
            _PackageType = 99
            _PackageDescription = ""
            _Pattern = 99
            _PatternDescription = ""
            _SlitWidth = 0
            _xPerf = False
            _PalletQuantity = 0
            _EngineeringNotes = ""
            _GrossWeight = 0
            _CasesPerPalletLTL = 0
            _CasesPerPalletTL = 0
            _PalletsPerTruck = 0
            _CasesPerTruck = 0
            _ShippingClass = ""
            _EachesPerMfgUnit = 0
            _CasesPerPalletLayer = 0
        End If

        dr.Close()
        dr = Nothing
    End Sub

    Public Sub Update(ByVal PRTNUM As String, ByVal Color As Integer, ByVal Pattern As Integer, ByVal Grade As Integer, ByVal Fold As Integer, ByVal PackageType As Integer, ByVal CustomerType As Integer, ByVal Label As Integer, ByVal CrossPerf As Boolean, ByVal CenterPerf As Boolean, ByVal Nascar As Boolean, ByVal SlitWidth As Decimal, ByVal FGBasisWeight As Decimal, ByVal RMBasisWeight As Decimal, ByVal PalletQuantity As Integer, ByVal EngineeringNotes As String, ByVal DueDaysID As Integer)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, PRTNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Color", SqlDbType.SmallInt, 0, Color, ParameterDirection.Input)
        oSQL.AddParameter("@Pattern", SqlDbType.SmallInt, 0, Pattern, ParameterDirection.Input)
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)
        oSQL.AddParameter("@FoldType", SqlDbType.SmallInt, 0, Fold, ParameterDirection.Input)
        oSQL.AddParameter("@PackageType", SqlDbType.SmallInt, 0, PackageType, ParameterDirection.Input)
        oSQL.AddParameter("@CustomerType", SqlDbType.SmallInt, 0, CustomerType, ParameterDirection.Input)
        oSQL.AddParameter("@Label", SqlDbType.SmallInt, 0, Label, ParameterDirection.Input)
        oSQL.AddParameter("@SlitWidth", SqlDbType.Float, 0, SlitWidth, ParameterDirection.Input)
        oSQL.AddParameter("@RMBasisWeight", SqlDbType.Float, 0, RMBasisWeight, ParameterDirection.Input)
        oSQL.AddParameter("@FGBasisWeight", SqlDbType.Float, 0, FGBasisWeight, ParameterDirection.Input)
        oSQL.AddParameter("@CrossPerf", SqlDbType.Bit, 0, CrossPerf, ParameterDirection.Input)
        oSQL.AddParameter("@CenterPerf", SqlDbType.Bit, 0, CenterPerf, ParameterDirection.Input)
        oSQL.AddParameter("@Nascar", SqlDbType.Bit, 0, Nascar, ParameterDirection.Input)
        oSQL.AddParameter("@PalletQuantity", SqlDbType.SmallInt, 0, PalletQuantity, ParameterDirection.Input)
        oSQL.AddParameter("@EngineeringNotes", SqlDbType.Text, 0, EngineeringNotes, ParameterDirection.Input)
        oSQL.AddParameter("@DueDaysID", SqlDbType.SmallInt, 0, DueDaysID, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdatePartCharacteristics")
    End Sub
End Class

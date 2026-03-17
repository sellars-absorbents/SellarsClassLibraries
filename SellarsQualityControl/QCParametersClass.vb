Imports Sellars.SQL
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient

Public Class QCParametersClass
    Private _ConnectionString As String = ""

    Private _Found As Boolean = False
    Private _Grade As Integer = 0

    Private _WeightGraphMin As Decimal = 0
    Private _WeightLRL As Decimal = 0
    Private _WeightLCL As Decimal = 0
    Private _WeightUCL As Decimal = 0
    Private _WeightURL As Decimal = 0
    Private _WeightGraphMax As Decimal = 0
    Private _WeightObjective As Decimal = 0

    Private _OpSideBulkGraphMin As Decimal = 0
    Private _OpSideBulkLRL As Decimal = 0
    Private _OpSideBulkLCL As Decimal = 0
    Private _OpSideBulkUCL As Decimal = 0
    Private _OpSideBulkURL As Decimal = 0
    Private _OpSideBulkGraphMax As Decimal = 0
    Private _OpSideBulkObjective As Decimal = 0

    Private _DriveSideBulkGraphMin As Decimal = 0
    Private _DriveSideBulkLRL As Decimal = 0
    Private _DriveSideBulkLCL As Decimal = 0
    Private _DriveSideBulkUCL As Decimal = 0
    Private _DriveSideBulkURL As Decimal = 0
    Private _DriveSideBulkGraphMax As Decimal = 0
    Private _DriveSideBulkObjective As Decimal = 0

    Private _MDTGraphMin As Decimal = 0
    Private _MDTLRL As Decimal = 0
    Private _MDTLCL As Decimal = 0
    Private _MDTUCL As Decimal = 0
    Private _MDTURL As Decimal = 0
    Private _MDTGraphMax As Decimal = 0
    Private _MDTObjective As Decimal = 0

    Private _MDTEGraphMin As Decimal = 0
    Private _MDTELRL As Decimal = 0
    Private _MDTELCL As Decimal = 0
    Private _MDTEUCL As Decimal = 0
    Private _MDTEURL As Decimal = 0
    Private _MDTEGraphMax As Decimal = 0
    Private _MDTEObjective As Decimal = 0

    Private _CDTDGraphMin As Decimal = 0
    Private _CDTDLRL As Decimal = 0
    Private _CDTDLCL As Decimal = 0
    Private _CDTDUCL As Decimal = 0
    Private _CDTDURL As Decimal = 0
    Private _CDTDGraphMax As Decimal = 0
    Private _CDTDObjective As Decimal = 0

    Private _CDTWGraphMin As Decimal = 0
    Private _CDTWLRL As Decimal = 0
    Private _CDTWLCL As Decimal = 0
    Private _CDTWUCL As Decimal = 0
    Private _CDTWURL As Decimal = 0
    Private _CDTWGraphMax As Decimal = 0
    Private _CDTWObjective As Decimal = 0

    Private _CDTCGraphMin As Decimal = 0
    Private _CDTCLRL As Decimal = 0
    Private _CDTCLCL As Decimal = 0
    Private _CDTCUCL As Decimal = 0
    Private _CDTCURL As Decimal = 0
    Private _CDTCGraphMax As Decimal = 0
    Private _CDTCObjective As Decimal = 0

    Private _TTGraphMin As Decimal = 0
    Private _TTLRL As Decimal = 0
    Private _TTLCL As Decimal = 0
    Private _TTUCL As Decimal = 0
    Private _TTURL As Decimal = 0
    Private _TTGraphMax As Decimal = 0
    Private _TTObjective As Decimal = 0

    Private _ZPeelGraphMin As Decimal = 0
    Private _ZPeelLRL As Decimal = 0
    Private _ZPeelLCL As Decimal = 0
    Private _ZPeelUCL As Decimal = 0
    Private _ZPeelURL As Decimal = 0
    Private _ZPeelGraphMax As Decimal = 0
    Private _ZPeelObjective As Decimal = 0

    Private _ContaminationGraphMin As Decimal = 0
    Private _ContaminationLRL As Decimal = 0
    Private _ContaminationLCL As Decimal = 0
    Private _ContaminationUCL As Decimal = 0
    Private _ContaminationURL As Decimal = 0
    Private _ContaminationGraphMax As Decimal = 0
    Private _ContaminationObjective As Decimal = 0

    Private _TWAGraphMin As Decimal = 0
    Private _TWALRL As Decimal = 0
    Private _TWALCL As Decimal = 0
    Private _TWAUCL As Decimal = 0
    Private _TWAURL As Decimal = 0
    Private _TWAGraphMax As Decimal = 0
    Private _TWAObjective As Decimal = 0

    Private _AbsorbRateGraphMin As Decimal = 0
    Private _AbsorbRateLRL As Decimal = 0
    Private _AbsorbRateLCL As Decimal = 0
    Private _AbsorbRateUCL As Decimal = 0
    Private _AbsorbRateURL As Decimal = 0
    Private _AbsorbRateGraphMax As Decimal = 0
    Private _AbsorbRateObjective As Decimal = 0

    Private _RWidthGraphMin As Decimal = 0
    Private _RWidthLRL As Decimal = 0
    Private _RWidthLCL As Decimal = 0
    Private _RWidthUCL As Decimal = 0
    Private _RWidthURL As Decimal = 0
    Private _RWidthGraphMax As Decimal = 0
    Private _RWidthObjective As Decimal = 0

    Private _BPWidthGraphMin As Decimal = 0
    Private _BPWidthLRL As Decimal = 0
    Private _BPWidthLCL As Decimal = 0
    Private _BPWidthUCL As Decimal = 0
    Private _BPWidthURL As Decimal = 0
    Private _BPWidthGraphMax As Decimal = 0
    Private _BPWidthObjective As Decimal = 0

    Private _PcntcGraphMin As Decimal = 0
    Private _PcntcLRL As Decimal = 0
    Private _PcntcLCL As Decimal = 0
    Private _PcntcUCL As Decimal = 0
    Private _PcntcURL As Decimal = 0
    Private _PcntcGraphMax As Decimal = 0
    Private _PcntcObjective As Decimal = 0

    Private _WDGraphMin As Decimal = 0
    Private _WDLRL As Decimal = 0
    Private _WDLCL As Decimal = 0
    Private _WDUCL As Decimal = 0
    Private _WDURL As Decimal = 0
    Private _WDGraphMax As Decimal = 0
    Private _WDObjective As Decimal = 0

    Public Enum FieldTypeEnum
        Unassigned = 0
        Weight = 1
        OpSideBulk = 2
        DriveSideBulk = 3
        MDT = 4
        MDTE = 5
        CDTD = 6
        CDTW = 7
        CDTC = 8
        TT = 9
        ZPeel = 10
        Contamination = 11
        TWA = 12
        Pcntc = 13
        WD = 14
        AbsorbRate = 15
        RWidth = 16
        BPWidth = 17
    End Enum

    Public Structure QCSpecificationsStructure
        Public GraphMin As Decimal
        Public LRL As Decimal
        Public LCL As Decimal
        Public UCL As Decimal
        Public URL As Decimal
        Public GraphMax As Decimal
        Public Objective As Decimal
    End Structure

    Public Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
        Set(ByVal Value As String)
            _ConnectionString = Value
        End Set
    End Property

    Public ReadOnly Property Found() As Boolean
        Get
            Return _Found
        End Get
    End Property

    Public Property Grade() As Integer
        Get
            Return _Grade
        End Get
        Set(ByVal Value As Integer)
            _Grade = Value
        End Set
    End Property

    Public Property OpSideBulkGraphMin() As Decimal
        Get
            Return _OpSideBulkGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkGraphMin = Value
        End Set
    End Property

    Public Property OpSideBulkLRL() As Decimal
        Get
            Return _OpSideBulkLRL
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkLRL = Value
        End Set
    End Property

    Public Property OpSideBulkLCL() As Decimal
        Get
            Return _OpSideBulkLCL
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkLCL = Value
        End Set
    End Property

    Public Property OpSideBulkUCL() As Decimal
        Get
            Return _OpSideBulkUCL
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkUCL = Value
        End Set
    End Property

    Public Property OpSideBulkURL() As Decimal
        Get
            Return _OpSideBulkURL
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkURL = Value
        End Set
    End Property

    Public Property OpSideBulkGraphMax() As Decimal
        Get
            Return _OpSideBulkGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkGraphMax = Value
        End Set
    End Property

    Public Property OpSideBulkObjective() As Decimal
        Get
            Return _OpSideBulkObjective
        End Get
        Set(ByVal Value As Decimal)
            _OpSideBulkObjective = Value
        End Set
    End Property

    Public Property DriveSideBulkGraphMin() As Decimal
        Get
            Return _DriveSideBulkGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkGraphMin = Value
        End Set
    End Property

    Public Property DriveSideBulkLRL() As Decimal
        Get
            Return _DriveSideBulkLRL
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkLRL = Value
        End Set
    End Property

    Public Property DriveSideBulkLCL() As Decimal
        Get
            Return _DriveSideBulkLCL
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkLCL = Value
        End Set
    End Property

    Public Property DriveSideBulkUCL() As Decimal
        Get
            Return _DriveSideBulkUCL
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkUCL = Value
        End Set
    End Property

    Public Property DriveSideBulkURL() As Decimal
        Get
            Return _DriveSideBulkURL
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkURL = Value
        End Set
    End Property

    Public Property DriveSideBulkGraphMax() As Decimal
        Get
            Return _DriveSideBulkGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkGraphMax = Value
        End Set
    End Property

    Public Property DriveSideBulkObjective() As Decimal
        Get
            Return _DriveSideBulkObjective
        End Get
        Set(ByVal Value As Decimal)
            _DriveSideBulkObjective = Value
        End Set
    End Property
    Public Property CDTCGraphMin() As Decimal
        Get
            Return _CDTCGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _CDTCGraphMin = Value
        End Set
    End Property

    Public Property CDTCLRL() As Decimal
        Get
            Return _CDTCLRL
        End Get
        Set(ByVal Value As Decimal)
            _CDTCLRL = Value
        End Set
    End Property

    Public Property CDTCLCL() As Decimal
        Get
            Return _CDTCLCL
        End Get
        Set(ByVal Value As Decimal)
            _CDTCLCL = Value
        End Set
    End Property

    Public Property CDTCUCL() As Decimal
        Get
            Return _CDTCUCL
        End Get
        Set(ByVal Value As Decimal)
            _CDTCUCL = Value
        End Set
    End Property

    Public Property CDTCURL() As Decimal
        Get
            Return _CDTCURL
        End Get
        Set(ByVal Value As Decimal)
            _CDTCURL = Value
        End Set
    End Property

    Public Property CDTCGraphMax() As Decimal
        Get
            Return _CDTCGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _CDTCGraphMax = Value
        End Set
    End Property

    Public Property CDTCObjective() As Decimal
        Get
            Return _CDTCObjective
        End Get
        Set(ByVal Value As Decimal)
            _CDTCObjective = Value
        End Set
    End Property

    Public Property CDTDGraphMin() As Decimal
        Get
            Return _CDTDGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _CDTDGraphMin = Value
        End Set
    End Property

    Public Property CDTDLRL() As Decimal
        Get
            Return _CDTDLRL
        End Get
        Set(ByVal Value As Decimal)
            _CDTDLRL = Value
        End Set
    End Property

    Public Property CDTDLCL() As Decimal
        Get
            Return _CDTDLCL
        End Get
        Set(ByVal Value As Decimal)
            _CDTDLCL = Value
        End Set
    End Property

    Public Property CDTDUCL() As Decimal
        Get
            Return _CDTDUCL
        End Get
        Set(ByVal Value As Decimal)
            _CDTDUCL = Value
        End Set
    End Property

    Public Property CDTDURL() As Decimal
        Get
            Return _CDTDURL
        End Get
        Set(ByVal Value As Decimal)
            _CDTDURL = Value
        End Set
    End Property

    Public Property CDTDGraphMax() As Decimal
        Get
            Return _CDTDGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _CDTDGraphMax = Value
        End Set
    End Property

    Public Property CDTDObjective() As Decimal
        Get
            Return _CDTDObjective
        End Get
        Set(ByVal Value As Decimal)
            _CDTDObjective = Value
        End Set
    End Property

    Public Property CDTWGraphMin() As Decimal
        Get
            Return _CDTWGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _CDTWGraphMin = Value
        End Set
    End Property

    Public Property CDTWLRL() As Decimal
        Get
            Return _CDTWLRL
        End Get
        Set(ByVal Value As Decimal)
            _CDTWLRL = Value
        End Set
    End Property

    Public Property CDTWLCL() As Decimal
        Get
            Return _CDTWLCL
        End Get
        Set(ByVal Value As Decimal)
            _CDTWLCL = Value
        End Set
    End Property

    Public Property CDTWURL() As Decimal
        Get
            Return _CDTWURL
        End Get
        Set(ByVal Value As Decimal)
            _CDTWURL = Value
        End Set
    End Property

    Public Property CDTWUCL() As Decimal
        Get
            Return _CDTWUCL
        End Get
        Set(ByVal Value As Decimal)
            _CDTWUCL = Value
        End Set
    End Property

    Public Property CDTWGraphMax() As Decimal
        Get
            Return _CDTWGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _CDTWGraphMax = Value
        End Set
    End Property

    Public Property CDTWObjective() As Decimal
        Get
            Return _CDTWObjective
        End Get
        Set(ByVal Value As Decimal)
            _CDTWObjective = Value
        End Set
    End Property

    Public Property MDTGraphMin() As Decimal
        Get
            Return _MDTGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _MDTGraphMin = Value
        End Set
    End Property

    Public Property MDTLRL() As Decimal
        Get
            Return _MDTLRL
        End Get
        Set(ByVal Value As Decimal)
            _MDTLRL = Value
        End Set
    End Property

    Public Property MDTLCL() As Decimal
        Get
            Return _MDTLCL
        End Get
        Set(ByVal Value As Decimal)
            _MDTLCL = Value
        End Set
    End Property

    Public Property MDTUCL() As Decimal
        Get
            Return _MDTUCL
        End Get
        Set(ByVal Value As Decimal)
            _MDTUCL = Value
        End Set
    End Property

    Public Property MDTURL() As Decimal
        Get
            Return _MDTURL
        End Get
        Set(ByVal Value As Decimal)
            _MDTURL = Value
        End Set
    End Property

    Public Property MDTGraphMax() As Decimal
        Get
            Return _MDTGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _MDTGraphMax = Value
        End Set
    End Property

    Public Property MDTObjective() As Decimal
        Get
            Return _MDTObjective
        End Get
        Set(ByVal Value As Decimal)
            _MDTObjective = Value
        End Set
    End Property

    Public Property MDTEGraphMin() As Decimal
        Get
            Return _MDTEGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _MDTEGraphMin = Value
        End Set
    End Property

    Public Property MDTELRL() As Decimal
        Get
            Return _MDTELRL
        End Get
        Set(ByVal Value As Decimal)
            _MDTELRL = Value
        End Set
    End Property

    Public Property MDTELCL() As Decimal
        Get
            Return _MDTELCL
        End Get
        Set(ByVal Value As Decimal)
            _MDTELCL = Value
        End Set
    End Property

    Public Property MDTEUCL() As Decimal
        Get
            Return _MDTEUCL
        End Get
        Set(ByVal Value As Decimal)
            _MDTEUCL = Value
        End Set
    End Property

    Public Property MDTEURL() As Decimal
        Get
            Return _MDTEURL
        End Get
        Set(ByVal Value As Decimal)
            _MDTEURL = Value
        End Set
    End Property

    Public Property MDTEGraphMax() As Decimal
        Get
            Return _MDTEGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _MDTEGraphMax = Value
        End Set
    End Property

    Public Property MDTEObjective() As Decimal
        Get
            Return _MDTEObjective
        End Get
        Set(ByVal Value As Decimal)
            _MDTEObjective = Value
        End Set
    End Property

    Public Property TTGraphMin() As Decimal
        Get
            Return _TTGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _TTGraphMin = Value
        End Set
    End Property

    Public Property TTLRL() As Decimal
        Get
            Return _TTLRL
        End Get
        Set(ByVal Value As Decimal)
            _TTLRL = Value
        End Set
    End Property

    Public Property TTLCL() As Decimal
        Get
            Return _TTLCL
        End Get
        Set(ByVal Value As Decimal)
            _TTLCL = Value
        End Set
    End Property

    Public Property TTUCL() As Decimal
        Get
            Return _TTUCL
        End Get
        Set(ByVal Value As Decimal)
            _TTUCL = Value
        End Set
    End Property

    Public Property TTURL() As Decimal
        Get
            Return _TTURL
        End Get
        Set(ByVal Value As Decimal)
            _TTURL = Value
        End Set
    End Property

    Public Property TTGraphMax() As Decimal
        Get
            Return _TTGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _TTGraphMax = Value
        End Set
    End Property

    Public Property TTObjective() As Decimal
        Get
            Return _TTObjective
        End Get
        Set(ByVal Value As Decimal)
            _TTObjective = Value
        End Set
    End Property

    Public Property WeightGraphMin() As Decimal
        Get
            Return _WeightGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _WeightGraphMin = Value
        End Set
    End Property

    Public Property WeightLRL() As Decimal
        Get
            Return _WeightLRL
        End Get
        Set(ByVal Value As Decimal)
            _WeightLRL = Value
        End Set
    End Property

    Public Property WeightLCL() As Decimal
        Get
            Return _WeightLCL
        End Get
        Set(ByVal Value As Decimal)
            _WeightLCL = Value
        End Set
    End Property

    Public Property WeightUCL() As Decimal
        Get
            Return _WeightUCL
        End Get
        Set(ByVal Value As Decimal)
            _WeightUCL = Value
        End Set
    End Property

    Public Property WeightURL() As Decimal
        Get
            Return _WeightURL
        End Get
        Set(ByVal Value As Decimal)
            _WeightURL = Value
        End Set
    End Property

    Public Property WeightGraphMax() As Decimal
        Get
            Return _WeightGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _WeightGraphMax = Value
        End Set
    End Property

    Public Property WeightObjective() As Decimal
        Get
            Return _WeightObjective
        End Get
        Set(ByVal Value As Decimal)
            _WeightObjective = Value
        End Set
    End Property

    Public Property ZPeelGraphMin() As Decimal
        Get
            Return _ZPeelGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelGraphMin = Value
        End Set
    End Property

    Public Property ZPeelLRL() As Decimal
        Get
            Return _ZPeelLRL
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelLRL = Value
        End Set
    End Property

    Public Property ZPeelLCL() As Decimal
        Get
            Return _ZPeelLCL
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelLCL = Value
        End Set
    End Property

    Public Property ZPeelUCL() As Decimal
        Get
            Return _ZPeelUCL
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelUCL = Value
        End Set
    End Property

    Public Property ZPeelURL() As Decimal
        Get
            Return _ZPeelURL
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelURL = Value
        End Set
    End Property

    Public Property ZPeelGraphMax() As Decimal
        Get
            Return _ZPeelGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelGraphMax = Value
        End Set
    End Property

    Public Property ZPeelObjective() As Decimal
        Get
            Return _ZPeelObjective
        End Get
        Set(ByVal Value As Decimal)
            _ZPeelObjective = Value
        End Set
    End Property

    Public Property ContaminationGraphMin() As Decimal
        Get
            Return _ContaminationGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationGraphMin = Value
        End Set
    End Property

    Public Property ContaminationLRL() As Decimal
        Get
            Return _ContaminationLRL
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationLRL = Value
        End Set
    End Property

    Public Property ContaminationLCL() As Decimal
        Get
            Return _ContaminationLCL
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationLCL = Value
        End Set
    End Property

    Public Property ContaminationUCL() As Decimal
        Get
            Return _ContaminationUCL
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationUCL = Value
        End Set
    End Property

    Public Property ContaminationURL() As Decimal
        Get
            Return _ContaminationURL
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationURL = Value
        End Set
    End Property

    Public Property ContaminationGraphMax() As Decimal
        Get
            Return _ContaminationGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationGraphMax = Value
        End Set
    End Property

    Public Property ContaminationObjective() As Decimal
        Get
            Return _ContaminationObjective
        End Get
        Set(ByVal Value As Decimal)
            _ContaminationObjective = Value
        End Set
    End Property

    Public Property TWAGraphMin() As Decimal
        Get
            Return _TWAGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _TWAGraphMin = Value
        End Set
    End Property

    Public Property TWALRL() As Decimal
        Get
            Return _TWALRL
        End Get
        Set(ByVal Value As Decimal)
            _TWALRL = Value
        End Set
    End Property

    Public Property TWALCL() As Decimal
        Get
            Return _TWALCL
        End Get
        Set(ByVal Value As Decimal)
            _TWALCL = Value
        End Set
    End Property

    Public Property TWAUCL() As Decimal
        Get
            Return _TWAUCL
        End Get
        Set(ByVal Value As Decimal)
            _TWAUCL = Value
        End Set
    End Property

    Public Property TWAURL() As Decimal
        Get
            Return _TWAURL
        End Get
        Set(ByVal Value As Decimal)
            _TWAURL = Value
        End Set
    End Property

    Public Property TWAGraphMax() As Decimal
        Get
            Return _TWAGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _TWAGraphMax = Value
        End Set
    End Property

    Public Property TWAObjective() As Decimal
        Get
            Return _TWAObjective
        End Get
        Set(ByVal Value As Decimal)
            _TWAObjective = Value
        End Set
    End Property


    Public Property AbsorbRateGraphMin() As Decimal
        Get
            Return _AbsorbRateGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateGraphMin = Value
        End Set
    End Property

    Public Property AbsorbRateLRL() As Decimal
        Get
            Return _AbsorbRateLRL
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateLRL = Value
        End Set
    End Property

    Public Property AbsorbRateLCL() As Decimal
        Get
            Return _AbsorbRateLCL
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateLCL = Value
        End Set
    End Property

    Public Property AbsorbRateUCL() As Decimal
        Get
            Return _AbsorbRateUCL
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateUCL = Value
        End Set
    End Property

    Public Property AbsorbRateURL() As Decimal
        Get
            Return _AbsorbRateURL
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateURL = Value
        End Set
    End Property

    Public Property AbsorbRateGraphMax() As Decimal
        Get
            Return _AbsorbRateGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateGraphMax = Value
        End Set
    End Property

    Public Property AbsorbRateObjective() As Decimal
        Get
            Return _AbsorbRateObjective
        End Get
        Set(ByVal Value As Decimal)
            _AbsorbRateObjective = Value
        End Set
    End Property

    Public Property RWidthGraphMin() As Decimal
        Get
            Return _RWidthGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _RWidthGraphMin = Value
        End Set
    End Property

    Public Property RWidthLRL() As Decimal
        Get
            Return _RWidthLRL
        End Get
        Set(ByVal Value As Decimal)
            _RWidthLRL = Value
        End Set
    End Property

    Public Property RWidthLCL() As Decimal
        Get
            Return _RWidthLCL
        End Get
        Set(ByVal Value As Decimal)
            _RWidthLCL = Value
        End Set
    End Property

    Public Property RWidthUCL() As Decimal
        Get
            Return _RWidthUCL
        End Get
        Set(ByVal Value As Decimal)
            _RWidthUCL = Value
        End Set
    End Property

    Public Property RWidthURL() As Decimal
        Get
            Return _RWidthURL
        End Get
        Set(ByVal Value As Decimal)
            _RWidthURL = Value
        End Set
    End Property

    Public Property RWidthGraphMax() As Decimal
        Get
            Return _RWidthGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _RWidthGraphMax = Value
        End Set
    End Property

    Public Property RWidthObjective() As Decimal
        Get
            Return _RWidthObjective
        End Get
        Set(ByVal Value As Decimal)
            _RWidthObjective = Value
        End Set
    End Property

    Public Property BPWidthGraphMin() As Decimal
        Get
            Return _BPWidthGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthGraphMin = Value
        End Set
    End Property

    Public Property BPWidthLRL() As Decimal
        Get
            Return _BPWidthLRL
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthLRL = Value
        End Set
    End Property

    Public Property BPWidthLCL() As Decimal
        Get
            Return _BPWidthLCL
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthLCL = Value
        End Set
    End Property

    Public Property BPWidthUCL() As Decimal
        Get
            Return _BPWidthUCL
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthUCL = Value
        End Set
    End Property

    Public Property BPWidthURL() As Decimal
        Get
            Return _BPWidthURL
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthURL = Value
        End Set
    End Property

    Public Property BPWidthGraphMax() As Decimal
        Get
            Return _BPWidthGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthGraphMax = Value
        End Set
    End Property

    Public Property BPWidthObjective() As Decimal
        Get
            Return _BPWidthObjective
        End Get
        Set(ByVal Value As Decimal)
            _BPWidthObjective = Value
        End Set
    End Property

    Public Property PcntcGraphMin() As Decimal
        Get
            Return _PcntcGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _PcntcGraphMin = Value
        End Set
    End Property

    Public Property PcntcLRL() As Decimal
        Get
            Return _PcntcLRL
        End Get
        Set(ByVal Value As Decimal)
            _PcntcLRL = Value
        End Set
    End Property

    Public Property PcntcLCL() As Decimal
        Get
            Return _PcntcLCL
        End Get
        Set(ByVal Value As Decimal)
            _PcntcLCL = Value
        End Set
    End Property

    Public Property PcntcUCL() As Decimal
        Get
            Return _PcntcUCL
        End Get
        Set(ByVal Value As Decimal)
            _PcntcUCL = Value
        End Set
    End Property

    Public Property PcntcURL() As Decimal
        Get
            Return _PcntcURL
        End Get
        Set(ByVal Value As Decimal)
            _PcntcURL = Value
        End Set
    End Property

    Public Property PcntcGraphMax() As Decimal
        Get
            Return _PcntcGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _PcntcGraphMax = Value
        End Set
    End Property

    Public Property PcntcObjective() As Decimal
        Get
            Return _PcntcObjective
        End Get
        Set(ByVal Value As Decimal)
            _PcntcObjective = Value
        End Set
    End Property

    Public Property WDGraphMin() As Decimal
        Get
            Return _WDGraphMin
        End Get
        Set(ByVal Value As Decimal)
            _WDGraphMin = Value
        End Set
    End Property

    Public Property WDLRL() As Decimal
        Get
            Return _WDLRL
        End Get
        Set(ByVal Value As Decimal)
            _WDLRL = Value
        End Set
    End Property

    Public Property WDLCL() As Decimal
        Get
            Return _WDLCL
        End Get
        Set(ByVal Value As Decimal)
            _WDLCL = Value
        End Set
    End Property

    Public Property WDUCL() As Decimal
        Get
            Return _WDUCL
        End Get
        Set(ByVal Value As Decimal)
            _WDUCL = Value
        End Set
    End Property

    Public Property WDURL() As Decimal
        Get
            Return _WDURL
        End Get
        Set(ByVal Value As Decimal)
            _WDURL = Value
        End Set
    End Property

    Public Property WDGraphMax() As Decimal
        Get
            Return _WDGraphMax
        End Get
        Set(ByVal Value As Decimal)
            _WDGraphMax = Value
        End Set
    End Property

    Public Property WDObjective() As Decimal
        Get
            Return _WDObjective
        End Get
        Set(ByVal Value As Decimal)
            _WDObjective = Value
        End Set
    End Property

    Public Sub New(ByVal connection As String)
        ConnectionString = connection
    End Sub

    Public Sub Add()
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Assign all the necessary properties to an SQL Parameter
        Call AssignParameters(oSQL)

        ' Run the stored procedure
        oSQL.RunProc("AddQCSpecification")
    End Sub

    Private Sub AssignParameters(ByVal oSQL As SqlService)
        ' Add the parameters to the command object
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)

        oSQL.AddParameter("@WeightGraphMin", SqlDbType.Float, 0, WeightGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@WeightLRL", SqlDbType.Float, 0, WeightLRL, ParameterDirection.Input)
        oSQL.AddParameter("@WeightLCL", SqlDbType.Float, 0, WeightLCL, ParameterDirection.Input)
        oSQL.AddParameter("@WeightUCL", SqlDbType.Float, 0, WeightUCL, ParameterDirection.Input)
        oSQL.AddParameter("@WeightURL", SqlDbType.Float, 0, WeightURL, ParameterDirection.Input)
        oSQL.AddParameter("@WeightGraphMax", SqlDbType.Float, 0, WeightGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@WeightObjective", SqlDbType.Float, 0, WeightObjective, ParameterDirection.Input)

        oSQL.AddParameter("@OpSideBulkGraphMin", SqlDbType.Float, 0, OpSideBulkGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulkLRL", SqlDbType.Float, 0, OpSideBulkLRL, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulkLCL", SqlDbType.Float, 0, OpSideBulkLCL, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulkUCL", SqlDbType.Float, 0, OpSideBulkUCL, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulkURL", SqlDbType.Float, 0, OpSideBulkURL, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulkGraphMax", SqlDbType.Float, 0, OpSideBulkGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@OpSideBulkObjective", SqlDbType.Float, 0, OpSideBulkObjective, ParameterDirection.Input)

        oSQL.AddParameter("@DriveSideBulkGraphMin", SqlDbType.Float, 0, DriveSideBulkGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulkLRL", SqlDbType.Float, 0, DriveSideBulkLRL, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulkLCL", SqlDbType.Float, 0, DriveSideBulkLCL, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulkUCL", SqlDbType.Float, 0, DriveSideBulkUCL, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulkURL", SqlDbType.Float, 0, DriveSideBulkURL, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulkGraphMax", SqlDbType.Float, 0, DriveSideBulkGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@DriveSideBulkObjective", SqlDbType.Float, 0, DriveSideBulkObjective, ParameterDirection.Input)

        oSQL.AddParameter("@MDTGraphMin", SqlDbType.Float, 0, MDTGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@MDTLRL", SqlDbType.Float, 0, MDTLRL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTLCL", SqlDbType.Float, 0, MDTLCL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTUCL", SqlDbType.Float, 0, MDTUCL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTURL", SqlDbType.Float, 0, MDTURL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTGraphMax", SqlDbType.Float, 0, MDTGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@MDTObjective", SqlDbType.Float, 0, MDTObjective, ParameterDirection.Input)

        oSQL.AddParameter("@MDTEGraphMin", SqlDbType.Float, 0, MDTEGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@MDTELRL", SqlDbType.Float, 0, MDTELRL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTELCL", SqlDbType.Float, 0, MDTELCL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTEUCL", SqlDbType.Float, 0, MDTEUCL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTEURL", SqlDbType.Float, 0, MDTEURL, ParameterDirection.Input)
        oSQL.AddParameter("@MDTEGraphMax", SqlDbType.Float, 0, MDTEGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@MDTEObjective", SqlDbType.Float, 0, MDTEObjective, ParameterDirection.Input)

        oSQL.AddParameter("@CDTDGraphMin", SqlDbType.Float, 0, CDTDGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@CDTDLRL", SqlDbType.Float, 0, CDTDLRL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTDLCL", SqlDbType.Float, 0, CDTDLCL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTDUCL", SqlDbType.Float, 0, CDTDUCL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTDURL", SqlDbType.Float, 0, CDTDURL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTDGraphMax", SqlDbType.Float, 0, CDTDGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@CDTDObjective", SqlDbType.Float, 0, CDTDObjective, ParameterDirection.Input)

        oSQL.AddParameter("@CDTWGraphMin", SqlDbType.Float, 0, CDTWGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@CDTWLRL", SqlDbType.Float, 0, CDTWLRL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTWLCL", SqlDbType.Float, 0, CDTWLCL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTWUCL", SqlDbType.Float, 0, CDTWUCL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTWURL", SqlDbType.Float, 0, CDTWURL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTWGraphMax", SqlDbType.Float, 0, CDTWGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@CDTWObjective", SqlDbType.Float, 0, CDTWObjective, ParameterDirection.Input)

        oSQL.AddParameter("@CDTCGraphMin", SqlDbType.Float, 0, CDTCGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@CDTCLRL", SqlDbType.Float, 0, CDTCLRL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTCLCL", SqlDbType.Float, 0, CDTCLCL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTCUCL", SqlDbType.Float, 0, CDTCUCL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTCURL", SqlDbType.Float, 0, CDTCURL, ParameterDirection.Input)
        oSQL.AddParameter("@CDTCGraphMax", SqlDbType.Float, 0, CDTCGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@CDTCObjective", SqlDbType.Float, 0, CDTCObjective, ParameterDirection.Input)

        oSQL.AddParameter("@TTGraphMin", SqlDbType.Float, 0, TTGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@TTLRL", SqlDbType.Float, 0, TTLRL, ParameterDirection.Input)
        oSQL.AddParameter("@TTLCL", SqlDbType.Float, 0, TTLCL, ParameterDirection.Input)
        oSQL.AddParameter("@TTUCL", SqlDbType.Float, 0, TTUCL, ParameterDirection.Input)
        oSQL.AddParameter("@TTURL", SqlDbType.Float, 0, TTURL, ParameterDirection.Input)
        oSQL.AddParameter("@TTGraphMax", SqlDbType.Float, 0, TTGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@TTObjective", SqlDbType.Float, 0, TTObjective, ParameterDirection.Input)

        oSQL.AddParameter("@ZPeelGraphMin", SqlDbType.Float, 0, ZPeelGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeelLRL", SqlDbType.Float, 0, ZPeelLRL, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeelLCL", SqlDbType.Float, 0, ZPeelLCL, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeelUCL", SqlDbType.Float, 0, ZPeelUCL, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeelURL", SqlDbType.Float, 0, ZPeelURL, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeelGraphMax", SqlDbType.Float, 0, ZPeelGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@ZPeelObjective", SqlDbType.Float, 0, ZPeelObjective, ParameterDirection.Input)

        oSQL.AddParameter("@ContaminationGraphMin", SqlDbType.Float, 0, ContaminationGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@ContaminationLRL", SqlDbType.Float, 0, ContaminationLRL, ParameterDirection.Input)
        oSQL.AddParameter("@ContaminationLCL", SqlDbType.Float, 0, ContaminationLCL, ParameterDirection.Input)
        oSQL.AddParameter("@ContaminationUCL", SqlDbType.Float, 0, ContaminationUCL, ParameterDirection.Input)
        oSQL.AddParameter("@ContaminationURL", SqlDbType.Float, 0, ContaminationURL, ParameterDirection.Input)
        oSQL.AddParameter("@ContaminationGraphMax", SqlDbType.Float, 0, ContaminationGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@ContaminationObjective", SqlDbType.Float, 0, ContaminationObjective, ParameterDirection.Input)

        oSQL.AddParameter("@TWAGraphMin", SqlDbType.Float, 0, TWAGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@TWALRL", SqlDbType.Float, 0, TWALRL, ParameterDirection.Input)
        oSQL.AddParameter("@TWALCL", SqlDbType.Float, 0, TWALCL, ParameterDirection.Input)
        oSQL.AddParameter("@TWAUCL", SqlDbType.Float, 0, TWAUCL, ParameterDirection.Input)
        oSQL.AddParameter("@TWAURL", SqlDbType.Float, 0, TWAURL, ParameterDirection.Input)
        oSQL.AddParameter("@TWAGraphMax", SqlDbType.Float, 0, TWAGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@TWAObjective", SqlDbType.Float, 0, TWAObjective, ParameterDirection.Input)

        oSQL.AddParameter("@AbsorbRateGraphMin", SqlDbType.Float, 0, AbsorbRateGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@AbsorbRateLRL", SqlDbType.Float, 0, AbsorbRateLRL, ParameterDirection.Input)
        oSQL.AddParameter("@AbsorbRateLCL", SqlDbType.Float, 0, AbsorbRateLCL, ParameterDirection.Input)
        oSQL.AddParameter("@AbsorbRateUCL", SqlDbType.Float, 0, AbsorbRateUCL, ParameterDirection.Input)
        oSQL.AddParameter("@AbsorbRateURL", SqlDbType.Float, 0, AbsorbRateURL, ParameterDirection.Input)
        oSQL.AddParameter("@AbsorbRateGraphMax", SqlDbType.Float, 0, AbsorbRateGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@AbsorbRateObjective", SqlDbType.Float, 0, AbsorbRateObjective, ParameterDirection.Input)

        oSQL.AddParameter("@RWidthGraphMin", SqlDbType.Float, 0, RWidthGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@RWidthLRL", SqlDbType.Float, 0, RWidthLRL, ParameterDirection.Input)
        oSQL.AddParameter("@RWidthLCL", SqlDbType.Float, 0, RWidthLCL, ParameterDirection.Input)
        oSQL.AddParameter("@RWidthUCL", SqlDbType.Float, 0, RWidthUCL, ParameterDirection.Input)
        oSQL.AddParameter("@RWidthURL", SqlDbType.Float, 0, RWidthURL, ParameterDirection.Input)
        oSQL.AddParameter("@RWidthGraphMax", SqlDbType.Float, 0, RWidthGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@RWidthObjective", SqlDbType.Float, 0, RWidthObjective, ParameterDirection.Input)

        oSQL.AddParameter("@BPWidthGraphMin", SqlDbType.Float, 0, BPWidthGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@BPWidthLRL", SqlDbType.Float, 0, BPWidthLRL, ParameterDirection.Input)
        oSQL.AddParameter("@BPWidthLCL", SqlDbType.Float, 0, BPWidthLCL, ParameterDirection.Input)
        oSQL.AddParameter("@BPWidthUCL", SqlDbType.Float, 0, BPWidthUCL, ParameterDirection.Input)
        oSQL.AddParameter("@BPWidthURL", SqlDbType.Float, 0, BPWidthURL, ParameterDirection.Input)
        oSQL.AddParameter("@BPWidthGraphMax", SqlDbType.Float, 0, BPWidthGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@BPWidthObjective", SqlDbType.Float, 0, BPWidthObjective, ParameterDirection.Input)

        oSQL.AddParameter("@PcntcGraphMin", SqlDbType.Float, 0, PcntcGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@PcntcLRL", SqlDbType.Float, 0, PcntcLRL, ParameterDirection.Input)
        oSQL.AddParameter("@PcntcLCL", SqlDbType.Float, 0, PcntcLCL, ParameterDirection.Input)
        oSQL.AddParameter("@PcntcUCL", SqlDbType.Float, 0, PcntcUCL, ParameterDirection.Input)
        oSQL.AddParameter("@PcntcURL", SqlDbType.Float, 0, PcntcURL, ParameterDirection.Input)
        oSQL.AddParameter("@PcntcGraphMax", SqlDbType.Float, 0, PcntcGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@PcntcObjective", SqlDbType.Float, 0, PcntcObjective, ParameterDirection.Input)

        oSQL.AddParameter("@WDGraphMin", SqlDbType.Float, 0, WDGraphMin, ParameterDirection.Input)
        oSQL.AddParameter("@WDLRL", SqlDbType.Float, 0, WDLRL, ParameterDirection.Input)
        oSQL.AddParameter("@WDLCL", SqlDbType.Float, 0, WDLCL, ParameterDirection.Input)
        oSQL.AddParameter("@WDUCL", SqlDbType.Float, 0, WDUCL, ParameterDirection.Input)
        oSQL.AddParameter("@WDURL", SqlDbType.Float, 0, WDURL, ParameterDirection.Input)
        oSQL.AddParameter("@WDGraphMax", SqlDbType.Float, 0, WDGraphMax, ParameterDirection.Input)
        oSQL.AddParameter("@WDObjective", SqlDbType.Float, 0, WDObjective, ParameterDirection.Input)
    End Sub

    Public Sub Delete(ByVal Grade As Integer)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("DeleteQCSpecification")
    End Sub

    Public Function Read() As DataSet
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcDataSet("ReadAllQCSpecificationGrades", "Grades")
    End Function

    Public Sub Read(ByVal Grade As Integer)
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@Grade", SqlDbType.SmallInt, 0, Grade, ParameterDirection.Input)

        ' Execute the stored procedure
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadGradeParameters")
            ' Read the first record
            If Not dr.Read() Then
                _Found = False
                ' Assign the values
                _WeightGraphMin = 0
                _WeightLRL = 0
                _WeightLCL = 0
                _WeightUCL = 0
                _WeightURL = 0
                _WeightGraphMax = 0
                _WeightObjective = 0

                _OpSideBulkGraphMin = 0
                _OpSideBulkLRL = 0
                _OpSideBulkLCL = 0
                _OpSideBulkUCL = 0
                _OpSideBulkURL = 0
                _OpSideBulkGraphMax = 0
                _OpSideBulkObjective = 0

                _DriveSideBulkGraphMin = 0
                _DriveSideBulkLRL = 0
                _DriveSideBulkLCL = 0
                _DriveSideBulkUCL = 0
                _DriveSideBulkURL = 0
                _DriveSideBulkGraphMax = 0
                _DriveSideBulkObjective = 0

                _MDTGraphMin = 0
                _MDTLRL = 0
                _MDTLCL = 0
                _MDTUCL = 0
                _MDTURL = 0
                _MDTGraphMax = 0
                _MDTObjective = 0

                _MDTEGraphMin = 0
                _MDTELRL = 0
                _MDTELCL = 0
                _MDTEUCL = 0
                _MDTEURL = 0
                _MDTEGraphMax = 0
                _MDTEObjective = 0

                _CDTDGraphMin = 0
                _CDTDLRL = 0
                _CDTDLCL = 0
                _CDTDUCL = 0
                _CDTDURL = 0
                _CDTDGraphMax = 0
                _CDTDObjective = 0

                _CDTWGraphMin = 0
                _CDTWLRL = 0
                _CDTWLCL = 0
                _CDTWUCL = 0
                _CDTWURL = 0
                _CDTWGraphMax = 0
                _CDTWObjective = 0

                _CDTCGraphMin = 0
                _CDTCLRL = 0
                _CDTCLCL = 0
                _CDTCUCL = 0
                _CDTCURL = 0
                _CDTCGraphMax = 0
                _CDTCObjective = 0

                _TTGraphMin = 0
                _TTLRL = 0
                _TTLCL = 0
                _TTUCL = 0
                _TTURL = 0
                _TTGraphMax = 0
                _TTObjective = 0

                _ZPeelGraphMin = 0
                _ZPeelLRL = 0
                _ZPeelLCL = 0
                _ZPeelUCL = 0
                _ZPeelURL = 0
                _ZPeelGraphMax = 0
                _ZPeelObjective = 0

                _ContaminationGraphMin = 0
                _ContaminationLRL = 0
                _ContaminationLCL = 0
                _ContaminationUCL = 0
                _ContaminationURL = 0
                _ContaminationGraphMax = 0
                _ContaminationObjective = 0

                _TWAGraphMin = 0
                _TWALRL = 0
                _TWALCL = 0
                _TWAUCL = 0
                _TWAURL = 0
                _TWAGraphMax = 0
                _TWAObjective = 0

                _AbsorbRateGraphMin = 0
                _AbsorbRateLRL = 0
                _AbsorbRateLCL = 0
                _AbsorbRateUCL = 0
                _AbsorbRateURL = 0
                _AbsorbRateGraphMax = 0
                _AbsorbRateObjective = 0

                _RWidthGraphMin = 0
                _RWidthLRL = 0
                _RWidthLCL = 0
                _RWidthUCL = 0
                _RWidthURL = 0
                _RWidthGraphMax = 0
                _RWidthObjective = 0

                _BPWidthGraphMin = 0
                _BPWidthLRL = 0
                _BPWidthLCL = 0
                _BPWidthUCL = 0
                _BPWidthURL = 0
                _BPWidthGraphMax = 0
                _BPWidthObjective = 0

                _PcntcGraphMin = 0
                _PcntcLRL = 0
                _PcntcLCL = 0
                _PcntcUCL = 0
                _PcntcURL = 0
                _PcntcGraphMax = 0
                _PcntcObjective = 0

                _WDGraphMin = 0
                _WDLRL = 0
                _WDLCL = 0
                _WDUCL = 0
                _WDURL = 0
                _WDGraphMax = 0
                _WDObjective = 0
            Else
                _Found = True
                ' Assign the values
                _WeightGraphMin = dr("WeightGraphMin")
                _WeightLRL = dr("WeightLRL")
                _WeightLCL = dr("WeightLCL")
                _WeightUCL = dr("WeightUCL")
                _WeightURL = dr("WeightURL")
                _WeightGraphMax = dr("WeightGraphMax")
                _WeightObjective = IIf(IsDBNull(dr("WeightObjective")), 0, dr("WeightObjective"))

                _OpSideBulkGraphMin = dr("OpSideBulkGraphMin")
                _OpSideBulkLRL = dr("OpSideBulkLRL")
                _OpSideBulkLCL = dr("OpSideBulkLCL")
                _OpSideBulkUCL = dr("OpSideBulkUCL")
                _OpSideBulkURL = dr("OpSideBulkURL")
                _OpSideBulkGraphMax = dr("OpSideBulkGraphMax")
                _OpSideBulkObjective = IIf(IsDBNull(dr("OpSideBulkObjective")), 0, dr("OpSideBulkObjective"))

                _DriveSideBulkGraphMin = dr("DriveSideBulkGraphMin")
                _DriveSideBulkLRL = dr("DriveSideBulkLRL")
                _DriveSideBulkLCL = dr("DriveSideBulkLCL")
                _DriveSideBulkUCL = dr("DriveSideBulkUCL")
                _DriveSideBulkURL = dr("DriveSideBulkURL")
                _DriveSideBulkGraphMax = dr("DriveSideBulkGraphMax")
                _DriveSideBulkObjective = IIf(IsDBNull(dr("DriveSideBulkObjective")), 0, dr("DriveSideBulkObjective"))

                _MDTGraphMin = dr("MDTGraphMin")
                _MDTLRL = dr("MDTLRL")
                _MDTLCL = dr("MDTLCL")
                _MDTUCL = dr("MDTUCL")
                _MDTURL = dr("MDTURL")
                _MDTGraphMax = dr("MDTGraphMax")
                _MDTObjective = IIf(IsDBNull(dr("MDTObjective")), 0, dr("MDTObjective"))

                _MDTEGraphMin = dr("MDTEGraphMin")
                _MDTELRL = dr("MDTELRL")
                _MDTELCL = dr("MDTELCL")
                _MDTEUCL = dr("MDTEUCL")
                _MDTEURL = dr("MDTEURL")
                _MDTEGraphMax = dr("MDTEGraphMax")
                _MDTEObjective = IIf(IsDBNull(dr("MDTEObjective")), 0, dr("MDTEObjective"))

                _CDTDGraphMin = dr("CDTDGraphMin")
                _CDTDLRL = dr("CDTDLRL")
                _CDTDLCL = dr("CDTDLCL")
                _CDTDUCL = dr("CDTDUCL")
                _CDTDURL = dr("CDTDURL")
                _CDTDGraphMax = dr("CDTDGraphMax")
                _CDTDObjective = IIf(IsDBNull(dr("CDTDObjective")), 0, dr("CDTDObjective"))

                _CDTWGraphMin = dr("CDTWGraphMin")
                _CDTWLRL = dr("CDTWLRL")
                _CDTWLCL = dr("CDTWLCL")
                _CDTWUCL = dr("CDTWUCL")
                _CDTWURL = dr("CDTWURL")
                _CDTWGraphMax = dr("CDTWGraphMax")
                _CDTWObjective = IIf(IsDBNull(dr("CDTWObjective")), 0, dr("CDTWObjective"))

                _CDTCGraphMin = dr("CDTCGraphMin")
                _CDTCLRL = dr("CDTCLRL")
                _CDTCLCL = dr("CDTCLCL")
                _CDTCUCL = dr("CDTCUCL")
                _CDTCURL = dr("CDTCURL")
                _CDTCGraphMax = dr("CDTCGraphMax")
                _CDTCObjective = IIf(IsDBNull(dr("CDTCObjective")), 0, dr("CDTCObjective"))

                _TTGraphMin = dr("TTGraphMin")
                _TTLRL = dr("TTLRL")
                _TTLCL = dr("TTLCL")
                _TTUCL = dr("TTUCL")
                _TTURL = dr("TTURL")
                _TTGraphMax = dr("TTGraphMax")
                _TTObjective = IIf(IsDBNull(dr("TTObjective")), 0, dr("TTObjective"))

                _ZPeelGraphMin = IIf(IsDBNull(dr("ZPeelGraphMin")), 0, dr("ZPeelGraphMin"))
                _ZPeelLRL = IIf(IsDBNull(dr("ZPeelLRL")), 0, dr("ZPeelLRL"))
                _ZPeelLCL = IIf(IsDBNull(dr("ZPeelLCL")), 0, dr("ZPeelLCL"))
                _ZPeelUCL = IIf(IsDBNull(dr("ZPeelUCL")), 0, dr("ZPeelUCL"))
                _ZPeelURL = IIf(IsDBNull(dr("ZPeelURL")), 0, dr("ZPeelURL"))
                _ZPeelGraphMax = IIf(IsDBNull(dr("ZPeelGraphMax")), 0, dr("ZPeelGraphMax"))
                _ZPeelObjective = IIf(IsDBNull(dr("ZPeelObjective")), 0, dr("ZPeelObjective"))

                _ContaminationGraphMin = dr("ContaminationGraphMin")
                _ContaminationLRL = dr("ContaminationLRL")
                _ContaminationLCL = dr("ContaminationLCL")
                _ContaminationUCL = dr("ContaminationUCL")
                _ContaminationURL = dr("ContaminationURL")
                _ContaminationGraphMax = dr("ContaminationGraphMax")
                _ContaminationObjective = IIf(IsDBNull(dr("ContaminationObjective")), 0, dr("ContaminationObjective"))

                _TWAGraphMin = dr("TWAGraphMin")
                _TWALRL = dr("TWALRL")
                _TWALCL = dr("TWALCL")
                _TWAUCL = dr("TWAUCL")
                _TWAURL = dr("TWAURL")
                _TWAGraphMax = dr("TWAGraphMax")
                _TWAObjective = IIf(IsDBNull(dr("TWAObjective")), 0, dr("TWAObjective"))

                _AbsorbRateGraphMin = dr("AbsorbRateGraphMin")
                _AbsorbRateLRL = dr("AbsorbRateLRL")
                _AbsorbRateLCL = dr("AbsorbRateLCL")
                _AbsorbRateUCL = dr("AbsorbRateUCL")
                _AbsorbRateURL = dr("AbsorbRateURL")
                _AbsorbRateGraphMax = dr("AbsorbRateGraphMax")
                _AbsorbRateObjective = IIf(IsDBNull(dr("AbsorbRateObjective")), 0, dr("AbsorbRateObjective"))

                _RWidthGraphMin = dr("RWidthGraphMin")
                _RWidthLRL = dr("RWidthLRL")
                _RWidthLCL = dr("RWidthLCL")
                _RWidthUCL = dr("RWidthUCL")
                _RWidthURL = dr("RWidthURL")
                _RWidthGraphMax = dr("RWidthGraphMax")
                _RWidthObjective = IIf(IsDBNull(dr("RWidthObjective")), 0, dr("RWidthObjective"))

                _BPWidthGraphMin = dr("BPWidthGraphMin")
                _BPWidthLRL = dr("BPWidthLRL")
                _BPWidthLCL = dr("BPWidthLCL")
                _BPWidthUCL = dr("BPWidthUCL")
                _BPWidthURL = dr("BPWidthURL")
                _BPWidthGraphMax = dr("BPWidthGraphMax")
                _BPWidthObjective = IIf(IsDBNull(dr("BPWidthObjective")), 0, dr("BPWidthObjective"))

                _PcntcGraphMin = dr("PcntcGraphMin")
                _PcntcLRL = dr("PcntcLRL")
                _PcntcLCL = dr("PcntcLCL")
                _PcntcUCL = dr("PcntcUCL")
                _PcntcURL = dr("PcntcURL")
                _PcntcGraphMax = dr("PcntcGraphMax")
                _PcntcObjective = IIf(IsDBNull(dr("PcntcObjective")), 0, dr("PcntcObjective"))

                _WDGraphMin = dr("WDGraphMin")
                _WDLRL = dr("WDLRL")
                _WDLCL = dr("WDLCL")
                _WDUCL = dr("WDUCL")
                _WDURL = dr("WDURL")
                _WDGraphMax = dr("WDGraphMax")
                _WDObjective = IIf(IsDBNull(dr("WDObjective")), 0, dr("WDObjective"))
            End If
        End Using
    End Sub

    Public Sub Update()
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Assign all the necessary properties to an SQL Parameter
        Call AssignParameters(oSQL)

        ' Run the stored procedure
        oSQL.RunProc("UpdateQCParameters")
    End Sub

    Public Sub LoadParameters(ByVal Grade As Integer, ByRef QCSpecifications() As QCSpecificationsStructure, ByRef BasisAdjustment As Decimal)
        Read(Grade)

        QCSpecifications(FieldTypeEnum.Weight).GraphMin = WeightGraphMin
        QCSpecifications(FieldTypeEnum.Weight).LRL = WeightLRL
        QCSpecifications(FieldTypeEnum.Weight).LCL = WeightLCL
        QCSpecifications(FieldTypeEnum.Weight).UCL = WeightUCL + BasisAdjustment
        QCSpecifications(FieldTypeEnum.Weight).URL = WeightURL
        QCSpecifications(FieldTypeEnum.Weight).GraphMax = WeightGraphMax
        QCSpecifications(FieldTypeEnum.Weight).Objective = WeightObjective

        QCSpecifications(FieldTypeEnum.OpSideBulk).GraphMin = OpSideBulkGraphMin
        QCSpecifications(FieldTypeEnum.OpSideBulk).LRL = OpSideBulkLRL
        QCSpecifications(FieldTypeEnum.OpSideBulk).LCL = OpSideBulkLCL
        QCSpecifications(FieldTypeEnum.OpSideBulk).UCL = OpSideBulkUCL
        QCSpecifications(FieldTypeEnum.OpSideBulk).URL = OpSideBulkURL
        QCSpecifications(FieldTypeEnum.OpSideBulk).GraphMax = OpSideBulkGraphMax
        QCSpecifications(FieldTypeEnum.OpSideBulk).Objective = OpSideBulkObjective

        QCSpecifications(FieldTypeEnum.DriveSideBulk).GraphMin = DriveSideBulkGraphMin
        QCSpecifications(FieldTypeEnum.DriveSideBulk).LRL = DriveSideBulkLRL
        QCSpecifications(FieldTypeEnum.DriveSideBulk).LCL = DriveSideBulkLCL
        QCSpecifications(FieldTypeEnum.DriveSideBulk).UCL = DriveSideBulkUCL
        QCSpecifications(FieldTypeEnum.DriveSideBulk).URL = DriveSideBulkURL
        QCSpecifications(FieldTypeEnum.DriveSideBulk).GraphMax = DriveSideBulkGraphMax
        QCSpecifications(FieldTypeEnum.DriveSideBulk).Objective = DriveSideBulkObjective

        QCSpecifications(FieldTypeEnum.MDT).GraphMin = MDTGraphMin
        QCSpecifications(FieldTypeEnum.MDT).LRL = MDTLRL
        QCSpecifications(FieldTypeEnum.MDT).LCL = MDTLCL
        QCSpecifications(FieldTypeEnum.MDT).UCL = MDTUCL
        QCSpecifications(FieldTypeEnum.MDT).URL = MDTURL
        QCSpecifications(FieldTypeEnum.MDT).GraphMax = MDTGraphMax
        QCSpecifications(FieldTypeEnum.MDT).Objective = MDTObjective

        QCSpecifications(FieldTypeEnum.MDTE).GraphMin = MDTEGraphMin
        QCSpecifications(FieldTypeEnum.MDTE).LRL = MDTELRL
        QCSpecifications(FieldTypeEnum.MDTE).LCL = MDTELCL
        QCSpecifications(FieldTypeEnum.MDTE).UCL = MDTEUCL
        QCSpecifications(FieldTypeEnum.MDTE).URL = MDTEURL
        QCSpecifications(FieldTypeEnum.MDTE).GraphMax = MDTEGraphMax
        QCSpecifications(FieldTypeEnum.MDTE).Objective = MDTEObjective

        QCSpecifications(FieldTypeEnum.CDTD).GraphMin = CDTDGraphMin
        QCSpecifications(FieldTypeEnum.CDTD).LRL = CDTDLRL
        QCSpecifications(FieldTypeEnum.CDTD).LCL = CDTDLCL
        QCSpecifications(FieldTypeEnum.CDTD).UCL = CDTDUCL
        QCSpecifications(FieldTypeEnum.CDTD).URL = CDTDURL
        QCSpecifications(FieldTypeEnum.CDTD).GraphMax = CDTDGraphMax
        QCSpecifications(FieldTypeEnum.CDTD).Objective = CDTDObjective

        QCSpecifications(FieldTypeEnum.CDTW).GraphMin = CDTWGraphMin
        QCSpecifications(FieldTypeEnum.CDTW).LRL = CDTWLRL
        QCSpecifications(FieldTypeEnum.CDTW).LCL = CDTWLCL
        QCSpecifications(FieldTypeEnum.CDTW).UCL = CDTWUCL
        QCSpecifications(FieldTypeEnum.CDTW).URL = CDTWURL
        QCSpecifications(FieldTypeEnum.CDTW).GraphMax = CDTWGraphMax
        QCSpecifications(FieldTypeEnum.CDTW).Objective = CDTWObjective

        QCSpecifications(FieldTypeEnum.CDTC).GraphMin = CDTCGraphMin
        QCSpecifications(FieldTypeEnum.CDTC).LRL = CDTCLRL
        QCSpecifications(FieldTypeEnum.CDTC).LCL = CDTCLCL
        QCSpecifications(FieldTypeEnum.CDTC).UCL = CDTCUCL
        QCSpecifications(FieldTypeEnum.CDTC).URL = CDTCURL
        QCSpecifications(FieldTypeEnum.CDTC).GraphMax = CDTCGraphMax
        QCSpecifications(FieldTypeEnum.CDTC).Objective = CDTCObjective

        QCSpecifications(FieldTypeEnum.TT).GraphMin = TTGraphMin
        QCSpecifications(FieldTypeEnum.TT).LRL = TTLRL
        QCSpecifications(FieldTypeEnum.TT).LCL = TTLCL
        QCSpecifications(FieldTypeEnum.TT).UCL = TTUCL
        QCSpecifications(FieldTypeEnum.TT).URL = TTURL
        QCSpecifications(FieldTypeEnum.TT).GraphMax = TTGraphMax
        QCSpecifications(FieldTypeEnum.TT).Objective = TTObjective

        QCSpecifications(FieldTypeEnum.ZPeel).GraphMin = ZPeelGraphMin
        QCSpecifications(FieldTypeEnum.ZPeel).LRL = ZPeelLRL
        QCSpecifications(FieldTypeEnum.ZPeel).LCL = ZPeelLCL
        QCSpecifications(FieldTypeEnum.ZPeel).UCL = ZPeelUCL
        QCSpecifications(FieldTypeEnum.ZPeel).URL = ZPeelURL
        QCSpecifications(FieldTypeEnum.ZPeel).GraphMax = ZPeelGraphMax
        QCSpecifications(FieldTypeEnum.ZPeel).Objective = ZPeelObjective

        QCSpecifications(FieldTypeEnum.Contamination).GraphMin = ContaminationGraphMin
        QCSpecifications(FieldTypeEnum.Contamination).LRL = ContaminationLRL
        QCSpecifications(FieldTypeEnum.Contamination).LCL = ContaminationLCL
        QCSpecifications(FieldTypeEnum.Contamination).UCL = ContaminationUCL
        QCSpecifications(FieldTypeEnum.Contamination).URL = ContaminationURL
        QCSpecifications(FieldTypeEnum.Contamination).GraphMax = ContaminationGraphMax
        QCSpecifications(FieldTypeEnum.Contamination).Objective = ContaminationObjective

        QCSpecifications(FieldTypeEnum.TWA).GraphMin = TWAGraphMin
        QCSpecifications(FieldTypeEnum.TWA).LRL = TWALRL
        QCSpecifications(FieldTypeEnum.TWA).LCL = TWALCL
        QCSpecifications(FieldTypeEnum.TWA).UCL = TWAUCL
        QCSpecifications(FieldTypeEnum.TWA).URL = TWAURL
        QCSpecifications(FieldTypeEnum.TWA).GraphMax = TWAGraphMax
        QCSpecifications(FieldTypeEnum.TWA).Objective = TWAObjective

        QCSpecifications(FieldTypeEnum.AbsorbRate).GraphMin = AbsorbRateGraphMin
        QCSpecifications(FieldTypeEnum.AbsorbRate).LRL = AbsorbRateLRL
        QCSpecifications(FieldTypeEnum.AbsorbRate).LCL = AbsorbRateLCL
        QCSpecifications(FieldTypeEnum.AbsorbRate).UCL = AbsorbRateUCL
        QCSpecifications(FieldTypeEnum.AbsorbRate).URL = AbsorbRateURL
        QCSpecifications(FieldTypeEnum.AbsorbRate).GraphMax = AbsorbRateGraphMax
        QCSpecifications(FieldTypeEnum.AbsorbRate).Objective = AbsorbRateObjective

        QCSpecifications(FieldTypeEnum.RWidth).GraphMin = RWidthGraphMin
        QCSpecifications(FieldTypeEnum.RWidth).LRL = RWidthLRL
        QCSpecifications(FieldTypeEnum.RWidth).LCL = RWidthLCL
        QCSpecifications(FieldTypeEnum.RWidth).UCL = RWidthUCL
        QCSpecifications(FieldTypeEnum.RWidth).URL = RWidthURL
        QCSpecifications(FieldTypeEnum.RWidth).GraphMax = RWidthGraphMax
        QCSpecifications(FieldTypeEnum.RWidth).Objective = RWidthObjective

        QCSpecifications(FieldTypeEnum.BPWidth).GraphMin = BPWidthGraphMin
        QCSpecifications(FieldTypeEnum.BPWidth).LRL = BPWidthLRL
        QCSpecifications(FieldTypeEnum.BPWidth).LCL = BPWidthLCL
        QCSpecifications(FieldTypeEnum.BPWidth).UCL = BPWidthUCL
        QCSpecifications(FieldTypeEnum.BPWidth).URL = BPWidthURL
        QCSpecifications(FieldTypeEnum.BPWidth).GraphMax = BPWidthGraphMax
        QCSpecifications(FieldTypeEnum.BPWidth).Objective = BPWidthObjective

        QCSpecifications(FieldTypeEnum.Pcntc).GraphMin = PcntcGraphMin
        QCSpecifications(FieldTypeEnum.Pcntc).LRL = PcntcLRL
        QCSpecifications(FieldTypeEnum.Pcntc).LCL = PcntcLCL
        QCSpecifications(FieldTypeEnum.Pcntc).UCL = PcntcUCL
        QCSpecifications(FieldTypeEnum.Pcntc).URL = PcntcURL
        QCSpecifications(FieldTypeEnum.Pcntc).GraphMax = PcntcGraphMax
        QCSpecifications(FieldTypeEnum.Pcntc).Objective = PcntcObjective

        QCSpecifications(FieldTypeEnum.WD).GraphMin = WDGraphMin
        QCSpecifications(FieldTypeEnum.WD).LRL = WDLRL
        QCSpecifications(FieldTypeEnum.WD).LCL = WDLCL
        QCSpecifications(FieldTypeEnum.WD).UCL = WDUCL
        QCSpecifications(FieldTypeEnum.WD).URL = WDURL
        QCSpecifications(FieldTypeEnum.WD).GraphMax = WDGraphMax
        QCSpecifications(FieldTypeEnum.WD).Objective = WDObjective
    End Sub
End Class

Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class PartClass
    Inherits ClassBase

    Private _Acttyp As String = ""
    Private _BOMUOM As String = ""
    Private _SLSUOM As String = ""
    Private _Committed As Integer = 0
    Private _Cost As Decimal = 0
    Private _Description As String = ""
    Private _Description2 As String = ""
    Private _OnHand As Integer = 0
    Private _QuantityCommitted As Integer = 0
    Private _PartNumber As String = ""
    Private _PoundsPerSalesUnit As Decimal = 0
    Private _PLANID As String = ""
    Private _REVLEV As String = ""
    Private _StockCode As String = ""
    Private _STAENG As String = ""
    Private _OrderPolicy As String = ""
    Private _SLSCNV As Decimal = 0

    Public ReadOnly Property ACTTYP() As String
        Get
            Return _Acttyp
        End Get
    End Property

    Public ReadOnly Property Available() As Integer
        Get
            Return OnHand - Committed
        End Get
    End Property

    Public ReadOnly Property QuantityAvailable() As Integer
        Get
            Return OnHand - QuantityCommitted
        End Get
    End Property

    Public ReadOnly Property BOMUOM() As String
        Get
            Return _BOMUOM
        End Get
    End Property

    Public ReadOnly Property SLSUOM() As String
        Get
            Return _SLSUOM
        End Get
    End Property

    Public ReadOnly Property Committed() As Integer
        Get
            Return _Committed
        End Get
    End Property

    Public ReadOnly Property QuantityCommitted() As Integer
        Get
            Return _QuantityCommitted
        End Get
    End Property

    Public ReadOnly Property Cost() As Decimal
        Get
            Return _Cost
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return _Description.Trim
        End Get
    End Property

    Public ReadOnly Property Description2() As String
        Get
            Return _Description2.Trim
        End Get
    End Property

    Public ReadOnly Property OnHand() As Integer
        Get
            Return _OnHand
        End Get
    End Property

    Public ReadOnly Property PartNumber() As String
        Get
            Return _PartNumber
        End Get
    End Property

    Public ReadOnly Property PlanCode() As String
        Get
            Return _PLANID
        End Get
    End Property

    Public ReadOnly Property PoundsPerSalesUnit() As Decimal
        Get
            Return _PoundsPerSalesUnit
        End Get
    End Property

    Public ReadOnly Property PrivateLabel() As Boolean
        Get
            Return _Description.IndexOf("*") <> -1
        End Get
    End Property

    Public ReadOnly Property REVLEV() As String
        Get
            Return _REVLEV
        End Get
    End Property

    Public ReadOnly Property STAENG() As String
        Get
            Return _STAENG
        End Get
    End Property

    Public ReadOnly Property StockCode() As String
        Get
            Return _StockCode
        End Get
    End Property

    Public ReadOnly Property MTS() As Boolean   'Made To Stock
        Get
            If _OrderPolicy = "P" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property MTO() As Boolean   'Made To Order
        Get
            If _OrderPolicy = "O" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property SLSCNV() As Decimal
        Get
            Return _SLSCNV
        End Get
    End Property

    Public Enum OrderTypes
        FinishedGoods = 1
        RollGoods = 2
    End Enum

    ' This constructor will only be used if someone wishes to access
    ' the generic READ method to get a list of all parts
    Public Sub New()

    End Sub

    Public Sub New(ByVal passPart As String, ByVal source As DataSource)
        _PartNumber = passPart
        Read(passPart, source)
    End Sub

    Public Sub New(ByVal passPart As String)
        _PartNumber = passPart
        ReadSellarsPart(passPart)
    End Sub

    '*********************************************************************
    ' LastProduced(LocationID, LineID)
    ' Returns the part number that was last produced on the specified line
    '*********************************************************************
    Public Function LastProduced(ByVal LocationID As Integer, ByVal LineID As Integer) As String
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@LocationID", SqlDbType.SmallInt, 0, LocationID, ParameterDirection.Input)
        oSQL.AddParameter("@LineID", SqlDbType.SmallInt, 0, LineID, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim Part As New Parameter("@Part", SqlDbType.NVarChar, 15, CStr(""), ParameterDirection.Output)
        oSQL.AddParameter(Part)

        ' Run the stored procedure
        oSQL.RunProc("PartLastProduced")

        ' Return the stored procedure output parameter values
        Return Part.Value
    End Function

    Public Function Read() As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure
        Return oSQL.RunProcReader("ReadAllPartSales")
    End Function

    Public Function ReadActive() As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure
        Return oSQL.RunProcReader("ReadAllPartDescriptions")
    End Function

    Public Function Read(ByVal PlanCode As String, ByVal REVLEV As String) As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PlanCode", SqlDbType.NVarChar, 3, PlanCode.Trim, ParameterDirection.Input)
        oSQL.AddParameter("@REVLEV", SqlDbType.NVarChar, 3, REVLEV.Trim, ParameterDirection.Input)

        ' Run the stored procedure
        Return oSQL.RunProcReader("ReadAllPartsByPlanCodes")
    End Function

    Private Sub Read(ByVal passpart As String, ByVal source As DataSource)
        If source = DataSource.Max Then
            ReadMaxPart(passpart)
        Else
            ReadSellarsPart(passpart)
        End If
    End Sub

    Private Sub ReadMaxPart(ByVal passpart As String)
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        'Dim strSQL_Original As String = "Select ACTTYP_01, BOMUOM_01, PMDES1_01, PMDES2_01, REVLEV_01, PLANID_01, COST_01, ONHAND_01, WGT_01, DELSTK_01, STAENG_01 " & _
        '                       "From ""Part Master"" " & _
        '                       "Where PRTNUM_01 = '" & passpart.Trim & "'"

        Dim strSQL As String = "Select ACTTYP_01, BOMUOM_01, PMDES1_01, PMDES2_01, REVLEV_01, PLANID_01, COST_01, ONHAND_01, WGT_01, DELSTK_01, STAENG_01" & _
                               ", ORDPOL_01, PMDES1_29, PMDES2_29, QTYCOM_29, SLSUOM_29, SLSCNV_29 " & _
                               " From ""Part_Master"", ""PART_SALES""" & _
                               " Where PRTNUM_01 = '" & passpart.Trim & "'" & _
                               " AND PRTNUM_01 = PRTNUM_29"

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim PartMasterReader As SqlDataReader = cmd.ExecuteReader()
        If PartMasterReader.Read() Then

            If IsDBNull(PartMasterReader("ACTTYP_01")) Then
                _Acttyp = ""
            Else
                _Acttyp = PartMasterReader("ACTTYP_01")
            End If

            If IsDBNull(PartMasterReader("BOMUOM_01")) Then
                _BOMUOM = ""
            Else
                _BOMUOM = PartMasterReader("BOMUOM_01")
            End If

            If IsDBNull(PartMasterReader("SLSUOM_29")) Then
                _SLSUOM = ""
            Else
                _SLSUOM = PartMasterReader("SLSUOM_29")
            End If

            If IsDBNull(PartMasterReader("PMDES1_29")) Then
                _Description = ""
            Else
                '_Description = PartMasterReader("PMDES1_01")
                _Description = PartMasterReader("PMDES1_29")
            End If

            If IsDBNull(PartMasterReader("PMDES2_29")) Then
                _Description2 = ""
            Else
                '_Description2 = PartMasterReader("PMDES2_01")
                _Description2 = PartMasterReader("PMDES2_29")
            End If

            If IsDBNull(PartMasterReader("REVLEV_01")) Then
                _REVLEV = ""
            Else
                _REVLEV = PartMasterReader("REVLEV_01")
            End If

            If IsDBNull(PartMasterReader("PLANID_01")) Then
                _PLANID = ""
            Else
                _PLANID = PartMasterReader("PLANID_01")
            End If

            If IsDBNull(PartMasterReader("COST_01")) Then
                _Cost = 0
            Else
                _Cost = PartMasterReader("COST_01")
            End If

            If IsDBNull(PartMasterReader("ONHAND_01")) Then
                _OnHand = 0
            Else
                _OnHand = PartMasterReader("ONHAND_01")
            End If

            If IsDBNull(PartMasterReader("WGT_01")) Then
                _PoundsPerSalesUnit = 0
            Else
                _PoundsPerSalesUnit = PartMasterReader("WGT_01")
            End If

            If IsDBNull(PartMasterReader("DELSTK_01")) Then
                _StockCode = ""
            Else
                _StockCode = PartMasterReader("DELSTK_01")
            End If

            If IsDBNull(PartMasterReader("STAENG_01")) Then
                _STAENG = ""
            Else
                _STAENG = PartMasterReader("STAENG_01")
            End If

            If IsDBNull(PartMasterReader("ORDPOL_01")) Then
                _OrderPolicy = ""
            Else
                _OrderPolicy = PartMasterReader("ORDPOL_01")
            End If

            If IsDBNull(PartMasterReader("QTYCOM_29")) Then
                _QuantityCommitted = 0
            Else
                _QuantityCommitted = PartMasterReader("QTYCOM_29")
            End If

            If IsDBNull(PartMasterReader("SLSCNV_29")) Then
                _SLSCNV = 0
            Else
                _SLSCNV = PartMasterReader("SLSCNV_29")
            End If
        Else
            _Acttyp = ""
            _BOMUOM = ""
            _Description = ""
            _Description2 = ""
            _REVLEV = ""
            _PLANID = ""
            _Cost = 0
            _OnHand = 0
            _PoundsPerSalesUnit = 0
            _StockCode = ""
            _STAENG = ""
            _OrderPolicy = ""
        End If

        PartMasterReader.Close()
        PartMasterReader = Nothing
        CloseMaxConnection()
    End Sub

    Private Sub ReadSellarsPart(ByVal passpart As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, passpart, ParameterDirection.Input)

        ' Get the datareader that contains the results
        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadPartRecord")

        ' Assign the variables from the database to properties
        If dr.Read() Then
            _BOMUOM = dr("BOMUOM_01")
            _Description = dr("PMDES1_01")
            _Description2 = dr("PMDES2_01")
            _REVLEV = dr("REVLEV_01")
            _PLANID = dr("PLANID_01")
            _Cost = dr("COST_01")
            _OnHand = dr("ONHAND_01")
            _Committed = dr("QTYCOM_29")
            _PoundsPerSalesUnit = dr("WGT_01")
        Else
            Throw New RecordNotOnDatabaseException(passpart, "PartMaster SQL")
        End If

        dr.Close()
        dr = Nothing
    End Sub

    Public Function Valid(ByVal PlanId As String, ByVal OrderType As OrderTypes) As Boolean
        Dim rtnVal As Boolean = False
        Try
            Using Conn As New SqlConnection(ConnectionString)
                Conn.Open()

                Dim strSQL As String = "select @Count = count(PlanId) from OrderTypePlanIds where PlanId = @PlanId and OrderType = @OrderType"

                Using Cmd As New SqlCommand(strSQL, Conn)

                    Cmd.CommandType = CommandType.Text
                    Cmd.Parameters.Add(New SqlParameter("@PlanId", PlanId))
                    Cmd.Parameters.Add(New SqlParameter("@OrderType", OrderType))

                    ' Add the parameter to return if there is a matching plan id for the specified application type
                    Dim prmCount As New SqlParameter("@Count", SqlDbType.SmallInt, 0)
                    prmCount.Direction = ParameterDirection.Output
                    Cmd.Parameters.Add(prmCount)

                    ' Execute the stored procedure to update the InvoiceMasterExt table
                    Cmd.ExecuteNonQuery()

                    ' Set the return flag to true
                    If Convert.ToInt16(prmCount.Value) > 0 Then
                        rtnVal = True
                    End If

                End Using
            End Using
        Catch ex As Exception
            Throw New ApplicationException("Valid function failed: " + ex.Message)
        End Try

        Return rtnVal
    End Function

End Class

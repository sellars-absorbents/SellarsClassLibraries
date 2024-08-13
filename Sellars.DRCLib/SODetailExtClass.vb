Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math
Imports System.Text

Public Class SODetailExtClass
    Inherits ClassBase

    Private _BasePaperWidth As Decimal = 0
    Private _BuyerPartNumber As String = ""
    Private _DiscountPercent As Decimal = 0
    Private _EstProduction As Date = DefaultDate
    Private _CoreSize As Decimal = 0
    Private _OutsideDiameter As Decimal = 0
    Private _Pallets As Boolean = False
    Private _SlitWidth As Decimal = 0
    Private _Priority As Decimal = 99
    Private _Produced As Boolean = False
    Private _PromotionCode As String = ""
    Private _UnisourceStatus As Integer = 0
    Private _UnitPrice As Decimal = 0
    Private _SentToUnisource As Date = DefaultDate
    Private _ShippedFromUnisource As Date = DefaultDate
    Private _PromotionAddedLine As Boolean = False
    Private _ShipFromWarehouse As String = ""
    Private _runDate As Date

    Public Enum Pallet
        No = 0
        Yes = 1
    End Enum

    Public ReadOnly Property BasePaperWidth() As String
        Get
            Return _BasePaperWidth
        End Get
    End Property

    Public ReadOnly Property BuyerPartNumber() As String
        Get
            Return _BuyerPartNumber
        End Get
    End Property

    Public ReadOnly Property CoreSize() As Decimal
        Get
            Return _CoreSize
        End Get
    End Property

    Public ReadOnly Property DiscountPercent() As Decimal
        Get
            Return _DiscountPercent
        End Get
    End Property

    Public ReadOnly Property EstProduction() As Date
        Get
            Return _EstProduction
        End Get
    End Property

    Public ReadOnly Property OutsideDiameter() As Decimal
        Get
            Return _OutsideDiameter
        End Get
    End Property

    Public ReadOnly Property Pallets() As Boolean
        Get
            Return _Pallets
        End Get
    End Property

    Public ReadOnly Property Priority() As Decimal
        Get
            Return _Priority
        End Get
    End Property

    Public ReadOnly Property Produced() As Boolean
        Get
            Return _Produced
        End Get
    End Property

    Public ReadOnly Property PromotionAddedLine() As Boolean
        Get
            Return _PromotionAddedLine
        End Get
    End Property

    Public ReadOnly Property PromotionCode() As String
        Get
            Return _PromotionCode
        End Get
    End Property

    Public ReadOnly Property SentToUnisource() As Date
        Get
            Return _SentToUnisource
        End Get
    End Property

    Public ReadOnly Property ShippedFromUnisource() As Date
        Get
            Return _ShippedFromUnisource
        End Get
    End Property

    Public ReadOnly Property ShipFromWarehouse() As String
        Get
            Return _ShipFromWarehouse
        End Get
    End Property

    Public ReadOnly Property SlitWidth() As Decimal
        Get
            Return _SlitWidth
        End Get
    End Property

    Public ReadOnly Property UnisourceStatus() As Integer
        Get
            Return _UnisourceStatus
        End Get
    End Property

    Public ReadOnly Property UnitPrice() As Decimal
        Get
            Return _UnitPrice
        End Get
    End Property

    Public ReadOnly Property RunDate() As Date
        Get
            Return _runDate
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal Ordnum As String, ByVal LINNUM As String, ByVal DELNUM As String)
        Read(Ordnum, LINNUM, DELNUM)
    End Sub

    Public Sub Add(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passUnisourceStatus As Integer, ByVal BuyerPartNumber As String, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal ShipFromWarehouse As String, ByVal WarehouseChangeReason As Short, ByVal runDate As Date, Optional ByVal creditMemo As Boolean = False)
        Add(ORDNUM, LINNUM, DELNUM, passCoreSize, passOutsideDiameter, passWidth, passPallets, passUnisourceStatus, BuyerPartNumber, passPromotionCode, passUnitPrice, passDiscountPercent, ShipFromWarehouse, WarehouseChangeReason, creditMemo)
        UpdateRunDate(ORDNUM, LINNUM, DELNUM, runDate)
    End Sub

    Public Sub Add(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passUnisourceStatus As Integer, ByVal BuyerPartNumber As String, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal ShipFromWarehouse As String, ByVal WarehouseChangeReason As Short, Optional ByVal creditMemo As Boolean = False)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Add():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@CoreSize", SqlDbType.Float, 0, passCoreSize, ParameterDirection.Input)
        oSQL.AddParameter("@OD", SqlDbType.Float, 0, passOutsideDiameter, ParameterDirection.Input)
        oSQL.AddParameter("@Width", SqlDbType.Float, 0, passWidth, ParameterDirection.Input)
        oSQL.AddParameter("@Pallets", SqlDbType.Bit, 0, passPallets, ParameterDirection.Input)
        oSQL.AddParameter("@UnisourceStatus", SqlDbType.Int, 0, passUnisourceStatus, ParameterDirection.Input)
        oSQL.AddParameter("@BuyerPartNumber", SqlDbType.NVarChar, 50, BuyerPartNumber, ParameterDirection.Input)
        oSQL.AddParameter("@PromotionCode", SqlDbType.NVarChar, 30, passPromotionCode, ParameterDirection.Input)
        oSQL.AddParameter("@UnitPrice", SqlDbType.Decimal, 0, passUnitPrice, ParameterDirection.Input)
        oSQL.AddParameter("@DiscountPercent", SqlDbType.Decimal, 0, passDiscountPercent, ParameterDirection.Input)

        'If creditMemo Then
        oSQL.AddParameter("@Produced", SqlDbType.Bit, 1, 0, ParameterDirection.Input)
        oSQL.AddParameter("@Scheduled", SqlDbType.Bit, 1, 0, ParameterDirection.Input)
        'End If
        oSQL.AddParameter("@ShipFromWarehouse", SqlDbType.NVarChar, 10, ShipFromWarehouse, ParameterDirection.Input)
        oSQL.AddParameter("@WarehouseChangeReason", SqlDbType.SmallInt, 0, WarehouseChangeReason, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSODetailExt")
    End Sub

    Public Sub AddFree(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passUnisourceStatus As Integer, ByVal BuyerPartNumber As String, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal ShipFromWarehouse As String, ByVal WarehouseChangeReason As Short, Optional ByVal creditMemo As Boolean = False)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Add():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@CoreSize", SqlDbType.Float, 0, passCoreSize, ParameterDirection.Input)
        oSQL.AddParameter("@OD", SqlDbType.Float, 0, passOutsideDiameter, ParameterDirection.Input)
        oSQL.AddParameter("@Width", SqlDbType.Float, 0, passWidth, ParameterDirection.Input)
        oSQL.AddParameter("@Pallets", SqlDbType.Bit, 0, passPallets, ParameterDirection.Input)
        oSQL.AddParameter("@UnisourceStatus", SqlDbType.Int, 0, passUnisourceStatus, ParameterDirection.Input)
        oSQL.AddParameter("@BuyerPartNumber", SqlDbType.NVarChar, 50, BuyerPartNumber, ParameterDirection.Input)
        oSQL.AddParameter("@PromotionCode", SqlDbType.NVarChar, 30, passPromotionCode, ParameterDirection.Input)
        oSQL.AddParameter("@UnitPrice", SqlDbType.Decimal, 0, passUnitPrice, ParameterDirection.Input)
        oSQL.AddParameter("@DiscountPercent", SqlDbType.Decimal, 0, passDiscountPercent, ParameterDirection.Input)
        oSQL.AddParameter("@ShipFromWarehouse", SqlDbType.NVarChar, 10, ShipFromWarehouse, ParameterDirection.Input)
        oSQL.AddParameter("@WarehouseChangeReason", SqlDbType.SmallInt, 0, WarehouseChangeReason, ParameterDirection.Input)

        If creditMemo Then
            oSQL.AddParameter("@Produced", SqlDbType.Bit, 1, passPallets, ParameterDirection.Input)
            oSQL.AddParameter("@Scheduled", SqlDbType.Bit, 0, passPallets, ParameterDirection.Input)
        End If

        ' Run the stored procedure
        oSQL.RunProc("AddSODetailExtFree")
    End Sub

    Public Sub Clone(ByVal OldORDNUM As String, ByVal NewORDNUM As String, ByVal SelectedPart As String, ByVal scheduledDate As DateTime)
        Dim strSQL As String = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE (TABLE_NAME = 'SalesOrderDetailExt')"
        Dim columnListTarget As New StringBuilder()
        Dim connectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString

        Using Conn As SqlConnection = New SqlConnection(connectionString)
            If Conn.State.Equals(ConnectionState.Closed) Then
                Conn.Open()
            End If

            Dim SQL As New StringBuilder()

            Using Cmd As SqlCommand = New SqlCommand(strSQL, Conn)
                Cmd.CommandType = CommandType.Text

                Using GetColumns As SqlDataReader = Cmd.ExecuteReader()
                    While GetColumns.Read
                        Dim colName As String = GetColumns("COLUMN_NAME")

                        Select Case colName
                            Case "ORDNUM"
                                columnListTarget.Append("'" + NewORDNUM + "'")

                            Case "Priority"
                                columnListTarget.Append(", 99")

                            Case "Produced",
                            "Scheduled"
                                columnListTarget.Append(", 0")

                            Case "EstProduction"
                                columnListTarget.Append(", '2000-01-01 12:01:00.00 AM'")

                            Case "RunDate"
                                columnListTarget.Append(", '" + scheduledDate.ToString("yyyy-MM-dd") + "'")

                            Case Else
                                columnListTarget.Append("," + colName)
                        End Select
                    End While
                End Using
            End Using

            SQL.Append("INSERT INTO SalesOrderDetailExt SELECT " + columnListTarget.ToString() + " FROM SalesOrderDetailExt WHERE ORDNUM = @ORDNUM")

            If SelectedPart <> "" Then
                SQL.Append(" and LINNUM in (select LINNUM_28 from ExactMAXSELLR..SO_Detail where ORDNUM_28 = @ORDNUM and PRTNUM_28 like @PRTNUM)")
            End If

            Using Cmd As SqlCommand = New SqlCommand(SQL.ToString(), Conn)
                Cmd.CommandType = CommandType.Text
                Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OldORDNUM))

                If SelectedPart <> "" Then
                    Cmd.Parameters.Add(New SqlParameter("@PRTNUM", SelectedPart + "%"))
                End If

                Cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function CloneRecord(ByVal ORDNUM As String, ByVal FromLINNUM As String, ByVal ToLINNUM As String, ByVal DELNUM As String, ByVal PromotionCode As String) As Boolean
        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the submitted flag to true, the submitted date to the current date and time, and the reference field to be the Sellars Distribution
        ' automatically generated invoice number for applicable orders
        Dim strSQL As String = "insert into SalesOrderDetailExt (ORDNUM, LINNUM, DELNUM, SlitWidth, CoreSize, OD, Pallets, BasePaperWidth, Priority, Produced, Scheduled, EstProduction, UnisourceStatus, SentToUnisource, ShippedFromUnisource, BuyerPartNumber, PromotionCode, UnitPrice, DiscountPercent, PromotionAddedLine, ShipFromWarehouse) select ORDNUM, @ToLINNUM, DELNUM, SlitWidth, CoreSize, OD, Pallets, BasePaperWidth, Priority, Produced, Scheduled, EstProduction, UnisourceStatus, SentToUnisource, ShippedFromUnisource, BuyerPartNumber, @PromotionCode, UnitPrice, UnitPrice, 1, ShipFromWarehouse from SalesOrderDetailExt where ORDNUM = @ORDNUM and LINNUM = @FromLINNUM and DELNUM = @DELNUM"

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add a parameter to the command that has the invoice number to update
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@FromLINNUM", FromLINNUM))
                cmd.Parameters.Add(New SqlParameter("@ToLINNUM", ToLINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@PromotionCode", PromotionCode))

                cmd.ExecuteNonQuery()

                ' Set the return flag to true
                rtnData = True
            End Using
        End Using

        Return rtnData
    End Function

    '*********************************************************************
    ' Delete()
    ' Deletes the customer note for a customer.
    '*********************************************************************
    Public Sub Delete(ByVal pOrder As String, ByVal pLINNUM As String, ByVal pDELNUM As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, pOrder, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, pLINNUM, ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, pDELNUM, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("DeleteSalesOrderDetailExt")
    End Sub

    Public Sub Read(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)

        ' Run the stored procedure
        Using dr As SqlDataReader = oSQL.RunProcReader("GetSODetailExt")
            ' if no record was read, then clear all the fields
            If dr Is Nothing Then
                ClearFields()
                Throw New Exception("SQL Database Not Found while reading SODetailExt")
                Exit Sub
            End If

            ' Assign the variables from the database to properties
            If dr.Read() Then
                If IsDBNull(dr("BuyerPartNumber")) Then
                    _BuyerPartNumber = 0
                Else
                    _BuyerPartNumber = dr("BuyerPartNumber")
                End If
                If IsDBNull(dr("CoreSize")) Then
                    _CoreSize = 0
                Else
                    _CoreSize = dr("CoreSize")
                End If
                If IsDBNull(dr("DiscountPercent")) Then
                    _DiscountPercent = 0
                Else
                    _DiscountPercent = dr("DiscountPercent")
                End If
                If IsDBNull(dr("OD")) Then
                    _OutsideDiameter = 0
                Else
                    _OutsideDiameter = dr("OD")
                End If
                If IsDBNull(dr("Pallets")) Then
                    _Pallets = False
                Else
                    _Pallets = dr("Pallets")
                End If
                If IsDBNull(dr("Priority")) Then
                    _Priority = 99
                Else
                    _Priority = dr("Priority")
                End If
                If IsDBNull(dr("EstProduction")) Then
                    _EstProduction = DefaultDate
                Else
                    _EstProduction = dr("EstProduction")
                End If
                If IsDBNull(dr("Produced")) Then
                    _Produced = False
                Else
                    _Produced = dr("Produced")
                End If
                If IsDBNull(dr("PromotionAddedLine")) Then
                    _PromotionAddedLine = False
                Else
                    _PromotionAddedLine = dr("PromotionAddedLine")
                End If
                If IsDBNull(dr("PromotionCode")) Then
                    _PromotionCode = ""
                Else
                    _PromotionCode = dr("PromotionCode")
                End If
                If IsDBNull(dr("SlitWidth")) Then
                    _SlitWidth = 0
                Else
                    _SlitWidth = dr("SlitWidth")
                End If
                If IsDBNull(dr("BasePaperWidth")) Then
                    _BasePaperWidth = 0
                Else
                    _BasePaperWidth = dr("BasePaperWidth")
                End If
                If IsDBNull(dr("UnisourceStatus")) Then
                    _UnisourceStatus = 0
                Else
                    _UnisourceStatus = dr("UnisourceStatus")
                End If
                If IsDBNull(dr("UnitPrice")) Then
                    _UnitPrice = 0
                Else
                    _UnitPrice = dr("UnitPrice")
                End If
                If IsDBNull(dr("SentToUnisource")) Then
                    _SentToUnisource = DefaultDate
                Else
                    _SentToUnisource = dr("SentToUnisource")
                End If
                If IsDBNull(dr("ShippedFromUnisource")) Then
                    _ShippedFromUnisource = DefaultDate
                Else
                    _ShippedFromUnisource = dr("ShippedFromUnisource")
                End If
                If IsDBNull(dr("ShipFromWarehouse")) Then
                    _ShipFromWarehouse = ""
                Else
                    _ShipFromWarehouse = dr("ShipFromWarehouse")
                End If
            Else
                ClearFields()
            End If
        End Using

        Dim sql As String = $"select IsNull(RunDate, sod.CUSDUE_28) as RunDate
                              from SalesOrderDetailExt sodx
                              join exactmaxsellr.dbo.SO_Detail sod on sod.ORDNUM_28 = sodx.ORDNUM and SOD.LINNUM_28 = sodx.LINNUM and SOD.DELNUM_28 = sodx.DELNUM
                              where sodx.ORDNUM = @OrderNumber and sodx.LINNUM = @LineNumber and sodx.DELNUM = @DelNumber"
        Dim runDate As Date = Date.Now
        Using connection As SqlConnection = New SqlConnection(ConnectionString)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@OrderNumber", ORDNUM))
                command.Parameters.Add(New SqlParameter("@LineNumber", LINNUM))
                command.Parameters.Add(New SqlParameter("@DelNumber", DELNUM))

                Dim runDateResult As Object = command.ExecuteScalar()

                If runDateResult IsNot Nothing AndAlso runDateResult IsNot DBNull.Value AndAlso Date.TryParse(runDateResult, runDate) Then
                    _runDate = runDate
                End If
            End Using
        End Using
    End Sub

    Private Sub ClearFields()
        _BuyerPartNumber = ""
        _CoreSize = 0
        _DiscountPercent = 0
        _EstProduction = DefaultDate
        _OutsideDiameter = 0
        _Pallets = False
        _PromotionAddedLine = False
        _PromotionCode = ""
        _SlitWidth = 0
        _UnisourceStatus = 0
        _UnitPrice = 0
        _SentToUnisource = DefaultDate
        _ShipFromWarehouse = ""
    End Sub

    Public Sub UpdateRunDate(orderNumber As String, lineNumber As String, delNumber As String, runDate As Date)
        Dim sql As String = "update SalesOrderDetailExt set RunDate = @RunDate where ORDNUM = @OrderNumber and LINNUM = @LineNumber and DELNUM = @DelNumber"

        Using connection As SqlConnection = New SqlConnection(ConnectionString)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@RunDate", runDate))
                command.Parameters.Add(New SqlParameter("@OrderNumber", orderNumber))
                command.Parameters.Add(New SqlParameter("@LineNumber", lineNumber))
                command.Parameters.Add(New SqlParameter("@DelNumber", delNumber))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@CoreSize", SqlDbType.Float, 0, passCoreSize, ParameterDirection.Input)
        oSQL.AddParameter("@OD", SqlDbType.Float, 0, passOutsideDiameter, ParameterDirection.Input)
        oSQL.AddParameter("@Width", SqlDbType.Float, 0, passWidth, ParameterDirection.Input)
        oSQL.AddParameter("@Pallets", SqlDbType.Bit, 0, passPallets, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSODetailExt")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal PassWarehouse As String, ByVal runDate As Date)
        Update(ORDNUM, LINNUM, DELNUM, passCoreSize, passOutsideDiameter, passWidth, passPallets, passPromotionCode, passUnitPrice, passDiscountPercent, PassWarehouse)
        UpdateRunDate(ORDNUM, LINNUM, DELNUM, runDate)
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal PassWarehouse As String)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@CoreSize", SqlDbType.Float, 0, passCoreSize, ParameterDirection.Input)
        oSQL.AddParameter("@OD", SqlDbType.Float, 0, passOutsideDiameter, ParameterDirection.Input)
        oSQL.AddParameter("@Width", SqlDbType.Float, 0, passWidth, ParameterDirection.Input)
        oSQL.AddParameter("@Pallets", SqlDbType.Bit, 0, passPallets, ParameterDirection.Input)
        oSQL.AddParameter("@PromotionCode", SqlDbType.NVarChar, 30, passPromotionCode, ParameterDirection.Input)
        oSQL.AddParameter("@UnitPrice", SqlDbType.Decimal, 0, passUnitPrice, ParameterDirection.Input)
        oSQL.AddParameter("@DiscountPercent", SqlDbType.Decimal, 0, passDiscountPercent, ParameterDirection.Input)
        oSQL.AddParameter("@Warehouse", SqlDbType.NVarChar, 10, PassWarehouse, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSODetailExtWUnitPrice")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal PassWarehouse As String, ByVal WarehouseChangeReason As Short, ByVal runDate As Date)
        Update(ORDNUM, LINNUM, DELNUM, passCoreSize, passOutsideDiameter, passWidth, passPallets, passPromotionCode, passUnitPrice, passDiscountPercent, PassWarehouse, WarehouseChangeReason)
        UpdateRunDate(ORDNUM, LINNUM, DELNUM, runDate)
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passCoreSize As Decimal, ByVal passOutsideDiameter As Decimal, ByVal passWidth As Decimal, ByVal passPallets As Boolean, ByVal passPromotionCode As String, ByVal passUnitPrice As Decimal, ByVal passDiscountPercent As Decimal, ByVal PassWarehouse As String, ByVal WarehouseChangeReason As Short)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@CoreSize", SqlDbType.Float, 0, passCoreSize, ParameterDirection.Input)
        oSQL.AddParameter("@OD", SqlDbType.Float, 0, passOutsideDiameter, ParameterDirection.Input)
        oSQL.AddParameter("@Width", SqlDbType.Float, 0, passWidth, ParameterDirection.Input)
        oSQL.AddParameter("@Pallets", SqlDbType.Bit, 0, passPallets, ParameterDirection.Input)
        oSQL.AddParameter("@PromotionCode", SqlDbType.NVarChar, 30, passPromotionCode, ParameterDirection.Input)
        oSQL.AddParameter("@UnitPrice", SqlDbType.Decimal, 0, passUnitPrice, ParameterDirection.Input)
        oSQL.AddParameter("@DiscountPercent", SqlDbType.Decimal, 0, passDiscountPercent, ParameterDirection.Input)
        oSQL.AddParameter("@Warehouse", SqlDbType.NVarChar, 10, PassWarehouse, ParameterDirection.Input)
        oSQL.AddParameter("@WarehouseChangeReason", SqlDbType.SmallInt, 0, WarehouseChangeReason, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSODetailExtWUnitPriceReason")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passPriority As Decimal)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@Priority", SqlDbType.Float, 0, passPriority, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSODetailExtPriority")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal passProduced As Boolean)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@Produced", SqlDbType.Bit, 0, passProduced, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSODetailExtProduced")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal EstProduction As Date)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@LINNUM", SqlDbType.NVarChar, 2, LINNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@DELNUM", SqlDbType.NVarChar, 2, DELNUM.PadLeft(2, "0"), ParameterDirection.Input)
        oSQL.AddParameter("@EstProduction", SqlDbType.SmallDateTime, 0, EstProduction, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSODetailExtEstProduction")
    End Sub

    Public Function Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal UnisourceStatus As Integer) As Boolean
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the UnisourceStatus for the applicable lines in the order
        Dim strSQL As String = "update SalesOrderDetailExt set UnisourceStatus = @UnisourceStatus where ORDNUM = @ORDNUM and LINNUM = @LINNUM and DELNUM = @DELNUM"

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            ' Set up a new SQL Command
            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add a parameter to the command that has the invoice number to update
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@UnisourceStatus", UnisourceStatus))

                cmd.ExecuteNonQuery()

                ' Set the return flag to true
                rtnData = True
            End Using
        End Using

        Return rtnData
    End Function

    Public Function Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal UnisourceStatus As Integer, ByVal ShipfromWarehouse As String, ByVal WarehouseChangeReason As Short) As Boolean
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the UnisourceStatus for the applicable lines in the order
        Dim strSQL As String = "update SalesOrderDetailExt set UnisourceStatus = @UnisourceStatus, ShipFromWarehouse = @ShipFromWarehouse, WarehouseChangeReasonCode = @WarehouseChangeReason where ORDNUM = @ORDNUM and LINNUM = @LINNUM and DELNUM = @DELNUM"

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            ' Set up a new SQL Command
            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)

                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add parameters to the command to fill in the sql command string defined above
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@UnisourceStatus", UnisourceStatus))
                cmd.Parameters.Add(New SqlParameter("@ShipFromWarehouse", ShipfromWarehouse))
                cmd.Parameters.Add(New SqlParameter("@WarehouseChangeReason", WarehouseChangeReason))

                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' Set the return flag to true
        rtnData = True

        Return rtnData
    End Function

    Public Function Update(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal ShipFromWarehouse As String, ByVal WarehouseChangeReason As Short) As Boolean
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the UnisourceStatus for the applicable lines in the order
        Dim strSQL As String = ""

        If ShipFromWarehouse = "DSC1" Or ShipFromWarehouse = "MIL2" Then
            strSQL = "update SalesOrderDetailExt set UnisourceStatus = 0, SentToUnisource = '12/31/2050 00:00:00 AM', AcknowledgedOn = '12/31/2050 00:00:00 AM', ShipFromWarehouse = @ShipFromWarehouse where ORDNUM = @ORDNUM And LINNUM = @LINNUM And DELNUM = @DELNUM"
        Else
            strSQL = "update SalesOrderDetailExt set UnisourceStatus = case when UnisourceStatus = 0 then 1 else UnisourceStatus end, ShipFromWarehouse = @ShipFromWarehouse where ORDNUM = @ORDNUM And LINNUM = @LINNUM And DELNUM = @DELNUM"
        End If

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add a parameter to the command that has the invoice number to update
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@ShipFromWarehouse", ShipFromWarehouse))
                cmd.Parameters.Add(New SqlParameter("@WarehouseChangeReason", WarehouseChangeReason))

                cmd.ExecuteNonQuery()

                ' Set the return flag to true
                rtnData = True
            End Using
        End Using

        Return rtnData
    End Function

    Public Function UpdateSentToUnisource(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal SentToUnisource As DateTime) As Boolean
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the submitted flag to true, the submitted date to the current date and time, and the reference field to be the Sellars Distribution
        ' automatically generated invoice number for applicable orders
        Dim strSQL As String = "update SalesOrderDetailExt set UnisourceStatus = 3 SentToUnisource = @SentToUnisource where ORDNUM = @ORDNUM and LINNUM = @LINNUM and DELNUM = @DELNUM"

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add a parameter to the command that has the invoice number to update
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@SentToUnisource", SentToUnisource))

                cmd.ExecuteNonQuery()

                ' Set the return flag to true
                rtnData = True
            End Using
        End Using

        Return rtnData
    End Function

    Public Function UpdateUnisourceStatus(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal UnisourceStatus As Integer) As Boolean
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:Update():Order number is empty.")
        End If

        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the submitted flag to true, the submitted date to the current date and time, and the reference field to be the Sellars Distribution
        ' automatically generated invoice number for applicable orders
        Dim strSQL As String = "update SalesOrderDetailExt set UnisourceStatus = @UnisourceStatus where ORDNUM = @ORDNUM and LINNUM = @LINNUM and DELNUM = @DELNUM"

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add a parameter to the command that has the invoice number to update
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@UnisourceStatus", UnisourceStatus))

                cmd.ExecuteNonQuery()

                ' Set the return flag to true
                rtnData = True
            End Using
        End Using

        Return rtnData
    End Function

    Public Function UpdatePromotionCodeWPrice(ByVal ORDNUM As String, ByVal LINNUM As String, ByVal DELNUM As String, ByVal PromotionCode As String, ByVal UnitPrice As Decimal, ByVal DiscountPercent As Decimal, ByVal QuoteIssue As Boolean) As Boolean
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SODetailExtClass:UpdatePromotionCode():Order number is empty.")
        End If

        ' Declare a variable which will be returned as a result of this procedure
        Dim rtnData As Boolean = False

        ' Set the submitted flag to true, the submitted date to the current date and time, and the reference field to be the Sellars Distribution
        ' automatically generated invoice number for applicable orders
        Dim strSQL As String = "update SalesOrderDetailExt set PromotionCode = @PromotionCode, UnitPrice = @UnitPrice, DiscountPercent = @DiscountPercent, QuoteIssue = @QuoteIssue where ORDNUM = @ORDNUM and LINNUM = @LINNUM and DELNUM = @DELNUM"

        Using conn As SqlConnection = New SqlConnection(ConnectionString)
            conn.Open()

            Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                ' Add a parameter to the command that has the invoice number to update
                cmd.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                cmd.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                cmd.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                cmd.Parameters.Add(New SqlParameter("@PromotionCode", PromotionCode))
                cmd.Parameters.Add(New SqlParameter("@UnitPrice", UnitPrice))
                cmd.Parameters.Add(New SqlParameter("@DiscountPercent", DiscountPercent))
                cmd.Parameters.Add(New SqlParameter("@QuoteIssue", QuoteIssue))

                cmd.ExecuteNonQuery()

                ' Set the return flag to true
                rtnData = True
            End Using
        End Using

        Return rtnData
    End Function
End Class

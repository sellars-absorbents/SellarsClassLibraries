Imports System.Collections.Generic
Imports System.Data.SqlClient

Public Class QueuedWebSalesOrderDetail
    Protected Friend Const DefaultDate As Date = #1/1/2000 12:01:00 AM#
    Private _LINNUM As String = ""                      ' Line Number
    Private _PRTNUM As String = ""                    ' Part Number
    Private _DUEDATE As Date = DefaultDate            ' Current Due Date
    Private _SLSUOM As String = ""                    ' Unit of Measure (e.g. 'EA', 'RL', etc.)
    Private _PRICE As Decimal = 0                      ' Unit Price
    Private _QUANTITYPURCHASED As Double = 0          ' Quantity the customer Purchased
    Private _QUANTITYFREE As Double = 0               ' Quantity that is free
    Private _DISC As Single = 0
    Private _BuyerPartNumber As String = ""
    Private _PromotionCode As String = ""
    Private _PromotionAddedLine As Boolean = False
    Private _FullPrice As Decimal = 0
    Private _DiscountPercent As Decimal = 0
    Private _GLXREF As String = ""
    Private _Notes As New SalesOrderDetailNotes
    Private _purchasedProcessed As Boolean = False
    Private _freeProcessed As Boolean = False
    Private _id As Guid
    Private _masterID As Guid

    Public Property ORDNUM() As String = ""

    Public Property BuyerPartNumber() As String
        Get
            Return _BuyerPartNumber
        End Get
        Set(ByVal value As String)
            _BuyerPartNumber = value
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

    Public Property GLXREF() As String
        Get
            Return _GLXREF
        End Get
        Set(ByVal value As String)
            _GLXREF = value
        End Set
    End Property

    Public Property LINNUM() As String
        Get
            Return _LINNUM
        End Get
        Set(ByVal value As String)
            _LINNUM = GetValue(value, "")
        End Set
    End Property

    Public Property DELNUM() As String = ""

    Public Property PRTNUM() As String
        Get
            Return _PRTNUM.Trim
        End Get
        Set(ByVal value As String)
            _PRTNUM = GetValue(value, "")
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

    Public Property DUEDATE() As Date
        Get
            Return _DUEDATE
        End Get
        Set(ByVal value As Date)
            _DUEDATE = GetValue(value, DefaultDate)
        End Set
    End Property

    Public Property FullPrice() As Decimal
        Get
            Return _FullPrice
        End Get
        Set(ByVal value As Decimal)
            _FullPrice = GetValue(value, 0)
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

    Public Property PRICE() As Decimal
        Get
            Return _PRICE
        End Get
        Set(ByVal value As Decimal)
            _PRICE = GetValue(value, 0)
        End Set
    End Property

    Public Property PromotionAddedLine() As Boolean
        Get
            Return _PromotionAddedLine
        End Get
        Set(ByVal value As Boolean)
            _PromotionAddedLine = value
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

    Public Property QuantityFree() As Double
        Get
            Return _QUANTITYFREE
        End Get
        Set(ByVal value As Double)
            _QUANTITYFREE = GetValue(value, 0)
        End Set
    End Property

    Public Property FreeProcessed As Boolean
        Get
            Return _freeProcessed
        End Get
        Private Set(value As Boolean)
            _freeProcessed = value
        End Set
    End Property

    Public Property QuantityPurchased() As Double
        Get
            Return _QUANTITYPURCHASED
        End Get
        Set(ByVal value As Double)
            _QUANTITYPURCHASED = GetValue(value, 0)
        End Set
    End Property

    Public Property PurchasedProcessed As Boolean
        Get
            Return _purchasedProcessed
        End Get
        Private Set(value As Boolean)
            _purchasedProcessed = value
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

    Public Property ID As Guid
        Get
            Return _id
        End Get
        Private Set(value As Guid)
            _id = value
        End Set
    End Property

    Public Property MasterID As Guid
        Get
            Return _masterID
        End Get
        Private Set(value As Guid)
            _masterID = value
        End Set
    End Property

    Private _onHold As Boolean = False
    Public Property OnHold As Boolean
        Get
            Return _onHold
        End Get
        Private Set(value As Boolean)
            _onHold = value
        End Set
    End Property

    Private _processTries As Integer = 0
    Public Property ProcessTries As Integer
        Get
            Return _processTries
        End Get
        Private Set(value As Integer)
            _processTries = value
        End Set
    End Property

    Private _processed As Boolean = False
    Public Property Processed As Boolean
        Get
            Return _processed
        End Get
        Private Set(value As Boolean)
            _processed = value
        End Set
    End Property

    Private _processedOn As Date = DefaultDate
    Public Property ProcessedOn As Date
        Get
            Return _processedOn
        End Get
        Private Set(value As Date)
            _processedOn = value
        End Set
    End Property

    Private _createdOn As Date = DefaultDate
    Public Property CreatedOn As Date
        Get
            Return _createdOn
        End Get
        Private Set(value As Date)
            _createdOn = value
        End Set
    End Property

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

    Public Function AddRecordToQueue(ByVal shopfloorConnection As String, ByVal masterID As Guid) As Guid
        Dim sql As String = "INSERT INTO [dbo].[QueuedWebSalesOrderDetail]([MasterID],[ORDNUM],[LINNUM],[DELNUM],[PRTNUM],[DUEDATE],[SLSUOM],[PRICE],[QuantityPurchased],[QuantityFree],[Disc],[BuyerPartNumber],[PromotionCode],[PromotionAddedLine],[FullPrice],[DiscountPercent],[GLXREF])
                             OUTPUT INSERTED.ID
                             VALUES
                                   (@MasterID,
                                    @ORDNUM,
                                    @LINNUM,
                                    @DELNUM,
                                   @PRTNUM,
                                   @DUEDATE,
                                   @SLSUOM,
                                   @PRICE,
                                   @QuantityPurchased,
                                   @QuantityFree,
                                   @Disc,
                                   @BuyerPartNumber,
                                   @PromotionCode,
                                   @PromotionAddedLine,
                                   @FullPrice,
                                   @DiscountPercent,
                                   @GLXREF)"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@MasterID", masterID))
                command.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                command.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                command.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                command.Parameters.Add(New SqlParameter("@PRTNUM", PRTNUM))
                command.Parameters.Add(New SqlParameter("@DUEDATE", DUEDATE))
                command.Parameters.Add(New SqlParameter("@SLSUOM", SLSUOM))
                command.Parameters.Add(New SqlParameter("@PRICE", PRICE))
                command.Parameters.Add(New SqlParameter("@QuantityPurchased", QuantityPurchased))
                command.Parameters.Add(New SqlParameter("@QuantityFree", QuantityFree))
                command.Parameters.Add(New SqlParameter("@Disc", DISC))
                command.Parameters.Add(New SqlParameter("@BuyerPartNumber", BuyerPartNumber))
                command.Parameters.Add(New SqlParameter("@PromotionCode", PromotionCode))
                command.Parameters.Add(New SqlParameter("@PromotionAddedLine", PromotionAddedLine))
                command.Parameters.Add(New SqlParameter("@FullPrice", FullPrice))
                command.Parameters.Add(New SqlParameter("@DiscountPercent", DiscountPercent))
                command.Parameters.Add(New SqlParameter("@GLXREF", GLXREF))

                Me.ID = CType(command.ExecuteScalar(), Guid)
                Me.MasterID = masterID
            End Using
        End Using

        Return Me.ID
    End Function

    Friend Shared Function GetQueuedRecords(ByVal connection As SqlConnection, ByVal masterID As Guid, ByVal onHold As Boolean, ByVal processed As Boolean) As List(Of QueuedWebSalesOrderDetail)
        Dim results As New List(Of QueuedWebSalesOrderDetail)
        Dim sql As String = "select *
                               from QueuedWebSalesOrderDetail
                               where MasterID = @MasterID
                               and OnHold = @OnHold
                               and Processed = @Processed"

        Using command As SqlCommand = New SqlCommand(sql, connection)
            command.Parameters.Add(New SqlParameter("@MasterID", masterID))
            command.Parameters.Add(New SqlParameter("@OnHold", onHold))
            command.Parameters.Add(New SqlParameter("@Processed", processed))

            Using dr As SqlDataReader = command.ExecuteReader()
                While dr.Read()
                    Dim detail As QueuedWebSalesOrderDetail = New QueuedWebSalesOrderDetail()

                    PopulateProperties(dr, detail)

                    results.Add(detail)
                End While
            End Using
        End Using

        For Each detail As QueuedWebSalesOrderDetail In results
            detail.Notes = SalesOrderDetailNotes.GetQueuedRecords(connection, detail.ID, onHold, processed)
        Next

        Return results
    End Function

    Private Shared Sub PopulateProperties(dr As SqlDataReader, detail As QueuedWebSalesOrderDetail)
        detail.ID = dr("ID")
        detail.MasterID = dr("MasterID")
        detail.ORDNUM = dr("ORDNUM")
        detail.LINNUM = dr("LINNUM")
        detail.DELNUM = dr("DELNUM")
        detail.PRTNUM = dr("PRTNUM")
        detail.DUEDATE = dr("DUEDATE")
        detail.SLSUOM = dr("SLSUOM")
        detail.PRICE = dr("PRICE")
        detail.QuantityPurchased = dr("QuantityPurchased")
        detail.PurchasedProcessed = dr("PurchasedProcessed")
        detail.QuantityFree = dr("QuantityFree")
        detail.FreeProcessed = dr("FreeProcessed")
        detail.DISC = dr("DISC")
        detail.BuyerPartNumber = dr("BuyerPartNumber")
        detail.PromotionCode = dr("PromotionCode")
        detail.PromotionAddedLine = dr("PromotionAddedLine")
        detail.FullPrice = dr("FullPrice")
        detail.DiscountPercent = dr("DiscountPercent")
        detail.GLXREF = dr("GLXREF")
        detail.OnHold = dr("OnHold")
        detail.ProcessTries = dr("ProcessTries")
        detail.Processed = dr("Processed")

        If (dr("ProcessedOn") IsNot DBNull.Value) Then
            detail.ProcessedOn = dr("ProcessedOn")
        Else
            detail.ProcessedOn = DefaultDate
        End If

        detail.CreatedOn = dr("CreatedOn")
    End Sub

    Public Sub IncrementProcessTries(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderDetail 
                             set ProcessTries = ProcessTries + 1
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub SetOrderInfo(ByVal shopfloorConnection As String, ByVal orderNumber As String, ByVal lineNumber As String, ByVal delNumber As String)
        Dim sql As String = "update QueuedWebSalesOrderDetail 
                             set ORDNUM = @ORDNUM,
                             LINNUM = @LINNUM,
                             DELNUM = @DELNUM
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ORDNUM", orderNumber))
                command.Parameters.Add(New SqlParameter("@LINNUM", lineNumber.PadLeft(2, " "c)))
                command.Parameters.Add(New SqlParameter("@DELNUM", delNumber.PadLeft(2, " "c)))
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub MarkPurchasedAsProcessed(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderDetail 
                             set PurchasedProcessed = 1
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub MarkFreeAsProcessed(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderDetail 
                             set FreeProcessed = 1
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub MarkAsProcessed(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderDetail 
                             set Processed = 1,
                             ProcessedOn = GETDATE()
                             where ID = @ID"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@ID", Me.ID))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class

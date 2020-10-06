Imports System.Data.SqlClient

Public Class SalesOrderDetailNotes
    Protected Friend Const DefaultDate As Date = #1/1/2000 12:01:00 AM#
    Private _id As Guid
    Private _detailID As Guid
    Private _DoNotPrint As String = ""
    Private _PrintOnInvoiceOnly As String = ""
    Private _PrintOnOrderOnly As String = ""
    Private _PrintOnBoth As String = ""

    Public Property ORDNUM As String = ""
    Public Property LINNUM As String = ""
    Public Property DELNUM As String = ""

    Public Property DoNotPrint() As String
        Get
            Return _DoNotPrint.Trim
        End Get
        Set(ByVal value As String)
            _DoNotPrint = GetValue(value, "")
        End Set
    End Property

    Public Property PrintOnBoth() As String
        Get
            Return _PrintOnBoth.Trim
        End Get
        Set(ByVal value As String)
            _PrintOnBoth = GetValue(value, "")
        End Set
    End Property

    Public Property PrintOnInvoiceOnly() As String
        Get
            Return _PrintOnInvoiceOnly.Trim
        End Get
        Set(ByVal value As String)
            _PrintOnInvoiceOnly = GetValue(value, "")
        End Set
    End Property

    Public Property PrintOnOrderOnly() As String
        Get
            Return _PrintOnOrderOnly.Trim
        End Get
        Set(ByVal value As String)
            _PrintOnOrderOnly = GetValue(value, "")
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

    Public Property DetailID As Guid
        Get
            Return _detailID
        End Get
        Private Set(value As Guid)
            _detailID = value
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

    Public Function AddRecordToQueue(ByVal shopfloorConnection As String, ByVal detailID As Guid)
        Dim sql As String = "INSERT INTO [dbo].[QueuedWebSalesOrderDetailNote]([DetailID],[ORDNUM],[LINNUM],[DELNUM],[DoNotPrint],[PrintOnInvoiceOnly],[PrintOnOrderOnly],[PrintOnBoth])
                             OUTPUT INSERTED.ID
                             VALUES
                                (@DetailID,
                                @ORDNUM,
                                @LINNUM,
                                @DELNUM,
                                @DoNotPrint,
                                @PrintOnInvoiceOnly,
                                @PrintOnOrderOnly,
                                @PrintOnBoth)"

        Using connection As SqlConnection = New SqlConnection(shopfloorConnection)
            connection.Open()

            Using command As SqlCommand = New SqlCommand(sql, connection)
                command.Parameters.Add(New SqlParameter("@DetailID", detailID))
                command.Parameters.Add(New SqlParameter("@ORDNUM", ORDNUM))
                command.Parameters.Add(New SqlParameter("@LINNUM", LINNUM))
                command.Parameters.Add(New SqlParameter("@DELNUM", DELNUM))
                command.Parameters.Add(New SqlParameter("@DoNotPrint", Me.DoNotPrint))
                command.Parameters.Add(New SqlParameter("@PrintOnInvoiceOnly", Me.PrintOnInvoiceOnly))
                command.Parameters.Add(New SqlParameter("@PrintOnOrderOnly", Me.PrintOnOrderOnly))
                command.Parameters.Add(New SqlParameter("@PrintOnBoth", Me.PrintOnBoth))

                Me.ID = CType(command.ExecuteScalar(), Guid)
                Me.DetailID = detailID
            End Using
        End Using

        Return Me.ID
    End Function

    Friend Shared Function GetQueuedRecords(ByVal connection As SqlConnection, ByVal detailID As Guid, ByVal onHold As Boolean, ByVal processed As Boolean) As SalesOrderDetailNotes
        Dim result As SalesOrderDetailNotes
        Dim sql As String = "select *
                               from QueuedWebSalesOrderDetailNote
                               where DetailID = @DetailID
                               and OnHold = @OnHold
                               and Processed = @Processed"

        Using command As SqlCommand = New SqlCommand(sql, connection)
            command.Parameters.Add(New SqlParameter("@DetailID", detailID))
            command.Parameters.Add(New SqlParameter("@OnHold", onHold))
            command.Parameters.Add(New SqlParameter("@Processed", processed))

            Using dr As SqlDataReader = command.ExecuteReader()
                While dr.Read()
                    result = New SalesOrderDetailNotes
                    PopulateProperties(dr, result)
                End While
            End Using
        End Using

        Return result
    End Function

    Private Shared Sub PopulateProperties(dr As SqlDataReader, nte As SalesOrderDetailNotes)
        nte.ID = dr("ID")
        nte.DetailID = dr("DetailID")
        nte.ORDNUM = dr("ORDNUM")
        nte.LINNUM = dr("LINNUM")
        nte.DELNUM = dr("DELNUM")
        nte.DoNotPrint = dr("DoNotPrint")
        nte.PrintOnInvoiceOnly = dr("PrintOnInvoiceOnly")
        nte.PrintOnOrderOnly = dr("PrintOnOrderOnly")
        nte.PrintOnBoth = dr("PrintOnBoth")
        nte.OnHold = dr("OnHold")
        nte.ProcessTries = dr("ProcessTries")
        nte.Processed = dr("Processed")

        If (dr("ProcessedOn") IsNot DBNull.Value) Then
            nte.ProcessedOn = dr("ProcessedOn")
        Else
            nte.ProcessedOn = DefaultDate
        End If

        nte.CreatedOn = dr("CreatedOn")
    End Sub

    Public Sub IncrementProcessTries(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderDetailNote 
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
        Dim sql As String = "update QueuedWebSalesOrderDetailNote 
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

    Public Sub MarkAsProcessed(ByVal shopfloorConnection As String)
        Dim sql As String = "update QueuedWebSalesOrderDetailNote 
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

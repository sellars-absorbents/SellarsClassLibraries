Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math
Imports System.Text

Public Class SOMasterExtClass
    Inherits ClassBase

    Private _AddressLocationNumber As String = ""
    Private _AllowFinish As Boolean = False
    Private _CarrierAlphaCode As String = ""
    Private _CarrierRouting As String = ""
    Private _CarrierTransMethodCode As String = ""
    Private _EnteredBy As String = ""
    Private _LastChanged As Date
    Private _LastPrinted As Date
    Private _LocationCodeQualifier As String = ""
    Private _Notes As String = ""
    Private _ProofedBy As String = ""
    Private _Hold As Boolean = False
    Private _HoldUserEmail As String = ""
    Private _DefaultStockID As String = ""
    Private _OrderType As String = ""
    Private _Finished As Boolean = False
    Private _ConversionID As Integer = 1
    Private _PickPrinted As Integer = 0
    Private _LastPickPrinted As DateTime = New DateTime(2050, 12, 31, 0, 0, 0)
    Private _CCTransactionID As String = ""
    Private _CCAuthCode As String = ""
    Private _CCCaptureStatus As Integer = 0
    Private _CCCaptureDate As Date = New DateTime(2050, 12, 31, 0, 0, 0)
    Private _ShipFromStockID As String = ""

    ' Set up the unique shipping fields
    Private _Carrier As Integer = 0
    Private _CarrierMethod As String = ""
    Private _CarrierThirdParty As String = ""
    Private _CarrierName As String = ""
    Private _ContactName As String = ""
    Private _ContactPhone As String = ""
    Private _ContactEmail As String = ""
    Private _EstimatedShipping As Decimal = 0

    Public ReadOnly Property AddressLocationNumber() As String
        Get
            Return _AddressLocationNumber
        End Get
    End Property

    Public ReadOnly Property AllowFinish() As Boolean
        Get
            Return _AllowFinish
        End Get
    End Property

    Public ReadOnly Property Carrier() As Integer
        Get
            Return _Carrier
        End Get
    End Property

    Public ReadOnly Property CarrierAlphaCode() As String
        Get
            Return _CarrierAlphaCode
        End Get
    End Property

    Public ReadOnly Property CarrierMethod() As String
        Get
            Return _CarrierMethod
        End Get
    End Property

    Public ReadOnly Property CarrierName() As String
        Get
            Return _CarrierName
        End Get
    End Property

    Public ReadOnly Property CarrierRouting() As String
        Get
            Return _CarrierRouting
        End Get
    End Property

    Public ReadOnly Property CarrierThirdParty() As String
        Get
            Return _CarrierThirdParty
        End Get
    End Property

    Public ReadOnly Property CarrierTransMethodCode() As String
        Get
            Return _CarrierTransMethodCode
        End Get
    End Property

    Public ReadOnly Property CCAuthCode() As String
        Get
            Return _CCAuthCode
        End Get
    End Property

    Public ReadOnly Property CCCaptureDate() As Date
        Get
            Return _CCCaptureDate
        End Get
    End Property

    Public ReadOnly Property CCCaptureStatus() As Integer
        Get
            Return _CCCaptureStatus
        End Get
    End Property

    Public ReadOnly Property CCTransactionID() As String
        Get
            Return _CCTransactionID
        End Get
    End Property

    Public ReadOnly Property ContactEmail() As String
        Get
            Return _ContactEmail
        End Get
    End Property

    Public ReadOnly Property ContactName() As String
        Get
            Return _ContactName
        End Get
    End Property

    Public ReadOnly Property ContactPhone() As String
        Get
            Return _ContactPhone
        End Get
    End Property

    Public ReadOnly Property ConversionID() As Integer
        Get
            Return _ConversionID
        End Get
    End Property

    Public ReadOnly Property DefaultStockID() As String
        Get
            Return _DefaultStockID
        End Get
    End Property

    Public ReadOnly Property EnteredBy() As String
        Get
            Return _EnteredBy
        End Get
    End Property

    Public ReadOnly Property Finished() As Boolean
        Get
            Return _Finished
        End Get
    End Property

    Public ReadOnly Property Hold() As Boolean
        Get
            Return _Hold
        End Get
    End Property

    Public ReadOnly Property HoldUserEmail() As String
        Get
            Return _HoldUserEmail
        End Get
    End Property

    Public ReadOnly Property LastChanged() As Date
        Get
            Return _LastChanged
        End Get
    End Property

    Public ReadOnly Property LastPickPrinted() As DateTime
        Get
            Return _LastPickPrinted
        End Get
    End Property

    Public ReadOnly Property LastPrinted() As Date
        Get
            Return _LastPrinted
        End Get
    End Property

    Public ReadOnly Property LocationCodeQualifier() As String
        Get
            Return _LocationCodeQualifier
        End Get
    End Property

    Public ReadOnly Property Notes() As String
        Get
            Return _Notes
        End Get
    End Property

    Public ReadOnly Property OrderType() As String
        Get
            Return _OrderType
        End Get
    End Property

    Public ReadOnly Property PickPrinted() As Integer
        Get
            Return _PickPrinted
        End Get
    End Property

    Public ReadOnly Property ProofedBy() As String
        Get
            Return _ProofedBy
        End Get
    End Property

    Public ReadOnly Property EstimatedShipping() As Decimal
        Get
            Return _EstimatedShipping
        End Get
    End Property

    Public ReadOnly Property ShipFromStockID() As String
        Get
            Return _ShipFromStockID
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal Ordnum As String)
        Read(Ordnum)
    End Sub

    Public Sub Add(ByVal ORDNUM As String, ByVal EnteredBy As String)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Add():Order number is empty.")
        End If
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@EnteredBy", SqlDbType.NVarChar, 15, EnteredBy, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSOMasterExt")
    End Sub

    Public Sub Add(ByVal ORDNUM As String, ByVal EnteredBy As String, ByVal CarrierID As Integer, ByVal CarrierMethod As String, ByVal CarrierThirdParty As String, ByVal CarrierName As String, ByVal ContactName As String, ByVal ContactPhone As String, ByVal ContactEmail As String, ByVal DefaultStockID As String, ByVal OrderType As String, ByVal ConversionID As Integer, ByVal ShipFromStockID As String, Optional ByVal SendAcknowledgement As Boolean = False)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Add():Order number is empty.")
        End If
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@EnteredBy", SqlDbType.NVarChar, 15, EnteredBy, ParameterDirection.Input)
        oSQL.AddParameter("@Carrier", SqlDbType.SmallInt, -1, CarrierID, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierMethod", SqlDbType.NVarChar, 50, CarrierMethod, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierThirdParty", SqlDbType.NVarChar, 50, CarrierThirdParty, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierName", SqlDbType.NVarChar, 50, CarrierName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactName", SqlDbType.NVarChar, 50, ContactName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactPhone", SqlDbType.NVarChar, 20, ContactPhone, ParameterDirection.Input)
        oSQL.AddParameter("@ContactEmail", SqlDbType.NVarChar, 100, ContactEmail, ParameterDirection.Input)
        oSQL.AddParameter("@DefaultStockID", SqlDbType.NVarChar, 10, DefaultStockID, ParameterDirection.Input)
        oSQL.AddParameter("@OrderType", SqlDbType.NVarChar, 2, OrderType, ParameterDirection.Input)
        oSQL.AddParameter("@ConversionID", SqlDbType.SmallInt, 0, ConversionID, ParameterDirection.Input)
        oSQL.AddParameter("@ShipFromStockID", SqlDbType.NVarChar, 10, ShipFromStockID, ParameterDirection.Input)
        oSQL.AddParameter("@SendAcknowledgement", SqlDbType.Bit, 0, SendAcknowledgement, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSOMasterExtWCarrier")
    End Sub

    Public Sub Add(ByVal ORDNUM As String, ByVal EnteredBy As String, ByVal CarrierID As Integer, ByVal CarrierMethod As String, ByVal CarrierThirdParty As String, ByVal CarrierName As String, ByVal ContactName As String, ByVal ContactPhone As String, ByVal ContactEmail As String, ByVal DefaultStockID As String, ByVal OrderType As String, ByVal ConversionID As Integer, ByVal CCTransactionID As String, ByVal CCAuthCode As String, ByVal EstimatedShipping As Decimal, ByVal Websitetotal As Decimal, ByVal ShipFromStockID As String, Optional ByVal SendAcknowledgement As Boolean = False)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Add():Order number is empty.")
        End If
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@EnteredBy", SqlDbType.NVarChar, 15, EnteredBy, ParameterDirection.Input)
        oSQL.AddParameter("@Carrier", SqlDbType.SmallInt, -1, CarrierID, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierMethod", SqlDbType.NVarChar, 50, CarrierMethod, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierThirdParty", SqlDbType.NVarChar, 50, CarrierThirdParty, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierName", SqlDbType.NVarChar, 50, CarrierName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactName", SqlDbType.NVarChar, 50, ContactName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactPhone", SqlDbType.NVarChar, 20, ContactPhone, ParameterDirection.Input)
        oSQL.AddParameter("@ContactEmail", SqlDbType.NVarChar, 100, ContactEmail, ParameterDirection.Input)
        oSQL.AddParameter("@DefaultStockID", SqlDbType.NVarChar, 10, DefaultStockID, ParameterDirection.Input)
        oSQL.AddParameter("@OrderType", SqlDbType.NVarChar, 2, OrderType, ParameterDirection.Input)
        oSQL.AddParameter("@ConversionID", SqlDbType.SmallInt, 0, ConversionID, ParameterDirection.Input)
        oSQL.AddParameter("@CCTransactionID", SqlDbType.NVarChar, 50, CCTransactionID, ParameterDirection.Input)
        oSQL.AddParameter("@CCAuthCode", SqlDbType.NVarChar, 20, CCAuthCode, ParameterDirection.Input)
        oSQL.AddParameter("@EstimatedShipping", SqlDbType.Decimal, 0, EstimatedShipping, ParameterDirection.Input)
        oSQL.AddParameter("@WebsiteTotal", SqlDbType.Decimal, 0, Websitetotal, ParameterDirection.Input)
        oSQL.AddParameter("@ShipFromStockID", SqlDbType.NVarChar, 10, ShipFromStockID, ParameterDirection.Input)
        oSQL.AddParameter("@SendAcknowledgement", SqlDbType.Bit, 0, SendAcknowledgement, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSOMasterExtWCarrierCC")
    End Sub

    Public Sub Add(ByVal ORDNUM As String, ByVal EnteredBy As String, ByVal LocationCodeQualifier As String, ByVal AddressLocationNumber As String)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Add():Order number is empty.")
        End If
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@EnteredBy", SqlDbType.NVarChar, 15, EnteredBy, ParameterDirection.Input)
        oSQL.AddParameter("@LocationCodeQualifier", SqlDbType.NVarChar, 20, LocationCodeQualifier, ParameterDirection.Input)
        oSQL.AddParameter("@AddressLocationNumber", SqlDbType.NVarChar, 10, AddressLocationNumber, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddSOMasterExtWLoc")
    End Sub

    Public Sub Clone(ByVal OldORDNUM As String, ByVal NewORDNUM As String, ByVal Username As String)
        Dim strSQL As String = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE (TABLE_NAME = 'SalesOrderMasterExt')"

        Dim columnListTarget As New StringBuilder()

        ' retrieve the connection string from the task config
        Dim connectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString

        Dim Cmd As SqlCommand

        ' Open up the database connection
        Dim Conn As SqlConnection = New SqlConnection(connectionString)
        If Conn.State.Equals(ConnectionState.Closed) Then
            Conn.Open()
        End If

        Try
            ' Comment added.
            Cmd = New SqlCommand(strSQL, Conn)
            Cmd.CommandType = CommandType.Text

            Dim GetColumns As SqlDataReader = Cmd.ExecuteReader()
            While GetColumns.Read

                Dim colName As String = GetColumns("COLUMN_NAME").ToString.Trim()
                Select Case colName
                    Case "ORDNUM"
                        columnListTarget.Append("'" + NewORDNUM + "'")

                    Case "EnteredBy"
                        columnListTarget.Append(", '" + Username + "'")

                    Case "LastChanged"
                        columnListTarget.Append(", '" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'")

                    Case "LocationCodeQualifier",
                        "AddressLocationNumber",
                        "RDCDescription",
                        "CarrierAlphaCode",
                        "CarrierRouting",
                        "CarrierTransMethodCode",
                        "CCTransactionID",
                        "CCAuthCode",
                        "ProofedBy"
                        columnListTarget.Append(", ''")

                    Case "AcknowledgementSent",
                        "CCCaptureDate",
                        "LastPrinted",
                        "LastPickPrinted"
                        columnListTarget.Append(", '12/31/2050 00:00:00 AM'")

                    Case "Finished",
                        "PickPrinted",
                        "SendAcknowledgement",
                        "CCCaptureStatus"
                        columnListTarget.Append(", 0")

                    Case Else
                        columnListTarget.Append(", " + colName)
                End Select

            End While

            GetColumns.Close()

            Dim SQL As String = "INSERT INTO SalesOrderMasterExt SELECT " + columnListTarget.ToString() + " FROM SalesOrderMasterExt WHERE ORDNUM = @ORDNUM"
            Cmd = New SqlCommand(SQL, Conn)
            Cmd.CommandType = CommandType.Text
            Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OldORDNUM))
            Cmd.ExecuteNonQuery()

        Catch ex As Exception
            Dim x As String = ex.Message
        End Try

    End Sub

    Public Sub Read(ByVal ORDNUM As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        Dim dr As SqlDataReader

        ' Run the stored procedure
        dr = oSQL.RunProcReader("GetSOMasterExt")

        ' Assign the variables from the database to properties
        If dr Is Nothing Then
            ClearFields()
            Throw New Exception("SQL Database Not Found while reading SOMasterExt")
            Exit Sub
        End If

        If dr.Read() Then
            If IsDBNull(dr("AddressLocationNumber")) Then
                _AddressLocationNumber = ""
            Else
                _AddressLocationNumber = dr("AddressLocationNumber")
            End If
            If IsDBNull(dr("AllowFinish")) Then
                _AllowFinish = False
            Else
                _AllowFinish = dr("AllowFinish")
            End If

            If IsDBNull(dr("Carrier")) Then
                _Carrier = 0
            Else
                _Carrier = Convert.ToInt16(dr("Carrier"))
            End If
            If IsDBNull(dr("CarrierAlphaCode")) Then
                _CarrierAlphaCode = ""
            Else
                _CarrierAlphaCode = dr("CarrierAlphaCode")
            End If
            If IsDBNull(dr("ContactEmail")) Then
                _ContactEmail = ""
            Else
                _ContactEmail = dr("ContactEmail")
            End If
            If IsDBNull(dr("ContactName")) Then
                _ContactName = ""
            Else
                _ContactName = dr("ContactName")
            End If
            If IsDBNull(dr("ContactPhone")) Then
                _ContactPhone = ""
            Else
                _ContactPhone = dr("ContactPhone")
            End If
            If IsDBNull(dr("CarrierMethod")) Then
                _CarrierMethod = 0
            Else
                _CarrierMethod = dr("CarrierMethod")
            End If
            If IsDBNull(dr("CarrierRouting")) Then
                _CarrierRouting = ""
            Else
                _CarrierRouting = dr("CarrierRouting")
            End If
            If IsDBNull(dr("CarrierThirdParty")) Then
                _CarrierThirdParty = ""
            Else
                _CarrierThirdParty = dr("CarrierThirdParty")
            End If
            If IsDBNull(dr("CarrierName")) Then
                _CarrierName = ""
            Else
                _CarrierName = dr("CarrierName")
            End If
            If IsDBNull(dr("CarrierTransMethodCode")) Then
                _CarrierTransMethodCode = ""
            Else
                _CarrierTransMethodCode = dr("CarrierTransMethodCode")
            End If
            If IsDBNull(dr("EnteredBy")) Then
                _EnteredBy = ""
            Else
                _EnteredBy = dr("EnteredBy")
            End If
            If IsDBNull(dr("LastChanged")) Then
                _LastChanged = DefaultDate
            Else
                _LastChanged = dr("LastChanged")
            End If
            If IsDBNull(dr("LastPrinted")) Then
                _LastPrinted = DefaultDate
            Else
                _LastPrinted = dr("LastPrinted")
            End If
            If IsDBNull(dr("LocationCodeQualifier")) Then
                _LocationCodeQualifier = ""
            Else
                _LocationCodeQualifier = dr("LocationCodeQualifier")
            End If
            If IsDBNull(dr("ProofedBy")) Then
                _ProofedBy = ""
            Else
                _ProofedBy = dr("ProofedBy")
            End If
            If IsDBNull(dr("Notes")) Then
                _Notes = ""
            Else
                _Notes = dr("Notes")
                _Notes.Replace("\r\n", Environment.NewLine)
            End If
            If IsDBNull(dr("Hold")) Then
                _Hold = False
            Else
                _Hold = dr("Hold")
            End If
            If IsDBNull(dr("HoldUserEmail")) Then
                _HoldUserEmail = ""
            Else
                _HoldUserEmail = dr("HoldUserEmail")
            End If
            If IsDBNull(dr("DefaultStockID")) Then
                _DefaultStockID = ""
            Else
                _DefaultStockID = dr("DefaultStockID")
            End If
            If IsDBNull(dr("OrderType")) Then
                _OrderType = ""
            Else
                _OrderType = dr("OrderType")
            End If
            If IsDBNull(dr("Finished")) Then
                _Finished = False
            Else
                _Finished = dr("Finished")
            End If
            If IsDBNull(dr("ConversionID")) Then
                _ConversionID = 1
            Else
                _ConversionID = dr("ConversionID")
            End If
            If IsDBNull(dr("PickPrinted")) Then
                _PickPrinted = 0
            Else
                _PickPrinted = dr("PickPrinted")
            End If
            If IsDBNull(dr("LastPickPrinted")) Then
                _LastPickPrinted = DefaultDate
            Else
                _LastPickPrinted = dr("LastPickPrinted")
            End If
            If IsDBNull(dr("CCAuthCode")) Then
                _CCAuthCode = ""
            Else
                _CCAuthCode = dr("CCAuthCode")
            End If
            If IsDBNull(dr("CCCaptureDate")) Then
                _CCCaptureDate = DefaultDate
            Else
                _CCCaptureDate = dr("CCCaptureDate")
            End If
            If IsDBNull(dr("CCCaptureStatus")) Then
                _CCCaptureStatus = 0
            Else
                _CCCaptureStatus = dr("CCCaptureStatus")
            End If
            If IsDBNull(dr("CCTransactionID")) Then
                _CCTransactionID = ""
            Else
                _CCTransactionID = dr("CCTransactionID")
            End If
            If IsDBNull(dr("EstimatedShipping")) Then
                _EstimatedShipping = 0
            Else
                _EstimatedShipping = dr("EstimatedShipping")
            End If
            If IsDBNull(dr("ShipFromStockID")) Then
                _ShipFromStockID = ""
            Else
                _ShipFromStockID = dr("ShipFromStockID")
            End If
        Else
            ClearFields()
        End If

        dr.Close()
        dr = Nothing
    End Sub

    Public Sub ReadHeld(ByRef DS As DataSet)
        ' Run the stored procedure
        Dim da As New SqlDataAdapter("Select ORDNUM from SalesOrderMasterExt where Hold = '1'", ConnectionString)
        da.Fill(DS, "HeldOrders")
    End Sub

    Private Sub ClearFields()
        _AddressLocationNumber = ""
        _AllowFinish = False
        _EnteredBy = ""
        _LastChanged = DefaultDate
        _LastPrinted = DefaultDate
        _LocationCodeQualifier = ""
        _ProofedBy = ""
        _Notes = ""
        _Hold = False
        _HoldUserEmail = ""
        _Carrier = 0
        _CarrierMethod = ""
        _CarrierThirdParty = ""
        _CCAuthCode = ""
        _CCCaptureDate = DefaultDate
        _CCCaptureStatus = 0
        _CCTransactionID = ""
        _ContactEmail = ""
        _ContactName = ""
        _ContactPhone = ""
        _DefaultStockID = ""
        _OrderType = ""
        _Finished = False
        _PickPrinted = 0
        _LastPickPrinted = DefaultDate
        _EstimatedShipping = 0
        _ShipFromStockID = ""
    End Sub

    Public Sub Update(ByVal ORDNUM As String)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtLastChanged")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal CarrierID As Integer, ByVal CarrierMethod As String, ByVal CarrierThirdParty As String, ByVal CarrierName As String, ByVal ContactName As String, ByVal ContactPhone As String, ByVal ContactEmail As String, ByVal ShipFromStockID As String, ByVal ConversionID As Integer)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Carrier", SqlDbType.SmallInt, -1, CarrierID, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierMethod", SqlDbType.NVarChar, 50, CarrierMethod, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierThirdParty", SqlDbType.NVarChar, 50, CarrierThirdParty, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierName", SqlDbType.NVarChar, 50, CarrierName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactName", SqlDbType.NVarChar, 50, ContactName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactPhone", SqlDbType.NVarChar, 20, ContactPhone, ParameterDirection.Input)
        oSQL.AddParameter("@ContactEmail", SqlDbType.NVarChar, 100, ContactEmail, ParameterDirection.Input)
        oSQL.AddParameter("@ShipFromStockID", SqlDbType.NVarChar, 10, ShipFromStockID, ParameterDirection.Input)
        oSQL.AddParameter("@ConversionID", SqlDbType.SmallInt, 0, ConversionID, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtWCarrier")
    End Sub

    Public Sub UpdateFromC(ByVal ORDNUM As String, ByVal CarrierID As Int32, ByVal CarrierMethod As String, ByVal CarrierThirdParty As String, ByVal CarrierName As String, ByVal ContactName As String, ByVal ContactPhone As String, ByVal ContactEmail As String, ByVal ShipFromStockID As String, ByVal ConversionID As Int32)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Carrier", SqlDbType.SmallInt, -1, CarrierID, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierMethod", SqlDbType.NVarChar, 50, CarrierMethod, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierThirdParty", SqlDbType.NVarChar, 50, CarrierThirdParty, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierName", SqlDbType.NVarChar, 50, CarrierName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactName", SqlDbType.NVarChar, 50, ContactName, ParameterDirection.Input)
        oSQL.AddParameter("@ContactPhone", SqlDbType.NVarChar, 20, ContactPhone, ParameterDirection.Input)
        oSQL.AddParameter("@ContactEmail", SqlDbType.NVarChar, 100, ContactEmail, ParameterDirection.Input)
        oSQL.AddParameter("@ShipFromStockID", SqlDbType.NVarChar, 10, ShipFromStockID, ParameterDirection.Input)
        oSQL.AddParameter("@ConversionID", SqlDbType.SmallInt, 0, ConversionID, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtWCarrier")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal ProofedBy As String)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@PROOFEDBY", SqlDbType.NVarChar, 15, ProofedBy, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtProofedBy")
    End Sub

    Public Sub Update(ByVal ORDNUM As String, ByVal Hold As Boolean, ByVal EMail As String)
        If String.IsNullOrEmpty(ORDNUM.Trim) Then
            Throw New ApplicationException("SalesOrderMasterExtClass:Update():Order number is empty.")
        End If

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Hold", SqlDbType.Bit, 0, Hold, ParameterDirection.Input)
        oSQL.AddParameter("@EMail", SqlDbType.NVarChar, 50, EMail, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtHold")
    End Sub

    Public Sub UpdateLastPrinted(ByVal StartORDNUM As String, ByVal EndORDNUM As String, ByVal Username As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@StartORDNUM", SqlDbType.NVarChar, 20, StartORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@EndORDNUM", SqlDbType.NVarChar, 20, EndORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@USERNAME", SqlDbType.NVarChar, 15, Username, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtLastPrinted2")
    End Sub

    Public Sub UpdateNotes(ByVal ORDNUM As String, ByVal Notes As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Notes", SqlDbType.NVarChar, 4000, Notes, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtNotes")
    End Sub

    Public Sub UpdateNotes(ByVal ORDNUM As String, ByVal Notes As String, ByVal LocationCodeQualifier As String, ByVal AddressLocationNumber As String, ByVal RDCDescription As String, ByVal CarrierAlphaCode As String, ByVal CarrierRouting As String, ByVal CarrierTransMethodCode As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, ORDNUM, ParameterDirection.Input)
        oSQL.AddParameter("@Notes", SqlDbType.NText, -1, Notes, ParameterDirection.Input)
        oSQL.AddParameter("@LocationCodeQualifier", SqlDbType.NVarChar, 10, LocationCodeQualifier, ParameterDirection.Input)
        oSQL.AddParameter("@AddressLocationNumber", SqlDbType.NVarChar, 10, AddressLocationNumber, ParameterDirection.Input)
        oSQL.AddParameter("@RDCDescription", SqlDbType.NVarChar, 100, RDCDescription, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierAlphaCode", SqlDbType.NVarChar, 100, CarrierAlphaCode, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierRouting", SqlDbType.NVarChar, 100, CarrierRouting, ParameterDirection.Input)
        oSQL.AddParameter("@CarrierTransMethodCode", SqlDbType.NVarChar, 100, CarrierTransMethodCode, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateSOMasterExtNotesWLoc")
    End Sub

    Public Sub UpdateEstimatedShipping(ByVal OrderNumber As String, ByVal EstimatedShipping As Decimal)
        Dim Cmd As SqlCommand

        ' Open up the database connection
        Dim Conn As SqlConnection = New SqlConnection(ConnectionString)
        If Conn.State.Equals(ConnectionState.Closed) Then
            Conn.Open()
        End If

        Dim strSQL As String = ""

        strSQL = "Update SalesOrderMasterExt set EstimatedShipping = @EstimatedShipping where ORDNUM = @ORDNUM"

        Try
            Cmd = New SqlCommand(strSQL, Conn)
            Cmd.CommandType = CommandType.Text
            Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OrderNumber))
            Cmd.Parameters.Add(New SqlParameter("@EstimatedShipping", EstimatedShipping))
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try

        ' Close the database connection and free up memory
        Conn.Close()
        Conn.Dispose()

        ' If the command is set, then dispose of it to free up memory
        If Cmd IsNot Nothing Then
            Cmd.Dispose()
        End If
    End Sub

    Public Sub UpdateFinished(ByVal OrderNumber As String, Optional ByVal Finished As Boolean = True, Optional ByVal PickPrinted As Boolean = False)
        Dim Cmd As SqlCommand

        ' Open up the database connection
        Dim Conn As SqlConnection = New SqlConnection(ConnectionString)
        If Conn.State.Equals(ConnectionState.Closed) Then
            Conn.Open()
        End If

        Dim strSQL As String = ""

        strSQL = "Update SalesOrderMasterExt set Finished = @Finished, PickPrinted = @PickPrinted where ORDNUM = @ORDNUM"

        Try
            Cmd = New SqlCommand(strSQL, Conn)
            Cmd.CommandType = CommandType.Text
            Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OrderNumber))
            Cmd.Parameters.Add(New SqlParameter("@Finished", Finished))
            Cmd.Parameters.Add(New SqlParameter("@PickPrinted", PickPrinted))
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try

        ' Close the database connection and free up memory
        Conn.Close()
        Conn.Dispose()

        ' If the command is set, then dispose of it to free up memory
        If Cmd IsNot Nothing Then
            Cmd.Dispose()
        End If
    End Sub

    Public Sub Reset(ByVal OrderNumber As String)
        Dim Cmd As SqlCommand

        ' Open up the database connection
        Dim Conn As SqlConnection = New SqlConnection(ConnectionString)
        If Conn.State.Equals(ConnectionState.Closed) Then
            Conn.Open()
        End If

        Dim strSQL As String = ""

        strSQL = "Update SalesOrderMasterExt set Finished = 0, PickPrinted = 0, LastPickPrinted = '12/31/2050 00:00:00 AM' where ORDNUM = @ORDNUM"

        Try
            Cmd = New SqlCommand(strSQL, Conn)
            Cmd.CommandType = CommandType.Text
            Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OrderNumber))
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try

        ' Close the database connection and free up memory
        Conn.Close()
        Conn.Dispose()

        ' If the command is set, then dispose of it to free up memory
        If Cmd IsNot Nothing Then
            Cmd.Dispose()
        End If
    End Sub

    Public Sub UpdateCCCaptureStatus(ByVal OrderNumber As String, ByVal Status As Integer)
        Dim Cmd As SqlCommand

        ' Open up the database connection
        Dim Conn As SqlConnection = New SqlConnection(ConnectionString)
        If Conn.State.Equals(ConnectionState.Closed) Then
            Conn.Open()
        End If

        Dim strSQL As String = ""

        strSQL = "Update SalesOrderMasterExt set CCCaptureStatus = @CCCaptureStatus where ORDNUM = @ORDNUM"

        Try
            Cmd = New SqlCommand(strSQL, Conn)
            Cmd.CommandType = CommandType.Text
            Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OrderNumber))
            Cmd.Parameters.Add(New SqlParameter("@CCCaptureStatus", Status))
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try

        ' Close the database connection and free up memory
        Conn.Close()
        Conn.Dispose()

        ' If the command is set, then dispose of it to free up memory
        If Cmd IsNot Nothing Then
            Cmd.Dispose()
        End If
    End Sub

    Public Sub UpdateCCPaymentInfo(ByVal OrderNumber As String, ByVal TransactionID As String, ByVal AuthCode As String)
        Dim Cmd As SqlCommand

        ' Open up the database connection
        Dim Conn As SqlConnection = New SqlConnection(ConnectionString)
        If Conn.State.Equals(ConnectionState.Closed) Then
            Conn.Open()
        End If

        Dim strSQL As String = ""

        strSQL = "Update SalesOrderMasterExt set CCCaptureStatus = 1, CCTransactionID = @TransactionID, CCAuthCode = @AuthCode, Finished = 1 where ORDNUM = @ORDNUM"

        Try
            Cmd = New SqlCommand(strSQL, Conn)
            Cmd.CommandType = CommandType.Text
            Cmd.Parameters.Add(New SqlParameter("@ORDNUM", OrderNumber))
            Cmd.Parameters.Add(New SqlParameter("@TransactionID", TransactionID))
            Cmd.Parameters.Add(New SqlParameter("@AuthCode", AuthCode))
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        End Try

        ' Close the database connection and free up memory
        Conn.Close()
        Conn.Dispose()

        ' If the command is set, then dispose of it to free up memory
        If Cmd IsNot Nothing Then
            Cmd.Dispose()
        End If
    End Sub

End Class
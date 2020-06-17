Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class CustomerSettingsClass
    Inherits ClassBase

    Private _ReadError As Boolean = False
    Private _GIP As Decimal = 0
    Private _Rebate As Decimal = 0
    Private _RebatePercent As Decimal = 0
    Private _Freight As Decimal = 0
    Private _Type As Integer = 1

    Public ReadOnly Property CustomerType() As Integer
        Get
            Return _Type
        End Get
    End Property

    Public ReadOnly Property Freight() As Decimal
        Get
            Return _Freight
        End Get
    End Property

    Public ReadOnly Property GIP() As Decimal
        Get
            Return _GIP
        End Get
    End Property

    Public ReadOnly Property Rebate() As Decimal
        Get
            Return _Rebate
        End Get
    End Property

    Public ReadOnly Property ReadError() As Boolean
        Get
            Return _ReadError
        End Get
    End Property

    Public ReadOnly Property RebatePercent() As Decimal
        Get
            Return _RebatePercent
        End Get
    End Property

    ' Get the settings for a single customer
    Public Sub New(ByVal passCustomer As String)
        Read(passCustomer)
    End Sub

    ' Get the settings for a specific part, in an invoice for a customer
    Public Sub New(ByVal passCustomer As String, ByVal Invoice As String, ByVal Part As String)
        Read(passCustomer, Invoice, Part)
    End Sub

    ' Get the settings for a part within all invoices received within a date range for a customer
    Public Sub New(ByVal passCustomer As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Part As String)
        Read(passCustomer, StartDate, EndDate, Part)
    End Sub

    Private Sub Read(ByVal CUSTID As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CUSTID, ParameterDirection.Input)

        ' Get the data from the stored procedure
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadCustomerSettingsRecord")
            ' look at the first resultset to see how many records were
            ' returned
            dr.Read()
            If dr("Count") = 0 Then
                _ReadError = True
            End If

            ' Set the properties from the data in the datareader if
            ' there was no error reading the record
            If Not ReadError Then
                ' First move to the next resultset
                dr.NextResult()
                ' Go grab the variables
                SetProperties(dr)
            End If
        End Using
    End Sub

    Private Sub Read(ByVal CUSTID As String, ByVal Invoice As String, ByVal Part As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CUSTID, ParameterDirection.Input)
        oSQL.AddParameter("@INVOICE", SqlDbType.NVarChar, 20, Invoice, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, Part, ParameterDirection.Input)

        ' Get the data from the stored procedure
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadCustomerSettingsRecordByInvoice")
            ' Set the properties from the data in the datareader
            SetProperties(dr)
        End Using
    End Sub

    Private Sub Read(ByVal CUSTID As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Part As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CUSTID, ParameterDirection.Input)
        oSQL.AddParameter("@StartDate", SqlDbType.DateTime, 0, StartDate, ParameterDirection.Input)
        oSQL.AddParameter("@EndDate", SqlDbType.DateTime, 0, EndDate, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, Part, ParameterDirection.Input)

        ' Get the data from the stored procedure
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadCustomerSettingsRecordByDate")
            ' Set the properties from the data in the datareader
            SetProperties(dr)
        End Using
    End Sub

    Private Sub SetProperties(ByRef dr As SqlClient.SqlDataReader)
        ' If a record was read, then set the properties
        ' otherwise set them to zero
        If dr.Read() Then
            _GIP = dr("GIP")
            _Rebate = dr("Rebate")
            _RebatePercent = dr("RebatePercent")
            _Freight = dr("Freight")
            _Type = dr("CustomerType")
        Else
            _GIP = 0
            _Rebate = 0
            _RebatePercent = 0
            _Freight = 0
            _Type = 1
        End If
    End Sub

    Public Sub Add(ByVal CUSTID As String, ByVal GIP As Decimal, ByVal Rebate As Decimal, ByVal RebatePercent As Decimal, ByVal Freight As Int16)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CUSTID, ParameterDirection.Input)
        oSQL.AddParameter("@GIP", SqlDbType.Float, 0, GIP, ParameterDirection.Input)
        oSQL.AddParameter("@Rebate", SqlDbType.Float, 0, Rebate, ParameterDirection.Input)
        oSQL.AddParameter("@RebatePercent", SqlDbType.Float, 0, RebatePercent, ParameterDirection.Input)
        oSQL.AddParameter("@Freight", SqlDbType.Int, 0, Freight, ParameterDirection.Input)
        oSQL.AddParameter("@CustomerType", SqlDbType.Int, 0, CustomerType, ParameterDirection.Input)

        ' Get the data from the stored procedure
        oSQL.RunProc("AddCustomerSettings")
    End Sub

    Public Sub Delete(ByVal CUSTID As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CUSTID, ParameterDirection.Input)

        ' Get the data from the stored procedure
        oSQL.RunProc("DeleteCustomerSettings")
    End Sub

    Public Sub Update(ByVal CUSTID As String, ByVal GIP As Decimal, ByVal Rebate As Decimal, ByVal RebatePercent As Decimal, ByVal Freight As Integer)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, CUSTID, ParameterDirection.Input)
        oSQL.AddParameter("@GIP", SqlDbType.Float, 0, GIP, ParameterDirection.Input)
        oSQL.AddParameter("@Rebate", SqlDbType.Float, 0, Rebate, ParameterDirection.Input)
        oSQL.AddParameter("@RebatePercent", SqlDbType.Float, 0, RebatePercent, ParameterDirection.Input)
        oSQL.AddParameter("@Freight", SqlDbType.Int, 0, Freight, ParameterDirection.Input)
        oSQL.AddParameter("@CustomerType", SqlDbType.Int, 0, CustomerType, ParameterDirection.Input)

        ' Get the data from the stored procedure
        oSQL.RunProc("UpdateCustomerSettings")
    End Sub

End Class

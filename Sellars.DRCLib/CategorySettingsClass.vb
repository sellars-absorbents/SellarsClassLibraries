Imports Sellars.SQL
Imports System.Collections
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Public Class CategorySettingsClass
    Inherits ClassBase

    Private _Commission As Decimal = 0
    Private _FreightPerUnit As Decimal = 0
    Private _NascarDiscount As Decimal = 0
    Private _TargetMargin As Decimal = 0
    Private _Description As String = ""

    Public ReadOnly Property Commission() As Decimal
        Get
            Return _Commission
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return _Description
        End Get
    End Property

    Public ReadOnly Property FreightPerUnit() As Decimal
        Get
            Return _FreightPerUnit
        End Get
    End Property

    Public Property NascarDiscount() As Decimal
        Get
            Return _NascarDiscount
        End Get
        Set(ByVal Value As Decimal)
            _NascarDiscount = Value
        End Set
    End Property

    Public ReadOnly Property TargetMargin() As Decimal
        Get
            Return _TargetMargin
        End Get
    End Property

    ' allow someone to initiate a class without any parameters
    ' usefull to do the read method to return all categories so they
    ' can be contained within a listbox control
    Public Sub New()
        MyBase.New()
    End Sub

    ' If you initialize the class with a plan and revlev, then read the specific 
    ' category and set the properties of the class appropriately
    Public Sub New(ByVal passPLANID As String, ByVal passREVLEV As String)
        MyBase.New()
        Read(passPLANID, passREVLEV)
        If TargetMargin = 0 Then
            Read(passPLANID, "")
        End If
    End Sub

    ' Get the settings for a specific part, in an invoice for a customer
    Public Sub New(ByVal passCustomer As String, ByVal Invoice As String, ByVal Part As String)
        MyBase.New()
        Read(passCustomer, Invoice, Part)
    End Sub

    ' Get the settings for a part within all invoices received within a date range for a customer
    Public Sub New(ByVal passCustomer As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Part As String)
        MyBase.New()
        Read(passCustomer, StartDate, EndDate, Part)
    End Sub

    Public Function Read() As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure and return a datareader
        Return oSQL.RunProcReader("ReadAllCategoryTypes")
    End Function

    Public Sub Read(ByVal PLANID As String, ByVal REVLEV As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PLANID", SqlDbType.NVarChar, 3, PLANID, ParameterDirection.Input)
        oSQL.AddParameter("@REVLEV", SqlDbType.NVarChar, 3, REVLEV, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadCategorySettingsRecord")
            ' Set the properties from the read
            SetProperties(dr)
        End Using
    End Sub

    Private Sub Read(ByVal Customer As String, ByVal Invoice As String, ByVal Part As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, Customer, ParameterDirection.Input)
        oSQL.AddParameter("@INVOICE", SqlDbType.NVarChar, 20, Invoice, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, Part, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadCategorySettingsRecordByInvoice")
            ' Set the properties from the read
            SetProperties(dr)
        End Using
    End Sub

    Private Sub Read(ByVal Customer As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Part As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, Customer, ParameterDirection.Input)
        oSQL.AddParameter("@StartDate", SqlDbType.DateTime, 0, StartDate, ParameterDirection.Input)
        oSQL.AddParameter("@EndDate", SqlDbType.DateTime, 0, EndDate, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, Part, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Using dr As SqlDataReader = oSQL.RunProcReader("ReadCategorySettingsRecordByDate")
            ' Set the properties from the read
            SetProperties(dr)
        End Using
    End Sub

    Private Sub SetProperties(ByRef dr As SqlClient.SqlDataReader)
        ' try to read the datareader, and 
        ' if a record found, use the values,
        ' otherwise set the properties to zero
        If dr.Read() Then
            _Commission = dr("Commission")
            _FreightPerUnit = dr("FreightPerPound")
            If IsDBNull(dr("Description")) Then
                _Description = ""
            Else
                _Description = dr("Description")
            End If
            _NascarDiscount = dr("NascarDiscount")
            _TargetMargin = dr("TargetMargin")
        Else
            _Commission = 0
            _FreightPerUnit = 0
            _Description = ""
            _NascarDiscount = 0
            _TargetMargin = 0
        End If
    End Sub

    Public Sub Add(ByVal PLANID As String, ByVal REVLEV As String, ByVal NascarDiscount As Decimal, ByVal TargetMargin As Decimal, ByVal Description As String, ByVal Commission As Decimal, ByVal FreightPerPound As Decimal)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PLANID", SqlDbType.NVarChar, 3, PLANID, ParameterDirection.Input)
        oSQL.AddParameter("@REVLEV", SqlDbType.NVarChar, 3, REVLEV, ParameterDirection.Input)
        oSQL.AddParameter("@NASCAR", SqlDbType.Float, 0, NascarDiscount, ParameterDirection.Input)
        oSQL.AddParameter("@MARGIN", SqlDbType.Float, 0, TargetMargin, ParameterDirection.Input)
        oSQL.AddParameter("@DESCRIPTION", SqlDbType.NVarChar, 50, Description, ParameterDirection.Input)
        oSQL.AddParameter("@Commission", SqlDbType.Float, 0, Commission, ParameterDirection.Input)
        oSQL.AddParameter("@FreightPerPound", SqlDbType.Float, 0, FreightPerPound, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("AddCategory")
    End Sub

    Public Sub Delete(ByVal PLANID As String, ByVal REVLEV As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PLANID", SqlDbType.NVarChar, 3, PLANID, ParameterDirection.Input)
        oSQL.AddParameter("@REVLEV", SqlDbType.NVarChar, 3, REVLEV, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("DeleteCategory")
    End Sub

    Public Sub Update(ByVal PLANID As String, ByVal REVLEV As String, ByVal NascarDiscount As Decimal, ByVal TargetMargin As Decimal, ByVal Description As String, ByVal Commission As Decimal, ByVal FreightPerPound As Decimal)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@PLANID", SqlDbType.NVarChar, 3, PLANID, ParameterDirection.Input)
        oSQL.AddParameter("@REVLEV", SqlDbType.NVarChar, 3, REVLEV, ParameterDirection.Input)
        oSQL.AddParameter("@NASCAR", SqlDbType.Float, 0, NascarDiscount, ParameterDirection.Input)
        oSQL.AddParameter("@MARGIN", SqlDbType.Float, 0, TargetMargin, ParameterDirection.Input)
        oSQL.AddParameter("@Description", SqlDbType.NVarChar, 50, Description, ParameterDirection.Input)
        oSQL.AddParameter("@Commission", SqlDbType.Float, 0, Commission, ParameterDirection.Input)
        oSQL.AddParameter("@FreightPerPound", SqlDbType.Float, 0, FreightPerPound, ParameterDirection.Input)

        ' Run the stored procedure
        oSQL.RunProc("UpdateCategory")
    End Sub
End Class
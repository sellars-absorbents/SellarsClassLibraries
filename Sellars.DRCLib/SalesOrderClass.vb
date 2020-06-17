Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class SalesOrderClass
    Inherits ClassBase

    Private _OrderNumber As String = ""
    Private _Status As String = ""
    Private _Name As String = ""
    Private _Address1 As String = ""
    Private _Address2 As String = ""
    Private _City As String = ""
    Private _State As String = ""
    Private _ZipCode As String = ""
    Private _CustPO As String = ""
    Private _Customer As String = ""

    Public Enum OrderStatus
        Closed = 3
        Open = 4
    End Enum

    Public ReadOnly Property Address1() As String
        Get
            Return _Address1.Trim
        End Get
    End Property

    Public ReadOnly Property Address2() As String
        Get
            Return _Address2.Trim
        End Get
    End Property

    Public ReadOnly Property City() As String
        Get
            Return _City.Trim
        End Get
    End Property

    Public ReadOnly Property CustPO() As String
        Get
            Return _CustPO.Trim
        End Get
    End Property

    Public ReadOnly Property Customer() As String
        Get
            Return _Customer.Trim
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _Name.Trim
        End Get
    End Property

    Public ReadOnly Property OrderNumber() As String
        Get
            Return _OrderNumber.Trim
        End Get
    End Property

    Public ReadOnly Property State() As String
        Get
            Return _State.Trim
        End Get
    End Property

    Public ReadOnly Property Status() As String
        Get
            Return _Status.Trim
        End Get
    End Property

    Public ReadOnly Property ZipCode() As String
        Get
            Return _ZipCode.Trim
        End Get
    End Property

    Public Sub New(ByVal passOrder As String)
        _OrderNumber = passOrder
        Read(passOrder)
    End Sub

    Private Sub Read(ByVal passorder As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@ORDNUM", SqlDbType.NVarChar, 20, passorder, ParameterDirection.Input)

        Using dr As SqlDataReader = oSQL.RunProcReader("ReadSalesOrderMaster")
            ' Assign the variables from the database to properties
            If dr.Read() Then
                _Status = dr("STATUS_27")
                _Address1 = dr("ADDR1_27")
                _Address2 = dr("ADDR2_27")
                _City = dr("CITY_27")
                _Name = dr("NAME_27")
                _State = dr("STATE_27")
                _ZipCode = dr("ZIPCD_27")
                _CustPO = dr("CUSTPO_27")
                _Customer = dr("CUSTID_27")
            Else
                _Status = ""
                _Address1 = ""
                _Address2 = ""
                _City = ""
                _Name = ""
                _State = ""
                _ZipCode = ""
                _CustPO = ""
                _Customer = ""
                Throw New RecordNotOnDatabaseException(passorder, "SalesOrderMaster")
            End If
        End Using
    End Sub
End Class

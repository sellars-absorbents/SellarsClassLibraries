Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class CustomerPartClass
    Inherits ClassBase

    Private _Found As Boolean = False
    Private _OrderNumber As String = ""
    Private _Customer As String = ""
    Private _Part As String = ""
    Private _CustomerPart As String = ""
    Private _Description1 As String = ""
    Private _Description2 As String = ""
    Private _Description3 As String = ""
    Private _Description4 As String = ""

    Public ReadOnly Property Customer() As String
        Get
            Return _Customer
        End Get
    End Property

    Public ReadOnly Property CustomerPart() As String
        Get
            Return _CustomerPart
        End Get
    End Property

    Public ReadOnly Property Description1() As String
        Get
            Return _Description1
        End Get
    End Property

    Public ReadOnly Property Description2() As String
        Get
            Return _Description2
        End Get
    End Property

    Public ReadOnly Property Description3() As String
        Get
            Return _Description3
        End Get
    End Property

    Public ReadOnly Property Description4() As String
        Get
            Return _Description4
        End Get
    End Property

    Public ReadOnly Property Found() As Boolean
        Get
            Return _Found
        End Get
    End Property

    Public ReadOnly Property Part() As String
        Get
            Return _Part
        End Get
    End Property

    Public Sub New(ByVal value As DataSource, ByVal passCustomer As String, ByVal passPart As String)
        _Customer = passCustomer
        _Part = passPart
        ' If the data source says max check the max customer part table, otherwise check the
        ' sql version of the tables
        If value = ClassBase.DataSource.Max Then
            ReadMax(passCustomer, passPart)
        Else
            ReadSellars(passCustomer, passPart)
        End If
    End Sub

    Public Sub New(ByVal value As DataSource, ByVal passCustomer As String, ByVal passPart As String, ByVal passWidth As Decimal)
        _Customer = passCustomer
        _Part = passPart
        ' If the data source says max check the max customer part table, otherwise check the
        ' sql version of the tables
        If value = ClassBase.DataSource.Max Then
            ReadMax(passCustomer, passPart)
        Else
            ReadSellars(passCustomer, passPart)
        End If

        ' If no part was found yet, then check the slit width customer part table
        If Not _Found Then
            ReadSellars(passCustomer, passPart, passWidth)
        End If
    End Sub

    Private Sub ReadMax(ByVal passCustomer As String, ByVal passPart As String)
        OpenMaxConnection()
        Dim strSQL As String = "Select * " & _
                               "From ""Customer_Part_Data"" " & _
                               "Where CUSTID_103 = '" & passCustomer & "' " & _
                               "and PRTNUM_103 = '" & passPart & "'"
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim myReader As SqlDataReader = cmd.ExecuteReader()

        If myReader.Read() Then
            _Found = True
            _CustomerPart = myReader("CUSTPRT_103")
            _Description1 = myReader("PRTDESC1_103")
            _Description2 = myReader("PRTDESC2_103")
            _Description3 = myReader("PRTDESC3_103")
            _Description4 = myReader("PRTDESC4_103")
        Else
            _Found = False
            _CustomerPart = ""
            _Description1 = ""
            _Description2 = ""
            _Description3 = ""
            _Description4 = ""
        End If
        myReader.Close()
        myReader = Nothing
        cmd = Nothing
        CloseMaxConnection()
    End Sub

    Private Sub ReadSellars(ByVal passCustomer As String, ByVal passPart As String)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, passCustomer, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, passPart, ParameterDirection.Input)

        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadCustomerPart")

        ' Assign the variables from the database to properties
        If dr.Read() Then
            _Found = True
            _CustomerPart = dr("CUSTPRT")
            _Description1 = dr("PRTDESC1")
            _Description2 = dr("PRTDESC2")
            _Description3 = dr("PRTDESC3")
            _Description4 = dr("PRTDESC4")
        Else
            _Found = False
            _CustomerPart = ""
            _Description1 = ""
            _Description2 = ""
            _Description3 = ""
            _Description4 = ""
        End If

        dr.Close()
        dr = Nothing

    End Sub

    Private Sub ReadSellars(ByVal passCustomer As String, ByVal passPart As String, ByVal passWidth As Decimal, Optional ByVal passOD As Decimal = 0)
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameter to the command object
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, passCustomer, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, passPart, ParameterDirection.Input)
        oSQL.AddParameter("@Width", SqlDbType.Decimal, 0, passWidth, ParameterDirection.Input)
        oSQL.AddParameter("@OD", SqlDbType.Decimal, 0, passOD, ParameterDirection.Input)

        ' Execute the stored procedure, and if it returned a record, then get the customer part
        Dim dr As SqlClient.SqlDataReader = oSQL.RunProcReader("ReadCustomerSlitWidthPart")
        If dr.Read() Then
            _Found = True
            _CustomerPart = dr("CUSTPRT")
            If IsDBNull(dr("PRTDESC1")) Then
                _Description1 = ""
            Else
                _Description1 = dr("PRTDESC1")
            End If
            If (_Description1 = "'") Or (_Description1 = "-") Then
                _Description1 = ""
            End If

            If IsDBNull(dr("PRTDESC2")) Then
                _Description2 = ""
            Else
                _Description2 = dr("PRTDESC2")
            End If
            If (_Description2 = "'") Or (_Description2 = "-") Then
                _Description2 = ""
            End If

            If IsDBNull(dr("PRTDESC3")) Then
                _Description3 = ""
            Else
                _Description3 = dr("PRTDESC3")
            End If
            If (_Description3 = "'") Or (_Description3 = "-") Then
                _Description3 = ""
            End If

            If IsDBNull(dr("PRTDESC4")) Then
                _Description4 = ""
            Else
                _Description4 = dr("PRTDESC4")
            End If
            If (_Description4 = "'") Or (_Description4 = "-") Then
                _Description4 = ""
            End If

        Else
            ' If we didn't find a match on the specified OD, then we need to try it with the dfault zero OD to see if there
            ' is an entry for that!
            If passOD <> 0 Then
                oSQL = New SqlService(ConnectionString)

                ' Add the parameter to the command object
                oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, passCustomer, ParameterDirection.Input)
                oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, passPart, ParameterDirection.Input)
                oSQL.AddParameter("@Width", SqlDbType.Decimal, 0, passWidth, ParameterDirection.Input)
                oSQL.AddParameter("@OD", SqlDbType.Decimal, 0, 0, ParameterDirection.Input)
                If dr.Read() Then
                    _Found = True
                    _CustomerPart = dr("CUSTPRT")
                    If IsDBNull(dr("PRTDESC1")) Then
                        _Description1 = ""
                    Else
                        _Description1 = dr("PRTDESC1")
                    End If
                    If (_Description1 = "'") Or (_Description1 = "-") Then
                        _Description1 = ""
                    End If

                    If IsDBNull(dr("PRTDESC2")) Then
                        _Description2 = ""
                    Else
                        _Description2 = dr("PRTDESC2")
                    End If
                    If (_Description2 = "'") Or (_Description2 = "-") Then
                        _Description2 = ""
                    End If

                    If IsDBNull(dr("PRTDESC3")) Then
                        _Description3 = ""
                    Else
                        _Description3 = dr("PRTDESC3")
                    End If
                    If (_Description3 = "'") Or (_Description3 = "-") Then
                        _Description3 = ""
                    End If

                    If IsDBNull(dr("PRTDESC4")) Then
                        _Description4 = ""
                    Else
                        _Description4 = dr("PRTDESC4")
                    End If
                    If (_Description4 = "'") Or (_Description4 = "-") Then
                        _Description4 = ""
                    End If
                Else
                    _Found = False
                    _CustomerPart = ""
                    _Description1 = ""
                    _Description2 = ""
                    _Description3 = ""
                    _Description4 = ""
                End If
            End If
        End If

        ' Close the dataset and free up memory
        dr.Close()
        dr = Nothing
        oSQL = Nothing
    End Sub
End Class
Imports System.Data
Imports System.Data.SqlClient

Public Class PaymentTermsClass
    Inherits ClassBase

    Private _Days As Integer = 0
    Private _Discount As Decimal = 0
    Private _DiscountDays As Integer = 0

    Public ReadOnly Property Days() As Integer
        Get
            Return _Days
        End Get
    End Property

    Public ReadOnly Property Discount() As Integer
        Get
            Return _Discount
        End Get
    End Property

    Public ReadOnly Property DiscountDays() As Integer
        Get
            Return _DiscountDays
        End Get
    End Property

    ' allow someone to initiate a class without any parameters
    ' usefull to do the read method to return all categories so they
    ' can be contained within a listbox control
    Public Sub New()
    End Sub

    Public Function Read() As SqlDataReader
        ' Add is only currently available for the max datasource

        Dim strSQL As String = "select CODE_36 as Code, DESC_36 as Description " & _
                                "from ""Code_Master"" " & _
                                "where CDEKEY_36 = 'TERM' " & _
                                "ORDER BY CODE_36 ASC"
        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Return cmd.ExecuteReader(Data.CommandBehavior.CloseConnection)
    End Function

    Public Function Read(ByVal Code As String) As Boolean
        ' Add is only currently available for the max datasource
        Dim rtnVal As Boolean = False

        Try

            Dim strSQL As String = "select @Days = DAYS_36, @Discount = DISC_36, @DiscountDays = DISCDY_36 " & _
                                    "from Code_Master " & _
                                    "where CDEKEY_36 = 'TERM' " & _
                                    "and CODE_36 = @TermCode"
            OpenMaxConnection()
            Dim cmd As New SqlCommand(strSQL, MaxConnection)

            Dim prmDays As New SqlParameter("@Days", SqlDbType.SmallInt, 0)
            prmDays.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmDays)

            Dim prmDiscount As New SqlParameter("@Discount", SqlDbType.SmallInt, 0)
            prmDiscount.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmDiscount)

            Dim prmDiscountDays As New SqlParameter("@DiscountDays", SqlDbType.SmallInt, 0)
            prmDiscountDays.Direction = ParameterDirection.Output
            cmd.Parameters.Add(prmDiscountDays)

            cmd.Parameters.Add(New SqlParameter("@TermCode", Code))

            ' Execute against the database
            cmd.ExecuteNonQuery()

            ' Set the return fields
            _Days = prmDays.Value
            _Discount = (prmDiscount.Value / 10000)
            _DiscountDays = prmDiscountDays.Value

            ' If it makes it here, set the return value to true
            rtnVal = True

        Catch ex As Exception
            rtnVal = False
        Finally
            ' Close the max connection
            CloseMaxConnection()
        End Try

        Return rtnVal
    End Function

    Public Function GetTermCodebyValues(ByVal _days As String, ByVal _disc As Decimal) As String
        Dim discount As Decimal = _disc * 100
        Dim strSQL As String = "select CODE_36 as Code " & _
                        "from ""Code_Master"" " & _
                        "where CDEKEY_36 = 'TERM' " & _
                        "AND DAYS_36 = '" & _days & "' " & _
                        "AND DISC_36 = " & discount.ToString()

        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim code As Object
        code = cmd.ExecuteScalar()
        CloseMaxConnection()

        Return code.ToString()

    End Function
End Class

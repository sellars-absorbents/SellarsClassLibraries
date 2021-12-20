Imports System
Imports System.Collections.Generic
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

    Public ReadOnly Property Discount() As Decimal
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

    Public Function Read() As List(Of PaymentTermsData)
        ' declare return value
        Dim rtnVal As New List(Of PaymentTermsData)

        Dim strSQL As String = "select CODE_36 as Code, DESC_36 as Description " &
                                "from Code_Master with (nolock) " &
                                "where CDEKEY_36 = 'TERM' " &
                                "ORDER BY CODE_36 ASC"

        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        Dim x As New PaymentTermsData()
                        x.Code = dr("Code").ToString().Trim
                        x.Description = dr("Description").ToString().Trim()
                        rtnVal.Add(x)
                    End While
                End Using
            End Using
        End Using

        Return rtnVal
    End Function

    Public Function Read(ByVal Code As String) As Boolean
        Dim rtnVal As Boolean = False
        Dim strSQL As String = "select @Days = DAYS_36, @Discount = DISC_36, @DiscountDays = DISCDY_36 " &
                               "from Code_Master with (nolock) " &
                               "where CDEKEY_36 = 'TERM' " &
                               "and CODE_36 = @TermCode"

        ' try to open another connection to the max database
        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

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
            End Using
        End Using

        Return rtnVal
    End Function

    Public Function GetTermCodebyValues(ByVal _days As String, ByVal _disc As Decimal) As String
        Dim rtnVal As String = ""
        Dim discount As Decimal = _disc * 100
        Dim strSQL As String = "select CODE_36 as Code " &
                        "from Code_Master with (nolock) " &
                        "where CDEKEY_36 = 'TERM' " &
                        "AND DAYS_36 = '" & _days & "' " &
                        "AND DISC_36 = " & discount.ToString()

        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

                rtnVal = cmd.ExecuteScalar().ToSting().Trim()
            End Using
        End Using

        Return rtnVal
    End Function
End Class

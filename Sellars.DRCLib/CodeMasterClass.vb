Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math

Public Class CodeMasterClass
    Inherits ClassBase

    Private _CodeKey As String = ""
    Private _Code As String = ""
    Private _Description As String = ""
    Private _TermDays As String = 0
    Private _DiscountPercent As Integer = 0
    Private _DiscountDays As Integer = 0

    Public ReadOnly Property Code() As String
        Get
            Return _Code
        End Get
    End Property

    Public ReadOnly Property CodeKey() As String
        Get
            Return _CodeKey
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return _Description
        End Get
    End Property

    Public ReadOnly Property Days() As String
        Get
            Return _TermDays
        End Get
    End Property

    Public ReadOnly Property DiscountPercent() As Integer
        Get
            Return _DiscountPercent
        End Get
    End Property

    Public ReadOnly Property DiscountDays() As Integer
        Get
            Return _DiscountDays
        End Get
    End Property

    Public Sub New(ByVal Source As DataSource, ByVal CodeKey As String, ByVal Code As String)
        ' Call the base class new function
        MyBase.New()

        _CodeKey = CodeKey
        _Code = Code

        If Source = ClassBase.DataSource.Sellars Then
            ReadSellars(CodeKey, Code)
        Else
            ReadMax(CodeKey, Code)
        End If
    End Sub

    Private Sub ReadSellars(ByVal CodeKey As String, ByVal Code As String)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@CodeKey", SqlDbType.NVarChar, 4, CodeKey, ParameterDirection.Input)
        oSQL.AddParameter("@Code", SqlDbType.NVarChar, 2, Code, ParameterDirection.Input)

        ' Run the stored procedure
        Dim dr As SqlDataReader = oSQL.RunProcReader("ReadCodeMasterRecord")

        ' Assign the variables from the database to properties
        If dr.Read() Then
            If IsDBNull(dr("DESC_36")) Then
                _Description = ""
            Else
                _Description = dr("DESC_36")
            End If
            If IsDBNull(dr("DAYS_36")) Then
                _TermDays = ""
            Else
                _TermDays = dr("DAYS_36")
            End If
            If IsDBNull(dr("DISC_36")) Then
                _DiscountPercent = 0
            Else
                _DiscountPercent = dr("DISC_36")
            End If
            If IsDBNull(dr("DISCDY_36")) Then
                _DiscountDays = 0
            Else
                _DiscountDays = dr("DISCDY_36")
            End If
        Else
            Throw New RecordNotOnDatabaseException("codekey: " & CodeKey & ", Code: " & Code, "CodeMaster")
        End If

        ' Close the datareader object and free up memory
        dr.Close()
        dr = Nothing
    End Sub

    Private Sub ReadMax(ByVal CodeKey As String, ByVal Code As String)

        Dim strSQL As String = "select * " & _
                               "From ""Code_Master"" " & _
                               "Where CDEKEY_36 = '" & CodeKey.Trim & "' " & _
                               "and CODE_36 = '" & Code.Trim & "'"

        ' Set up the new Sql command
        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim dr As SqlDataReader = cmd.ExecuteReader()

        ' Assign the variables from the database to properties
        If dr.Read() Then
            If IsDBNull(dr("DESC_36")) Then
                _Description = ""
            Else
                _Description = dr("DESC_36")
            End If
            If IsDBNull(dr("DAYS_36")) Then
                _TermDays = ""
            Else
                _TermDays = dr("DAYS_36")
            End If
            If IsDBNull(dr("DISC_36")) Then
                _DiscountPercent = 0
            Else
                _DiscountPercent = dr("DISC_36")
            End If
            If IsDBNull(dr("DISCDY_36")) Then
                _DiscountDays = 0
            Else
                _DiscountDays = dr("DISCDY_36")
            End If
        Else
            Throw New RecordNotOnDatabaseException("codekey: " & CodeKey & ", Code: " & Code, "CodeMaster")
        End If

        ' Close the datareader object and free up memory
        dr.Close()
        dr = Nothing
    End Sub

End Class
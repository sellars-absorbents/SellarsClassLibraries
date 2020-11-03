Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class CustomerMasterExtClass
    Inherits ClassBase

    Private _AcknowledgementLogo As Short = 1

    Public Property AcknowledgementLogo As Short
        Get
            Return _AcknowledgementLogo
        End Get
        Set(value As Short)
            _AcknowledgementLogo = value
        End Set
    End Property

    Private Sub ClearFields()
        _AcknowledgementLogo = 1
    End Sub

    Public Sub New()

    End Sub

    Public Sub New(ByVal CUSTID As String)
        ReadLogo(CUSTID)
    End Sub

    Public Sub ReadLogo(ByVal CUSTID As String)
        ' retrieve the connection string from the task config
        Dim connectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString

        ' Set the submitted flag to true, the submitted date to the current date and time, and the reference field to be the Sellars Distribution
        ' automatically generated invoice number for applicable orders
        Dim strSQL As String = "select @AcknowledgementLogo = AcknowledgementLogo from CustomerMasterExt where CUSTID = @CUSTID"

        Using conn As SqlConnection = New SqlConnection(connectionString)
            ' First try to update the CustomerMasterExt, if not there, then try to add it.
            Try
                ' Open the database connection
                conn.Open()

                ' Set up a new SQL Command
                Using cmd As SqlCommand = New SqlCommand(strSQL, conn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 0

                    ' Add a parameter to the command that has the invoice number to update
                    cmd.Parameters.Add(New SqlParameter("@CUSTID", CUSTID))

                    Dim parmLogo As New SqlParameter("@AcknowledgementLogo", SqlDbType.SmallInt, 0)
                    parmLogo.Direction = ParameterDirection.Output
                    parmLogo.Value = Nothing
                    cmd.Parameters.Add(parmLogo)

                    ' Execute the stored procedure to update the InvoiceMasterExt table
                    cmd.ExecuteNonQuery()

                    ' Set the acknowledgement Logo value
                    _AcknowledgementLogo = parmLogo.Value
                End Using
            Catch ex As Exception
                _AcknowledgementLogo = 1
            End Try
        End Using
    End Sub

    Public Function AddUpdate(ByVal CUSTID As String, ByVal AcknowledgementLogo As Short) As Boolean
        Dim rtnData As Boolean = False

        If String.IsNullOrEmpty(CUSTID.Trim) Then
            Throw New ApplicationException("CustomerMasterExtClass:AddUpdate():CUSTID is empty.")
        End If

        ' retrieve the connection string from the task config
        Dim connectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("Shopfloor").ConnectionString

        ' Set the submitted flag to true, the submitted date to the current date and time, and the reference field to be the Sellars Distribution
        ' automatically generated invoice number for applicable orders
        Dim strSQLUpdate As String = "update CustomerMasterExt set AcknowledgementLogo = @AcknowledgementLogo where CUSTID = @CUSTID"
        Dim strSQLAdd As String = "insert into CustomerMasterExt (CUSTID, AcknowledgementLogo) values (@CUSTID, @AcknowledgementLogo)"

        Using conn As SqlConnection = New SqlConnection(connectionString)
            ' First try to update the CustomerMasterExt, if not there, then try to add it.
            Try
                ' Open the database connection
                conn.Open()

                Dim recsChanged As Integer = 0

                ' Set up a new SQL Command
                Using cmd As SqlCommand = New SqlCommand(strSQLUpdate, conn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 0

                    ' Add a parameter to the command that has the invoice number to update
                    cmd.Parameters.Add(New SqlParameter("@CUSTID", CUSTID))
                    cmd.Parameters.Add(New SqlParameter("@AcknowledgementLogo", AcknowledgementLogo))

                    ' Execute the stored procedure to update the InvoiceMasterExt table
                    recsChanged = cmd.ExecuteNonQuery()
                End Using

                If recsChanged = 0 Then
                    Try
                        ' Set up a new SQL Command
                        Using cmd As SqlCommand = New SqlCommand(strSQLAdd, conn)
                            cmd.CommandType = CommandType.Text
                            cmd.CommandTimeout = 0

                            ' Add a parameter to the command that has the invoice number to update
                            cmd.Parameters.Add(New SqlParameter("@CUSTID", CUSTID))
                            cmd.Parameters.Add(New SqlParameter("@AcknowledgementLogo", AcknowledgementLogo))

                            ' Execute the stored procedure to update the InvoiceMasterExt table
                            cmd.ExecuteNonQuery()
                        End Using

                        rtnData = True
                    Catch exa As Exception
                        rtnData = False
                    End Try
                Else
                    ' Set the return flag to true
                    rtnData = True
                End If
            Catch exu As Exception
                rtnData = False
            End Try
        End Using
    End Function
End Class

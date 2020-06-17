Imports System.Data
Imports System.Data.SqlClient

Public Class TaxMasterClass
    Inherits ClassBase

    Public Function Read(ByVal value As DataSource) As DataTable
        ' Add is only currently available for the max datasource
        If value = DataSource.Sellars Then
            Exit Function
        End If

        ' Create a temporary table
        Dim _TaxList = New DataTable("TaxList")
        _TaxList.Columns.Add("Code", Type.GetType("System.String"))
        _TaxList.Columns.Add("Description", Type.GetType("System.String"))

        ' Add the default settings
        Dim workRow As DataRow
        workRow = _TaxList.NewRow()
        workRow("Code") = ""
        workRow("Description") = "< None >"
        _TaxList.Rows.Add(workRow)

        ' read the data from the database
        Dim strSQL As String = "select TAXCDE_25 as Code, TAXDES_25 as Description " &
                               "from Tax_Master with (nolock) " &
                               "order by TAXDES_25"

        Using MaxConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        workRow = _TaxList.NewRow()
                        workRow("Code") = dr("Code").ToString().Trim()
                        workRow("Description") = dr("Description").ToString().Trim()
                        _TaxList.Rows.Add(workRow)
                    End While
                End Using
            End Using
        End Using

        Return _TaxList
    End Function

    Public Function Read(ByVal TaxCode As String) As Decimal
        ' If there was no supplied tax code, then return zero
        If TaxCode.Trim = "" Then
            Return 0
            Exit Function
        End If

        ' Declare necessary local variables
        Dim TaxRate As Decimal

        ' Set up the SQL Query to get the tax rate
        Dim strSQL As String = "Select TAXRTE_25 " &
                               "From TAX_MASTER with (nolock) " &
                               "Where TAXCDE_25 = '" & TaxCode & "'"

        Using MaxConnection = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        ' Need to multiple by .01 to get the actual rate based on how max stores the data
                        TaxRate = Convert.ToDecimal(dr("TAXRTE_25")) * 0.01
                    Else
                        TaxRate = 0
                    End If
                End Using
            End Using
        End Using

        ' Return the retrieved tax rate
        Return TaxRate
    End Function
End Class

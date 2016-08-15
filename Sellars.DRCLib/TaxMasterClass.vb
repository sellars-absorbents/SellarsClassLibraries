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

        ' try to open another connection to the max database
        OpenMaxConnection()

        ' read the data from the database
        Dim strSQL As String = "select TAXCDE_25 as Code, TAXDES_25 as Description " & _
                               "from ""Tax_Master"" " & _
                               "order by TAXDES_25"
        Dim da As New SqlDataAdapter(strSQL.ToString, MaxConnection)
        Dim ds As New DataSet
        da.Fill(ds, "TaxList")

        ' Insert each row into the taxlist table
        For iRow As Integer = 0 To ds.Tables(0).Rows.Count - 1
            ' Add the default settings
            workRow = _TaxList.NewRow()
            workRow("Code") = ds.Tables(0).Rows(iRow)("Code")
            workRow("Description") = ds.Tables(0).Rows(iRow)("Description")
            _TaxList.Rows.Add(workRow)
        Next

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

        ' Get all the notes that match the criteria
        OpenMaxConnection()
        Dim strSQL As String = "Select TAXRTE_25 " & _
                               "From ""TAX_MASTER"" " & _
                               "Where TAXCDE_25 = '" & TaxCode & "'"
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim myReader As SqlDataReader = cmd.ExecuteReader()

        ' If a record was found, return the tax rate, otherwise return zero for the rate
        If myReader.Read() Then
            ' Need to multiple by .01 to get the actual rate based on how max stores the data
            TaxRate = myReader("TAXRTE_25") * 0.01
        Else
            TaxRate = 0
        End If

        ' Close the datareader/max connection, and free up memory
        myReader.Close()
        myReader = Nothing
        CloseMaxConnection()

        ' Return the retrieved tax rate
        Return TaxRate
    End Function
End Class

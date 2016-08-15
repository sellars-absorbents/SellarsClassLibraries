Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient

Public Class MaxCompanies
    Inherits ClassBase

    Public Sub New()

    End Sub

    Public Function Read() As SqlDataReader
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Run the stored procedure
        Return oSQL.RunProcReader("GetMaxCompanies")
    End Function

    Public Function Read(ByVal Path As String) As String
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@Path", SqlDbType.NVarChar, 50, Path, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Dim dr As SqlDataReader = oSQL.RunProcReader("GetMaxCompanyByPath")

        ' Set the properties from the read
        Dim MaxCompany As String = ""
        If dr.Read() Then
            MaxCompany = dr("Name")
        End If

        ' Close the dataset, connection and free up memory
        dr.Close()
        dr = Nothing

        ' Return the Company Name
        Return MaxCompany
    End Function
End Class
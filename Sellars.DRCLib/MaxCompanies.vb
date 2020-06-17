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
        Dim MaxCompany As String = ""
        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@Path", SqlDbType.NVarChar, 50, Path, ParameterDirection.Input)

        ' Run the stored procedure and return a datareader
        Using dr As SqlDataReader = oSQL.RunProcReader("GetMaxCompanyByPath")

            If dr.Read() Then
                MaxCompany = dr("Name")
            End If
        End Using

        ' Return the Company Name
        Return MaxCompany
    End Function
End Class
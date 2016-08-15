Imports System.Data
Imports System.Data.SqlClient

Public Class DocumentLinkClass
    Inherits ClassBase

    Public Enum DocumentType
        CUSTID = 0
        ORDNUM = 1
        PRTNUM = 2
        VENID = 3
    End Enum

    Private DocumentTypes() As String = {"CUSTID", "ORDNUM", "PRTNUM", "VENID"}

    Public Function Read(ByVal passType As DocumentType, ByVal passValue As String) As SqlDataReader
        OpenMaxConnection()
        Dim strSQL As String = "Select PATH_85 as Path " & _
                               "From ""Document_Link_Mst"" " & _
                               "Where TYPE_85 = '" & DocumentTypes(passType) & "' AND VALUE_85 = '" & passValue & "' "
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Return cmd.ExecuteReader(CommandBehavior.CloseConnection)
    End Function

End Class

Imports System.Collections.Generic
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

    Public Function Read(ByVal passType As DocumentType, ByVal passValue As String) As List(Of String)
        Dim rtnValues As New List(Of String)

        Dim strSQL As String = "Select PATH_85 as Path " &
                               "From Document_Link_Mst with (nolock) " &
                               "Where TYPE_85 = '" & DocumentTypes(passType) & "' AND VALUE_85 = '" & passValue & "' "

        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

                Using dr As SqlDataReader = cmd.ExecuteReader()

                    While dr.Read()
                        rtnValues.Add(dr("Path").ToString().Trim())
                    End While

                End Using

            End Using
        End Using

        Return rtnValues
    End Function

End Class

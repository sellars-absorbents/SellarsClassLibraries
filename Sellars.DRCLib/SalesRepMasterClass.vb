Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math

Public Class SalesRepMasterClass
    Inherits ClassBase

    Private _SalesRepID As String = ""
    Private _Name As String = ""

    Public Sub New(ByVal Source As DataSource, ByVal ID As String)
        _SalesRepID = ID
        Read(Source, ID)
    End Sub

    Public ReadOnly Property ID() As String
        Get
            Return _SalesRepID
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    Private Sub Read(ByVal Source As DataSource, ByVal ID As String)

        ' Declare necessary local variables and initialize them
        Dim CustomerPrice As Decimal = 0
        Dim strSQL As String = "Select SLSNME_26 " &
                                "From Sales_Rep_Master with (nolock) " &
                                "Where SLSREP_26 = '" & ID.Trim & "' "

        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            Using cmd As New SqlCommand(strSQL, MaxConnection)
                cmd.CommandType = CommandType.Text

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        If IsDBNull(dr("SLSNME_26")) Then
                            _Name = ""
                        Else
                            _Name = dr("SLSNME_26").ToString().Trim()
                        End If
                    Else
                        _Name = ""
                    End If
                End Using
            End Using
        End Using
    End Sub

End Class

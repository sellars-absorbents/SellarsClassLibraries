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

        If Source = DataSource.Sellars Then
            ' Declare the SQL data layer class
            Dim oSQL As New SqlService(ConnectionString)

            ' Add the parameters to the command object
            oSQL.AddParameter("@ID", SqlDbType.NVarChar, 4, ID, ParameterDirection.Input)

            Dim dr As SqlDataReader = oSQL.RunProcReader("GetMaxSalesRepMaster")

            ' Assign the variables from the database to properties
            If dr.Read() Then
                If IsDBNull(dr("Name")) Then
                    _Name = ""
                Else
                    _Name = dr("Name")
                End If
            End If

            dr.Close()
            dr = Nothing
        Else
            ' try to open another connection to the max database
            OpenMaxConnection()

            ' Declare necessary local variables and initialize them
            Dim CustomerPrice As Decimal = 0
            Dim strSQL As String = "Select SLSNME_26 " & _
                                   "From ""Sales_Rep_Master"" " & _
                                   "Where SLSREP_26 = '" & ID.Trim & "' "

            ' Set up the new Sql command and retrieve the data from the database
            Dim cmd As New SqlCommand(strSQL, MaxConnection)
            Dim SalesRepMasterReader As SqlDataReader = cmd.ExecuteReader()
            If SalesRepMasterReader.Read() Then
                If IsDBNull(SalesRepMasterReader("SLSNME_26")) Then
                    _Name = ""
                Else
                    _Name = SalesRepMasterReader("SLSNME_26")
                End If
            Else
                _Name = ""
            End If
            SalesRepMasterReader.Close()
            SalesRepMasterReader = Nothing
            CloseMaxConnection()
        End If

    End Sub

End Class

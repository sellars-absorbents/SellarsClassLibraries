Imports System.Data
Imports System.Data.SqlClient

Public Class StockMasterClass
    Inherits ClassBase

    Dim _recordfound As Boolean = False

    Public ReadOnly Property RecordFound() As Boolean
        Get
            Return _recordfound
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal passStockCode As String)
        _recordfound = Read(passStockCode)
    End Sub

    Public Function Read(ByVal passStockCode As String) As Boolean
        Dim bRecordFound As Boolean = False

        ' try to open another connection to the max database
        OpenMaxConnection()

        Dim strSQL As String = "Select * " & _
                               "From ""Stock_Master"" " & _
                               "Where STK_05 = '" & passStockCode.Trim & "' "

        ' Set up the new Sql command
        Using cmd As New SqlCommand(strSQL, MaxConnection)
            Using PartSalesReader As SqlDataReader = cmd.ExecuteReader()
                If PartSalesReader.Read() Then
                    bRecordFound = True
                Else
                    bRecordFound = False
                End If
            End Using
        End Using

        CloseMaxConnection()

        ' Return whether or not the stock code was found on the Stock Master table
        Return bRecordFound
    End Function

End Class

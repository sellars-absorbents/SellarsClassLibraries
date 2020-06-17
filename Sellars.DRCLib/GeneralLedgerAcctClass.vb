Imports System.Data
Imports System.Data.SqlClient

Public Class GeneralLedgerAcctClass
    Inherits ClassBase

    Private _GLAcct As String

    Public ReadOnly Property GLAcct() As String
        Get
            Return _GLAcct.Trim
        End Get
    End Property

    Public Function Read(ByVal passacttyp As String) As String
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        _GLAcct = ""
        Dim strSQL As String = "Select GLACCT_19 " & _
                               "From ""General_Ledger_Acct"" " & _
                               "Where KEY_19 = 'CSTOFSLS' " & _
                               "and RECTYP_19 = 'D' " & _
                               "and ACTTYP_19 = '" & passacttyp.Trim & "' "

        ' Set up the new Sql command
        Using cmd As New SqlCommand(strSQL, MaxConnection)
            Using glReader As SqlDataReader = cmd.ExecuteReader()
                glReader.Read()
                _GLAcct = glReader("GLACCT_19")
            End Using
        End Using

        CloseMaxConnection()

        ' Return the price from the Price Breaks table that is appropriate for the 
        ' item / customer type / quantity submitted
        Return GLAcct.Trim
    End Function

    Public Function Read() As String
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        _GLAcct = ""
        Dim strSQL As String = "Select GLACCT_19 " & _
                               "From ""General_Ledger_Acct"" " & _
                               "Where KEY_19 = 'SLSACT' " & _
                               "and RECTYP_19 = 'D' " & _
                               "and ACTTYP_19 = 'A' "

        ' Set up the new Sql command
        Using cmd As New SqlCommand(strSQL, MaxConnection)
            Using glReader As SqlDataReader = cmd.ExecuteReader()
                glReader.Read()
                _GLAcct = glReader("GLACCT_19")
            End Using
        End Using

        CloseMaxConnection()

        ' Return the price from the Price Breaks table that is appropriate for the 
        ' item / customer type / quantity submitted
        Return GLAcct.Trim
    End Function

    Public Function Verify(ByVal passact As String) As Boolean
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them
        Dim bVerified As Boolean = False
        Dim strSQL As String = "Select GLACCT_19 " &
                               "From ""General_Ledger_Acct"" " &
                               "Where GLACCT_19 = '" & passact.Trim & "' "

        ' Set up the new Sql command
        Using cmd As New SqlCommand(strSQL, MaxConnection)
            Using glReader As SqlDataReader = cmd.ExecuteReader()
                bVerified = glReader.HasRows
            End Using
        End Using

        CloseMaxConnection()

        ' Return the price from the Price Breaks table that is appropriate for the 
        ' item / customer type / quantity submitted
        Return bVerified
    End Function
End Class
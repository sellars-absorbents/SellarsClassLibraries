Public Class PrinterNotFoundException
    Inherits ApplicationException

    Private _Printer As Integer
    Private _Location As Integer
    Private _Line As Integer

    Public ReadOnly Property Printer() As Integer
        Get
            Return _Printer
        End Get
    End Property

    Public ReadOnly Property Line() As Integer
        Get
            Return _Line
        End Get
    End Property

    Public ReadOnly Property Location() As Integer
        Get
            Return _Location
        End Get
    End Property

    Public Overrides ReadOnly Property Message() As String
        Get
            Return "Printer number: " & Printer & " for location: " & Location & ", Line: " & Line & " was not in the Printers table."
        End Get
    End Property

    Public Sub New(ByVal Location As Integer, ByVal Line As Integer, ByVal Printer As Integer)
        _Printer = Printer
        _Location = Location
        _Line = Line
    End Sub
End Class

Public Class RecordNotOnDatabaseException
    Inherits ApplicationException

    Private _strKey As String
    Private _iKey As String
    Private _Table As String

    Private _DataType As messagetype
    Enum messagetype
        stringkey = 0
        Integerkey = 1
    End Enum

    Public ReadOnly Property IntegerKey() As Integer
        Get
            Return _iKey
        End Get
    End Property

    Public ReadOnly Property StringKey() As String
        Get
            Return _strKey
        End Get
    End Property

    Public ReadOnly Property TableName() As String
        Get
            Return _Table
        End Get
    End Property

    Public Overrides ReadOnly Property Message() As String
        Get
            Select Case _DataType
                Case messagetype.Integerkey
                    Return "Record with key value of: " & IntegerKey.ToString.Trim & " was not in the " & TableName & " table."
                Case messagetype.stringkey
                    Return "Record with key value of: " & StringKey & " was not in the " & TableName & " table."
            End Select
        End Get
    End Property

    Public Sub New(ByVal Key As String, ByVal tablename As String)
        _strKey = Key
        _Table = tablename
        _DataType = messagetype.stringkey
    End Sub

    Public Sub New(ByVal Key As Integer, ByVal tablename As String)
        _iKey = Key
        _Table = tablename
        _DataType = messagetype.Integerkey
    End Sub
End Class

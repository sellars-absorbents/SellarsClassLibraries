Public Class NumberClass

    Private _BaseNumber As Double

    Public Sub New()

    End Sub

    Public Sub New(ByVal x As Double)
        _BaseNumber = x
    End Sub

    Public ReadOnly Property SignificantDigits() As Int16
        Get
            Dim _BaseString As String = _BaseNumber.ToString
            If _BaseString.IndexOf(".") = -1 Then
                Return 0
            Else
                Return _BaseString.Length - _BaseString.IndexOf(".")
            End If
        End Get
    End Property

    Public ReadOnly Property SignificantDigits(ByVal _PassNumber As Double) As Int16
        Get
            Dim _BaseString As String = _PassNumber.ToString
            If _BaseString.IndexOf(".") = -1 Then
                Return 0
            Else
                ' Need to have the -1 to start after the decimal place!
                Return _BaseString.Length - _BaseString.IndexOf(".") - 1
            End If
        End Get
    End Property

End Class

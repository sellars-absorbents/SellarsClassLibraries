Imports System
Imports System.Collections.Generic
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.ServiceModel

Public Class SalesOrderDetailNotes

    Private _DoNotPrint As String = ""
    Private _PrintOnInvoiceOnly As String = ""
    Private _PrintOnOrderOnly As String = ""
    Private _PrintOnBoth As String = ""

    Public Property DoNotPrint() As String
        Get
            Return _DoNotPrint.Trim
        End Get
        Set(ByVal value As String)
            _DoNotPrint = GetValue(value, "")
        End Set
    End Property

    Public Property PrintOnBoth() As String
        Get
            Return _PrintOnBoth.Trim
        End Get
        Set(ByVal value As String)
            _PrintOnBoth = GetValue(value, "")
        End Set
    End Property

    Public Property PrintOnInvoiceOnly() As String
        Get
            Return _PrintOnInvoiceOnly.Trim
        End Get
        Set(ByVal value As String)
            _PrintOnInvoiceOnly = GetValue(value, "")
        End Set
    End Property

    Public Property PrintOnOrderOnly() As String
        Get
            Return _PrintOnOrderOnly.Trim
        End Get
        Set(ByVal value As String)
            _PrintOnOrderOnly = GetValue(value, "")
        End Set
    End Property


    Private Function GetValue(Of T)(ByVal value As T, ByVal defaultValue As T) As T
        If value Is Nothing Then
            Return defaultValue
        Else
            ' If a string was pased in, then make sure that we return the trimmed value
            If TypeOf (value) Is String Then
                Return Convert.ChangeType(value.ToString().Trim(), GetType(T))
            Else
                Return value
            End If
        End If
    End Function

End Class

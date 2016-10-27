Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient

Public Class PartSalesClass
    Inherits ClassBase

    Private _Price As Decimal
    Private _Description1 As String
    Private _Description2 As String
    Private _Conversion As Decimal = 0

    Private _PriceBreaks As New Collection

    Public ReadOnly Property Conversion() As Decimal
        Get
            Return _Conversion
        End Get
    End Property

    Public ReadOnly Property Price() As Decimal
        Get
            Return _Price
        End Get
    End Property

    Public ReadOnly Property Description1() As String
        Get
            Return _Description1
        End Get
    End Property

    Public ReadOnly Property Description2() As String
        Get
            Return _Description2
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal passpart As String)
        Read(passpart)
    End Sub

    Public Function IsTaxable(ByVal passpart As String) As Boolean
        ' Declare necessary local variables
        Dim taxabl As Boolean = False

        ' try to open another connection to the max database
        OpenMaxConnection()

        Dim strSQL As String = "Select * " & _
                               "From ""Part_Sales"" " & _
                               "Where PRTNUM_29 = '" & passpart.Trim & "' "

        ' Set up the new Sql command and read the database
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim PartSalesReader As SqlDataReader = cmd.ExecuteReader()

        ' if the record was found, and the taxabl indicator set to yes, return true from this function
        If PartSalesReader.Read() Then
            If PartSalesReader("TAXABL_29") = "Y" Then
                taxabl = True
            End If
        End If

        ' Close the reader, max connection and free up memory
        PartSalesReader.Close()
        PartSalesReader = Nothing
        closeMaxConnection()

        ' return the appropriate value showing whether the part is taxable
        Return taxabl
    End Function

    Public Sub Read(ByVal passpart As String)
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them

        ' Loop through and clear off the old price break data
        For Each pbd As PriceBreakData In _PriceBreaks
            _PriceBreaks.Remove(1)
        Next

        Dim CustomerPrice As Decimal = 0
        _Price = 0
        Dim strSQL As String = "Select PMDES1_29, PMDES2_29, SLSCNV_29  " & _
                               "From ""Part_Sales"" " & _
                               "Where PRTNUM_29 = '" & passpart.Trim & "' "

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim PartSalesReader As SqlDataReader = cmd.ExecuteReader()
        If PartSalesReader.Read() Then
            _Description1 = PartSalesReader("PMDES1_29")
            _Description2 = PartSalesReader("PMDES2_29")
            _Conversion = PartSalesReader("SLSCNV_29")
        Else
            _Description1 = ""
            _Description2 = ""
            _Conversion = 1
        End If

        PartSalesReader.Close()
        PartSalesReader = Nothing
        CloseMaxConnection()
    End Sub

    Public Function Read(ByVal passpart As String, ByVal passqty As Decimal) As Decimal
        ' try to open another connection to the max database
        OpenMaxConnection()

        ' Declare necessary local variables and initialize them

        ' Loop through and clear off the old price break data
        For Each pbd As PriceBreakData In _PriceBreaks
            _PriceBreaks.Remove(1)
        Next

        Dim CustomerPrice As Decimal = 0
        _Price = 0
        Dim strSQL As String = "Select * " & _
                               "From ""Part_Sales"" " & _
                               "Where PRTNUM_29 = '" & passpart.Trim & "' "

        ' Set up the new Sql command
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        Dim PartSalesReader As SqlDataReader = cmd.ExecuteReader()
        PartSalesReader.Read()
        _Price = PartSalesReader("PRICE_29")

        SetPriceBreak(PartSalesReader("PRICE1_29"), PartSalesReader("BREAK1_29"), PartSalesReader("DISC1_29"))
        SetPriceBreak(PartSalesReader("PRICE2_29"), PartSalesReader("BREAK2_29"), PartSalesReader("DISC2_29"))
        SetPriceBreak(PartSalesReader("PRICE3_29"), PartSalesReader("BREAK3_29"), PartSalesReader("DISC3_29"))
        SetPriceBreak(PartSalesReader("PRICE4_29"), PartSalesReader("BREAK4_29"), PartSalesReader("DISC4_29"))
        SetPriceBreak(PartSalesReader("PRICE5_29"), PartSalesReader("BREAK5_29"), PartSalesReader("DISC5_29"))
        SetPriceBreak(PartSalesReader("PRICE6_29"), PartSalesReader("BREAK6_29"), PartSalesReader("DISC6_29"))
        SetPriceBreak(PartSalesReader("PRICE7_29"), PartSalesReader("BREAK7_29"), PartSalesReader("DISC7_29"))
        SetPriceBreak(PartSalesReader("PRICE8_29"), PartSalesReader("BREAK8_29"), PartSalesReader("DISC8_29"))
        SetPriceBreak(PartSalesReader("PRICE9_29"), PartSalesReader("BREAK9_29"), PartSalesReader("DISC9_29"))

        PartSalesReader.Close()
        PartSalesReader = Nothing
        closeMaxConnection()

        If _PriceBreaks.Count = 0 Then
            CustomerPrice = _Price
        Else
            For Each pbd As PriceBreakData In _PriceBreaks
                If passqty >= pbd.BreakQty Then
                    CustomerPrice = pbd.Price
                End If
            Next
        End If

        ' If we get here that means that we had a price break record for that 
        ' part/customer combination, and not enough quantity to meet the minimum
        ' so simply return the generic price from the record
        If CustomerPrice = 0 Then
            CustomerPrice = Price
        End If

        ' Return the price from the Price Breaks table that is appropriate for the 
        ' item / customer type / quantity submitted
        Return CustomerPrice
    End Function

    Private Sub SetPriceBreak(ByVal passPrice, ByVal passQTY, ByVal passDisc)
        If passQTY > 0 Then
            Dim pbd As New PriceBreakData(passPrice, passQTY, passDisc)
            _PriceBreaks.Add(pbd)
            pbd = Nothing
        End If
    End Sub

    Private Class PriceBreakData
        Public Price As Decimal
        Public BreakQty As Decimal
        Public DiscountPercent As Decimal

        Public Sub New(ByVal passPrice As Decimal, ByVal passBreak As Decimal, ByVal passDisc As Decimal)
            Price = passPrice
            BreakQty = passBreak
            DiscountPercent = passDisc
        End Sub
    End Class
End Class

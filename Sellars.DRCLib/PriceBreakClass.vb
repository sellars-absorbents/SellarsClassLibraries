Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient

Public Class PriceBreakClass
    Inherits ClassBase

    Private _Price As Decimal

    Private _PriceBreaks As New Collection

    Public ReadOnly Property Price() As Decimal
        Get
            Return _Price
        End Get
    End Property

    Public Function Read(ByVal passpart As String, ByVal passcustyp As String, ByVal passqty As Decimal) As Decimal
        ' Loop through and clear off the old price break data
        For Each pbd As PriceBreakData In _PriceBreaks
            _PriceBreaks.Remove(1)
        Next

        ' Declare necessary local variables
        Dim CustomerPrice As Decimal = 0
        _Price = 0

        Dim strSQL As String = "Select * " &
                               "From Price_Breaks with (nolock) " &
                               "Where PRTNUM_88 = '" & passpart.Trim & "' " &
                               "and CUSTYP_88 = '" & passcustyp.Trim & "' "

        ' try to open another connection to the max database
        Using MaxConnection As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("MaxData").ConnectionString)
            MaxConnection.Open()

            ' Set up the new Sql command
            Using cmd As New SqlCommand(strSQL, MaxConnection)

                Using PartSalesReader As SqlDataReader = cmd.ExecuteReader()
                    PartSalesReader.Read()
                    _Price = PartSalesReader("PRICE_88")

                    SetPriceBreak(PartSalesReader("PRICE1_88"), PartSalesReader("BREAK1_88"), PartSalesReader("DISC1_88"))
                    SetPriceBreak(PartSalesReader("PRICE2_88"), PartSalesReader("BREAK2_88"), PartSalesReader("DISC2_88"))
                    SetPriceBreak(PartSalesReader("PRICE3_88"), PartSalesReader("BREAK3_88"), PartSalesReader("DISC3_88"))
                    SetPriceBreak(PartSalesReader("PRICE4_88"), PartSalesReader("BREAK4_88"), PartSalesReader("DISC4_88"))
                    SetPriceBreak(PartSalesReader("PRICE5_88"), PartSalesReader("BREAK5_88"), PartSalesReader("DISC5_88"))
                    SetPriceBreak(PartSalesReader("PRICE6_88"), PartSalesReader("BREAK6_88"), PartSalesReader("DISC6_88"))
                    SetPriceBreak(PartSalesReader("PRICE7_88"), PartSalesReader("BREAK7_88"), PartSalesReader("DISC7_88"))
                    SetPriceBreak(PartSalesReader("PRICE8_88"), PartSalesReader("BREAK8_88"), PartSalesReader("DISC8_88"))
                End Using
            End Using
        End Using

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
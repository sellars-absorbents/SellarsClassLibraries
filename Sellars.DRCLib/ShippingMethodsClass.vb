Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Public Class ShippingMethodsClass
    Inherits ClassBase

    ' allow someone to initiate a class without any parameters
    ' usefull to do the read method to return all categories so they
    ' can be contained within a listbox control
    Public Sub New()
    End Sub

    ' Function to add a new record to the Max Code Master table
    Public Sub Add(ByVal CodeKey As String, ByVal Code As String, ByVal Description As String)
        ' try to open another connection to the max database
        ' Declare necessary local variables and initialize them

        Dim ActTyp As String = ""
        Dim Curr As String = ""
        Dim DayOmn As Integer = 0
        Dim Days As String = ""
        Dim DecPlc As Integer = 0
        Dim Desc As String = ""
        Dim Disc As Integer = 0
        Dim DiscDy As Integer = 0
        Dim ExcRte As Integer = 0
        Dim Filler As String = ""
        Dim Mcomp As String = ""
        Dim Msite As String = ""
        Dim Symbol As String = ""
        Dim UdfKey As String = ""
        Dim UdfRef As String = ""

        Dim strSQL As String = "Insert into ""Code_Master"" (ACTTYP_36, CDEKEY_36, CODE_36, CURR_36, DAYOMN_36, DAYS_36, DECPLC_36, DESC_36, DISC_36, DISCDY_36, EXCRTE_36, FILLER_36, MCOMP_36, MSITE_36, SYMBOL_36, UDFKEY_36, UDFREF_36) " & _
                               "values ('" & ActTyp & "','" & CodeKey.ToUpper & "','" & Code.ToUpper & "','" & Curr & "'," & DayOmn & ",'" & Days & "'," & DecPlc & ",'" & Description & "'," & Disc & "," & DiscDy & "," & ExcRte & ",'" & Filler & "','" & Mcomp & "','" & Msite & "','" & Symbol & "','" & UdfKey & "','" & UdfRef & "')"
        OpenMaxConnection()
        Dim cmd As New SqlCommand(strSQL, MaxConnection)
        cmd.ExecuteNonQuery()
        CloseMaxConnection()
    End Sub

    Public Function Read() As DataSet
        ' Add is only currently available for the max datasource

        Dim strSQL As String = "select CODE_36 as Code, DESC_36 as Description, DESC_36 + ' (' + CODE_36 + ')' as DescriptionAndCode " & _
                                "from ""Code_Master"" " & _
                                "where CDEKEY_36 = 'SHIP' " & _
                                "ORDER BY CODE_36 ASC"
        OpenMaxConnection()
        Dim da As New SqlDataAdapter(strSQL.ToString, MaxConnection)
        Dim ds As New DataSet
        da.Fill(ds, "ShipVias")
        Return ds
    End Function
End Class

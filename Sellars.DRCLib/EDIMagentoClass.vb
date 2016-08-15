Imports Sellars.SQL

Public Class EDIMagentoClass
    Inherits ClassBase

    Private _termsCode As String
    Private _shipViaCode As String
    Private _OrderedBy As String
    Private _FOB As String
    Private _customer As CustomerClass
    Private _TaxCode1 As String
    Private _TaxCode2 As String
    Private _TaxCode3 As String
    Private _TaxRate1 As Decimal
    Private _TaxRate2 As Decimal
    Private _TaxRate3 As Decimal
    Private _ShipToCode As String
    Private _ShipToName As String
    Private _ShipToAddress1 As String
    Private _ShipToAddress2 As String
    Private _ShipToCity As String
    Private _ShipToState As String
    Private _ShipToZip As String
    Private _ShipToCountry As String
    Private _OrderLines As DataSet
    Private _TotalPrice As Decimal
    Private _ExtNotes As String
    Private _Comment1 As String
    Private _Comment2 As String
    Private _Comment3 As String
    Private _PONumber As String

    Public ReadOnly Property TermsCode() As String
        Get
            Return _termsCode.Trim()
        End Get
    End Property

    Public ReadOnly Property ShipViaCode() As String
        Get
            Return _shipViaCode.Trim()
        End Get
    End Property

    Public ReadOnly Property OrderedBy() As String
        Get
            Return _OrderedBy.Trim()
        End Get
    End Property

    Public ReadOnly Property FOB() As String
        Get
            Return _FOB.Trim()
        End Get
    End Property

    Public ReadOnly Property Customer() As CustomerClass
        Get
            Return _customer
        End Get
    End Property

    Public ReadOnly Property ShipToCode() As String
        Get
            Return _ShipToCode.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToName() As String
        Get
            Return _ShipToName.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToAddress1() As String
        Get
            Return _ShipToAddress1.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToAddress2() As String
        Get
            Return _ShipToAddress2.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToCity() As String
        Get
            Return _ShipToCity.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToState() As String
        Get
            Return _ShipToState.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToZip() As String
        Get
            Return _ShipToZip.Trim
        End Get
    End Property

    Public ReadOnly Property ShipToCountry() As String
        Get
            Return _ShipToCountry.Trim
        End Get
    End Property

    Public ReadOnly Property TaxCode1() As String
        Get
            Return _TaxCode1.Trim
        End Get
    End Property

    Public ReadOnly Property TaxCode2() As String
        Get
            Return _TaxCode2.Trim
        End Get
    End Property

    Public ReadOnly Property TaxCode3() As String
        Get
            Return _TaxCode3.Trim
        End Get
    End Property

    Public ReadOnly Property TaxRate1() As Decimal
        Get
            Return _TaxRate1
        End Get
    End Property

    Public ReadOnly Property TaxRate2() As Decimal
        Get
            Return _TaxRate2
        End Get
    End Property

    Public ReadOnly Property TaxRate3() As Decimal
        Get
            Return _TaxRate3
        End Get
    End Property

    Public ReadOnly Property OrderLines() As DataSet
        Get
            Return _OrderLines
        End Get

    End Property

    Public ReadOnly Property TotalPrice() As Decimal
        Get
            Return _TotalPrice
        End Get
    End Property

    Public ReadOnly Property ExtNotes() As String
        Get
            Return _ExtNotes.Trim
        End Get
    End Property

    Public ReadOnly Property Comment1() As String
        Get
            Return _Comment1.Trim
        End Get
    End Property

    Public ReadOnly Property Comment2() As String
        Get
            Return _Comment2.Trim
        End Get
    End Property

    Public ReadOnly Property Comment3() As String
        Get
            Return _Comment3.Trim
        End Get
    End Property

    Public ReadOnly Property PONumber() As String
        Get
            Return _PONumber.Trim
        End Get
    End Property

    Public Sub New(ByVal _dsOrderHeader As DataSet, ByVal _dsOrderLines As DataSet, ByVal _dsCustomerXref As DataSet)

        'PO Number
        If Not (IsDBNull(_dsOrderHeader.Tables(0).Rows(0).Item("CUST_PO"))) Then
            _PONumber = _dsOrderHeader.Tables(0).Rows(0).Item("CUST_PO")
        Else
            _PONumber = "WEB" & _dsOrderHeader.Tables(0).Rows(0).Item("MAGENTO_ORD_NUM")
        End If

        'ShipVia
        _shipViaCode = "10"

        'Ordered By
        _OrderedBy = "WEB"

        'FOB
        _FOB = "ORIGIN"

        'Customer
        Dim _custid As String = _dsCustomerXref.Tables(0).Rows(0).Item("CUST_ID")
        _customer = New CustomerClass(_custid, DataSource.Max)

        'Terms
        _termsCode = _customer.TermCode

        'Tax Codes
        Dim myTax As New TaxMasterClass

        'Ship To Address
        _ShipToCode = ""

        'Ship To Company Name
        If (Not IsDBNull(_dsOrderHeader.Tables(0).Rows(0).Item("SHIP_COMPANY_NAME"))) Then
            If _dsOrderHeader.Tables(0).Rows(0).Item("SHIP_COMPANY_NAME").ToString <> "" Then
                _ShipToName = _dsOrderHeader.Tables(0).Rows(0).Item("SHIP_COMPANY_NAME")
            Else
                _ShipToName = _customer.Name
            End If
        Else
            _ShipToName = _customer.Name
        End If

        _ShipToAddress1 = UCase(_dsOrderHeader.Tables(0).Rows(0).Item("SHIP_ADDR1"))
        _ShipToAddress2 = ""
        _ShipToCity = UCase(_dsOrderHeader.Tables(0).Rows(0).Item("SHIP_CITY"))
        _ShipToState = GetState(_dsOrderHeader.Tables(0).Rows(0).Item("SHIP_STATE"))
        _ShipToZip = UCase(_dsOrderHeader.Tables(0).Rows(0).Item("SHIP_ZIP"))
        _ShipToCountry = UCase(_dsOrderHeader.Tables(0).Rows(0).Item("SHIP_COUNTRY"))

        'Tax Information
        _TaxCode1 = _customer.TaxCode1
        _TaxCode2 = _customer.TaxCode2
        _TaxCode3 = _customer.TaxCode3
        _TaxRate1 = myTax.Read(_customer.TaxCode1)
        _TaxRate2 = myTax.Read(_customer.TaxCode2)
        _TaxRate3 = myTax.Read(_customer.TaxCode3)

        'Lines
        Dim myDataTable As DataTable = New DataTable("SODetail")
        Dim myDataColumn As DataColumn
        Dim myDataRow As DataRow

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "LINNUM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "DELNUM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PRTNUM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "CURQTY"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "CURDUE"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.DateTime")
        myDataColumn.ColumnName = "CUSDUE"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "UOM"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PrtUnitPrice"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "Pallets"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "GLCode"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "LINDEL"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "DUESHIP"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "PARTEXT"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "BothDesc"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "UOMPALLETS"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "EXTPRICES"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "REQSHP"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "ShortComment"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Boolean")
        myDataColumn.ColumnName = "Closed"
        myDataColumn.ReadOnly = False
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "Taxable"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "TaxCode1"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "TaxCode2"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "TaxCode3"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "TaxRate1"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "TaxRate2"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.Decimal")
        myDataColumn.ColumnName = "TaxRate3"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "GLACCOUNT"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "StockCode"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        myDataColumn = New DataColumn
        myDataColumn.DataType = System.Type.GetType("System.String")
        myDataColumn.ColumnName = "BothNotes"
        myDataColumn.ReadOnly = True
        myDataTable.Columns.Add(myDataColumn)

        Dim dsLines As New DataSet
        dsLines.Tables.Add(myDataTable)

        Dim myRow As DataRow
        For Each myRow In _dsOrderLines.Tables("OrderLines").Rows
            Dim newRow As DataRow = dsLines.Tables(0).NewRow
            newRow("DELNUM") = "01"
            If Trim(myRow("LINNUM")).Length < 2 Then
                newRow("LINDEL") = "0" & Trim(myRow("LINNUM")) & "-" & "01"
                newRow("LINNUM") = "0" & Trim(myRow("LINNUM"))
            Else
                newRow("LINDEL") = Trim(myRow("LINNUM")) & "-" & "01"
                newRow("LINNUM") = Trim(myRow("LINNUM"))
            End If

            newRow("CURQTY") = myRow("QTY_ORD")
            newRow("DUESHIP") = FormatNumber(myRow("QTY_ORD"), 0) & vbCrLf & FormatNumber("0", 0)

            Dim dueDate As Date = _dsOrderHeader.Tables("OrderHeader").Rows(0).Item("PO_DATE")
            Select Case dueDate.DayOfWeek
                Case DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday
                    dueDate = dueDate.AddDays(3)
                Case DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday
                    dueDate = dueDate.AddDays(5)
                Case DayOfWeek.Saturday
                    dueDate = dueDate.AddDays(4)
            End Select

            newRow("CURDUE") = dueDate
            newRow("CUSDUE") = dueDate

            Dim partNum As String
            If IsDBNull(myRow("PARTNUM")) Then
                partNum = ""
            Else
                partNum = myRow("PARTNUM")
            End If

            If partNum <> "" Then
                Dim myPart As New PartClass(partNum, DataSource.Max)
                Dim _partChar As New PartCharacteristicsClass(partNum)

                newRow("PARTEXT") = partNum
                newRow("PRTNUM") = partNum
                newRow("BOTHDESC") = myPart.Description.Trim & vbCrLf & myPart.Description2.Trim
                Dim _generalLedgerAcctClass As New GeneralLedgerAcctClass
                newRow("GLACCOUNT") = _generalLedgerAcctClass.Read(myPart.ACTTYP)
                Dim _partClass As New PartClass(myPart.PartNumber, ClassBase.DataSource.Max)
                newRow("StockCode") = _partClass.StockCode
                newRow("UOM") = myPart.BOMUOM
                newRow("Pallets") = "No"
                newRow("UOMPALLETS") = myPart.BOMUOM & vbCrLf & "No"
            Else
                newRow("PARTEXT") = ""
                newRow("PRTNUM") = ""
                newRow("BOTHDESC") = ""
                newRow("BothNotes") = ""
                newRow("UOM") = ""
                newRow("Pallets") = "No"
                newRow("UOMPALLETS") = "" & vbCrLf & "No"
            End If

            Dim clsNumber As New NumberClass
            ' Figure out how many significant decimal digits there are in the unit price,
            ' and if less than 2 always show at least two decimals
            Dim iDigits As Integer = clsNumber.SignificantDigits(myRow("PRICE"))
            If iDigits < 2 Then
                iDigits = 2
            End If
            newRow("PrtUnitPrice") = FormatCurrency(myRow("PRICE"), iDigits)

            newRow("EXTPRICES") = FormatCurrency(myRow("PRICE") * myRow("QTY_ORD"), 2) & vbCrLf & "$0.00"
            _TotalPrice = _TotalPrice + myRow("PRICE") * myRow("QTY_ORD")

            newRow("REQSHP") = dueDate
            newRow("ShortComment") = "Line Notes"
            newRow("Closed") = False

            newRow("Taxable") = _customer.Taxable
            newRow("TaxCode1") = _TaxCode1
            newRow("TaxCode2") = _TaxCode2
            newRow("TaxCode3") = _TaxCode3
            newRow("TaxRate1") = _TaxRate1
            newRow("TaxRate2") = _TaxRate2
            newRow("TaxRate3") = _TaxRate3

            dsLines.Tables(0).Rows.Add(newRow)
        Next

        _OrderLines = dsLines

        'Comment 1
        _Comment1 = ""

        'Comment 2
        _Comment2 = ""

        'Comment 3
        _Comment3 = ""

        'Ext Notes
        _ExtNotes = ""

    End Sub

    Private Function GetState(ByVal _state As String) As String
        Dim ds As DataSet
        Dim oSQL As New SqlService(ConnectionString)
        Dim strSQL As String = "SELECT ABBREVIATION FROM STATES WHERE (STATES.NAME = '" & UCase(_state) & "')"

        ds = oSQL.RunSql(strSQL, "States")

        Return UCase(ds.Tables("States").Rows(0).Item("Abbreviation"))

    End Function

End Class

Public Class EDIAutoZone850Class
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
    Private _LocationCodeQualifier As String
    Private _AddressLocationNumber As String

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

    Public ReadOnly Property AddressLocationNumber() As String
        Get
            Return _AddressLocationNumber.Trim
        End Get
    End Property

    Public ReadOnly Property LocationCodeQualifier() As String
        Get
            Return _LocationCodeQualifier.Trim
        End Get
    End Property

    Public Sub New(ByVal _dsOrderHeader As DataSet, ByVal _dsOrderLines As DataSet, ByVal _dsCustomerXref As DataSet)

        'PO Number
        _PONumber = _dsOrderHeader.Tables(0).Rows(0).Item("CUST_PO")

        ' Location Code Qualifier
        _LocationCodeQualifier = _dsOrderHeader.Tables(0).Rows(0).Item("LOCATION_CODE_QUALIFIER")

        ' Address Location Number
        _AddressLocationNumber = _dsOrderHeader.Tables(0).Rows(0).Item("ADDR_LOC_CODE")

        'Terms
        ' Use Default Term Code
        _termsCode = _dsCustomerXref.Tables(0).Rows(0).Item("TERMS")

        'ShipVia
        _shipViaCode = _dsCustomerXref.Tables(0).Rows(0).Item("SHPVIA")

        'Ordered By
        _OrderedBy = "EDI"

        'FOB
        _FOB = "DESTINATION"

        'Customer
        Dim _custid As String = _dsCustomerXref.Tables(0).Rows(0).Item("CUST_ID")
        _customer = New CustomerClass(_custid, DataSource.Max)

        'Tax Codes
        Dim myTax As New TaxMasterClass

        'Ship To Address
        If IsDBNull(_dsOrderHeader.Tables(0).Rows(0).Item("ADDR_LOC_CODE")) Then
            _ShipToCode = ""
            _ShipToName = _customer.Name
            _ShipToAddress1 = _customer.Address1
            _ShipToAddress2 = _customer.Address2
            _ShipToCity = _customer.City
            _ShipToState = _customer.State
            _ShipToZip = _customer.ZipCode
            _ShipToCountry = _customer.Country

            'Tax Information
            _TaxCode1 = _customer.TaxCode1
            _TaxCode2 = _customer.TaxCode2
            _TaxCode3 = _customer.TaxCode3
            _TaxRate1 = myTax.Read(_customer.TaxCode1)
            _TaxRate2 = myTax.Read(_customer.TaxCode2)
            _TaxRate3 = myTax.Read(_customer.TaxCode3)

        Else
            Dim _custshipcode As String = _dsOrderHeader.Tables(0).Rows(0).Item("ADDR_LOC_CODE")
            Dim _shippingMaster As New ShippingMasterClass(_custid, _custshipcode, DataSource.Max)
            _ShipToCode = _custshipcode
            _ShipToName = _shippingMaster.Name
            _ShipToAddress1 = _shippingMaster.Address1
            _ShipToAddress2 = _shippingMaster.Address2
            _ShipToCity = _shippingMaster.City
            _ShipToState = _shippingMaster.State
            _ShipToZip = _shippingMaster.ZipCode
            _ShipToCountry = _shippingMaster.Country

            'Tax Information
            _TaxCode1 = _shippingMaster.TaxCode1
            _TaxCode2 = _shippingMaster.TaxCode2
            _TaxCode3 = _shippingMaster.TaxCode3
            _TaxRate1 = myTax.Read(_shippingMaster.TaxCode1)
            _TaxRate2 = myTax.Read(_shippingMaster.TaxCode2)
            _TaxRate3 = myTax.Read(_shippingMaster.TaxCode3)

        End If

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

            Dim myDate As Date = _dsOrderHeader.Tables("OrderHeader").Rows(0).Item("DATE_2")
            'Create new date with year 2030
            Dim newDate As Date = New Date(2030, myDate.Month, myDate.Day)

            newRow("CURDUE") = newDate
            newRow("CUSDUE") = newDate

            Dim partNum As String
            If IsDBNull(myRow("VENDOR_PARTNUM")) Then
                partNum = ""
            Else
                partNum = myRow("VENDOR_PARTNUM")
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

                'Calculate Carton Count
                Dim cartonCount As Integer = Math.Round(myRow("QTY_ORD") * myPart.SLSCNV)

                newRow("BothNotes") = cartonCount & " CS"
            Else
                newRow("PARTEXT") = ""
                newRow("PRTNUM") = ""
                newRow("BOTHDESC") = ""
                newRow("BothNotes") = ""
            End If

            newRow("UOM") = myRow("UOM")
            newRow("UOMPALLETS") = myRow("UOM") & vbCrLf & "Yes"

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

            newRow("REQSHP") = Format(newDate, "MM/dd/yy")
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
        _Comment1 = _dsCustomerXref.Tables(0).Rows(0).Item("COMNT1")

        'Comment 2
        _Comment2 = _dsCustomerXref.Tables(0).Rows(0).Item("COMNT2")

        'Comment 3
        _Comment3 = ""

        'Ext Notes
        _ExtNotes = _dsCustomerXref.Tables(0).Rows(0).Item("NOTES")
        _ExtNotes = _ExtNotes.Replace("[LINE3]", "* SHIP NO LATER THAN " & _dsOrderHeader.Tables("OrderHeader").Rows(0).Item("DATE_2") & " TO ARRIVE BY " & _dsOrderHeader.Tables("OrderHeader").Rows(0).Item("DATE_1") & ".")

    End Sub

End Class

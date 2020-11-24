Public Class PartData
    ' Data from the PartMaster and PartMasterCharacteristics tables retrieved via the GetPartData function
    Public Property OrderNumber As String = ""
    Public Property AccountType As String = ""
    Public Property Description As String = ""
    Public Property PartNumber As String = ""
    Public Property StockCode As String = ""
    Public Property SalesUom As String = ""
    Public Property SalesConversion As Decimal = 0
    Public Property CasesPerPallet As Integer = 0
    Public Property CasesPerPalletLayer As Integer = 0
    Public Property PalletsPerTruck As Integer = 0
    Public Property EachesPerMfgUnit As Integer = 0
    Public Property Pallets As Integer = 0
    Public Property CustomerPartNumber As String = ""
    Public Property LineNumber As String = ""
    Public Property POPrice As Decimal = 0.0D
    Public Property POTax As Decimal = 0.0D
    Public Property OriginalPrice As Decimal = 0.0D
    Public Property Discount As Decimal = 0.0D
    Public Property DiscountedPrice As Decimal = 0.0D
    Public Property OrderQuantity As Integer = 0
    Public Property CaseQuantity As Integer = 0
    Public Property RemainingLayers As Integer = 0
    Public Property RemainingCases As Integer = 0
    Public Property GTIN As String = ""
    Public Property ConsumerPackageCode As String = ""
    Public Property Weight As Decimal = 0.0D
    Public Property CaseLength As Decimal = 0.0D
    Public Property CaseHeight As Decimal = 0.0D
    Public Property CaseWidth As Decimal = 0.0D
    Public Property ORDPOL As String = ""
    Public Property PurchaseOrderLineNumber As String = ""
    Public Property GenericCustomerField As String = ""
    Public Property PalletGroupCode As String = ""
End Class

Imports Sellars.SQL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math

Public Class PriceQuoteClass
    Inherits ClassBase

    ' declare other business objects, make public so all properties are available
    ' to the calling application as:   PriceQuoteClass.CustomerSettings.GIP, etc...
    Public Customer As CustomerClass
    Public CustomerSettings As CustomerSettingsClass
    Public CategorySettings As CategorySettingsClass
    Public Part As PartClass
    Public FreightSettings As FreightSettingsClass
    Public ShippingInfo As ShippingMasterClass

    Private MaxIterations As Integer = 51

    Private _CommissionPercent As Decimal = 0

    Private _SalesRep As String
    Private _Quantity As Decimal = 0

    Private _EstCommission As Decimal = 0
    Private _EstGIP As Decimal = 0
    Private _EstPartCost As Decimal = 0
    Private _EstRebate As Decimal = 0
    Private _EstRebatePercent As Decimal = 0
    Private _EstNascar As Decimal = 0
    Private _EstFreight As Decimal = 0

    Private _Price As Decimal = 0
    Private _EstRevenue As Decimal = 0
    Private _EstCost As Decimal = 0
    Private _EstMargin As Decimal = 0

    Private _GIPAmount As Decimal = 0
    Private _CommissionAmount As Decimal = 0
    Private _RebateAmount As Decimal = 0
    Private _RebatePercentAmount As Decimal = 0
    Private _FreightAmount As Decimal = 0

    Private _NascarAmount As Decimal = 0

    Private PartCharacteristics As PartCharacteristicsClass

    Public ReadOnly Property ActualIterations() As Decimal
        Get
            Return 51 - MaxIterations
        End Get
    End Property

    Public ReadOnly Property CommissionPercent() As Decimal
        Get
            Return _CommissionPercent
        End Get
    End Property

    Public ReadOnly Property EstimatedCommission() As Decimal
        Get
            Return Round(_EstCommission, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedCost() As Decimal
        Get
            Return Round(_EstCost, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedFreight() As Decimal
        Get
            Return Round(_EstFreight, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedGIP() As Decimal
        Get
            Return Round(_EstGIP, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedMargin() As Decimal
        Get
            Return _EstMargin
        End Get
    End Property

    Public ReadOnly Property EstimatedNascar() As Decimal
        Get
            Return Round(_EstNascar, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedPrice() As Decimal
        Get
            Return Round(_Price, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedPartCost() As Decimal
        Get
            Return Round(_EstPartCost, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedRebate() As Decimal
        Get
            Return Round(_EstRebate, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedRebatePercentage() As Decimal
        Get
            Return Round(_EstRebatePercent, 2)
        End Get
    End Property

    Public ReadOnly Property EstimatedRevenue() As Decimal
        Get
            Return Round(_EstRevenue, 2)
        End Get
    End Property

    Public ReadOnly Property SalesRep() As String
        Get
            Return _SalesRep
        End Get
    End Property

    Public ReadOnly Property EstimatedWeight() As Decimal
        Get
            Dim PackageWeight As Decimal
            If Part.BOMUOM = "LB" Then
                PackageWeight = Math.Round(CDec(1 * _Quantity), 0)
            Else
                PackageWeight = Math.Round(PartCharacteristics.FGBasisWeight * _Quantity, 0)
            End If
            Return PackageWeight
        End Get
    End Property

    Public Sub New(ByVal CustomerID As String, ByVal SalesRep As String, ByRef CustomerObj As CustomerClass)
        ' Get the customer settings when we initialize
        ' Set the customer object so we can reference the variables
        ' in the parent class object
        Customer = CustomerObj

        ' Create a new customer settings object
        CustomerSettings = New CustomerSettingsClass(CustomerID)

        ' Set the sales rep id
        _SalesRep = SalesRep
    End Sub

    Public Function Calculate(ByVal PartNumber As String, ByVal Quantity As Decimal, ByVal ShipToCode As String) As Decimal
        ' Capture the quantity and set to a property
        _Quantity = Quantity

        ShippingInfo = New ShippingMasterClass(ShipToCode, ClassBase.DataSource.Sellars)
        ShippingInfo.Read(Customer.CustomerNumber, ShipToCode, ClassBase.DataSource.Sellars)

        ' Get the state specific freight settings which contains
        ' the mulitplier to use based on the state being shipped to
        FreightSettings = New FreightSettingsClass(ShippingInfo.State)

        ' First of all get the part information
        ' we really need the planid and revlev in order to retrieve the category settings
        Part = New PartClass(PartNumber)
        PartCharacteristics = New PartCharacteristicsClass(PartNumber)

        ' Get the target margin and nascar discounts from the part settings
        CategorySettings = New CategorySettingsClass(Part.PlanCode, Part.REVLEV)
        ' If the PartCharacteristics NASCAR flag is 0, then zero out the NASCAR
        ' percentage on the CategorySettings
        If PartCharacteristics.Nascar = False Then
            CategorySettings.NascarDiscount = 0
        End If

        ' Only try to do the price calculation if there is a cost, and a target margin
        If Part.Cost > 0 And CategorySettings.TargetMargin > 0 Then
            CalculatePrice()
            Return _Price
        Else
            Return 0
        End If
    End Function

    Private Sub CalculatePrice()
        Dim Commission As Decimal
        Dim TestMargin As Decimal
        Dim SeedPrice As Decimal
        Dim SimpleRevenue As Decimal
        Dim Increment As Decimal
        Dim GoalMargin As Decimal
        Dim TempCommission As Decimal

        Increment = 50
        MaxIterations = 51
        GoalMargin = CategorySettings.TargetMargin

        ' Set the base testing price
        SimpleRevenue = Part.Cost * CDec(_Quantity)

        ' We will only use commission if the customer record
        ' has a sales rep assigned to it.
        If SalesRep = "" Then
            _CommissionPercent = 0
        Else
            _CommissionPercent = CategorySettings.Commission
        End If
        Commission = SimpleRevenue * CommissionPercent

        ' Max uses the WGT field in the partmaster to store other
        ' information for DRC Roll Goods, so we need to have this
        ' little buypass in for LB items until Max gets cleaned up
        If Part.BOMUOM = "LB" Then
            _EstFreight = (CategorySettings.FreightPerUnit * FreightSettings.FreightMultiplier)
        Else
            _EstFreight = (CategorySettings.FreightPerUnit * Part.PoundsPerSalesUnit * FreightSettings.FreightMultiplier)
        End If

        SeedPrice = Part.Cost + (CDec(_EstFreight) / CDec(_Quantity)) + ((1 + GoalMargin) * (Commission + CustomerSettings.Rebate))

        While (Increment > 0.002 Or Increment < -0.002) And MaxIterations > 0
            TestMargin = CalcMargin(SeedPrice)
            If (TestMargin < GoalMargin And Increment < 0) Or (TestMargin > GoalMargin And Increment > 0) Then
                Increment = Increment * -0.5
            End If
            SeedPrice = SeedPrice + Increment
            MaxIterations = MaxIterations - 1
        End While

        _Price = SeedPrice
    End Sub

    Private Function CalcMargin(ByVal TestPrice As Decimal) As Decimal

        ' Max uses the WGT field in the partmaster to store other
        ' information for DRC Roll Goods, so we need to have this
        ' little buypass in for LB items until Max gets cleaned up
        If Part.BOMUOM = "LB" Then
            _EstFreight = (CategorySettings.FreightPerUnit * _Quantity) * FreightSettings.FreightMultiplier
        Else
            _EstFreight = (CategorySettings.FreightPerUnit * (_Quantity * Part.PoundsPerSalesUnit)) * FreightSettings.FreightMultiplier
        End If

        _EstRevenue = System.Math.Round(TestPrice * _Quantity, 2)
        _EstPartCost = (Part.Cost * _Quantity)
        _EstCommission = (CommissionPercent * _EstRevenue)
        _EstGIP = (CustomerSettings.GIP * _EstRevenue)
        _EstRebate = (CustomerSettings.Rebate * _Quantity)
        _EstRebatePercent = (CustomerSettings.RebatePercent * _EstRevenue)

        ' Only allow NASCAR pricing if the part characteristics NASCAR flag is true
        If PartCharacteristics.Nascar Then
            _EstNascar = (CategorySettings.NascarDiscount * _EstRevenue)
        Else
            _EstNascar = 0
        End If

        _EstCost = _EstPartCost + _EstFreight + _EstCommission + _EstGIP + _EstRebate + _EstRebatePercent + _EstNascar

        _EstMargin = (_EstRevenue - _EstCost) / _EstRevenue

        Return _EstMargin
    End Function

    Public Function Add(ByVal PartNumber As String, ByVal Quantity As Decimal, ByVal ShipToCode As String, ByVal Remarks As String, ByVal UserName As String, ByVal UserSalesRep As String, ByVal FullName As String) As Date

        ' Calculate all the price settings
        Calculate(PartNumber, Quantity, ShipToCode)

        ' Declare the SQL data layer class
        Dim oSQL As New SqlService(ConnectionString)

        ' Add the parameters to the command object
        oSQL.AddParameter("@UserName", SqlDbType.NVarChar, 20, UserName, ParameterDirection.Input)
        oSQL.AddParameter("@FullName", SqlDbType.NVarChar, 50, FullName, ParameterDirection.Input)
        oSQL.AddParameter("@UserSalesRep", SqlDbType.NVarChar, 7, UserSalesRep, ParameterDirection.Input)
        oSQL.AddParameter("@SalesRep", SqlDbType.NVarChar, 7, SalesRep, ParameterDirection.Input)
        oSQL.AddParameter("@CUSTID", SqlDbType.NVarChar, 20, Customer.CustomerNumber, ParameterDirection.Input)
        oSQL.AddParameter("@CustomerName", SqlDbType.NVarChar, 30, Customer.Name, ParameterDirection.Input)
        oSQL.AddParameter("@SalesTerms", SqlDbType.NVarChar, 20, Customer.Terms.Description, ParameterDirection.Input)
        oSQL.AddParameter("@PRTNUM", SqlDbType.NVarChar, 30, Part.PartNumber, ParameterDirection.Input)
        oSQL.AddParameter("@PartDescription", SqlDbType.NVarChar, 25, Part.Description, ParameterDirection.Input)
        oSQL.AddParameter("@OnHand", SqlDbType.Int, 0, Part.OnHand, ParameterDirection.Input)
        oSQL.AddParameter("@Committed", SqlDbType.Int, 0, Part.Committed, ParameterDirection.Input)
        oSQL.AddParameter("@Available", SqlDbType.Int, 0, Part.Available, ParameterDirection.Input)
        oSQL.AddParameter("@ShipToName", SqlDbType.NVarChar, 30, ShippingInfo.Name, ParameterDirection.Input)
        oSQL.AddParameter("@ShipToAddress1", SqlDbType.NVarChar, 30, ShippingInfo.Address1, ParameterDirection.Input)
        oSQL.AddParameter("@ShipToAddress2", SqlDbType.NVarChar, 30, ShippingInfo.Address2, ParameterDirection.Input)
        oSQL.AddParameter("@ShipToCity", SqlDbType.NVarChar, 15, ShippingInfo.City, ParameterDirection.Input)
        oSQL.AddParameter("@ShipToState", SqlDbType.NVarChar, 30, ShippingInfo.State, ParameterDirection.Input)
        oSQL.AddParameter("@ShipToZip", SqlDbType.NVarChar, 30, ShippingInfo.ZipCode, ParameterDirection.Input)
        oSQL.AddParameter("@FormattedShipToAddress", SqlDbType.NVarChar, 1000, ShippingInfo.MailingAddress, ParameterDirection.Input)
        oSQL.AddParameter("@Quantity", SqlDbType.Int, 0, Me._Quantity, ParameterDirection.Input)
        oSQL.AddParameter("@UnitPrice", SqlDbType.Money, 0, EstimatedPrice, ParameterDirection.Input)
        oSQL.AddParameter("@TotalPrice", SqlDbType.Money, 0, EstimatedRevenue, ParameterDirection.Input)
        oSQL.AddParameter("@FreightBasis", SqlDbType.Money, 0, CategorySettings.FreightPerUnit, ParameterDirection.Input)
        oSQL.AddParameter("@FreightMultiplier", SqlDbType.Float, 0, FreightSettings.FreightMultiplier, ParameterDirection.Input)
        oSQL.AddParameter("@FreightTotal", SqlDbType.Money, 0, EstimatedFreight, ParameterDirection.Input)
        oSQL.AddParameter("@UnitWeight", SqlDbType.Float, 0, PartCharacteristics.FGBasisWeight, ParameterDirection.Input)
        oSQL.AddParameter("@TotalWeight", SqlDbType.Float, 0, EstimatedWeight, ParameterDirection.Input)
        oSQL.AddParameter("@GIPBasis", SqlDbType.Float, 0, CustomerSettings.GIP, ParameterDirection.Input)
        oSQL.AddParameter("@GIPTotal", SqlDbType.Money, 0, EstimatedGIP, ParameterDirection.Input)
        oSQL.AddParameter("@CommissionBasis", SqlDbType.Float, 0, CategorySettings.Commission, ParameterDirection.Input)
        oSQL.AddParameter("@CommissionTotal", SqlDbType.Money, 0, EstimatedCommission, ParameterDirection.Input)
        oSQL.AddParameter("@Rebate$Basis", SqlDbType.Float, 0, CustomerSettings.Rebate, ParameterDirection.Input)
        oSQL.AddParameter("@Rebate$Total", SqlDbType.Money, 0, EstimatedRebate, ParameterDirection.Input)
        oSQL.AddParameter("@RebatePcntBasis", SqlDbType.Float, 0, CustomerSettings.RebatePercent, ParameterDirection.Input)
        oSQL.AddParameter("@RebatePcntTotal", SqlDbType.Money, 0, EstimatedRebatePercentage, ParameterDirection.Input)
        oSQL.AddParameter("@NascarBasis", SqlDbType.Float, 0, CategorySettings.NascarDiscount, ParameterDirection.Input)
        oSQL.AddParameter("@NascarTotal", SqlDbType.Money, 0, EstimatedNascar, ParameterDirection.Input)
        oSQL.AddParameter("@UnitCost", SqlDbType.Money, 0, Part.Cost, ParameterDirection.Input)
        oSQL.AddParameter("@TotalCost", SqlDbType.Money, 0, EstimatedCost, ParameterDirection.Input)
        oSQL.AddParameter("@TargetMargin", SqlDbType.Float, 0, CategorySettings.TargetMargin, ParameterDirection.Input)
        oSQL.AddParameter("@QuoteMargin", SqlDbType.Float, 0, EstimatedMargin, ParameterDirection.Input)
        oSQL.AddParameter("@Remarks", SqlDbType.NVarChar, 1000, Remarks, ParameterDirection.Input)

        ' Add the output parameters to the command object
        Dim ID = New Parameter("@ID", SqlDbType.DateTime, 0, CDate("01/01/2001 12:01:00 AM"), ParameterDirection.Output)
        oSQL.AddParameter(ID)

        ' Execute the stored procedure
        oSQL.RunProc("AddPriceQuote")

        ' Return the new price quote ID
        Return ID.Value
    End Function

End Class

Public Class DTCCTransaction

    ' Public properties to hold the data from the columns in the provided image.
    Public Property ExecBroker As String
    Public Property AssetClassSecCode As String
    Public Property Security As String
    Public Property TranCode As String
    Public Property Price As Decimal
    Public Property AcctID As String
    Public Property AllocQty As Decimal
    Public Property TrdAmt As Decimal

    Public Property ChargesTaxesFeesAmount1 As Decimal
    Public Property ChargesTaxesFeesAmount2 As Decimal
    Public Property ChargesTaxesFeesAmount3 As Decimal

    Public Property NetCashAmount As Decimal
    Public Property SettleAmount As Decimal

    Public Property AllocSettleCurr As String

    Public Property Comm As Decimal

    Public Property TradeDate As Date


    ' A constructor to initialize all properties when a new object is created.
    Public Sub New(ByVal execBroker As String,
                   ByVal assetClassSecCode As String,
                   ByVal security As String,
                   ByVal tranCode As String,
                   ByVal price As Decimal,
                   ByVal acctID As String,
                   ByVal allocQty As Decimal,
                   ByVal trdAmt As Decimal,
                   ByVal chargesTaxesFeesAmount1 As Decimal,
    ByVal chargesTaxesFeesAmount2 As Decimal,
    ByVal chargesTaxesFeesAmount3 As Decimal,
    ByVal netCashAmount As Decimal,
    ByVal settleAmount As Decimal,
    ByVal allocSettleCurr As String,
    ByVal comm As Decimal,
                   ByVal tradeDate As Date
           )

        ' Assign the values passed to the constructor to the properties.
        Me.ExecBroker = execBroker
        Me.AssetClassSecCode = assetClassSecCode
        Me.Security = security
        Me.TranCode = tranCode
        Me.Price = price
        Me.AcctID = acctID
        Me.AllocQty = allocQty
        Me.TrdAmt = trdAmt
        Me.ChargesTaxesFeesAmount1 = chargesTaxesFeesAmount1
        Me.ChargesTaxesFeesAmount2 = chargesTaxesFeesAmount2
        Me.ChargesTaxesFeesAmount3 = chargesTaxesFeesAmount3
        Me.NetCashAmount = netCashAmount
        Me.SettleAmount = settleAmount
        Me.AllocSettleCurr = allocSettleCurr
        Me.Comm = comm
        Me.TradeDate = tradeDate


    End Sub

End Class
Public Class Transaction

    ' Public properties to hold the data from your columns.
    Public Property TranCode As String
    Public Property Quantity As String
    Public Property Description As String
    Public Property Price As Decimal
    Public Property NetAmount As Decimal
    Public Property BrokerId As String

    Public Property SecurityCurrency As String

    Public Property TradeDate As Date

    Public Property SettleDate As Date

    Public Property NetFees As Decimal
    Public Property BrokerName As String
    Public Property Cusip As String
    Public Property Commission As Decimal

    Public Property OtherFees As Decimal

    Public Property PortId As String

    Public Property Isin As String

    Public Property Fund As String

    Public Property Sedol As String

    Public Property ReflowFlag As String
    Public Property LotSelectionMethod As String

    ' A constructor to initialize all properties when a new object is created.
    Public Sub New(ByVal tranCode As String,
                   ByVal quantity As String,
                   ByVal descriptionParam As String,
                   ByVal priceParam As Decimal,
                   ByVal netAmountParam As Decimal,
                   ByVal brokerId As String,
                   ByVal securityCurrency As String,
                   ByVal tradeDate As Date,
                   ByVal settleDate As Date,
                   ByVal netFees As Decimal,
                    ByVal brokerName As String,
                    ByVal cusip As String,
                    ByVal commission As Decimal,
                    ByVal otherFees As Decimal,
                    ByVal portId As String,
                    ByVal isin As String,
                   ByVal fund As String,
                   ByVal sedol As String,
                      ByVal reflowFlag As String,
                         ByVal lotSelectionMethod As String
        )

        ' Assign the values passed to the constructor to the properties.
        Me.TranCode = tranCode
        Me.Quantity = quantity
        Me.Description = descriptionParam
        Me.Price = priceParam
        Me.NetAmount = netAmountParam
        Me.BrokerId = brokerId
        Me.SecurityCurrency = securityCurrency
        Me.TradeDate = tradeDate
        Me.SettleDate = settleDate
        Me.NetFees = netFees
        Me.BrokerName = brokerName
        Me.Cusip = cusip
        Me.Commission = commission
        Me.OtherFees = otherFees
        Me.PortId = portId
        Me.Isin = isin
        Me.Fund = fund
        Me.Sedol = sedol
        Me.ReflowFlag = reflowFlag
        Me.LotSelectionMethod = lotSelectionMethod
    End Sub

End Class


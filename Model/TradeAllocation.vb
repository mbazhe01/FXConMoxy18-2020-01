
Public Class TradeAllocation

    ' Public properties to hold the data from the columns.
    Public Property ExecBroker As String
    Public Property AssetClass As String
    Public Property SecCode As String
    Public Property Security As String
    Public Property BS As String ' Renamed to match sanitized header
    Public Property Price As Decimal
    Public Property AcctID As String
    Public Property AllocQtyType As String
    Public Property QtyAlloc As Decimal
    Public Property TrdAmt As Decimal
    Public Property ChargesTaxesFeesType1 As String
    Public Property ChargesTaxesFeesAmount1 As Decimal
    Public Property ChargesTaxesFeesType2 As String
    Public Property ChargesTaxesFeesAmount2 As Decimal
    Public Property ChargesTaxesFeesType3 As String
    Public Property ChargesTaxesFeesAmount3 As Decimal
    Public Property ChargesTaxesFeesType4 As String
    Public Property ChargesTaxesFeesAmount4 As Decimal
    Public Property ChargesTaxesFeesType5 As String
    Public Property ChargesTaxesFeesAmount5 As Decimal
    Public Property Fees As Decimal
    Public Property AccruedInterest As Decimal
    Public Property NetCashAmount As Decimal
    Public Property SettleAmt As Decimal
    Public Property AllocSettleCurr As String
    Public Property SettleCond As String
    Public Property PSET As String
    Public Property ClientAllocRef As String
    Public Property AllocPartyCapacity As String
    Public Property CommType1 As String
    Public Property CommBasis1 As String
    Public Property CommAmount1 As Decimal
    Public Property CommReason1 As String
    Public Property CommType2 As String
    Public Property CommBasis2 As String
    Public Property CommAmount2 As Decimal
    Public Property CommReason2 As String
    Public Property CommType3 As String
    Public Property CommBasis3 As String
    Public Property CommAmount3 As Decimal
    Public Property CommReason3 As String
    Public Property Comm As Decimal
    Public Property AllocErr As String
    Public Property TradeDate As Date
    Public Property SettleDate As Date

    ' NOTE: The constructor is no longer needed since we are using reflection to set properties.
    ' This makes the class simpler and more flexible.
End Class
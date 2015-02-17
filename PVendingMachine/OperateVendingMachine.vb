Public Class OperateVendingMachine
  Private prvVendingMachine As CVendingMachine

  Public Sub New()
    prvVendingMachine = New CVendingMachine()
  End Sub


  Public Function DisplayMessage() As String
    DisplayMessage = prvVendingMachine.DisplayMessage
  End Function


  Public Function CurrentAmount() As Decimal
    CurrentAmount = prvVendingMachine.CurrentAmount
  End Function


  Public Function LastReturnedAmount() As Decimal
    LastReturnedAmount = prvVendingMachine.LastReturnedAmount
  End Function


  Public Function InsertCoin(ByVal Coin) As Boolean

    Dim bFoundCoin As Boolean
    Dim bIsCoinAccepted As Boolean

    Dim dblSizeInMillimeters As Double
    Dim dblWeightInGrams As Double

    Dim sCoinName As String

    Dim aCoin(2) As Object


    dblSizeInMillimeters = 0
    dblWeightInGrams = 0

    'Compare the provided coin name to the set of defined coins
    For iCoin = 0 To CoinNameSizeWeight.GetUpperBound(0)
      aCoin = CoinNameSizeWeight(iCoin)
      sCoinName = aCoin(0)

      bFoundCoin = (Coin = sCoinName)
      If (bFoundCoin) Then
        'This coin name exists.  Use its parameters
        dblSizeInMillimeters = aCoin(1)
        dblWeightInGrams = aCoin(2)
        Exit For
      End If
    Next iCoin

    If (bFoundCoin) Then
      With prvVendingMachine
        bIsCoinAccepted = .AcceptCoin(dblSizeInMillimeters, dblWeightInGrams)
      End With
    End If

    InsertCoin = bIsCoinAccepted
  End Function


  Public Function SelectProduct(Product As String) As Boolean

    Dim bFoundProduct As Boolean
    Dim bIsProductDispensed As Boolean

    Dim decProductPrice As Decimal

    Dim sProductName As String

    Dim aProduct(1) As Object


    decProductPrice = 0

    'Compare the provided Product name to the set of defined Products
    For iProduct = 0 To ProductNamePrice.GetUpperBound(0)
      aProduct = ProductNamePrice(iProduct)
      sProductName = aProduct(0)

      bFoundProduct = (Product = sProductName)
      If (bFoundProduct) Then
        'This Product name is offered by this vending machine
        Exit For
      End If
    Next iProduct

    If (bFoundProduct) Then
      With prvVendingMachine
        bIsProductDispensed = .DispenseProduct(Product)
      End With
    End If

    SelectProduct = bIsProductDispensed
  End Function


  Public Function ReturnCurrentAmount() As Boolean

    Dim bCoinsReturned As Boolean


    With prvVendingMachine
      ReturnCurrentAmount = .ReturnCurrentAmount
    End With

    bCoinsReturned = bCoinsReturned

  End Function

End Class

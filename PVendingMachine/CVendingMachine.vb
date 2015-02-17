Public Class CVendingMachine

  'Stock the vending machine.  There is one item of each product in stock
  Private prvProductCandy As New CVendingMachineProduct("Candy", 0.65, 1)
  Private prvProductChips As New CVendingMachineProduct("Chips", 0.5, 1)
  Private prvProductCola As New CVendingMachineProduct("Cola", 1.0, 1)

  'Initialize the coin changer.  There are ten of each coin in the changer
  Private prvCoinChanger As New CCoinChanger(10, 10, 10)


  Public Sub New()
    'Initialize the vending machine
    ResetDisplayMessage()
  End Sub


  'The default display message depends upon the coin content of the coin changer
  Public Sub ResetDisplayMessage()

    Dim bCanReturnAllPrices As Boolean
    With prvCoinChanger
      'NOTE: With a full implementation of CanReturnAmount, this section could be significantly
      'improved to only indicate EXACT CHANGE ONLY when the specific product that was selected
      'cannot be purchased with the current amount with the precise coins that are in the machine
      'This implementation allows the conditions to be produced for test
      bCanReturnAllPrices = (.CanReturnAmount(prvProductCandy.Price))
      bCanReturnAllPrices = (bCanReturnAllPrices And .CanReturnAmount(prvProductChips.Price))
      bCanReturnAllPrices = (bCanReturnAllPrices And .CanReturnAmount(prvProductCola.Price))
    End With

    If (bCanReturnAllPrices) Then
      prvDisplayMessage = "INSERT COINS"
    Else
      prvDisplayMessage = "EXACT CHANGE ONLY"
    End If
  End Sub


  'The current value of coins accepted by the vending machine and not applied to a purchase
  Private prvCurrentAmount As Decimal
  Public ReadOnly Property CurrentAmount() As Decimal
    Get
      Return prvCurrentAmount
    End Get
  End Property


  'The value of the most-recently returned amount issued by the vending machine
  Private prvLastReturnedAmount As Decimal
  Public ReadOnly Property LastReturnedAmount() As Decimal
    Get
      Dim decLastReturnedAmount As Decimal

      decLastReturnedAmount = prvLastReturnedAmount
      'The last returned amount is a one-shot value that represents coins that were
      'returned to the vending machine coin return.  Once checked, the value is set
      'to zero to avoid the suggestion of additional returned amounts
      prvLastReturnedAmount = 0

      Return decLastReturnedAmount
    End Get
  End Property


  'The display text of the vending machine's interface panel
  Private prvDisplayMessage As String
  Public ReadOnly Property DisplayMessage() As String
    Get
      Dim decCurrentAmount As Decimal
      Dim sDisplayMessage As String

      sDisplayMessage = prvDisplayMessage
      decCurrentAmount = prvCurrentAmount

      'Many display messages are temporary status messages related to the operator's prior action.
      'Once the message has been displayed, the next display will revert to the current amount, or
      'to the instruction to insert coins, if the current amount is zero.
      If (prvCurrentAmount = 0) Then
        ResetDisplayMessage()
      Else
        prvDisplayMessage = Format(decCurrentAmount, "$0.00")
      End If

      'Return the value of the display message at the time this method was called
      Return sDisplayMessage
    End Get
  End Property


  'This method accepts a coin and either keeps the coin and credits the current amount with the value
  'of the coin or rejects the coin.  It returns true when a coin was accepted
  Public Function AcceptCoin(SizeInMillimeters As Double, _
                             WeightInGrams As Double _
                             ) As Boolean

    Dim bIsValidCoin As Boolean
    Dim decCoinValue As Decimal

    decCoinValue = IdentifyCoinValueInDollars(SizeInMillimeters, WeightInGrams)
    bIsValidCoin = (decCoinValue > 0)

    If (bIsValidCoin) Then
      prvCurrentAmount = prvCurrentAmount + decCoinValue
      With prvCoinChanger
        Select Case decCoinValue
          Case 0.25
            .AddQuarter()
          Case 0.1
            .AddDime()
          Case 0.05
            .AddNickel()
        End Select
      End With

      prvDisplayMessage = Format(prvCurrentAmount, "$0.00")
    End If

    AcceptCoin = bIsValidCoin

  End Function


  'When the current amount is equal to or exceeds the price of the selected product, this method dispenses
  'a product and returns any portion of the current amount that exceeds the product price to the coin return.
  'When the current amount is insufficient to purchase the product, it updates the display message to
  'indicate the product price.
  'It returns true when a product was dispensed
  Public Function DispenseProduct(ProductName As String) As Boolean

    Dim bIsProductAvailable As Boolean
    Dim bIsProductInStock As Boolean
    Dim bAreFundsAvailable As Boolean
    Dim bCanRemainingAmountBeReturned As Boolean
    Dim bWasProductDispensed As Boolean

    Dim decCurrentAmount As Decimal
    Dim decPrice As Decimal
    Dim decReturnAmount As Decimal

    Dim oProduct As CVendingMachineProduct

    oProduct = Nothing

    Select Case ProductName
      Case "Candy"
        oProduct = prvProductCandy
      Case "Chips"
        oProduct = prvProductChips
      Case "Cola"
        oProduct = prvProductCola
    End Select

    bIsProductAvailable = Not (oProduct Is Nothing)

    If (bIsProductAvailable) Then
      'The product is being sold by the vending machine
      With oProduct
        bIsProductInStock = (.QuantityInStock > 0)

        If (bIsProductInStock) Then
          'Some of this kind of product is present in the vending machine
          decCurrentAmount = prvCurrentAmount
          decPrice = .Price
          bAreFundsAvailable = (decCurrentAmount >= decPrice)

          If (bAreFundsAvailable) Then
            decReturnAmount = decCurrentAmount - decPrice
            With prvCoinChanger
              bCanRemainingAmountBeReturned = .CanReturnAmount(decReturnAmount)
            End With

            If (bCanRemainingAmountBeReturned) Then
              'Enough coins have been inserted to purchase the product and the change can be returned.
              'Dispense the product.
              'TODO: This is the point where the command to physically dispense the product would occur
              .SetQuantityInStock(.QuantityInStock - 1)
              bWasProductDispensed = True
              prvDisplayMessage = "THANK YOU"
              If (decReturnAmount > 0) Then
                With prvCoinChanger
                  .ReturnAmount(decReturnAmount)
                  prvLastReturnedAmount = decReturnAmount
                End With
              End If
              'ASSUMPTION:  All coin returns succeed, so the current amount is always zero at this point
              prvCurrentAmount = 0
            Else
              prvDisplayMessage = "EXACT CHANGE ONLY"
            End If
          Else
            prvDisplayMessage = "PRICE " & Format(.Price, "$0.00")
          End If
        Else
          prvDisplayMessage = "SOLD OUT"
        End If
      End With
    Else
      'If the product is not available, an error in the integration test product name exists
      'No changes to the vending machine display are specified for this condition.  Exit silently
    End If

    DispenseProduct = bWasProductDispensed

  End Function


  'This method returns the current amount to the vending machine coin return and resets the current amount to zero
  Public Function ReturnCurrentAmount() As Boolean

    Dim bHasCurrentAmount As Boolean
    Dim bCoinsReturned As Boolean
    Dim bCanReturnAmount As Boolean

    Dim decCurrentAmount As Decimal
    Dim decReturnedAmount As Decimal

    decCurrentAmount = prvCurrentAmount
    bHasCurrentAmount = (decCurrentAmount > 0)

    decReturnedAmount = 0

    If (bHasCurrentAmount) Then
      With prvCoinChanger
        bCanReturnAmount = .CanReturnAmount(decCurrentAmount)
        If (bCanReturnAmount) Then
          bCoinsReturned = .ReturnAmount(decCurrentAmount)
          If (bCoinsReturned) Then
            decReturnedAmount = decCurrentAmount
          End If
          prvCurrentAmount = 0
          ResetDisplayMessage() 'Note:  The Kata specifies INSERT COIN, but this is inconsistent with the prior use of INSERT COINS
        Else
          prvDisplayMessage = "ERROR: CANNOT RETURN AMOUNT"
        End If
      End With
    End If

    prvLastReturnedAmount = decReturnedAmount

    ReturnCurrentAmount = bCoinsReturned

  End Function


  Public Function EmptyCoinChanger(RemainingQuarterCount As Long, _
                                   RemainingDimeCount As Long, _
                                   RemainingNickelCount As Long _
                                   ) As Boolean

    Dim bCoinChangerEmptied As Boolean

    bCoinChangerEmptied = prvCoinChanger.EmptyCoinChanger(RemainingQuarterCount, RemainingDimeCount, RemainingNickelCount)
    ResetDisplayMessage()
    Return bCoinChangerEmptied

  End Function

End Class

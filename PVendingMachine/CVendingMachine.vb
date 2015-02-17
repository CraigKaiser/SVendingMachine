Public Class CVendingMachine

  'Stock the vending machine.  There is one item of each product in stock
  Private prvProductCandy As New CVendingMachineProduct("Candy", 0.65, 1)
  Private prvProductChips As New CVendingMachineProduct("Chips", 0.5, 1)
  Private prvProductCola As New CVendingMachineProduct("Cola", 1.0, 1)


  'The current value of coins accepted by the vending machine and not applied to a purchase
  Private prvCurrentAmount As Decimal
  Public ReadOnly Property CurrentAmount() As Decimal
    Get
      Return prvCurrentAmount
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
        prvDisplayMessage = "INSERT COINS"
      Else
        prvDisplayMessage = Format(decCurrentAmount, "$0.00")
      End If

      'Return the value of the display message at the time this method was called
      Return sDisplayMessage
    End Get
  End Property


  'This method accepts a coin and either keeps the coin and credits the current amount with the value
  'of the coin or rejects the coin.  It returns true when a coin was accepted
  Public Function AcceptCoin(ByVal SizeInMillimeters As Double, _
                             ByVal WeightInGrams As Double _
                             ) As Boolean

    Dim bIsValidCoin As Boolean
    Dim decCoinValue As Decimal

    decCoinValue = IdentifyCoinValueInDollars(SizeInMillimeters, WeightInGrams)
    bIsValidCoin = (decCoinValue > 0)

    If (bIsValidCoin) Then
      prvCurrentAmount = prvCurrentAmount + decCoinValue
      prvDisplayMessage = Format(prvCurrentAmount, "$0.00")
    End If

    AcceptCoin = bIsValidCoin

  End Function


  'This method dispenses a product and updates the current amount when the current amount is enough
  'to purchase the product.  It updates the display message to indicate the product price when the
  'current amount is insufficient to purchase the product.
  'It returns true when a product was dispensed
  Public Function DispenseProduct(ByVal ProductName As String) As Boolean

    Dim bIsProductAvailable As Boolean
    Dim bIsProductInStock As Boolean
    Dim bAreFundsAvailable As Boolean
    Dim bWasProductDispensed As Boolean

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
          bAreFundsAvailable = (prvCurrentAmount >= .Price)

          If (bAreFundsAvailable) Then
            'Enough coins have been inserted to purchase the product.  Dispense it.
            'TODO: This is the point where the command to physically dispense the product would occur
            .SetQuantityInStock(.QuantityInStock - 1)
            bWasProductDispensed = True
            prvDisplayMessage = "THANK YOU"
            prvCurrentAmount = 0 'Note: Per design, extra coins are kept by the vending machine at the time of purchase

          Else
            prvDisplayMessage = "PRICE " & Format(.Price, "$0.00")
          End If
        End If
      End With
    Else
      'If the product is not available, an error in the integration test product name exists
      'No changes to the vending machine display are specified for this condition.  Exit silently
    End If

    DispenseProduct = bWasProductDispensed

  End Function


  Public Sub New()
    'Initialize the vending machine
    prvDisplayMessage = "INSERT COINS"
  End Sub

End Class

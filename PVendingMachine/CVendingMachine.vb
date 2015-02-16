Public Class CVendingMachine


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
      Return prvDisplayMessage
    End Get
  End Property


  'This method accepts a coin and either keeps the coin and credits the current amount with the value
  'of the coin rejects the coin.  It returns true when a coin was accepted
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


  Public Sub New()
    'Initialize the vending machine
    prvDisplayMessage = "INSERT COINS"
  End Sub

End Class

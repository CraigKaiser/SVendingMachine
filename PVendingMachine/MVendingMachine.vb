Module MVendingMachine

  'Definition of coin names that could be inserted into the vending machine and their parameters
  'Name, Size (mm), Weight (g)
  Public CoinNameSizeWeight()() As Object = {({"Penny", 19.05, 2.5}), ({"Nickel", 21.21, 5.0}), ({"Dime", 17.91, 2.268}), ({"Quarter", 24.26, 5.67}), ({"Half Dollar", 30.61, 11.34})}

  'Definition of product names that could be selected from the vending machine and their parameters
  'Name, Price (US Dollars)
  Public ProductNamePrice()() As Object = {({"Cola", 1}), ({"Chips", 0.5}), ({"Candy", 0.65})}


  'Definition of coins that will be accepted by the vending machine:  US nickel, dime, and quarter
  'Size (mm), Weight (g), Value (US Dollars)
  Private prvValidCoinSizeWeightValue()() As Double = {({21.21, 5.0, 0.05}), ({17.91, 2.268, 0.1}), ({24.26, 5.67, 0.25})}


  'This function compares the coin parameters supplied to the set of valid coins that the vending machine will accept
  'If the parameters match a valid coin, the dollar value of that coin is returned.
  'Otherwise, zero is returned (the coin should not be accepted)

  Public Function IdentifyCoinValueInDollars(ByVal SizeInMillimeters As Double, _
                                             ByVal WeightInGrams As Double _
                                             ) As Decimal

    Dim dblSizeInMillimeters As Double
    Dim dblWeightInGrams As Double

    Dim decValueInDollars As Decimal
    Dim decCoinValue As Decimal

    Dim aCoin(2) As Double

    decCoinValue = 0

    'Compare the provided coin parameters to the set of valid coins
    For iCoin = 0 To prvValidCoinSizeWeightValue.GetUpperBound(0)
      aCoin = prvValidCoinSizeWeightValue(iCoin)
      dblSizeInMillimeters = aCoin(0)
      dblWeightInGrams = aCoin(1)
      decValueInDollars = aCoin(2)

      If ((SizeInMillimeters = dblSizeInMillimeters) And (WeightInGrams = dblWeightInGrams)) Then
        'This coin matches all parameters of a valid coin definition.  Return the value
        decCoinValue = decValueInDollars
        Exit For
      End If
    Next iCoin

    IdentifyCoinValueInDollars = decCoinValue

  End Function
End Module

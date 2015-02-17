'FitNesse test fixtures to automate acceptance tests
'Parameters to be passed into methods are declared as public properties
'The test fixtures perform the call to the underlying feature with the parameters it requires and returns the value

Public Class CTestFixture
  Inherits fit.ColumnFixture

  Public SizeInMillimeters As Double
  Public WeightInGrams As Double

  Public CoinName As String
  Public ProductName As String

  Public VendingMachine As New CVendingMachine


  Public Function testIdentifyCoinValueInDollars() As Decimal
    testIdentifyCoinValueInDollars = IdentifyCoinValueInDollars(SizeInMillimeters, WeightInGrams)
  End Function


  Public Function testInsertCoin() As Boolean

    Dim bFoundCoin As Boolean
    Dim bIsCoinAccepted As Boolean

    Dim dblSizeInMillimeters As Double
    Dim dblWeightInGrams As Double

    Dim sCoinName As String

    Dim aCoin(2) As Object

    dblSizeInMillimeters = 0
    dblWeightInGrams = 0

    sCoinName = CoinName
    'Compare the provided coin name to the set of defined coins
    For iCoin = 0 To CoinNameSizeWeight.GetUpperBound(0)
      aCoin = CoinNameSizeWeight(iCoin)
      sCoinName = aCoin(0)

      bFoundCoin = (CoinName = sCoinName)
      If (bFoundCoin) Then
        'This coin name exists.  Use its parameters
        dblSizeInMillimeters = aCoin(1)
        dblWeightInGrams = aCoin(2)
        Exit For
      End If
    Next iCoin

    If (bFoundCoin) Then
      With VendingMachine
        bIsCoinAccepted = .AcceptCoin(dblSizeInMillimeters, dblWeightInGrams)
      End With
    End If

    testInsertCoin = bIsCoinAccepted
  End Function


  Public Function testDisplayMessage() As String
    testDisplayMessage = VendingMachine.DisplayMessage
  End Function


  Public Function testCurrentValue() As Decimal
    testCurrentValue = VendingMachine.CurrentAmount
  End Function


  Public Function testSelectProduct() As Boolean

    Dim bFoundProduct As Boolean
    Dim bIsProductDispensed As Boolean

    Dim decProductPrice As Decimal

    Dim sProductName As String

    Dim aProduct(1) As Object

    decProductPrice = 0

    sProductName = ProductName
    'Compare the provided Product name to the set of defined Products
    For iProduct = 0 To ProductNamePrice.GetUpperBound(0)
      aProduct = ProductNamePrice(iProduct)
      sProductName = aProduct(0)

      bFoundProduct = (ProductName = sProductName)
      If (bFoundProduct) Then
        'This Product name exists.  Use its parameters
        decProductPrice = aProduct(1)
        Exit For
      End If
    Next iProduct

    If (bFoundProduct) Then
      With VendingMachine
        bIsProductDispensed = .DispenseProduct(sProductName)
      End With
    End If

    testSelectProduct = bIsProductDispensed
  End Function

End Class

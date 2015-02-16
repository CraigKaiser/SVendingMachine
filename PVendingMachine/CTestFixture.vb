'FitNesse test fixtures to automate acceptance tests
'Parameters to be passed into methods are declared as public properties
'The test fixtures perform the call to the underlying feature with the parameters it requires and returns the value

Public Class CTestFixture
  Inherits fit.ColumnFixture

  Public SizeInMillimeters As Double
  Public WeightInGrams As Double

  Public Function testIdentifyCoinValueInDollars() As Decimal
    testIdentifyCoinValueInDollars = IdentifyCoinValueInDollars(SizeInMillimeters, WeightInGrams)
  End Function

End Class

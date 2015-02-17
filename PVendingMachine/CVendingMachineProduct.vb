Public Class CVendingMachineProduct

  'The name of the product
  Private prvName As String
  Public ReadOnly Property Name() As String
    Get
      Return prvName
    End Get
  End Property

  Friend Sub SetName(ByVal Name As Decimal)
    prvName = Name
  End Sub


  'The selling price of the product 
  Private prvPrice As Decimal
  Public ReadOnly Property Price() As Decimal
    Get
      Return prvPrice
    End Get
  End Property

  Friend Sub SetPrice(ByVal Price As Decimal)
    prvPrice = Price
  End Sub


  'The current number of products in stock in the vending machine
  Private prvQuantityInStock As Long
  Public ReadOnly Property QuantityInStock As Long
    Get
      Return prvQuantityInStock
    End Get
  End Property

  Friend Sub SetQuantityInStock(ByVal QuantityInStock As Long)
    prvQuantityInStock = QuantityInStock
  End Sub


  Public Sub New(ByVal Name As String, _
                 ByVal Price As Decimal, _
                 ByVal QuantityInStock As Long)

    prvName = Name
    prvPrice = Price
    prvQuantityInStock = QuantityInStock
  End Sub


  Public Sub New(ByVal Name As String, _
                 ByVal Price As Decimal)

    prvName = Name
    prvPrice = Price
  End Sub
End Class

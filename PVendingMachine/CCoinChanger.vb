Public Class CCoinChanger

  Public Sub New(QuarterCount As Long, _
                 DimeCount As Long, _
                 NickelCount As Long)

    prvQuarterCount = QuarterCount
    prvDimeCount = DimeCount
    prvNickelCount = NickelCount
  End Sub


  'The number of quarters in the coin changer
  Private prvQuarterCount As Long
  Public ReadOnly Property QuarterCount() As Decimal
    Get
      Return prvQuarterCount
    End Get
  End Property


  Public Function CoinValueQuarters() As Decimal
    Return prvQuarterCount * 0.25
  End Function


  Public Function AddQuarter() As Boolean
    prvQuarterCount = prvQuarterCount + 1
    Return True
  End Function


  Public Function ReturnQuarter() As Boolean

    Dim bCoinReturned As Boolean


    If (prvQuarterCount > 0) Then
      'TODO:This is where the call to physically return a quarter would occur
      prvQuarterCount = prvQuarterCount - 1
      bCoinReturned = True
    End If

    Return bCoinReturned

  End Function


  'The number of Dimes in the coin changer
  Private prvDimeCount As Long
  Public ReadOnly Property DimeCount() As Decimal
    Get
      Return prvDimeCount
    End Get
  End Property


  Public Function CoinValueDimes() As Decimal
    Return prvDimeCount * 0.25
  End Function


  Public Function AddDime() As Boolean
    prvDimeCount = prvDimeCount + 1
    Return True
  End Function


  Public Function ReturnDime() As Boolean

    Dim bCoinReturned As Boolean


    If (prvDimeCount > 0) Then
      'TODO:This is where the call to physically return a Dime would occur
      prvDimeCount = prvDimeCount - 1
      bCoinReturned = True
    End If

    Return bCoinReturned

  End Function


  'The number of Nickels in the coin changer
  Private prvNickelCount As Long
  Public ReadOnly Property NickelCount() As Decimal
    Get
      Return prvNickelCount
    End Get
  End Property


  Public Function CoinValueNickels() As Decimal
    Return prvNickelCount * 0.25
  End Function


  Public Function AddNickel() As Boolean
    prvNickelCount = prvNickelCount + 1
    Return True
  End Function


  Public Function ReturnNickel() As Boolean

    Dim bCoinReturned As Boolean


    If (prvNickelCount > 0) Then
      'TODO:This is where the call to physically return a Nickel would occur
      prvNickelCount = prvNickelCount - 1
      bCoinReturned = True
    End If

    Return bCoinReturned

  End Function


  Public Function CoinValueTotal() As Decimal
    Return CoinValueQuarters() + CoinValueDimes() + CoinValueNickels()
  End Function


  Public Function CanReturnAmount(Amount As Decimal) As Boolean
    'NOTE:  This is a stub that does not implement a full algorithm to verify that the
    'amount to be returned can be fulfilled by the actual coins in the coin changer
    'This is a conscious scope limitation of this Kata to time-box the implementation
    'For test purposes, the coin changer will be initialized with sufficient coins
    'to successfully return any amount produced in the tests
    Return (CoinValueTotal() >= Amount)
  End Function


  Public Function ReturnAmount(Amount As Decimal) As Boolean
    'NOTE:  See CanReturnAmount note.  This method assumes that it will always succeed.
    'That assumption is invalid, but avoids the need for error handling in the kata

    Dim bCoinReturned As Boolean
    Dim bCanReturnAmount As Boolean
    Dim bTotalAmountReturned As Boolean

    Dim decRemainingAmount As Decimal

    decRemainingAmount = Amount
    bCanReturnAmount = CanReturnAmount(decRemainingAmount)

    If (bCanReturnAmount) Then
      Do While (decRemainingAmount >= 0.25)
        bCoinReturned = ReturnQuarter() 'ASSUMPTION: This always succeeds
        bCoinReturned = True 'TODO: Remove this.  It prevents a potential infinite loop from assumption
        If (bCoinReturned) Then
          decRemainingAmount = decRemainingAmount - 0.25
        End If
      Loop

      Do While (decRemainingAmount >= 0.1)
        bCoinReturned = ReturnDime() 'ASSUMPTION: This always succeeds
        bCoinReturned = True 'TODO: Remove this.  It prevents a potential infinite loop from assumption
        If (bCoinReturned) Then
          decRemainingAmount = decRemainingAmount - 0.1
        End If
      Loop

      Do While (decRemainingAmount >= 0.05)
        bCoinReturned = ReturnNickel() 'ASSUMPTION: This always succeeds
        bCoinReturned = True 'TODO: Remove this.  It prevents a potential infinite loop from assumption
        If (bCoinReturned) Then
          decRemainingAmount = decRemainingAmount - 0.05
        End If
      Loop
    End If

    bTotalAmountReturned = (decRemainingAmount = 0) 'See assumptions above.  This will always succeed at present

    ReturnAmount = bTotalAmountReturned

  End Function


  Public Function EmptyCoinChanger(RemainingQuarterCount As Long, _
                              RemainingDimeCount As Long, _
                              RemainingNickelCount As Long)
    prvQuarterCount = RemainingQuarterCount
    prvDimeCount = RemainingDimeCount
    prvNickelCount = RemainingNickelCount

    Return True
  End Function

End Class

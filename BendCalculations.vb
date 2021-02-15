Public Class BendCalculations
    Public Function CountBends(ByVal bend1X As Double, ByVal bend2X As Double, ByVal bend3X As Double, ByVal bend4X As Double, ByVal bend5X As Double, ByVal bend6X As Double, ByVal bend7X As Double, ByVal bend8X As Double, ByVal bend9X As Double, ByVal bend10X As Double, ByVal bend11X As Double) As Integer
        Dim countBendX As Integer


        If bend1X > 0 Then
            countBendX = 1
        End If

        If bend2X > 0 Then
            countBendX = 2
        End If

        If bend3X > 0 Then
            countBendX = 3
        End If

        If bend4X > 0 Then
            countBendX = 4
        End If

        If bend5X > 0 Then
            countBendX = 5
        End If

        If bend6X > 0 Then
            countBendX = 6
        End If

        If bend7X > 0 Then
            countBendX = 7
        End If

        If bend8X > 0 Then
            countBendX = 8
        End If

        If bend9X > 0 Then
            countBendX = 9
        End If

        If bend10X > 0 Then
            countBendX = 10
        End If

        If bend11X > 0 Then
            countBendX = 11
        End If

        CountBends = countBendX
    End Function
    Public Function CalculateSubjectLength(ByVal sheetthicknes As Double, ByVal bendX As Double, ByVal bend1X As Double, ByVal bend2X As Double, ByVal bend3X As Double, ByVal bend4X As Double, ByVal bend5X As Double, ByVal bend6X As Double, ByVal bend7X As Double, ByVal bend8X As Double, ByVal bend9X As Double, ByVal bend10X As Double, ByVal bend11X As Double) As Double
        Dim countBendX As Integer

        countBendX = CountBends(bend1X, bend2X, bend3X, bend4X, bend5X, bend6X, bend7X, bend8X, bend9X, bend10X, bend11X)
        CalculateSubjectLength = ((bendX + bend1X + bend2X + bend3X + bend4X + bend5X + bend6X + bend7X + bend8X + bend9X + bend10X + bend11X) - (sheetthicknes * 1.8 * countBendX))


    End Function
    Public Function CalculateAddition(ByVal bendaddition As ArrayList)
        Dim i As Integer
        Dim Sum As Double


        For i = 0 To bendaddition.Count - 1
            Sum = Sum + (SumBends(bendaddition, i) * 0.0006)
        Next i
        CalculateAddition = Sum
    End Function

    Private Function SumBends(ByVal Bends As ArrayList, ByVal Cursor As Integer)
        Dim i As Integer
        Dim Sum As Double


        For i = Cursor To Bends.Count - 1
            Sum = Sum + Bends.Item(i)
        Next i
        SumBends = Sum
    End Function
End Class

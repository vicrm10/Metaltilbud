Public Class CombiLaserCalculations

    Public Function CalculateLaserTimes(ByVal iNumberOfPlates As Integer, ByVal iPlateX As Integer, ByVal iPlateY As Integer, ByVal iSubjectX As Integer, ByVal iSubjectY As Integer, ByVal holes1 As Integer, ByVal holes2 As Integer, ByVal holes3 As Integer, ByVal holes4 As Integer, ByVal cuttinglength As Double, ByVal matrgruppe As Integer, ByVal sheetthickness As Double, ByVal Difficultfactor As Double, ByVal Antal As Double) As Double

        Dim subjectcuttinglength As Double
        Dim holecuttinglength1 As Double
        Dim holecuttinglength2 As Double
        Dim holecuttinglength3 As Double
        Dim holecuttinglength4 As Double
        Dim Totalcuttinglength As Double
        Dim Countholes As Integer
        Dim ilgangmin As Double
        Dim CuttingTimemin As Double
        Dim Removesubjectmin As Double



        subjectcuttinglength = (iSubjectX + iSubjectY) * 2
        holecuttinglength1 = holes1 * 20
        holecuttinglength2 = holes2 * 100
        holecuttinglength3 = holes3 * 230
        holecuttinglength4 = cuttinglength

        Totalcuttinglength = subjectcuttinglength + holecuttinglength1 + holecuttinglength2 + holecuttinglength3 + holecuttinglength4

        Countholes = holes1 + holes2 + holes3 + holes4

        ilgangmin = ((Math.Sqrt(Math.Pow(iSubjectX, 2) + (Math.Pow(iSubjectY, 2))) / (3 * 60000)) * Countholes * Antal)


        CuttingTimemin = (Totalcuttinglength * Antal) / (1000 * calculatecuttingspeed(Val(matrgruppe), Val(sheetthickness)))

        Removesubjectmin = (3 * Antal * Difficultfactor) / 60

        CalculateLaserTimes = ilgangmin + CuttingTimemin


    End Function

    Private Function calculatecuttingspeed(ByVal matrgruppe As Integer, ByVal sheetthickness As Double) As Double

        Select Case matrgruppe
            Case 1
                calculatecuttingspeed = ((1 / (Math.Pow(sheetthickness + 3, 1.3))) * 50) - 1 '(jern)

            Case 2
                calculatecuttingspeed = ((1 / (Math.Pow(sheetthickness + 1.1, 1.3))) * 32) - 2.2 '(rustfri)

            Case 3
                calculatecuttingspeed = ((1 / (Math.Pow(sheetthickness + 1.1, 1.3))) * 41) - 3.95 '(ALU)

            Case 4
                calculatecuttingspeed = 1 '(messing)

            Case 5
                calculatecuttingspeed = 1.5 '(aluzink)


        End Select
    End Function


End Class

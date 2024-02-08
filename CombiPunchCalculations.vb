Public Class CombiPunchCalculations

    Public Function CalculatePunchTimes(ByVal iNumberOfPlates As Integer, ByVal iPlateX As Integer, ByVal iPlateY As Integer, ByVal iSubjectX As Integer, ByVal iSubjectY As Integer, ByVal iToolChanges As Integer, ByVal iNumberOfHoles As Integer, ByVal Difficultfactor As Double, ByVal Antal As Double) As Double
        Dim objDatabaseIO As DatabaseIO

        Dim punchmin As Double
        Dim ilgang1 As Double
        Dim ilgang2 As Double
        Dim ilgangmin As Double
        Dim toolshiftmin As Double
        Dim subjectremovemin As Double
        Dim Dtrumpfmin As Double
        Dim sheetshift As Double






        objDatabaseIO = New DatabaseIO
        punchmin = (iNumberOfHoles * Antal) / (5 * 60)

        ilgang1 = ((Math.Sqrt(Math.Pow(iSubjectX, 2) + (Math.Pow(iSubjectY, 2))) / 3) * iNumberOfHoles * Antal)
        ilgang2 = (Math.Sqrt(Math.Pow(iPlateX, 2) + (Math.Pow(iPlateY, 2))) * iNumberOfPlates)
        ilgangmin = (ilgang1 + ilgang2) / 60000

        toolshiftmin = (iToolChanges * iNumberOfPlates * 5) / 60

        If iSubjectX > 350 Or iSubjectY > 350 Then
            subjectremovemin = (15 * Antal) / 60
        Else
            subjectremovemin = (5 * Antal) / 60
        End If

        Dtrumpfmin = punchmin + ilgangmin + toolshiftmin + (subjectremovemin * Difficultfactor)

        If iNumberOfHoles < 1 Then
            Dtrumpfmin = 0
        End If

        sheetshift = ((iNumberOfPlates * 30 * Difficultfactor) + 16) / 60

        CalculatePunchTimes = Dtrumpfmin + sheetshift





    End Function





End Class

Public Class LaserCalculations

    Public Function CalculateLaserTimes(ByVal iNumberOfPlates As Integer, ByVal iPlateX As Integer, ByVal iPlateY As Integer, ByVal iSubjectX As Integer, ByVal iSubjectY As Integer, ByVal holes1 As Integer, ByVal holes2 As Integer, ByVal holes3 As Integer, ByVal holes4 As Integer, ByVal cuttinglength As Double, ByVal matrgruppe As Integer, ByVal sheetthickness As Double, ByVal iMaterialType As Integer, ByVal Difficultfactor As Double, ByVal Antal As Double, ByVal Modulefactor As Integer) As Double
        Dim objDatabaseIO As DatabaseIO

        Dim subjectcuttinglength As Double
        Dim holecuttinglength1 As Double
        Dim holecuttinglength2 As Double
        Dim holecuttinglength3 As Double
        Dim holecuttinglength4 As Double
        Dim Totalcuttinglength As Double
        Dim Countholes As Integer
        Dim ilgangmin As Double
        Dim sheetshiftmin As Double
        Dim cuttingtimemin As Double
        Dim Removesubjectmin As Double
        Dim totalpiercetime As Double
        'Dim dblTypeThicknessFactor As Double

        If Difficultfactor = 0 Then
            Exit Function
        End If

        objDatabaseIO = New DatabaseIO
        'dblTypeThicknessFactor = objDatabaseIO.GetLaserTypeThicknessFactor(iMaterialType, sheetthickness)
        subjectcuttinglength = (iSubjectX + iSubjectY) * 2
        holecuttinglength1 = holes1 * 20
        holecuttinglength2 = holes2 * 100
        holecuttinglength3 = holes3 * 230
        holecuttinglength4 = cuttinglength

        Totalcuttinglength = subjectcuttinglength + holecuttinglength1 + holecuttinglength2 + holecuttinglength3 + holecuttinglength4

        Countholes = holes1 + holes2 + holes3 + holes4

        ilgangmin = ((Math.Sqrt(Math.Pow(iSubjectX, 2) + (Math.Pow(iSubjectY, 2))) / (3 * 60000)) * Countholes * Antal)

        sheetshiftmin = ((iNumberOfPlates * (30 + 105)) / 60) * Difficultfactor

        totalpiercetime = (calculatepiercetime(Val(matrgruppe), Val(sheetthickness)) * (1 + Countholes)) / 60 * Antal

        cuttingtimemin = ((Totalcuttinglength * Antal) / (1000 * calculatecuttingspeed(Val(matrgruppe), Val(sheetthickness))) + totalpiercetime)

        'Removesubjectmin = ((3 * Antal * Difficultfactor * Math.Sqrt(Modulefactor) / 60))
        Removesubjectmin = ((2 * Antal * Difficultfactor * Modulefactor / 60))

        CalculateLaserTimes = ilgangmin + cuttingtimemin + sheetshiftmin + Removesubjectmin

        'cuttingtime = cuttingtimemin / 60
    End Function
    Public Function calculatecuttingtime(ByVal iNumberOfPlates As Integer, ByVal iPlateX As Integer, ByVal iPlateY As Integer, ByVal iSubjectX As Integer, ByVal iSubjectY As Integer, ByVal holes1 As Integer, ByVal holes2 As Integer, ByVal holes3 As Integer, ByVal holes4 As Integer, ByVal cuttinglength As Double, ByVal matrgruppe As Integer, ByVal sheetthickness As Double, ByVal iMaterialType As Integer, ByVal Difficultfactor As Double, ByVal Antal As Double, ByVal Modulefactor As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim subjectcuttinglength As Double
        Dim holecuttinglength1 As Double
        Dim holecuttinglength2 As Double
        Dim holecuttinglength3 As Double
        Dim holecuttinglength4 As Double
        Dim Totalcuttinglength As Double
        Dim Countholes As Integer
        Dim totalpiercetime As Double


        If Difficultfactor = 0 Then
            Exit Function
        End If

        objDatabaseIO = New DatabaseIO
        'dblTypeThicknessFactor = objDatabaseIO.GetLaserTypeThicknessFactor(iMaterialType, sheetthickness)
        subjectcuttinglength = (iSubjectX + iSubjectY) * 2
        holecuttinglength1 = holes1 * 20
        holecuttinglength2 = holes2 * 100
        holecuttinglength3 = holes3 * 230
        holecuttinglength4 = cuttinglength

        Totalcuttinglength = subjectcuttinglength + holecuttinglength1 + holecuttinglength2 + holecuttinglength3 + holecuttinglength4

        Countholes = holes1 + holes2 + holes3 + holes4


        totalpiercetime = (calculatepiercetime(Val(matrgruppe), Val(sheetthickness)) * (1 + Countholes)) / 60 * Antal

        calculatecuttingtime = ((Totalcuttinglength * Antal) / (1000 * calculatecuttingspeed(Val(matrgruppe), Val(sheetthickness))) + totalpiercetime)


    End Function


    Private Function calculatecuttingspeed(ByVal matrgruppe As Integer, ByVal sheetthickness As Double) As Double

        Select Case matrgruppe
            Case 1
                'calculatecuttingspeed = (1 / (Math.Pow(sheetthickness + 3, 1.3))) * 50
                'matrgruppe 1 jern

                If sheetthickness <= 1 Then
                    calculatecuttingspeed = 60
                End If
                If sheetthickness > 1 Then
                    calculatecuttingspeed = 50
                End If
                If sheetthickness > 1.5 Then
                    calculatecuttingspeed = 30
                End If
                If sheetthickness > 2.5 Then
                    calculatecuttingspeed = 15
                End If
                If sheetthickness > 4 Then
                    calculatecuttingspeed = 8.5
                End If
                If sheetthickness > 5 Then
                    calculatecuttingspeed = 5.5
                End If
                If sheetthickness > 6 Then
                    calculatecuttingspeed = 3.5
                End If
                If sheetthickness > 8 Then
                    calculatecuttingspeed = 2.2
                End If
                If sheetthickness > 12 Then
                    calculatecuttingspeed = 1.8
                End If
                If sheetthickness > 15 Then
                    calculatecuttingspeed = 1
                End If
            Case 2
                'calculatecuttingspeed = ((1 / (Math.Pow(sheetthickness + 1.1, 1.3))) * 32) - 0.2
                'matrgruppe 2 rustfri

                If sheetthickness <= 1 Then
                    calculatecuttingspeed = 60
                End If
                If sheetthickness > 1 Then
                    calculatecuttingspeed = 52
                End If
                If sheetthickness > 1.5 Then
                    calculatecuttingspeed = 40
                End If
                If sheetthickness > 2 Then
                    calculatecuttingspeed = 29
                End If
                If sheetthickness > 2.5 Then
                    calculatecuttingspeed = 21
                End If
                If sheetthickness > 3 Then
                    calculatecuttingspeed = 9.5
                End If
                If sheetthickness > 4 Then
                    calculatecuttingspeed = 7
                End If
                If sheetthickness > 5 Then
                    calculatecuttingspeed = 5.5
                End If
                If sheetthickness > 6 Then
                    calculatecuttingspeed = 3.4
                End If
                If sheetthickness > 8 Then
                    calculatecuttingspeed = 2.2
                End If
                If sheetthickness > 10 Then
                    calculatecuttingspeed = 1.5
                End If
                If sheetthickness > 12 Then
                    calculatecuttingspeed = 1
                End If
                If sheetthickness > 15 Then
                    calculatecuttingspeed = 0.25
                End If
                If sheetthickness > 20 Then
                    calculatecuttingspeed = 0.17
                End If
                If sheetthickness > 25 Then
                    calculatecuttingspeed = 0.15
                End If
            Case 3
                'calculatecuttingspeed = ((1 / (Math.Pow(sheetthickness + 1.1, 1.3))) * 41) - 0.95
                'matrgruppe 3 alu

                If sheetthickness <= 1 Then
                    calculatecuttingspeed = 30
                End If
                If sheetthickness > 1 Then
                    calculatecuttingspeed = 20
                End If
                If sheetthickness > 1.5 Then
                    calculatecuttingspeed = 10
                End If
                If sheetthickness > 2.5 Then
                    calculatecuttingspeed = 6.5
                End If
                If sheetthickness > 4 Then
                    calculatecuttingspeed = 7.5
                End If
                If sheetthickness > 6 Then
                    calculatecuttingspeed = 4.8
                End If
                If sheetthickness > 8 Then
                    calculatecuttingspeed = 2
                End If
                If sheetthickness > 12 Then
                    calculatecuttingspeed = 1.2
                End If
                If sheetthickness > 15 Then
                    calculatecuttingspeed = 0.7
                End If
                If sheetthickness > 20 Then
                    calculatecuttingspeed = 0.3
                End If
                
            Case 4
                'calculatecuttingspeed = 1.5
                'matrgruppe 4 messing-kobber

                If sheetthickness <= 1 Then
                    calculatecuttingspeed = 35
                End If
                If sheetthickness > 1 Then
                    calculatecuttingspeed = 31
                End If
                If sheetthickness > 1.5 Then
                    calculatecuttingspeed = 22
                End If
                If sheetthickness > 2.5 Then
                    calculatecuttingspeed = 14
                End If
                If sheetthickness > 3 Then
                    calculatecuttingspeed = 8
                End If
                If sheetthickness > 5 Then
                    calculatecuttingspeed = 6
                End If
                If sheetthickness > 6 Then
                    calculatecuttingspeed = 3
                End If
                If sheetthickness > 8 Then
                    calculatecuttingspeed = 1
                End If

            Case 5
                'calculatecuttingspeed = 2
                'matrgruppe 5 aluzink-elgalv-varmgalv

                If sheetthickness <= 1 Then
                    calculatecuttingspeed = 35
                End If
                If sheetthickness > 1 Then
                    calculatecuttingspeed = 31
                End If
                If sheetthickness > 1.5 Then
                    calculatecuttingspeed = 22
                End If
                If sheetthickness > 2.5 Then
                    calculatecuttingspeed = 14
                End If
                If sheetthickness > 3 Then
                    calculatecuttingspeed = 10
                End If
        End Select
    End Function
    Private Function calculatepiercetime(ByVal matrgruppe As Integer, ByVal sheetthickness As Double) As Double
        Dim piercetime As Double

        Select Case matrgruppe
            Case 1 'jern

                If sheetthickness > 6 Then
                    piercetime = 0.6
                End If
                If sheetthickness > 8 Then
                    piercetime = 0.7
                End If
                If sheetthickness > 10 Then
                    piercetime = 0.7
                End If
                If sheetthickness > 12 Then
                    piercetime = 1.5
                End If
                If sheetthickness > 15 Then
                    piercetime = 2
                End If

            Case 2 'rustfri

                If sheetthickness > 6 Then
                    piercetime = 0.4
                End If
                If sheetthickness > 8 Then
                    piercetime = 0.7
                End If
                If sheetthickness > 10 Then
                    piercetime = 2.4
                End If
                If sheetthickness > 12 Then
                    piercetime = 20
                End If
                If sheetthickness > 15 Then
                    piercetime = 48
                End If
                If sheetthickness > 20 Then
                    piercetime = 80
                End If
                If sheetthickness > 25 Then
                    piercetime = 100
                End If

            Case 3 'alu

                If sheetthickness > 6 Then
                    piercetime = 0.25
                End If
                If sheetthickness > 8 Then
                    piercetime = 0.6
                End If
                If sheetthickness > 10 Then
                    piercetime = 1
                End If
                If sheetthickness > 12 Then
                    piercetime = 1.8
                End If
                If sheetthickness > 15 Then
                    piercetime = 3
                End If
                If sheetthickness > 20 Then
                    piercetime = 20
                End If

            Case 4 'messing

                If sheetthickness > 6 Then
                    piercetime = 0.23
                End If
                If sheetthickness > 8 Then
                    piercetime = 1
                End If

            Case 5 'aluzink

                If sheetthickness > 6 Then
                    piercetime = 0.23
                End If
                If sheetthickness > 8 Then
                    piercetime = 1
                End If

        End Select

        calculatepiercetime = piercetime

    End Function
    Public Function calculatenitrogen(ByVal matrgruppe As Integer, ByVal sheetthickness As Double) As Double
        Dim nitrogen As Double
        Select Case matrgruppe

            Case 1 'jern

                If sheetthickness < 1.25 Then
                    nitrogen = 17
                End If
                If sheetthickness > 1.5 Then
                    nitrogen = 30
                End If
                If sheetthickness > 2 Then
                    nitrogen = 40
                End If
                If sheetthickness > 3 Then
                    nitrogen = 75
                End If
                If sheetthickness > 4 Then
                    nitrogen = 61
                End If
                If sheetthickness > 5 Then
                    nitrogen = 78
                End If
                If sheetthickness > 8 Then
                    nitrogen = 70
                End If

            Case 2 'rustfri

                If sheetthickness < 1.25 Then
                    nitrogen = 37
                End If
                If sheetthickness > 1.5 Then
                    nitrogen = 45
                End If
                If sheetthickness > 2 Then
                    nitrogen = 62
                End If
                If sheetthickness > 3 Then
                    nitrogen = 66
                End If
                If sheetthickness > 4 Then
                    nitrogen = 66
                End If
                If sheetthickness > 5 Then
                    nitrogen = 70
                End If
                If sheetthickness > 6 Then
                    nitrogen = 74
                End If
                If sheetthickness > 8 Then
                    nitrogen = 78
                End If
                If sheetthickness > 10 Then
                    nitrogen = 87
                End If
                If sheetthickness > 12 Then
                    nitrogen = 123
                End If
                If sheetthickness > 15 Then
                    nitrogen = 119
                End If
                If sheetthickness > 20 Then
                    nitrogen = 110
                End If
                If sheetthickness > 25 Then
                    nitrogen = 77
                End If

            Case 3 'alu

                If sheetthickness < 1.25 Then
                    nitrogen = 9.4
                End If
                If sheetthickness > 1 Then
                    nitrogen = 13
                End If
                If sheetthickness > 1.5 Then
                    nitrogen = 22
                End If
                If sheetthickness > 2 Then
                    nitrogen = 25
                End If
                If sheetthickness > 3 Then
                    nitrogen = 30
                End If
                If sheetthickness > 4 Then
                    nitrogen = 47
                End If
                If sheetthickness > 5 Then
                    nitrogen = 50
                End If
                If sheetthickness > 6 Then
                    nitrogen = 55
                End If
                If sheetthickness > 8 Then
                    nitrogen = 61
                End If
                If sheetthickness > 10 Then
                    nitrogen = 71
                End If
                If sheetthickness > 12 Then
                    nitrogen = 70
                End If
                If sheetthickness > 15 Then
                    nitrogen = 82
                End If
                If sheetthickness > 20 Then
                    nitrogen = 90
                End If

            Case 4 'messing

                If sheetthickness < 1.25 Then
                    nitrogen = 17
                End If
                If sheetthickness > 1 Then
                    nitrogen = 17
                End If
                If sheetthickness > 1.5 Then
                    nitrogen = 24
                End If
                If sheetthickness > 2 Then
                    nitrogen = 28
                End If
                If sheetthickness > 3 Then
                    nitrogen = 57
                End If
                If sheetthickness > 4 Then
                    nitrogen = 57
                End If
                If sheetthickness > 5 Then
                    nitrogen = 68
                End If
                If sheetthickness > 6 Then
                    nitrogen = 66
                End If
                If sheetthickness > 8 Then
                    nitrogen = 74
                End If

            Case 5 'aluzink-galv.

                If sheetthickness < 1.25 Then
                    nitrogen = 17
                End If
                If sheetthickness > 1 Then
                    nitrogen = 17
                End If
                If sheetthickness > 1.5 Then
                    nitrogen = 24
                End If
                If sheetthickness > 2 Then
                    nitrogen = 28
                End If
                If sheetthickness > 3 Then
                    nitrogen = 57
                End If
                If sheetthickness > 4 Then
                    nitrogen = 57
                End If
                If sheetthickness > 5 Then
                    nitrogen = 68
                End If
                If sheetthickness > 6 Then
                    nitrogen = 66
                End If
                If sheetthickness > 8 Then
                    nitrogen = 74
                End If

        End Select

        calculatenitrogen = nitrogen

    End Function


End Class

Imports Microsoft.SharePoint.Client
Imports SP = Microsoft.SharePoint.Client
Public Class DatabaseIO



    Public Function getnumbermultiplier(ByVal number As Integer) As Double

        If number >= 1000 Then
            getnumbermultiplier = 0.7
            Exit Function

        End If
        If number >= 500 Then
            getnumbermultiplier = 0.75 - ((number - 500) * 0.0001)
            Exit Function

        End If
        If number >= 250 Then
            getnumbermultiplier = 0.9 - ((number - 250) * 0.0006)
            Exit Function

        End If
        If number >= 100 Then
            getnumbermultiplier = 1 - ((number - 100) * 0.0007)
            Exit Function

        End If
        If number >= 50 Then
            getnumbermultiplier = 1.25 - ((number - 50) * 0.005)
            Exit Function

        End If
        If number >= 25 Then
            getnumbermultiplier = 1.4 - ((number - 25) * 0.006)
            Exit Function

        End If
        If number >= 5 Then
            getnumbermultiplier = 2 - ((number - 5) * 0.03)
            Exit Function

        End If

        getnumbermultiplier = 3 - ((number - 1) * 0.25)


    End Function

    Public Function getmodulemultiplier(ByVal modulesize As Integer) As Double
        If modulesize <= 1 Then
            getmodulemultiplier = 1
            Exit Function
        End If

        If modulesize <= 2 Then
            getmodulemultiplier = 1.5
            Exit Function
        End If

        If modulesize <= 3 Then
            getmodulemultiplier = 2
            Exit Function
        End If

        If modulesize <= 4 Then
            getmodulemultiplier = 2.5
            Exit Function
        End If

        If modulesize <= 5 Then
            getmodulemultiplier = 3
            Exit Function
        End If

        If modulesize <= 6 Then
            getmodulemultiplier = 7
            Exit Function
        End If

        If modulesize <= 7 Then
            getmodulemultiplier = 10
            Exit Function
        End If

    End Function

    Public Function getcountersinkmultiplier(ByVal materialID As Integer) As Double
        Select Case materialID
            Case 1
                getcountersinkmultiplier = 1
            Case 2
                getcountersinkmultiplier = 2
            Case 3
                getcountersinkmultiplier = 1.5
            Case 4
                getcountersinkmultiplier = 1
            Case 5
                getcountersinkmultiplier = 1

        End Select
    End Function

    Public Function getthreadmultiplier(ByVal materialID As Integer) As Double
        Select Case materialID
            Case 1
                getthreadmultiplier = 1
            Case 2
                getthreadmultiplier = 2
            Case 3
                getthreadmultiplier = 1.5
            Case 4
                getthreadmultiplier = 1
            Case 5
                getthreadmultiplier = 1
        End Select
    End Function
    Public Function getpressnutprice(ByVal materialID As Integer) As Double
        Select Case materialID
            Case 1
                getpressnutprice = 0.8
            Case 2
                getpressnutprice = 1.3
            Case 3
                getpressnutprice = 1.6
            Case 4
                getpressnutprice = 3
            Case 5
                getpressnutprice = 0.8
        End Select
    End Function
    Public Function GetPressnut100min(ByVal modulesize As Integer) As Double

        Select Case modulesize

            Case 1
                GetPressnut100min = 10
            Case 2
                GetPressnut100min = 10
            Case 3
                GetPressnut100min = 11
            Case 4
                GetPressnut100min = 12
            Case 5
                GetPressnut100min = 25
            Case 6
                GetPressnut100min = 30
            Case 7
                GetPressnut100min = 40
        End Select


    End Function
    Public Function getpresstagprice(ByVal materialID As Integer) As Double
        Select Case materialID
            Case 1
                getpresstagprice = 1.1
            Case 2
                getpresstagprice = 1.7
            Case 3
                getpresstagprice = 2.6
            Case 4
                getpresstagprice = 3.5
            Case 5
                getpresstagprice = 1.1
        End Select
    End Function
    Public Function GetPresstag100min(ByVal modulesize As Integer) As Double

        Select Case modulesize

            Case 1
                GetPresstag100min = 15
            Case 2
                GetPresstag100min = 15
            Case 3
                GetPresstag100min = 17
            Case 4
                GetPresstag100min = 18
            Case 5
                GetPresstag100min = 38
            Case 6
                GetPresstag100min = 45
            Case 7
                GetPresstag100min = 60
        End Select


    End Function

    Public Function GetPressnutStartup() As Double
        'opstart = 15 min.
        'kontrol = 5 min.

        GetPressnutStartup = 15 + 5

    End Function

    Public Function GetPresstagStartup() As Double
        'opstart = 15 min.
        'kontrol = 5 min.

        GetPresstagStartup = 15 + 5

    End Function


    Public Function getboltprice(ByVal materialID As Integer) As Double
        Select Case materialID
            Case 1
                getboltprice = 0.65
            Case 2
                getboltprice = 1
            Case 3
                getboltprice = 1.5
            Case 4
                getboltprice = 2
            Case 5
                getboltprice = 0.65
        End Select
    End Function
    Public Function GetBoltweld100min(ByVal modulesize As Integer) As Double

        Select Case modulesize

            Case 1
                GetBoltweld100min = 9
            Case 2
                GetBoltweld100min = 10
            Case 3
                GetBoltweld100min = 10.5
            Case 4
                GetBoltweld100min = 11
            Case 5
                GetBoltweld100min = 22
            Case 6
                GetBoltweld100min = 25
            Case 7
                GetBoltweld100min = 35
        End Select


    End Function
    Public Function GetBoltweldingStartup(ByVal inumberofbolts As Double) As Double
        'opstart = 30 min. indtil 20 stk.
        'derefter 15 min. pr 10 stk.
        'kontrol = 5 min.
        If inumberofbolts <= 20 Then
            GetBoltweldingStartup = 15 + 5
        Else
            GetBoltweldingStartup = 30 + 5 + (((inumberofbolts - 20) / 10) * 15)
        End If

    End Function
    Public Function GetDeburringTime100(ByVal modulesize As Integer) As Double
        Select Case modulesize

            Case 1
                GetDeburringTime100 = 12
            Case 2
                GetDeburringTime100 = 27
            Case 3
                GetDeburringTime100 = 34
            Case 4
                GetDeburringTime100 = 55
            Case 5
                GetDeburringTime100 = 150
            Case 6
                GetDeburringTime100 = 340
            Case 7
                GetDeburringTime100 = 600
        End Select

    End Function
    Public Function GetDeburringstartup() As Double
        'opstart 10 min.
        'kontrol 5 min.
        GetDeburringstartup = 10 + 5
    End Function

    Public Function GetSteelmasterTime100(ByVal modulesize As Integer) As Double
        Select Case modulesize

            Case 1
                GetSteelmasterTime100 = 10
            Case 2
                GetSteelmasterTime100 = 20
            Case 3
                GetSteelmasterTime100 = 30
            Case 4
                GetSteelmasterTime100 = 50
            Case 5
                GetSteelmasterTime100 = 120
            Case 6
                GetSteelmasterTime100 = 200
            Case 7
                GetSteelmasterTime100 = 400
        End Select

    End Function
    Public Function GetSteelmasterstartup() As Double
        'opstart 5 min.
        'kontrol 5 min.
        GetSteelmasterstartup = 5 + 5
    End Function
    Public Function GetSteelmastermultiplier(ByVal modulesize As Integer, ByVal inumberofsubjects As Integer) As Integer
        If modulesize = 1 Then
            If inumberofsubjects > 60 Then
                GetSteelmastermultiplier = 2
            Else
                GetSteelmastermultiplier = 1
            End If
        End If
        If modulesize = 2 Then
            If inumberofsubjects > 20 Then
                GetSteelmastermultiplier = 2
            Else
                GetSteelmastermultiplier = 1
            End If
        End If
        If modulesize = 3 Then
            If inumberofsubjects > 6 Then
                GetSteelmastermultiplier = 2
            Else
                GetSteelmastermultiplier = 1
            End If
        End If
        If modulesize = 4 Then
            If inumberofsubjects > 2 Then
                GetSteelmastermultiplier = 2
            Else
                GetSteelmastermultiplier = 1
            End If
        End If
        If modulesize > 4 Then
            GetSteelmastermultiplier = 2
        End If

    End Function
    Public Function GetStraigtheningTime100(ByVal modulesize As Integer) As Double
        Select Case modulesize

            Case 1
                GetStraigtheningTime100 = 12
            Case 2
                GetStraigtheningTime100 = 27
            Case 3
                GetStraigtheningTime100 = 34
            Case 4
                GetStraigtheningTime100 = 48
            Case 5
                GetStraigtheningTime100 = 75
            Case 6
                GetStraigtheningTime100 = 150
            Case 7
                GetStraigtheningTime100 = 200
        End Select

    End Function
    Public Function GetStraigtheningstartup() As Double
        'opstart 5 min.
        'kontrol 5 min.
        GetStraigtheningstartup = 5 + 5
    End Function
    Public Function GetStraigtheningmultiplier(ByVal modulesize As Integer, ByVal inumberofsubjects As Integer) As Integer
        If modulesize <= 3 Then
            GetStraigtheningmultiplier = 1
        Else
            GetStraigtheningmultiplier = 2
        End If

    End Function

    Public Function GetBrushingTime100(ByVal modulesize As Integer) As Double
        Select Case modulesize

            Case 1
                GetBrushingTime100 = 12
            Case 2
                GetBrushingTime100 = 27
            Case 3
                GetBrushingTime100 = 34
            Case 4
                GetBrushingTime100 = 55
            Case 5
                GetBrushingTime100 = 150
            Case 6
                GetBrushingTime100 = 340
            Case 7
                GetBrushingTime100 = 600
        End Select

    End Function

    Public Function GetBrushingstartup() As Double
        'opstart 10 min.
        'kontrol 5 min.
        GetBrushingstartup = 10 + 5
    End Function

    Public Function GetPlateSizes(ByVal Checkstate_over As String, ByVal Checkstate_stor As String, ByVal Checkstate_mellem As String, ByVal Checkstate_lille As String) As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList







        alReturn = New ArrayList(4)
        alRow = New ArrayList(2)
        alRow.Add(1000)
        alRow.Add(2000)
        alReturn.Add(alRow)
        alRow = New ArrayList(2)
        alRow.Add(1250)
        alRow.Add(2500)
        alReturn.Add(alRow)
        alRow = New ArrayList(2)
        alRow.Add(1500)
        alRow.Add(3000)
        alReturn.Add(alRow)
        alRow = New ArrayList(2)
        alRow.Add(2000)
        alRow.Add(4000)
        alReturn.Add(alRow)


        If Checkstate_over = "1" Then
            'lille,mellem,stor

            alReturn = New ArrayList(3)
            alRow = New ArrayList(2)
            alRow.Add(1000)
            alRow.Add(2000)
            alReturn.Add(alRow)
            alRow = New ArrayList(2)
            alRow.Add(1250)
            alRow.Add(2500)
            alReturn.Add(alRow)
            alRow = New ArrayList(2)
            alRow.Add(1500)
            alRow.Add(3000)
            alReturn.Add(alRow)

            If Checkstate_mellem = "1" Then
                'lille,stor

                alReturn = New ArrayList(2)
                alRow = New ArrayList(2)
                alRow.Add(1000)
                alRow.Add(2000)
                alReturn.Add(alRow)
                alRow = New ArrayList(2)
                alRow.Add(1500)
                alRow.Add(3000)
                alReturn.Add(alRow)

                If Checkstate_lille = "1" Then
                    'stor
                    alReturn = New ArrayList(1)
                    alRow = New ArrayList(2)
                    alRow.Add(1500)
                    alRow.Add(3000)
                    alReturn.Add(alRow)
                End If
            End If
        End If

        If Checkstate_lille = "1" Then
            'over,stor,mellem
            alReturn = New ArrayList(3)
            alRow = New ArrayList(2)
            alRow.Add(1250)
            alRow.Add(2500)
            alReturn.Add(alRow)
            alRow = New ArrayList(2)
            alRow.Add(1500)
            alRow.Add(3000)
            alReturn.Add(alRow)
            alRow = New ArrayList(2)
            alRow.Add(2000)
            alRow.Add(4000)
            alReturn.Add(alRow)


            If Checkstate_stor = "1" Then
                'over,mellem

                alReturn = New ArrayList(2)
                alRow = New ArrayList(2)
                alRow.Add(1250)
                alRow.Add(2500)
                alReturn.Add(alRow)
                alRow = New ArrayList(2)
                alRow.Add(2000)
                alRow.Add(4000)
                alReturn.Add(alRow)

                If Checkstate_mellem = "1" Then
                    'over

                    alReturn = New ArrayList(1)
                    alRow = New ArrayList(2)
                    alRow.Add(2000)
                    alRow.Add(4000)
                    alReturn.Add(alRow)

                End If
            End If
        End If



        If Checkstate_stor = "1" Then
            'lille,mellem,over

            alReturn = New ArrayList(3)
            alRow = New ArrayList(2)
            alRow.Add(1000)
            alRow.Add(2000)
            alReturn.Add(alRow)
            alRow = New ArrayList(2)
            alRow.Add(1250)
            alRow.Add(2500)
            alReturn.Add(alRow)
            alRow = New ArrayList(2)
            alRow.Add(2000)
            alRow.Add(4000)
            alReturn.Add(alRow)

            If Checkstate_mellem = "1" Then
                'lille,over

                alReturn = New ArrayList(2)
                alRow = New ArrayList(2)
                alRow.Add(1000)
                alRow.Add(2000)
                alReturn.Add(alRow)
                alRow = New ArrayList(2)
                alRow.Add(2000)
                alRow.Add(4000)
                alReturn.Add(alRow)

                If Checkstate_over = "1" Then

                    alReturn = New ArrayList(1)
                    alRow = New ArrayList(2)
                    alRow.Add(1000)
                    alRow.Add(2000)
                    alReturn.Add(alRow)

                End If
            End If
        End If

        If Checkstate_over = "1" Then
            If Checkstate_lille = "1" Then
                'mellem,stor

                alReturn = New ArrayList(2)
                alRow = New ArrayList(2)
                alRow.Add(1250)
                alRow.Add(2500)
                alReturn.Add(alRow)
                alRow = New ArrayList(2)
                alRow.Add(1500)
                alRow.Add(3000)
                alReturn.Add(alRow)

                If Checkstate_stor = "1" Then
                    'mellem

                    alReturn = New ArrayList(1)
                    alRow = New ArrayList(2)
                    alRow.Add(1250)
                    alRow.Add(2500)
                    alReturn.Add(alRow)
                End If
            End If
        End If
        If Checkstate_lille = "1" Then
            If Checkstate_mellem = "1" Then
                'over,stor

                alReturn = New ArrayList(2)
                alRow = New ArrayList(2)
                alRow.Add(1500)
                alRow.Add(3000)
                alReturn.Add(alRow)
                alRow = New ArrayList(2)
                alRow.Add(2000)
                alRow.Add(4000)
                alReturn.Add(alRow)

                If Checkstate_over = "1" Then
                    'stor

                    alReturn = New ArrayList(1)
                    alRow = New ArrayList(2)
                    alRow.Add(1500)
                    alRow.Add(3000)
                    alReturn.Add(alRow)

                    If Checkstate_stor = "1" Then
                        'over

                        alReturn = New ArrayList(1)
                        alRow = New ArrayList(2)
                        alRow.Add(2000)
                        alRow.Add(4000)
                        alReturn.Add(alRow)

                    End If
                End If
            End If
        End If

        If Checkstate_over = "1" Then
            If Checkstate_stor = "1" Then
                'lille,mellem

                alReturn = New ArrayList(2)
                alRow = New ArrayList(2)
                alRow.Add(1000)
                alRow.Add(2000)
                alReturn.Add(alRow)
                alRow = New ArrayList(2)
                alRow.Add(1250)
                alRow.Add(2500)
                alReturn.Add(alRow)

                If Checkstate_mellem = "1" Then
                    'lille

                    alReturn = New ArrayList(1)
                    alRow = New ArrayList(2)
                    alRow.Add(1000)
                    alRow.Add(2000)
                    alReturn.Add(alRow)
                End If

            End If
        End If



        GetPlateSizes = alReturn
    End Function
    Public Function GetWasteMinimumSizes(ByVal matrgruppe As Integer, ByVal dblThickness As Double) As Double
        Dim wastemin As Double

        If matrgruppe = 1 Then
            If dblThickness <= 2 Then
                wastemin = 1000
            End If
            If dblThickness >= 2.5 Then
                wastemin = 700
            End If
            If dblThickness >= 5 Then
                wastemin = 700
            End If
            If dblThickness > 8 Then
                wastemin = 700
            End If
        End If
        If matrgruppe = 2 Then
            If dblThickness <= 2 Then
                wastemin = 700
            End If
            If dblThickness >= 2.5 Then
                wastemin = 700
            End If
            If dblThickness >= 5 Then
                wastemin = 500
            End If
            If dblThickness > 8 Then
                wastemin = 500
            End If
        End If
        If matrgruppe = 3 Then
            If dblThickness <= 2 Then
                wastemin = 1000
            End If
            If dblThickness >= 2.5 Then
                wastemin = 700
            End If
            If dblThickness >= 5 Then
                wastemin = 500
            End If
            If dblThickness > 8 Then
                wastemin = 500
            End If
        End If
        If matrgruppe = 4 Then
            If dblThickness <= 2 Then
                wastemin = 700
            End If
            If dblThickness >= 2.5 Then
                wastemin = 700
            End If
            If dblThickness >= 5 Then
                wastemin = 500
            End If
            If dblThickness > 8 Then
                wastemin = 500
            End If
        End If
        If matrgruppe = 5 Then
            If dblThickness <= 2 Then
                wastemin = 1000
            End If
            If dblThickness >= 2.5 Then
                wastemin = 700
            End If
            If dblThickness >= 5 Then
                wastemin = 700
            End If
            If dblThickness > 8 Then
                wastemin = 700
            End If
        End If

        GetWasteMinimumSizes = wastemin


    End Function
    Public Function GetLaserpriceDK(ByVal matrgruppe As Integer, ByVal dblThickness As Double) As Double
        Dim Laserprice As Double
        Dim Manprice As Double
        Dim price1 As Double
        Dim price2 As Double
        Dim price3 As Double
        'laserpriser Danmark
        price1 = 850
        price2 = 1150
        price3 = 1450
        Manprice = 550

        'jern
        If matrgruppe = 1 Then
            If dblThickness <= 8 Then
                Laserprice = price1
            End If
            If dblThickness > 8 Then
                Laserprice = price3
            End If
        End If
        'el-galv
        If matrgruppe = 5 Then
            If dblThickness <= 8 Then
                Laserprice = price1
            End If
            If dblThickness > 8 Then
                Laserprice = price3
            End If
        End If
        'rustfri
        If matrgruppe = 2 Then
            If dblThickness <= 2 Then
                Laserprice = price1
            End If
            If dblThickness >= 2.5 Then
                Laserprice = price2
            End If
            If dblThickness >= 10 Then
                Laserprice = price3
            End If
        End If
        'alu
        If matrgruppe = 3 Then
            If dblThickness <= 4 Then
                Laserprice = price1
            End If
            If dblThickness >= 5 Then
                Laserprice = price2
            End If
        End If
        'messing
        If matrgruppe = 4 Then
            If dblThickness <= 2 Then
                Laserprice = price1
            End If
            If dblThickness >= 2.5 Then
                Laserprice = price2
            End If
            If dblThickness >= 10 Then
                Laserprice = price3
            End If
        End If

        'diverse omkostninger = omkostninger til lamelsæt + rengøring = Kr. 53.5
        GetLaserpriceDK = Laserprice + Manprice + 53.5



    End Function
    Public Function GetLaserpricePL(ByVal matrgruppe As Integer, ByVal dblThickness As Double) As Double
        Dim Laserprice As Double
        Dim Manprice As Double
        Dim price1 As Double
        Dim price2 As Double
        Dim price3 As Double
        'laserpriser Polen
        price1 = 650
        price2 = 850
        price3 = 1050
        Manprice = 300

        'jern
        If matrgruppe = 1 Then
            If dblThickness <= 8 Then
                Laserprice = price1
            End If
            If dblThickness >= 10 Then
                Laserprice = price3
            End If
        End If
        'el-galv
        If matrgruppe = 5 Then
            If dblThickness <= 8 Then
                Laserprice = price1
            End If
            If dblThickness >= 10 Then
                Laserprice = price3
            End If
        End If
        'rustfri
        If matrgruppe = 2 Then
            If dblThickness <= 2 Then
                Laserprice = price1
            End If
            If dblThickness >= 2.5 Then
                Laserprice = price2
            End If
            If dblThickness >= 10 Then
                Laserprice = price3
            End If
        End If
        'alu
        If matrgruppe = 3 Then
            If dblThickness <= 4 Then
                Laserprice = price1
            End If
            If dblThickness >= 5 Then
                Laserprice = price2
            End If
        End If
        'messing
        If matrgruppe = 4 Then
            If dblThickness <= 2 Then
                Laserprice = price1
            End If
            If dblThickness >= 2.5 Then
                Laserprice = price2
            End If
            If dblThickness >= 10 Then
                Laserprice = price3
            End If
        End If


        GetLaserpricePL = Laserprice + Manprice


    End Function

    Public Function GetMaterials() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList
        Dim prices As New ArrayList()
        Try
            prices = retreiveSharepointPrices()
        Catch ex As Exception

            MsgBox("Unable to get prices from Intranet. Old values will be use, Beware")

        End Try

        alReturn = New ArrayList(15)
        alRow = New ArrayList(7)
        alRow.Add(1)
        alRow.Add("Jern, P01 AM, 12.03")
        If prices.Count > 0 Then
            alRow.Add(prices(0))
        Else
            alRow.Add(8.5)
        End If
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(2)
        alRow.Add("ALU, 2S Al99")
        If prices.Count > 0 Then
            alRow.Add(prices(1))
        Else
            alRow.Add(25)
        End If
        alRow.Add(3)
        alRow.Add(1)
        alRow.Add(2)
        alRow.Add(3)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(3)
        alRow.Add("ALU, AlMg3, Søvandsbestandig")
        If prices.Count > 0 Then
            alRow.Add(prices(2))
        Else
            alRow.Add(26.8)
        End If
        alRow.Add(3)
        alRow.Add(1)
        alRow.Add(2)
        alRow.Add(3)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(4)
        alRow.Add("Rustfri AISI 304, 1.4301")
        If prices.Count > 0 Then
            alRow.Add(prices(3))
        Else
            alRow.Add(22.4)
        End If
        alRow.Add(2)
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(5)
        alRow.Add("Rustfri AISI 316L, 1.4404")
        If prices.Count > 0 Then
            alRow.Add(prices(4))
        Else
            alRow.Add(31.8)
        End If
        alRow.Add(2)
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(6)
        alRow.Add("Messing")
        If prices.Count > 0 Then
            alRow.Add(prices(5))
        Else
            alRow.Add(81)
        End If
        alRow.Add(4)
        alRow.Add(1)
        alRow.Add(2)
        alRow.Add(8.4)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(7)
        alRow.Add("Elgalv, Fe P01 ZE, Zintec")
        If prices.Count > 0 Then
            alRow.Add(prices(6))
        Else
            alRow.Add(10.6)
        End If
        alRow.Add(5)
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(8)
        alRow.Add("Aluzink, B500A")
        If prices.Count > 0 Then
            alRow.Add(prices(7))
        Else
            alRow.Add(10.6)
        End If
        alRow.Add(5)
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(9)
        alRow.Add("ALU, AlMg3, Søvandsbest.m.PVC")
        If prices.Count > 0 Then
            alRow.Add(prices(8))
        Else
            alRow.Add(28.1)
        End If
        alRow.Add(3)
        alRow.Add(1.75)
        alRow.Add(2)
        alRow.Add(3)
        alReturn.Add(alRow)
        alRow = New ArrayList(7)
        alRow.Add(10)
        alRow.Add("ALU, AlMg1,m.PVC")
        If prices.Count > 0 Then
            alRow.Add(prices(9))
        Else
            alRow.Add(26.3)
        End If
        alRow.Add(3)
        alRow.Add(1.75)
        alRow.Add(2)
        alRow.Add(3)
        alReturn.Add(alRow)
        alRow = New ArrayList(11)
        alRow.Add(11)
        alRow.Add("Rustfri AISI 304, 1.4301 m.PVC")
        If prices.Count > 0 Then
            alRow.Add(prices(10))
        Else
            alRow.Add(23.6)
        End If
        alRow.Add(2)
        alRow.Add(1.75)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(12)
        alRow.Add(12)
        alRow.Add("Rustfri AISI 304, 1.4301 Slebet")
        If prices.Count > 0 Then
            alRow.Add(prices(11))
        Else
            alRow.Add(26.8)
        End If
        alRow.Add(2)
        alRow.Add(1)
        alRow.Add(2)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(13)
        alRow.Add(13)
        alRow.Add("Rustfri AISI 316L, 1.4404 m.PVC")
        If prices.Count > 0 Then
            alRow.Add(prices(12))
        Else
            alRow.Add(33.1)
        End If
        alRow.Add(2)
        alRow.Add(1.75)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(14)
        alRow.Add(14)
        alRow.Add("Kobber")
        If prices.Count > 0 Then
            alRow.Add(prices(13))
        Else
            alRow.Add(93.8)
        End If
        alRow.Add(4)
        alRow.Add(1.75)
        alRow.Add(1)
        alRow.Add(8)
        alReturn.Add(alRow)
        alRow = New ArrayList(15)
        alRow.Add(15)
        alRow.Add("Varmgalvaniseret")
        If prices.Count > 0 Then
            alRow.Add(prices(14))
        Else
            alRow.Add(10)
        End If
        alRow.Add(5)
        alRow.Add(1)
        alRow.Add(1)
        alRow.Add(9)
        alReturn.Add(alRow)

        GetMaterials = alReturn
    End Function

    Private Function retreiveSharepointPrices() As ArrayList
        Dim prices As New ArrayList()
        Dim urlSite As String
        urlSite = "http://intranet/Metalindustri/"
        Dim listName = "Materiale liste"
        Dim trackingList As List
        Dim camlXmlQuery = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                    "<Value Type='Number'>0</Value></Geq></Where></Query><RowLimit>500</RowLimit></View>"

        Dim clientContext As New SP.ClientContext(urlSite)
        trackingList = clientContext.Web.Lists.GetByTitle(listName)
        Dim camlQuery As New CamlQuery()
        camlQuery.ViewXml = camlXmlQuery
        Dim collTrackings As ListItemCollection
        collTrackings = trackingList.GetItems(camlQuery)


        clientContext.Load(collTrackings)
        clientContext.ExecuteQuery()

        For Each item As ListItem In collTrackings

            prices.Add(item("Pris"))

        Next
        retreiveSharepointPrices = prices
    End Function

    'Public Function GetMaterials() As ArrayList
    '    Dim alReturn As ArrayList
    '    Dim alRow As ArrayList

    '    alReturn = New ArrayList(15)
    '    alRow = New ArrayList(7)
    '    alRow.Add(1)
    '    alRow.Add("Jern, P01 AM, 12.03")
    '    alRow.Add(8.5)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(2)
    '    alRow.Add("ALU, 2S Al99")
    '    alRow.Add(25)
    '    alRow.Add(3)
    '    alRow.Add(1)
    '    alRow.Add(2)
    '    alRow.Add(3)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(3)
    '    alRow.Add("ALU, AlMg3, Søvandsbestandig")
    '    alRow.Add(26.8)
    '    alRow.Add(3)
    '    alRow.Add(1)
    '    alRow.Add(2)
    '    alRow.Add(3)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(4)
    '    alRow.Add("Rustfri AISI 304, 1.4301")
    '    alRow.Add(22.4)
    '    alRow.Add(2)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(5)
    '    alRow.Add("Rustfri AISI 316L, 1.4404")
    '    alRow.Add(31.8)
    '    alRow.Add(2)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(6)
    '    alRow.Add("Messing")
    '    alRow.Add(81)
    '    alRow.Add(4)
    '    alRow.Add(1)
    '    alRow.Add(2)
    '    alRow.Add(8.4)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(7)
    '    alRow.Add("Elgalv, Fe P01 ZE, Zintec")
    '    alRow.Add(10.6)
    '    alRow.Add(5)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(8)
    '    alRow.Add("Aluzink, B500A")
    '    alRow.Add(10.6)
    '    alRow.Add(5)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(9)
    '    alRow.Add("ALU, AlMg3, Søvandsbest.m.PVC")
    '    alRow.Add(28.1)
    '    alRow.Add(3)
    '    alRow.Add(1.75)
    '    alRow.Add(2)
    '    alRow.Add(3)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(7)
    '    alRow.Add(10)
    '    alRow.Add("ALU, AlMg1,m.PVC")
    '    alRow.Add(26.3)
    '    alRow.Add(3)
    '    alRow.Add(1.75)
    '    alRow.Add(2)
    '    alRow.Add(3)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(11)
    '    alRow.Add(11)
    '    alRow.Add("Rustfri AISI 304, 1.4301 m.PVC")
    '    alRow.Add(23.6)
    '    alRow.Add(2)
    '    alRow.Add(1.75)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(12)
    '    alRow.Add(12)
    '    alRow.Add("Rustfri AISI 304, 1.4301 Slebet")
    '    alRow.Add(26.8)
    '    alRow.Add(2)
    '    alRow.Add(1)
    '    alRow.Add(2)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(13)
    '    alRow.Add(13)
    '    alRow.Add("Rustfri AISI 316L, 1.4404 m.PVC")
    '    alRow.Add(33.1)
    '    alRow.Add(2)
    '    alRow.Add(1.75)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(14)
    '    alRow.Add(14)
    '    alRow.Add("Kobber")
    '    alRow.Add(93.8)
    '    alRow.Add(4)
    '    alRow.Add(1.75)
    '    alRow.Add(1)
    '    alRow.Add(8)
    '    alReturn.Add(alRow)
    '    alRow = New ArrayList(15)
    '    alRow.Add(15)
    '    alRow.Add("Varmgalvaniseret")
    '    alRow.Add(10)
    '    alRow.Add(5)
    '    alRow.Add(1)
    '    alRow.Add(1)
    '    alRow.Add(9)
    '    alReturn.Add(alRow)

    '    GetMaterials = alReturn
    'End Function
    Public Function GetDifficultclass(ByVal Klasse As Integer, ByVal Thickness As Double) As Double
        Dim Difficult As Integer

        If Klasse = 1 Then
            If Thickness < 1 Then
                Difficult = 3
            End If
            If Thickness >= 1 Then
                Difficult = 1
            End If
            If Thickness >= 2.5 Then
                Difficult = 2
            End If
            If Thickness >= 4 Then
                Difficult = 3
            End If
            If Thickness > 9 Then
                Difficult = 4
            End If
            If Thickness > 15 Then
                Difficult = 5
            End If
        End If

        If Klasse = 2 Then
            If Thickness < 1 Then
                Difficult = 3
            End If
            If Thickness >= 1 Then
                Difficult = 2
            End If
            If Thickness >= 2.5 Then
                Difficult = 3
            End If
            If Thickness >= 4 Then
                Difficult = 4
            End If
            If Thickness > 9 Then
                Difficult = 5
            End If
            If Thickness > 15 Then
                Difficult = 6
            End If
        End If

        GetDifficultclass = Difficult

    End Function
    Public Function GetDifficultmultiplier(ByVal difficultclass As Integer) As Double
        Dim multiplier As Double

        If difficultclass = 1 Then
            multiplier = 1
        End If
        If difficultclass = 2 Then
            multiplier = 1.25
        End If
        If difficultclass = 3 Then
            multiplier = 1.5
        End If
        If difficultclass = 4 Then
            multiplier = 1.8
        End If
        If difficultclass = 5 Then
            multiplier = 2
        End If
        If difficultclass = 6 Then
            multiplier = 2.2
        End If

        GetDifficultmultiplier = multiplier
    End Function
    Public Function GetMaterialWeight(ByVal matrgruppe As Integer) As Double

        Select Case matrgruppe
            Case 1
                GetMaterialWeight = 8

            Case 2
                GetMaterialWeight = 8

            Case 3
                GetMaterialWeight = 2.7

            Case 4
                GetMaterialWeight = 8.4

            Case 5
                GetMaterialWeight = 8

        End Select

    End Function

    Public Function GetPunchStartup() As Double
        GetPunchStartup = (240 + 630 + 90 + 60 + 30 + 120) / 60
    End Function


    Public Function GetLaserStartup(ByVal Thickness As Double) As Double
        'GetLaserStartup = (180 + 90 + 120) / 60

        If Thickness < 7 Then
            GetLaserStartup = 15
        End If
        If Thickness >= 7 Then
            GetLaserStartup = 15
        End If
        If Thickness >= 10 Then
            GetLaserStartup = 15
        End If
        If Thickness > 15 Then
            GetLaserStartup = 30
        End If
        If Thickness > 20 Then
            GetLaserStartup = 60
        End If
    End Function

    Public Function GetCombiStartup() As Double
        '18,5min
        GetCombiStartup = (240 + 630 + 90 + 60 + 30 + 120) / 60
    End Function

    Public Function GetLaserTypeThicknessFactor(ByVal iMaterialType As Integer, ByVal dblThickness As Double) As Double
        GetLaserTypeThicknessFactor = 1.25
    End Function
    Public Function gettimeprcut(ByVal modulesize As Integer) As Double
        If modulesize <= 1 Then
            gettimeprcut = 2.5
            Exit Function
        End If

        If modulesize <= 2 Then
            gettimeprcut = 4
            Exit Function
        End If

        If modulesize <= 3 Then
            gettimeprcut = 8
            Exit Function
        End If

        If modulesize <= 4 Then
            gettimeprcut = 12
            Exit Function
        End If

        If modulesize <= 5 Then
            gettimeprcut = 18
            Exit Function
        End If

        If modulesize <= 6 Then
            gettimeprcut = 30
            Exit Function
        End If

        If modulesize <= 7 Then
            gettimeprcut = 40
            Exit Function
        End If
    End Function

    Public Function GetThicknessmultiplier(ByVal sheetthickness As Double) As Integer

        GetThicknessmultiplier = 1

        If sheetthickness >= 4 Then
            GetThicknessmultiplier = 2
            Exit Function
        End If

        If sheetthickness > 8 Then
            GetThicknessmultiplier = 3
            Exit Function
        End If

    End Function


    Public Function GetSuppliers() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(21)
        alRow = New ArrayList(2)
        alRow.Add(0)
        alRow.Add("Ingen Overfladebehandling")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Laduco")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Miltona")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Næstved Pulverlakering")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("Jowis")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(5)
        alRow.Add("Greve Pulverlakering")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(6)
        alRow.Add("Cuwi")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(7)
        alRow.Add("Stjernecrom")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(8)
        alRow.Add("Croma")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(9)
        alRow.Add("Værløse Galvaniske")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(10)
        alRow.Add("Aluscan")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(11)
        alRow.Add("CGB")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(12)
        alRow.Add("Dan-Pol Serigrafi")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(13)
        alRow.Add("Mekalod Serigrafi")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(14)
        alRow.Add("Gladsaxe Metalsliberi")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(15)
        alRow.Add("Letech")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(16)
        alRow.Add("Dansk Indistricoating")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(17)
        alRow.Add("Karlslunde Metalindistri")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(18)
        alRow.Add("Unicoat")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(19)
        alRow.Add("Elplatek")
        alReturn.Add(alRow)
        GetSuppliers = alReturn
        alRow = New ArrayList(2)
        alRow.Add(20)
        alRow.Add("DOT")
        alReturn.Add(alRow)
        GetSuppliers = alReturn

    End Function

    Public Function GetSpotweldspeed(ByVal sheetthickness As Double, ByVal Matrgroup As Integer) As Double
        Dim weldspeed As Double
        'jern minutter pr punktsvejsning
        If Matrgroup = 1 Then
            If sheetthickness < 2 Then
                weldspeed = 0.05
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.05
            End If
            If sheetthickness >= 3 Then
                weldspeed = 0.06
            End If
            If sheetthickness >= 4 Then
                weldspeed = 0.07
            End If
            If sheetthickness >= 5 Then
                weldspeed = 0.09
            End If
            If sheetthickness = 6 Then
                weldspeed = 0.1
            End If
            'If sheetthickness >= 8 Then
            'weldspeed = 0.3
            'End If
            'If sheetthickness >= 10 Then
            'weldspeed = 0.4
            'End If
        End If
        'rustfri
        If Matrgroup = 2 Then
            If sheetthickness < 2 Then
                weldspeed = 0.05
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.05
            End If
            If sheetthickness >= 3 Then
                weldspeed = 0.06
            End If
            If sheetthickness >= 4 Then
                weldspeed = 0.07
            End If
            If sheetthickness >= 5 Then
                weldspeed = 0.09
            End If
            If sheetthickness = 6 Then
                weldspeed = 0.1
            End If
            'If sheetthickness >= 8 Then
            'weldspeed = 0.3
            'End If
            'If sheetthickness >= 10 Then
            'weldspeed = 0.4
            'End If
        End If
        'alu
        If Matrgroup = 3 Then
            If sheetthickness < 2 Then
                weldspeed = 0.2
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.2
            End If
            If sheetthickness >= 3 Then
                weldspeed = 0.3
            End If
            If sheetthickness = 4 Then
                weldspeed = 0.4
            End If
            'If sheetthickness >= 5 Then
            'weldspeed = 0.5
            'End If
            'If sheetthickness >= 6 Then
            'weldspeed = 0.6
            ' End If
            'If sheetthickness >= 8 Then
            'weldspeed = 0.7
            'End If
            'If sheetthickness >= 10 Then
            'weldspeed = 0.8
            'End If
        End If
        GetSpotweldspeed = weldspeed

    End Function

    Public Function GetSpotWeldingShift(ByVal modulesize As Integer) As Double
        Dim subjectshift As Double

        If modulesize = 1 Then
            subjectshift = 0.5
        End If
        If modulesize = 2 Then
            subjectshift = 1
        End If
        If modulesize = 3 Then
            subjectshift = 1.5
        End If
        If modulesize = 4 Then
            subjectshift = 2
        End If
        If modulesize = 5 Then
            subjectshift = 3
        End If
        If modulesize = 6 Then
            subjectshift = 8
        End If
        If modulesize = 7 Then
            subjectshift = 10
        End If

        GetSpotWeldingShift = subjectshift
    End Function
    Public Function GetSpotWeldingStartstop(ByVal modulesize As Integer) As Double
        Dim weldStartstop As Double

        If modulesize = 1 Then
            weldStartstop = 0.1
        End If
        If modulesize = 2 Then
            weldStartstop = 0.2
        End If
        If modulesize = 3 Then
            weldStartstop = 0.3
        End If
        If modulesize = 4 Then
            weldStartstop = 0.5
        End If
        If modulesize = 5 Then
            weldStartstop = 1
        End If
        If modulesize = 6 Then
            weldStartstop = 2
        End If
        If modulesize = 7 Then
            weldStartstop = 3
        End If

        GetSpotWeldingStartstop = weldStartstop



    End Function

    Public Function GetweldspeedTIG100(ByVal sheetthickness As Double, ByVal Matrgroup As Integer) As Double
        Dim weldspeed As Double
        'jern 100mm svejsning
        If Matrgroup = 1 Then
            If sheetthickness < 2 Then
                weldspeed = 1
            End If
            If sheetthickness >= 2 Then
                weldspeed = 1
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 5 Then
                weldspeed = 2
            End If
            If sheetthickness >= 6 Then
                weldspeed = 3
            End If
            If sheetthickness >= 8 Then
                weldspeed = 4
            End If
            If sheetthickness >= 10 Then
                weldspeed = 6
            End If
        End If
        'rustfri
        If Matrgroup = 2 Then
            If sheetthickness < 2 Then
                weldspeed = 1
            End If
            If sheetthickness >= 2 Then
                weldspeed = 1
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 5 Then
                weldspeed = 2
            End If
            If sheetthickness >= 6 Then
                weldspeed = 3
            End If
            If sheetthickness >= 8 Then
                weldspeed = 4
            End If
            If sheetthickness >= 10 Then
                weldspeed = 6
            End If
        End If
        'alu
        If Matrgroup = 3 Then
            If sheetthickness < 2 Then
                weldspeed = 2
            End If
            If sheetthickness >= 2 Then
                weldspeed = 2
            End If
            If sheetthickness >= 3 Then
                weldspeed = 2.5
            End If
            If sheetthickness >= 4 Then
                weldspeed = 3
            End If
            If sheetthickness >= 5 Then
                weldspeed = 5
            End If
            If sheetthickness >= 6 Then
                weldspeed = 7
            End If
            If sheetthickness >= 8 Then
                weldspeed = 8
            End If
            If sheetthickness >= 10 Then
                weldspeed = 9
            End If
        End If
        GetweldspeedTIG100 = weldspeed
    End Function
    Public Function GetweldspeedMIG100(ByVal sheetthickness As Double, ByVal Matrgroup As Integer) As Double
        Dim weldspeed As Double
        'jern 100mm svejsning
        If Matrgroup = 1 Then
            If sheetthickness < 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1.25
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1.25
            End If
            If sheetthickness >= 5 Then
                weldspeed = 2
            End If
            If sheetthickness >= 6 Then
                weldspeed = 3
            End If
            If sheetthickness >= 8 Then
                weldspeed = 3
            End If
            If sheetthickness >= 10 Then
                weldspeed = 4
            End If
        End If
        'rustfri
        If Matrgroup = 2 Then
            If sheetthickness < 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1.25
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1.25
            End If
            If sheetthickness >= 5 Then
                weldspeed = 1.75
            End If
            If sheetthickness >= 6 Then
                weldspeed = 3
            End If
            If sheetthickness >= 8 Then
                weldspeed = 3
            End If
            If sheetthickness >= 10 Then
                weldspeed = 4
            End If
        End If
        'alu
        If Matrgroup = 3 Then
            If sheetthickness < 2 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 2 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 3 Then
                weldspeed = 2
            End If
            If sheetthickness >= 4 Then
                weldspeed = 2.5
            End If
            If sheetthickness >= 5 Then
                weldspeed = 3
            End If
            If sheetthickness >= 6 Then
                weldspeed = 4
            End If
            If sheetthickness >= 8 Then
                weldspeed = 4
            End If
            If sheetthickness >= 10 Then
                weldspeed = 5
            End If
        End If
        GetweldspeedMIG100 = weldspeed

    End Function
    Public Function GettackweldspeedTIG100(ByVal sheetthickness As Double, ByVal Matrgroup As Integer) As Double
        Dim weldspeed As Double
        'jern 100mm svejsning
        If Matrgroup = 1 Then
            If sheetthickness < 2 Then
                weldspeed = 0.5
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 5 Then
                weldspeed = 2
            End If
            If sheetthickness >= 6 Then
                weldspeed = 2.5
            End If
            If sheetthickness >= 8 Then
                weldspeed = 3
            End If
            If sheetthickness >= 10 Then
                weldspeed = 3
            End If
        End If
        'rustfri
        If Matrgroup = 2 Then
            If sheetthickness < 2 Then
                weldspeed = 0.5
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 5 Then
                weldspeed = 2
            End If
            If sheetthickness >= 6 Then
                weldspeed = 2.5
            End If
            If sheetthickness >= 8 Then
                weldspeed = 3
            End If
            If sheetthickness >= 10 Then
                weldspeed = 3
            End If
        End If
        'alu
        If Matrgroup = 3 Then
            If sheetthickness < 2 Then
                weldspeed = 1
            End If
            If sheetthickness >= 2 Then
                weldspeed = 1
            End If
            If sheetthickness >= 3 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 4 Then
                weldspeed = 2
            End If
            If sheetthickness >= 5 Then
                weldspeed = 3
            End If
            If sheetthickness >= 6 Then
                weldspeed = 3.5
            End If
            If sheetthickness >= 8 Then
                weldspeed = 4
            End If
            If sheetthickness >= 10 Then
                weldspeed = 4
            End If
        End If
        GettackweldspeedTIG100 = weldspeed

    End Function
    Public Function GettackweldspeedMIG100(ByVal sheetthickness As Double, ByVal Matrgroup As Integer) As Double
        Dim weldspeed As Double
        'jern 100mm svejsning
        If Matrgroup = 1 Then
            If sheetthickness < 2 Then
                weldspeed = 0.5
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.5
            End If
            If sheetthickness >= 3 Then
                weldspeed = 0.5
            End If
            If sheetthickness >= 4 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 5 Then
                weldspeed = 1
            End If
            If sheetthickness >= 6 Then
                weldspeed = 1
            End If
            If sheetthickness >= 8 Then
                weldspeed = 1
            End If
            If sheetthickness >= 10 Then
                weldspeed = 1
            End If
        End If
        'rustfri
        If Matrgroup = 2 Then
            If sheetthickness < 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 3 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1
            End If
            If sheetthickness >= 5 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 6 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 8 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 10 Then
                weldspeed = 1.5
            End If
        End If
        'alu
        If Matrgroup = 3 Then
            If sheetthickness < 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 2 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 3 Then
                weldspeed = 0.75
            End If
            If sheetthickness >= 4 Then
                weldspeed = 1
            End If
            If sheetthickness >= 5 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 6 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 8 Then
                weldspeed = 1.5
            End If
            If sheetthickness >= 10 Then
                weldspeed = 1.5
            End If
        End If
        GettackweldspeedMIG100 = weldspeed


    End Function
    Public Function GetgrindtackweldspeedTIG100(ByVal sheetthickness As Double) As Double
        Dim weldspeed As Double

        If sheetthickness < 2 Then
            weldspeed = 1
        End If
        If sheetthickness >= 2 Then
            weldspeed = 1
        End If
        If sheetthickness >= 3 Then
            weldspeed = 1.5
        End If
        If sheetthickness >= 4 Then
            weldspeed = 2
        End If
        If sheetthickness >= 5 Then
            weldspeed = 2.5
        End If
        If sheetthickness >= 6 Then
            weldspeed = 3
        End If
        If sheetthickness >= 8 Then
            weldspeed = 3
        End If
        If sheetthickness >= 10 Then
            weldspeed = 3.5
        End If

        GetgrindtackweldspeedTIG100 = weldspeed
    End Function
    Public Function GetgrindweldspeedTIG100(ByVal sheetthickness As Double) As Double
        Dim weldspeed As Double

        If sheetthickness < 2 Then
            weldspeed = 2
        End If
        If sheetthickness >= 2 Then
            weldspeed = 2
        End If
        If sheetthickness >= 3 Then
            weldspeed = 3
        End If
        If sheetthickness >= 4 Then
            weldspeed = 3
        End If
        If sheetthickness >= 5 Then
            weldspeed = 4
        End If
        If sheetthickness >= 6 Then
            weldspeed = 5
        End If
        If sheetthickness >= 8 Then
            weldspeed = 6
        End If
        If sheetthickness >= 10 Then
            weldspeed = 6
        End If

        GetgrindweldspeedTIG100 = weldspeed

    End Function
    Public Function GetgrindtackweldspeedMIG100(ByVal sheetthickness As Double) As Double
        Dim weldspeed As Double

        If sheetthickness < 2 Then
            weldspeed = 1.5
        End If
        If sheetthickness >= 2 Then
            weldspeed = 1.5
        End If
        If sheetthickness >= 3 Then
            weldspeed = 1.5
        End If
        If sheetthickness >= 4 Then
            weldspeed = 2
        End If
        If sheetthickness >= 5 Then
            weldspeed = 2.5
        End If
        If sheetthickness >= 6 Then
            weldspeed = 2.5
        End If
        If sheetthickness >= 8 Then
            weldspeed = 3
        End If
        If sheetthickness >= 10 Then
            weldspeed = 3
        End If

        GetgrindtackweldspeedMIG100 = weldspeed

    End Function
    Public Function GetgrindweldspeedMIG100(ByVal sheetthickness As Double) As Double
        Dim weldspeed As Double

        If sheetthickness < 2 Then
            weldspeed = 2
        End If
        If sheetthickness >= 2 Then
            weldspeed = 2
        End If
        If sheetthickness >= 3 Then
            weldspeed = 3
        End If
        If sheetthickness >= 4 Then
            weldspeed = 4
        End If
        If sheetthickness >= 5 Then
            weldspeed = 5
        End If
        If sheetthickness >= 6 Then
            weldspeed = 6
        End If
        If sheetthickness >= 8 Then
            weldspeed = 6
        End If
        If sheetthickness >= 10 Then
            weldspeed = 6.5
        End If

        GetgrindweldspeedMIG100 = weldspeed

    End Function

    Public Function GetWeldingShift(ByVal modulesize As Integer) As Double
        Dim subjectshift As Double

        If modulesize = 1 Then
            subjectshift = 0.17
        End If
        If modulesize = 2 Then
            subjectshift = 0.25
        End If
        If modulesize = 3 Then
            subjectshift = 0.33
        End If
        If modulesize = 4 Then
            subjectshift = 3
        End If
        If modulesize = 5 Then
            subjectshift = 12
        End If
        If modulesize = 6 Then
            subjectshift = 15
        End If
        If modulesize = 7 Then
            subjectshift = 15
        End If

        GetWeldingShift = subjectshift * 0.5
    End Function
    Public Function GetWeldingStartstop(ByVal modulesize As Integer) As Double
        Dim weldStartstop As Double

        If modulesize = 1 Then
            weldStartstop = 0.05
        End If
        If modulesize = 2 Then
            weldStartstop = 0.05
        End If
        If modulesize = 3 Then
            weldStartstop = 0.13
        End If
        If modulesize = 4 Then
            weldStartstop = 0.13
        End If
        If modulesize = 5 Then
            weldStartstop = 0.17
        End If
        If modulesize = 6 Then
            weldStartstop = 0.33
        End If
        If modulesize = 7 Then
            weldStartstop = 0.33
        End If

        GetWeldingStartstop = weldStartstop * 0.5



    End Function

    Public Function GetSurfacetreatment() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(6)
        alRow = New ArrayList(2)
        alRow.Add(0)
        alRow.Add("Ingen Overfladebehandling")
        alReturn.Add(alRow)
        GetSurfacetreatment = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Vådlak")
        alReturn.Add(alRow)
        GetSurfacetreatment = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Pulverlak")
        alReturn.Add(alRow)
        GetSurfacetreatment = alReturn
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Chromit (jern)")
        alReturn.Add(alRow)
        GetSurfacetreatment = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("ChromitAL (alu)")
        alReturn.Add(alRow)
        GetSurfacetreatment = alReturn
        alRow = New ArrayList(2)
        alRow.Add(5)
        alRow.Add("Eloxering (natur/sort)")
        alReturn.Add(alRow)
        GetSurfacetreatment = alReturn


    End Function
    Public Function GetSurfacetreatment1() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(12)
        alRow = New ArrayList(2)
        alRow.Add(0)
        alRow.Add("Ingen Overfladebehandling")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Vådlak")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Pulverlak")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Chromit (jern)")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("ChromitAL (alu)")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(5)
        alRow.Add("Eloxering (natur/sort)")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(6)
        alRow.Add("Silketryk")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(7)
        alRow.Add("Fornikling/Fortinning")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(8)
        alRow.Add("Black Oxide (brunering)")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(9)
        alRow.Add("Glasblæsning")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(10)
        alRow.Add("Afsyring")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
        alRow = New ArrayList(2)
        alRow.Add(11)
        alRow.Add("Varmgalvanisering")
        alReturn.Add(alRow)
        GetSurfacetreatment1 = alReturn
    End Function

    Public Function GetSupplierspulver() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList

        alReturn = New ArrayList(7)
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Næstved Pulverlakering")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Laduco")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Miltona")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("Jowis")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn
        alRow = New ArrayList(2)
        alRow.Add(5)
        alRow.Add("Greve Pulverlakering")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn
        alRow = New ArrayList(2)
        alRow.Add(6)
        alRow.Add("Cuwi")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn
        alRow = New ArrayList(2)
        alRow.Add(7)
        alRow.Add("MalerFlex")
        alReturn.Add(alRow)
        GetSupplierspulver = alReturn

    End Function
    Public Function GetSuppliersvåd() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(5)
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Næstved Pulverlakering")
        alReturn.Add(alRow)
        GetSuppliersvåd = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Laduco")
        alReturn.Add(alRow)
        GetSuppliersvåd = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Miltona")
        alReturn.Add(alRow)
        GetSuppliersvåd = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("Jowis")
        alReturn.Add(alRow)
        GetSuppliersvåd = alReturn
        alRow = New ArrayList(2)
        alRow.Add(5)
        alRow.Add("MalerFlex")
        alReturn.Add(alRow)
        GetSuppliersvåd = alReturn

    End Function

    Public Function GetSupplierschromitAL() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(4)
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Værløse Galvaniske")
        alReturn.Add(alRow)
        GetSupplierschromitAL = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Stjernecrom")
        alReturn.Add(alRow)
        GetSupplierschromitAL = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Croma")
        alReturn.Add(alRow)
        GetSupplierschromitAL = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("Aluscan")
        alReturn.Add(alRow)
        GetSupplierschromitAL = alReturn

    End Function

    Public Function GetSupplierschromit() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(3)
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Værløse Galvaniske")
        alReturn.Add(alRow)
        GetSupplierschromit = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Stjernecrom")
        alReturn.Add(alRow)
        GetSupplierschromit = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Croma")
        alReturn.Add(alRow)
        GetSupplierschromit = alReturn

    End Function

    Public Function GetSuppliereloxering() As ArrayList
        Dim alReturn As ArrayList
        Dim alRow As ArrayList


        alReturn = New ArrayList(4)
        alRow = New ArrayList(2)
        alRow.Add(3)
        alRow.Add("Værløse Galvaniske")
        alReturn.Add(alRow)
        GetSuppliereloxering = alReturn
        alRow = New ArrayList(2)
        alRow.Add(1)
        alRow.Add("Stjernecrom")
        alReturn.Add(alRow)
        GetSuppliereloxering = alReturn
        alRow = New ArrayList(2)
        alRow.Add(2)
        alRow.Add("Croma")
        alReturn.Add(alRow)
        GetSuppliereloxering = alReturn
        alRow = New ArrayList(2)
        alRow.Add(4)
        alRow.Add("Aluscan")
        alReturn.Add(alRow)
        GetSuppliereloxering = alReturn

    End Function


    'Public Sub GetPulverlakData(ByVal supplyer As String, ByVal start As Double, ByVal coverholes As Double, ByVal m2price As Double)

    'If supplyer = "Laduco" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    ' End If
    'If supplyer = "Milling Industrilakering" Then
    'start = 500
    'coverholes = 2
    ' m2price = 120
    ' End If
    'If supplyer = "Næstved Pulverlakering" Then
    'start = 82
    'coverholes = 1.6
    ' m2price = 86
    ' End If
    'If supplyer = "Jowis" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    'End If
    'If supplyer = "Greve Pulverlakering" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    'End If
    'If supplyer = "Cuwi" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    'End If
    'GetPulverlakData = 0
    ' End Sub
    'Public Function GetVådlakData(ByVal supplyer As String, ByVal start As Double, ByVal coverholes As Double, ByVal m2price As Double) As Double

    'If supplyer = "Laduco" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 360
    'End If
    'If supplyer = "Milling Industrilakering" Then
    'start = 500
    'coverholes = 2
    'm2price = 360
    'End If
    'If supplyer = "Næstved Pulverlakering" Then
    'start = 82
    'coverholes = 1.6
    'm2price = 260
    ' End If
    'If supplyer = "Jowis" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    'End If
    'If supplyer = "Greve Pulverlakering" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    'End If
    'If supplyer = "Cuwi" Then
    'start = 400
    'coverholes = 1.25
    'm2price = 120
    'End If
    'End Function

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

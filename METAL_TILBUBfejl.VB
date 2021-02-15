Public Class metal_tilbud
    Private pobjListMaterials As ListMaterials
    Private pobjListSuppliers1 As ListSuppliers
    Private pobjListSuppliers2 As ListSuppliers
    Private pobjListSuppliers3 As ListSuppliers
    'Private pobjListSuppliers4 As ListSuppliers
    Private pobjListSurface1 As Listsurfaces
    Private pobjListSurface2 As Listsurfaces
    Private pobjListSurface3 As Listsurfaces
    'Private pobjListSurface4 As Listsurfaces

    Private pobjFilter_tb_tilbud1 As FilterKeys
    Private pobjFilter_tb_tilbud2 As FilterKeys
    Private pobjFilter_tb_tilbud3 As FilterKeys
    Private pobjFilter_tb_tilbud4 As FilterKeys
    private pobjFilter_tb_avance as FilterKeys 
    'buk
    Private pobjFilter_tb_tilbud5 As FilterKeys
    Private pobjFilter_tb_bukmax_x As FilterKeys
    Private pobjFilter_tb_bukmax_Y As FilterKeys
    Private pobjFilter_tb_buk1_x As FilterKeys
    Private pobjFilter_tb_buk1_Y As FilterKeys
    Private pobjFilter_tb_buk2_x As FilterKeys
    Private pobjFilter_tb_buk2_Y As FilterKeys
    Private pobjFilter_tb_buk3_x As FilterKeys
    Private pobjFilter_tb_buk3_Y As FilterKeys
    Private pobjFilter_tb_buk4_x As FilterKeys
    Private pobjFilter_tb_buk4_Y As FilterKeys
    Private pobjFilter_tb_buk5_x As FilterKeys
    Private pobjFilter_tb_buk5_Y As FilterKeys
    Private pobjFilter_tb_buk6_x As FilterKeys
    Private pobjFilter_tb_buk6_Y As FilterKeys
    Private pobjFilter_tb_buk7_x As FilterKeys
    Private pobjFilter_tb_buk7_Y As FilterKeys
    Private pobjFilter_tb_buk8_x As FilterKeys
    Private pobjFilter_tb_buk8_Y As FilterKeys
    Private pobjFilter_tb_buk9_x As FilterKeys
    Private pobjFilter_tb_buk9_Y As FilterKeys
    Private pobjFilter_tb_buk10_x As FilterKeys
    Private pobjFilter_tb_buk10_Y As FilterKeys
    Private pobjFilter_tb_buk11_x As FilterKeys
    Private pobjFilter_tb_buk11_Y As FilterKeys
    Private pobjFilter_tb_buk_uk As FilterKeys
    Private pobjFilter_tb_buk_opst_uk As FilterKeys
    'stans
    Private pobjFilter_tb_toolshift As FilterKeys
    Private pobjFilter_tb_slag_til_huller As FilterKeys
    Private pobjFilter_tb_CNCmin_uk As FilterKeys
    Private pobjFilter_tb_gruppe1_opstart_uk As FilterKeys
    'laser
    Private pobjFilter_tb_hulantal_1C As FilterKeys
    Private pobjFilter_tb_hulantal_2C As FilterKeys
    Private pobjFilter_tb_hulantal_3C As FilterKeys
    Private pobjFilter_tb_cuttinglength_C As FilterKeys
    Private pobjFilter_tb_laserCNC_tid_uk As FilterKeys
    Private pobjFilter_tb_laser_opstart_uk As FilterKeys
    'kombi
    Private pobjFilter_tb_toolshift_B As FilterKeys
    Private pobjFilter_tb_slag_til_huller_B As FilterKeys
    Private pobjFilter_tb_hulantal_1B As FilterKeys
    Private pobjFilter_tb_hulantal_2B As FilterKeys
    Private pobjFilter_tb_hulantal_3B As FilterKeys
    Private pobjFilter_tb_cuttinglength_B As FilterKeys
    Private pobjFilter_tb_combiCNC_tid_uk As FilterKeys
    Private pobjFilter_tb_combi_opstart_uk As FilterKeys
    'klip
    Private pobjFilter_tb_klip_tid_uk As FilterKeys
    Private pobjFilter_tb_klip_opstart_uk As FilterKeys
    'gruppe2
    Private pobjFilter_tb_afgrat_uk As FilterKeys
    Private pobjFilter_tb_grinding_uk As FilterKeys
    Private pobjFilter_tb_brush_uk As FilterKeys
    Private pobjFilter_tb_vibration_uk As FilterKeys
    Private pobjFilter_tb_rette_uk As FilterKeys
    Private pobjFilter_tb_stans_manuel_uk As FilterKeys
    Private pobjFilter_tb_countersink_uk As FilterKeys
    Private pobjFilter_tb_gevind_uk As FilterKeys
    Private pobjFilter_tb_pressnut_uk As FilterKeys
    Private pobjFilter_tb_pressnut_kr_uk As FilterKeys
    Private pobjFilter_tb_svejsestag_kr_uk As FilterKeys
    Private pobjFilter_tb_boltesvejs_uk As FilterKeys
    Private pobjFilter_tb_m2 As FilterKeys
    Private pobjFilter_tb_m2_5 As FilterKeys
    Private pobjFilter_tb_m3 As FilterKeys
    Private pobjFilter_tb_m4 As FilterKeys
    Private pobjFilter_tb_m5 As FilterKeys
    Private pobjFilter_tb_m6 As FilterKeys
    Private pobjFilter_tb_m8 As FilterKeys
    Private pobjFilter_tb_m10 As FilterKeys
    Private pobjFilter_tb_1 As FilterKeys
    Private pobjFilter_tb_2 As FilterKeys
    Private pobjFilter_tb_3 As FilterKeys
    Private pobjFilter_tb_4 As FilterKeys
    'gruppe3
    Private pobjFilter_tb_numberofspotwelds As FilterKeys
    Private pobjFilter_tb_spotweld_uk As FilterKeys
    'gruppe4
    Private pobjFilter_tb_numberofwelds As FilterKeys
    Private pobjFilter_tb_weldlength As FilterKeys
    Private pobjFilter_tb_tackweld_uk As FilterKeys
    Private pobjFilter_tb_weld_uk As FilterKeys
    Private pobjFilter_tb_grind_weld_uk As FilterKeys
    'gruppe5
    Private pobjFilter_tb_kontor_uk As FilterKeys
    Private pobjFilter_tb_kontrol_uk As FilterKeys
    'admin
    Private pobjFilter_tb_timesats_mand As FilterKeys
    Private pobjFilter_tb_timesats_D As FilterKeys
    Private pobjFilter_tb_timesats_B As FilterKeys
    Private pobjFilter_tb_timesats_C As FilterKeys
    'opstart
    Private pobjFilter_tb_antal_opstart_uk As FilterKeys
    Private pobjFilter_tb_opstart_kr_uk As FilterKeys
    Private pobjFilter_tb_antal_program_uk As FilterKeys
    Private pobjFilter_tb_program_kr_uk As FilterKeys
    Private pobjFilter_tb_opstart_avance As FilterKeys
    Private pobjFilter_tb_opstart_afgivettilbud As FilterKeys
    Private pobjFilter_tb_pladetykkelse As FilterKeys
    Private pobjFilter_tb_sværhed_uk As FilterKeys
    'indkøb
    Private pobjFilter_tb_overfl_opstart1 As FilterKeys
    Private pobjFilter_tb_overfl_opstart2 As FilterKeys
    Private pobjFilter_tb_overfl_opstart3 As FilterKeys
    Private pobjFilter_tb_overfl_opstart4 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk1 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk2 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk3 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk4 As FilterKeys
    Private pobjFilter_tb_overfl_Pris1 As FilterKeys
    Private pobjFilter_tb_overfl_Pris2 As FilterKeys
    Private pobjFilter_tb_overfl_Pris3 As FilterKeys
    Private pobjFilter_tb_overfl_Pris4 As FilterKeys
    Private pobjFilter_tb_overfl_Avance1 As FilterKeys
    Private pobjFilter_tb_overfl_Avance2 As FilterKeys
    Private pobjFilter_tb_overfl_Avance3 As FilterKeys
    Private pobjFilter_tb_overfl_Avance4 As FilterKeys

    Private pbLoading As Boolean


    Private Sub metal_tilbud_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FillCustomerNames()
    End Sub


    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_Reset.Click


    'samlet_tilbud.Hide()


    'samlet_tilbud.Show()


    'End Sub

    'Private Sub Label81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label81.Click

    'End Sub

    Private Sub FillCustomerNames()
        Dim objodbcconn As Odbc.OdbcConnection
        Dim objodbccom As Odbc.OdbcCommand
        Dim objodbcReader As Odbc.OdbcDataReader


        objodbcconn = New Odbc.OdbcConnection
        objodbcconn.ConnectionString = "DSN=MetalTilbud"
        objodbccom = New Odbc.OdbcCommand
        objodbccom.CommandText = "SELECT ID, Name FROM Customer"
        objodbcconn.Open()
        objodbccom.Connection = objodbcconn
        objodbcReader = objodbccom.ExecuteReader()
        While objodbcReader.Read() = True
            cb_kunde.Items.Add(objodbcReader.Item(1).ToString())
        End While
    End Sub

    Private Sub cb_kunde_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_kunde.SelectedIndexChanged
        UpdateDrawings()
    End Sub

    Private Sub UpdateDrawings()
        Dim objodbcconn As Odbc.OdbcConnection
        Dim objodbccom As Odbc.OdbcCommand
        Dim objodbcReader As Odbc.OdbcDataReader

        'Dim i As Long


        objodbcconn = New Odbc.OdbcConnection
        objodbcconn.ConnectionString = "DSN=MetalTilbud"
        objodbccom = New Odbc.OdbcCommand
        objodbccom.CommandText = "SELECT DrawingNumber FROM Offer, Customer WHERE Offer.CustID=Customer.ID AND Customer.Name='" + cb_kunde.Text + "'"
        objodbcconn.Open()
        objodbccom.Connection = objodbcconn
        objodbcReader = objodbccom.ExecuteReader()
        cb_Tegning.Items.Clear()
        While objodbcReader.Read() = True
            cb_Tegning.Items.Add(objodbcReader.Item(0).ToString())
        End While
    End Sub

    'GEVIND

    '-----------------------------------------------------------------------------------------------------------------

    'Private Sub bu_gev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'lb_gevind.Text = CalculateThreads(tb_antal1.Text)

    'End Sub

    Private Function CalculateThreads(ByVal number As Integer) As Double
        'Dim fthread_input As thread_input
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim totaltime As Double
        Dim countthread As Integer
        Dim modulesize As Integer

        totaltime = 10


        modulesize = Val(lb_modulstr.Text)

        objDatabaseIO = New DatabaseIO
        'fthread_input = New thread_input
        'fthread_input.ShowDialog(Me)

        countthread = (Val(tb_m2.Text) + Val(tb_m2_5.Text) + Val(tb_m3.Text) + Val(tb_m4.Text) + Val(tb_m5.Text) + Val(tb_m6.Text) + Val(tb_m8.Text) + Val(tb_m10.Text))

        If countthread = 0 Then
            lb_gevind.Text = ""
            lb_gevind_antal.Text = ""
            Exit Function
        End If

        If tb_m2.Text <> "" Then
            'If fthread_input.tb_m2.Text <> "" Then
            totaltime = totaltime + 2
        End If
        If tb_m2_5.Text <> "" Then
            'If fthread_input.tb_m2_5.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m3.Text <> "" Then
            'If fthread_input.tb_m3.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m4.Text <> "" Then
            'If fthread_input.tb_m4.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m5.Text <> "" Then
            'If fthread_input.tb_m5.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m6.Text <> "" Then
            'If fthread_input.tb_m6.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m8.Text <> "" Then
            'If fthread_input.tb_m8.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m10.Text <> "" Then
            'If fthread_input.tb_m10.Text <> "" Then
            totaltime = totaltime + 4
        End If

        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m2.Text))) * 2) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m2.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m2_5.Text))) * 2) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m2_5.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m3.Text))) * 2.5) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m3.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m4.Text))) * 3) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m4.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m5.Text))) * 3.5) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m5.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m6.Text))) * 4) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m6.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m8.Text))) * 4.5) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m8.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatethread(Val(tb_m10.Text))) * 5) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_m10.Text))) / 60)

        If lb_faktor.Text = "" Then
            Exit Function
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        CalculateThreads = totaltime * Difficultfactor

        lb_gevind.Text = CalculateThreads
        lb_gevind_antal.Text = countthread

    End Function

    Private Function calculatethread(ByVal count As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim materialID As Integer
        Dim sheetthickness As Integer

        If tb_pladetykkelse.Text = "" Then
            Exit Function
        End If

        sheetthickness = tb_pladetykkelse.Text
        materialID = Val(Lb_matrgruppe.Text)

        objDatabaseIO = New DatabaseIO


        calculatethread = (3 * (count * objDatabaseIO.getthreadmultiplier(materialID) * (1 + (0.111 * (sheetthickness - 1)))))



    End Function


    'UNDERSÆNKNING

    '--------------------------------------------------------------------------------------------------------------
    'Private Sub bu_unders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'tb_countersink_uk.Text = ""
    'lb_countersink.Text = CalculateCountersinks(tb_antal1.Text)
    'End Sub

    Private Function CalculateCountersinks(ByVal number As Integer) As Double
        'Dim fcountersink_input As countersink_input
        Dim objDatabaseIO As DatabaseIO
        Dim countcountersink As Integer
        Dim totaltime As Double
        Dim modulesize As Integer
        Dim Difficultfactor As Double
        Dim Countersinks As Double

        countcountersink = ((Val(tb_1.Text) + Val(tb_2.Text) + Val(tb_3.Text) + Val(tb_4.Text)))
        lb_undersænk_antal.Text = countcountersink

        If countcountersink = 0 Then
            lb_countersink.Text = ""
            lb_undersænk_antal.Text = ""
            Exit Function
        End If


        totaltime = 10

        modulesize = Val(lb_modulstr.Text)

        objDatabaseIO = New DatabaseIO
        'fcountersink_input = New countersink_input
        'fcountersink_input.ShowDialog(Me)
        If tb_1.Text <> "" Then
            'If fcountersink_input.tb_1.Text <> "" Then
            totaltime = totaltime + 2
        End If
        If tb_2.Text <> "" Then
            'If fcountersink_input.tb_2.Text <> "" Then
            totaltime = totaltime + 5
        End If
        If tb_3.Text <> "" Then
            'If fcountersink_input.tb_3.Text <> "" Then
            totaltime = totaltime + 5
        End If
        If tb_4.Text <> "" Then
            'If fcountersink_input.tb_4.Text <> "" Then
            totaltime = totaltime + 5
        End If

        totaltime = totaltime + ((number * (((calculatecountersink(Val(tb_1.Text))) * 1) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_1.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatecountersink(Val(tb_2.Text))) * 1.5) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_2.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatecountersink(Val(tb_3.Text))) * 2) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_3.Text))) / 60)
        totaltime = totaltime + ((number * (((calculatecountersink(Val(tb_4.Text))) * 3) + (objDatabaseIO.gettimeprcut(modulesize)) * Val(tb_4.Text))) / 60)

        If lb_faktor.Text = "" Then
            Exit Function
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        Countersinks = totaltime * Difficultfactor

        lb_countersink.Text = Countersinks

    End Function

    Private Function calculatecountersink(ByVal count As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim materialID As Integer

        materialID = Val(Lb_matrgruppe.Text)

        objDatabaseIO = New DatabaseIO
        If count = 0 Then
            calculatecountersink = 0
        Else
            calculatecountersink = (5 * (count * objDatabaseIO.getcountersinkmultiplier(materialID)))
        End If

    End Function

    Private Function CalculatePressnut(ByVal iNumberOfSubjects As Integer, ByVal iNumberOfNuts As Integer) As Double
        Dim objPressnutCalculations As PressnutCalculations
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        If iNumberOfNuts = 0 Then
            lb_pressnut.Text = ""
            tb_pressnut_uk.Text = ""
            lb_pressnut_kr.Text = ""
            tb_pressnut_kr_uk.Text = ""
            Exit Function
        End If

        objPressnutCalculations = New PressnutCalculations
        objDatabaseIO = New DatabaseIO

        CalculatePressnut = (objPressnutCalculations.CalculatePressnutTime(lb_modulstr.Text, iNumberOfNuts) / 100 * iNumberOfSubjects) + objDatabaseIO.GetPressnutStartup

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        lb_pressnut.Text = CalculatePressnut * Difficultfactor

        objDatabaseIO = New DatabaseIO
        lb_pressnut_kr.Text = objDatabaseIO.getpressnutprice(Val(Lb_matrgruppe.Text)) * iNumberOfNuts

    End Function
    Private Function CalculateBoltweld(ByVal iNumberOfSubjects As Integer, ByVal iNumberOfBolts As Integer) As Double
        Dim objPressnutCalculations As PressnutCalculations
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        If iNumberOfBolts = 0 Then
            lb_boltesvejs.Text = ""
            tb_boltesvejs_uk.Text = ""
            lb_svejsestag_kr.Text = ""
            tb_svejsestag_kr_uk.Text = ""
            Exit Function
        End If

        objPressnutCalculations = New PressnutCalculations
        objDatabaseIO = New DatabaseIO

        CalculateBoltweld = (objPressnutCalculations.CalculateBoltweldTime(lb_modulstr.Text, iNumberOfBolts) / 100 * iNumberOfSubjects) + objDatabaseIO.GetBoltweldingStartup(iNumberOfBolts)

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        lb_boltesvejs.Text = CalculateBoltweld * Difficultfactor

        objDatabaseIO = New DatabaseIO
        lb_svejsestag_kr.Text = objDatabaseIO.getboltprice(Val(Lb_matrgruppe.Text)) * iNumberOfBolts

    End Function
    Private Function CalculateDeburring(ByVal iNumberOfSubjects As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO

        If cb_afgrat.CheckState = CheckState.Unchecked Then
            lb_afgrat.Text = ""
            tb_afgrat_uk.Text = ""
            Exit Function
        Else
            CalculateDeburring = (objDatabaseIO.GetDeburringTime100(lb_modulstr.Text) / 100 * iNumberOfSubjects) + objDatabaseIO.GetDeburringstartup * Difficultfactor

        End If

        lb_afgrat.Text = CalculateDeburring


    End Function
    Private Function CalculateSteelmaster(ByVal iNumberOfSubjects As Integer, ByVal SubjectX As Double, ByVal SubjectY As Double) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO


        If cb_steelmaster.CheckState = CheckState.Unchecked Then
            lb_grinding.Text = ""
            tb_grinding_uk.Text = ""
            Exit Function
        End If

        If SubjectX < 150 Then
            If SubjectY < 150 Then
                MsgBox("Emnet er for lille til Steelmasteren")
                lb_grinding.Text = ""
                tb_grinding_uk.Text = ""
                cb_steelmaster.CheckState = CheckState.Unchecked
                Exit Function
            End If
        End If
        If SubjectX > 1250 Then
            If SubjectY > 1250 Then
                MsgBox("Emnet er for stort til Steelmasteren")
                lb_grinding.Text = ""
                tb_grinding_uk.Text = ""
                cb_steelmaster.CheckState = CheckState.Unchecked
                Exit Function
            End If
        End If

        CalculateSteelmaster = ((objDatabaseIO.GetSteelmasterTime100(lb_modulstr.Text) / 100 * iNumberOfSubjects) * objDatabaseIO.GetSteelmastermultiplier(lb_modulstr.Text, iNumberOfSubjects)) + objDatabaseIO.GetSteelmasterstartup * Difficultfactor


        lb_grinding.Text = CalculateSteelmaster


    End Function
    Private Function CalculateStraigthening(ByVal iNumberOfSubjects As Integer, ByVal SubjectX As Double, ByVal SubjectY As Double, ByVal sheetthickness As Double) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO


        If cb_rette.CheckState = CheckState.Unchecked Then
            lb_rette.Text = ""
            tb_rette_uk.Text = ""
            Exit Function
        End If
        If sheetthickness > 3 Then
            MsgBox("Pladetykkelsen er for stort til Rettemaskinen")
            lb_rette.Text = ""
            tb_rette_uk.Text = ""
            Exit Function
        End If


        If SubjectX > 550 Then
            If SubjectY > 550 Then
                MsgBox("Emnet er for stort til Rettemaskinen")
                lb_rette.Text = ""
                tb_rette_uk.Text = ""
                cb_rette.CheckState = CheckState.Unchecked
                Exit Function
            End If
        End If

        CalculateStraigthening = ((objDatabaseIO.GetStraigtheningTime100(lb_modulstr.Text) / 100 * iNumberOfSubjects) * objDatabaseIO.GetStraigtheningmultiplier(lb_modulstr.Text, iNumberOfSubjects)) + objDatabaseIO.GetStraigtheningstartup * Difficultfactor


        lb_rette.Text = CalculateStraigthening

    End Function

    Private Function CalculateBrushing(ByVal iNumberOfSubjects As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO

        If cb_brush.CheckState = CheckState.Unchecked Then
            lb_brush.Text = ""
            tb_brush_uk.Text = ""
            Exit Function
        Else
            CalculateBrushing = (objDatabaseIO.GetBrushingTime100(lb_modulstr.Text) / 100 * iNumberOfSubjects) + objDatabaseIO.GetBrushingstartup * Difficultfactor
        End If

        lb_brush.Text = CalculateBrushing

    End Function
    Private Function CalculateBestPlate(ByVal iNumberOfSubjects As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim alPlateSizes As ArrayList
        Dim alPlate As ArrayList
        Dim alAllPlateValues As ArrayList
        Dim alPlateValues As ArrayList

        Dim dblSubjectX As Double
        Dim dblSubjectY As Double
        Dim dblLowestWaste As Double
        Dim dblTotalWaste As Double

        Dim iLowestWasteIndex As Integer
        Dim iNumberOfPlates As Integer
        Dim iSubjectsPerPlate As Integer
        Dim iPlateX As Integer
        Dim iPlateY As Integer
        Dim i As Long


        '-------------------------------------------------
        ' Read input-boxes.
        '-------------------------------------------------
        dblSubjectX = lb_udfold_x.Text
        dblSubjectY = lb_udfold_y.Text
        '-------------------------------------------------
        objDatabaseIO = New DatabaseIO
        alPlateSizes = objDatabaseIO.GetPlateSizes(cb_fravælg.CheckState)
        alAllPlateValues = New ArrayList(alPlateSizes.Count)
        For i = 0 To alPlateSizes.Count - 1
            alPlate = alPlateSizes(i)
            alPlateValues = CalculatePlateValues(iNumberOfSubjects, alPlate(0), alPlate(1), dblSubjectX, dblSubjectY)
            alAllPlateValues.Add(alPlateValues)
        Next
        dblLowestWaste = 1000000000
        For i = 0 To alAllPlateValues.Count - 1
            alPlateValues = alAllPlateValues(i)
            If alPlateValues(1) < dblLowestWaste Then
                dblLowestWaste = alPlateValues(1)
                iLowestWasteIndex = i
            End If
        Next
        alPlateValues = alAllPlateValues(iLowestWasteIndex)
        '-------------------------------------------------
        ' Output results.
        '-------------------------------------------------
        iNumberOfPlates = alPlateValues(0)
        dblTotalWaste = alPlateValues(1)
        iSubjectsPerPlate = alPlateValues(2)
        alPlate = alPlateSizes(iLowestWasteIndex)
        iPlateX = alPlate(0)
        iPlateY = alPlate(1)

        If iSubjectsPerPlate = 0 Then

            MsgBox("Emnet kan ikke være på 1 plade")
            CalculateBestPlate = 1
            Exit Function
        End If

        CalculateWasteValues(iNumberOfSubjects, iPlateX, iPlateY, dblSubjectX, dblSubjectY)

        lb_antalplader.Text = iNumberOfPlates
        lb_pladeformatX.Text = iPlateX
        lb_pladeformatY.Text = iPlateY
        If iPlateY = 4000 Then
            If rb_C_laser.Checked = False Then
                MsgBox("Kun C-LASER KAN anvendes til format 2000x4000")
                rb_C_laser.Checked = True
            End If
        End If
        lb_emner_prplade.Text = iSubjectsPerPlate
        lb_spild.Text = dblTotalWaste / 1000000

        '-------------------------------------------------
    End Function

    Private Function CalculatePlateValues(ByVal iNumberOfSubjects As Integer, ByVal PlateX As Double, ByVal PlateY As Double, ByVal SubjectX As Double, ByVal SubjectY As Double) As ArrayList
        Dim alReturn As ArrayList
        Dim objDatabaseIO As DatabaseIO
        Dim dblWaste As Double
        Dim dblPlateArea As Double
        Dim dblPlateUsedArea As Double

        Dim iXCountX As Integer
        Dim iYCountY As Integer
        Dim iXCountY As Integer
        Dim iYCountX As Integer
        Dim iLastPlateCount As Integer
        Dim iSubjectsPerPlate As Integer
        Dim iNumberOfPlates As Integer
        Dim wasteX As Double
        Dim wasteY As Double
        Dim subjectaddition As Double
        Dim CountRowX As Double
        Dim CountRowY As Double
        Dim iCountRow As Integer
        Dim platemultiplier As Integer

        objDatabaseIO = New DatabaseIO
        platemultiplier = 1


        wasteX = 40
        wasteY = 100
        subjectaddition = 20

        If rb_klip.Checked = True Then
            wasteX = 0
            wasteY = 0
            subjectaddition = 0
        End If
        If rb_C_laser.Checked Then
            wasteX = 20
            wasteY = 20
            subjectaddition = 10
        End If

        alReturn = New ArrayList(3)
        iXCountX = (PlateX - wasteX) \ (SubjectX + subjectaddition)
        iYCountY = (PlateY - wasteY) \ (SubjectY + subjectaddition)
        iXCountY = (PlateX - wasteX) \ (SubjectY + subjectaddition)
        iYCountX = (PlateY - wasteY) \ (SubjectX + subjectaddition)
        If iXCountX * iYCountY >= iXCountY * iYCountX Then
            iSubjectsPerPlate = iXCountX * iYCountY
            CountRowX = iXCountX
            CountRowY = iYCountY
        Else
            iSubjectsPerPlate = iXCountY * iYCountX
            CountRowX = iXCountY
            CountRowY = iYCountX
        End If

        If CountRowX < CountRowY Then
            iCountRow = CountRowX
        Else
            iCountRow = CountRowY
        End If

        If iSubjectsPerPlate = 0 Then
            dblWaste = 1000000000
            alReturn.Add(iNumberOfPlates)
            alReturn.Add(dblWaste)
            alReturn.Add(iSubjectsPerPlate)
            CalculatePlateValues = alReturn

            Exit Function
        End If
        'If iSubjectsPerPlate = 0 Then
        'isubjectsPerPlate = 1
        'MsgBox("Emnet kan ikke være på 1 plade")
        'platemultiplier = 2
        'End If

        lb_rækkeantal.Text = iCountRow
        iNumberOfPlates = iNumberOfSubjects \ iSubjectsPerPlate * platemultiplier
        dblPlateArea = PlateX * PlateY
        dblPlateUsedArea = iSubjectsPerPlate * SubjectX * SubjectY
        dblWaste = (dblPlateArea - dblPlateUsedArea) * iNumberOfPlates
        iLastPlateCount = iNumberOfSubjects Mod iSubjectsPerPlate
        If iLastPlateCount > 0 Then
            dblWaste += dblPlateArea - (iLastPlateCount * SubjectX * SubjectY)
            iNumberOfPlates += 1
        End If


        alReturn.Add(iNumberOfPlates)
        alReturn.Add(dblWaste)
        alReturn.Add(iSubjectsPerPlate)
        CalculatePlateValues = alReturn
    End Function

    Private Function CalculateWasteValues(ByVal iNumberOfSubjects As Integer, ByVal PlateX As Double, ByVal PlateY As Double, ByVal SubjectX As Double, ByVal SubjectY As Double) As Double

        Dim objDatabaseIO As DatabaseIO

        Dim iXCountX As Integer
        Dim iYCountY As Integer
        Dim iXCountY As Integer
        Dim iYCountX As Integer
        Dim iLastPlateCount As Integer
        Dim iSubjectsPerPlate As Integer
        Dim iNumberOfPlates As Integer
        Dim wasteX As Double
        Dim wasteY As Double
        Dim StartwasteX As Double
        Dim StartwasteY As Double
        Dim subjectaddition As Double
        Dim CountRowX As Double
        Dim CountRowY As Double
        Dim iCountRow As Integer
        Dim platemultiplier As Integer
        Dim RowX As Double
        Dim RowY As Double
        Dim Areareduce As Integer
        Dim WasteMinimumSize As Double
        Dim Waste As Double
        Dim Wastenetto As Double
        Dim SpildTilLager As Double
        Dim Areareducelastplate As Integer
        Dim Wastenettolastplate As Double
        Dim alwasteValues As ArrayList


        objDatabaseIO = New DatabaseIO
        platemultiplier = 1
        WasteMinimumSize = objDatabaseIO.GetWasteMinimumSizes(Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text))


        StartwasteX = 40
        StartwasteY = 100
        subjectaddition = 20

        If rb_klip.Checked = True Then
            StartwasteX = 0
            StartwasteY = 0
            subjectaddition = 0
        End If
        If rb_C_laser.Checked Then
            StartwasteX = 20
            StartwasteY = 20
            subjectaddition = 10
        End If

        iXCountX = (PlateX - StartwasteX) \ (SubjectX + subjectaddition)
        iYCountY = (PlateY - StartwasteY) \ (SubjectY + subjectaddition)
        iXCountY = (PlateX - StartwasteX) \ (SubjectY + subjectaddition)
        iYCountX = (PlateY - StartwasteY) \ (SubjectX + subjectaddition)
        If iXCountX * iYCountY >= iXCountY * iYCountX Then
            iSubjectsPerPlate = iXCountX * iYCountY
            CountRowX = iXCountX
            CountRowY = iYCountY
            RowX = SubjectX + subjectaddition
            RowY = SubjectY + subjectaddition
        Else
            iSubjectsPerPlate = iXCountY * iYCountX
            CountRowX = iXCountY
            CountRowY = iYCountX
            RowX = SubjectY
            RowY = SubjectX
        End If

        If CountRowX < CountRowY Then
            iCountRow = CountRowX
        Else
            iCountRow = CountRowY
        End If


        lb_rækkeantal.Text = iCountRow
        iNumberOfPlates = iNumberOfSubjects \ iSubjectsPerPlate * platemultiplier

        iLastPlateCount = iNumberOfSubjects Mod iSubjectsPerPlate

        If iLastPlateCount = 0 Then


            wasteX = PlateX - (RowX * CountRowX) + StartwasteY
            wasteY = PlateY - (RowY * CountRowY) + StartwasteX
            If wasteX > wasteY Then
                Waste = wasteX
                Wastenetto = (wasteX * PlateY * (iNumberOfPlates - 1)) / 1000000
            Else
                Waste = wasteY
                Wastenetto = (wasteY * PlateX * (iNumberOfPlates - 1)) / 1000000
            End If
            If Wastenetto < 0 Then
                Wastenetto = Wastenetto * -1
            Else
                iLastPlateCount = iNumberOfSubjects
            End If
        End If

        alwasteValues = CalculateWasteLastplate(iLastPlateCount, PlateX, PlateY, SubjectX, SubjectY)

        Wastenettolastplate = alwasteValues(0)
        Areareducelastplate = alwasteValues(1)

        If Waste >= WasteMinimumSize Then
            Areareduce = 1
            lb_spildtype.Text = "TIL LAGER"
            SpildTilLager = Wastenetto
            lb_spildnetto.Text = SpildTilLager
        Else
            Areareduce = 0
            lb_spildtype.Text = ""
            lb_spildnetto.Text = ""
        End If

        If Areareducelastplate = 1 Then
            lb_spildtype.Text = "TIL LAGER"
            SpildTilLager = SpildTilLager + Wastenettolastplate
            lb_spildnetto.Text = SpildTilLager

        End If
        If rb_brutto.Checked = True Then
            lb_spildtype.Text = ""
            lb_spildnetto.Text = ""
        End If


    End Function

    Private Function CalculateWasteLastplate(ByVal iLastPlateCount As Integer, ByVal PlateX As Double, ByVal PlateY As Double, ByVal SubjectX As Double, ByVal SubjectY As Double) As ArrayList
        Dim objDatabaseIO As DatabaseIO

        Dim alReturn As ArrayList
        Dim iXCountX As Integer
        Dim iYCountY As Integer
        Dim iXCountY As Integer
        Dim iYCountX As Integer
        Dim subjectaddition As Double
        Dim WasteMinimumSize As Double
        Dim Waste As Double
        Dim Waste1 As Double
        Dim wasteX As Double
        Dim wasteY As Double
        Dim wasteX1 As Double
        Dim wasteY1 As Double
        Dim StartwasteX As Double
        Dim StartwasteY As Double
        Dim iCountRowX As Double
        Dim iCountRowY As Double
        Dim RowX As Double
        Dim RowY As Double
        Dim Wastenettolastplate As Double
        Dim Wastenettolastplate1 As Double
        Dim UsedCountRowX As Integer
        Dim UsedCountRowY As Integer
        Dim UsedCountRowX1 As Integer
        Dim UsedCountRowY1 As Integer
        Dim Areareducelastplate As Integer

        objDatabaseIO = New DatabaseIO

        WasteMinimumSize = objDatabaseIO.GetWasteMinimumSizes(Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text))

        startwasteX = 40
        startwasteY = 100
        subjectaddition = 20

        If rb_klip.Checked = True Then
            startwasteX = 0
            startwasteY = 0
            subjectaddition = 0
        End If
        If rb_C_laser.Checked Then
            startwasteX = 20
            startwasteY = 20
            subjectaddition = 10
        End If

        alReturn = New ArrayList(2)

        iXCountX = (PlateX - startwasteX) \ (SubjectX + subjectaddition)
        iYCountY = (PlateY - startwasteY) \ (SubjectY + subjectaddition)
        iXCountY = (PlateX - startwasteX) \ (SubjectY + subjectaddition)
        iYCountX = (PlateY - startwasteY) \ (SubjectX + subjectaddition)

        If iXCountX * iYCountY >= iXCountY * iYCountX Then
            iCountRowX = iXCountX
            iCountRowY = iYCountY
            RowX = SubjectX + subjectaddition
            RowY = SubjectY + subjectaddition
        Else
            iCountRowX = iXCountY
            iCountRowY = iYCountX
            RowX = SubjectY + subjectaddition
            RowY = SubjectX + subjectaddition
        End If

        UsedCountRowX = (iLastPlateCount / iCountRowX) + 0.49
        UsedCountRowY = (iLastPlateCount / iCountRowY) + 0.49

        wasteX = PlateX - (RowX * iYCountY) + StartwasteY
        wasteY = PlateY - (RowY * iXCountX) + StartwasteX

        If wasteX > wasteY Then
            Waste = wasteX
            Wastenettolastplate = (wasteX * PlateY) / 1000000
        Else
            Waste = wasteY
            Wastenettolastplate = (wasteY * PlateX) / 1000000
        End If

        UsedCountRowX1 = (iLastPlateCount / iYCountX) + 0.49
        UsedCountRowY1 = (iLastPlateCount / iXCountY) + 0.49

        wasteX1 = PlateY - (RowX * UsedCountRowY1) + StartwasteX
        wasteY1 = PlateX - (RowY * UsedCountRowX1) + StartwasteY

        If wasteX1 > wasteY1 Then
            Waste1 = wasteX1
            Wastenettolastplate1 = (wasteX1 * PlateY) / 1000000
        Else
            Waste1 = wasteY1
            Wastenettolastplate1 = (wasteY1 * PlateX) / 1000000
        End If

        If Wastenettolastplate1 > Wastenettolastplate Then
            Wastenettolastplate = Wastenettolastplate1
            Waste = Waste1
        End If

        If Waste >= WasteMinimumSize Then
            Areareducelastplate = 1
        Else
            Areareducelastplate = 0
        End If

        alReturn.Add(Wastenettolastplate)
        alReturn.Add(Areareducelastplate)
        CalculateWasteLastplate = alReturn


    End Function

    '-----------------------------------------------------------------------------
    '   EMNESTØRRELSE
    '------------------------------------------------------------------------
    Private Function CalculateSubjectSize() As Double
        Dim objDatabaseIO As DatabaseIO
        Dim objbendcalculations As BendCalculations
        Dim FModuleSize As Integer
        Dim Difficultclass As Double
        Dim Difficultmultiplier As Double
        objbendcalculations = New BendCalculations

        lb_udfold_x.Text = objbendcalculations.CalculateSubjectLength(tb_pladetykkelse.Text, Val(tb_bukmax_x.Text), Val(tb_buk1_x.Text), Val(tb_buk2_x.Text), Val(tb_buk3_x.Text), Val(tb_buk4_x.Text), Val(tb_buk5_x.Text), Val(tb_buk6_x.Text), Val(tb_buk7_x.Text), Val(tb_buk8_x.Text), Val(tb_buk9_x.Text), Val(tb_buk10_x.Text), Val(tb_buk11_x.Text))
        lb_udfold_y.Text = objbendcalculations.CalculateSubjectLength(tb_pladetykkelse.Text, Val(tb_bukmax_y.Text), Val(tb_buk1_y.Text), Val(tb_buk2_y.Text), Val(tb_buk3_y.Text), Val(tb_buk4_y.Text), Val(tb_buk5_y.Text), Val(tb_buk6_y.Text), Val(tb_buk7_y.Text), Val(tb_buk8_y.Text), Val(tb_buk9_y.Text), Val(tb_buk10_y.Text), Val(tb_buk11_y.Text))


        FModuleSize = CalculateModuleSize(lb_udfold_x.Text, lb_udfold_y.Text, FModuleSize)
        lb_modulstr.Text = FModuleSize


        'getting difficult factor
        objDatabaseIO = New DatabaseIO

        Difficultclass = objDatabaseIO.GetDifficultclass(Lb_klasse.Text, tb_pladetykkelse.Text)
        lb_sværhed.Text = Difficultclass
        Difficultmultiplier = objDatabaseIO.GetDifficultmultiplier(Val(pobjFilter_tb_sværhed_uk.GetValue))
        lb_faktor.Text = Difficultmultiplier

    End Function


    Private Function CalculateModuleSize(ByVal LengthX As Double, ByVal LengthY As Double, ByVal FModuleSize As Integer) As Integer

        Dim nextSize As Integer
        Dim ModuleSize As Integer

        nextSize = 0

        If 0.5 < (LengthX / LengthY) > 2 Then
            nextSize = 1

        End If

        If LengthX * LengthY <= 6400 Then
            ModuleSize = 1
        End If

        If 6400 < LengthX * LengthY And LengthX * LengthY <= 62500 Then
            ModuleSize = 2
        End If

        If 62500 < LengthX * LengthY And LengthX * LengthY <= 250000 Then
            ModuleSize = 3
        End If

        If 250000 < LengthX * LengthY And LengthX * LengthY <= 1000000 Then
            ModuleSize = 4
        End If

        If 1000000 < LengthX * LengthY And LengthX * LengthY <= 2000000 Then
            ModuleSize = 5
        End If

        If 2000000 < LengthX * LengthY And LengthX * LengthY <= 3125000 Then
            ModuleSize = 6
        End If

        If LengthX * LengthY > 3125000 Then
            ModuleSize = 7
        End If

        CalculateModuleSize = ModuleSize + nextSize



    End Function

    Private Sub CalculateBendTime(ByVal subjectweight As Double)

        Dim objbendcalculations As BendCalculations
        Dim sumY As Double
        Dim BendsY As Integer
        Dim bend As ArrayList
        Dim sumX As Double
        Dim sum As Double
        Dim bendmin100 As Double
        Dim BendsX As Integer

        objbendcalculations = New BendCalculations
        bend = New ArrayList
        bend.Add(Val(tb_buk1_x.Text))
        bend.Add(Val(tb_buk2_x.Text))
        bend.Add(Val(tb_buk3_x.Text))
        bend.Add(Val(tb_buk4_x.Text))
        bend.Add(Val(tb_buk5_x.Text))
        bend.Add(Val(tb_buk6_x.Text))
        bend.Add(Val(tb_buk7_x.Text))
        bend.Add(Val(tb_buk8_x.Text))
        bend.Add(Val(tb_buk9_x.Text))
        bend.Add(Val(tb_buk10_x.Text))
        bend.Add(Val(tb_buk11_x.Text))

        BendsX = objbendcalculations.CountBends(Val(tb_buk1_x.Text), Val(tb_buk2_x.Text), Val(tb_buk3_x.Text), Val(tb_buk4_x.Text), Val(tb_buk5_x.Text), Val(tb_buk6_x.Text), Val(tb_buk7_x.Text), Val(tb_buk8_x.Text), Val(tb_buk9_x.Text), Val(tb_buk10_x.Text), Val(tb_buk11_x.Text))

        sumX = objbendcalculations.CalculateAddition(bend)

        bend.Clear()

        bend.Add(Val(tb_buk1_y.Text))
        bend.Add(Val(tb_buk2_y.Text))
        bend.Add(Val(tb_buk3_y.Text))
        bend.Add(Val(tb_buk4_y.Text))
        bend.Add(Val(tb_buk5_y.Text))
        bend.Add(Val(tb_buk6_y.Text))
        bend.Add(Val(tb_buk7_y.Text))
        bend.Add(Val(tb_buk8_y.Text))
        bend.Add(Val(tb_buk9_y.Text))
        bend.Add(Val(tb_buk10_y.Text))
        bend.Add(Val(tb_buk11_y.Text))

        BendsY = objbendcalculations.CountBends(Val(tb_buk1_y.Text), Val(tb_buk2_y.Text), Val(tb_buk3_y.Text), Val(tb_buk4_y.Text), Val(tb_buk5_y.Text), Val(tb_buk6_y.Text), Val(tb_buk7_y.Text), Val(tb_buk8_y.Text), Val(tb_buk9_y.Text), Val(tb_buk10_y.Text), Val(tb_buk11_y.Text))

        sumY = objbendcalculations.CalculateAddition(bend)

        bendmin100 = CalculateBendMin(Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), subjectweight)


        sum = (bendmin100 * (BendsX + BendsY)) + (sumX * bendmin100) + (sumY * bendmin100)
        lb_buk_tid.Text = sum.ToString()
        lb_buk_opst.Text = CalculateBendSetup(Val(lb_udfold_x.Text), BendsY) + CalculateBendSetup(Val(lb_udfold_y.Text), BendsX)


    End Sub

    Private Function CalculateBendSetup(ByVal length As Double, ByVal countbends As Integer)
        Dim Difficultfactor As Double

        CalculateBendSetup = (length / 160 + 4) + ((length / 500 + 4) * countbends)

        If lb_faktor.Text = "" Then
            Exit Function
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        CalculateBendSetup = CalculateBendSetup * Difficultfactor

    End Function

    Private Function CalculateBendMin(ByVal lengthX As Double, ByVal lengthY As Double, ByVal SubjectWeight As Double)
        Dim Difficultfactor As Double

        If SubjectWeight < 19 Then
            CalculateBendMin = (Math.Pow(((Math.Sqrt(lengthX) + 2) * (Math.Sqrt(lengthY) + 1)), 1.15) / 20) - (Math.Sqrt(lengthX * lengthY) / 10) + 10

        ElseIf SubjectWeight < 48 Then
            CalculateBendMin = (Math.Pow(((Math.Sqrt(lengthX) + 4) * (Math.Sqrt(lengthY))), 1.25) / 20) - (Math.Sqrt(lengthX * lengthY) / 3.6) + 35

        Else
            CalculateBendMin = (lengthX * lengthY) / Math.Sqrt(lengthX * lengthY) * 0.15
        End If

        If lb_faktor.Text = "" Then
            Exit Function
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        CalculateBendMin = CalculateBendMin * Difficultfactor

    End Function
    Private Function CalculateWelding(ByVal antal1 As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim countwelds As Integer
        Dim weldlength As Double
        Dim weldspeed As Double
        Dim weldingtimemin As Double
        Dim weldinghandling As Double
        Dim weldspeedmultiplier As Double
        Dim opstart As Double

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO
        countwelds = Val(tb_numberofwelds.Text)
        weldlength = Val(tb_weldlength.Text)
        weldinghandling = 0
        'justering af svejsehastighed
        weldspeedmultiplier = 0.66

        opstart = 15
        weldinghandling = opstart + antal1 * ((objDatabaseIO.GetWeldingShift(Val(lb_modulstr.Text)) + (objDatabaseIO.GetWeldingStartstop(Val(lb_modulstr.Text)) * countwelds))) * Difficultfactor

        If cb_weld.CheckState = CheckState.Checked Then
            If tb_numberofwelds.Text = "" Then
                MsgBox("Antal svejsninger og svejselængde skal være udfyldt")
                cb_weld.CheckState = CheckState.Unchecked
                Exit Function
            End If
            If tb_weldlength.Text = "" Then
                MsgBox("Antal svejsninger og svejselængde skal være udfyldt")
                cb_weld.CheckState = CheckState.Unchecked
                Exit Function
            End If
            If rb_tig.Checked = True Then
                weldspeed = objDatabaseIO.GetweldspeedTIG100(Val(tb_pladetykkelse.Text), Val(Lb_matrgruppe.Text))
            Else
                weldspeed = objDatabaseIO.GetweldspeedMIG100(Val(tb_pladetykkelse.Text), Val(Lb_matrgruppe.Text))
            End If

            weldingtimemin = antal1 * weldspeed * weldlength * countwelds * weldspeedmultiplier / 100

            lb_weld.Text = FormatNumber(weldingtimemin + weldinghandling, 0)
        End If


        If cb_tackweld.CheckState = CheckState.Checked Then
            If tb_numberofwelds.Text = "" Then
                MsgBox("Antal svejsninger og svejselængde skal være udfyldt")
                cb_tackweld.CheckState = CheckState.Unchecked
                Exit Function
            End If
            If tb_weldlength.Text = "" Then
                MsgBox("Antal svejsninger og svejselængde skal være udfyldt")
                cb_tackweld.CheckState = CheckState.Unchecked
                Exit Function
            End If

            If rb_tig.Checked = True Then
                weldspeed = objDatabaseIO.GettackweldspeedTIG100(Val(tb_pladetykkelse.Text), Val(Lb_matrgruppe.Text))
            Else
                weldspeed = objDatabaseIO.GettackweldspeedMIG100(Val(tb_pladetykkelse.Text), Val(Lb_matrgruppe.Text))
            End If

            weldingtimemin = antal1 * weldspeed * weldlength * countwelds * weldspeedmultiplier / 100

            lb_tackweld.Text = FormatNumber(weldingtimemin + weldinghandling, 0)
        End If

        If cb_weld.CheckState = CheckState.Unchecked Then
            lb_weld.Text = ""
            tb_weld_uk.Text = ""
        End If
        If cb_tackweld.CheckState = CheckState.Unchecked Then
            lb_tackweld.Text = ""
            tb_tackweld_uk.Text = ""
        End If


    End Function
    Private Sub ClearAllCalculations()
        If rb_klip.Checked = False Then
            pobjFilter_tb_klip_tid_uk.ClearValues()
            pobjFilter_tb_klip_opstart_uk.ClearValues()
        End If
        If rb_D_stans.Checked = False Then
            pobjFilter_tb_CNCmin_uk.ClearValues()
            pobjFilter_tb_gruppe1_opstart_uk.ClearValues()
        End If
        If rb_C_laser.Checked = False Then
            pobjFilter_tb_laserCNC_tid_uk.ClearValues()
            pobjFilter_tb_laser_opstart_uk.ClearValues()
        End If
        If rb_B_kombi.Checked = False Then
            pobjFilter_tb_combi_opstart_uk.ClearValues()
            pobjFilter_tb_combiCNC_tid_uk.ClearValues()
        End If
    End Sub

    Private Sub CalculateManHourSubtotal()
        Dim objDatabaseIO As DatabaseIO

        Dim dblSubtotal As Double
        Dim dblSubtotalstart As Double
        Dim dblFactor As Double
        Dim dblSubtotalTime As Double
        Dim SubtotalTime As Double
        Dim dblSubtotalstartTime As Double
        objDatabaseIO = New DatabaseIO


        'Summing
        dblSubtotalTime = ((Val(pobjFilter_tb_buk_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_klip_tid_uk.GetValue) + Val(pobjFilter_tb_afgrat_uk.GetValue) + Val(pobjFilter_tb_grinding_uk.GetValue) + Val(pobjFilter_tb_brush_uk.GetValue) + Val(pobjFilter_tb_vibration_uk.GetValue) + Val(pobjFilter_tb_rette_uk.GetValue) + Val(pobjFilter_tb_stans_manuel_uk.GetValue) + Val(pobjFilter_tb_countersink_uk.GetValue) + Val(pobjFilter_tb_gevind_uk.GetValue) + Val(pobjFilter_tb_pressnut_uk.GetValue) + Val(pobjFilter_tb_boltesvejs_uk.GetValue) + Val(pobjFilter_tb_spotweld_uk.GetValue) + Val(pobjFilter_tb_tackweld_uk.GetValue) + Val(pobjFilter_tb_weld_uk.GetValue) + Val(pobjFilter_tb_grind_weld_uk.GetValue)) / 60)
        dblSubtotal = dblSubtotalTime * Val(tb_timesats_mand.Text)
        dblSubtotalstartTime = dblSubtotalstartTime + (Val(pobjFilter_tb_buk_opst_uk.GetValue) + Val(pobjFilter_tb_klip_opstart_uk.GetValue) + Val(pobjFilter_tb_kontor_uk.GetValue) + Val(pobjFilter_tb_kontrol_uk.GetValue)) / 60
        dblSubtotalstart = dblSubtotalstartTime * Val(tb_timesats_mand.Text)

        'dblSubtotal = dblSubtotal + dblSubtotalstart
        'dblSubtotalTime = FormatNumber((dblSubtotalTime + dblSubtotalstartTime), 2)
        dblSubtotalTime = FormatNumber(dblSubtotalTime, 2)
        dblSubtotalstartTime = FormatNumber(dblSubtotalstartTime, 2)

        'Factor application
        lb_mand1.Text = Convert.ToString(dblSubtotal + dblSubtotalstart)
        lb_mand1_tid.Text = Convert.ToString(dblSubtotalTime + dblSubtotalstartTime)

        lb_mand1.Text = Val(lb_mand1.Text)
        'lb_mand1.Text = FormatNumber(lb_mand1.Text, 2)

        If (Val(tb_antal2.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal2.Text)
            lb_mand2.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal2.Text) / 100)) + dblSubtotalstart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal2.Text) / 100)) + dblSubtotalstartTime), 2)
            lb_mand2_tid.Text = Convert.ToString(SubtotalTime)
            lb_mand2.Text = Val(lb_mand2.Text)
            'lb_mand2.Text = FormatNumber(lb_mand2.Text, 2)
        End If
        If (Val(tb_antal3.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal3.Text)
            lb_mand3.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal3.Text) / 100)) + dblSubtotalstart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal3.Text) / 100)) + dblSubtotalstartTime), 2)
            lb_mand3_tid.Text = Convert.ToString(SubtotalTime)
            lb_mand3.Text = Val(lb_mand3.Text)
            'lb_mand3.Text = FormatNumber(lb_mand3.Text, 2)
        End If
        If (Val(tb_antal4.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal4.Text)
            lb_mand4.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal4.Text) / 100)) + dblSubtotalstart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal4.Text) / 100)) + dblSubtotalstartTime), 2)
            lb_mand4_tid.Text = Convert.ToString(SubtotalTime)
            lb_mand4.Text = Val(lb_mand4.Text)
            'lb_mand4.Text = FormatNumber(lb_mand4.Text, 2)
        End If
        If (Val(tb_antal5.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal5.Text)
            lb_mand5.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal5.Text) / 100)) + dblSubtotalstart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal5.Text) / 100)) + dblSubtotalstartTime), 2)
            lb_mand5_tid.Text = Convert.ToString(SubtotalTime)
            lb_mand5.Text = Val(lb_mand5.Text)
            'lb_mand5.Text = FormatNumber(lb_mand5.Text, 2)
        End If

    End Sub

    Private Sub CalculateCNCHourSubtotal()
        Dim objDatabaseIO As DatabaseIO
        Dim SubtotalTime As Double
        Dim dblSubtotal As Double
        Dim dblSubStart As Double
        Dim dblFactor As Double
        Dim dblSubtotalTime As Double
        Dim dblSubStartTime As Double


        objDatabaseIO = New DatabaseIO
        'Summing
        If pobjFilter_tb_CNCmin_uk.GetValue <> 0 Then
            dblSubtotal = dblSubtotal + (Convert.ToDouble(pobjFilter_tb_CNCmin_uk.GetValue) * Val(tb_timesats_D.Text) / 60)
            dblSubtotalTime = dblSubtotalTime + (Convert.ToDouble(pobjFilter_tb_CNCmin_uk.GetValue) / 60)
            dblSubStart = dblSubStart + (Convert.ToDouble(pobjFilter_tb_gruppe1_opstart_uk.GetValue) * Val(tb_timesats_D.Text) / 60)
            dblSubStartTime = dblSubStartTime + (Convert.ToDouble(pobjFilter_tb_gruppe1_opstart_uk.GetValue) / 60)
        End If
        If pobjFilter_tb_laserCNC_tid_uk.GetValue <> 0 Then
            dblSubtotal = dblSubtotal + (Convert.ToDouble(pobjFilter_tb_laserCNC_tid_uk.GetValue) * Val(tb_timesats_C.Text) / 60)
            dblSubtotalTime = dblSubtotalTime + (Convert.ToDouble(pobjFilter_tb_laserCNC_tid_uk.GetValue) / 60)
            dblSubStart = dblSubStart + (Convert.ToDouble(pobjFilter_tb_laser_opstart_uk.GetValue) * Val(tb_timesats_C.Text) / 60)
            dblSubStartTime = dblSubStartTime + (Convert.ToDouble(pobjFilter_tb_laser_opstart_uk.GetValue) / 60)
        End If
        If pobjFilter_tb_combiCNC_tid_uk.GetValue <> 0 Then
            dblSubtotal = dblSubtotal + (Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) * Val(tb_timesats_B.Text) / 60)
            dblSubtotalTime = dblSubtotalTime + (Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) / 60)
            dblSubStart = dblSubStart + (Convert.ToDouble(pobjFilter_tb_combi_opstart_uk.GetValue) * Val(tb_timesats_B.Text) / 60)
            dblSubStartTime = dblSubStartTime + (Convert.ToDouble(pobjFilter_tb_combi_opstart_uk.GetValue) / 60)

        End If
        'dblSubtotal = FormatNumber(dblSubtotal, 0)
        'dblSubStart = FormatNumber(dblSubStart, 0)
        dblSubtotalTime = FormatNumber(dblSubtotalTime, 2)
        dblSubStartTime = FormatNumber(dblSubStartTime, 2)

        'Factor application
        lb_cnc1.Text = Convert.ToString(dblSubtotal + dblSubStart)
        lb_CNC1_tid.Text = Convert.ToString(dblSubtotalTime + dblSubStartTime)

        'lb_cnc1.Text = Val(lb_cnc1.Text)
        lb_cnc1.Text = FormatNumber(lb_cnc1.Text, 0)
        If (Val(tb_antal2.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal2.Text)
            lb_cnc2.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal2.Text) / 100)) + dblSubStart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal2.Text) / 100)) + dblSubStartTime), 2)
            lb_CNC2_tid.Text = Convert.ToString(SubtotalTime)
            lb_cnc2.Text = Val(lb_cnc2.Text)
            'lb_cnc2.Text = FormatNumber(lb_cnc2.Text, 2)
        End If
        If (Val(tb_antal3.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal3.Text)
            lb_cnc3.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal3.Text) / 100)) + dblSubStart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal3.Text) / 100)) + dblSubStartTime), 2)
            lb_CNC3_tid.Text = Convert.ToString(SubtotalTime)
            lb_cnc3.Text = Val(lb_cnc3.Text)
            'lb_cnc3.Text = FormatNumber(lb_cnc3.Text, 2)
        End If
        If (Val(tb_antal4.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal4.Text)
            lb_cnc4.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal4.Text) / 100)) + dblSubStart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal4.Text) / 100)) + dblSubStartTime), 2)
            lb_CNC4_tid.Text = Convert.ToString(SubtotalTime)
            lb_cnc4.Text = Val(lb_cnc4.Text)
            'lb_cnc4.Text = FormatNumber(lb_cnc4.Text, 2)
        End If
        If (Val(tb_antal5.Text) > 0) Then
            dblFactor = objDatabaseIO.getnumbermultiplier(tb_antal5.Text)
            lb_cnc5.Text = Convert.ToString(dblFactor * (dblSubtotal * (Val(tb_antal5.Text) / 100)) + dblSubStart)
            SubtotalTime = FormatNumber((dblFactor * (dblSubtotalTime * (Val(tb_antal5.Text) / 100)) + dblSubStartTime), 2)
            lb_CNC5_tid.Text = Convert.ToString(SubtotalTime)
            lb_cnc5.Text = Val(lb_cnc5.Text)
            'lb_cnc5.Text = FormatNumber(lb_cnc5.Text, 2)
        End If
    End Sub

    Private Sub CalculatePurchasingSubtotal()

        Dim pressnut As Double
        Dim boltesvejs As Double
        boltesvejs = Val(pobjFilter_tb_svejsestag_kr_uk.GetValue)
        pressnut = Val(pobjFilter_tb_pressnut_kr_uk.GetValue)
        lb_indkøb1.Text = CalculateSurfaceTreatment(Val(tb_antal1.Text)) + ((Val(pobjFilter_tb_pressnut_kr_uk.GetValue) + (Val(pobjFilter_tb_svejsestag_kr_uk.GetValue))) * Val(tb_antal1.Text))

        If Val(tb_antal2.Text) > 0 Then

            lb_indkøb2.Text = CalculateSurfaceTreatment(Val(tb_antal2.Text)) + ((Val(pobjFilter_tb_pressnut_kr_uk.GetValue) + (Val(pobjFilter_tb_svejsestag_kr_uk.GetValue))) * Val(tb_antal2.Text))
        End If
        If Val(tb_antal3.Text) > 0 Then

            lb_indkøb3.Text = CalculateSurfaceTreatment(Val(tb_antal3.Text)) + ((Val(pobjFilter_tb_pressnut_kr_uk.GetValue) + (Val(pobjFilter_tb_svejsestag_kr_uk.GetValue))) * Val(tb_antal3.Text))
        End If
        If Val(tb_antal4.Text) > 0 Then

            lb_indkøb4.Text = CalculateSurfaceTreatment(Val(tb_antal4.Text)) + ((Val(pobjFilter_tb_pressnut_kr_uk.GetValue) + (Val(pobjFilter_tb_svejsestag_kr_uk.GetValue))) * Val(tb_antal4.Text))
        End If
        If Val(tb_antal5.Text) > 0 Then

            lb_indkøb5.Text = CalculateSurfaceTreatment(Val(tb_antal5.Text)) + ((Val(pobjFilter_tb_pressnut_kr_uk.GetValue) + (Val(pobjFilter_tb_svejsestag_kr_uk.GetValue))) * Val(tb_antal5.Text))
        End If

    End Sub

    Private Function CalculateRawGoods(ByVal Numberofsheets As Integer, ByVal SheetformatX As Double, ByVal SheetformatY As Double, ByVal Sheetthickness As Double, ByVal Materialprice As Double, ByVal materialgroup As Integer, ByVal WasteToStore As Double) As Double
        Dim objDatabaseIO As DatabaseIO


        objDatabaseIO = New DatabaseIO
        CalculateRawGoods = ((Numberofsheets * ((SheetformatX * SheetformatY) / 1000000)) - WasteToStore) * Sheetthickness * Materialprice * objDatabaseIO.GetMaterialWeight(materialgroup)


    End Function

    Private Sub CalculateRawGoodsSubtotal()
        Dim råvarer As Double
        'If råvarer = 0 Then
        'Exit Sub
        'End If
        If lb_spildnetto.Text = "" Then
            lb_spildnetto.Text = 0
        End If
        råvarer = CalculateRawGoods(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(pobjFilter_tb_pladetykkelse.GetValue), FormatNumber(Lb_Kilopris.Text, 2), Val(Lb_matrgruppe.Text), FormatNumber(lb_spildnetto.Text, 2))
        lb_råvarer1.Text = FormatNumber(råvarer, 0)

        If Val(tb_antal2.Text) > 0 Then
            CalculateBestPlate(tb_antal2.Text)

            If lb_spildnetto.Text = "" Then
                lb_spildnetto.Text = 0
            End If

            råvarer = CalculateRawGoods(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(pobjFilter_tb_pladetykkelse.GetValue), FormatNumber(Lb_Kilopris.Text, 2), Val(Lb_matrgruppe.Text), FormatNumber(lb_spildnetto.Text, 2))
            lb_råvarer2.Text = FormatNumber(råvarer, 0)

        End If

        If Val(tb_antal3.Text) > 0 Then
            CalculateBestPlate(tb_antal3.Text)

            If lb_spildnetto.Text = "" Then
                lb_spildnetto.Text = 0
            End If

            råvarer = CalculateRawGoods(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(pobjFilter_tb_pladetykkelse.GetValue), FormatNumber(Lb_Kilopris.Text, 2), Val(Lb_matrgruppe.Text), FormatNumber(lb_spildnetto.Text, 2))
            lb_råvarer3.Text = FormatNumber(råvarer, 0)


        End If
        If Val(tb_antal4.Text) > 0 Then
            CalculateBestPlate(tb_antal4.Text)

            If lb_spildnetto.Text = "" Then
                lb_spildnetto.Text = 0
            End If

            råvarer = CalculateRawGoods(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(pobjFilter_tb_pladetykkelse.GetValue), FormatNumber(Lb_Kilopris.Text, 2), Val(Lb_matrgruppe.Text), FormatNumber(lb_spildnetto.Text, 2))
            lb_råvarer4.Text = FormatNumber(råvarer, 0)

        End If
        If Val(tb_antal5.Text) > 0 Then
            CalculateBestPlate(tb_antal5.Text)

            If lb_spildnetto.Text = "" Then
                lb_spildnetto.Text = 0
            End If

            råvarer = CalculateRawGoods(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(pobjFilter_tb_pladetykkelse.GetValue), FormatNumber(Lb_Kilopris.Text, 2), Val(Lb_matrgruppe.Text), FormatNumber(lb_spildnetto.Text, 2))
            lb_råvarer5.Text = FormatNumber(råvarer, 0)

        End If
        CalculateBestPlate(tb_antal1.Text)
    End Sub

    Private Sub CalculateSubtotals()
        CalculateManHourSubtotal()
        CalculateCNCHourSubtotal()
        CalculatePurchasingSubtotal()
        CalculateRawGoodsSubtotal()
    End Sub
    Private Sub CalculateTotalPrice()
        Dim samlet1 As Double
        Dim samlet2 As Double
        Dim samlet3 As Double
        Dim samlet4 As Double
        Dim samlet5 As Double

        If lb_råvarer1.Text = "" Then
            Exit Sub
        End If

        lb_opstart_program.Text = pobjFilter_tb_opstart_kr_uk.GetValue + pobjFilter_tb_program_kr_uk.GetValue
        lb_opst_prog_brutto.Text = Val(lb_opstart_program.Text) + (Val(lb_opstart_program.Text) * pobjFilter_tb_opstart_avance.GetValue / 100)


        lb_timer1.Text = FormatNumber(Convert.ToDouble(lb_mand1.Text) + Convert.ToDouble(lb_cnc1.Text), 2)
        lb_samlet1.Text = FormatNumber(((Convert.ToDouble(lb_timer1.Text) + Convert.ToDouble(lb_indkøb1.Text) + Convert.ToDouble(lb_råvarer1.Text)) / Convert.ToDouble(tb_antal1.Text)), 2) + " / " + FormatNumber(Convert.ToDouble(lb_timer1.Text) + Convert.ToDouble(lb_indkøb1.Text) + Convert.ToDouble(lb_råvarer1.Text), 0)
        samlet1 = Convert.ToDouble(lb_timer1.Text) + Convert.ToDouble(lb_indkøb1.Text) + Convert.ToDouble(lb_råvarer1.Text)
        lb_salg1.Text = FormatNumber((samlet1 + samlet1 * pobjFilter_tb_avance.GetValue / 100) / Convert.ToDouble(tb_antal1.Text), 2) + " / " + FormatNumber((samlet1 + samlet1 * pobjFilter_tb_avance.GetValue / 100), 0)


        If (Val(tb_antal2.Text) > 0) Then
            lb_timer2.Text = FormatNumber(Convert.ToDouble(lb_mand2.Text) + Convert.ToDouble(lb_cnc2.Text), 2)
            lb_samlet2.Text = FormatNumber(((Convert.ToDouble(lb_timer2.Text) + Convert.ToDouble(lb_indkøb2.Text) + Convert.ToDouble(lb_råvarer2.Text)) / Convert.ToDouble(tb_antal2.Text)), 2) + " / " + FormatNumber(Convert.ToDouble(lb_timer2.Text) + Convert.ToDouble(lb_indkøb2.Text) + Convert.ToDouble(lb_råvarer2.Text), 0)
            samlet2 = Convert.ToDouble(lb_timer2.Text) + Convert.ToDouble(lb_indkøb2.Text) + Convert.ToDouble(lb_råvarer2.Text)
            lb_salg2.Text = FormatNumber((samlet2 + samlet2 * pobjFilter_tb_avance.GetValue / 100) / Convert.ToDouble(tb_antal2.Text), 2) + " / " + FormatNumber((samlet2 + samlet2 * pobjFilter_tb_avance.GetValue / 100), 0)
        End If
        If (Val(tb_antal3.Text) > 0) Then
            lb_timer3.Text = FormatNumber(Convert.ToDouble(lb_mand3.Text) + Convert.ToDouble(lb_cnc3.Text), 2)
            lb_samlet3.Text = FormatNumber(((Convert.ToDouble(lb_timer3.Text) + Convert.ToDouble(lb_indkøb3.Text) + Convert.ToDouble(lb_råvarer3.Text)) / Convert.ToDouble(tb_antal3.Text)), 2) + " / " + FormatNumber(Convert.ToDouble(lb_timer3.Text) + Convert.ToDouble(lb_indkøb3.Text) + Convert.ToDouble(lb_råvarer3.Text), 0)
            samlet3 = Convert.ToDouble(lb_timer3.Text) + Convert.ToDouble(lb_indkøb3.Text) + Convert.ToDouble(lb_råvarer3.Text)
            lb_salg3.Text = FormatNumber((samlet3 + samlet3 * pobjFilter_tb_avance.GetValue / 100) / Convert.ToDouble(tb_antal3.Text), 2) + " / " + FormatNumber((samlet3 + samlet3 * pobjFilter_tb_avance.GetValue / 100), 0)
        End If
        If (Val(tb_antal4.Text) > 0) Then
            lb_timer4.Text = FormatNumber(Convert.ToDouble(lb_mand4.Text) + Convert.ToDouble(lb_cnc4.Text), 2)
            lb_samlet4.Text = FormatNumber(((Convert.ToDouble(lb_timer4.Text) + Convert.ToDouble(lb_indkøb4.Text) + Convert.ToDouble(lb_råvarer4.Text)) / Convert.ToDouble(tb_antal4.Text)), 2) + " / " + FormatNumber(Convert.ToDouble(lb_timer4.Text) + Convert.ToDouble(lb_indkøb4.Text) + Convert.ToDouble(lb_råvarer4.Text), 0)
            samlet4 = Convert.ToDouble(lb_timer4.Text) + Convert.ToDouble(lb_indkøb4.Text) + Convert.ToDouble(lb_råvarer4.Text)
            lb_salg4.Text = FormatNumber((samlet4 + samlet4 * pobjFilter_tb_avance.GetValue / 100) / Convert.ToDouble(tb_antal4.Text), 2) + " / " + FormatNumber((samlet4 + samlet4 * pobjFilter_tb_avance.GetValue / 100), 0)
        End If
        If (Val(tb_antal5.Text) > 0) Then
            lb_timer5.Text = FormatNumber(Convert.ToDouble(lb_mand5.Text) + Convert.ToDouble(lb_cnc5.Text), 2)
            lb_samlet5.Text = FormatNumber(((Convert.ToDouble(lb_timer5.Text) + Convert.ToDouble(lb_indkøb5.Text) + Convert.ToDouble(lb_råvarer5.Text)) / Convert.ToDouble(tb_antal5.Text)), 2) + " / " + FormatNumber(Convert.ToDouble(lb_timer5.Text) + Convert.ToDouble(lb_indkøb5.Text) + Convert.ToDouble(lb_råvarer5.Text), 0)
            samlet5 = Convert.ToDouble(lb_timer5.Text) + Convert.ToDouble(lb_indkøb5.Text) + Convert.ToDouble(lb_råvarer5.Text)
            lb_salg5.Text = FormatNumber((samlet5 + samlet5 * pobjFilter_tb_avance.GetValue / 100) / Convert.ToDouble(tb_antal5.Text), 2) + " / " + FormatNumber((samlet5 + samlet5 * pobjFilter_tb_avance.GetValue / 100), 0)
        End If

    End Sub


    Public Sub CalculateAll()
        Dim subjectweight As Double
        Dim stopvalue As Double


        If pbLoading = True Then
            Exit Sub
        End If
        If (pobjFilter_tb_bukmax_x.GetValue = 0 Or pobjFilter_tb_bukmax_Y.GetValue = 0 Or pobjFilter_tb_pladetykkelse.GetValue = 0) Then
            Exit Sub
        End If
        ClearAllCalculations()
        CalculateSubjectSize()
        stopvalue = CalculateBestPlate(tb_antal1.Text)
        If stopvalue = 1 Then
            Exit Sub
        End If
        subjectweight = CalculateSubjectweight()



        If Val(tb_buk1_x.Text) + Val(tb_buk1_y.Text) <> 0 Then
            CalculateBendTime(subjectweight)
        End If
        If rb_klip.Checked = True Then
            CalculateCuttingTime()
        End If
        If rb_D_stans.Checked = True Then
            CalculatePunchTime()
        End If
        If rb_C_laser.Checked = True Then
            CalculateLaserTime()
        End If
        If rb_B_kombi.Checked = True Then
            CalculateCombiTime()
        End If
        CalculateThreads(tb_antal1.Text)
        CalculateCountersinks(tb_antal1.Text)
        CalculatePressnut(Val(tb_antal1.Text), Val(tb_presmøtrik_antal.Text))
        CalculateBoltweld(Val(tb_antal1.Text), Val(tb_boltesvejs_antal.Text))
        CalculateDeburring(Val(tb_antal1.Text))
        CalculateSteelmaster(Val(tb_antal1.Text), lb_udfold_x.Text, lb_udfold_y.Text)
        CalculateBrushing(Val(tb_antal1.Text))
        CalculateStraigthening(Val(tb_antal1.Text), lb_udfold_x.Text, lb_udfold_y.Text, tb_pladetykkelse.Text)
        CalculateWelding(tb_antal1.Text)
        CalculateSubtotals()
        CalculatePurchasingSubtotal()
        CalculateTotalPrice()
    End Sub

    Private Function CalculateSubjectweight()
        Dim objDatabaseIO As DatabaseIO
        Dim LengthX As Double
        Dim LengthY As Double
        Dim thickness As Double
        Dim materialgroup As Integer
        Dim materialweight As Double
        Dim dblSubjectWeight As Double

        objDatabaseIO = New DatabaseIO
        LengthX = Val(lb_udfold_x.Text)

        LengthY = Val(lb_udfold_y.Text)
        thickness = Val(tb_pladetykkelse.Text)
        materialgroup = Val(Lb_matrgruppe.Text)

        materialweight = objDatabaseIO.GetMaterialWeight(materialgroup)

        dblSubjectWeight = LengthX * LengthY / 1000000 * materialweight * thickness

        lb_emnevægt.Text = dblSubjectWeight
        CalculateSubjectweight = dblSubjectWeight

    End Function
    Public Function CalculateSurfaceTreatment(ByVal antal As Double)
        Dim surfacetreatmentcost As Double
        Dim surfacetreatment As Double
        Dim antal100 As Double

        antal100 = 100


        surfacetreatment = 0
        If Val(tb_overfl_pris1.Text) <> 0 Then
            surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart1.GetValue), Val(pobjFilter_tb_overfl_afdæk1.GetValue), Val(pobjFilter_tb_overfl_Pris1.GetValue), Val(pobjFilter_tb_overfl_Avance1.GetValue), antal)
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_1.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart1.GetValue), Val(pobjFilter_tb_overfl_afdæk1.GetValue), Val(pobjFilter_tb_overfl_Pris1.GetValue), Val(pobjFilter_tb_overfl_Avance1.GetValue), antal100)) / antal100, 2)
        End If
        If Val(tb_overfl_pris2.Text) <> 0 Then
            surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart2.GetValue), Val(pobjFilter_tb_overfl_afdæk2.GetValue), Val(pobjFilter_tb_overfl_Pris2.GetValue), Val(pobjFilter_tb_overfl_Avance2.GetValue), antal)
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_2.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart2.GetValue), Val(pobjFilter_tb_overfl_afdæk2.GetValue), Val(pobjFilter_tb_overfl_Pris2.GetValue), Val(pobjFilter_tb_overfl_Avance2.GetValue), antal100)) / antal100, 2)
        End If
        If Val(tb_overfl_pris3.Text) <> 0 Then
            surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart3.GetValue), Val(pobjFilter_tb_overfl_afdæk3.GetValue), Val(pobjFilter_tb_overfl_Pris3.GetValue), Val(pobjFilter_tb_overfl_Avance3.GetValue), antal)
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_3.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart3.GetValue), Val(pobjFilter_tb_overfl_afdæk3.GetValue), Val(pobjFilter_tb_overfl_Pris3.GetValue), Val(pobjFilter_tb_overfl_Avance3.GetValue), antal100)) / antal100, 2)
        End If
        If Val(tb_overfl_pris4.Text) <> 0 Then
            surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart4.GetValue), Val(pobjFilter_tb_overfl_afdæk4.GetValue), Val(pobjFilter_tb_overfl_Pris4.GetValue), Val(pobjFilter_tb_overfl_Avance4.GetValue), antal)
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_4.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart4.GetValue), Val(pobjFilter_tb_overfl_afdæk4.GetValue), Val(pobjFilter_tb_overfl_Pris4.GetValue), Val(pobjFilter_tb_overfl_Avance4.GetValue), antal100)) / antal100, 2)
        End If
        CalculateSurfaceTreatment = surfacetreatmentcost
    End Function

    Public Function CalculateSurfaceTreatmentCost(ByVal opstart As Double, ByVal afdækning As Double, ByVal pris As Double, ByVal avance As Double, ByVal antal As Double)
        CalculateSurfaceTreatmentCost = opstart + ((afdækning + pris) * antal)

        CalculateSurfaceTreatmentCost = CalculateSurfaceTreatmentCost + CalculateSurfaceTreatmentCost * avance / 100


    End Function
    Private Sub CalculatePunchTime()
        Dim objCombiPunchCalculations As PunchCalculations
        Dim Difficultfactor As Double
        Dim startup As Double
        Dim objDatabaseIO As DatabaseIO


        objCombiPunchCalculations = New PunchCalculations
        lb_gruppe1_tid.Text = objCombiPunchCalculations.CalculatePunchTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_toolshift.Text), Val(tb_slag_til_huller.Text), Val(Convert.ToDouble(lb_faktor.Text)))
        objDatabaseIO = New DatabaseIO

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        startup = Difficultfactor * objDatabaseIO.GetPunchStarup

        lb_gruppe1_opstart.Text = startup



    End Sub


    Private Sub CalculateLaserTime()
        Dim objLaserCalculations As LaserCalculations
        Dim Difficultfactor As Double
        Dim objDatabaseIO As DatabaseIO
        Dim startup As Double


        objLaserCalculations = New LaserCalculations
        lb_laserCNC_tid.Text = objLaserCalculations.CalculateLaserTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_hulantal_1C.Text), Val(tb_hulantal_2C.Text), Val(tb_hulantal_3C.Text), Val(tb_cuttinglength_C.Text), Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text), pobjListMaterials.GetMaterialType(), Val(Convert.ToDouble(lb_faktor.Text)))
        objDatabaseIO = New DatabaseIO

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        startup = Difficultfactor * objDatabaseIO.GetLaserStartup

        lb_laser_opstart.Text = startup


    End Sub

    Private Sub CalculateCombiTime()
        Dim objCombiLaserCalculations As CombiLaserCalculations
        Dim objCombiPunchCalculations As CombiPunchCalculations
        Dim objDatabaseIO As DatabaseIO
        Dim PunchtimeCNC As Double
        Dim LasertimeCNC As Double
        Dim Difficultfactor As Double

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        objCombiLaserCalculations = New CombiLaserCalculations
        LasertimeCNC = objCombiLaserCalculations.CalculateLaserTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_hulantal_1B.Text), Val(tb_hulantal_2B.Text), Val(tb_hulantal_3B.Text), Val(tb_cuttinglength_B.Text), Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text), Val(Convert.ToDouble(lb_faktor.Text)))

        objCombiPunchCalculations = New CombiPunchCalculations
        PunchtimeCNC = objCombiPunchCalculations.CalculatePunchTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_toolshift_B.Text), Val(tb_slag_til_huller_B.Text), Val(Convert.ToDouble(lb_faktor.Text)))


        objDatabaseIO = New DatabaseIO

        lb_Combi_opstart.Text = objDatabaseIO.GetCombiStartup * Difficultfactor

        lb_CombiCNC_tid.Text = LasertimeCNC + PunchtimeCNC



    End Sub
    Private Sub CalculateCuttingTime()


        Dim objCuttingCalculations As CuttingCalculations
        Dim CountRow As Integer
        Dim Subjects1Sheet As Integer
        Dim NumberofSheets As Integer
        Dim Cuttingtime As Double
        Dim Difficultfactor As Double

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        CountRow = Convert.ToInt32(lb_rækkeantal.Text)
        Subjects1Sheet = Val(lb_emner_prplade.Text)
        NumberofSheets = Val(lb_antalplader.Text)

        objCuttingCalculations = New CuttingCalculations
        Cuttingtime = objCuttingCalculations.CalculateCuttingTime(Val(lb_rækkeantal.Text), Val(lb_emner_prplade.Text), Val(lb_antalplader.Text), Val(tb_pladetykkelse.Text), Val(lb_modulstr.Text))


        lb_klip_opstart.Text = Difficultfactor * objCuttingCalculations.GetCuttingStartup(Val(lb_antalplader.Text), Val(tb_pladetykkelse.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text))
        lb_klip_tid.Text = Difficultfactor * Cuttingtime.ToString()



    End Sub
    Public Sub New()
        pbLoading = True
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        pobjListMaterials = New ListMaterials(cb_materiale, Lb_matrgruppe, Lb_klasse, Lb_Kilopris)
        pobjListMaterials.List()
        pobjListSuppliers1 = New ListSuppliers(cb_overfl_leverandør1)
        pobjListSuppliers1.List()
        pobjListSuppliers2 = New ListSuppliers(cb_overfl_leverandør2)
        pobjListSuppliers2.List()
        pobjListSuppliers3 = New ListSuppliers(cb_overfl_leverandør3)
        pobjListSuppliers3.List()
        'pobjListSuppliers4 = New ListSuppliers(cb_overfl_leverandør4)
        'pobjListSuppliers4.List()
        pobjListsurface1 = New Listsurfaces(cb_overfl_beh1)
        pobjListsurface1.List()
        pobjListsurface2 = New Listsurfaces(cb_overfl_beh2)
        pobjListsurface2.List()
        pobjListsurface3 = New Listsurfaces(cb_overfl_beh3)
        pobjListsurface3.List()
        'pobjListsurface4 = New Listsurfaces(cb_overfl_beh4)
        'pobjListsurface4.List()

        pobjFilter_tb_tilbud1 = New FilterKeys(tb_tilbud1, Nothing, Me)
        pobjFilter_tb_tilbud2 = New FilterKeys(tb_tilbud2, Nothing, Me)
        pobjFilter_tb_tilbud3 = New FilterKeys(tb_tilbud3, Nothing, Me)
        pobjFilter_tb_tilbud4 = New FilterKeys(tb_tilbud4, Nothing, Me)
        pobjFilter_tb_tilbud5 = New FilterKeys(tb_tilbud5, Nothing, Me)
        pobjFilter_tb_avance = New FilterKeys(tb_avance, Nothing, Me)

        'buk
        pobjFilter_tb_bukmax_x = New FilterKeys(tb_bukmax_x, Nothing, Me)
        pobjFilter_tb_bukmax_Y = New FilterKeys(tb_bukmax_y, Nothing, Me)
        pobjFilter_tb_buk1_x = New FilterKeys(tb_buk1_x, Nothing, Me)
        pobjFilter_tb_buk1_Y = New FilterKeys(tb_buk1_y, Nothing, Me)
        pobjFilter_tb_buk2_x = New FilterKeys(tb_buk2_x, Nothing, Me)
        pobjFilter_tb_buk2_Y = New FilterKeys(tb_buk2_y, Nothing, Me)
        pobjFilter_tb_buk3_x = New FilterKeys(tb_buk3_x, Nothing, Me)
        pobjFilter_tb_buk3_Y = New FilterKeys(tb_buk3_y, Nothing, Me)
        pobjFilter_tb_buk4_x = New FilterKeys(tb_buk4_x, Nothing, Me)
        pobjFilter_tb_buk4_Y = New FilterKeys(tb_buk4_y, Nothing, Me)
        pobjFilter_tb_buk5_x = New FilterKeys(tb_buk5_x, Nothing, Me)
        pobjFilter_tb_buk5_Y = New FilterKeys(tb_buk5_y, Nothing, Me)
        pobjFilter_tb_buk6_x = New FilterKeys(tb_buk6_x, Nothing, Me)
        pobjFilter_tb_buk6_Y = New FilterKeys(tb_buk6_y, Nothing, Me)
        pobjFilter_tb_buk7_x = New FilterKeys(tb_buk7_x, Nothing, Me)
        pobjFilter_tb_buk7_Y = New FilterKeys(tb_buk7_y, Nothing, Me)
        pobjFilter_tb_buk8_x = New FilterKeys(tb_buk8_x, Nothing, Me)
        pobjFilter_tb_buk8_Y = New FilterKeys(tb_buk8_y, Nothing, Me)
        pobjFilter_tb_buk9_x = New FilterKeys(tb_buk9_x, Nothing, Me)
        pobjFilter_tb_buk9_Y = New FilterKeys(tb_buk9_y, Nothing, Me)
        pobjFilter_tb_buk10_x = New FilterKeys(tb_buk10_x, Nothing, Me)
        pobjFilter_tb_buk10_Y = New FilterKeys(tb_buk10_y, Nothing, Me)
        pobjFilter_tb_buk11_x = New FilterKeys(tb_buk11_x, Nothing, Me)
        pobjFilter_tb_buk11_Y = New FilterKeys(tb_buk11_y, Nothing, Me)
        pobjFilter_tb_buk_uk = New FilterKeys(tb_buk_uk, lb_buk_tid, Me)
        pobjFilter_tb_buk_opst_uk = New FilterKeys(tb_buk_opst_uk, lb_buk_opst, Me)
        'stans
        pobjFilter_tb_toolshift = New FilterKeys(tb_toolshift, Nothing, Me)
        pobjFilter_tb_slag_til_huller = New FilterKeys(tb_slag_til_huller, Nothing, Me)
        pobjFilter_tb_CNCmin_uk = New FilterKeys(tb_CNCmin_uk, lb_gruppe1_tid, Me)
        pobjFilter_tb_gruppe1_opstart_uk = New FilterKeys(tb_gruppe1_opstart_uk, lb_gruppe1_opstart, Me)
        'laser
        pobjFilter_tb_hulantal_1C = New FilterKeys(tb_hulantal_1C, Nothing, Me)
        pobjFilter_tb_hulantal_2C = New FilterKeys(tb_hulantal_2C, Nothing, Me)
        pobjFilter_tb_hulantal_3C = New FilterKeys(tb_hulantal_3C, Nothing, Me)
        pobjFilter_tb_cuttinglength_C = New FilterKeys(tb_cuttinglength_C, Nothing, Me)
        pobjFilter_tb_laserCNC_tid_uk = New FilterKeys(tb_laserCNC_tid_uk, lb_laserCNC_tid, Me)
        pobjFilter_tb_laser_opstart_uk = New FilterKeys(tb_laser_opstart_uk, lb_laser_opstart, Me)
        'kombi
        pobjFilter_tb_toolshift_B = New FilterKeys(tb_toolshift_B, Nothing, Me)
        pobjFilter_tb_slag_til_huller_B = New FilterKeys(tb_slag_til_huller_B, Nothing, Me)
        pobjFilter_tb_hulantal_1B = New FilterKeys(tb_hulantal_1B, Nothing, Me)
        pobjFilter_tb_hulantal_2B = New FilterKeys(tb_hulantal_2B, Nothing, Me)
        pobjFilter_tb_hulantal_3B = New FilterKeys(tb_hulantal_3B, Nothing, Me)
        pobjFilter_tb_cuttinglength_B = New FilterKeys(tb_cuttinglength_B, Nothing, Me)
        pobjFilter_tb_combiCNC_tid_uk = New FilterKeys(tb_combiCNC_tid_uk, lb_CombiCNC_tid, Me)
        pobjFilter_tb_combi_opstart_uk = New FilterKeys(tb_combi_opstart_uk, lb_Combi_opstart, Me)
        'klip
        pobjFilter_tb_klip_tid_uk = New FilterKeys(tb_klip_tid_uk, lb_klip_tid, Me)
        pobjFilter_tb_klip_opstart_uk = New FilterKeys(tb_klip_opstart_uk, lb_klip_opstart, Me)
        'gruppe2
        pobjFilter_tb_afgrat_uk = New FilterKeys(tb_afgrat_uk, lb_afgrat, Me)
        pobjFilter_tb_grinding_uk = New FilterKeys(tb_grinding_uk, lb_grinding, Me)
        pobjFilter_tb_brush_uk = New FilterKeys(tb_brush_uk, lb_brush, Me)
        pobjFilter_tb_vibration_uk = New FilterKeys(tb_vibration_uk, lb_vibration, Me)
        pobjFilter_tb_rette_uk = New FilterKeys(tb_rette_uk, lb_rette, Me)
        pobjFilter_tb_stans_manuel_uk = New FilterKeys(tb_stans_manuel_uk, lb_stans_manuel, Me)
        pobjFilter_tb_countersink_uk = New FilterKeys(tb_countersink_uk, lb_countersink, Me)
        pobjFilter_tb_gevind_uk = New FilterKeys(tb_gevind_uk, lb_gevind, Me)
        pobjFilter_tb_pressnut_uk = New FilterKeys(tb_pressnut_uk, lb_pressnut, Me)
        pobjFilter_tb_pressnut_kr_uk = New FilterKeys(tb_pressnut_kr_uk, lb_pressnut_kr, Me)
        pobjFilter_tb_boltesvejs_uk = New FilterKeys(tb_boltesvejs_uk, lb_boltesvejs, Me)
        pobjFilter_tb_svejsestag_kr_uk = New FilterKeys(tb_svejsestag_kr_uk, lb_svejsestag_kr, Me)
        pobjFilter_tb_m2 = New FilterKeys(tb_m2, Nothing, Me)
        pobjFilter_tb_m2_5 = New FilterKeys(tb_m2_5, Nothing, Me)
        pobjFilter_tb_m3 = New FilterKeys(tb_m3, Nothing, Me)
        pobjFilter_tb_m4 = New FilterKeys(tb_m4, Nothing, Me)
        pobjFilter_tb_m5 = New FilterKeys(tb_m5, Nothing, Me)
        pobjFilter_tb_m6 = New FilterKeys(tb_m6, Nothing, Me)
        pobjFilter_tb_m8 = New FilterKeys(tb_m8, Nothing, Me)
        pobjFilter_tb_m10 = New FilterKeys(tb_m10, Nothing, Me)
        pobjFilter_tb_1 = New FilterKeys(tb_1, Nothing, Me)
        pobjFilter_tb_2 = New FilterKeys(tb_2, Nothing, Me)
        pobjFilter_tb_3 = New FilterKeys(tb_3, Nothing, Me)
        pobjFilter_tb_4 = New FilterKeys(tb_4, Nothing, Me)
        'gruppe3
        pobjFilter_tb_numberofspotwelds = New FilterKeys(tb_numberofspotwelds, Nothing, Me)
        pobjFilter_tb_spotweld_uk = New FilterKeys(tb_spotweld_uk, lb_spotweld, Me)
        'gruppe4
        pobjFilter_tb_numberofwelds = New FilterKeys(tb_numberofwelds, Nothing, Me)
        pobjFilter_tb_weldlength = New FilterKeys(tb_weldlength, Nothing, Me)
        pobjFilter_tb_tackweld_uk = New FilterKeys(tb_tackweld_uk, lb_tackweld, Me)
        pobjFilter_tb_weld_uk = New FilterKeys(tb_weld_uk, lb_weld, Me)
        pobjFilter_tb_grind_weld_uk = New FilterKeys(tb_grind_weld_uk, lb_grind_weld, Me)
        'gruppe5
        pobjFilter_tb_kontor_uk = New FilterKeys(tb_kontor_uk, lb_kontor, Me)
        pobjFilter_tb_kontrol_uk = New FilterKeys(tb_kontrol_uk, lb_kontrol, Me)
        'admin
        pobjFilter_tb_timesats_mand = New FilterKeys(tb_timesats_mand, Nothing, Me)
        pobjFilter_tb_timesats_D = New FilterKeys(tb_timesats_D, Nothing, Me)
        pobjFilter_tb_timesats_B = New FilterKeys(tb_timesats_B, Nothing, Me)
        pobjFilter_tb_timesats_C = New FilterKeys(tb_timesats_C, Nothing, Me)
        'opstart
        pobjFilter_tb_antal_opstart_uk = New FilterKeys(tb_antal_opstart_uk, lb_antal_opstart, Me)
        pobjFilter_tb_opstart_kr_uk = New FilterKeys(tb_opstart_kr_uk, lb_opstart_kr, Me)
        pobjFilter_tb_antal_program_uk = New FilterKeys(tb_antal_program_uk, lb_antal_program, Me)
        pobjFilter_tb_program_kr_uk = New FilterKeys(tb_Program_kr_uk, lb_program_kr, Me)
        pobjFilter_tb_opstart_avance = New FilterKeys(tb_opstart_avance, Nothing, Me)
        pobjFilter_tb_opstart_afgivettilbud = New FilterKeys(tb_opstart_afgivettilbud, Nothing, Me)
        pobjFilter_tb_sværhed_uk = New FilterKeys(tb_sværhed_uk, lb_sværhed, Me)

        'indkøb
        pobjFilter_tb_overfl_opstart1 = New FilterKeys(tb_overfl_opstart1, Nothing, Me)
        pobjFilter_tb_overfl_opstart2 = New FilterKeys(tb_overfl_opstart2, Nothing, Me)
        pobjFilter_tb_overfl_opstart3 = New FilterKeys(tb_overfl_opstart3, Nothing, Me)
        pobjFilter_tb_overfl_opstart4 = New FilterKeys(tb_overfl_opstart4, Nothing, Me)
        pobjFilter_tb_overfl_afdæk1 = New FilterKeys(tb_overfl_afdæk1, Nothing, Me)
        pobjFilter_tb_overfl_afdæk2 = New FilterKeys(tb_overfl_afdæk2, Nothing, Me)
        pobjFilter_tb_overfl_afdæk3 = New FilterKeys(tb_overfl_afdæk3, Nothing, Me)
        pobjFilter_tb_overfl_afdæk4 = New FilterKeys(tb_overfl_afdæk4, Nothing, Me)
        pobjFilter_tb_overfl_Pris1 = New FilterKeys(tb_overfl_pris1, Nothing, Me)
        pobjFilter_tb_overfl_Pris2 = New FilterKeys(tb_overfl_pris2, Nothing, Me)
        pobjFilter_tb_overfl_Pris3 = New FilterKeys(tb_overfl_pris3, Nothing, Me)
        pobjFilter_tb_overfl_Pris4 = New FilterKeys(tb_overfl_pris4, Nothing, Me)
        pobjFilter_tb_overfl_Avance1 = New FilterKeys(tb_overfl_avance1, Nothing, Me)
        pobjFilter_tb_overfl_Avance2 = New FilterKeys(tb_overfl_avance2, Nothing, Me)
        pobjFilter_tb_overfl_Avance3 = New FilterKeys(tb_overfl_avance3, Nothing, Me)
        pobjFilter_tb_overfl_Avance4 = New FilterKeys(tb_overfl_avance4, Nothing, Me)

        pobjFilter_tb_pladetykkelse = New FilterKeys(tb_pladetykkelse, Nothing, Me)
        pbLoading = False
        lb_dato.Text = Now
        lb_operatør.Text = System.Environment.UserName
    End Sub
    'OPDATERING AF UDREGNINGER VED ÆNDRINGER

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_afgrat.CheckedChanged

    End Sub

    Private Sub lb_børste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_brush.Click

    End Sub

    Private Sub tb_antal2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_antal2.TextChanged

        If tb_antal2.Text = "" Then
            lb_mand2.Text = ""
            lb_cnc2.Text = ""
            lb_mand2_tid.Text = ""
            lb_CNC2_tid.Text = ""
            lb_timer2.Text = ""
            lb_indkøb2.Text = ""
            lb_råvarer2.Text = ""
            lb_samlet2.Text = ""
            lb_salg2.Text = ""
            tb_tilbud2.Text = ""
        Else
            CalculateAll()
        End If
    End Sub
    Private Sub tb_antal3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_antal3.TextChanged

        If tb_antal3.Text = "" Then
            lb_mand3.Text = ""
            lb_cnc3.Text = ""
            lb_mand3_tid.Text = ""
            lb_CNC3_tid.Text = ""
            lb_timer3.Text = ""
            lb_indkøb3.Text = ""
            lb_råvarer3.Text = ""
            lb_samlet3.Text = ""
            lb_salg3.Text = ""
            tb_tilbud3.Text = ""
        Else
            CalculateAll()
        End If
    End Sub
    Private Sub tb_antal4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_antal4.TextChanged
        If tb_antal4.Text = "" Then
            lb_mand4.Text = ""
            lb_cnc4.Text = ""
            lb_mand4_tid.Text = ""
            lb_CNC4_tid.Text = ""
            lb_timer4.Text = ""
            lb_indkøb4.Text = ""
            lb_råvarer4.Text = ""
            lb_samlet4.Text = ""
            lb_salg4.Text = ""
            tb_tilbud4.Text = ""
        Else
            CalculateAll()
        End If
    End Sub
    Private Sub tb_antal5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_antal5.TextChanged
        If tb_antal5.Text = "" Then
            lb_mand5.Text = ""
            lb_cnc5.Text = ""
            lb_mand5_tid.Text = ""
            lb_CNC5_tid.Text = ""
            lb_timer5.Text = ""
            lb_indkøb5.Text = ""
            lb_råvarer5.Text = ""
            lb_samlet5.Text = ""
            lb_salg5.Text = ""
            tb_tilbud5.Text = ""
        Else
            CalculateAll()
        End If
    End Sub

    Private Sub rb_D_stans_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_D_stans.CheckedChanged
        CalculateAll()
    End Sub

    Private Sub cmdRecalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CalculateAll()
    End Sub

    Private Sub rb_C_laser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_C_laser.CheckedChanged
        CalculateAll()
    End Sub

    Private Sub rb_B_kombi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_B_kombi.CheckedChanged
        CalculateAll()
    End Sub

    Private Sub rb_klip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_klip.CheckedChanged
        CalculateAll()
    End Sub
    Private Sub rb_netto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_netto.CheckedChanged
        CalculateAll()
    End Sub
    Private Sub rb_brutto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_brutto.CheckedChanged
        CalculateAll()
    End Sub
    Private Sub rb_tig_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_tig.CheckedChanged
        CalculateAll()
    End Sub
    Private Sub tb_numberofwelds_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_numberofwelds.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_weldlength_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_weldlength.TextChanged
        CalculateAll()
    End Sub
    Private Sub cb_weld_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_weld.CheckStateChanged
        CalculateAll()
    End Sub
    Private Sub cb_tackweld_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tackweld.CheckStateChanged
        CalculateAll()
    End Sub
    Private Sub cb_fravælg_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_fravælg.CheckStateChanged
        If cb_fravælg.CheckState = CheckState.Unchecked Then
            'rb_C_laser.Checked = True
        End If
        If cb_fravælg.CheckState = CheckState.Checked Then
            ResetCalculations()
        End If
        CalculateAll()
    End Sub

    Private Sub cb_materiale_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_materiale.SelectedIndexChanged
        CalculateAll()
    End Sub

    Private Sub lb_antal_opstart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_antal_opstart.Click
        CalculateAll()
    End Sub
    Private Sub cb_afgrat_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_afgrat.CheckStateChanged
        CalculateAll()
    End Sub
    Private Sub cb_steelmaster_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_steelmaster.CheckStateChanged
        CalculateAll()
    End Sub
    Private Sub cb_brush_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_brush.CheckStateChanged
        CalculateAll()
    End Sub
    Private Sub cb_rette_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_rette.CheckStateChanged
        CalculateAll()
    End Sub
    Private Sub lb_countersink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_countersink.TextChanged
        CalculateAll()
    End Sub
    Private Sub lb_gevind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_gevind.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m2.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m2_5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m2_5.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m3.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m4.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m5.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m6.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m8.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_m10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m10.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_1.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_2.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_3.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_4.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_presmøtrik_antal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_presmøtrik_antal.TextChanged
        CalculateAll()
    End Sub
    Private Sub tb_boltesvejs_antal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_boltesvejs_antal.TextChanged
        CalculateAll()
    End Sub

    Private Sub bu_Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_Reset.Click
        ResetAll()
    End Sub
    Private Sub bu_udskriv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_udskriv.Click
        'printscreen
    End Sub

    Private Sub ResetCalculations()
        lb_mand1.Text = ""
        lb_cnc1.Text = ""
        lb_mand1_tid.Text = ""
        lb_CNC1_tid.Text = ""
        lb_timer1.Text = ""
        lb_indkøb1.Text = ""
        lb_råvarer1.Text = ""
        lb_samlet1.Text = ""
        lb_salg1.Text = ""
        lb_mand2.Text = ""
        lb_cnc2.Text = ""
        lb_mand2_tid.Text = ""
        lb_CNC2_tid.Text = ""
        lb_timer2.Text = ""
        lb_indkøb2.Text = ""
        lb_råvarer2.Text = ""
        lb_samlet2.Text = ""
        lb_salg2.Text = ""
        lb_mand3.Text = ""
        lb_cnc3.Text = ""
        lb_mand3_tid.Text = ""
        lb_CNC3_tid.Text = ""
        lb_timer3.Text = ""
        lb_indkøb3.Text = ""
        lb_råvarer3.Text = ""
        lb_samlet3.Text = ""
        lb_salg3.Text = ""
        lb_mand4.Text = ""
        lb_cnc4.Text = ""
        lb_mand4_tid.Text = ""
        lb_CNC4_tid.Text = ""
        lb_timer4.Text = ""
        lb_indkøb4.Text = ""
        lb_råvarer4.Text = ""
        lb_samlet4.Text = ""
        lb_salg4.Text = ""
        lb_mand5.Text = ""
        lb_cnc5.Text = ""
        lb_mand5_tid.Text = ""
        lb_CNC5_tid.Text = ""
        lb_timer5.Text = ""
        lb_indkøb5.Text = ""
        lb_råvarer5.Text = ""
        lb_samlet5.Text = ""
        lb_salg5.Text = ""
        lb_gruppe1_tid.Text = ""
        lb_gruppe1_opstart.Text = ""
        tb_CNCmin_uk.Text = ""
        tb_gruppe1_opstart_uk.Text = ""
        tb_laserCNC_tid_uk.Text = ""
        tb_laser_opstart_uk.Text = ""
        tb_combiCNC_tid_uk.Text = ""
        tb_combi_opstart_uk.Text = ""
        tb_klip_tid_uk.Text = ""
        tb_klip_opstart_uk.Text = ""
        tb_tilbud1.Text = ""
        tb_tilbud2.Text = ""
        tb_tilbud3.Text = ""
        tb_tilbud4.Text = ""
        tb_tilbud5.Text = ""
        lb_pladeformatX.Text = ""
        lb_pladeformatY.Text = ""
        lb_antalplader.Text = ""
        lb_emner_prplade.Text = ""
        lb_spild.Text = ""
        lb_emnevægt.Text = ""
        tb_pressnut_uk.Text = ""

    End Sub
    Private Sub ResetAll()
        cb_fravælg.CheckState = CheckState.Checked
        rb_D_stans.Checked = True
        rb_netto.Checked = True
        lb_mand1.Text = ""
        lb_cnc1.Text = ""
        lb_mand1_tid.Text = ""
        lb_CNC1_tid.Text = ""
        lb_timer1.Text = ""
        lb_indkøb1.Text = ""
        lb_råvarer1.Text = ""
        lb_samlet1.Text = ""
        lb_salg1.Text = ""
        lb_mand2.Text = ""
        lb_cnc2.Text = ""
        lb_mand2_tid.Text = ""
        lb_CNC2_tid.Text = ""
        lb_timer2.Text = ""
        lb_indkøb2.Text = ""
        lb_råvarer2.Text = ""
        lb_samlet2.Text = ""
        lb_salg2.Text = ""
        lb_mand3.Text = ""
        lb_cnc3.Text = ""
        lb_mand3_tid.Text = ""
        lb_CNC3_tid.Text = ""
        lb_timer3.Text = ""
        lb_indkøb3.Text = ""
        lb_råvarer3.Text = ""
        lb_samlet3.Text = ""
        lb_salg3.Text = ""
        lb_mand4.Text = ""
        lb_cnc4.Text = ""
        lb_mand4_tid.Text = ""
        lb_CNC4_tid.Text = ""
        lb_timer4.Text = ""
        lb_indkøb4.Text = ""
        lb_råvarer4.Text = ""
        lb_samlet4.Text = ""
        lb_salg4.Text = ""
        lb_mand5.Text = ""
        lb_cnc5.Text = ""
        lb_mand5_tid.Text = ""
        lb_CNC5_tid.Text = ""
        lb_timer5.Text = ""
        lb_indkøb5.Text = ""
        lb_råvarer5.Text = ""
        lb_samlet5.Text = ""
        lb_salg5.Text = ""
        tb_pladetykkelse.Text = ""
        tb_antal2.Text = ""
        tb_antal3.Text = ""
        tb_antal4.Text = ""
        tb_antal5.Text = ""
        cb_kunde.Text = ""
        tb_emne.Text = ""
        cb_Tegning.Text = ""
        tb_revision.Text = ""
        cb_materiale.Text = "1 Jern, P01 AM, 12.03"
        tb_avance.Text = 20
        tb_opstart_avance.Text = 20
        tb_bukmax_x.Text = ""
        tb_bukmax_y.Text = ""
        tb_buk1_x.Text = ""
        tb_buk2_x.Text = ""
        tb_buk3_x.Text = ""
        tb_buk4_x.Text = ""
        tb_buk5_x.Text = ""
        tb_buk6_x.Text = ""
        tb_buk7_x.Text = ""
        tb_buk8_x.Text = ""
        tb_buk9_x.Text = ""
        tb_buk10_x.Text = ""
        tb_buk11_x.Text = ""
        tb_buk1_y.Text = ""
        tb_buk2_y.Text = ""
        tb_buk3_y.Text = ""
        tb_buk4_y.Text = ""
        tb_buk5_y.Text = ""
        tb_buk6_y.Text = ""
        tb_buk7_y.Text = ""
        tb_buk8_y.Text = ""
        tb_buk9_y.Text = ""
        tb_buk10_y.Text = ""
        tb_buk11_y.Text = ""
        lb_udfold_x.Text = ""
        lb_udfold_y.Text = ""
        lb_buk_tid.Text = ""
        lb_buk_opst.Text = ""
        tb_buk_uk.Text = ""
        tb_buk_opst_uk.Text = ""
        tb_toolshift.Text = ""
        tb_slag_til_huller.Text = ""
        lb_gruppe1_tid.Text = ""
        lb_gruppe1_opstart.Text = ""
        tb_CNCmin_uk.Text = ""
        tb_gruppe1_opstart_uk.Text = ""
        tb_hulantal_1C.Text = ""
        tb_hulantal_2C.Text = ""
        tb_hulantal_3C.Text = ""
        tb_cuttinglength_C.Text = ""
        tb_laserCNC_tid_uk.Text = ""
        tb_laser_opstart_uk.Text = ""
        tb_toolshift_B.Text = ""
        tb_slag_til_huller_B.Text = ""
        tb_hulantal_1B.Text = ""
        tb_hulantal_2B.Text = ""
        tb_hulantal_3B.Text = ""
        tb_cuttinglength_B.Text = ""
        tb_combiCNC_tid_uk.Text = ""
        tb_combi_opstart_uk.Text = ""
        tb_klip_tid_uk.Text = ""
        tb_klip_opstart_uk.Text = ""
        cb_afgrat.CheckState = CheckState.Unchecked
        lb_afgrat.Text = ""
        lb_grinding.Text = ""
        cb_brush.CheckState = CheckState.Unchecked
        lb_brush.Text = ""
        lb_vibration.Text = ""
        lb_rette.Text = ""
        cb_afgrat.Checked = False
        cb_rette.Checked = False
        cb_vibrationsafgr.Checked = False
        cb_brush.Checked = False
        cb_steelmaster.Checked = False
        lb_countersink.Text = ""
        lb_gevind.Text = ""
        lb_pressnut.Text = ""
        lb_boltesvejs.Text = ""
        lb_undersænk_antal.Text = ""
        lb_gevind_antal.Text = ""
        tb_presmøtrik_antal.Text = ""
        tb_boltesvejs_antal.Text = ""
        tb_afgrat_uk.Text = ""
        tb_grinding_uk.Text = ""
        tb_brush_uk.Text = ""
        tb_vibration_uk.Text = ""
        tb_rette_uk.Text = ""
        tb_stans_manuel_uk.Text = ""
        tb_countersink_uk.Text = ""
        tb_gevind_uk.Text = ""
        tb_pressnut_uk.Text = ""
        tb_boltesvejs_uk.Text = ""
        lb_antal_opstart.Text = ""
        lb_opstart_kr.Text = ""
        lb_antal_program.Text = ""
        lb_program_kr.Text = ""
        tb_antal_opstart_uk.Text = ""
        tb_opstart_kr_uk.Text = ""
        tb_antal_program_uk.Text = ""
        tb_Program_kr_uk.Text = ""
        lb_opstart_program.Text = ""
        lb_opst_prog_brutto.Text = ""
        tb_opstart_afgivettilbud.Text = ""
        tb_tilbud1.Text = ""
        tb_tilbud2.Text = ""
        tb_tilbud3.Text = ""
        tb_tilbud4.Text = ""
        tb_tilbud5.Text = ""
        tb_m2.Text = ""
        tb_m2_5.Text = ""
        tb_m3.Text = ""
        tb_m4.Text = ""
        tb_m5.Text = ""
        tb_m6.Text = ""
        tb_m8.Text = ""
        tb_m10.Text = ""
        tb_1.Text = ""
        tb_2.Text = ""
        tb_3.Text = ""
        tb_4.Text = ""
        tb_numberofspotwelds.Text = ""
        lb_spotweld.Text = ""
        tb_spotweld_uk.Text = ""
        tb_numberofwelds.Text = ""
        tb_weldlength.Text = ""
        cb_tackweld.CheckState = CheckState.Unchecked
        cb_weld.CheckState = CheckState.Unchecked
        cb_grind_weld.CheckState = CheckState.Unchecked
        lb_tackweld.Text = ""
        lb_weld.Text = ""
        lb_grind_weld.Text = ""
        tb_tackweld_uk.Text = ""
        tb_weld_uk.Text = ""
        tb_grind_weld_uk.Text = ""
        rb_tig.Checked = True
        lb_kontor.Text = "60"
        lb_kontrol.Text = "60"
        tb_kontor_uk.Text = ""
        tb_kontrol_uk.Text = ""
        cb_overfl_leverandør1.Text = "0 Ingen Overfladebehandling"
        cb_overfl_leverandør2.Text = "0 Ingen Overfladebehandling"
        cb_overfl_leverandør3.Text = "0 Ingen Overfladebehandling"
        cb_overfl_leverandør4.Text = "0 Ingen Overfladebehandling"
        cb_overfl_beh1.Text = ""
        cb_overfl_beh2.Text = ""
        cb_overfl_beh3.Text = ""
        cb_overfl_beh4.Text = ""
        tb_overfl_opstart1.Text = ""
        tb_overfl_opstart2.Text = ""
        tb_overfl_opstart3.Text = ""
        tb_overfl_opstart4.Text = ""
        tb_overfl_afdæk1.Text = ""
        tb_overfl_afdæk2.Text = ""
        tb_overfl_afdæk3.Text = ""
        tb_overfl_afdæk4.Text = ""
        tb_overfl_pris1.Text = ""
        tb_overfl_pris2.Text = ""
        tb_overfl_pris3.Text = ""
        tb_overfl_pris4.Text = ""
        tb_overfl_avance1.Text = 15
        tb_overfl_avance2.Text = 15
        tb_overfl_avance3.Text = 15
        tb_overfl_avance4.Text = 15
        tb_overfl_pris100_1.Text = ""
        tb_overfl_pris100_2.Text = ""
        tb_overfl_pris100_3.Text = ""
        tb_overfl_pris100_4.Text = ""
        lb_pladeformatX.Text = ""
        lb_pladeformatY.Text = ""
        lb_modulstr.Text = ""
        lb_faktor.Text = ""
        lb_sværhed.Text = ""
        tb_sværhed_uk.Text = ""
        lb_antalplader.Text = ""
        lb_emner_prplade.Text = ""
        lb_spild.Text = ""
        lb_emnevægt.Text = ""
        lb_spildtype.Text = ""
        lb_spildnetto.Text = ""
        tb_timesats_mand.Text = 475
        tb_timesats_D.Text = 850
        tb_timesats_B.Text = 1490
        tb_timesats_C.Text = 1900

    End Sub


    Private Sub lb_antal_opstart_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lb_antal_opstart.LocationChanged

    End Sub

    Private Sub rb_netto_CursorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rb_netto.CursorChanged

    End Sub
End Class

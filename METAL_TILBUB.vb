Imports System.Data.SqlClient
Imports System.Configuration


Public Class metal_tilbud
    Private pobjListMaterials As ListMaterials
    'Private pobjListSuppliers1 As ListSuppliers
    'Private pobjListSuppliers2 As ListSuppliers
    'Private pobjListSuppliers3 As ListSuppliers
    'Private pobjListSuppliers4 As ListSuppliers
    'Private pobjListSuppliers5 As ListSuppliers
    Private pobjListSurface1 As Listsurfaces
    Private pobjListSurface2 As Listsurfaces
    Private pobjListSurface3 As Listsurfaces
    Private pobjListSurface4 As Listsurfaces1
    Private pobjListSurface5 As Listsurfaces1
    Private pobjFilter_tb_tilbud1 As FilterKeys
    Private pobjFilter_tb_tilbud2 As FilterKeys
    Private pobjFilter_tb_tilbud3 As FilterKeys
    Private pobjFilter_tb_tilbud4 As FilterKeys
    Private pobjFilter_tb_avance As FilterKeys
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
    Private pobjFilter_tb_valsning_uk As FilterKeys
    Private pobjFilter_tb_stepbuk_uk As FilterKeys
    Private pobjFilter_tb_stepantal As FilterKeys
    'stans
    Private pobjFilter_tb_toolshift As FilterKeys
    Private pobjFilter_tb_slag_til_huller As FilterKeys
    Private pobjFilter_tb_CNCmin_uk As FilterKeys
    Private pobjFilter_tb_gruppe1_opstart_uk As FilterKeys
    'laser
    Private pobjFilter_tb_hulantal_1C As FilterKeys
    Private pobjFilter_tb_hulantal_2C As FilterKeys
    Private pobjFilter_tb_hulantal_3C As FilterKeys
    Private pobjFilter_tb_hulantal_4C As FilterKeys
    Private pobjFilter_tb_cuttinglength_C As FilterKeys
    Private pobjFilter_tb_laserCNC_tid_uk As FilterKeys
    Private pobjFilter_tb_laser_opstart_uk As FilterKeys
    'kombi
    Private pobjFilter_tb_toolshift_B As FilterKeys
    Private pobjFilter_tb_slag_til_huller_B As FilterKeys
    Private pobjFilter_tb_hulantal_1B As FilterKeys
    Private pobjFilter_tb_hulantal_2B As FilterKeys
    Private pobjFilter_tb_hulantal_3B As FilterKeys
    Private pobjFilter_tb_hulantal_4B As FilterKeys

    Private pobjFilter_tb_cuttinglength_B As FilterKeys
    Private pobjFilter_tb_combiCNC_tid_uk As FilterKeys
    Private pobjFilter_tb_combiCNCstans_tid_uk As FilterKeys
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
    Private pobjFilter_tb_presstag_uk As FilterKeys
    Private pobjFilter_tb_presstag_kr_uk As FilterKeys
    Private pobjFilter_tb_pressnut_kr_uk As FilterKeys
    Private pobjFilter_tb_svejsestag_kr_uk As FilterKeys
    Private pobjFilter_tb_tilsatsmatr_kr_uk As FilterKeys
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
    Private pobjFilter_tb_numberofspots As FilterKeys
    Private pobjFilter_tb_numberofspotweldseams As FilterKeys
    Private pobjFilter_tb_spotweld_uk As FilterKeys
    'gruppe4
    Private pobjFilter_tb_numberofwelds As FilterKeys
    Private pobjFilter_tb_weldlength As FilterKeys
    Private pobjFilter_tb_tackweld_uk As FilterKeys
    Private pobjFilter_tb_weld_uk As FilterKeys
    Private pobjFilter_tb_grind_weld_uk As FilterKeys
    Private pobjFilter_tb_rettesvejs_uk As FilterKeys
    Private pobjFilter_tb_tapantal As FilterKeys
    Private pobjFilter_tb_tapsvejs_uk As FilterKeys
    'gruppe5
    Private pobjFilter_tb_glasbl_uk As FilterKeys
    Private pobjFilter_tb_slib_uk As FilterKeys
    'gruppe6
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
    Private pobjFilter_tb_Kilopris_uk As FilterKeys
    'indkøb
    Private pobjFilter_tb_overfl_opstart1 As FilterKeys
    Private pobjFilter_tb_overfl_opstart2 As FilterKeys
    Private pobjFilter_tb_overfl_opstart3 As FilterKeys
    Private pobjFilter_tb_overfl_opstart4 As FilterKeys
    Private pobjFilter_tb_overfl_opstart5 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk1 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk2 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk3 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk4 As FilterKeys
    Private pobjFilter_tb_overfl_afdæk5 As FilterKeys
    Private pobjFilter_tb_overfl_Pris1 As FilterKeys
    Private pobjFilter_tb_overfl_Pris2 As FilterKeys
    Private pobjFilter_tb_overfl_Pris3 As FilterKeys
    Private pobjFilter_tb_overfl_Pris1_uk As FilterKeys
    Private pobjFilter_tb_overfl_Pris2_uk As FilterKeys
    Private pobjFilter_tb_overfl_Pris3_uk As FilterKeys
    Private pobjFilter_tb_overfl_Pris4 As FilterKeys
    Private pobjFilter_tb_overfl_Pris5 As FilterKeys
    Private pobjFilter_tb_overfl_Avance1 As FilterKeys
    Private pobjFilter_tb_overfl_Avance2 As FilterKeys
    Private pobjFilter_tb_overfl_Avance3 As FilterKeys
    Private pobjFilter_tb_overfl_Avance4 As FilterKeys
    Private pobjFilter_tb_overfl_Avance5 As FilterKeys

    Private pobjFilter_tb_overfl_tilbudpris1_uk As FilterKeys
    Private pobjFilter_tb_overfl_tilbudpris2_uk As FilterKeys
    Private pobjFilter_tb_overfl_tilbudpris3_uk As FilterKeys




    'Private pobjFilter_tb_numberofsurface As FilterKeys


    'Private cnn As System.Data.SqlClient.SqlConnection



    'Private pobjFilter_tb_antal1 As FilterKeys
    'Private pobjFilter_tb_antal2 As FilterKeys
    'Private pobjFilter_tb_antal3 As FilterKeys
    'Private pobjFilter_tb_antal4 As FilterKeys
    'Private pobjFilter_tb_antal5 As FilterKeys


    Private pbLoading As Boolean
    Dim Metaltilbud1 As Object

    'Private Shared ConStr As String = ConfigurationManager.ConnectionStrings("ConStr").ConnectionString

    'Dropdown menu, Customer filled.
    'Private Sub FillCustomerNames()

    '    If ConfigurationManager.AppSettings("ConnectionType").Trim() = "ConStr" Then
    '        Dim cnn As New System.Data.SqlClient.SqlConnection(ConStr)
    '        Dim cmd As New System.Data.SqlClient.SqlCommand(ConfigurationManager.AppSettings("SQLSelect").Trim(), cnn)
    '        Dim dr As System.Data.SqlClient.SqlDataReader = Nothing

    '         Open a connection.
    '        Try
    '            cnn.Open()

    '            dr = cmd.ExecuteReader()
    '            If dr.Read = True Then
    '                While dr.Read()
    '                    cb_kunde.Items.Add(dr.Item(0).ToString())
    '                End While

    '            End If

    '        Finally
    '             Close the connection.
    '            cnn.Close()

    '        End Try

    '    End If'

    ' Creates a connection string used for APPLSRV01\SQLEXPRESS
    Private Function GetConnectionString() As String
        Return "SERVER=APPLSRV01\SQLEXPRESS;User ID=MetalTilbud;Password=MetalTilbud;DATABASE=MetalTilbud;"
        ' Return "SERVER=APPLSRV02-2016\SQLEXPRESS;User ID=MetalTilbud;Password=MetalTilbud;DATABASE=MetalTilbud;"
    End Function

    ' Varibles Used for the sql connection
    Dim cnn As New System.Data.SqlClient.SqlConnection(GetConnectionString())
    Dim cmd As New System.Data.SqlClient.SqlCommand("SELECT ID, Name FROM Customer", cnn)
    Dim dr As System.Data.SqlClient.SqlDataReader = Nothing

    Private Sub FillCustomerNames()
        'Open Sql connection
        Try
            cnn.Open()

            'Add data into sql data into cb_kunde
            dr = cmd.ExecuteReader()
            While dr.Read() = True
                cb_kunde.Items.Add(dr.Item(1).ToString())
            End While
        Catch ex As Exception

        Finally
            cnn.Close()
        End Try

    End Sub

    Private Sub metal_tilbud_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'FillCustomerNames()
    End Sub

    Private Sub cb_kunde_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_kunde.SelectedIndexChanged
        UpdateDrawings()
    End Sub

    Private Sub UpdateDrawings()
    End Sub

    'GEVIND

    '-----------------------------------------------------------------------------------------------------------------


    Private Function CalculateThreads(ByVal number As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim totaltime As Double
        Dim countthread As Integer
        Dim modulesize As Integer

        totaltime = 1

        modulesize = Val(lb_modulstr.Text)

        objDatabaseIO = New DatabaseIO

        countthread = (Val(tb_m2.Text) + Val(tb_m2_5.Text) + Val(tb_m3.Text) + Val(tb_m4.Text) + Val(tb_m5.Text) + Val(tb_m6.Text) + Val(tb_m8.Text) + Val(tb_m10.Text))

        If countthread = 0 Then
            lb_gevind.Text = ""
            lb_gevind_antal.Text = ""
            Exit Function
        End If

        If tb_m2.Text <> "" Then
            totaltime = totaltime + 2
        End If
        If tb_m2_5.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m3.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m4.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m5.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m6.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m8.Text <> "" Then
            totaltime = totaltime + 4
        End If
        If tb_m10.Text <> "" Then
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

        lb_gevind.Text = FormatNumber(CalculateThreads, 0)
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


        calculatethread = (count * objDatabaseIO.getthreadmultiplier(materialID) * (1 + (0.111 * (sheetthickness - 1))))



    End Function


    'UNDERSÆNKNING



    Private Function CalculateCountersinks(ByVal number As Integer) As Double
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
        If tb_1.Text <> "" Then
            totaltime = totaltime + 2
        End If
        If tb_2.Text <> "" Then
            totaltime = totaltime + 5
        End If
        If tb_3.Text <> "" Then
            totaltime = totaltime + 5
        End If
        If tb_4.Text <> "" Then
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
        Countersinks = FormatNumber(totaltime * Difficultfactor, 0)


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
        Dim matrgruppe As Integer

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

        matrgruppe = Val(Lb_matrgruppe.Text)

        If rb_rustfri.Checked = True Then
            matrgruppe = 2
        End If

        objDatabaseIO = New DatabaseIO
        lb_pressnut_kr.Text = objDatabaseIO.getpressnutprice(matrgruppe) * iNumberOfNuts

    End Function
    Private Function CalculatePresstag(ByVal iNumberOfSubjects As Integer, ByVal iNumberOfNuts As Integer) As Double
        Dim objPresstagCalculations As PressnutCalculations
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim matrgruppe As Integer

        If iNumberOfNuts = 0 Then
            lb_presstag.Text = ""
            tb_presstag_uk.Text = ""
            lb_presstag_kr.Text = ""
            tb_presstag_kr_uk.Text = ""
            Exit Function
        End If

        objPresstagCalculations = New PressnutCalculations
        objDatabaseIO = New DatabaseIO

        CalculatePresstag = (objPresstagCalculations.CalculatePresstagTime(lb_modulstr.Text, iNumberOfNuts) / 100 * iNumberOfSubjects) + objDatabaseIO.GetPresstagStartup

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        lb_presstag.Text = CalculatePresstag * Difficultfactor

        matrgruppe = Val(Lb_matrgruppe.Text)

        If rb_rustfri.Checked = True Then
            matrgruppe = 2
        End If

        objDatabaseIO = New DatabaseIO
        lb_presstag_kr.Text = objDatabaseIO.getpresstagprice(matrgruppe) * iNumberOfNuts

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
    Private Function CalculateSteelmaster(ByVal iNumberOfSubjects As Integer, ByVal SubjectX As Double, ByVal SubjectY As Double, ByVal matrgruppe As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO


        If cb_steelmaster.CheckState = CheckState.Unchecked Then
            lb_grinding.Text = ""
            tb_grinding_uk.Text = ""
            Exit Function
        End If

        If matrgruppe = 1 Then

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
        End If
        If matrgruppe = 2 Then
            If SubjectX < 25 Then
                If SubjectY < 25 Then
                    MsgBox("Emnet er for lille til Grindingmasteren")
                    lb_grinding.Text = ""
                    tb_grinding_uk.Text = ""
                    cb_steelmaster.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If
            If SubjectX > 1350 Then
                If SubjectY > 1350 Then
                    MsgBox("Emnet er for stort til Grindingmasteren")
                    lb_grinding.Text = ""
                    tb_grinding_uk.Text = ""
                    cb_steelmaster.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If
        End If
        If matrgruppe = 3 Then
            If SubjectX < 25 Then
                If SubjectY < 25 Then
                    MsgBox("Emnet er for lille til Grindingmasteren")
                    lb_grinding.Text = ""
                    tb_grinding_uk.Text = ""
                    cb_steelmaster.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If
            If SubjectX > 1350 Then
                If SubjectY > 1350 Then
                    MsgBox("Emnet er for stort til Grindingmasteren")
                    lb_grinding.Text = ""
                    tb_grinding_uk.Text = ""
                    cb_steelmaster.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If
        End If
        If matrgruppe = 4 Then
            If SubjectX < 25 Then
                If SubjectY < 25 Then
                    MsgBox("Emnet er for lille til Grindingmasteren")
                    lb_grinding.Text = ""
                    tb_grinding_uk.Text = ""
                    cb_steelmaster.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If
            If SubjectX > 1350 Then
                If SubjectY > 1350 Then
                    MsgBox("Emnet er for stort til Grindingmasteren")
                    lb_grinding.Text = ""
                    tb_grinding_uk.Text = ""
                    cb_steelmaster.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If
            If matrgruppe = 5 Then

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
    Private Sub CalculateAdministration(ByVal Antal As Integer)
        Dim administration As Integer
        Dim kontrol As Integer
        Dim kontor As Integer

        If Antal <= 100 Then
            kontor = 60
        Else
            kontor = 60 + 60 / 100 * Antal * 0.1
        End If
        lb_kontor.Text = kontor

        If Antal <= 100 Then
            kontrol = 60
        Else
            kontrol = 60 + 60 / 100 * Antal * 0.3
        End If
        lb_kontrol.Text = kontrol

        'sammentællinger
        administration = FormatNumber(pobjFilter_tb_kontor_uk.GetValue + pobjFilter_tb_kontrol_uk.GetValue, 0)
        'lb_administration_ialt.Text = administration
        'lb_tekniker.Text = FormatNumber((pobjFilter_tb_kontor_uk.GetValue / 60), 2)
        'lb_control.Text = FormatNumber((pobjFilter_tb_kontrol_uk.GetValue / 60), 2)

    End Sub

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
        Dim iSubjectsPerPlate As Double
        Dim iPlateX As Integer
        Dim iPlateY As Integer
        Dim i As Long


        '-------------------------------------------------
        ' Read input-boxes.
        '-------------------------------------------------
        If cb_fravælg.CheckState = "1" Then
            If cb_fravælg_1500_3000.CheckState = "1" Then
                If cb_fravælg_1250_2500.CheckState = "1" Then
                    If cb_fravælg_1000_2000.CheckState = "1" Then
                        MsgBox("ALLE FORMATER ER FRAVALGT")
                        cb_fravælg_1250_2500.CheckState = "0"
                    End If
                End If
            End If
        End If


        dblSubjectX = lb_udfold_x.Text
        dblSubjectY = lb_udfold_y.Text
        '-------------------------------------------------
        objDatabaseIO = New DatabaseIO
        alPlateSizes = objDatabaseIO.GetPlateSizes(cb_fravælg.CheckState, cb_fravælg_1500_3000.CheckState, cb_fravælg_1250_2500.CheckState, cb_fravælg_1000_2000.CheckState)
        alAllPlateValues = New ArrayList(alPlateSizes.Count)
        For i = 0 To alPlateSizes.Count - 1
            alPlate = alPlateSizes(i)
            alPlateValues = CalculatePlateValues(iNumberOfSubjects, alPlate(0), alPlate(1), dblSubjectX, dblSubjectY)
            alAllPlateValues.Add(alPlateValues)
        Next
        dblLowestWaste = 100000000000000
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
        lb_spild.Text = FormatNumber(dblTotalWaste / 1000000, 2)

        '-------------------------------------------------
    End Function

    Private Function CalculatePlateValues(ByVal iNumberOfSubjects As Double, ByVal PlateX As Double, ByVal PlateY As Double, ByVal SubjectX As Double, ByVal SubjectY As Double) As ArrayList
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
        '(Forberedt til 2 plader pr emne)
        platemultiplier = 1

        '(kombi stans)
        wasteX = 100
        wasteY = 20
        subjectaddition = 20

        If rb_B_kombi.Checked = True Then
            wasteX = 70
            wasteY = 10
            subjectaddition = 10
        End If
        If rb_klip.Checked = True Then
            wasteX = 0
            wasteY = 0
            subjectaddition = 0
        End If
        If rb_C_laser.Checked Then
            wasteX = 10
            wasteY = 10
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
            dblWaste = 100000000000000
            alReturn.Add(iNumberOfPlates)
            alReturn.Add(dblWaste)
            alReturn.Add(iSubjectsPerPlate)
            CalculatePlateValues = alReturn

            Exit Function
        End If


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


        StartwasteX = 100
        StartwasteY = 20
        subjectaddition = 20

        If rb_B_kombi.Checked = True Then
            wasteX = 70
            wasteY = 10
            subjectaddition = 10
        End If

        If rb_klip.Checked = True Then
            StartwasteX = 0
            StartwasteY = 0
            subjectaddition = 0
        End If
        If rb_C_laser.Checked Then
            StartwasteX = 10
            StartwasteY = 10
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

        If iLastPlateCount >= 1 Then

            If Waste >= WasteMinimumSize Then
                Areareduce = 1
                lb_spildtype.Text = "TIL LAGER"
                SpildTilLager = Wastenetto
                lb_spildnetto.Text = FormatNumber(SpildTilLager, 2)
            Else
                Areareduce = 0
                lb_spildtype.Text = ""
                lb_spildnetto.Text = ""
            End If

        End If

        If Areareducelastplate = 1 Then
            lb_spildtype.Text = "TIL LAGER"
            SpildTilLager = SpildTilLager + Wastenettolastplate
            lb_spildnetto.Text = FormatNumber(SpildTilLager, 2)

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
        Dim wasteXX As Double
        Dim wasteXY As Double
        Dim wasteYX As Double
        Dim wasteYY As Double
        Dim StartwasteX As Double
        Dim StartwasteY As Double
        Dim RowX As Double
        Dim RowY As Double
        Dim Wastenettolastplate As Double
        Dim UsedCountRowXX As Integer
        Dim UsedCountRowXY As Integer
        Dim UsedCountRowYX As Integer
        Dim UsedCountRowYY As Integer
        Dim Areareducelastplate As Integer
        Dim wasteareaXX As Double
        Dim wasteareaXY As Double
        Dim wasteareaYX As Double
        Dim wasteareaYY As Double

        objDatabaseIO = New DatabaseIO

        WasteMinimumSize = objDatabaseIO.GetWasteMinimumSizes(Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text))

        StartwasteX = 100
        StartwasteY = 40
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

        alReturn = New ArrayList(2)


        RowX = SubjectX + subjectaddition
        RowY = SubjectY + subjectaddition

        iXCountX = (PlateY - StartwasteY) \ (SubjectX + subjectaddition)
        If iXCountX > 0 Then
            UsedCountRowXX = (iLastPlateCount / iXCountX) + 0.49
            wasteXX = PlateX - StartwasteX - (RowY * UsedCountRowXX)
            If wasteXX < WasteMinimumSize Then
                wasteXX = 0
            End If
            wasteareaXX = (wasteXX * PlateY) / 1000000
        Else
            wasteareaXX = 0
        End If

        iXCountY = (PlateX - StartwasteX) \ (SubjectY + subjectaddition)
        If iXCountY > 0 Then
            UsedCountRowXY = (iLastPlateCount / iXCountY) + 0.49
            wasteXY = PlateY - StartwasteY - (RowX * UsedCountRowXY)
            If wasteXY < WasteMinimumSize Then
                wasteXY = 0
            End If
            wasteareaXY = (wasteXY * PlateX) / 1000000
        Else
            wasteareaXY = 0
        End If

        iYCountX = (PlateY - StartwasteY) \ (SubjectY + subjectaddition)
        If iYCountX > 0 Then
            UsedCountRowYX = (iLastPlateCount / iYCountX) + 0.49
            wasteYX = PlateX - StartwasteX - (RowX * UsedCountRowYX)
            If wasteYX < WasteMinimumSize Then
                wasteYX = 0
            End If
            wasteareaYX = (wasteYX * PlateY) / 1000000
        Else
            wasteareaYX = 0
        End If

        iYCountY = (PlateX - StartwasteX) \ (SubjectX + subjectaddition)
        If iYCountY > 0 Then
            UsedCountRowYY = (iLastPlateCount / iYCountY) + 0.49
            wasteYY = PlateY - StartwasteY - (RowY * UsedCountRowYY)
            If wasteYY < WasteMinimumSize Then
                wasteYY = 0
            End If
            wasteareaYY = (wasteYY * PlateX) / 1000000
        Else
            wasteareaYY = 0
        End If





        If wasteareaXX > wasteareaXY Then
            Wastenettolastplate = wasteareaXX
            Waste = wasteXX
        Else
            Wastenettolastplate = wasteareaXY
            Waste = wasteXY
        End If

        If wasteareaYX > Wastenettolastplate Then
            Wastenettolastplate = wasteareaYX
            Waste = wasteYX
        End If

        If wasteareaYY > Wastenettolastplate Then
            Wastenettolastplate = wasteareaYY
            Waste = wasteYY
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
        'Dim sheetthickness As Double

        'sheetthickness = (Val(tb_pladetykkelse.Text))

        'If sheetthickness > 10 Then
        'MsgBox("Buk i pladetykkelser over 10mm kan ikke beregnes, spørg bukoperatøren")
        'tb_pladetykkelse.Text = 0
        'ResetCalculations()
        'Exit Function
        'End If

        objbendcalculations = New BendCalculations

        lb_udfold_x.Text = objbendcalculations.CalculateSubjectLength(Val(tb_pladetykkelse.Text), Val(tb_bukmax_x.Text), Val(tb_buk1_x.Text), Val(tb_buk2_x.Text), Val(tb_buk3_x.Text), Val(tb_buk4_x.Text), Val(tb_buk5_x.Text), Val(tb_buk6_x.Text), Val(tb_buk7_x.Text), Val(tb_buk8_x.Text), Val(tb_buk9_x.Text), Val(tb_buk10_x.Text), Val(tb_buk11_x.Text))
        lb_udfold_y.Text = objbendcalculations.CalculateSubjectLength(Val(tb_pladetykkelse.Text), Val(tb_bukmax_y.Text), Val(tb_buk1_y.Text), Val(tb_buk2_y.Text), Val(tb_buk3_y.Text), Val(tb_buk4_y.Text), Val(tb_buk5_y.Text), Val(tb_buk6_y.Text), Val(tb_buk7_y.Text), Val(tb_buk8_y.Text), Val(tb_buk9_y.Text), Val(tb_buk10_y.Text), Val(tb_buk11_y.Text))

        lb_nettoareal.Text = ((Val(lb_udfold_x.Text) * Val(lb_udfold_y.Text)) - ((Val(tb_buk1_x.Text) + Val(tb_buk2_x.Text) + Val(tb_buk3_x.Text) + Val(tb_buk4_x.Text) + Val(tb_buk5_x.Text) + Val(tb_buk6_x.Text) + Val(tb_buk7_x.Text) + Val(tb_buk8_x.Text) + Val(tb_buk9_x.Text) + Val(tb_buk10_x.Text) + Val(tb_buk11_x.Text)) * (Val(tb_buk1_y.Text) + Val(tb_buk2_y.Text) + Val(tb_buk3_y.Text) + Val(tb_buk4_y.Text) + Val(tb_buk5_y.Text) + Val(tb_buk6_y.Text) + Val(tb_buk7_y.Text) + Val(tb_buk8_y.Text) + Val(tb_buk9_y.Text) + Val(tb_buk10_y.Text) + Val(tb_buk11_y.Text)))) * 0.000001


        FModuleSize = CalculateModuleSize(lb_udfold_x.Text, lb_udfold_y.Text, FModuleSize)
        lb_modulstr.Text = FModuleSize


        'getting difficult factor
        objDatabaseIO = New DatabaseIO

        Difficultclass = objDatabaseIO.GetDifficultclass(Lb_klasse.Text, Val(tb_pladetykkelse.Text))
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

    Private Sub CalculateBendTime(ByVal subjectweight As Double, ByVal Antal As Double)
        Dim objDatabaseIO As DatabaseIO

        Dim objbendcalculations As BendCalculations
        Dim sumY As Double
        Dim BendsY As Integer
        Dim bend As ArrayList
        Dim sumX As Double
        Dim sum As Double
        Dim bendmin As Double
        Dim BendsX As Integer
        Dim bendstart As Double
        Dim stepbend As Double

        objbendcalculations = New BendCalculations
        objDatabaseIO = New DatabaseIO


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

        bendmin = (CalculateBendMin(Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), subjectweight)) * (objDatabaseIO.getnumbermultiplier(Antal) / 100 * Antal)

        If Val(tb_stepantal.Text) <> 0 Then
            Calculatestepbend(subjectweight, Val(Antal))
        Else
            lb_stepbuk.Text = ""
            tb_stepbuk_uk.Text = ""
        End If

        sum = FormatNumber((bendmin * (BendsX + BendsY)) + (sumX * bendmin) + (sumY * bendmin), 0)
        lb_buk_tid.Text = sum.ToString()
        bendstart = FormatNumber(CalculateBendSetup(Val(lb_udfold_x.Text), BendsY) + CalculateBendSetup(Val(lb_udfold_y.Text), BendsX), 0)
        lb_buk_opst.Text = bendstart
        stepbend = Val(lb_stepbuk.Text)
        lb_buk_ialt.Text = FormatNumber(sum + bendstart + stepbend, 0)


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
        Dim sheetthickness As Double
        Dim bendfactor As Double

        sheetthickness = (Val(tb_pladetykkelse.Text))

        bendfactor = 1
        If sheetthickness > 10 Then
            bendfactor = 1.1
        End If
        If sheetthickness > 14 Then
            bendfactor = 1.15
        End If
        If sheetthickness > 19 Then
            bendfactor = 1.25
        End If

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
        CalculateBendMin = CalculateBendMin * Difficultfactor * bendfactor

    End Function
    Private Function Calculatestepbend(ByVal subjectweight As Double, ByVal Antal As Double) As Double
        Dim Difficultfactor As Double
        Dim lengthx As Double
        Dim lengthy As Double
        Dim stepantal As Double
        Dim steptid As Double
        Dim modulsize As Double
        Dim stepbendmin As Double


        modulsize = Val(lb_modulstr.Text)
        stepantal = Val(tb_stepantal.Text)
        steptid = 1

        If lb_faktor.Text = "" Then
            Exit Function
        End If

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        lengthx = Val(lb_udfold_x.Text)
        lengthy = Val(lb_udfold_y.Text)

        If modulsize = 1 Then
            steptid = 1
        End If
        If modulsize = 2 Then
            steptid = 1.5
        End If
        If modulsize = 3 Then
            steptid = 2
        End If
        If modulsize = 4 Then
            steptid = 3
        End If
        If modulsize = 5 Then
            steptid = 5
        End If
        If modulsize = 6 Then
            steptid = 7
        End If
        If modulsize = 7 Then
            steptid = 10
        End If

        stepbendmin = FormatNumber(((steptid * stepantal * Antal) / 60), 0)
        lb_stepbuk.Text = stepbendmin

    End Function

    Private Function CalculateSpotWelding(ByVal antal As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim countspots As Integer
        Dim countweldseams As Double
        Dim weldspeed As Double
        Dim weldingtimemin As Double
        Dim spotweldinghandling As Double
        Dim weldspeedmultiplier As Double
        Dim opstart As Double
        Dim totalspotweldingtime As Double


        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO
        countspots = Val(tb_numberofspots.Text)
        countweldseams = Val(tb_numberofspotweldseams.Text)
        spotweldinghandling = 0
        'justering af svejsehastighed
        weldspeedmultiplier = 1

        opstart = 20
        spotweldinghandling = opstart + antal * ((objDatabaseIO.GetSpotWeldingShift(Val(lb_modulstr.Text)) + (objDatabaseIO.GetSpotWeldingStartstop(Val(lb_modulstr.Text)) * countweldseams))) * Difficultfactor

        If cb_spotweld.CheckState = CheckState.Checked Then

            If tb_numberofspotweldseams.Text = "" Then
                MsgBox("Antal svejsesømme og antal punktsvejsninger skal være udfyldt")
                cb_spotweld.CheckState = CheckState.Unchecked
                Exit Function
            End If
            If tb_numberofspots.Text = "" Then
                MsgBox("Antal svejsesømme og antal punktsvejsninger skal være udfyldt")
                cb_spotweld.CheckState = CheckState.Unchecked
                Exit Function
            End If

            If Val(Lb_matrgruppe.Text) = 3 Then
                If Val(tb_pladetykkelse.Text) > 4 Then
                    MsgBox("Pladetykkelse er max 4mm for punktsvejsning i Aluminium")
                    cb_spotweld.CheckState = CheckState.Unchecked
                    Exit Function
                End If
            End If

            If Val(tb_pladetykkelse.Text) > 6 Then
                MsgBox("Pladetykkelse er max 6mm for punktsvejsning i jern og rustfri")
                cb_spotweld.CheckState = CheckState.Unchecked
                Exit Function
            End If

            weldspeed = objDatabaseIO.GetSpotweldspeed(Val(tb_pladetykkelse.Text), Val(Lb_matrgruppe.Text))

            weldingtimemin = antal * weldspeed * countspots * weldspeedmultiplier / 100
            totalspotweldingtime = FormatNumber(weldingtimemin + spotweldinghandling, 0)
            lb_spotweld.Text = totalspotweldingtime
        End If




        If cb_spotweld.CheckState = CheckState.Unchecked Then
            lb_spotweld.Text = ""
            tb_spotweld_uk.Text = ""
        End If



    End Function

    Private Function CalculateWelding(ByVal antal As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim countwelds As Integer
        Dim weldlength As Double
        Dim weldspeed As Double
        Dim weldingtimemin As Double
        Dim weldinghandling As Double
        Dim weldspeedmultiplier As Double
        Dim opstart As Double
        Dim Totalweldingtime As Double
        Dim Totaltackweldingtime As Double
        Dim oprettetid As Double
        Dim svejsetraad As Double
        Dim svejsetraadpris As Double
        Dim tapsvejstid As Double
        Dim tapsvejstid100 As Double
        Dim weldingfactor As Double
        Dim calculated_weldingtime As Double
        Dim check As Integer

        If cb_weld.CheckState = CheckState.Unchecked Then
            If cb_tackweld.CheckState = CheckState.Unchecked Then
                lb_weld.Text = ""
                tb_weld_uk.Text = ""
                lb_tackweld.Text = ""
                tb_tackweld_uk.Text = ""
                lb_tilsatsmatr_kr.Text = ""
                'Exit Function
            End If
        End If
        If cb_weld.CheckState = CheckState.Checked Then
            If Val(Lb_matrgruppe.Text) = 4 Then
                MsgBox("Materialet kan ikke svejses")
                lb_weld.Text = ""
                tb_weld_uk.Text = ""
                lb_tackweld.Text = ""
                tb_tackweld_uk.Text = ""
                lb_tilsatsmatr_kr.Text = ""
                cb_weld.CheckState = CheckState.Unchecked
                cb_tackweld.CheckState = CheckState.Unchecked
                Exit Function
            End If
            If Val(Lb_matrgruppe.Text) = 5 Then
                MsgBox("Materialet kan ikke svejses")
                lb_weld.Text = ""
                tb_weld_uk.Text = ""
                lb_tackweld.Text = ""
                tb_tackweld_uk.Text = ""
                lb_tilsatsmatr_kr.Text = ""
                cb_weld.CheckState = CheckState.Unchecked
                cb_tackweld.CheckState = CheckState.Unchecked
                Exit Function
            End If
        End If

        If cb_tackweld.CheckState = CheckState.Checked Then
            If Val(Lb_matrgruppe.Text) = 4 Then
                MsgBox("Materialet kan ikke svejses")
                lb_weld.Text = ""
                tb_weld_uk.Text = ""
                lb_tackweld.Text = ""
                tb_tackweld_uk.Text = ""
                lb_tilsatsmatr_kr.Text = ""
                cb_weld.CheckState = CheckState.Unchecked
                cb_tackweld.CheckState = CheckState.Unchecked
                Exit Function
            End If
            If Val(Lb_matrgruppe.Text) = 5 Then
                MsgBox("Materialet kan ikke svejses")
                lb_weld.Text = ""
                tb_weld_uk.Text = ""
                lb_tackweld.Text = ""
                tb_tackweld_uk.Text = ""
                lb_tilsatsmatr_kr.Text = ""
                cb_weld.CheckState = CheckState.Unchecked
                cb_tackweld.CheckState = CheckState.Unchecked
                Exit Function
            End If
        End If

        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO
        countwelds = Val(tb_numberofwelds.Text)
        weldlength = Val(tb_weldlength.Text)
        weldinghandling = 0

        'justering af svejsehastighed
        weldspeedmultiplier = 1
        check = 0
        weldingfactor = 1
        If rb_factor1.Checked = True Then
            weldingfactor = 1
            check = 1
        End If
        If rb_factor2.Checked = True Then
            weldingfactor = 1.2
            check = 1
        End If
        If rb_factor3.Checked = True Then
            weldingfactor = 1.5
            check = 1
        End If
        If rb_factor4.Checked = True Then
            weldingfactor = 2
            check = 1
        End If
        If rb_factor5.Checked = True Then
            weldingfactor = 3
            check = 1
        End If
        If check = 0 Then
            rb_factor1.Checked = True
            weldingfactor = 1
        End If
        opstart = 15
        weldinghandling = opstart + antal * ((objDatabaseIO.GetWeldingShift(Val(lb_modulstr.Text)) + (objDatabaseIO.GetWeldingStartstop(Val(lb_modulstr.Text)) * countwelds))) * Difficultfactor

        If cb_weld.CheckState = CheckState.Checked Then
            cb_tackweld.CheckState = CheckState.Unchecked

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

            weldingtimemin = antal * weldspeed * weldlength * weldspeedmultiplier / 100
            Totalweldingtime = FormatNumber((weldingtimemin + weldinghandling) * weldingfactor, 0)
            lb_weld.Text = Totalweldingtime


            svejsetraad = Val(tb_pladetykkelse.Text) * Val(tb_pladetykkelse.Text) * weldlength * 0.5 * 0.001
            If Val(Lb_matrgruppe.Text) = 1 Then
                svejsetraadpris = Format(svejsetraad * 8 * 45 / 1000, 0)
            End If
            If Val(Lb_matrgruppe.Text) = 2 Then
                svejsetraadpris = Format(svejsetraad * 8 * 75 / 1000, 0)
            End If
            If Val(Lb_matrgruppe.Text) = 3 Then
                svejsetraadpris = Format(svejsetraad * 2.8 * 220 / 1000, 0)
            End If
            lb_tilsatsmatr_kr.Text = svejsetraad
        End If

        oprettetid = FormatNumber(Totalweldingtime * 0.25, 0)

        If cb_rettesvejs.CheckState = CheckState.Checked Then
            lb_rettesvejs_tid.Text = oprettetid
        Else
            lb_rettesvejs_tid.Text = ""
            tb_rettesvejs_uk.Text = ""
            oprettetid = 0

        End If

        If cb_tackweld.CheckState = CheckState.Checked Then
            cb_weld.CheckState = CheckState.Unchecked

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

            weldingtimemin = antal * weldspeed * weldlength * weldspeedmultiplier / 100
            Totaltackweldingtime = FormatNumber((weldingtimemin + weldinghandling) * weldingfactor, 0)
            lb_tackweld.Text = Totaltackweldingtime

            svejsetraad = Val(tb_pladetykkelse.Text) * Val(tb_pladetykkelse.Text) * weldlength * antal * 0.5 * 0.001 * 0.33

            If Val(Lb_matrgruppe.Text) = 1 Then
                svejsetraadpris = Format(svejsetraad * 8 * 45 / 1000, 0)
            End If
            If Val(Lb_matrgruppe.Text) = 2 Then
                svejsetraadpris = Format(svejsetraad * 8 * 75 / 1000, 0)
            End If
            If Val(Lb_matrgruppe.Text) = 3 Then
                svejsetraadpris = Format(svejsetraad * 2.8 * 220 / 1000, 0)
            End If
            lb_tilsatsmatr_kr.Text = svejsetraadpris
        End If

        If cb_weld.CheckState = CheckState.Unchecked Then
            lb_weld.Text = ""
            tb_weld_uk.Text = ""
        End If
        If cb_tackweld.CheckState = CheckState.Unchecked Then
            lb_tackweld.Text = ""
            tb_tackweld_uk.Text = ""
        End If
        'tapsvejsning
        If tb_tapantal.Text <> "" Then
            tapsvejstid100 = 30
            If Totalweldingtime + Totaltackweldingtime + oprettetid + Val(lb_grind_weld.Text) > 15 Then
                opstart = 0
            End If
            tapsvejstid = Format((Val(tb_tapantal.Text) * tapsvejstid100 / 100 * antal * weldingfactor) + opstart, 0)
        End If

        lb_tapsvejs_tid.Text = tapsvejstid

        If tb_tapantal.Text = "" Then
            lb_tapsvejs_tid.Text = ""
            tb_tapsvejs_uk.Text = ""
        End If

        calculated_weldingtime = Totalweldingtime + Totaltackweldingtime + oprettetid + Val(lb_grind_weld.Text) + tapsvejstid
        tb_svejstid_ialt.Text = Format(calculated_weldingtime, 0)

        If calculated_weldingtime = 0 Then
            tb_svejstid_ialt.Text = ""
        End If


    End Function

    Private Function CalculateGrindWelding(ByVal antal As Integer) As Double
        Dim objDatabaseIO As DatabaseIO
        Dim Difficultfactor As Double
        Dim countwelds As Integer
        Dim weldlength As Double
        Dim Grindweldspeed As Double
        Dim Grindweldingtimemin As Double
        Dim Grindweldinghandling As Double
        Dim Grindweldspeedmultiplier As Double
        Dim opstart As Double
        Dim totalgrindweldingtime As Double


        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        objDatabaseIO = New DatabaseIO
        countwelds = Val(tb_numberofwelds.Text)
        weldlength = Val(tb_weldlength.Text)
        Grindweldinghandling = 0
        'justering af svejsehastighed
        Grindweldspeedmultiplier = 0.5

        opstart = 5
        Grindweldinghandling = opstart + antal * ((objDatabaseIO.GetWeldingShift(Val(lb_modulstr.Text)) + (objDatabaseIO.GetWeldingStartstop(Val(lb_modulstr.Text)) * countwelds))) * Difficultfactor

        If cb_grind_weld.CheckState = CheckState.Checked Then

            If cb_tackweld.CheckState = CheckState.Checked Then

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
                    Grindweldspeed = objDatabaseIO.GetgrindtackweldspeedTIG100(Val(tb_pladetykkelse.Text))
                Else
                    Grindweldspeed = objDatabaseIO.GetgrindtackweldspeedMIG100(Val(tb_pladetykkelse.Text))
                End If

                Grindweldingtimemin = antal * Grindweldspeed * weldlength * Grindweldspeedmultiplier / 100

                lb_grind_weld.Text = FormatNumber(Grindweldingtimemin + Grindweldinghandling, 0)
            End If

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
                    Grindweldspeed = objDatabaseIO.GetgrindweldspeedTIG100(Val(tb_pladetykkelse.Text))
                Else
                    Grindweldspeed = objDatabaseIO.GetgrindweldspeedMIG100(Val(tb_pladetykkelse.Text))
                End If

                Grindweldingtimemin = antal * Grindweldspeed * weldlength * Grindweldspeedmultiplier / 100
                totalgrindweldingtime = FormatNumber(Grindweldingtimemin + Grindweldinghandling, 0)
                lb_grind_weld.Text = totalgrindweldingtime
            End If

        End If

        If cb_weld.CheckState = CheckState.Unchecked Then
            If cb_tackweld.CheckState = CheckState.Unchecked Then
                cb_grind_weld.CheckState = CheckState.Unchecked
            End If
        End If

        If cb_grind_weld.CheckState = CheckState.Unchecked Then
            lb_grind_weld.Text = ""
            tb_grind_weld_uk.Text = ""
        End If

    End Function

    Private Function Calculateglasbl(ByVal antal As Integer) As Double
        Dim opstart As Double
        Dim areal As Double
        Dim m2tid As Double

        If cb_glasbl.CheckState = CheckState.Checked Then

            opstart = 15
            m2tid = 20
            areal = lb_nettoareal.Text
            If areal < 0.01 Then
                areal = 0.01
            End If
            Calculateglasbl = opstart + (areal * antal * m2tid)
            lb_glasbl.Text = FormatNumber(Calculateglasbl, 0)

        Else
            Calculateglasbl = 0
            lb_glasbl.Text = ""
            tb_glasbl_uk.Text = ""
        End If

    End Function
    Private Function Calculateslib(ByVal antal As Integer) As Double
        Dim opstart As Double
        Dim areal As Double
        Dim m2tid As Double

        If cb_slib.CheckState = CheckState.Checked Then

            opstart = 15
            m2tid = 30
            areal = lb_nettoareal.Text
            If areal < 0.01 Then
                areal = 0.01
            End If
            Calculateslib = opstart + (areal * antal * m2tid)
            lb_slib.Text = FormatNumber(Calculateslib, 0)

        Else
            Calculateslib = 0
            lb_slib.Text = ""
            tb_slib_uk.Text = ""

        End If

    End Function
    Private Sub ClearAllCalculations()
        If rb_klip.Checked = False Then
            pobjFilter_tb_klip_tid_uk.ClearValues()
            pobjFilter_tb_klip_opstart_uk.ClearValues()
            lb_klip_ialt.Text = ""
            lb_klip_ialtuk.Text = ""

        End If
        If rb_D_stans.Checked = False Then
            pobjFilter_tb_CNCmin_uk.ClearValues()
            pobjFilter_tb_gruppe1_opstart_uk.ClearValues()
            lb_stans_ialt.Text = ""
            lb_stans_ialtuk.Text = ""

        End If
        If rb_C_laser.Checked = False Then
            pobjFilter_tb_laserCNC_tid_uk.ClearValues()
            pobjFilter_tb_laser_opstart_uk.ClearValues()
            lb_laser_ialt.Text = ""
            lb_laser_ialtuk.Text = ""

        End If
        If rb_B_kombi.Checked = False Then
            pobjFilter_tb_combi_opstart_uk.ClearValues()
            pobjFilter_tb_combiCNC_tid_uk.ClearValues()
            pobjFilter_tb_combiCNCstans_tid_uk.ClearValues()
            lb_combi_ialt.Text = ""
            lb_combi_ialtuk.Text = ""

        End If
    End Sub

    Private Sub CalculateManHourSubtotal(ByVal Antal As Double)
        Dim objDatabaseIO As DatabaseIO

        Dim dblSubtotal As Double
        Dim dblSubtotalstart As Double
        Dim dblSubtotalTime As Double
        Dim uk_correction As Double
        Dim uk_correction1 As Double
        Dim dblSubcalculatedTime As Double
        Dim dblSubtotalstartTime As Double
        objDatabaseIO = New DatabaseIO
        Dim SumcalculatedTime As Single
        Dim gevind As Integer
        Dim addprice_glassb As Double
        Dim addprice_steelm As Double
        Dim addprice_vibration As Double
        Dim addprice_grinding As Double
        Dim addprice As Double


        If lb_gevind.Text = "" Then
            gevind = 0
        Else
            gevind = lb_gevind.Text
        End If

        If rb_danmark.Checked = True Then
            addprice_glassb = 150
            addprice_steelm = 150
            addprice_vibration = 150
            addprice_grinding = 250
        Else
            addprice_glassb = 100
            addprice_steelm = 100
            addprice_vibration = 100
            addprice_grinding = 100
        End If

        'Summing
        SumcalculatedTime = Convert.ToDouble(Val(lb_buk_tid.Text)) + Convert.ToDouble(Val(lb_stepbuk.Text)) + Convert.ToDouble(Val(lb_valsetid.Text)) + Convert.ToDouble(Val(lb_klip_tid.Text)) + Convert.ToDouble(Val(lb_afgrat.Text)) + Convert.ToDouble(Val(lb_grinding.Text)) + Convert.ToDouble(Val(lb_brush.Text)) + Convert.ToDouble(Val(lb_vibration.Text)) + Convert.ToDouble(Val(lb_rette.Text)) + Convert.ToDouble(Val(lb_stans_manuel.Text)) + Convert.ToDouble(Val(lb_countersink.Text)) + gevind + Convert.ToDouble(Val(lb_pressnut.Text)) + Convert.ToDouble(Val(lb_presstag.Text)) + Convert.ToDouble(Val(lb_boltesvejs.Text)) + Convert.ToDouble(Val(lb_spotweld.Text)) + Convert.ToDouble(Val(lb_tackweld.Text)) + Convert.ToDouble(Val(lb_weld.Text)) + Convert.ToDouble(Val(lb_rettesvejs_tid.Text)) + Convert.ToDouble(Val(lb_grind_weld.Text)) + Convert.ToDouble(Val(lb_tapsvejs_tid.Text)) + Convert.ToDouble(Val(lb_glasbl.Text)) + Convert.ToDouble(Val(lb_slib.Text))

        dblSubcalculatedTime = ((Convert.ToDouble(Val(lb_buk_tid.Text)) + Convert.ToDouble(Val(lb_stepbuk.Text)) + Convert.ToDouble(Val(lb_valsetid.Text)) + Convert.ToDouble(Val(lb_klip_tid.Text)) + Convert.ToDouble(Val(lb_afgrat.Text)) + Convert.ToDouble(Val(lb_grinding.Text)) + Convert.ToDouble(Val(lb_brush.Text)) + Convert.ToDouble(Val(lb_vibration.Text)) + Convert.ToDouble(Val(lb_rette.Text)) + Convert.ToDouble(Val(lb_stans_manuel.Text)) + Convert.ToDouble(Val(lb_countersink.Text)) + gevind + Convert.ToDouble(Val(lb_pressnut.Text)) + Convert.ToDouble(Val(lb_presstag.Text)) + Convert.ToDouble(Val(lb_boltesvejs.Text)) + Convert.ToDouble(Val(lb_spotweld.Text)) + Convert.ToDouble(Val(lb_tackweld.Text)) + Convert.ToDouble(Val(lb_weld.Text)) + Convert.ToDouble(Val(lb_rettesvejs_tid.Text)) + Convert.ToDouble(Val(lb_grind_weld.Text)) + Convert.ToDouble(Val(lb_tapsvejs_tid.Text)) + Convert.ToDouble(Val(lb_glasbl.Text))) / 60)
        dblSubtotalTime = ((Val(pobjFilter_tb_buk_uk.GetValue) + Val(pobjFilter_tb_stepbuk_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_valsning_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_klip_tid_uk.GetValue) + Val(pobjFilter_tb_afgrat_uk.GetValue) + Val(pobjFilter_tb_grinding_uk.GetValue) + Val(pobjFilter_tb_brush_uk.GetValue) + Val(pobjFilter_tb_vibration_uk.GetValue) + Val(pobjFilter_tb_rette_uk.GetValue) + Val(pobjFilter_tb_stans_manuel_uk.GetValue) + Val(pobjFilter_tb_countersink_uk.GetValue) + Val(pobjFilter_tb_gevind_uk.GetValue) + Val(pobjFilter_tb_pressnut_uk.GetValue) + Val(pobjFilter_tb_presstag_uk.GetValue) + Val(pobjFilter_tb_boltesvejs_uk.GetValue) + Val(pobjFilter_tb_spotweld_uk.GetValue) + Val(pobjFilter_tb_tackweld_uk.GetValue) + Val(pobjFilter_tb_weld_uk.GetValue) + Val(pobjFilter_tb_rettesvejs_uk.GetValue) + Val(pobjFilter_tb_grind_weld_uk.GetValue) + Val(pobjFilter_tb_tapsvejs_uk.GetValue) + Val(pobjFilter_tb_glasbl_uk.GetValue) + Val(pobjFilter_tb_slib_uk.GetValue)) / 60)

        If Antal = Convert.ToDouble(tb_antal1.Text) Then
            uk_correction1 = (dblSubtotalTime - dblSubcalculatedTime)
            lb_ukTime.Text = FormatNumber(uk_correction1, 1)
        End If

        uk_correction = Convert.ToDouble(lb_ukTime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal

        addprice = (Val(pobjFilter_tb_grinding_uk.GetValue) / 60 * addprice_steelm) + ((Val(pobjFilter_tb_grind_weld_uk.GetValue) + Val(pobjFilter_tb_slib_uk.GetValue)) / 60 * addprice_grinding) + (Val(pobjFilter_tb_vibration_uk.GetValue) / 60 * addprice_vibration) + (Val(pobjFilter_tb_glasbl_uk.GetValue) / 60 * addprice_glassb)

        dblSubtotal = ((dblSubcalculatedTime + uk_correction) * Val(tb_timesats_mand.Text)) + addprice
        dblSubtotalstartTime = dblSubtotalstartTime + (Val(pobjFilter_tb_buk_opst_uk.GetValue) + Val(pobjFilter_tb_klip_opstart_uk.GetValue) + Val(pobjFilter_tb_kontor_uk.GetValue) + Val(pobjFilter_tb_kontrol_uk.GetValue)) / 60
        dblSubtotalstart = dblSubtotalstartTime * Val(tb_timesats_mand.Text)

        dblSubtotalTime = dblSubcalculatedTime + uk_correction


        dblSubtotalTime = FormatNumber(dblSubtotalTime, 2)
        dblSubtotalstartTime = FormatNumber(dblSubtotalstartTime, 2)
        'kr
        lb_mand1.Text = FormatNumber(Convert.ToString(dblSubtotal + dblSubtotalstart), 2)
        'timer
        lb_mand1_tid.Text = Convert.ToString(dblSubtotalTime + dblSubtotalstartTime)



    End Sub

    Private Sub CalculateCNCHourSubtotal(ByVal Antal As Double)
        Dim objDatabaseIO As DatabaseIO
        Dim dblSubtotal As Double
        Dim dblSubCalculatedTime As Double
        Dim dblSubCalculatedTimestans As Double
        Dim dblSubStart As Double
        Dim dblSubtotalTime As Double
        Dim dblSubtotalTimestans As Double
        Dim dblSubStartTime As Double
        Dim uk_correction As Double
        Dim uk_correction1 As Double
        Dim uk_correction2 As Double
        Dim uk_correction_stans As Double



        objDatabaseIO = New DatabaseIO
        'Summing
        If pobjFilter_tb_CNCmin_uk.GetValue <> 0 Then
            dblSubCalculatedTime = dblSubCalculatedTime + (Convert.ToDouble(lb_gruppe1_tid.Text) / 60)
            dblSubtotalTime = dblSubtotalTime + (Convert.ToDouble(pobjFilter_tb_CNCmin_uk.GetValue) / 60)
            If Antal = Val(tb_antal1.Text) Then
                uk_correction1 = (dblSubtotalTime - dblSubCalculatedTime)
                lb_ukCNCtime.Text = uk_correction1
            End If
            uk_correction = Convert.ToDouble(lb_ukCNCtime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal
            dblSubtotal = dblSubtotal + (dblSubCalculatedTime + uk_correction) * Val(tb_timesats_D.Text)
            dblSubStart = dblSubStart + (Convert.ToDouble(pobjFilter_tb_gruppe1_opstart_uk.GetValue) * Val(tb_timesats_D.Text) / 60)
            dblSubStartTime = dblSubStartTime + (Convert.ToDouble(pobjFilter_tb_gruppe1_opstart_uk.GetValue) / 60)

            'LASER
        End If
        If pobjFilter_tb_laserCNC_tid_uk.GetValue <> 0 Then
            dblSubCalculatedTime = dblSubCalculatedTime + (Convert.ToDouble(lb_laserCNC_tid.Text) / 60)
            dblSubtotalTime = dblSubtotalTime + (Convert.ToDouble(pobjFilter_tb_laserCNC_tid_uk.GetValue) / 60)
            If Antal = Val(tb_antal1.Text) Then
                uk_correction1 = (dblSubtotalTime - dblSubCalculatedTime)
                lb_ukCNCtime.Text = uk_correction1
            End If
            uk_correction = Convert.ToDouble(lb_ukCNCtime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal
            dblSubtotal = (dblSubCalculatedTime + uk_correction) * Val(tb_timesats_C.Text)
            dblSubStart = (Convert.ToDouble(pobjFilter_tb_laser_opstart_uk.GetValue) * Val(tb_timesats_C.Text) / 60)
            dblSubStartTime = (Convert.ToDouble(pobjFilter_tb_laser_opstart_uk.GetValue) / 60)
        End If

        'COMBI

        If pobjFilter_tb_combiCNC_tid_uk.GetValue <> 0 Then
            dblSubCalculatedTime = dblSubCalculatedTime + (Convert.ToDouble(lb_CombiCNC_tid.Text) / 60)
            dblSubtotalTime = dblSubtotalTime + (Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) / 60)
            If Antal = Val(tb_antal1.Text) Then
                uk_correction1 = (dblSubtotalTime - dblSubCalculatedTime)
                lb_ukCNCtime.Text = uk_correction1
            End If

            If pobjFilter_tb_combiCNCstans_tid_uk.GetValue <> 0 Then
                dblSubCalculatedTimestans = dblSubCalculatedTimestans + (Convert.ToDouble(lb_CombiCNCstans_tid.Text) / 60)
                dblSubtotalTimestans = dblSubtotalTimestans + (Convert.ToDouble(pobjFilter_tb_combiCNCstans_tid_uk.GetValue) / 60)
                If Antal = Val(tb_antal1.Text) Then
                    uk_correction2 = (dblSubtotalTimestans - dblSubCalculatedTimestans)
                    lb_ukCNCstanstime.Text = uk_correction2
                End If



                'If Antal = Val(tb_antal1.Text) Then
                'dblSubtotal = dblSubtotal + ((Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) * Val(tb_timesats_B.Text)) + (Convert.ToDouble(pobjFilter_tb_combiCNCstans_tid_uk.GetValue) * Val(tb_timesats_D.Text))) / 60

                uk_correction = Convert.ToDouble(lb_ukCNCtime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal
                'uk_correction_stans = Convert.ToDouble(lb_ukCNCstanstime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal
                'Else
                'uk_correction = Convert.ToDouble(lb_ukCNCtime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal
                uk_correction_stans = Convert.ToDouble(lb_ukCNCstanstime.Text) / Convert.ToDouble(tb_antal1.Text) * Antal

                'dblSubtotal = dblSubtotal + ((Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) / Val(tb_antal1.Text) * Antal) * Val(tb_timesats_B.Text) / 60) + ((Convert.ToDouble(pobjFilter_tb_combiCNCstans_tid_uk.GetValue) / Val(tb_antal1.Text) * Antal) * Val(tb_timesats_D.Text) / 60)
                'End If

                dblSubtotal = ((dblSubCalculatedTime + uk_correction) * Val(tb_timesats_C.Text)) + ((dblSubCalculatedTimestans + uk_correction_stans) * Val(tb_timesats_D.Text))


                dblSubStart = dblSubStart + (Convert.ToDouble(pobjFilter_tb_combi_opstart_uk.GetValue) * ((Val(tb_timesats_B.Text) + Val(tb_timesats_D.Text)) / 2) / 60)
                dblSubStartTime = dblSubStartTime + (Convert.ToDouble(pobjFilter_tb_combi_opstart_uk.GetValue) / 60)
            End If
        End If
        dblSubtotalTime = dblSubCalculatedTime + uk_correction + dblSubCalculatedTimestans + uk_correction_stans
        dblSubtotalTime = FormatNumber(dblSubtotalTime, 2)
        dblSubStartTime = FormatNumber(dblSubStartTime, 2)

        lb_cnc1.Text = Convert.ToString(dblSubtotal + dblSubStart)
        lb_CNC1_tid.Text = Convert.ToString(dblSubtotalTime + dblSubStartTime)
        lb_cnc1.Text = FormatNumber(lb_cnc1.Text, 0)

        'sammentælling til ialt uk tider CNC
        lb_buk_ialtuk.Text = (Convert.ToDouble(pobjFilter_tb_buk_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_buk_opst_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_stepbuk_uk.GetValue))
        If rb_klip.Checked = True Then
            lb_klip_ialtuk.Text = (Convert.ToDouble(pobjFilter_tb_klip_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_klip_opstart_uk.GetValue))
        End If
        If rb_D_stans.Checked = True Then
            lb_stans_ialtuk.Text = (Convert.ToDouble(pobjFilter_tb_CNCmin_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_gruppe1_opstart_uk.GetValue))
        End If
        If rb_C_laser.Checked = True Then
            lb_laser_ialtuk.Text = (Convert.ToDouble(pobjFilter_tb_laserCNC_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_laser_opstart_uk.GetValue))
        End If
        If rb_B_kombi.Checked = True Then
            lb_combi_ialtuk.Text = (Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_combiCNCstans_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_combi_opstart_uk.GetValue))
        End If

    End Sub

    Private Sub CalculatePurchasingSubtotal(ByVal Antal As Double)

        Dim pressnut As Double
        Dim boltesvejs As Double
        Dim presstag As Double
        'Dim svejsestag As Double
        Dim tilsatsmatr As Double
        Dim indkøb1 As Double
        Dim nitrogenpris As Double

        boltesvejs = Val(pobjFilter_tb_svejsestag_kr_uk.GetValue)
        pressnut = Val(pobjFilter_tb_pressnut_kr_uk.GetValue)
        presstag = Val(pobjFilter_tb_presstag_kr_uk.GetValue)
        tilsatsmatr = Val(pobjFilter_tb_tilsatsmatr_kr_uk.GetValue)
        nitrogenpris = Val(lb_nitrogen.Text) * 1.15

        'lb_indkøb1.Text = FormatNumber(CalculateSurfaceTreatment(Antal) + ((Val(pobjFilter_tb_pressnut_kr_uk.GetValue) + (Val(pobjFilter_tb_presstag_kr_uk.GetValue) + (Val(pobjFilter_tb_svejsestag_kr_uk.GetValue) + Val(pobjFilter_tb_tilsatsmatr_kr_uk.GetValue))) * Antal)), 2)
        indkøb1 = (CalculateSurfaceTreatment(Antal) + (pressnut + presstag + boltesvejs + tilsatsmatr) * Antal) + nitrogenpris
        lb_indkøb1.Text = FormatNumber(indkøb1, 2)
    End Sub

    Private Function CalculateRawGoods(ByVal Numberofsheets As Integer, ByVal SheetformatX As Double, ByVal SheetformatY As Double, ByVal Sheetthickness As Double, ByVal Materialprice As Double, ByVal materialgroup As Integer, ByVal WasteToStore As Double) As Double
        Dim m2subject As Double
        Dim m2prsheet As Double
        Dim m2subtotal As Double
        Dim objDatabaseIO As DatabaseIO
        Dim decimalsheet As Double

        objDatabaseIO = New DatabaseIO
        CalculateRawGoods = ((Numberofsheets * ((SheetformatX * SheetformatY) / 1000000)) - WasteToStore) * Sheetthickness * Materialprice * objDatabaseIO.GetMaterialWeight(materialgroup)
        'm2 pr emne
        m2subtotal = ((Numberofsheets * ((SheetformatX * SheetformatY) / 1000000)) - WasteToStore)
        m2subject = m2subtotal / Val(tb_antal1.Text)
        m2prsheet = ((SheetformatX * SheetformatY) / 1000000)
        decimalsheet = m2subtotal / m2prsheet
        lb_pladeforbrug.Text = decimalsheet


    End Function

    Private Sub CalculateRawGoodsSubtotal(ByVal Antal As Double)
        Dim råvarer As Double
        Dim råvarerstk As Double


        If lb_spildnetto.Text = "" Then
            lb_spildnetto.Text = 0
        End If
        råvarer = CalculateRawGoods(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(pobjFilter_tb_pladetykkelse.GetValue), FormatNumber((pobjFilter_tb_Kilopris_uk.GetValue), 2), Val(Lb_matrgruppe.Text), FormatNumber(lb_spildnetto.Text, 2))
        råvarerstk = råvarer / Antal
        lb_råvarerstk1.Text = FormatNumber(råvarerstk, 2)
        lb_råvarer1.Text = FormatNumber(råvarer, 0)


        CalculateBestPlate(tb_antal1.Text)
    End Sub

    Private Sub CalculateSubtotals(ByVal Antal As Double)
        Dim objDatabaseIO As DatabaseIO
        Dim LaserpricesDK As Double
        Dim LaserpricesPL As Double

        If rb_danmark.Checked = True Then
            objDatabaseIO = New DatabaseIO
            LaserpricesDK = objDatabaseIO.GetLaserpriceDK(Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text))
            tb_timesats_C.Text = LaserpricesDK
            tb_timesats_B.Text = LaserpricesDK
        End If
        If rb_polen.Checked = True Then
            objDatabaseIO = New DatabaseIO
            LaserpricesPL = objDatabaseIO.GetLaserpricePL(Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text))
            tb_timesats_C.Text = LaserpricesPL
            tb_timesats_B.Text = LaserpricesPL
        End If

        CalculateManHourSubtotal(Antal)
        CalculateCNCHourSubtotal(Antal)
        CalculatePurchasingSubtotal(Antal)
        CalculateRawGoodsSubtotal(Antal)
    End Sub
    Private Sub CalculateTotalPrice(ByVal Antal As Double)
        Dim samlet1 As Double

        If lb_råvarerstk1.Text = "" Then
            Exit Sub
        End If

        lb_opstart_program.Text = pobjFilter_tb_opstart_kr_uk.GetValue + pobjFilter_tb_program_kr_uk.GetValue + pobjFilter_tb_antal_opstart_uk.GetValue
        lb_opst_prog_brutto.Text = Val(lb_opstart_program.Text) + (Val(lb_opstart_program.Text) * pobjFilter_tb_opstart_avance.GetValue / 100)


        lb_timer1.Text = FormatNumber(Convert.ToDouble(lb_mand1.Text) + Convert.ToDouble(lb_cnc1.Text), 2)
        lb_samlet1.Text = FormatNumber(((Convert.ToDouble(lb_timer1.Text) + Convert.ToDouble(lb_indkøb1.Text) + Convert.ToDouble(lb_råvarer1.Text)) / Antal), 2) + " / " + FormatNumber(Convert.ToDouble(lb_timer1.Text) + Convert.ToDouble(lb_indkøb1.Text) + Convert.ToDouble(lb_råvarer1.Text), 0)
        samlet1 = Convert.ToDouble(lb_timer1.Text) + Convert.ToDouble(lb_indkøb1.Text) + Convert.ToDouble(lb_råvarer1.Text)
        lb_salg1.Text = FormatNumber((samlet1 + samlet1 * pobjFilter_tb_avance.GetValue / 100) / Antal, 2) + " / " + FormatNumber((samlet1 + samlet1 * pobjFilter_tb_avance.GetValue / 100), 0)



    End Sub


    Public Sub CalculateOrdrestr()
        If pbLoading = True Then
            Exit Sub
        End If
        If (pobjFilter_tb_bukmax_x.GetValue = 0 Or pobjFilter_tb_bukmax_Y.GetValue = 0 Or pobjFilter_tb_pladetykkelse.GetValue = 0) Then
            Exit Sub
        End If

        If tb_antal1.Text = "" Then
            tb_antal1.Text = 100
        End If
        CalculateAll(tb_antal1.Text)


        If Val(tb_antal2.Text) >= 1 Then
            CalculateAll(tb_antal2.Text)
            CalculateOrdrestr2()
        End If

        If Val(tb_antal3.Text) >= 1 Then
            CalculateAll(tb_antal3.Text)
            CalculateOrdrestr3()
        End If

        If Val(tb_antal4.Text) >= 1 Then
            CalculateAll(tb_antal4.Text)
            CalculateOrdrestr4()
        End If

        If Val(tb_antal5.Text) >= 1 Then
            CalculateAll(tb_antal5.Text)
            CalculateOrdrestr5()
        End If



        CalculateAll(tb_antal1.Text)

        If (Val(tb_pladetykkelse.Text)) > 10 Then
            If Val(tb_buk1_x.Text) + Val(tb_buk1_y.Text) <> 0 Then
                MsgBox("Ved buk i pladetykkelser over 10mm er beregningen usikker, spørg bukoperatøren om buk tid ")
            End If
        End If

    End Sub

    Private Sub CalculateOrdrestr2()


        lb_mand2.Text = lb_mand1.Text
        lb_mand2_tid.Text = lb_mand1_tid.Text
        lb_cnc2.Text = lb_cnc1.Text
        lb_CNC2_tid.Text = lb_CNC1_tid.Text
        lb_timer2.Text = lb_timer1.Text
        lb_indkøb2.Text = lb_indkøb1.Text
        lb_råvarerstk2.Text = lb_råvarerstk1.Text
        lb_råvarer2.Text = lb_råvarer1.Text
        lb_samlet2.Text = lb_samlet1.Text
        lb_salg2.Text = lb_salg1.Text


    End Sub
    Private Sub CalculateOrdrestr3()


        lb_mand3.Text = lb_mand1.Text
        lb_mand3_tid.Text = lb_mand1_tid.Text
        lb_cnc3.Text = lb_cnc1.Text
        lb_CNC3_tid.Text = lb_CNC1_tid.Text
        lb_timer3.Text = lb_timer1.Text
        lb_indkøb3.Text = lb_indkøb1.Text
        lb_råvarerstk3.Text = lb_råvarerstk1.Text
        lb_råvarer3.Text = lb_råvarer1.Text
        lb_samlet3.Text = lb_samlet1.Text
        lb_salg3.Text = lb_salg1.Text


    End Sub
    Private Sub CalculateOrdrestr4()


        lb_mand4.Text = lb_mand1.Text
        lb_mand4_tid.Text = lb_mand1_tid.Text
        lb_cnc4.Text = lb_cnc1.Text
        lb_CNC4_tid.Text = lb_CNC1_tid.Text
        lb_timer4.Text = lb_timer1.Text
        lb_indkøb4.Text = lb_indkøb1.Text
        lb_råvarerstk4.Text = lb_råvarerstk1.Text
        lb_råvarer4.Text = lb_råvarer1.Text
        lb_samlet4.Text = lb_samlet1.Text
        lb_salg4.Text = lb_salg1.Text


    End Sub
    Private Sub CalculateOrdrestr5()


        lb_mand5.Text = lb_mand1.Text
        lb_mand5_tid.Text = lb_mand1_tid.Text
        lb_cnc5.Text = lb_cnc1.Text
        lb_CNC5_tid.Text = lb_CNC1_tid.Text
        lb_timer5.Text = lb_timer1.Text
        lb_indkøb5.Text = lb_indkøb1.Text
        lb_råvarerstk5.Text = lb_råvarerstk1.Text
        lb_råvarer5.Text = lb_råvarer1.Text
        lb_samlet5.Text = lb_samlet1.Text
        lb_salg5.Text = lb_salg1.Text


    End Sub

    Public Sub CalculateGrouphours()
        lb_gruppetid1.Text = FormatNumber((Convert.ToDouble(pobjFilter_tb_CNCmin_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_gruppe1_opstart_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_laserCNC_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_laser_opstart_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_combiCNC_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_combiCNCstans_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_combi_opstart_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_klip_tid_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_klip_opstart_uk.GetValue)) / 60, 2)
        If lb_gruppetid1.Text = 0.0 Then
            lb_gruppetid1.Text = ""
        End If
        lb_gruppetid2.Text = FormatNumber((Val(pobjFilter_tb_afgrat_uk.GetValue) + Val(pobjFilter_tb_grinding_uk.GetValue) + Val(pobjFilter_tb_brush_uk.GetValue) + Val(pobjFilter_tb_vibration_uk.GetValue) + Val(pobjFilter_tb_rette_uk.GetValue) + Val(pobjFilter_tb_stans_manuel_uk.GetValue) + Val(pobjFilter_tb_countersink_uk.GetValue) + Val(pobjFilter_tb_gevind_uk.GetValue) + Val(pobjFilter_tb_presstag_uk.GetValue) + Val(pobjFilter_tb_pressnut_uk.GetValue) + Val(pobjFilter_tb_boltesvejs_uk.GetValue)) / 60, 2)
        If lb_gruppetid2.Text = 0.0 Then
            lb_gruppetid2.Text = ""
        End If
        lb_gruppetid3.Text = FormatNumber((Convert.ToDouble(pobjFilter_tb_buk_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_stepbuk_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_buk_opst_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_valsning_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_spotweld_uk.GetValue)) / 60, 2)
        If lb_gruppetid3.Text = 0.0 Then
            lb_gruppetid3.Text = ""
        End If
        lb_gruppetid4.Text = FormatNumber((Convert.ToDouble(pobjFilter_tb_tackweld_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_weld_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_rettesvejs_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_grind_weld_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_tapsvejs_uk.GetValue)) / 60, 2)
        If lb_gruppetid4.Text = 0.0 Then
            lb_gruppetid4.Text = ""
        End If
        lb_gruppetid5.Text = FormatNumber((Convert.ToDouble(pobjFilter_tb_kontrol_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_kontor_uk.GetValue)) / 60, 2)
        If lb_gruppetid5.Text = 0.0 Then
            lb_gruppetid5.Text = ""
        End If
        lb_gruppetid5slib.Text = FormatNumber((Convert.ToDouble(pobjFilter_tb_glasbl_uk.GetValue) + Convert.ToDouble(pobjFilter_tb_slib_uk.GetValue)) / 60, 2)
        If lb_gruppetid5slib.Text = 0.0 Then
            lb_gruppetid5slib.Text = ""
        End If
    End Sub



    Public Sub CalculateAll(ByVal Antal As Double)

        Dim subjectweight As Double
        Dim stopvalue As Double
        Dim sheetthickness As Double
        Dim materiale As String

        sheetthickness = (Val(tb_pladetykkelse.Text))

        If tb_vibration_uk.Text <> "" Then
            tb_antal2.Text = ""
            tb_antal3.Text = ""
            tb_antal4.Text = ""
            tb_antal5.Text = ""
        End If
        'If tb_stans_manuel_uk.Text <> "" Then
        'tb_antal2.Text = ""
        'tb_antal3.Text = ""
        'tb_antal4.Text = ""
        'tb_antal5.Text = ""
        'End If

        'If sheetthickness > 10 Then
        'If Val(tb_buk1_x.Text) + Val(tb_buk1_y.Text) <> 0 Then
        'MsgBox("Buk i pladetykkelser over 10mm kan ikke beregnes, spørg bukoperatøren om buk tid (Indtast udfoldningsmål under største mål for videre beregning)")
        'MsgBox("Ved buk i pladetykkelser over 10mm er beregningen usikker, spørg bukoperatøren om buk tid ")
        'tb_buk1_y.Text = ""
        'tb_buk1_x.Text = ""
        'tb_pladetykkelse.Text = ""
        'ResetCalculations()
        'Exit Function
        'End If
        'End If

        If pbLoading = True Then
            Exit Sub
        End If
        If (pobjFilter_tb_bukmax_x.GetValue = 0 Or pobjFilter_tb_bukmax_Y.GetValue = 0 Or pobjFilter_tb_pladetykkelse.GetValue = 0) Then
            Exit Sub
        End If
        ClearAllCalculations()
        CalculateSubjectSize()


        stopvalue = CalculateBestPlate(Antal)
        If stopvalue = 1 Then
            Exit Sub
        End If
        subjectweight = CalculateSubjectweight()




        If Val(tb_buk1_x.Text) + Val(tb_buk1_y.Text) <> 0 Then
            CalculateBendTime(subjectweight, Val(Antal))
        Else
            lb_buk_tid.Text = ""
            lb_buk_opst.Text = ""
        End If


        If rb_klip.Checked = True Then
            CalculateCuttingTime(Antal)
        End If
        If rb_D_stans.Checked = True Then
            CalculatePunchTime(Antal)
        End If
        If rb_C_laser.Checked = True Then
            CalculateLaserTime(Antal)
        End If
        If rb_B_kombi.Checked = True Then
            CalculateCombiTime(Antal)
        End If
        CalculateThreads(Antal)
        CalculateCountersinks(Antal)
        CalculatePressnut(Val(Antal), Val(tb_presmøtrik_antal.Text))
        CalculatePresstag(Val(Antal), Val(tb_presstag_antal.Text))
        CalculateBoltweld(Val(Antal), Val(tb_boltesvejs_antal.Text))
        CalculateDeburring(Val(Antal))
        CalculateSteelmaster(Val(Antal), lb_udfold_x.Text, lb_udfold_y.Text, Lb_matrgruppe.Text)
        CalculateBrushing(Val(Antal))
        CalculateStraigthening(Val(Antal), lb_udfold_x.Text, lb_udfold_y.Text, Val(tb_pladetykkelse.Text))
        CalculateWelding(Antal)
        CalculateSpotWelding(Antal)
        CalculateGrindWelding(Antal)
        Calculateglasbl(Antal)
        Calculateslib(Antal)
        CalculateSurface1()
        CalculateSurface2()
        CalculateSurface3()
        'CalculateAdministration(Antal)
        CalculateSubtotals(Antal)
        CalculatePurchasingSubtotal(Antal)
        CalculateTotalPrice(Antal)
        CalculateGrouphours()
        materiale = cb_materiale.Text
        cb_materiale.Text = materiale

        'If sheetthickness > 10 Then
        'If Val(tb_buk1_x.Text) + Val(tb_buk1_y.Text) <> 0 Then
        'MsgBox("Ved buk i pladetykkelser over 10mm er beregningen usikker, spørg bukoperatøren om buk tid ")
        'End If
        'End If

    End Sub


    Private Function CalculateSubjectweight()
        Dim objDatabaseIO As DatabaseIO

        Dim thickness As Double
        Dim materialgroup As Integer
        Dim materialweight As Double
        Dim dblSubjectWeight As Double
        Dim area As Double
        objDatabaseIO = New DatabaseIO
        area = lb_nettoareal.Text
        thickness = Val(tb_pladetykkelse.Text)
        materialgroup = Val(Lb_matrgruppe.Text)

        materialweight = objDatabaseIO.GetMaterialWeight(materialgroup)

        dblSubjectWeight = area * materialweight * thickness

        lb_emnevægt.Text = FormatNumber(dblSubjectWeight, 2)
        CalculateSubjectweight = dblSubjectWeight

    End Function
    Public Function Getsupplierlist1() As String

        Dim objListSuppliersPulver As ListSuppliersPulver
        Dim objListSuppliersvåd As ListSuppliersVåd
        Dim objListSupplierschromit As ListSuppliersChromit
        Dim objListSupplierschromitAL As ListSuppliersChromital
        Dim objListSuppliersEloxering As ListSuppliersEloxering
        If cb_overfl_beh1.Text = "Ingen Overfladebehandling" Then
            Getsupplierlist1 = ""
            cb_overfl_leverandør1.Text = ""
            tb_overfl_opstart1.Text = ""
            tb_overfl_afdæk1.Text = ""
            tb_overfl_pris1.Text = ""
            tb_overfl_pris1_uk.Text = ""
            tb_overfl_pris100_1.Text = ""
            Exit Function
        Else
            If cb_overfl_beh1.Text = "Pulverlak" Then
                objListSuppliersPulver = New ListSuppliersPulver(cb_overfl_leverandør1)
                objListSuppliersPulver.List()
            End If
            If cb_overfl_beh1.Text = "Vådlak" Then
                objListSuppliersvåd = New ListSuppliersVåd(cb_overfl_leverandør1)
                objListSuppliersvåd.List()
            End If
            If cb_overfl_beh1.Text = "Chromit (jern)" Then
                objListSupplierschromit = New ListSuppliersChromit(cb_overfl_leverandør1)
                objListSupplierschromit.List()
            End If
            If cb_overfl_beh1.Text = "ChromitAL (alu)" Then
                objListSupplierschromitAL = New ListSuppliersChromital(cb_overfl_leverandør1)
                objListSupplierschromitAL.List()
            End If
            If cb_overfl_beh1.Text = "Eloxering (natur/sort)" Then
                objListSuppliersEloxering = New ListSuppliersEloxering(cb_overfl_leverandør1)
                objListSuppliersEloxering.List()
            End If
        End If
        Getsupplierlist1 = ""
    End Function
    Public Function Getsupplierlist2() As String

        Dim objListSuppliersPulver As ListSuppliersPulver
        Dim objListSuppliersvåd As ListSuppliersVåd
        Dim objListSupplierschromit As ListSuppliersChromit
        Dim objListSupplierschromitAL As ListSuppliersChromital
        Dim objListSuppliersEloxering As ListSuppliersEloxering

        If cb_overfl_beh2.Text = "Ingen Overfladebehandling" Then
            Getsupplierlist2 = ""
            cb_overfl_leverandør2.Text = ""
            tb_overfl_opstart2.Text = ""
            tb_overfl_afdæk2.Text = ""
            tb_overfl_pris2.Text = ""
            tb_overfl_pris2_uk.Text = ""
            tb_overfl_pris100_2.Text = ""
            Exit Function
        Else
            If cb_overfl_beh2.Text = "Pulverlak" Then
                objListSuppliersPulver = New ListSuppliersPulver(cb_overfl_leverandør2)
                objListSuppliersPulver.List()
            End If
            If cb_overfl_beh2.Text = "Vådlak" Then
                objListSuppliersvåd = New ListSuppliersVåd(cb_overfl_leverandør2)
                objListSuppliersvåd.List()
            End If
            If cb_overfl_beh2.Text = "Chromit (jern)" Then
                objListSupplierschromit = New ListSuppliersChromit(cb_overfl_leverandør2)
                objListSupplierschromit.List()
            End If
            If cb_overfl_beh2.Text = "ChromitAL (alu)" Then
                objListSupplierschromitAL = New ListSuppliersChromital(cb_overfl_leverandør2)
                objListSupplierschromitAL.List()
            End If
            If cb_overfl_beh2.Text = "Eloxering (natur/sort)" Then
                objListSuppliersEloxering = New ListSuppliersEloxering(cb_overfl_leverandør2)
                objListSuppliersEloxering.List()
            End If
        End If
        Getsupplierlist2 = ""

    End Function
    Public Function Getsupplierlist3() As String

        Dim objListSuppliersPulver As ListSuppliersPulver
        Dim objListSuppliersvåd As ListSuppliersVåd
        Dim objListSupplierschromit As ListSuppliersChromit
        Dim objListSupplierschromitAL As ListSuppliersChromital
        Dim objListSuppliersEloxering As ListSuppliersEloxering

        If cb_overfl_beh3.Text = "Ingen Overfladebehandling" Then
            Getsupplierlist3 = ""
            cb_overfl_leverandør3.Text = ""
            tb_overfl_opstart3.Text = ""
            tb_overfl_afdæk3.Text = ""
            tb_overfl_pris3.Text = ""
            tb_overfl_pris3_uk.Text = ""
            tb_overfl_pris100_3.Text = ""
            Exit Function
        Else
            If cb_overfl_beh3.Text = "Pulverlak" Then
                objListSuppliersPulver = New ListSuppliersPulver(cb_overfl_leverandør3)
                objListSuppliersPulver.List()
            End If
            If cb_overfl_beh3.Text = "Vådlak" Then
                objListSuppliersvåd = New ListSuppliersVåd(cb_overfl_leverandør3)
                objListSuppliersvåd.List()
            End If
            If cb_overfl_beh3.Text = "Chromit (jern)" Then
                objListSupplierschromit = New ListSuppliersChromit(cb_overfl_leverandør3)
                objListSupplierschromit.List()
            End If
            If cb_overfl_beh3.Text = "ChromitAL (alu)" Then
                objListSupplierschromitAL = New ListSuppliersChromital(cb_overfl_leverandør3)
                objListSupplierschromitAL.List()
            End If
            If cb_overfl_beh3.Text = "Eloxering (natur/sort)" Then
                objListSuppliersEloxering = New ListSuppliersEloxering(cb_overfl_leverandør3)
                objListSuppliersEloxering.List()
            End If
        End If
        CalculateSurfaceTreatment(Val(tb_antal1.Text))
        Getsupplierlist3 = ""
    End Function

    Public Function Getsupplierlist4() As String

        Dim objListSuppliers As ListSuppliers

        If cb_overfl_beh4.Text = "Ingen Overfladebehandling" Then
            Getsupplierlist4 = ""
            cb_overfl_leverandør4.Text = ""
            tb_overfl_opstart4.Text = ""
            tb_overfl_afdæk4.Text = ""
            tb_overfl_pris4.Text = ""
            'tb_overfl_pris4_uk.Text = ""
            tb_overfl_pris100_4.Text = ""
            Exit Function
        Else
            objListSuppliers = New ListSuppliers(cb_overfl_leverandør4)
            objListSuppliers.List()
        End If
        CalculateSurfaceTreatment(Val(tb_antal1.Text))
        Getsupplierlist4 = ""
    End Function
    Public Function Getsupplierlist5() As String

        Dim objListSuppliers As ListSuppliers

        If cb_overfl_beh5.Text = "Ingen Overfladebehandling" Then
            Getsupplierlist5 = ""
            cb_overfl_leverandør5.Text = ""
            tb_overfl_opstart5.Text = ""
            tb_overfl_afdæk5.Text = ""
            tb_overfl_pris5.Text = ""
            'tb_overfl_pris5_uk.Text = ""
            tb_overfl_pris100_5.Text = ""
            Exit Function
        Else
            objListSuppliers = New ListSuppliers(cb_overfl_leverandør5)
            objListSuppliers.List()
        End If
        CalculateSurfaceTreatment(Val(tb_antal1.Text))
        Getsupplierlist5 = ""
    End Function
    Public Sub CalculateSurface1()
        Dim Faktor As Double
        Dim stkppris As Double
        Dim areal As String
        Dim surfacefactor As Double
        Dim surfacenumber As Double
        Dim paintprotection As Double

        If lb_nettoareal.Text = "" Then
            Exit Sub
        End If
        areal = lb_nettoareal.Text
        If areal < 0.01 Then
            areal = 0.01
        End If

        If cb_1side_1.CheckState = CheckState.Checked Then
            Faktor = 1
        Else : Faktor = 2
        End If

        paintprotection = Val(tb_presstag_antal.Text) + Val(tb_presmøtrik_antal.Text) + Val(tb_boltesvejs_antal.Text)

        If paintprotection = 0 Then
            paintprotection = 1
        End If

        surfacenumber = (Val(tb_numberofsurface.Text))

        If surfacenumber = 1 Then
            surfacefactor = 1
        End If
        If surfacenumber = 2 Then
            surfacefactor = 1.1
        End If
        If surfacenumber = 3 Then
            surfacefactor = 1.2
        End If
        If surfacenumber = 4 Then
            surfacefactor = 1.3
        End If
        If surfacenumber = 5 Then
            surfacefactor = 1.4
        End If
        If surfacenumber = 6 Then
            surfacefactor = 1.5
        End If

        If cb_overfl_beh1.Text = "Pulverlak" Then
            If cb_overfl_leverandør1.Text = "Laduco" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Miltona" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Næstved Pulverlakering" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Jowis" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Greve Pulverlakering" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Cuwi" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "MalerFlex" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (250 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If

        End If
        If cb_overfl_beh1.Text = "Vådlak" Then
            If cb_overfl_leverandør1.Text = "Laduco" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Miltona" Then
                tb_overfl_opstart1.Text = 400
                tb_overfl_afdæk1.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Næstved Pulverlakering" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 2
                stkppris = (325 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Jowis" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "MalerFlex" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk1.Text) * paintprotection)
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If

        End If
        If cb_overfl_beh1.Text = "Chromit (jern)" Then
            If cb_overfl_leverandør1.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 115 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Stjernecrom" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Croma" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 160 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
        End If
        If cb_overfl_beh1.Text = "ChromitAL (alu)" Then
            If cb_overfl_leverandør1.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 125 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Stjernecrom" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Croma" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Aluscan" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 175 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
        End If
        If cb_overfl_beh1.Text = "Eloxering (natur/sort)" Then
            If cb_overfl_leverandør1.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 600 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Stjernecrom" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Croma" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør1.Text = "Aluscan" Then
                tb_overfl_opstart1.Text = 300
                tb_overfl_afdæk1.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris1.Text = FormatNumber(stkppris, 2)
            End If
        End If


    End Sub
    Public Sub CalculateSurface2()
        Dim Faktor As Double
        Dim stkppris As Double
        Dim areal As String
        Dim surfacefactor As Double
        Dim surfacenumber As Double
        Dim paintprotection As Double

        If lb_nettoareal.Text = "" Then
            Exit Sub
        End If
        areal = lb_nettoareal.Text
        If areal < 0.01 Then
            areal = 0.01
        End If

        If cb_1side_2.CheckState = CheckState.Checked Then
            Faktor = 1
        Else : Faktor = 2
        End If

        paintprotection = Val(tb_presstag_antal.Text) + Val(tb_presmøtrik_antal.Text) + Val(tb_boltesvejs_antal.Text)

        If paintprotection = 0 Then
            paintprotection = 1
        End If


        surfacenumber = (Val(tb_numberofsurface.Text))

        If surfacenumber = 1 Then
            surfacefactor = 1
        End If
        If surfacenumber = 2 Then
            surfacefactor = 1.1
        End If
        If surfacenumber = 3 Then
            surfacefactor = 1.2
        End If
        If surfacenumber = 4 Then
            surfacefactor = 1.3
        End If
        If surfacenumber = 5 Then
            surfacefactor = 1.4
        End If
        If surfacenumber = 6 Then
            surfacefactor = 1.5
        End If

        If cb_overfl_beh2.Text = "Pulverlak" Then
            If cb_overfl_leverandør2.Text = "Laduco" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Miltona" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Næstved Pulverlakering" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Jowis" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Greve Pulverlakering" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Cuwi" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "MalerFlex" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (250 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If

        End If
        If cb_overfl_beh2.Text = "Vådlak" Then
            If cb_overfl_leverandør2.Text = "Laduco" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Miltona" Then
                tb_overfl_opstart2.Text = 400
                tb_overfl_afdæk2.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Næstved Pulverlakering" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 2
                stkppris = (325 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Jowis" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "MalerFlex" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk2.Text) * paintprotection)
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If

        End If
        If cb_overfl_beh2.Text = "Chromit (jern)" Then
            If cb_overfl_leverandør2.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 115 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Stjernecrom" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Croma" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 160 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
        End If
        If cb_overfl_beh2.Text = "ChromitAL (alu)" Then
            If cb_overfl_leverandør2.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 125 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Stjernecrom" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Croma" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Aluscan" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 175 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
        End If
        If cb_overfl_beh2.Text = "Eloxering (natur/sort)" Then
            If cb_overfl_leverandør2.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 600 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Stjernecrom" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Croma" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør2.Text = "Aluscan" Then
                tb_overfl_opstart2.Text = 300
                tb_overfl_afdæk2.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris2.Text = FormatNumber(stkppris, 2)
            End If
        End If
    End Sub
    Public Sub CalculateSurface3()

        Dim Faktor As Double
        Dim stkppris As Double
        Dim areal As String
        Dim surfacefactor As Double
        Dim surfacenumber As Double
        Dim paintprotection As Double

        If lb_nettoareal.Text = "" Then
            Exit Sub
        End If
        areal = lb_nettoareal.Text
        If areal < 0.01 Then
            areal = 0.01
        End If

        If cb_1side_3.CheckState = CheckState.Checked Then
            Faktor = 1
        Else : Faktor = 2
        End If

        paintprotection = Val(tb_presstag_antal.Text) + Val(tb_presmøtrik_antal.Text) + Val(tb_boltesvejs_antal.Text)

        If paintprotection = 0 Then
            paintprotection = 1
        End If


        surfacenumber = (Val(tb_numberofsurface.Text))

        If surfacenumber = 1 Then
            surfacefactor = 1
        End If
        If surfacenumber = 2 Then
            surfacefactor = 1.1
        End If
        If surfacenumber = 3 Then
            surfacefactor = 1.2
        End If
        If surfacenumber = 4 Then
            surfacefactor = 1.3
        End If
        If surfacenumber = 5 Then
            surfacefactor = 1.4
        End If
        If surfacenumber = 6 Then
            surfacefactor = 1.5
        End If

        If cb_overfl_beh3.Text = "Pulverlak" Then
            If cb_overfl_leverandør3.Text = "Laduco" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Miltona" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Næstved Pulverlakering" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Jowis" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Greve Pulverlakering" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Cuwi" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (200 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "MalerFlex" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (250 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If

        End If
        If cb_overfl_beh3.Text = "Vådlak" Then
            If cb_overfl_leverandør3.Text = "Laduco" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Miltona" Then
                tb_overfl_opstart3.Text = 400
                tb_overfl_afdæk3.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Næstved Pulverlakering" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 2
                stkppris = (325 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Jowis" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "MalerFlex" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 2
                stkppris = (450 * areal * Faktor) + (Val(tb_overfl_afdæk3.Text) * paintprotection)
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If

        End If
        If cb_overfl_beh3.Text = "Chromit (jern)" Then
            If cb_overfl_leverandør3.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 115 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Stjernecrom" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Croma" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 160 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
        End If
        If cb_overfl_beh3.Text = "ChromitAL (alu)" Then
            If cb_overfl_leverandør3.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 125 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Stjernecrom" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Croma" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 150 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Aluscan" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 175 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
        End If
        If cb_overfl_beh3.Text = "Eloxering (natur/sort)" Then
            If cb_overfl_leverandør3.Text = "Værløse Galvaniske" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 600 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Stjernecrom" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Croma" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
            If cb_overfl_leverandør3.Text = "Aluscan" Then
                tb_overfl_opstart3.Text = 300
                tb_overfl_afdæk3.Text = 0
                stkppris = 700 * areal * 2 * surfacefactor
                tb_overfl_pris3.Text = FormatNumber(stkppris, 2)
            End If
        End If


    End Sub



    Public Function CalculateSurfaceTreatment(ByVal antal As Double)
        Dim surfacetreatmentcost As Double
        Dim surfacetreatment As Double
        Dim stykpris As Double
        Dim antal100 As Double

        'antal100 = 100
        antal100 = antal

        'lb_overfl_antal.Text = tb_antal1.Text

        surfacetreatment = 0
        If Val(tb_overfl_pris1.Text) = 0 Then
            tb_overfl_pris100_1.Text = ""
        End If
        If Val(tb_overfl_pris1.Text) <> 0 Then
            If Val(tb_overfl_pris1_uk.Text) = 0.0 Then
                stykpris = Val(pobjFilter_tb_overfl_Pris1.GetValue)
            Else
                stykpris = Val(tb_overfl_pris1_uk.Text)
            End If
            'surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart1.GetValue), Val(pobjFilter_tb_overfl_afdæk1.GetValue), stykpris, Val(pobjFilter_tb_overfl_Avance1.GetValue), antal)
            'surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_1.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart1.GetValue), Val(pobjFilter_tb_overfl_afdæk1.GetValue), stykpris, Val(pobjFilter_tb_overfl_Avance1.GetValue), antal100)) / antal100, 2)
            surfacetreatment = pobjFilter_tb_overfl_tilbudpris1_uk.GetValue * antal
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment

        End If
        If Val(tb_overfl_pris2.Text) = 0 Then
            tb_overfl_pris100_2.Text = ""
        End If
        If Val(tb_overfl_pris2.Text) <> 0 Then
            If Val(tb_overfl_pris2_uk.Text) = 0.0 Then
                stykpris = Val(pobjFilter_tb_overfl_Pris2.GetValue)
            Else
                stykpris = Val(tb_overfl_pris2_uk.Text)
            End If
            'surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart2.GetValue), Val(pobjFilter_tb_overfl_afdæk2.GetValue), stykpris, Val(pobjFilter_tb_overfl_Avance2.GetValue), antal)
            'surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_2.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart2.GetValue), Val(pobjFilter_tb_overfl_afdæk2.GetValue), stykpris, Val(pobjFilter_tb_overfl_Avance2.GetValue), antal100)) / antal100, 2)
            surfacetreatment = pobjFilter_tb_overfl_tilbudpris2_uk.GetValue * antal
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
        End If
        If Val(tb_overfl_pris3.Text) = 0 Then
            tb_overfl_pris100_3.Text = ""
        End If
        If Val(tb_overfl_pris3.Text) <> 0 Then
            If Val(tb_overfl_pris3_uk.Text) = 0.0 Then
                stykpris = Val(pobjFilter_tb_overfl_Pris3.GetValue)
            Else
                stykpris = Val(tb_overfl_pris3_uk.Text)
            End If
            'surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart3.GetValue), Val(pobjFilter_tb_overfl_afdæk3.GetValue), stykpris, Val(pobjFilter_tb_overfl_Avance3.GetValue), antal)
            'surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_3.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart3.GetValue), Val(pobjFilter_tb_overfl_afdæk3.GetValue), stykpris, Val(pobjFilter_tb_overfl_Avance3.GetValue), antal100)) / antal100, 2)
            surfacetreatment = pobjFilter_tb_overfl_tilbudpris3_uk.GetValue * antal
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
        End If

        If cb_overfl_beh4.Text = "Ingen Overfladebehandling" Then
            cb_overfl_leverandør4.Text = ""
            tb_overfl_opstart4.Text = ""
            tb_overfl_afdæk4.Text = ""
            tb_overfl_pris4.Text = ""
            tb_overfl_pris100_4.Text = ""
        End If

        If Val(tb_overfl_pris4.Text) <> 0 Then
            surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart4.GetValue), Val(pobjFilter_tb_overfl_afdæk4.GetValue), Val(pobjFilter_tb_overfl_Pris4.GetValue), Val(pobjFilter_tb_overfl_Avance4.GetValue), antal)
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_4.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart4.GetValue), Val(pobjFilter_tb_overfl_afdæk4.GetValue), Val(pobjFilter_tb_overfl_Pris4.GetValue), Val(pobjFilter_tb_overfl_Avance4.GetValue), antal100)) / antal100, 2)
        Else
            tb_overfl_pris100_4.Text = ""
        End If

        If cb_overfl_beh5.Text = "Ingen Overfladebehandling" Then
            cb_overfl_leverandør5.Text = ""
            tb_overfl_opstart5.Text = ""
            tb_overfl_afdæk5.Text = ""
            tb_overfl_pris5.Text = ""
            tb_overfl_pris100_5.Text = ""
        End If

        If Val(tb_overfl_pris5.Text) <> 0 Then
            surfacetreatment = CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart5.GetValue), Val(pobjFilter_tb_overfl_afdæk5.GetValue), Val(pobjFilter_tb_overfl_Pris5.GetValue), Val(pobjFilter_tb_overfl_Avance5.GetValue), antal)
            surfacetreatmentcost = surfacetreatmentcost + surfacetreatment
            tb_overfl_pris100_5.Text = FormatNumber((CalculateSurfaceTreatmentCost(Val(pobjFilter_tb_overfl_opstart5.GetValue), Val(pobjFilter_tb_overfl_afdæk5.GetValue), Val(pobjFilter_tb_overfl_Pris5.GetValue), Val(pobjFilter_tb_overfl_Avance5.GetValue), antal100)) / antal100, 2)
        Else
            tb_overfl_pris100_5.Text = ""
        End If
        CalculateSurfaceTreatment = surfacetreatmentcost
    End Function

    Public Function CalculateSurfaceTreatmentCost(ByVal opstart As Double, ByVal afdækning As Double, ByVal pris As Double, ByVal avance As Double, ByVal antal As Double)
        CalculateSurfaceTreatmentCost = opstart + (((afdækning * ((Val(lb_gevind_antal.Text) + Val(tb_presmøtrik_antal.Text) + Val(tb_boltesvejs_antal.Text))) + pris) * antal))

        CalculateSurfaceTreatmentCost = CalculateSurfaceTreatmentCost + CalculateSurfaceTreatmentCost * avance / 100


    End Function
    Private Sub CalculatePunchTime(ByVal antal As Double)
        Dim objCombiPunchCalculations As PunchCalculations
        Dim Difficultfactor As Double
        Dim objkommakontrol As Kommakontrol
        Dim cnctime As Double
        Dim startup As Double
        Dim objDatabaseIO As DatabaseIO
        Dim stansialt As Double
        objkommakontrol = New Kommakontrol


        objCombiPunchCalculations = New PunchCalculations

        cnctime = objCombiPunchCalculations.CalculatePunchTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_toolshift.Text), Val(tb_slag_til_huller.Text), Val(Convert.ToDouble(lb_faktor.Text)), antal)
        lb_gruppe1_tid.Text = objkommakontrol.Decimalkontrol(cnctime)
        objDatabaseIO = New DatabaseIO

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        startup = Difficultfactor * objDatabaseIO.GetPunchStartup

        If startup < 15 Then
            startup = 15
        End If


        lb_gruppe1_opstart.Text = objkommakontrol.Decimalkontrol(startup)
        stansialt = cnctime + startup
        lb_stans_ialt.Text = objkommakontrol.Decimalkontrol(stansialt)


    End Sub


    Private Sub CalculateLaserTime(ByVal antal As Double)
        Dim objLaserCalculations As LaserCalculations
        Dim objkommakontrol As Kommakontrol
        Dim Difficultfactor As Double
        Dim objDatabaseIO As DatabaseIO
        Dim cnctime As Double
        Dim startup As Double
        Dim laserialt As Double
        Dim nitrogen As Double
        Dim nitrogencost As Double
        Dim cuttingtime As Double
        'Dim calculetacuttingtime As Double


        objkommakontrol = New Kommakontrol
        objLaserCalculations = New LaserCalculations
        cnctime = objLaserCalculations.CalculateLaserTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_hulantal_1C.Text), Val(tb_hulantal_2C.Text), Val(tb_hulantal_3C.Text), Val(tb_hulantal_4C.Text), Val(tb_cuttinglength_C.Text), Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text), pobjListMaterials.GetMaterialType(), Val(Convert.ToDouble(lb_faktor.Text)), antal, Val(lb_modulstr.Text))

        objLaserCalculations = New LaserCalculations
        nitrogen = objLaserCalculations.calculatenitrogen(Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text))

        objLaserCalculations = New LaserCalculations
        cuttingtime = objLaserCalculations.calculatecuttingtime(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_hulantal_1C.Text), Val(tb_hulantal_2C.Text), Val(tb_hulantal_3C.Text), Val(tb_hulantal_4C.Text), Val(tb_cuttinglength_C.Text), Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text), pobjListMaterials.GetMaterialType(), Val(Convert.ToDouble(lb_faktor.Text)), antal, Val(lb_modulstr.Text))

        lb_laserCNC_tid.Text = objkommakontrol.Decimalkontrol(cnctime)
        objDatabaseIO = New DatabaseIO

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)
        'startup = Difficultfactor * objDatabaseIO.GetLaserStartup(Val(tb_pladetykkelse.Text))
        startup = objDatabaseIO.GetLaserStartup(Val(tb_pladetykkelse.Text))
        'If startup < 15 Then
        'startup = 15
        'End If

        If tb_samkørsel.Text <> "" Then
            startup = startup * 0.1
        End If

        lb_laser_opstart.Text = objkommakontrol.Decimalkontrol(startup)
        laserialt = cnctime + startup
        lb_laser_ialt.Text = objkommakontrol.Decimalkontrol(laserialt)

        nitrogencost = nitrogen * cuttingtime / 60 * 2.22 'nitrogenpris 2.22 kr./M3

        lb_nitrogen.Text = objkommakontrol.Decimalkontrol(nitrogencost)

    End Sub

    Private Sub CalculateCombiTime(ByVal antal As Double)
        Dim objCombiLaserCalculations As CombiLaserCalculations
        Dim objCombiPunchCalculations As CombiPunchCalculations
        Dim objkommakontrol As Kommakontrol
        'Dim cnctime As Double
        Dim mantime As Double
        Dim objDatabaseIO As DatabaseIO
        Dim PunchtimeCNC As Double
        Dim LasertimeCNC As Double
        Dim Difficultfactor As Double
        Dim combiialt As Double
        objkommakontrol = New Kommakontrol

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        objCombiLaserCalculations = New CombiLaserCalculations
        LasertimeCNC = objCombiLaserCalculations.CalculateLaserTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_hulantal_1B.Text), Val(tb_hulantal_2B.Text), Val(tb_hulantal_3B.Text), Val(tb_hulantal_4B.Text), Val(tb_cuttinglength_B.Text), Val(Lb_matrgruppe.Text), Val(tb_pladetykkelse.Text), Val(Convert.ToDouble(lb_faktor.Text)), antal)
        If LasertimeCNC < 1 Then
            LasertimeCNC = 1
        End If
        objCombiPunchCalculations = New CombiPunchCalculations
        PunchtimeCNC = objCombiPunchCalculations.CalculatePunchTimes(Val(lb_antalplader.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text), Val(lb_udfold_x.Text), Val(lb_udfold_y.Text), Val(tb_toolshift_B.Text), Val(tb_slag_til_huller_B.Text), Val(Convert.ToDouble(lb_faktor.Text)), antal)


        objDatabaseIO = New DatabaseIO

        mantime = objDatabaseIO.GetCombiStartup * Difficultfactor
        If mantime < 15 Then
            mantime = 15
        End If

        lb_Combi_opstart.Text = objkommakontrol.Decimalkontrol(mantime)

        'cnctime = LasertimeCNC + PunchtimeCNC
        lb_CombiCNC_tid.Text = objkommakontrol.Decimalkontrol(LasertimeCNC)
        lb_CombiCNCstans_tid.Text = objkommakontrol.Decimalkontrol(PunchtimeCNC)

        combiialt = LasertimeCNC + PunchtimeCNC + mantime
        lb_combi_ialt.Text = objkommakontrol.Decimalkontrol(combiialt)

    End Sub
    Private Sub CalculateCuttingTime(ByVal antal As Double)


        Dim objCuttingCalculations As CuttingCalculations
        Dim objkommakontrol As Kommakontrol

        Dim CountRow As Integer
        Dim Subjects1Sheet As Integer
        Dim NumberofSheets As Integer
        Dim Cuttingtime As Double
        Dim Difficultfactor As Double
        Dim klipopstart As Double
        Dim kliptid As Double
        Dim kliptidialt As Double
        objkommakontrol = New Kommakontrol

        If lb_faktor.Text = "" Then
            Exit Sub
        End If
        Difficultfactor = Convert.ToDouble(lb_faktor.Text)

        CountRow = Convert.ToInt32(lb_rækkeantal.Text)
        Subjects1Sheet = Val(lb_emner_prplade.Text)
        NumberofSheets = Val(lb_antalplader.Text)

        objCuttingCalculations = New CuttingCalculations
        Cuttingtime = objCuttingCalculations.CalculateCuttingTime(Val(lb_rækkeantal.Text), Val(lb_emner_prplade.Text), Val(lb_antalplader.Text), Val(tb_pladetykkelse.Text), Val(lb_modulstr.Text), antal)


        klipopstart = Difficultfactor * objCuttingCalculations.GetCuttingStartup(Val(lb_antalplader.Text), Val(tb_pladetykkelse.Text), Val(lb_pladeformatX.Text), Val(lb_pladeformatY.Text))
        If klipopstart < 10 Then
            klipopstart = 10
        End If
        kliptid = Difficultfactor * Cuttingtime.ToString()

        kliptidialt = kliptid + klipopstart

        lb_klip_opstart.Text = objkommakontrol.Decimalkontrol(klipopstart)
        lb_klip_tid.Text = objkommakontrol.Decimalkontrol(kliptid)
        lb_klip_ialt.Text = objkommakontrol.Decimalkontrol(kliptidialt)
    End Sub
    Public Sub New()
        pbLoading = True
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        pobjListMaterials = New ListMaterials(cb_materiale, Lb_matrgruppe, Lb_klasse, Lb_Kilopris)
        pobjListMaterials.List()

        pobjListSurface1 = New Listsurfaces(cb_overfl_beh1)
        pobjListSurface1.List()
        pobjListSurface2 = New Listsurfaces(cb_overfl_beh2)
        pobjListSurface2.List()
        pobjListSurface3 = New Listsurfaces(cb_overfl_beh3)
        pobjListSurface3.List()
        pobjListSurface4 = New Listsurfaces1(cb_overfl_beh4)
        pobjListSurface4.List()
        pobjListSurface5 = New Listsurfaces1(cb_overfl_beh5)
        pobjListSurface5.List()

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
        pobjFilter_tb_valsning_uk = New FilterKeys(tb_valsning_uk, lb_valsetid, Me)
        pobjFilter_tb_stepbuk_uk = New FilterKeys(tb_stepbuk_uk, lb_stepbuk, Me)
        pobjFilter_tb_stepantal = New FilterKeys(tb_stepantal, Nothing, Me)
        'stans
        pobjFilter_tb_toolshift = New FilterKeys(tb_toolshift, Nothing, Me)
        pobjFilter_tb_slag_til_huller = New FilterKeys(tb_slag_til_huller, Nothing, Me)
        pobjFilter_tb_CNCmin_uk = New FilterKeys(tb_CNCmin_uk, lb_gruppe1_tid, Me)
        pobjFilter_tb_gruppe1_opstart_uk = New FilterKeys(tb_gruppe1_opstart_uk, lb_gruppe1_opstart, Me)
        'laser
        pobjFilter_tb_hulantal_1C = New FilterKeys(tb_hulantal_1C, Nothing, Me)
        pobjFilter_tb_hulantal_2C = New FilterKeys(tb_hulantal_2C, Nothing, Me)
        pobjFilter_tb_hulantal_3C = New FilterKeys(tb_hulantal_3C, Nothing, Me)
        pobjFilter_tb_hulantal_4C = New FilterKeys(tb_hulantal_4C, Nothing, Me)
        pobjFilter_tb_cuttinglength_C = New FilterKeys(tb_cuttinglength_C, Nothing, Me)
        pobjFilter_tb_laserCNC_tid_uk = New FilterKeys(tb_laserCNC_tid_uk, lb_laserCNC_tid, Me)
        pobjFilter_tb_laser_opstart_uk = New FilterKeys(tb_laser_opstart_uk, lb_laser_opstart, Me)
        'kombi
        pobjFilter_tb_toolshift_B = New FilterKeys(tb_toolshift_B, Nothing, Me)
        pobjFilter_tb_slag_til_huller_B = New FilterKeys(tb_slag_til_huller_B, Nothing, Me)
        pobjFilter_tb_hulantal_1B = New FilterKeys(tb_hulantal_1B, Nothing, Me)
        pobjFilter_tb_hulantal_2B = New FilterKeys(tb_hulantal_2B, Nothing, Me)
        pobjFilter_tb_hulantal_3B = New FilterKeys(tb_hulantal_3B, Nothing, Me)
        pobjFilter_tb_hulantal_4B = New FilterKeys(tb_hulantal_4B, Nothing, Me)

        pobjFilter_tb_cuttinglength_B = New FilterKeys(tb_cuttinglength_B, Nothing, Me)
        pobjFilter_tb_combiCNC_tid_uk = New FilterKeys(tb_combiCNC_tid_uk, lb_CombiCNC_tid, Me)
        pobjFilter_tb_combiCNCstans_tid_uk = New FilterKeys(tb_CombiCNCstans_tid_uk, lb_CombiCNCstans_tid, Me)
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
        pobjFilter_tb_presstag_uk = New FilterKeys(tb_presstag_uk, lb_presstag, Me)
        pobjFilter_tb_presstag_kr_uk = New FilterKeys(tb_presstag_kr_uk, lb_presstag_kr, Me)
        pobjFilter_tb_boltesvejs_uk = New FilterKeys(tb_boltesvejs_uk, lb_boltesvejs, Me)
        pobjFilter_tb_svejsestag_kr_uk = New FilterKeys(tb_svejsestag_kr_uk, lb_svejsestag_kr, Me)
        pobjFilter_tb_tilsatsmatr_kr_uk = New FilterKeys(tb_tilsatsmatr_kr_uk, lb_tilsatsmatr_kr, Me)
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
        pobjFilter_tb_numberofspots = New FilterKeys(tb_numberofspots, Nothing, Me)
        pobjFilter_tb_numberofspotweldseams = New FilterKeys(tb_numberofspotweldseams, Nothing, Me)
        pobjFilter_tb_spotweld_uk = New FilterKeys(tb_spotweld_uk, lb_spotweld, Me)
        'gruppe4
        pobjFilter_tb_numberofwelds = New FilterKeys(tb_numberofwelds, Nothing, Me)
        pobjFilter_tb_weldlength = New FilterKeys(tb_weldlength, Nothing, Me)
        pobjFilter_tb_tackweld_uk = New FilterKeys(tb_tackweld_uk, lb_tackweld, Me)
        pobjFilter_tb_weld_uk = New FilterKeys(tb_weld_uk, lb_weld, Me)
        pobjFilter_tb_grind_weld_uk = New FilterKeys(tb_grind_weld_uk, lb_grind_weld, Me)
        pobjFilter_tb_rettesvejs_uk = New FilterKeys(tb_rettesvejs_uk, lb_rettesvejs_tid, Me)
        pobjFilter_tb_tapsvejs_uk = New FilterKeys(tb_tapsvejs_uk, lb_tapsvejs_tid, Me)
        pobjFilter_tb_tapantal = New FilterKeys(tb_tapantal, Nothing, Me)
        'gruppe5
        pobjFilter_tb_glasbl_uk = New FilterKeys(tb_glasbl_uk, lb_glasbl, Me)
        pobjFilter_tb_slib_uk = New FilterKeys(tb_slib_uk, lb_slib, Me)
        'gruppe6
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
        pobjFilter_tb_Kilopris_uk = New FilterKeys(tb_Kilopris_uk, Lb_Kilopris, Me)

        'indkøb
        pobjFilter_tb_overfl_opstart1 = New FilterKeys(tb_overfl_opstart1, Nothing, Me)
        pobjFilter_tb_overfl_opstart2 = New FilterKeys(tb_overfl_opstart2, Nothing, Me)
        pobjFilter_tb_overfl_opstart3 = New FilterKeys(tb_overfl_opstart3, Nothing, Me)
        pobjFilter_tb_overfl_opstart4 = New FilterKeys(tb_overfl_opstart4, Nothing, Me)
        pobjFilter_tb_overfl_opstart5 = New FilterKeys(tb_overfl_opstart5, Nothing, Me)
        pobjFilter_tb_overfl_afdæk1 = New FilterKeys(tb_overfl_afdæk1, Nothing, Me)
        pobjFilter_tb_overfl_afdæk2 = New FilterKeys(tb_overfl_afdæk2, Nothing, Me)
        pobjFilter_tb_overfl_afdæk3 = New FilterKeys(tb_overfl_afdæk3, Nothing, Me)
        pobjFilter_tb_overfl_afdæk4 = New FilterKeys(tb_overfl_afdæk4, Nothing, Me)
        pobjFilter_tb_overfl_afdæk5 = New FilterKeys(tb_overfl_afdæk5, Nothing, Me)
        pobjFilter_tb_overfl_Pris1 = New FilterKeys(tb_overfl_pris1, Nothing, Me)
        pobjFilter_tb_overfl_Pris2 = New FilterKeys(tb_overfl_pris2, Nothing, Me)
        pobjFilter_tb_overfl_Pris3 = New FilterKeys(tb_overfl_pris3, Nothing, Me)
        pobjFilter_tb_overfl_Pris1_uk = New FilterKeys(tb_overfl_pris1_uk, Nothing, Me)
        pobjFilter_tb_overfl_Pris2_uk = New FilterKeys(tb_overfl_pris2_uk, Nothing, Me)
        pobjFilter_tb_overfl_Pris3_uk = New FilterKeys(tb_overfl_pris3_uk, Nothing, Me)
        pobjFilter_tb_overfl_Pris4 = New FilterKeys(tb_overfl_pris4, Nothing, Me)
        pobjFilter_tb_overfl_Pris5 = New FilterKeys(tb_overfl_pris5, Nothing, Me)
        pobjFilter_tb_overfl_Avance1 = New FilterKeys(tb_overfl_avance1, Nothing, Me)
        pobjFilter_tb_overfl_Avance2 = New FilterKeys(tb_overfl_avance2, Nothing, Me)
        pobjFilter_tb_overfl_Avance3 = New FilterKeys(tb_overfl_avance3, Nothing, Me)
        pobjFilter_tb_overfl_Avance4 = New FilterKeys(tb_overfl_avance4, Nothing, Me)
        pobjFilter_tb_overfl_Avance5 = New FilterKeys(tb_overfl_avance5, Nothing, Me)
        pobjFilter_tb_overfl_tilbudpris1_uk = New FilterKeys(tb_overfl_tilbudpris1_uk, tb_overfl_pris100_1, Me)
        pobjFilter_tb_overfl_tilbudpris2_uk = New FilterKeys(tb_overfl_tilbudpris2_uk, tb_overfl_pris100_2, Me)
        pobjFilter_tb_overfl_tilbudpris3_uk = New FilterKeys(tb_overfl_tilbudpris3_uk, tb_overfl_pris100_3, Me)

        'pobjFilter_tb_numberofsurface = New FilterKeys(tb_numberofsurface, Nothing, Me)
        pobjFilter_tb_pladetykkelse = New FilterKeys(tb_pladetykkelse, Nothing, Me)


        'pobjFilter_tb_antal1 = New FilterKeys(tb_antal1, Nothing, Me)
        'pobjFilter_tb_antal2 = New FilterKeys(tb_antal2, Nothing, Me)
        'pobjFilter_tb_antal3 = New FilterKeys(tb_antal3, Nothing, Me)
        'pobjFilter_tb_antal4 = New FilterKeys(tb_antal4, Nothing, Me)
        'pobjFilter_tb_antal5 = New FilterKeys(tb_antal5, Nothing, Me)

        pbLoading = False
        lb_dato.Text = Today
        lb_dato_opr.Text = Today

        'lb_operatør.Text = System.Environment.UserName
    End Sub
    'OPDATERING AF UDREGNINGER VED ÆNDRINGER

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_afgrat.CheckedChanged

    End Sub
    Private Sub lb_udfold_x_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_udfold_x.TextChanged
        lb_buk_tid.Text = ""
        lb_buk_opst.Text = ""
        CalculateOrdrestr()
    End Sub
    Private Sub lb_udfold_y_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_udfold_y.TextChanged
        lb_buk_tid.Text = ""
        lb_buk_opst.Text = ""
        CalculateOrdrestr()
    End Sub
    Private Sub lb_børste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_brush.Click

    End Sub
    Private Sub tb_antal1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_antal1.TextChanged

        Resetunderkend()

        If tb_antal1.Text = "" Then
            lb_mand1.Text = ""
            lb_cnc1.Text = ""
            lb_mand1_tid.Text = ""
            lb_CNC1_tid.Text = ""
            lb_timer1.Text = ""
            lb_indkøb1.Text = ""
            lb_råvarer1.Text = ""
            lb_råvarerstk1.Text = ""
            lb_samlet1.Text = ""
            lb_salg1.Text = ""
            tb_tilbud1.Text = ""
        Else
            CalculateOrdrestr()

        End If
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
            lb_råvarerstk2.Text = ""
            lb_samlet2.Text = ""
            lb_salg2.Text = ""
            tb_tilbud2.Text = ""
        Else
            CalculateOrdrestr()
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
            lb_råvarerstk3.Text = ""
            lb_samlet3.Text = ""
            lb_salg3.Text = ""
            tb_tilbud3.Text = ""
        Else
            CalculateOrdrestr()
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
            lb_råvarerstk4.Text = ""
            lb_samlet4.Text = ""
            lb_salg4.Text = ""
            tb_tilbud4.Text = ""
        Else
            CalculateOrdrestr()
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
            lb_råvarerstk5.Text = ""
            lb_samlet5.Text = ""
            lb_salg5.Text = ""
            tb_tilbud5.Text = ""
        Else
            CalculateOrdrestr()
        End If
    End Sub

    Private Sub rb_D_stans_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_D_stans.CheckedChanged
        CalculateOrdrestr()
    End Sub

    Private Sub tb_vibration_uk_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_vibration_uk.TextChanged
        tb_antal2.Text = ""
        tb_antal3.Text = ""
        tb_antal4.Text = ""
        tb_antal5.Text = ""
    End Sub
    'Private Sub tb_stans_manuel_uk_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_stans_manuel_uk.TextChanged
    'tb_antal2.Text = ""
    'tb_antal3.Text = ""
    'tb_antal4.Text = ""
    'tb_antal5.Text = ""
    'End Sub


    Private Sub rb_C_laser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_C_laser.CheckedChanged
        CalculateOrdrestr()
    End Sub

    Private Sub rb_B_kombi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_B_kombi.CheckedChanged
        CalculateOrdrestr()
    End Sub

    Private Sub rb_klip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_klip.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_netto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_netto.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_brutto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_brutto.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_tig_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_tig.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_rustfri_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_rustfri.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_jern_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_jern.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_numberofwelds_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_numberofwelds.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_weldlength_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_weldlength.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_Spotweld_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_spotweld.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_weld_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_weld.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_tackweld_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tackweld.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_Grind_weld_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_grind_weld.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_factor1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_factor1.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_factor2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_factor2.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_factor3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_factor3.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_factor4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_factor4.CheckedChanged
        CalculateOrdrestr()
    End Sub
    Private Sub rb_factor5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_factor5.CheckedChanged
        If tb_tapantal.Text <> "" Then
            CalculateOrdrestr()
        End If
    End Sub
    Private Sub tb_tapantal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_tapantal.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_fravælg_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_fravælg.CheckStateChanged

        Dim kilopris As String
        kilopris = Lb_Kilopris.Text
        If cb_fravælg.CheckState = CheckState.Unchecked Then
        End If
        If cb_fravælg.CheckState = CheckState.Checked Then
            ResetCalculations()
        End If
        Lb_Kilopris.Text = kilopris
        CalculateOrdrestr()
    End Sub
    Private Sub cb_fravælg_1500_3000_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_fravælg_1500_3000.CheckStateChanged
        Dim kilopris As String
        kilopris = Lb_Kilopris.Text

        If cb_fravælg_1500_3000.CheckState = CheckState.Unchecked Then
        End If
        If cb_fravælg_1500_3000.CheckState = CheckState.Checked Then
            ResetCalculations()
        End If
        Lb_Kilopris.Text = kilopris
        CalculateOrdrestr()
    End Sub
    Private Sub cb_fravælg_1250_2500_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_fravælg_1250_2500.CheckStateChanged
        Dim kilopris As String
        kilopris = Lb_Kilopris.Text

        If cb_fravælg_1250_2500.CheckState = CheckState.Unchecked Then
        End If
        If cb_fravælg_1250_2500.CheckState = CheckState.Checked Then
            ResetCalculations()
        End If
        Lb_Kilopris.Text = kilopris

        CalculateOrdrestr()
    End Sub
    Private Sub cb_fravælg_1000_2000_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_fravælg_1000_2000.CheckStateChanged
        Dim kilopris As String
        kilopris = Lb_Kilopris.Text


        If cb_fravælg_1000_2000.CheckState = CheckState.Unchecked Then
        End If
        If cb_fravælg_1000_2000.CheckState = CheckState.Checked Then
            ResetCalculations()
        End If
        Lb_Kilopris.Text = kilopris

        CalculateOrdrestr()
    End Sub

    Private Sub cb_materiale_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_materiale.SelectedIndexChanged
        Dim valgtmateriale As String
        valgtmateriale = cb_materiale.SelectedItem

        If valgtmateriale = "1 Jern, P01 AM, 12.03" Then
            rb_jern.Checked = True
        End If
        If valgtmateriale = "2 ALU, 2S Al99" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "3 ALU, AlMg3, Søvandsbestandig" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "4 Rustfri AISI 304, 1.4301" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "5 Rustfri AISI 316L, 1.4404" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "6 Messing" Then
            rb_rustfri.Checked = False
            rb_jern.Checked = False
        End If
        If valgtmateriale = "7 Elgalv, Fe P01 ZE, Zintec" Then
            rb_jern.Checked = True
        End If
        If valgtmateriale = "8 Aluzink, B500A" Then
            rb_jern.Checked = True
        End If
        If valgtmateriale = "9 ALU, AlMg3, Søvandsbest.m.PVC" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "10 ALU, AlMg1,m.PVC" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "11 Rustfri AISI 304, 1.4301 m.PVC" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "12 Rustfri AISI 304, 1.4301 Slebet" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "13 Rustfri AISI 316L, 1.4404 m.PVC" Then
            rb_rustfri.Checked = True
        End If
        If valgtmateriale = "14 Aluzink, B500A m.PVC" Then
            rb_jern.Checked = True
        End If
        If valgtmateriale = "15 Varmgalvaniseret" Then
            rb_jern.Checked = True
        End If

        CalculateOrdrestr()

    End Sub

    Private Sub lb_antal_opstart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_antal_opstart.Click
        CalculateOrdrestr()
    End Sub
    Private Sub cb_afgrat_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_afgrat.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_steelmaster_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_steelmaster.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_brush_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_brush.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_rette_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_rette.CheckStateChanged
        CalculateOrdrestr()
    End Sub

    Private Sub tb_m2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m2.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m2_5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m2_5.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m3.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m4.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m5.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m6.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m8.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_m10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_m10.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_1.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_2.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_3.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_4.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_presmøtrik_antal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_presmøtrik_antal.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_presstag_antal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_presstag_antal.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_boltesvejs_antal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_boltesvejs_antal.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_numberofsurface_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_numberofsurface.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_overfl_beh1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_beh1.TextChanged
        Getsupplierlist1()
    End Sub
    Private Sub cb_overfl_beh2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_beh2.TextChanged
        Getsupplierlist2()
    End Sub
    Private Sub cb_overfl_beh3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_beh3.TextChanged
        Getsupplierlist3()
    End Sub
    Private Sub cb_overfl_beh4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_beh4.TextChanged
        Getsupplierlist4()
        'CalculateOrdrestr()
    End Sub
    Private Sub cb_overfl_beh5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_beh5.TextChanged
        Getsupplierlist5()
    End Sub
    Private Sub cb_overfl_leverandør1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_leverandør1.TextChanged
        CalculateSurface1()
        CalculateOrdrestr()
    End Sub
    Private Sub cb_overfl_leverandør2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_leverandør2.TextChanged
        CalculateSurface2()
        CalculateOrdrestr()
    End Sub
    Private Sub cb_overfl_leverandør3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_overfl_leverandør3.TextChanged
        CalculateSurface3()
        CalculateOrdrestr()
    End Sub
    Private Sub cb_1side_1_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_1side_1.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_1side_2_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_1side_2.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_1side_3_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_1side_3.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_rettesvejs_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_rettesvejs.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_glasbl_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_glasbl.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub cb_slib_CheckstateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_slib.CheckStateChanged
        CalculateOrdrestr()
    End Sub
    Private Sub tb_samkørsel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_samkørsel.TextChanged
        CalculateOrdrestr()
    End Sub
    Private Sub bu_Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_Reset.Click
        ResetAll()
        'FillCustomerNames()
    End Sub

    Private Sub rb_danmark_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_danmark.CheckedChanged
        If rb_danmark.Checked = True Then
            tb_timesats_mand.Text = 550
            tb_timesats_D.Text = 1000
            tb_timesats_B.Text = 1400
            tb_timesats_C.Text = 2000
        Else
            tb_timesats_mand.Text = 300
            tb_timesats_D.Text = 700
            tb_timesats_B.Text = 1000
            tb_timesats_C.Text = 1200

        End If

        CalculateOrdrestr()
    End Sub
    Private Sub bu_gem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_gem.Click
        Dim ex As New Microsoft.Office.Interop.Excel.Application
        Dim wb As Microsoft.Office.Interop.Excel.Workbook
        'Dim ex As New Excel.Application
        'Dim wb As Excel.Workbook

        Dim sFilename As String
        Dim sKundenavn As String
        Dim sRevision As String
        Dim filnavn As String


        If cb_Tegning.Text = "" Then
            MsgBox("Tegningsnummer mangler")
            Exit Sub
        End If
        If cb_kunde.Text = "" Then
            MsgBox("Kundenavn mangler")
            Exit Sub
        End If
        If tb_tilbudnr.Text = "" Then
            MsgBox("Tilbudsnummer mangler")
            Exit Sub
        End If

        sFilename = cb_Tegning.Text
        sKundenavn = cb_kunde.Text
        sRevision = tb_revision.Text

        lb_filnavn.Text = sKundenavn & "_" & sFilename & "_" & sRevision & ".xls"
        filnavn = lb_filnavn.Text
        filnavn = filnavn.Replace("/", "-")
        lb_filnavn.Text = filnavn

        If lb_operatør_opr.Text = "" Then
            lb_operatør_opr.Text = System.Environment.UserName
            lb_dato_opr.Text = Today
        End If

        lb_operatør.Text = System.Environment.UserName
        lb_dato.Text = Today

        Try
            'Åben template xls filen DK
            'wb = ex.Workbooks.Open("C:\TilbudsFiler\tilbudsdata.xlt")
            wb = ex.Workbooks.Open("W:\Tilbud\TilbudsFiler\tilbudsdata.xlt")
            'Åben template xls filen POL
            'wb = ex.Workbooks.Open("\\Akspol\AKS Gruppen Dokumenter Polen\Tilbud\Metaltilbud\Tilbudsfiler\tilbudsdata.xlt")


        Catch err As Exception
            'Xls filen kan ikke åbnes.
            MsgBox("Excel templaten kan ikke åbnes")
            ex.Quit()
            Exit Sub
        End Try

        'Åben det det første ark i regnearket.
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets.Item(1)
        ' Dim sheet As Excel.Worksheet = wb.Worksheets.Item(1)
        'Læs data ud i kolonne D
        sheet.Cells(2, 4) = Me.lb_filnavn.Text
        sheet.Cells(3, 4) = Me.lb_operatør_opr.Text
        sheet.Cells(4, 4) = Me.lb_dato_opr.Text
        sheet.Cells(5, 4) = Me.cb_kunde.Text
        sheet.Cells(6, 4) = Me.tb_emne.Text
        sheet.Cells(7, 4) = Me.cb_Tegning.Text
        sheet.Cells(8, 4) = Me.tb_revision.Text
        sheet.Cells(9, 4) = Me.cb_materiale.Text
        sheet.Cells(10, 4) = Me.tb_pladetykkelse.Text
        sheet.Cells(11, 4) = Me.tb_sværhed_uk.Text
        sheet.Cells(12, 4) = Me.tb_Kilopris_uk.Text
        sheet.Cells(13, 4) = Me.rb_netto.Checked
        sheet.Cells(14, 4) = Me.rb_brutto.Checked
        sheet.Cells(15, 4) = Me.cb_fravælg.CheckState
        sheet.Cells(16, 4) = Me.tb_bukmax_x.Text
        sheet.Cells(17, 4) = Me.tb_buk1_x.Text
        sheet.Cells(18, 4) = Me.tb_buk2_x.Text
        sheet.Cells(19, 4) = Me.tb_buk3_x.Text
        sheet.Cells(20, 4) = Me.tb_buk4_x.Text
        sheet.Cells(21, 4) = Me.tb_buk5_x.Text
        sheet.Cells(22, 4) = Me.tb_buk6_x.Text
        sheet.Cells(23, 4) = Me.tb_buk7_x.Text
        sheet.Cells(24, 4) = Me.tb_buk8_x.Text
        sheet.Cells(25, 4) = Me.tb_buk9_x.Text
        sheet.Cells(26, 4) = Me.tb_buk10_x.Text
        sheet.Cells(27, 4) = Me.tb_buk11_x.Text
        sheet.Cells(28, 4) = Me.tb_bukmax_y.Text
        sheet.Cells(29, 4) = Me.tb_buk1_y.Text
        sheet.Cells(30, 4) = Me.tb_buk2_y.Text
        sheet.Cells(31, 4) = Me.tb_buk3_y.Text
        sheet.Cells(32, 4) = Me.tb_buk4_y.Text
        sheet.Cells(33, 4) = Me.tb_buk5_y.Text
        sheet.Cells(34, 4) = Me.tb_buk6_y.Text
        sheet.Cells(35, 4) = Me.tb_buk7_y.Text
        sheet.Cells(36, 4) = Me.tb_buk8_y.Text
        sheet.Cells(37, 4) = Me.tb_buk9_y.Text
        sheet.Cells(38, 4) = Me.tb_buk10_y.Text
        sheet.Cells(39, 4) = Me.tb_buk11_y.Text
        sheet.Cells(40, 4) = Me.lb_buk_tid.Text
        sheet.Cells(41, 4) = Me.lb_buk_opst.Text
        sheet.Cells(42, 4) = Me.tb_buk_uk.Text
        sheet.Cells(43, 4) = Me.tb_buk_opst_uk.Text
        sheet.Cells(45, 4) = Me.tb_valsning_uk.Text
        sheet.Cells(47, 4) = Me.rb_D_stans.Checked
        sheet.Cells(48, 4) = Me.tb_toolshift.Text
        sheet.Cells(49, 4) = Me.tb_slag_til_huller.Text
        sheet.Cells(50, 4) = Me.lb_gruppe1_tid.Text
        sheet.Cells(51, 4) = Me.lb_gruppe1_opstart.Text
        sheet.Cells(52, 4) = Me.tb_CNCmin_uk.Text
        sheet.Cells(53, 4) = Me.tb_gruppe1_opstart_uk.Text
        sheet.Cells(54, 4) = Me.rb_C_laser.Checked
        sheet.Cells(55, 4) = Me.tb_hulantal_1C.Text
        sheet.Cells(56, 4) = Me.tb_hulantal_2C.Text
        sheet.Cells(57, 4) = Me.tb_hulantal_3C.Text
        sheet.Cells(58, 4) = Me.tb_cuttinglength_C.Text
        sheet.Cells(59, 4) = Me.lb_laserCNC_tid.Text
        sheet.Cells(60, 4) = Me.lb_laser_opstart.Text
        sheet.Cells(61, 4) = Me.tb_laserCNC_tid_uk.Text
        sheet.Cells(62, 4) = Me.tb_laser_opstart_uk.Text
        sheet.Cells(63, 4) = Me.rb_B_kombi.Checked
        sheet.Cells(64, 4) = Me.tb_toolshift_B.Text
        sheet.Cells(65, 4) = Me.tb_slag_til_huller_B.Text
        sheet.Cells(66, 4) = Me.tb_hulantal_1B.Text
        sheet.Cells(67, 4) = Me.tb_hulantal_2B.Text
        sheet.Cells(68, 4) = Me.tb_hulantal_3B.Text
        sheet.Cells(69, 4) = Me.tb_cuttinglength_B.Text
        sheet.Cells(70, 4) = Me.lb_CombiCNC_tid.Text
        sheet.Cells(71, 4) = Me.lb_Combi_opstart.Text
        sheet.Cells(72, 4) = Me.tb_combiCNC_tid_uk.Text
        sheet.Cells(73, 4) = Me.tb_combi_opstart_uk.Text
        sheet.Cells(74, 4) = Me.rb_klip.Checked
        sheet.Cells(75, 4) = Me.lb_klip_tid.Text
        sheet.Cells(76, 4) = Me.lb_klip_opstart.Text
        sheet.Cells(77, 4) = Me.tb_klip_tid_uk.Text
        sheet.Cells(78, 4) = Me.tb_klip_opstart_uk.Text
        sheet.Cells(79, 4) = Me.lb_CombiCNCstans_tid.Text
        sheet.Cells(80, 4) = Me.cb_afgrat.CheckState
        sheet.Cells(81, 4) = Me.lb_afgrat.Text
        sheet.Cells(82, 4) = Me.tb_afgrat_uk.Text
        sheet.Cells(83, 4) = Me.cb_steelmaster.CheckState
        sheet.Cells(84, 4) = Me.lb_grinding.Text
        sheet.Cells(85, 4) = Me.tb_grinding_uk.Text
        sheet.Cells(86, 4) = Me.cb_brush.CheckState
        sheet.Cells(87, 4) = Me.lb_brush.Text
        sheet.Cells(88, 4) = Me.tb_brush_uk.Text
        sheet.Cells(89, 4) = Me.cb_vibrationsafgr.CheckState
        sheet.Cells(90, 4) = Me.lb_vibration.Text
        sheet.Cells(91, 4) = Me.tb_vibration_uk.Text
        sheet.Cells(92, 4) = Me.cb_rette.CheckState
        sheet.Cells(93, 4) = Me.lb_rette.Text
        sheet.Cells(94, 4) = Me.tb_rette_uk.Text
        sheet.Cells(95, 4) = Me.tb_stans_manuel_uk.Text
        sheet.Cells(96, 4) = Me.tb_presmøtrik_antal.Text
        sheet.Cells(97, 4) = Me.lb_pressnut.Text
        sheet.Cells(98, 4) = Me.tb_pressnut_uk.Text
        sheet.Cells(99, 4) = Me.tb_boltesvejs_antal.Text
        sheet.Cells(100, 4) = Me.lb_boltesvejs.Text
        sheet.Cells(101, 4) = Me.tb_boltesvejs_uk.Text
        sheet.Cells(103, 4) = Me.tb_m2.Text
        sheet.Cells(104, 4) = Me.tb_m2_5.Text
        sheet.Cells(105, 4) = Me.tb_m3.Text
        sheet.Cells(106, 4) = Me.tb_m4.Text
        sheet.Cells(107, 4) = Me.tb_m5.Text
        sheet.Cells(108, 4) = Me.tb_m6.Text
        sheet.Cells(109, 4) = Me.tb_m8.Text
        sheet.Cells(110, 4) = Me.tb_m10.Text
        sheet.Cells(111, 4) = Me.lb_gevind.Text
        sheet.Cells(112, 4) = Me.tb_gevind_uk.Text
        sheet.Cells(114, 4) = Me.tb_1.Text
        sheet.Cells(115, 4) = Me.tb_2.Text
        sheet.Cells(116, 4) = Me.tb_3.Text
        sheet.Cells(117, 4) = Me.tb_4.Text
        sheet.Cells(118, 4) = Me.lb_countersink.Text
        sheet.Cells(119, 4) = Me.tb_countersink_uk.Text
        sheet.Cells(120, 4) = Me.cb_spotweld.CheckState
        sheet.Cells(121, 4) = Me.tb_numberofspotweldseams.Text
        sheet.Cells(122, 4) = Me.tb_numberofspots.Text
        sheet.Cells(123, 4) = Me.lb_spotweld.Text
        sheet.Cells(124, 4) = Me.tb_spotweld_uk.Text
        sheet.Cells(126, 4) = Me.tb_numberofwelds.Text
        sheet.Cells(127, 4) = Me.tb_weldlength.Text
        sheet.Cells(128, 4) = Me.rb_tig.Checked
        sheet.Cells(129, 4) = Me.rb_migmag.Checked
        sheet.Cells(130, 4) = Me.cb_tackweld.CheckState
        sheet.Cells(131, 4) = Me.lb_tackweld.Text
        sheet.Cells(132, 4) = Me.tb_tackweld_uk.Text
        sheet.Cells(133, 4) = Me.cb_weld.CheckState
        sheet.Cells(134, 4) = Me.lb_weld.Text
        sheet.Cells(135, 4) = Me.tb_weld_uk.Text
        sheet.Cells(136, 4) = Me.cb_grind_weld.CheckState
        sheet.Cells(137, 4) = Me.lb_grind_weld.Text
        sheet.Cells(138, 4) = Me.tb_grind_weld_uk.Text
        sheet.Cells(140, 4) = Me.lb_kontor.Text
        sheet.Cells(141, 4) = Me.tb_kontor_uk.Text
        sheet.Cells(142, 4) = Me.lb_kontrol.Text
        sheet.Cells(143, 4) = Me.tb_kontrol_uk.Text
        sheet.Cells(145, 4) = Me.rb_danmark.Checked
        sheet.Cells(146, 4) = Me.rb_polen.Checked
        sheet.Cells(148, 4) = Me.tb_presstag_kr_uk.Text
        sheet.Cells(149, 4) = Me.lb_pressnut_kr.Text
        sheet.Cells(150, 4) = Me.tb_pressnut_kr_uk.Text
        sheet.Cells(151, 4) = Me.lb_svejsestag_kr.Text
        sheet.Cells(152, 4) = Me.tb_svejsestag_kr_uk.Text
        sheet.Cells(153, 4) = Me.lb_tilsatsmatr_kr.Text
        sheet.Cells(154, 4) = Me.tb_tilsatsmatr_kr_uk.Text
        sheet.Cells(155, 4) = Me.lb_gruppetid1.Text
        sheet.Cells(156, 4) = Me.lb_gruppetid2.Text
        sheet.Cells(157, 4) = Me.lb_gruppetid3.Text
        sheet.Cells(158, 4) = Me.lb_gruppetid4.Text
        sheet.Cells(159, 4) = Me.lb_gruppetid5.Text
        sheet.Cells(160, 4) = Me.lb_gruppetid5slib.Text
        sheet.Cells(161, 4) = Me.rtb_bem.Text
        sheet.Cells(162, 4) = Me.cb_overfl_beh1.Text
        sheet.Cells(163, 4) = Me.cb_overfl_leverandør1.Text
        sheet.Cells(164, 4) = Me.cb_1side_1.CheckState
        sheet.Cells(165, 4) = Me.tb_overfl_avance1.Text
        sheet.Cells(166, 4) = Me.tb_overfl_pris100_1.Text
        sheet.Cells(167, 4) = Me.tb_overfl_pris1_uk.Text
        sheet.Cells(168, 4) = Me.cb_overfl_beh2.Text
        sheet.Cells(169, 4) = Me.cb_overfl_leverandør2.Text
        sheet.Cells(170, 4) = Me.cb_1side_2.CheckState
        sheet.Cells(171, 4) = Me.tb_overfl_avance2.Text
        sheet.Cells(172, 4) = Me.tb_overfl_pris100_2.Text
        sheet.Cells(173, 4) = Me.tb_overfl_pris2_uk.Text
        sheet.Cells(174, 4) = Me.cb_overfl_beh3.Text
        sheet.Cells(175, 4) = Me.cb_overfl_leverandør3.Text
        sheet.Cells(176, 4) = Me.cb_1side_3.CheckState
        sheet.Cells(177, 4) = Me.tb_overfl_avance3.Text
        sheet.Cells(178, 4) = Me.tb_overfl_pris100_3.Text
        sheet.Cells(179, 4) = Me.tb_overfl_pris3_uk.Text
        sheet.Cells(180, 4) = Me.cb_overfl_beh4.Text
        sheet.Cells(181, 4) = Me.cb_overfl_leverandør4.Text
        sheet.Cells(182, 4) = Me.tb_overfl_opstart4.Text
        sheet.Cells(183, 4) = Me.tb_overfl_afdæk4.Text
        sheet.Cells(184, 4) = Me.tb_overfl_pris4.Text
        sheet.Cells(185, 4) = Me.tb_overfl_avance4.Text
        sheet.Cells(186, 4) = Me.tb_overfl_pris100_4.Text
        sheet.Cells(188, 4) = Me.tb_antal1.Text
        sheet.Cells(189, 4) = Me.tb_tilbud1.Text
        sheet.Cells(190, 4) = Me.tb_antal2.Text
        sheet.Cells(191, 4) = Me.tb_tilbud2.Text
        sheet.Cells(192, 4) = Me.tb_antal3.Text
        sheet.Cells(193, 4) = Me.tb_tilbud3.Text
        sheet.Cells(194, 4) = Me.tb_antal4.Text
        sheet.Cells(195, 4) = Me.tb_tilbud4.Text
        sheet.Cells(196, 4) = Me.tb_antal5.Text
        sheet.Cells(197, 4) = Me.tb_tilbud5.Text
        sheet.Cells(198, 4) = Me.tb_avance.Text
        sheet.Cells(199, 4) = Me.lb_antal_opstart.Text
        sheet.Cells(200, 4) = Me.tb_antal_opstart_uk.Text
        sheet.Cells(201, 4) = Me.lb_opstart_kr.Text
        sheet.Cells(202, 4) = Me.tb_opstart_kr_uk.Text
        sheet.Cells(203, 4) = Me.lb_antal_program.Text
        sheet.Cells(204, 4) = Me.tb_antal_program_uk.Text
        sheet.Cells(205, 4) = Me.lb_program_kr.Text
        sheet.Cells(206, 4) = Me.tb_Program_kr_uk.Text
        sheet.Cells(207, 4) = Me.tb_opstart_avance.Text
        sheet.Cells(208, 4) = Me.tb_opstart_afgivettilbud.Text
        sheet.Cells(210, 4) = Me.lb_antalplader.Text
        sheet.Cells(211, 4) = Me.lb_pladeformatX.Text
        sheet.Cells(212, 4) = Me.lb_pladeformatY.Text
        sheet.Cells(213, 4) = Me.lb_emner_prplade.Text
        sheet.Cells(214, 4) = Me.lb_spild.Text
        sheet.Cells(215, 4) = Me.Lb_Kilopris.Text
        sheet.Cells(216, 4) = Me.tb_presstag_antal.Text
        sheet.Cells(217, 4) = Me.lb_presstag.Text
        sheet.Cells(218, 4) = Me.tb_presstag_uk.Text
        sheet.Cells(219, 4) = Me.tb_glasbl_uk.Text
        sheet.Cells(221, 4) = Me.cb_fravælg_1500_3000.CheckState
        sheet.Cells(222, 4) = Me.cb_fravælg_1250_2500.CheckState
        sheet.Cells(223, 4) = Me.cb_fravælg_1000_2000.CheckState
        sheet.Cells(225, 4) = Me.tb_rettesvejs_uk.Text
        sheet.Cells(226, 4) = Me.cb_rettesvejs.CheckState
        sheet.Cells(227, 4) = Me.lb_rettesvejs_tid.Text

        sheet.Cells(230, 4) = Me.tb_overfl_tilbudpris1_uk.Text
        sheet.Cells(231, 4) = Me.tb_overfl_tilbudpris2_uk.Text
        sheet.Cells(232, 4) = Me.tb_overfl_tilbudpris3_uk.Text
        sheet.Cells(233, 4) = Me.tb_overfltilbud_antal.Text

        sheet.Cells(235, 4) = Me.cb_overfl_beh5.Text
        sheet.Cells(236, 4) = Me.cb_overfl_leverandør5.Text
        sheet.Cells(237, 4) = Me.tb_overfl_afdæk5.Text
        sheet.Cells(238, 4) = Me.tb_overfl_pris5.Text
        sheet.Cells(239, 4) = Me.tb_overfl_avance5.Text
        sheet.Cells(240, 4) = Me.tb_overfl_opstart5.Text
        sheet.Cells(241, 4) = Me.rb_jern.Checked
        sheet.Cells(242, 4) = Me.rb_rustfri.Checked

        sheet.Cells(243, 4) = Me.tb_hulantal_4C.Text
        sheet.Cells(244, 4) = Me.tb_hulantal_4B.Text

        sheet.Cells(245, 4) = Me.tb_stepantal.Text
        sheet.Cells(246, 4) = Me.tb_stepbuk_uk.Text
        sheet.Cells(247, 4) = Me.tb_tilbudnr.Text

        sheet.Cells(248, 4) = Me.lb_operatør.Text
        sheet.Cells(249, 4) = Me.lb_dato.Text
        sheet.Cells(250, 4) = tb_numberofsurface.Text

        sheet.Cells(251, 4) = Me.cb_glasbl.CheckState
        sheet.Cells(252, 4) = Me.tb_glasbl_uk.Text
        sheet.Cells(253, 4) = Me.cb_slib.CheckState
        sheet.Cells(254, 4) = Me.tb_slib_uk.Text

        sheet.Cells(255, 4) = Me.tb_tapantal.Text
        sheet.Cells(256, 4) = Me.tb_tapsvejs_uk.Text
        sheet.Cells(257, 4) = Me.rb_factor1.Checked
        sheet.Cells(258, 4) = Me.rb_factor2.Checked
        sheet.Cells(259, 4) = Me.rb_factor3.Checked
        sheet.Cells(260, 4) = Me.rb_factor4.Checked
        sheet.Cells(261, 4) = Me.rb_factor5.Checked

        'Gem xls filen med fil navnet i DK
        'wb.SaveAs("C:\TilbudsFiler\" & filnavn & " ")
        wb.SaveAs("W:\Tilbud\TilbudsFiler\" & filnavn & " ")

        'wb.SaveCopyAs("\\Win2000server\AKS Metal\Tilbud\Tilbudsfiler\" & filnavn & " ")


        'Gem xls filen med fil navnet i Polen
        'wb.SaveAs("\\Akspol\AKS Gruppen Dokumenter Polen\Tilbud\Metaltilbud\Tilbudsfiler\" & filnavn & " ")

        'Luk regnearket
        wb.Close()
        'Afslut Excel
        ex.Quit()
    End Sub
    Private Sub bu_hent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_hent.Click
        Dim FileDialog As New OpenFileDialog
        Dim sFilename As String
        Dim wb As Microsoft.Office.Interop.Excel.Workbook

        'Dim materiale As String

        'Dim wb As Excel.Workbook

        'Sæt fil dialog properties
        FileDialog.DefaultExt = "xls"
        FileDialog.Filter = "Excel Filer (*.xls)|*.xls"
        'DK
        'FileDialog.InitialDirectory = "C:\TilbudsFiler"
        FileDialog.InitialDirectory = "W:\Tilbud\TilbudsFiler"
        'Polen
        'FileDialog.InitialDirectory = "\\Akspol\AKS Gruppen Dokumenter Polen\Tilbud\Metaltilbud\Tilbudsfiler\"
        'Vis åben fil dialogen
        If FileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Hent filnavnet fra dialog boksen
            sFilename = FileDialog.FileName
        Else
            Exit Sub
        End If

        Dim ex As New Microsoft.Office.Interop.Excel.Application
        ' Dim ex As New Excel.Application

        Try
            'Åben xls filen
            wb = ex.Workbooks.Open(sFilename)

            'Åben det det første ark i regnearket.
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets.Item(1)
            ' Dim sheet As Excel.Worksheet = wb.Worksheets.Item(1)

            ResetAll()

            'Indlæs kolonne B fra regnearket til Text boks.


            Me.lb_filnavn.Text = sheet.Cells(2, 4).value
            Me.lb_operatør_opr.Text = sheet.Cells(3, 4).value
            Me.lb_dato_opr.Text = sheet.Cells(4, 4).value
            Me.cb_kunde.Text = sheet.Cells(5, 4).value
            Me.tb_emne.Text = sheet.Cells(6, 4).value
            Me.cb_Tegning.Text = sheet.Cells(7, 4).value
            Me.tb_revision.Text = sheet.Cells(8, 4).value
            Me.cb_materiale.Text = sheet.Cells(9, 4).value
            Me.tb_sværhed_uk.Text = sheet.Cells(11, 4).value
            Me.tb_Kilopris_uk.Text = sheet.Cells(12, 4).value
            Me.rb_netto.Checked = sheet.Cells(13, 4).value
            Me.rb_brutto.Checked = sheet.Cells(14, 4).value
            Me.cb_fravælg.CheckState = sheet.Cells(15, 4).value
            Me.tb_bukmax_x.Text = sheet.Cells(16, 4).value
            Me.tb_buk1_x.Text = sheet.Cells(17, 4).value
            Me.tb_buk2_x.Text = sheet.Cells(18, 4).value
            Me.tb_buk3_x.Text = sheet.Cells(19, 4).value
            Me.tb_buk4_x.Text = sheet.Cells(20, 4).value
            Me.tb_buk5_x.Text = sheet.Cells(21, 4).value
            Me.tb_buk6_x.Text = sheet.Cells(22, 4).value
            Me.tb_buk7_x.Text = sheet.Cells(23, 4).value
            Me.tb_buk8_x.Text = sheet.Cells(24, 4).value
            Me.tb_buk9_x.Text = sheet.Cells(25, 4).value
            Me.tb_buk10_x.Text = sheet.Cells(26, 4).value
            Me.tb_buk11_x.Text = sheet.Cells(27, 4).value
            Me.tb_bukmax_y.Text = sheet.Cells(28, 4).value
            Me.tb_buk1_y.Text = sheet.Cells(29, 4).value
            Me.tb_buk2_y.Text = sheet.Cells(30, 4).value
            Me.tb_buk3_y.Text = sheet.Cells(31, 4).value
            Me.tb_buk4_y.Text = sheet.Cells(32, 4).value
            Me.tb_buk5_y.Text = sheet.Cells(33, 4).value
            Me.tb_buk6_y.Text = sheet.Cells(34, 4).value
            Me.tb_buk7_y.Text = sheet.Cells(35, 4).value
            Me.tb_buk8_y.Text = sheet.Cells(36, 4).value
            Me.tb_buk9_y.Text = sheet.Cells(37, 4).value
            Me.tb_buk10_y.Text = sheet.Cells(38, 4).value
            Me.tb_buk11_y.Text = sheet.Cells(39, 4).value
            Me.tb_buk_uk.Text = sheet.Cells(42, 4).value
            Me.tb_buk_opst_uk.Text = sheet.Cells(43, 4).value
            Me.rb_D_stans.Checked = sheet.Cells(47, 4).value
            Me.tb_toolshift.Text = sheet.Cells(48, 4).value
            Me.tb_slag_til_huller.Text = sheet.Cells(49, 4).value
            Me.tb_CNCmin_uk.Text = sheet.Cells(52, 4).value
            Me.tb_gruppe1_opstart_uk.Text = sheet.Cells(53, 4).value
            Me.rb_C_laser.Checked = sheet.Cells(54, 4).value
            Me.tb_hulantal_1C.Text = sheet.Cells(55, 4).value
            Me.tb_hulantal_2C.Text = sheet.Cells(56, 4).value
            Me.tb_hulantal_3C.Text = sheet.Cells(57, 4).value
            Me.tb_hulantal_4C.Text = sheet.Cells(243, 4).value
            Me.tb_cuttinglength_C.Text = sheet.Cells(58, 4).value
            Me.tb_laserCNC_tid_uk.Text = sheet.Cells(61, 4).value
            Me.tb_laser_opstart_uk.Text = sheet.Cells(62, 4).value
            Me.rb_B_kombi.Checked = sheet.Cells(63, 4).value
            Me.tb_toolshift_B.Text = sheet.Cells(64, 4).value
            Me.tb_slag_til_huller_B.Text = sheet.Cells(65, 4).value
            Me.tb_hulantal_1B.Text = sheet.Cells(66, 4).value
            Me.tb_hulantal_2B.Text = sheet.Cells(67, 4).value
            Me.tb_hulantal_3B.Text = sheet.Cells(68, 4).value
            Me.tb_hulantal_4B.Text = sheet.Cells(244, 4).value
            Me.tb_cuttinglength_B.Text = sheet.Cells(69, 4).value
            Me.tb_combiCNC_tid_uk.Text = sheet.Cells(72, 4).value
            Me.tb_combi_opstart_uk.Text = sheet.Cells(73, 4).value
            Me.rb_klip.Checked = sheet.Cells(74, 4).value
            Me.tb_klip_tid_uk.Text = sheet.Cells(77, 4).value
            Me.tb_klip_opstart_uk.Text = sheet.Cells(78, 4).value
            Me.tb_CombiCNCstans_tid_uk.Text = sheet.Cells(79, 4).value
            Me.cb_afgrat.CheckState = sheet.Cells(80, 4).value
            Me.tb_afgrat_uk.Text = sheet.Cells(82, 4).value
            Me.cb_steelmaster.CheckState = sheet.Cells(83, 4).value
            Me.tb_grinding_uk.Text = sheet.Cells(85, 4).value
            Me.cb_brush.CheckState = sheet.Cells(86, 4).value
            Me.tb_brush_uk.Text = sheet.Cells(88, 4).value
            Me.cb_vibrationsafgr.CheckState = sheet.Cells(89, 4).value
            Me.tb_vibration_uk.Text = sheet.Cells(91, 4).value
            Me.cb_rette.CheckState = sheet.Cells(92, 4).value
            Me.tb_rette_uk.Text = sheet.Cells(94, 4).value
            Me.tb_stans_manuel_uk.Text = sheet.Cells(95, 4).value
            Me.tb_presmøtrik_antal.Text = sheet.Cells(96, 4).value
            Me.tb_pressnut_uk.Text = sheet.Cells(98, 4).value
            Me.tb_presstag_antal.Text = sheet.Cells(216, 4).value
            Me.tb_presstag_uk.Text = sheet.Cells(218, 4).value
            Me.tb_boltesvejs_antal.Text = sheet.Cells(99, 4).value
            Me.tb_boltesvejs_uk.Text = sheet.Cells(101, 4).value
            Me.tb_m2.Text = sheet.Cells(103, 4).value
            Me.tb_m2_5.Text = sheet.Cells(104, 4).value
            Me.tb_m3.Text = sheet.Cells(105, 4).value
            Me.tb_m4.Text = sheet.Cells(106, 4).value
            Me.tb_m5.Text = sheet.Cells(107, 4).value
            Me.tb_m6.Text = sheet.Cells(108, 4).value
            Me.tb_m8.Text = sheet.Cells(109, 4).value
            Me.tb_m10.Text = sheet.Cells(110, 4).value
            Me.tb_gevind_uk.Text = sheet.Cells(112, 4).value
            Me.tb_1.Text = sheet.Cells(114, 4).value
            Me.tb_2.Text = sheet.Cells(115, 4).value
            Me.tb_3.Text = sheet.Cells(116, 4).value
            Me.tb_4.Text = sheet.Cells(117, 4).value
            Me.tb_countersink_uk.Text = sheet.Cells(119, 4).value
            Me.cb_spotweld.CheckState = sheet.Cells(120, 4).value
            Me.tb_numberofspotweldseams.Text = sheet.Cells(121, 4).value
            Me.tb_numberofspots.Text = sheet.Cells(122, 4).value
            Me.tb_spotweld_uk.Text = sheet.Cells(124, 4).value
            Me.tb_numberofwelds.Text = sheet.Cells(126, 4).value
            Me.tb_weldlength.Text = sheet.Cells(127, 4).value
            Me.rb_tig.Checked = sheet.Cells(128, 4).value
            Me.cb_tackweld.CheckState = sheet.Cells(130, 4).value
            Me.tb_tackweld_uk.Text = sheet.Cells(132, 4).value
            Me.cb_weld.CheckState = sheet.Cells(133, 4).value
            Me.tb_weld_uk.Text = sheet.Cells(135, 4).value
            Me.cb_grind_weld.CheckState = sheet.Cells(136, 4).value
            Me.tb_grind_weld_uk.Text = sheet.Cells(138, 4).value
            Me.tb_kontor_uk.Text = sheet.Cells(141, 4).value
            Me.tb_kontrol_uk.Text = sheet.Cells(143, 4).value
            Me.rb_danmark.Checked = sheet.Cells(145, 4).value
            Me.rb_polen.Checked = sheet.Cells(146, 4).value
            Me.tb_presstag_kr_uk.Text = sheet.Cells(148, 4).value
            Me.tb_pressnut_kr_uk.Text = sheet.Cells(150, 4).value
            Me.lb_svejsestag_kr.Text = sheet.Cells(151, 4).value
            Me.tb_svejsestag_kr_uk.Text = sheet.Cells(152, 4).value
            Me.tb_tilsatsmatr_kr_uk.Text = sheet.Cells(154, 4).value
            Me.rtb_bem.Text = sheet.Cells(161, 4).value
            Me.cb_overfl_beh1.Text = sheet.Cells(162, 4).value
            Me.cb_overfl_leverandør1.Text = sheet.Cells(163, 4).value
            Me.cb_1side_1.CheckState = sheet.Cells(164, 4).value
            Me.tb_overfl_avance1.Text = sheet.Cells(165, 4).value
            Me.tb_overfl_pris1_uk.Text = sheet.Cells(167, 4).value
            Me.cb_overfl_beh2.Text = sheet.Cells(168, 4).value
            Me.cb_overfl_leverandør2.Text = sheet.Cells(169, 4).value
            Me.cb_1side_2.CheckState = sheet.Cells(170, 4).value
            Me.tb_overfl_avance2.Text = sheet.Cells(171, 4).value
            Me.tb_overfl_pris2_uk.Text = sheet.Cells(173, 4).value
            Me.cb_overfl_beh3.Text = sheet.Cells(174, 4).value
            Me.cb_overfl_leverandør3.Text = sheet.Cells(175, 4).value
            Me.cb_1side_3.CheckState = sheet.Cells(176, 4).value
            Me.tb_overfl_avance3.Text = sheet.Cells(177, 4).value
            Me.tb_overfl_pris3_uk.Text = sheet.Cells(179, 4).value
            Me.cb_overfl_beh4.Text = sheet.Cells(180, 4).value
            Me.cb_overfl_leverandør4.Text = sheet.Cells(181, 4).value
            Me.tb_overfl_opstart4.Text = sheet.Cells(182, 4).value
            Me.tb_overfl_afdæk4.Text = sheet.Cells(183, 4).value
            Me.tb_overfl_pris4.Text = sheet.Cells(184, 4).value
            Me.tb_overfl_avance4.Text = sheet.Cells(185, 4).value
            Me.tb_antal1.Text = "100"

            'Me.tb_antal1.Text = sheet.Cells(188, 4).value
            Me.tb_tilbud1.Text = sheet.Cells(189, 4).value
            'Me.tb_antal2.Text = sheet.Cells(190, 4).value
            Me.tb_tilbud2.Text = sheet.Cells(191, 4).value
            'Me.tb_antal3.Text = sheet.Cells(192, 4).value
            Me.tb_tilbud3.Text = sheet.Cells(193, 4).value
            'Me.tb_antal4.Text = sheet.Cells(194, 4).value
            Me.tb_tilbud4.Text = sheet.Cells(195, 4).value
            'Me.tb_antal5.Text = sheet.Cells(196, 4).value
            Me.tb_tilbud5.Text = sheet.Cells(197, 4).value
            Me.tb_avance.Text = sheet.Cells(198, 4).value
            Me.tb_antal_opstart_uk.Text = sheet.Cells(200, 4).value
            Me.tb_opstart_kr_uk.Text = sheet.Cells(202, 4).value
            Me.tb_antal_program_uk.Text = sheet.Cells(204, 4).value
            Me.tb_Program_kr_uk.Text = sheet.Cells(206, 4).value
            Me.tb_opstart_avance.Text = sheet.Cells(207, 4).value
            Me.tb_opstart_afgivettilbud.Text = sheet.Cells(208, 4).value
            Me.tb_pladetykkelse.Text = sheet.Cells(10, 4).value
            Me.tb_glasbl_uk.Text = sheet.Cells(219, 4).value
            Me.cb_fravælg_1500_3000.CheckState = sheet.Cells(221, 4).value
            Me.cb_fravælg_1250_2500.CheckState = sheet.Cells(222, 4).value
            Me.cb_fravælg_1000_2000.CheckState = sheet.Cells(223, 4).value
            Me.cb_rettesvejs.CheckState = sheet.Cells(226, 4).value
            Me.lb_rettesvejs_tid.Text = sheet.Cells(227, 4).value

            'CalculateOrdrestr2()
            Me.tb_antal1.Text = sheet.Cells(188, 4).value
            Me.tb_antal2.Text = sheet.Cells(190, 4).value
            If tb_antal2.Text = "" Then
                tb_antal2.Text = 0
                tb_antal2.Text = ""
            End If
            Me.tb_antal3.Text = sheet.Cells(192, 4).value
            Me.tb_antal4.Text = sheet.Cells(194, 4).value
            Me.tb_antal5.Text = sheet.Cells(196, 4).value

            Me.tb_buk_uk.Text = sheet.Cells(42, 4).value
            Me.tb_buk_opst_uk.Text = sheet.Cells(43, 4).value
            Me.tb_CNCmin_uk.Text = sheet.Cells(52, 4).value
            Me.tb_gruppe1_opstart_uk.Text = sheet.Cells(53, 4).value
            Me.tb_laserCNC_tid_uk.Text = sheet.Cells(61, 4).value
            Me.tb_laser_opstart_uk.Text = sheet.Cells(62, 4).value
            Me.tb_combiCNC_tid_uk.Text = sheet.Cells(72, 4).value
            Me.tb_combi_opstart_uk.Text = sheet.Cells(73, 4).value
            Me.tb_klip_tid_uk.Text = sheet.Cells(77, 4).value
            Me.tb_klip_opstart_uk.Text = sheet.Cells(78, 4).value
            Me.tb_rettesvejs_uk.Text = sheet.Cells(225, 4).value
            Me.tb_valsning_uk.Text = sheet.Cells(45, 4).value
            Me.tb_glasbl_uk.Text = sheet.Cells(219, 4).value

            Me.tb_afgrat_uk.Text = sheet.Cells(82, 4).value
            Me.tb_grinding_uk.Text = sheet.Cells(85, 4).value
            Me.tb_brush_uk.Text = sheet.Cells(88, 4).value
            Me.tb_vibration_uk.Text = sheet.Cells(91, 4).value
            Me.tb_rette_uk.Text = sheet.Cells(94, 4).value
            Me.tb_stans_manuel_uk.Text = sheet.Cells(95, 4).value
            Me.tb_pressnut_uk.Text = sheet.Cells(98, 4).value
            Me.tb_presstag_uk.Text = sheet.Cells(218, 4).value
            Me.tb_boltesvejs_uk.Text = sheet.Cells(101, 4).value
            Me.tb_gevind_uk.Text = sheet.Cells(112, 4).value
            Me.tb_countersink_uk.Text = sheet.Cells(119, 4).value
            Me.tb_spotweld_uk.Text = sheet.Cells(124, 4).value
            Me.tb_tackweld_uk.Text = sheet.Cells(132, 4).value
            Me.tb_weld_uk.Text = sheet.Cells(135, 4).value
            Me.tb_grind_weld_uk.Text = sheet.Cells(138, 4).value
            Me.tb_kontor_uk.Text = sheet.Cells(141, 4).value
            Me.tb_kontrol_uk.Text = sheet.Cells(143, 4).value

            Me.Lb_Kilopris.Text = sheet.Cells(215, 4).value

            Me.tb_overfl_tilbudpris1_uk.Text = sheet.Cells(230, 4).value
            Me.tb_overfl_tilbudpris2_uk.Text = sheet.Cells(231, 4).value
            Me.tb_overfl_tilbudpris3_uk.Text = sheet.Cells(232, 4).value
            Me.tb_overfltilbud_antal.Text = sheet.Cells(233, 4).value

            Me.cb_overfl_beh5.Text = sheet.Cells(235, 4).value
            Me.cb_overfl_leverandør5.Text = sheet.Cells(236, 4).value
            Me.tb_overfl_afdæk5.Text = sheet.Cells(237, 4).value
            Me.tb_overfl_pris5.Text = sheet.Cells(238, 4).value
            Me.tb_overfl_avance5.Text = sheet.Cells(239, 4).value
            Me.tb_overfl_opstart5.Text = sheet.Cells(240, 4).value
            Me.rb_jern.Checked = sheet.Cells(241, 4).value
            Me.rb_rustfri.Checked = sheet.Cells(242, 4).value

            Me.tb_stepantal.Text = sheet.Cells(245, 4).value
            Me.tb_stepbuk_uk.Text = sheet.Cells(246, 4).value
            Me.tb_tilbudnr.Text = sheet.Cells(247, 4).value

            Me.lb_operatør.Text = sheet.Cells(248, 4).value
            Me.lb_dato.Text = sheet.Cells(249, 4).value
            tb_numberofsurface.Text = sheet.Cells(250, 4).value
            If tb_numberofsurface.Text = "" Then
                tb_numberofsurface.Text = 1
            End If
            Me.cb_glasbl.CheckState = sheet.Cells(251, 4).value
            Me.tb_glasbl_uk.Text = sheet.Cells(252, 4).value
            Me.cb_slib.CheckState = sheet.Cells(253, 4).value
            Me.tb_slib_uk.Text = sheet.Cells(254, 4).value
            'materiale = cb_materiale.Text
            Me.tb_tapantal.Text = sheet.Cells(255, 4).value
            Me.tb_tapsvejs_uk.Text = sheet.Cells(256, 4).value
            Me.rb_factor1.Checked = sheet.Cells(257, 4).value
            Me.rb_factor2.Checked = sheet.Cells(258, 4).value
            Me.rb_factor3.Checked = sheet.Cells(259, 4).value
            Me.rb_factor4.Checked = sheet.Cells(260, 4).value
            Me.rb_factor5.Checked = sheet.Cells(261, 4).value

        Catch err As Exception
            'Der er opstået en fejl. Vis fejl beskeden
            MsgBox(err.Message)

        Finally
            'Luk regnearket
            'wb.Close()
            'Afslut Excel
            ex.Quit()

        End Try


        CalculateOrdrestr()
        'cb_materiale.Text = materiale
        'CalculateAll()

    End Sub


    Private Sub bu_udskriv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bu_udskriv.Click


        'Print(Me.Metaltilbud1.metal_tilbub)

    End Sub
    Public Sub Resetunderkend()

        tb_CNCmin_uk.Text = ""
        tb_laserCNC_tid_uk.Text = ""
        tb_combiCNC_tid_uk.Text = ""
        tb_CombiCNCstans_tid_uk.Text = ""
        tb_klip_tid_uk.Text = ""
        tb_klip_tid_uk.Text = ""
        tb_buk_uk.Text = ""
        tb_afgrat_uk.Text = ""
        tb_grinding_uk.Text = ""
        tb_brush_uk.Text = ""
        tb_vibration_uk.Text = ""
        tb_rette_uk.Text = ""
        tb_stans_manuel_uk.Text = ""
        tb_countersink_uk.Text = ""
        tb_gevind_uk.Text = ""
        tb_pressnut_uk.Text = ""
        tb_presstag_uk.Text = ""
        tb_boltesvejs_uk.Text = ""
        tb_spotweld_uk.Text = ""
        tb_tackweld_uk.Text = ""
        tb_weld_uk.Text = ""
        tb_grind_weld_uk.Text = ""
        tb_glasbl_uk.Text = ""
        tb_slib_uk.Text = ""
        tb_valsning_uk.Text = ""
        tb_rettesvejs_uk.Text = ""
    End Sub

    Private Sub ResetCalculations()
        lb_mand1.Text = ""
        lb_cnc1.Text = ""
        lb_mand1_tid.Text = ""
        lb_CNC1_tid.Text = ""
        lb_timer1.Text = ""
        lb_indkøb1.Text = ""
        lb_råvarerstk1.Text = ""
        lb_råvarer1.Text = ""
        lb_samlet1.Text = ""
        lb_salg1.Text = ""
        lb_mand2.Text = ""
        lb_cnc2.Text = ""
        lb_mand2_tid.Text = ""
        lb_CNC2_tid.Text = ""
        lb_timer2.Text = ""
        lb_indkøb2.Text = ""
        lb_råvarerstk2.Text = ""
        lb_råvarer2.Text = ""
        lb_samlet2.Text = ""
        lb_salg2.Text = ""
        lb_mand3.Text = ""
        lb_cnc3.Text = ""
        lb_mand3_tid.Text = ""
        lb_CNC3_tid.Text = ""
        lb_timer3.Text = ""
        lb_indkøb3.Text = ""
        lb_råvarerstk3.Text = ""
        lb_råvarer3.Text = ""
        lb_samlet3.Text = ""
        lb_salg3.Text = ""
        lb_mand4.Text = ""
        lb_cnc4.Text = ""
        lb_mand4_tid.Text = ""
        lb_CNC4_tid.Text = ""
        lb_timer4.Text = ""
        lb_indkøb4.Text = ""
        lb_råvarerstk4.Text = ""
        lb_råvarer4.Text = ""
        lb_samlet4.Text = ""
        lb_salg4.Text = ""
        lb_mand5.Text = ""
        lb_cnc5.Text = ""
        lb_mand5_tid.Text = ""
        lb_CNC5_tid.Text = ""
        lb_timer5.Text = ""
        lb_indkøb5.Text = ""
        lb_råvarerstk5.Text = ""
        lb_råvarer5.Text = ""
        lb_samlet5.Text = ""
        lb_salg5.Text = ""
        lb_gruppe1_tid.Text = ""
        lb_gruppe1_opstart.Text = ""
        tb_CNCmin_uk.Text = ""
        lb_CombiCNCstans_tid.Text = ""
        lb_combi_ialt.Text = ""
        lb_stans_ialt.Text = ""
        lb_buk_ialt.Text = ""
        lb_klip_ialt.Text = ""
        lb_laser_ialt.Text = ""

        tb_gruppe1_opstart_uk.Text = ""
        tb_laserCNC_tid_uk.Text = ""
        tb_laser_opstart_uk.Text = ""
        tb_combiCNC_tid_uk.Text = ""
        tb_CombiCNCstans_tid_uk.Text = ""
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
        tb_presstag_kr_uk.Text = ""
        tb_pressnut_kr_uk.Text = ""
        tb_svejsestag_kr_uk.Text = ""
        tb_tilsatsmatr_kr_uk.Text = ""
        tb_glasbl_uk.Text = ""
        tb_slib_uk.Text = ""
        tb_tilbudnr.Text = ""
        lb_laserCNC_tid.Text = ""
        lb_laser_opstart.Text = ""
        lb_nitrogen.Text = ""
        lb_pladeforbrug.Text = ""
        Lb_Kilopris.Text = ""
        tb_samkørsel.Text = ""
        tb_tapantal.Text = ""
        lb_tapsvejs_tid.Text = ""
        tb_tapsvejs_uk.Text = ""
        tb_svejstid_ialt.Text = ""
        rb_factor1.Checked = True
    End Sub
    Private Sub ResetAll()
        cb_fravælg.CheckState = CheckState.Checked
        cb_fravælg_1500_3000.CheckState = CheckState.Unchecked
        cb_fravælg_1250_2500.CheckState = CheckState.Unchecked
        cb_fravælg_1000_2000.CheckState = CheckState.Unchecked
        rb_C_laser.Checked = True
        rb_netto.Checked = True
        lb_mand1.Text = ""
        lb_cnc1.Text = ""
        lb_mand1_tid.Text = ""
        lb_CNC1_tid.Text = ""
        lb_timer1.Text = ""
        lb_indkøb1.Text = ""
        lb_råvarerstk1.Text = ""
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
        lb_råvarerstk2.Text = ""
        lb_samlet2.Text = ""
        lb_salg2.Text = ""
        lb_mand3.Text = ""
        lb_cnc3.Text = ""
        lb_mand3_tid.Text = ""
        lb_CNC3_tid.Text = ""
        lb_timer3.Text = ""
        lb_indkøb3.Text = ""
        lb_råvarer3.Text = ""
        lb_råvarerstk3.Text = ""
        lb_samlet3.Text = ""
        lb_salg3.Text = ""
        lb_mand4.Text = ""
        lb_cnc4.Text = ""
        lb_mand4_tid.Text = ""
        lb_CNC4_tid.Text = ""
        lb_timer4.Text = ""
        lb_indkøb4.Text = ""
        lb_råvarer4.Text = ""
        lb_råvarerstk4.Text = ""
        lb_samlet4.Text = ""
        lb_salg4.Text = ""
        lb_mand5.Text = ""
        lb_cnc5.Text = ""
        lb_mand5_tid.Text = ""
        lb_CNC5_tid.Text = ""
        lb_timer5.Text = ""
        lb_indkøb5.Text = ""
        lb_råvarer5.Text = ""
        lb_råvarerstk5.Text = ""
        lb_samlet5.Text = ""
        lb_salg5.Text = ""
        tb_pladetykkelse.Text = ""
        tb_antal1.Text = "100"
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
        tb_stepantal.Text = ""
        lb_udfold_x.Text = ""
        lb_udfold_y.Text = ""
        lb_buk_tid.Text = ""
        lb_buk_opst.Text = ""
        tb_buk_uk.Text = ""
        tb_buk_opst_uk.Text = ""
        lb_buk_ialt.Text = ""
        lb_buk_ialtuk.Text = ""
        tb_stepbuk_uk.Text = ""
        lb_stepbuk.Text = ""
        lb_nettoareal.Text = ""
        tb_toolshift.Text = ""
        tb_slag_til_huller.Text = ""
        lb_gruppe1_tid.Text = ""
        lb_gruppe1_opstart.Text = ""
        tb_CNCmin_uk.Text = ""
        tb_gruppe1_opstart_uk.Text = ""
        lb_stans_ialt.Text = ""
        lb_stans_ialtuk.Text = ""
        tb_hulantal_1C.Text = ""
        tb_hulantal_2C.Text = ""
        tb_hulantal_3C.Text = ""
        tb_hulantal_4C.Text = ""
        tb_cuttinglength_C.Text = ""
        tb_laserCNC_tid_uk.Text = ""
        tb_laser_opstart_uk.Text = ""
        lb_laser_ialt.Text = ""
        lb_laser_ialtuk.Text = ""
        tb_toolshift_B.Text = ""
        tb_slag_til_huller_B.Text = ""
        tb_hulantal_1B.Text = ""
        tb_hulantal_2B.Text = ""
        tb_hulantal_3B.Text = ""
        tb_hulantal_4B.Text = ""
        tb_cuttinglength_B.Text = ""
        tb_combiCNC_tid_uk.Text = ""
        tb_CombiCNCstans_tid_uk.Text = ""
        tb_combi_opstart_uk.Text = ""
        lb_combi_ialt.Text = ""
        lb_combi_ialtuk.Text = ""
        tb_klip_tid_uk.Text = ""
        tb_klip_opstart_uk.Text = ""
        lb_klip_ialt.Text = ""
        lb_klip_ialtuk.Text = ""
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
        lb_presstag.Text = ""
        lb_boltesvejs.Text = ""
        lb_undersænk_antal.Text = ""
        lb_gevind_antal.Text = ""
        tb_presmøtrik_antal.Text = ""
        tb_presstag_antal.Text = ""
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
        tb_presstag_uk.Text = ""
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
        tb_numberofspots.Text = ""
        tb_numberofspotweldseams.Text = ""
        lb_spotweld.Text = ""
        tb_spotweld_uk.Text = ""
        tb_numberofwelds.Text = ""
        tb_weldlength.Text = ""
        cb_tackweld.CheckState = CheckState.Unchecked
        cb_weld.CheckState = CheckState.Unchecked
        cb_spotweld.CheckState = CheckState.Unchecked
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
        lb_filnavn.Text = ""
        lb_operatør.Text = ""
        lb_operatør_opr.Text = ""
        lb_dato_opr.Text = ""
        lb_dato.Text = ""
        cb_overfl_leverandør1.Text = ""
        cb_overfl_leverandør2.Text = ""
        cb_overfl_leverandør3.Text = ""
        cb_overfl_leverandør4.Text = ""
        cb_overfl_leverandør5.Text = ""
        cb_overfl_beh1.Text = ""
        cb_overfl_beh2.Text = ""
        cb_overfl_beh3.Text = ""
        cb_overfl_beh4.Text = ""
        cb_overfl_beh5.Text = ""
        tb_overfl_opstart1.Text = ""
        tb_overfl_opstart2.Text = ""
        tb_overfl_opstart3.Text = ""
        tb_overfl_opstart4.Text = ""
        tb_overfl_opstart5.Text = ""
        tb_overfl_afdæk1.Text = ""
        tb_overfl_afdæk2.Text = ""
        tb_overfl_afdæk3.Text = ""
        tb_overfl_afdæk4.Text = ""
        tb_overfl_afdæk5.Text = ""
        tb_overfl_pris1.Text = ""
        tb_overfl_pris2.Text = ""
        tb_overfl_pris3.Text = ""
        tb_overfl_pris4.Text = ""
        tb_overfl_pris5.Text = ""
        tb_overfl_pris1_uk.Text = ""
        tb_overfl_pris2_uk.Text = ""
        tb_overfl_pris3_uk.Text = ""
        tb_overfl_avance1.Text = 15
        tb_overfl_avance2.Text = 15
        tb_overfl_avance3.Text = 15
        tb_overfl_avance4.Text = 15
        tb_overfl_avance5.Text = 15
        tb_overfl_pris100_1.Text = ""
        tb_overfl_pris100_2.Text = ""
        tb_overfl_pris100_3.Text = ""
        tb_overfl_pris100_4.Text = ""
        tb_overfl_pris100_5.Text = ""
        tb_overfl_tilbudpris1_uk.Text = ""
        tb_overfl_tilbudpris2_uk.Text = ""
        tb_overfl_tilbudpris3_uk.Text = ""
        tb_overfltilbud_antal.Text = ""
        cb_1side_1.CheckState = CheckState.Unchecked
        cb_1side_2.CheckState = CheckState.Unchecked
        cb_1side_3.CheckState = CheckState.Unchecked
        lb_pladeformatX.Text = ""
        lb_pladeformatY.Text = ""
        lb_modulstr.Text = ""
        lb_faktor.Text = ""
        lb_sværhed.Text = ""
        tb_sværhed_uk.Text = ""
        tb_Kilopris_uk.Text = ""
        lb_antalplader.Text = ""
        lb_emner_prplade.Text = ""
        lb_spild.Text = ""
        lb_emnevægt.Text = ""
        lb_spildtype.Text = ""
        lb_spildnetto.Text = ""
        lb_gruppetid1.Text = ""
        lb_gruppetid2.Text = ""
        lb_gruppetid3.Text = ""
        lb_gruppetid4.Text = ""
        lb_gruppetid5.Text = ""
        lb_gruppetid5slib.Text = ""
        lb_pressnut_kr.Text = ""
        lb_presstag_kr.Text = ""
        lb_svejsestag_kr.Text = ""
        tb_presstag_kr_uk.Text = ""
        tb_pressnut_kr_uk.Text = ""
        tb_svejsestag_kr_uk.Text = ""
        tb_tilsatsmatr_kr_uk.Text = ""
        rtb_bem.Text = ""
        tb_valsning_uk.Text = ""
        tb_rettesvejs_uk.Text = ""
        cb_rettesvejs.CheckState = CheckState.Unchecked
        lb_rettesvejs_tid.Text = ""
        tb_glasbl_uk.Text = ""
        tb_slib_uk.Text = ""
        lb_glasbl.Text = ""
        lb_slib.Text = ""
        tb_tilbudnr.Text = ""
        lb_laserCNC_tid.Text = ""
        lb_laser_opstart.Text = ""
        lb_nitrogen.Text = ""
        lb_pladeforbrug.Text = ""
        Lb_Kilopris.Text = ""
        tb_samkørsel.Text = ""
        tb_numberofsurface.Text = 1
        cb_glasbl.CheckState = CheckState.Unchecked
        cb_slib.CheckState = CheckState.Unchecked
        rb_factor1.Checked = True
        rb_danmark.Checked = True
        tb_tapantal.Text = ""
        lb_tapsvejs_tid.Text = ""
        tb_tapsvejs_uk.Text = ""
        tb_svejstid_ialt.Text = ""
        'rb_polen.Checked = True

        rb_jern.Checked = True

        If rb_danmark.Checked = True Then
            tb_timesats_mand.Text = 550
            tb_timesats_D.Text = 1000
            tb_timesats_B.Text = 1400
            tb_timesats_C.Text = 2000
        Else
            tb_timesats_mand.Text = 300
            tb_timesats_D.Text = 700
            tb_timesats_B.Text = 1000
            tb_timesats_C.Text = 1200

        End If

    End Sub
    Private Sub ResetUK()
        tb_buk_uk.Text = ""
        tb_buk_opst_uk.Text = ""
        tb_stepbuk_uk.Text = ""
        tb_CNCmin_uk.Text = ""
        tb_gruppe1_opstart_uk.Text = ""
        tb_laserCNC_tid_uk.Text = ""
        tb_laser_opstart_uk.Text = ""
        tb_combiCNC_tid_uk.Text = ""
        tb_CombiCNCstans_tid_uk.Text = ""
        tb_combi_opstart_uk.Text = ""
        tb_klip_tid_uk.Text = ""
        tb_klip_opstart_uk.Text = ""
        tb_afgrat_uk.Text = ""
        tb_grinding_uk.Text = ""
        tb_brush_uk.Text = ""
        tb_vibration_uk.Text = ""
        tb_rette_uk.Text = ""
        tb_stans_manuel_uk.Text = ""
        tb_countersink_uk.Text = ""
        tb_gevind_uk.Text = ""
        tb_pressnut_uk.Text = ""
        tb_presstag_uk.Text = ""
        tb_boltesvejs_uk.Text = ""
        tb_opstart_kr_uk.Text = ""
        tb_Program_kr_uk.Text = ""
        tb_tackweld_uk.Text = ""
        tb_weld_uk.Text = ""
        tb_grind_weld_uk.Text = ""
        tb_glasbl_uk.Text = ""
        tb_slib_uk.Text = ""
        tb_valsning_uk.Text = ""
        tb_rettesvejs_uk.Text = ""
        tb_tilbudnr.Text = ""
        tb_tapsvejs_uk.Text = ""

    End Sub

    Private Sub lb_antal_opstart_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lb_antal_opstart.LocationChanged

    End Sub

    Private Sub rb_netto_CursorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rb_netto.CursorChanged

    End Sub


    Private Sub bu_hent_CursorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles bu_hent.CursorChanged

    End Sub








    Private Sub Label257_Click(sender As System.Object, e As System.EventArgs) Handles Label257.Click

    End Sub
End Class

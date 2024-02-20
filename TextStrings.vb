Public Class TextStrings

    Private client As String
    Private part_name As String
    Private drawing_num As String
    Private language As String

    Public Sub New(lang As String)
        language = lang
    End Sub

    Public Sub New()

    End Sub

    Public Function setLanguage(lang As String)


        language = lang


    End Function

    Public Function getClientString() As String


        If language = "E" Then
            getClientString = "CLIENT"

        Else
            getClientString = "KUNDE"
        End If


    End Function

    Public Function getPartNameString() As String


        If language = "E" Then
            getPartNameString = "PART NAME"

        Else
            getPartNameString = "BENÆVNELSE"
        End If


    End Function

    Public Function getDrawingNumbString() As String


        If language = "E" Then
            getDrawingNumbString = "Drawing Nr."

        Else
            getDrawingNumbString = "TEGN.NR"
        End If


    End Function

    Public Function getPriceSchemeString() As String


        If language = "E" Then
            getPriceSchemeString = "Price Scheme"

        Else
            getPriceSchemeString = "Prisskema"
        End If

    End Function

    Public Function getQtyO1String() As String


        If language = "E" Then
            getQtyO1String = "Order qty. 1"

        Else
            getQtyO1String = "Ordre str. 1"
        End If


    End Function

    Public Function getQtyO2String() As String


        If language = "E" Then
            getQtyO2String = "Order qty. 2"

        Else
            getQtyO2String = "Ordre str. 2"
        End If


    End Function

    Public Function getQtyO3String() As String


        If language = "E" Then
            getQtyO3String = "Order qty. 3"

        Else
            getQtyO3String = "Ordre str. 3"
        End If


    End Function

    Public Function getQtyO4String() As String


        If language = "E" Then
            getQtyO4String = "Order qty. 4"

        Else
            getQtyO4String = "Ordre str. 4"
        End If


    End Function

    Public Function getQtyO5String() As String


        If language = "E" Then
            getQtyO5String = "Order qty. 5"

        Else
            getQtyO5String = "Ordre str. 5"
        End If


    End Function

    Public Function getQtyString() As String


        If language = "E" Then
            getQtyString = "Qty"

        Else
            getQtyString = "antal"
        End If


    End Function

    Public Function getMansCostString() As String


        If language = "E" Then
            getMansCostString = "Man cost dkk-hours"

        Else
            getMansCostString = "Mand timer kr.-timer"
        End If


    End Function

    Public Function getCNCCostString() As String


        If language = "E" Then
            getCNCCostString = "CNC cost dkk-hours"

        Else
            getCNCCostString = "CNC timer kr.-timer"
        End If


    End Function

    Public Function getTimeCostString() As String


        If language = "E" Then
            getTimeCostString = "Time cost dkk"

        Else
            getTimeCostString = "Timer ialt kr."
        End If


    End Function

    Public Function getOutsourceCostString() As String


        If language = "E" Then
            getOutsourceCostString = "Outsourced serv. dkk"

        Else
            getOutsourceCostString = "Indkøb kr."
        End If


    End Function

    Public Function getRawMatCostString() As String


        If language = "E" Then
            getRawMatCostString = "Raw mat pc/tot dkk"

        Else
            getRawMatCostString = "Råvarer stk/ialt kr."
        End If


    End Function

    Public Function getSumCostString() As String


        If language = "E" Then
            getSumCostString = "Sum pc/tot dkk"

        Else
            getSumCostString = "Samlet stk/ialt kr."
        End If


    End Function

    Public Function getProfitString() As String


        If language = "E" Then
            getProfitString = "Profit"

        Else
            getProfitString = "Avance"
        End If


    End Function

    Public Function getSalepriceString() As String


        If language = "E" Then
            getSalepriceString = "Saleprice pc/tot dkk"

        Else
            getSalepriceString = "Salgpris stk/ialt kr."
        End If


    End Function

    Public Function getActualOfferString() As String


        If language = "E" Then
            getActualOfferString = "Actual offer dkk"

        Else
            getActualOfferString = "Afgivet tilbud kr."
        End If


    End Function

    Public Function getCNCSetupProgramString() As String


        If language = "E" Then
            getCNCSetupProgramString = "CNC Setup/program"

        Else
            getCNCSetupProgramString = "CNC Opstart/program"
        End If


    End Function

    Public Function getSetup1timeString() As String


        If language = "E" Then
            getSetup1timeString = "Setup 1st time"

        Else
            getSetup1timeString = "Opstart 1. gang"
        End If


    End Function

    Public Function getSetupDkkString() As String


        If language = "E" Then
            getSetupDkkString = "Setup dkk"

        Else
            getSetupDkkString = "Opstart kr."
        End If


    End Function

    Public Function getProgramQtyString() As String


        If language = "E" Then
            getProgramQtyString = "Program qty"

        Else
            getProgramQtyString = "Program antal"
        End If


    End Function

    Public Function getSetupProgramsTotalDkkString() As String


        If language = "E" Then
            getSetupProgramsTotalDkkString = "Setup+programs total dkk"

        Else
            getSetupProgramsTotalDkkString = "Opstar+programmer ialt kr."
        End If


    End Function

    Public Function getSalespriceSetupProgramDkkString() As String


        If language = "E" Then
            getSalespriceSetupProgramDkkString = "Saleprice Setup+Prog. dkk"

        Else
            getSalespriceSetupProgramDkkString = "Salgpris Opst.+Prog. kr."
        End If


    End Function


    Public Function getProgramsDkkString() As String


        If language = "E" Then
            getProgramsDkkString = "Programs dkk"

        Else
            getProgramsDkkString = "Programmer kr."
        End If


    End Function

    Public Function getOverruleString() As String


        If language = "E" Then
            getOverruleString = "Overrule"

        Else
            getOverruleString = "Underkend"
        End If


    End Function

    Public Function getGroupTimesByOQty1String() As String


        If language = "E" Then
            getGroupTimesByOQty1String = "Group times by Order qty. 1"

        Else
            getGroupTimesByOQty1String = "GRUPPETIDER ved Ordre str. 1"
        End If


    End Function

    Public Function getHoursString() As String


        If language = "E" Then
            getHoursString = "Hours"

        Else
            getHoursString = "Timer"
        End If


    End Function

    Public Function getGroup1String() As String


        If language = "E" Then
            getGroup1String = "Group 1"

        Else
            getGroup1String = "Gruppe 1"
        End If


    End Function

    Public Function getGroup2String() As String


        If language = "E" Then
            getGroup2String = "Group 2"

        Else
            getGroup2String = "Gruppe 2"
        End If


    End Function

    Public Function getGroup3String() As String


        If language = "E" Then
            getGroup3String = "Group 3"

        Else
            getGroup3String = "Gruppe 3"
        End If


    End Function

    Public Function getGroup4String() As String


        If language = "E" Then
            getGroup4String = "Group 4"

        Else
            getGroup4String = "Gruppe 4"
        End If


    End Function

    Public Function getGroup5String() As String


        If language = "E" Then
            getGroup5String = "Group 5"

        Else
            getGroup5String = "Gruppe 5"
        End If


    End Function

    Public Function getGroup6String() As String


        If language = "E" Then
            getGroup6String = "Group 6"

        Else
            getGroup6String = "Gruppe 6"
        End If


    End Function

    Public Function getCncGiloString() As String


        If language = "E" Then
            getCncGiloString = "CNC + Guillotine"

        Else
            getCncGiloString = "CNC + Klip"
        End If


    End Function

    Public Function getManualString() As String


        If language = "E" Then
            getManualString = "Manual"

        Else
            getManualString = "Myg"
        End If


    End Function

    Public Function getBendSpotRollString() As String


        If language = "E" Then
            getBendSpotRollString = "Bend + spotweld + roll"

        Else
            getBendSpotRollString = "Buk + punktsvejse + valse"
        End If


    End Function

    Public Function getWeldingStraiString() As String


        If language = "E" Then
            getWeldingStraiString = "Welding + straightening"

        Else
            getWeldingStraiString = "Svejsning + opretning"
        End If


    End Function

    Public Function getGrindingString() As String


        If language = "E" Then
            getGrindingString = "Grinding"

        Else
            getGrindingString = "Slibning"
        End If


    End Function

    Public Function getOfficeControlString() As String


        If language = "E" Then
            getOfficeControlString = "Office + Control"

        Else
            getOfficeControlString = "Kontor + Kontrol"
        End If


    End Function

    Public Function getDkkNetByOrderQ1String() As String


        If language = "E" Then
            getDkkNetByOrderQ1String = "dkk net by Order qty. 1"

        Else
            getDkkNetByOrderQ1String = "Kr. netto ved Ordre str. 1"
        End If


    End Function

    Public Function getSelectMaterialString() As String


        If language = "E" Then
            getSelectMaterialString = "SELECT MATERIAL"

        Else
            getSelectMaterialString = "VÆLG MATERIALE"
        End If


    End Function

    Public Function getplateThickString() As String


        If language = "E" Then
            getplateThickString = "plate thickness"

        Else
            getplateThickString = "pl. tykkelse"
        End If


    End Function

    Public Function getModuleSizeString() As String


        If language = "E" Then
            getModuleSizeString = "Module size"

        Else
            getModuleSizeString = "Modulstørrelse"
        End If


    End Function

    Public Function getMaterialGroupString() As String


        If language = "E" Then
            getMaterialGroupString = "Material group"

        Else
            getMaterialGroupString = "Materialegruppe"
        End If


    End Function

    Public Function getGradeString() As String


        If language = "E" Then
            getGradeString = "Grade"

        Else
            getGradeString = "Klasse"
        End If


    End Function

    Public Function getDifficultyString() As String


        If language = "E" Then
            getDifficultyString = "Difficulty"

        Else
            getDifficultyString = "Sværhedsgrad"
        End If


    End Function

    Public Function getPlateQtyString() As String


        If language = "E" Then
            getPlateQtyString = "plate Qty"

        Else
            getPlateQtyString = "Antal plader"
        End If


    End Function

    Public Function getPlateFormatString() As String


        If language = "E" Then
            getPlateFormatString = "Plate format"

        Else
            getPlateFormatString = "Pladeformat"
        End If


    End Function

    Public Function getItemsPerPlateString() As String


        If language = "E" Then
            getItemsPerPlateString = "Items per plate"

        Else
            getItemsPerPlateString = "Emner pr. plade"
        End If


    End Function


    Public Function getPlateConsumptionString() As String


        If language = "E" Then
            getPlateConsumptionString = "Plate consumption"

        Else
            getPlateConsumptionString = "Pladeforbrug"
        End If


    End Function

    Public Function getWasteString() As String


        If language = "E" Then
            getWasteString = "Waste"

        Else
            getWasteString = "Splid"
        End If


    End Function

    Public Function getCostPriceString() As String


        If language = "E" Then
            getCostPriceString = "Cost price dkk/Kg"

        Else
            getCostPriceString = "Salgspris Kr./Kg"
        End If


    End Function

    Public Function getOverruleCostPriceString() As String


        If language = "E" Then
            getOverruleCostPriceString = "Overrule cost price"

        Else
            getOverruleCostPriceString = "Underk. Salgspr."
        End If


    End Function

    Public Function getFactorString() As String


        If language = "E" Then
            getFactorString = "Factor"

        Else
            getFactorString = "Faktor"
        End If


    End Function

    Public Function getUKString() As String


        If language = "E" Then
            getUKString = "Overrule"

        Else
            getUKString = "UK"
        End If


    End Function

    Public Function getOneTo6String() As String


        If language = "E" Then
            getOneTo6String = "1 to 6"

        Else
            getOneTo6String = "1 til 6"
        End If


    End Function

    Public Function getForOQ1String() As String


        If language = "E" Then
            getForOQ1String = "for order qty 1"

        Else
            getForOQ1String = "ved ordrestr. 1"
        End If


    End Function

    Public Function getItemWeightString() As String


        If language = "E" Then
            getItemWeightString = "Item weight"

        Else
            getItemWeightString = "Emnevægt"
        End If


    End Function

    Public Function getKrKgString() As String


        If language = "E" Then
            getKrKgString = "dkk/Kg"

        Else
            getKrKgString = "Kr./Kg"
        End If


    End Function


    Public Function getNetRawMaterialString() As String


        If language = "E" Then
            getNetRawMaterialString = "NET raw material"

        Else
            getNetRawMaterialString = "Råvare NETTO"
        End If


    End Function

    Public Function getGrossRawMaterialString() As String


        If language = "E" Then
            getGrossRawMaterialString = "GROSS raw material"

        Else
            getGrossRawMaterialString = "Råvare BRUTTO"
        End If


    End Function

    Public Function getExclude2000x4000String() As String


        If language = "E" Then
            getExclude2000x4000String = "EXCLUDE size 2000x4000"

        Else
            getExclude2000x4000String = "FRAVÆLG overstørrelse 2000x4000"
        End If


    End Function

    Public Function getExclude1500x3000String() As String


        If language = "E" Then
            getExclude1500x3000String = "EXCLUDE size 1500x3000"

        Else
            getExclude1500x3000String = "FRAVÆLG overstørrelse 1500x3000"
        End If


    End Function

    Public Function getExclude1250x2500String() As String


        If language = "E" Then
            getExclude1250x2500String = "EXCLUDE size 1250x2500"

        Else
            getExclude1250x2500String = "FRAVÆLG overstørrelse 1250x2500"
        End If


    End Function

    Public Function getExclude1000x2000String() As String


        If language = "E" Then
            getExclude1000x2000String = "EXCLUDE size 1000x2000"

        Else
            getExclude1000x2000String = "FRAVÆLG overstørrelse 1000x2000"
        End If


    End Function

    Public Function getBendingGroup3String() As String


        If language = "E" Then
            getBendingGroup3String = "Bending (Group 3)"

        Else
            getBendingGroup3String = "Buk (Gruppe3)"
        End If


    End Function

    Public Function getBiggestMeasureString() As String


        If language = "E" Then
            getBiggestMeasureString = "Biggest Measure"

        Else
            getBiggestMeasureString = "Største mål"
        End If


    End Function

    Public Function getBendString(numb As String) As String


        If language = "E" Then
            getBendString = "Bend " + numb

        Else
            getBendString = "Buk " + numb
        End If


    End Function

    Public Function getUnfoldingString() As String


        If language = "E" Then
            getUnfoldingString = "Unfolding"

        Else
            getUnfoldingString = "Udfoldning"
        End If


    End Function

    Public Function getBendMinOQ1String() As String


        If language = "E" Then
            getBendMinOQ1String = "Bend min./O.qty.1"

        Else
            getBendMinOQ1String = "Buk min./O.str.1"
        End If


    End Function

    Public Function getAreaM2String() As String


        If language = "E" Then
            getAreaM2String = "Area M2"

        Else
            getAreaM2String = "Areal M2"
        End If


    End Function

    Public Function getPaint1SideString() As String


        If language = "E" Then
            getPaint1SideString = "(paint) 1 side"

        Else
            getPaint1SideString = "(maling) 1 side"
        End If


    End Function

    Public Function getStepBendNumbStepsString() As String


        If language = "E" Then
            getStepBendNumbStepsString = "Stepbend, nmr of steps per item"

        Else
            getStepBendNumbStepsString = "Stepbuk, Antal step pr. emne"
        End If


    End Function

    Public Function getStepBendString() As String


        If language = "E" Then
            getStepBendString = "Step bend"

        Else
            getStepBendString = "Step buk"
        End If


    End Function

    Public Function getSetupBendingString() As String


        If language = "E" Then
            getSetupBendingString = "Setup bending"

        Else
            getSetupBendingString = "Opstilling buk"
        End If


    End Function

    Public Function getTotalBendString() As String


        If language = "E" Then
            getTotalBendString = "Total bend time"

        Else
            getTotalBendString = "Buk ialt"
        End If


    End Function

    Public Function getRollingString() As String


        If language = "E" Then
            getRollingString = "Rolling"

        Else
            getRollingString = "Valsning"
        End If


    End Function

    Public Function getCutString() As String


        If language = "E" Then
            getCutString = "Guillotine"

        Else
            getCutString = "Klip"
        End If


    End Function

    Public Function getPunchString() As String


        If language = "E" Then
            getPunchString = "B-punch"

        Else
            getPunchString = "B-stans"
        End If


    End Function

    Public Function getCutMinOQ1String() As String


        If language = "E" Then
            getCutMinOQ1String = "Cut min./O.qty.1"

        Else
            getCutMinOQ1String = "Klip min./O.str.1"
        End If


    End Function

    Public Function getSetupMinString() As String


        If language = "E" Then
            getSetupMinString = "Setup min."

        Else
            getSetupMinString = "Opstart min."
        End If


    End Function

    Public Function getCutTotalOQ1String() As String


        If language = "E" Then
            getCutTotalOQ1String = "Cut total/O.qty.1"

        Else
            getCutTotalOQ1String = "Klip ialt/O.str.1"
        End If


    End Function

    Public Function getNumbHoles2to10String() As String


        If language = "E" Then
            getNumbHoles2to10String = "Numb. of holes (ø2-ø10)"

        Else
            getNumbHoles2to10String = "Antal huller (ø2-ø10)"
        End If


    End Function

    Public Function getNumbHoles11to50String() As String


        If language = "E" Then
            getNumbHoles11to50String = "Numb. of holes (ø11-ø50)"

        Else
            getNumbHoles11to50String = "Antal huller (ø11-ø50)"
        End If


    End Function

    Public Function getNumbHoles51to100String() As String


        If language = "E" Then
            getNumbHoles51to100String = "Numb. of holes (ø51-ø100)"

        Else
            getNumbHoles51to100String = "Antal huller (ø51-ø100)"
        End If


    End Function

    Public Function getTotCuttingLen100String() As String


        If language = "E" Then
            getTotCuttingLen100String = "Tot cutting length (>ø100)"

        Else
            getTotCuttingLen100String = "Skærelængde ialt (>ø100)"
        End If


    End Function

    Public Function getCNCMinOQ1String() As String


        If language = "E" Then
            getCNCMinOQ1String = "CNC min./O.qty.1"

        Else
            getCNCMinOQ1String = "CNC min./O.str.1"
        End If

    End Function

    Public Function getCLaserTotOQ1String() As String


        If language = "E" Then
            getCLaserTotOQ1String = "C-laser total /O.qty.1"

        Else
            getCLaserTotOQ1String = "C-laser ialt /O.str.1"
        End If

    End Function

    Public Function getToolChangeString() As String


        If language = "E" Then
            getToolChangeString = "Tool change"

        Else
            getToolChangeString = "Værktøjsskift"
        End If

    End Function

    Public Function getNumbOfStrokesString() As String


        If language = "E" Then
            getNumbOfStrokesString = "Strokes per holes"

        Else
            getNumbOfStrokesString = "Antal slag til huller"
        End If

    End Function

    Public Function getDStansTotOQ1String() As String


        If language = "E" Then
            getDStansTotOQ1String = "D-stans tot/O.qty.1"

        Else
            getDStansTotOQ1String = "D-stans ialt/O.str.1"
        End If

    End Function

    Public Function getCncMinLaserOQ1String() As String


        If language = "E" Then
            getCncMinLaserOQ1String = "CNC min.laser/O.qty.1"

        Else
            getCncMinLaserOQ1String = "CNC min.laser/O.str.1"
        End If

    End Function

    Public Function getCncMinStansOQ1String() As String


        If language = "E" Then
            getCncMinStansOQ1String = "CNC min.stans/O.qty.1"

        Else
            getCncMinStansOQ1String = "CNC min.stans/O.str.1"
        End If

    End Function

    Public Function getBKombiTotOQ1String() As String


        If language = "E" Then
            getBKombiTotOQ1String = "B-combi tot/O.qty.1"

        Else
            getBKombiTotOQ1String = "B-combi ialt/O.str.1"
        End If

    End Function

    Public Function getPerEmneString() As String


        If language = "E" Then
            getPerEmneString = "per item"

        Else
            getPerEmneString = "pr.emne"
        End If

    End Function

    Public Function getManualDeburringString() As String


        If language = "E" Then
            getManualDeburringString = "Manual deburring"

        Else
            getManualDeburringString = "Manuel Afrat/fil"
        End If

    End Function


    Public Function getMinOQ1String() As String


        If language = "E" Then
            getMinOQ1String = "min./O.qty.1"

        Else
            getMinOQ1String = "min./O.str.1"
        End If

    End Function

    Public Function getMinOQ1String2() As String


        If language = "E" Then
            getMinOQ1String2 = "min./" + Environment.NewLine + "O.qty.1"

        Else
            getMinOQ1String2 = "min./" + Environment.NewLine + "O.str.1"
        End If

    End Function

    Public Function getBrushDeburringString() As String


        If language = "E" Then
            getBrushDeburringString = "Brush deburring"

        Else
            getBrushDeburringString = "Børsteafgratning"
        End If

    End Function

    Public Function getVibrateDeburringString() As String


        If language = "E" Then
            getVibrateDeburringString = "Vibrate deburring"

        Else
            getVibrateDeburringString = "Vibrationsafgratn."
        End If

    End Function

    Public Function getStraightMachineString() As String


        If language = "E" Then
            getStraightMachineString = "Straightener machine"

        Else
            getStraightMachineString = "Rettemaskine"
        End If

    End Function

    Public Function getManualPucnchString() As String


        If language = "E" Then
            getManualPucnchString = "Manual Punch"

        Else
            getManualPucnchString = "Stans manuel "
        End If

    End Function

    Public Function getPresstagString() As String


        If language = "E" Then
            getPresstagString = "Presstag" + Environment.NewLine + "qty per item"

        Else
            getPresstagString = "Presstag" + Environment.NewLine + "antal pr emne"
        End If

    End Function

    Public Function getPresnutString() As String


        If language = "E" Then
            getPresnutString = "Presnut" + Environment.NewLine + "qty per item"

        Else
            getPresnutString = "Presmøtrik" + Environment.NewLine + "antal pr emne"
        End If

    End Function

    Public Function getScrewWeldingString() As String


        If language = "E" Then
            getScrewWeldingString = "ScrewWeld" + Environment.NewLine + "qty per item"

        Else
            getScrewWeldingString = "Svejsestag" + Environment.NewLine + "antal pr emne"
        End If

    End Function

    Public Function getThreadingString() As String


        If language = "E" Then
            getThreadingString = "THREADING"

        Else
            getThreadingString = "GEVIND"
        End If

    End Function

    Public Function getChamferingString() As String


        If language = "E" Then
            getChamferingString = "CHAMFERING"

        Else
            getChamferingString = "UNDERSÆNKNING"
        End If

    End Function

    Public Function getThreadSizeString() As String


        If language = "E" Then
            getThreadSizeString = "Thread size"

        Else
            getThreadSizeString = "Gevind str."
        End If

    End Function

    Public Function getQtyPerItemString() As String


        If language = "E" Then
            getQtyPerItemString = "Qty per item"

        Else
            getQtyPerItemString = "Antal pr emne"
        End If

    End Function

    Public Function getDiameterThreadSizeString() As String


        If language = "E" Then
            getDiameterThreadSizeString = "Diameter" + Environment.NewLine + "(Thread size)"

        Else
            getDiameterThreadSizeString = "Diameter" + Environment.NewLine + "(Gevind str.)"
        End If

    End Function

    Public Function getTotalQtyString() As String


        If language = "E" Then
            getTotalQtyString = "Total qty"

        Else
            getTotalQtyString = "antal ialt"
        End If

    End Function

    Public Function getOver15mmString() As String


        If language = "E" Then
            getOver15mmString = "over 15 mm(M10-M12)"

        Else
            getOver15mmString = "over 15 mm(M10-M12)"
        End If

    End Function

    Public Function getOnlyOQ1String() As String


        If language = "E" Then
            getOnlyOQ1String = "(only Order qty. 1)"

        Else
            getOnlyOQ1String = "(kun Ordre str. 1)"
        End If

    End Function

    Public Function getGroup3SpotWeldString() As String


        If language = "E" Then
            getGroup3SpotWeldString = "Group 3 SpotWelding"

        Else
            getGroup3SpotWeldString = "Gruppe 3 Punktsvejsning"
        End If

    End Function

    Public Function getnumberOfWeldsPerItemString() As String


        If language = "E" Then
            getnumberOfWeldsPerItemString = "Number of welds per item"

        Else
            getnumberOfWeldsPerItemString = "antal svejsesømme pr emne"
        End If

    End Function

    Public Function getnumberOfSpotWeldsPerItemString() As String


        If language = "E" Then
            getnumberOfSpotWeldsPerItemString = "Number of spotWelds per item"

        Else
            getnumberOfSpotWeldsPerItemString = "antal punktsvejsninger pr emne"
        End If

    End Function

    Public Function getSpotWeldingString() As String


        If language = "E" Then
            getSpotWeldingString = "Spotwelding"

        Else
            getSpotWeldingString = "punktsvejsn."
        End If

    End Function

    Public Function getGroup4WeldingString() As String


        If language = "E" Then
            getGroup4WeldingString = "Group 4 Welding"

        Else
            getGroup4WeldingString = "Gruppe 4 Svejsning"
        End If

    End Function

    Public Function getWeldingLenghtPerItemString() As String


        If language = "E" Then
            getWeldingLenghtPerItemString = "welding length per item in mm"

        Else
            getWeldingLenghtPerItemString = "svejselængde pr emne i mm"
        End If

    End Function

    Public Function getOnlyTackWeldingString() As String


        If language = "E" Then
            getOnlyTackWeldingString = "Only tack welding"

        Else
            getOnlyTackWeldingString = "kun hæfte svejsn."
        End If

    End Function

    Public Function getFullWeldString() As String


        If language = "E" Then
            getFullWeldString = "Full welding"

        Else
            getFullWeldString = "fuld svejsning"
        End If

    End Function

    Public Function getStraighteningString() As String


        If language = "E" Then
            getStraighteningString = "Straightening"

        Else
            getStraighteningString = "opretning"
        End If

    End Function

    Public Function getTapQtyString() As String


        If language = "E" Then
            getTapQtyString = "Tap Qty"

        Else
            getTapQtyString = "antal tappe"
        End If

    End Function

    Public Function getTapWeldingString() As String


        If language = "E" Then
            getTapWeldingString = "Tap welding"

        Else
            getTapWeldingString = "tap svejsning"
        End If

    End Function


    Public Function getWeldingDifficultyString() As String


        If language = "E" Then
            getWeldingDifficultyString = "WELDING DIFFICULTY"

        Else
            getWeldingDifficultyString = "SVEJSNING SVÆRHEDSGRAD"
        End If

    End Function

    Public Function getTotalWeldingTimeString() As String


        If language = "E" Then
            getTotalWeldingTimeString = "Total welding time"

        Else
            getTotalWeldingTimeString = "Svejsning ialt"
        End If

    End Function


    Public Function getSandblastingString() As String


        If language = "E" Then
            getSandblastingString = "Sandblasting"

        Else
            getSandblastingString = "Glasblæsning"
        End If

    End Function

    Public Function getWeldGrindingString() As String


        If language = "E" Then
            getWeldGrindingString = "Weld grindng"

        Else
            getWeldGrindingString = "slibe svejsn."
        End If

    End Function

    Public Function getMinutesString() As String


        If language = "E" Then
            getMinutesString = "Minutes"

        Else
            getMinutesString = "minutter"
        End If

    End Function

    Public Function getOfficeString() As String


        If language = "E" Then
            getOfficeString = "Office"

        Else
            getOfficeString = "Kontor"
        End If

    End Function

    Public Function getControlString() As String


        If language = "E" Then
            getControlString = "Control"

        Else
            getControlString = "Kontrol"
        End If

    End Function



    Public Function getPresnutPresstagString() As String


        If language = "E" Then
            getPresnutPresstagString = "Presnut / Presstag"

        Else
            getPresnutPresstagString = "Presmøtrikker/stag"
        End If

    End Function
    Public Function getIronString() As String


        If language = "E" Then
            getIronString = "Iron"

        Else
            getIronString = "Jern"
        End If

    End Function

    Public Function getStainlessString() As String


        If language = "E" Then
            getStainlessString = "Stainless"

        Else
            getStainlessString = "Rustfri"
        End If

    End Function


    Public Function getSpendingsString() As String


        If language = "E" Then
            getSpendingsString = "Spendings"

        Else
            getSpendingsString = "Forbrug"
        End If

    End Function


    Public Function getPresstagString2() As String


        If language = "E" Then
            getPresstagString2 = "Presstag"

        Else
            getPresstagString2 = "Presstag"
        End If

    End Function

    Public Function getPresnutString2() As String


        If language = "E" Then
            getPresnutString2 = "Presnut"

        Else
            getPresnutString2 = "Presmøtrik"
        End If

    End Function

    Public Function getScreWeldString2() As String


        If language = "E" Then
            getScreWeldString2 = "Screw weld"

        Else
            getScreWeldString2 = "Svejsestag"
        End If

    End Function

    Public Function getAdditionalMatString() As String


        If language = "E" Then
            getAdditionalMatString = "Additional mat."

        Else
            getAdditionalMatString = "tilsatsmatr"
        End If

    End Function

    Public Function getDkkPerItemString() As String


        If language = "E" Then
            getDkkPerItemString = "dkk per item"

        Else
            getDkkPerItemString = "Kr. pr. emne"
        End If

    End Function

    Public Function getAdditionalInfoString() As String


        If language = "E" Then
            getAdditionalInfoString = "ADDITIONAL INFO"

        Else
            getAdditionalInfoString = "BEMÆRKNINGER"
        End If

    End Function


    Public Function getSamkorselString() As String


        If language = "E" Then
            getSamkorselString = "Samkørsel with (reduced setup time)"

        Else
            getSamkorselString = "Samkørsel med (reduceret opstartstid)"
        End If

    End Function

    Public Function getOprDateString() As String


        If language = "E" Then
            getOprDateString = "Opr. Date"

        Else
            getOprDateString = "Opr. Dato"
        End If

    End Function

    Public Function getOperatorString() As String


        If language = "E" Then
            getOperatorString = "Operator"

        Else
            getOperatorString = "Operatør"
        End If

    End Function

    Public Function getUpdateDateString() As String


        If language = "E" Then
            getUpdateDateString = "Update date"

        Else
            getUpdateDateString = "Ændr.Dato"
        End If

    End Function

    Public Function getFileNameString() As String


        If language = "E" Then
            getFileNameString = "File name"

        Else
            getFileNameString = "Filnavn"
        End If

    End Function

    Public Function getHoursPricesString() As String


        If language = "E" Then
            getHoursPricesString = "Hour prices dkk"

        Else
            getHoursPricesString = "Timesats Kr."
        End If

    End Function

    Public Function getDenmarkString() As String


        If language = "E" Then
            getDenmarkString = "Denmark"

        Else
            getDenmarkString = "Danmark"
        End If

    End Function

    Public Function getPolandString() As String


        If language = "E" Then
            getPolandString = "Poland"

        Else
            getPolandString = "Polen"
        End If

    End Function
    Public Function getmanString() As String


        If language = "E" Then
            getmanString = "Man"

        Else
            getmanString = "mand"
        End If

    End Function

    Public Function getPunch2String() As String


        If language = "E" Then
            getPunch2String = "B punch"

        Else
            getPunch2String = "B stans"
        End If

    End Function


    Public Function getSaveString() As String


        If language = "E" Then
            getSaveString = "SAVE OFFER"

        Else
            getSaveString = "GEM TILBUD"
        End If

    End Function

    Public Function getLoadString() As String


        If language = "E" Then
            getLoadString = "LOAD OFFER"

        Else
            getLoadString = "HENT TILBUD"
        End If

    End Function

    Public Function getOfferNumbString() As String


        If language = "E" Then
            getOfferNumbString = "OFFER NR."

        Else
            getOfferNumbString = "TILBUD NR."
        End If

    End Function


    Public Function getSurfaceTreatmentOutSourceString() As String


        If language = "E" Then
            getSurfaceTreatmentOutSourceString = "SURFACE TREATMENT / OUTSOURCE"

        Else
            getSurfaceTreatmentOutSourceString = "OVERFLADEBEHANDLING / INDKØB"
        End If

    End Function

    Public Function getSupplierString() As String


        If language = "E" Then
            getSupplierString = "SUPPLIER"

        Else
            getSupplierString = "LEVERANDØR"
        End If

    End Function

    Public Function getSurfaceQtyString() As String


        If language = "E" Then
            getSurfaceQtyString = "Surface Qty"

        Else
            getSurfaceQtyString = "Antal flader"
        End If

    End Function

    Public Function getCoatingDkkPcString() As String


        If language = "E" Then
            getCoatingDkkPcString = "Coating dkk/pc"

        Else
            getCoatingDkkPcString = "Afdækn. Kr./stk"
        End If

    End Function

    Public Function getPriceDkkPcString() As String


        If language = "E" Then
            getPriceDkkPcString = "Price dkk/pc"

        Else
            getPriceDkkPcString = "Pris Kr./stk"
        End If

    End Function

    Public Function getDkkPcString() As String


        If language = "E" Then
            getDkkPcString = "dkk pc"

        Else
            getDkkPcString = "kr stk"
        End If

    End Function

    Public Function getOfferDkkPcString() As String


        If language = "E" Then
            getOfferDkkPcString = "Offer dkk/pc"

        Else
            getOfferDkkPcString = "Tilbud Kr./stk"
        End If

    End Function

    Public Function getOfferAppliesToString() As String


        If language = "E" Then
            getOfferAppliesToString = "Offer applies to"

        Else
            getOfferAppliesToString = "TILBUD GÆLDER VED"
        End If

    End Function

    Public Function getPcsString() As String


        If language = "E" Then
            getPcsString = "Pcs"

        Else
            getPcsString = "STK."
        End If

    End Function

    Public Function getProfitString2() As String


        If language = "E" Then
            getProfitString2 = "Profit %"

        Else
            getProfitString2 = "Avance %"
        End If

    End Function

    Public Function getEnterSomeDataFirstString() As String


        If language = "E" Then
            getEnterSomeDataFirstString = "Enter some data first"

        Else
            getEnterSomeDataFirstString = "Indtast nogle data først"
        End If

    End Function

    Public Function getCustomerIsMissingString() As String


        If language = "E" Then
            getCustomerIsMissingString = "Customer name is missing"

        Else
            getCustomerIsMissingString = "Kundenavn mangle"
        End If

    End Function

    Public Function getDrawingIsMissingString() As String


        If language = "E" Then
            getDrawingIsMissingString = "Drawing number is missing"

        Else
            getDrawingIsMissingString = "Tegningsnummer mangler"
        End If

    End Function

    Public Function getLoadingPleaseWaitString() As String


        If language = "E" Then
            getLoadingPleaseWaitString = "LOADING...PLEASE WAIT"

        Else
            getLoadingPleaseWaitString = "HENTER...VENT VENLIGST"
        End If

    End Function

    Public Function getExcelTemplateCouldNotBeOpenString() As String


        If language = "E" Then
            getExcelTemplateCouldNotBeOpenString = "Excel template cannot be open"

        Else
            getExcelTemplateCouldNotBeOpenString = "Excel templaten kan ikke åbnes"
        End If

    End Function

    Public Function getSavingPleaseWaitString() As String


        If language = "E" Then
            getSavingPleaseWaitString = "SAVING...PLEASE WAIT"

        Else
            getSavingPleaseWaitString = "GEMMER... VENT VENLIGST"
        End If

    End Function

    Public Function getNoSurfaceTreatmentString() As String


        If language = "E" Then
            getNoSurfaceTreatmentString = "No Surface Treatment"

        Else
            getNoSurfaceTreatmentString = "Ingen Overfladebehandling"
        End If

    End Function

    Public Function getPowderCoatingString() As String


        If language = "E" Then
            getPowderCoatingString = "Powder Coating"

        Else
            getPowderCoatingString = "Pulverlak"
        End If

    End Function


    Public Function getWetVarnishString() As String


        If language = "E" Then
            getWetVarnishString = "Wet Varnish"

        Else
            getWetVarnishString = "Vådlak"
        End If

    End Function


    Public Function getChromiteString() As String


        If language = "E" Then
            getChromiteString = "Chromite (Iron)"

        Else
            getChromiteString = "Chromit (jern)"
        End If

    End Function

    Public Function getAnodizingString() As String


        If language = "E" Then
            getAnodizingString = "Anodizing (natural / black)"

        Else
            getAnodizingString = "Eloxering (natur/sort)"
        End If

    End Function

    Public Function getForBendInPlateThicknessElipsisisString() As String


        If language = "E" Then
            getForBendInPlateThicknessElipsisisString = "For bends in plate thicknesses over 10mm, the calculation is uncertain, ask the bending operator about bending time"

        Else
            getForBendInPlateThicknessElipsisisString = "Ved buk i pladetykkelser over 10mm er beregningen usikker, spørg bukoperatøren om buk tid "
        End If

    End Function

    Public Function getNumberOfWeldsAndWeldLenghtString() As String


        If language = "E" Then
            getNumberOfWeldsAndWeldLenghtString = "Number of welds and weld length must be setted"

        Else
            getNumberOfWeldsAndWeldLenghtString = "Antal svejsninger og svejselængde skal være udfyld"
        End If

    End Function

    Public Function getThatMaterialCannotBeWeldedString() As String


        If language = "E" Then
            getThatMaterialCannotBeWeldedString = "That material cannot be welded"

        Else
            getThatMaterialCannotBeWeldedString = "Materialet kan ikke svejses"
        End If

    End Function


    Public Function getPlateThicknessIsMax6String() As String


        If language = "E" Then
            getPlateThicknessIsMax6String = "Plate thickness is max 6mm for spot welding in iron and stainless steel"

        Else
            getPlateThicknessIsMax6String = "Pladetykkelse er max 6mm for punktsvejsning i jern og rustfri"
        End If

    End Function

    Public Function getPlateThicknessIsMax4ForSpotString() As String


        If language = "E" Then
            getPlateThicknessIsMax4ForSpotString = "Plate thickness is max 4mm for spot welding in Aluminum"

        Else
            getPlateThicknessIsMax4ForSpotString = "Pladetykkelse er max 4mm for punktsvejsning i Aluminium"
        End If

    End Function

    Public Function getInStockString() As String


        If language = "E" Then
            getInStockString = "IN STOCK"

        Else
            getInStockString = "TIL LAGER"
        End If

    End Function


    Public Function getOnlyCLaserCanBeUsedString() As String


        If language = "E" Then
            getOnlyCLaserCanBeUsedString = "Only C-LASER CAN be used for 2000x4000 format"

        Else
            getOnlyCLaserCanBeUsedString = "Kun C-LASER KAN anvendes til format 2000x4000"
        End If

    End Function

    Public Function getTheItemDontFitInThePlateString() As String


        If language = "E" Then
            getTheItemDontFitInThePlateString = "The item does not fit on one plate"

        Else
            getTheItemDontFitInThePlateString = "Emnet kan ikke være på 1 plade"
        End If

    End Function

    Public Function getAllFormatsAreDeselectedString() As String


        If language = "E" Then
            getAllFormatsAreDeselectedString = "ALL FORMATS ARE DESELECTED"

        Else
            getAllFormatsAreDeselectedString = "ALLE FORMATER ER FRAVALGT"
        End If

    End Function


    Public Function getTheItemIsTooBigForTheStraigString() As String


        If language = "E" Then
            getTheItemIsTooBigForTheStraigString = "The item is too big for the Straightener machine"

        Else
            getTheItemIsTooBigForTheStraigString = "Emnet er for stort til Rettemaskinen"
        End If

    End Function

    Public Function getThePlateThicknessIsTooBigForTheStraigString() As String


        If language = "E" Then
            getThePlateThicknessIsTooBigForTheStraigString = "The plate thickness is too big for the Straightener machine"

        Else
            getThePlateThicknessIsTooBigForTheStraigString = "Pladetykkelsen er for stort til Rettemaskinen"
        End If

    End Function

    Public Function getTheItemIsTooBigForTheSteelmasterString() As String


        If language = "E" Then
            getTheItemIsTooBigForTheSteelmasterString = "The item is too big for the Steelmaster"

        Else
            getTheItemIsTooBigForTheSteelmasterString = "Emnet er for stort til Steelmasteren"
        End If

    End Function

    Public Function getTheItemIsTooSmallForTheSteelmasterString() As String


        If language = "E" Then
            getTheItemIsTooSmallForTheSteelmasterString = "The item is too small for the Steelmaster"

        Else
            getTheItemIsTooSmallForTheSteelmasterString = "Emnet er for lille til Steelmasteren"
        End If

    End Function

    Public Function getTheItemIsTooBigForTheGrindmasterString() As String


        If language = "E" Then
            getTheItemIsTooBigForTheGrindmasterString = "The item is too big for the Grindingmaster"

        Else
            getTheItemIsTooBigForTheGrindmasterString = "Emnet er for stort til Grindingmasteren"
        End If

    End Function

    Public Function getTheItemIsTooSmallForTheGrindmasterString() As String


        If language = "E" Then
            getTheItemIsTooSmallForTheGrindmasterString = "The item is too small for the Grindingmaster"

        Else
            getTheItemIsTooSmallForTheGrindmasterString = "Emnet er for lille til Grindingmasteren"
        End If

    End Function

    Public Function getNumberOfWeldsAndNUmberOfSpotWeldsString() As String


        If language = "E" Then
            getNumberOfWeldsAndNUmberOfSpotWeldsString = "Number of welds and number of spot welds must be setted"

        Else
            getNumberOfWeldsAndNUmberOfSpotWeldsString = "Antal svejsesømme og antal punktsvejsninger skal være udfyldt"
        End If

    End Function






End Class

Public Class oper_sequence
    Dim totalPiece As Integer
    Dim client As String
    Dim operatorr As String
    Dim drawing As String
    Dim revision As String
    Dim itemName As String
    Dim officeTime As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If validateEntries() Then

            createMethodCard()

        End If
    End Sub

    Public Sub setLabelValues(ByRef runTime, ByRef setupTime, ByRef totalPieces, aoperatorr, adrawing, arevision, aitemName, aclient)
        totalPiece = totalPieces
        client = aclient
        operatorr = aoperatorr
        drawing = adrawing
        revision = arevision
        itemName = aitemName
        officeTime = runTime(27)
        Dim i = 0
        For i = 0 To 27
            If isMKOperation(i) = True Then
                If runTime(i) <> 0 Then
                    run_array(i).Text = CStr(runTime(i))
                    setup_array(i).Text = CStr(setupTime(i))
                Else
                    combo_array(i).Enabled = False

                End If
            End If
        Next


    End Sub

    Private Sub createMethodCard()

        Dim i = 0
        Dim ex As New Microsoft.Office.Interop.Excel.Application
        Dim wb As Microsoft.Office.Interop.Excel.Workbook


        Enabled = False
        Try
            'Åben template xls filen DK
            wb = ex.Workbooks.Open("\\akspol\AKS Gruppen Dokumenter\Production documents\AKS Poland\Zamówienia\Metal Tilbud Files\Method Cards\Template\MK2.xltx")
            'Åben template xls filen POL
            'wb = ex.Workbooks.Open("\\Akspol\AKS Gruppen Dokumenter Polen\Tilbud\Metaltilbud\Tilbudsfiler\tilbudsdata.xlt")


        Catch err As Exception
            'Xls filen kan ikke åbnes.
            Enabled = True
            MsgBox("The Excel template cannot be opened")
            ex.Quit()
            Exit Sub
        End Try

        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets.Item(1)
        Dim sheet2 As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets.Item(2)
        ' Dim sheet As Excel.Worksheet = wb.Worksheets.Item(1)
        'Læs data ud i kolonne D
        sheet.Cells(2, 9) = operatorr
        sheet.Cells(3, 9) = drawing
        sheet.Cells(3, 13) = revision
        sheet.Cells(4, 9) = itemName
        sheet.Cells(2, 22) = totalPiece
        sheet.Cells(6, 27) = officeTime

        Dim operationPosition = 6

        Dim selecteds As New List(Of KeyValuePair(Of Integer, Integer))
        For i = 0 To 27

            If combo_array(i).SelectedItem <> "" Then

                Dim operSequence = Convert.ToInt32(Convert.ToInt32(combo_array(i).SelectedItem))
                selecteds.Add(New KeyValuePair(Of Integer, Integer)(operSequence, i))

            End If

        Next

        selecteds.Sort(Function(kv1, kv2) CInt(kv1.Key).CompareTo(CInt(kv2.Key)))


        Dim count = selecteds.Count()
        Dim runTime(count) As Decimal
        Dim setupTime(count) As Decimal
        For i = 0 To count - 1
            Dim operation = selecteds.Item(i).Value
            If isMKOperation(operation) = True Then
                setupTime(i) = Convert.ToDecimal(setup_array(operation).Text)
                runTime(i) = Convert.ToDecimal(run_array(operation).Text)
            Else
                setupTime(i) = Convert.ToDecimal(tb_setup_array(operation).Text)
                runTime(i) = Convert.ToDecimal(tb_run_array(operation).Text)
            End If

            setOperationInMK(runTime(i), setupTime(i), sheet, sheet2, operationPosition, operation + 1)
        Next

        wb.SaveAs("\\akspol\AKS Gruppen Dokumenter\Production documents\AKS Poland\Zamówienia\Metal Tilbud Files\Method Cards\" & client & "_" & drawing & ".xlsx" & " ")
        wb.Close()
        ex.Quit()
        'Process.Start("C:\Users\admbd\Desktop\" & drawing & ".xlsx")
        Enabled = True
        Dim result As DialogResult = MessageBox.Show("Do you want to open it?", "Method Card Saved Succesfully", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Process.Start("\\akspol\AKS Gruppen Dokumenter\Production documents\AKS Poland\Zamówienia\Metal Tilbud Files\Method Cards\" & client & "_" & drawing & ".xlsx")
        ElseIf result = DialogResult.No Then
        End If

        Close()

    End Sub

    Public Sub setOperationInMK(ByVal runTime As Decimal, ByVal setupTime As Decimal, ByRef sheet As Microsoft.Office.Interop.Excel.Worksheet, ByRef sheet2 As Microsoft.Office.Interop.Excel.Worksheet, ByRef operationPosition As Integer, ByRef operationNumber As Integer)

        Dim total = runTime
        Dim timePerPiece = total / totalPiece
        Dim setUp = setupTime
        Dim timePerpieceStr As String
        If timePerPiece < 1 Then

            timePerpieceStr = Fix(timePerPiece * 60).ToString + "s"



        Else
            If timePerPiece > 60 Then
                timePerpieceStr = Math.Round(timePerPiece / 60, 2).ToString + "h"

            Else
                timePerpieceStr = Math.Round(timePerPiece, 1).ToString + "m"
            End If
        End If



        sheet.Cells(operationPosition, 19) = operationNumber
        If setUp <> 0 Then
            sheet.Cells(operationPosition, 21) = setUp
        Else
            sheet.Cells(operationPosition, 21) = sheet2.Cells(operationNumber + 1, 3)
        End If
        sheet.Cells(operationPosition, 22) = timePerpieceStr
        operationPosition = operationPosition + 2


    End Sub

    Private Function validateEntries() As Boolean

        Dim i = 0

        Dim selecteds As New ArrayList
        For i = 0 To 27
            If combo_array(i).SelectedItem <> "" Then


                Dim operNumber = Convert.ToInt32(Convert.ToInt32(combo_array(i).SelectedItem))

                'If the operation is one of the inputted manually and thier text boxes are not good then throw error message
                If isMKOperation(i) = False And (Not IsNumeric(tb_setup_array(i).Text) Or Not IsNumeric(tb_run_array(i).Text)) Then
                    MsgBox("Values for setup and run time must be set correctly for: " + lb_array(i).Text)
                    Return False
                End If

                If Not selecteds.Contains(operNumber) Then
                    selecteds.Add(operNumber)
                Else
                    MsgBox("Please Correct the duplicates in the sequence")
                    Return False
                End If

            Else
                If run_array(i).Text <> "" Then

                    MsgBox("Operation <" + lb_array(i).Text + "> must be included in the sequence")
                    Return False

                End If

                If tb_setup_array(i).Text <> "" Or tb_run_array(i).Text <> "" Then

                    MsgBox("Operation <" + lb_array(i).Text + "> has some inputs but it doesn't have sequence number")
                    Return False

                End If

            End If

        Next


        If Not selecteds.Contains(1) Then
            MsgBox("The first operation must be <1>")
            Return False
        End If

        selecteds.Sort()
        Dim count = selecteds.Count()
        For i = 0 To count

            If (i + 1 < count) Then
                If selecteds.Item(i + 1) - 1 <> selecteds.Item(i) Then
                    MsgBox("The sequence is not correct, can't jump from " + CStr(selecteds.Item(i)) + " To " + CStr(selecteds.Item(i + 1)))
                    Return False
                End If
            End If

        Next


        Return True

    End Function

End Class
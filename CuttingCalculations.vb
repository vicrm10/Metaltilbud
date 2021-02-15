Public Class CuttingCalculations

    Public Function CalculateCuttingTime(ByVal CountRow As Integer, ByVal subjects1sheet As Integer, ByVal numberofsheets As Integer, ByVal sheetthickness As Double, ByVal modulesize As Integer, ByVal Antal As Double) As Double
        Dim objDatabaseIO As DatabaseIO

        objDatabaseIO = New DatabaseIO

        'CalculateCuttingTime = ((subjects1sheet + CountRow) * numberofsheets) * (objDatabaseIO.gettimeprcut(modulesize) / 60) * objDatabaseIO.GetThicknessmultiplier(sheetthickness)

        CalculateCuttingTime = (Antal + CountRow) * (objDatabaseIO.gettimeprcut(modulesize) / 60) * objDatabaseIO.GetThicknessmultiplier(sheetthickness)


    End Function
    Public Function GetCuttingStartup(ByVal Numberofsheets As Double, ByVal Sheetthickness As Double, ByVal SheetX As Double, ByVal SheetY As Double) As Double
        Dim sheetsize As Double
        Dim Sheetsizemultiplier As Double
        Dim objDatabaseIO As DatabaseIO

        sheetsize = SheetX * SheetY
        Sheetsizemultiplier = 40 / 60

        If sheetsize = 2000000 Then
            Sheetsizemultiplier = 30 / 60
        End If
        If sheetsize = 4500000 Then
            Sheetsizemultiplier = 60 / 60
        End If

        objDatabaseIO = New DatabaseIO


        GetCuttingStartup = 5 + ((Numberofsheets * Sheetsizemultiplier) * objDatabaseIO.GetThicknessmultiplier(Sheetthickness))

    End Function
End Class

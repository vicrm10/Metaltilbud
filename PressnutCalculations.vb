Public Class PressnutCalculations
    Public Function CalculatePressnutTime(ByVal modulesize As Integer, ByVal iNumberOfNuts As Integer) As Double
        Dim objDatabaseIO As DatabaseIO

        objDatabaseIO = New DatabaseIO

        CalculatePressnutTime = objDatabaseIO.GetPressnut100min(modulesize) * iNumberOfNuts

    End Function
    Public Function CalculateBoltweldTime(ByVal modulesize As Integer, ByVal iNumberOfNuts As Integer) As Double
        Dim objDatabaseIO As DatabaseIO

        objDatabaseIO = New DatabaseIO

        CalculateBoltweldTime = objDatabaseIO.GetBoltweld100min(modulesize) * iNumberOfNuts
    End Function
    Public Function CalculatePresstagTime(ByVal modulesize As Integer, ByVal iNumberOfNuts As Integer) As Double
        Dim objDatabaseIO As DatabaseIO

        objDatabaseIO = New DatabaseIO

        CalculatePresstagTime = objDatabaseIO.GetPresstag100min(modulesize) * iNumberOfNuts
    End Function
End Class

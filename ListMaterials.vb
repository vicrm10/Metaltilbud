Public Class ListMaterials
    Private palMaterials As ArrayList
    Private WithEvents pcmbTarget As ComboBox
    Private ptxtMaterialGroup As Label
    Private ptxtKlasse As Label
    Private ptxtKilopris As Label


    Public Sub New(ByRef cmbTarget As ComboBox, ByRef txtMaterialGroup As Label, ByRef txtKlasse As Label, ByRef txtKilopris As Label)
        pcmbTarget = cmbTarget
        ptxtMaterialGroup = txtMaterialGroup
        ptxtKlasse = txtKlasse
        ptxtKilopris = txtKilopris
        RefreshData()
    End Sub

    Public Sub RefreshData()
        Dim objDatabaseIO As DatabaseIO


        objDatabaseIO = New DatabaseIO
        palMaterials = objDatabaseIO.GetMaterials()
    End Sub

    Public Sub RefreshData(lang As String)
        Dim objDatabaseIO As DatabaseIO
        Dim materialStrings As ArrayList
        Dim i As Integer

        objDatabaseIO = New DatabaseIO
        materialStrings = objDatabaseIO.GetMaterialsStrings(lang)
        For i = 0 To palMaterials.Count - 1
            palMaterials(i)(1) = materialStrings(i)

        Next

    End Sub

    Public Sub List()
        Dim alItem As ArrayList

        Dim i As Integer


        pcmbTarget.Items.Clear()
        For i = 0 To palMaterials.Count - 1
            alItem = palMaterials(i)
            pcmbTarget.Items.Add(alItem(0).ToString() + " " + alItem(1).ToString())
        Next
        If palMaterials.Count > 0 Then
            pcmbTarget.SelectedIndex = 0
        End If
    End Sub

    Public Function GetMaterialType() As Integer
        Dim alItem As ArrayList


        alItem = palMaterials(pcmbTarget.SelectedIndex)
        GetMaterialType = Convert.ToInt32(alItem(0))
    End Function

    Private Sub pcmbTarget_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pcmbTarget.SelectedIndexChanged
        Dim alItem As ArrayList


        alItem = palMaterials(pcmbTarget.SelectedIndex)
        ptxtMaterialGroup.Text = alItem(3).ToString()
        ptxtKlasse.Text = alItem(5).ToString()
        ptxtKilopris.Text = alItem(2).ToString()
    End Sub
End Class

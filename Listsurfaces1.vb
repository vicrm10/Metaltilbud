Public Class Listsurfaces1

    Private palSuppliers As ArrayList
    Private WithEvents pcmbTarget As ComboBox

    Public Sub New(ByRef cmbTarget As ComboBox)
        pcmbTarget = cmbTarget
        RefreshData()
    End Sub

    Public Sub RefreshData()
        Dim objDatabaseIO As DatabaseIO


        objDatabaseIO = New DatabaseIO
        palSuppliers = objDatabaseIO.GetSurfacetreatment1()
    End Sub

    Public Sub List()
        Dim alItem As ArrayList

        Dim i As Integer


        pcmbTarget.Items.Clear()
        For i = 0 To palSuppliers.Count - 1
            alItem = palSuppliers(i)
            'pcmbTarget.Items.Add(alItem(0).ToString() + " " + alItem(1).ToString())
            pcmbTarget.Items.Add(alItem(1).ToString())
        Next
        If palSuppliers.Count > 0 Then
            pcmbTarget.SelectedIndex = 0
        End If
    End Sub

    Private Sub pcmbTarget_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pcmbTarget.SelectedIndexChanged
        '        Dim alItem As ArrayList
        '
        '
        '        alItem = palMaterials(pcmbTarget.SelectedIndex)
        '        ptxtMaterialGroup.Text = alItem(3).ToString()
    End Sub

End Class

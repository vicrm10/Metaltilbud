Public Class FilterKeys
    Private WithEvents ptxtTarget As TextBox
    Private WithEvents plblSource As Label

    Private pobjForm As metal_tilbud
    Private pbNonNumberEntered As Boolean = False
    Private pbPeriodPressed As Boolean = False


    Private Sub ptxtTarget_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ptxtTarget.KeyDown
        pbNonNumberEntered = False
        pbPeriodPressed = False
        ' Determine whether the keystroke is a number from the top of the keyboard.
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then
            ' Determine whether the keystroke is a number from the keypad.
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then
                ' Determine whether the keystroke is a backspace.
                If e.KeyCode <> Keys.Back And e.KeyCode <> Keys.Delete Then
                    If e.KeyCode <> Keys.Decimal And e.KeyCode <> Keys.Oemcomma Then
                        If e.KeyCode = Keys.OemPeriod Then
                            pbPeriodPressed = True
                        Else
                            ' A non-numerical keystroke was pressed. 
                            ' Set the flag to true and evaluate in KeyPress event.
                            pbNonNumberEntered = True
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ptxtTarget_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ptxtTarget.KeyPress
        ' Check for the flag being set in the KeyDown event.
        If pbNonNumberEntered = True Then
            ' Stop the character from being entered into the control since it is non-numerical.
            e.Handled = True
        End If
        If pbPeriodPressed = True Then
            e.KeyChar = ","
        End If
        If e.KeyChar = "," Then
            If ptxtTarget.Text.IndexOf(",") > 0 Then
                e.KeyChar = ""
            End If
        End If
    End Sub

    Public Sub New(ByRef txtTarget As TextBox, ByRef lblSource As Label, ByRef objForm As metal_tilbud)
        ptxtTarget = txtTarget
        plblSource = lblSource
        pobjForm = objForm
    End Sub

    Public Function GetValue() As Double
        If plblSource Is Nothing Then
            If ptxtTarget.Text <> "" Then
                GetValue = Convert.ToDouble(ptxtTarget.Text)
            Else
                GetValue = 0
            End If
        Else
            If ptxtTarget.Text = "" Then
                If plblSource.Text <> "" Then
                    GetValue = Convert.ToDouble(plblSource.Text)
                Else
                    GetValue = 0
                End If
            Else
                GetValue = Convert.ToDouble(ptxtTarget.Text)
            End If
        End If
    End Function

    Public Sub ClearValues()
        ptxtTarget.Text = ""
        If Not plblSource Is Nothing Then
            plblSource.Text = ""
        End If
    End Sub

    Private Sub ptxtTarget_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ptxtTarget.KeyUp
        pobjForm.CalculateOrdrestr()
    End Sub
End Class

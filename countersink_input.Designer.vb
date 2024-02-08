<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class countersink_input
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tb_1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tb_4 = New System.Windows.Forms.TextBox
        Me.tb_3 = New System.Windows.Forms.TextBox
        Me.tb_2 = New System.Windows.Forms.TextBox
        Me.bu_ok = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'tb_1
        '
        Me.tb_1.Location = New System.Drawing.Point(165, 60)
        Me.tb_1.Name = "tb_1"
        Me.tb_1.Size = New System.Drawing.Size(79, 20)
        Me.tb_1.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(159, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 27)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Antal pr emne"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tb_4
        '
        Me.tb_4.Location = New System.Drawing.Point(165, 120)
        Me.tb_4.Name = "tb_4"
        Me.tb_4.Size = New System.Drawing.Size(79, 20)
        Me.tb_4.TabIndex = 3
        '
        'tb_3
        '
        Me.tb_3.Location = New System.Drawing.Point(165, 100)
        Me.tb_3.Name = "tb_3"
        Me.tb_3.Size = New System.Drawing.Size(79, 20)
        Me.tb_3.TabIndex = 2
        '
        'tb_2
        '
        Me.tb_2.Location = New System.Drawing.Point(165, 80)
        Me.tb_2.Name = "tb_2"
        Me.tb_2.Size = New System.Drawing.Size(79, 20)
        Me.tb_2.TabIndex = 1
        '
        'bu_ok
        '
        Me.bu_ok.Location = New System.Drawing.Point(198, 154)
        Me.bu_ok.Name = "bu_ok"
        Me.bu_ok.Size = New System.Drawing.Size(78, 20)
        Me.bu_ok.TabIndex = 4
        Me.bu_ok.Text = "ok (Fortsæt)"
        Me.bu_ok.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(40, 79)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 20)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "6-10 mm (M3-M5)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(40, 99)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 20)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "10-15 mm (M6-M8)"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(27, 119)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(117, 20)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "over 15 mm (M10-M12)"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(40, 36)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(134, 20)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "Diameter (Gevind str.)"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(41, 59)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(103, 20)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "4-5 mm (M2-M2.5)"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'countersink_input
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(385, 238)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.bu_ok)
        Me.Controls.Add(Me.tb_2)
        Me.Controls.Add(Me.tb_3)
        Me.Controls.Add(Me.tb_4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tb_1)
        Me.Name = "countersink_input"
        Me.Text = "UNDERSÆNKNING"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tb_1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tb_4 As System.Windows.Forms.TextBox
    Friend WithEvents tb_3 As System.Windows.Forms.TextBox
    Friend WithEvents tb_2 As System.Windows.Forms.TextBox
    Friend WithEvents bu_ok As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPickGame
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.NoGames = New System.Windows.Forms.TextBox
        Me.LblMaxGames = New System.Windows.Forms.Label
        Me.OKButton = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Go to game"
        '
        'NoGames
        '
        Me.NoGames.Location = New System.Drawing.Point(70, 8)
        Me.NoGames.Name = "NoGames"
        Me.NoGames.Size = New System.Drawing.Size(42, 20)
        Me.NoGames.TabIndex = 1
        Me.NoGames.Text = "12345"
        '
        'LblMaxGames
        '
        Me.LblMaxGames.AutoSize = True
        Me.LblMaxGames.Location = New System.Drawing.Point(118, 11)
        Me.LblMaxGames.Name = "LblMaxGames"
        Me.LblMaxGames.Size = New System.Drawing.Size(61, 13)
        Me.LblMaxGames.TabIndex = 2
        Me.LblMaxGames.Text = "of XXXXXX"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(185, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(43, 23)
        Me.OKButton.TabIndex = 3
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'frmPickGame
        '
        Me.AcceptButton = Me.OKButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(234, 31)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.LblMaxGames)
        Me.Controls.Add(Me.NoGames)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPickGame"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Pick game"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents NoGames As System.Windows.Forms.TextBox
    Friend WithEvents LblMaxGames As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
End Class

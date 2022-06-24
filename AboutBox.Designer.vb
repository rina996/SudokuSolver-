<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutBox
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LabelProductName = New System.Windows.Forms.Label
        Me.LabelVersion = New System.Windows.Forms.Label
        Me.LabelCopyright = New System.Windows.Forms.Label
        Me.LabelDescription = New System.Windows.Forms.Label
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'LabelProductName
        '
        Me.LabelProductName.AutoSize = True
        Me.LabelProductName.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelProductName.Location = New System.Drawing.Point(9, 9)
        Me.LabelProductName.Name = "LabelProductName"
        Me.LabelProductName.Padding = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.LabelProductName.Size = New System.Drawing.Size(41, 23)
        Me.LabelProductName.TabIndex = 3
        Me.LabelProductName.Text = "Name"
        '
        'LabelVersion
        '
        Me.LabelVersion.AutoSize = True
        Me.LabelVersion.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelVersion.Location = New System.Drawing.Point(9, 32)
        Me.LabelVersion.Name = "LabelVersion"
        Me.LabelVersion.Padding = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.LabelVersion.Size = New System.Drawing.Size(48, 23)
        Me.LabelVersion.TabIndex = 4
        Me.LabelVersion.Text = "Version"
        '
        'LabelCopyright
        '
        Me.LabelCopyright.AutoSize = True
        Me.LabelCopyright.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelCopyright.Location = New System.Drawing.Point(9, 55)
        Me.LabelCopyright.Name = "LabelCopyright"
        Me.LabelCopyright.Padding = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.LabelCopyright.Size = New System.Drawing.Size(57, 23)
        Me.LabelCopyright.TabIndex = 5
        Me.LabelCopyright.Text = "Copyright"
        '
        'LabelDescription
        '
        Me.LabelDescription.AutoSize = True
        Me.LabelDescription.Dock = System.Windows.Forms.DockStyle.Top
        Me.LabelDescription.Location = New System.Drawing.Point(9, 78)
        Me.LabelDescription.Name = "LabelDescription"
        Me.LabelDescription.Padding = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.LabelDescription.Size = New System.Drawing.Size(66, 23)
        Me.LabelDescription.TabIndex = 7
        Me.LabelDescription.Text = "Description"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.LinkLabel1.Location = New System.Drawing.Point(9, 101)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Padding = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.LinkLabel1.Size = New System.Drawing.Size(327, 23)
        Me.LinkLabel1.TabIndex = 8
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "www.codeproject.com/KB/game/VBSudokuSolverGenerator.aspx"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Location = New System.Drawing.Point(9, 124)
        Me.Label1.Name = "Label1"
        Me.Label1.Padding = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.Label1.Size = New System.Drawing.Size(288, 23)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Click link above for more information or to submit feedback"
        '
        'AboutBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(391, 153)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.LabelDescription)
        Me.Controls.Add(Me.LabelCopyright)
        Me.Controls.Add(Me.LabelVersion)
        Me.Controls.Add(Me.LabelProductName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AboutBox"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "AboutBox"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelProductName As System.Windows.Forms.Label
    Friend WithEvents LabelVersion As System.Windows.Forms.Label
    Friend WithEvents LabelCopyright As System.Windows.Forms.Label
    Friend WithEvents LabelDescription As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class

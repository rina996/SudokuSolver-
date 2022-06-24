<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInitial
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
        Me.components = New System.ComponentModel.Container
        Me.GameType = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.NotifyIconProgram = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PuzzleGenerationOptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NotifyIconGenerator = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.AniTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GameType
        '
        Me.GameType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.GameType.FormattingEnabled = True
        Me.GameType.ItemHeight = 13
        Me.GameType.Items.AddRange(New Object() {"Classic", "Samurai"})
        Me.GameType.Location = New System.Drawing.Point(7, 7)
        Me.GameType.Name = "GameType"
        Me.GameType.Size = New System.Drawing.Size(137, 21)
        Me.GameType.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(149, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(59, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(214, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(59, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Exit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'NotifyIconProgram
        '
        Me.NotifyIconProgram.Visible = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem, Me.PuzzleGenerationOptionsToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(211, 48)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'PuzzleGenerationOptionsToolStripMenuItem
        '
        Me.PuzzleGenerationOptionsToolStripMenuItem.Name = "PuzzleGenerationOptionsToolStripMenuItem"
        Me.PuzzleGenerationOptionsToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.PuzzleGenerationOptionsToolStripMenuItem.Text = "Puzzle generation options"
        '
        'NotifyIconGenerator
        '
        Me.NotifyIconGenerator.Text = "NotifyIcon2"
        Me.NotifyIconGenerator.Visible = True
        '
        'AniTimer
        '
        Me.AniTimer.Enabled = True
        Me.AniTimer.Interval = 250
        '
        'frmInitial
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(279, 36)
        Me.Controls.Add(Me.GameType)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmInitial"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select game type"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GameType As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents NotifyIconProgram As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PuzzleGenerationOptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyIconGenerator As System.Windows.Forms.NotifyIcon
    Friend WithEvents AniTimer As System.Windows.Forms.Timer
End Class

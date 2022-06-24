' Copyright (c) Microsoft Corporation. All rights reserved.
Partial Public Class frmGame
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
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
        Me.GameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ClearBoardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator
        Me.LoadFromFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator
        Me.CopyCluesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CopyFilledToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CopySolutionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SolveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NextStepToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ShowAllStepsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator16 = New System.Windows.Forms.ToolStripSeparator
        Me.CheckConsistencyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator
        Me.SolveUpToToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FindOnlyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ShowSolutionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.UndoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RedoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator
        Me.OptionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExtraToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PuzzleLibraryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator18 = New System.Windows.Forms.ToolStripSeparator
        Me.OptimiseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OptimisePuzzleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OptimisePuzzleSymmetryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BatchOptimiseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator
        Me.BatchSolveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FeedbackBox = New System.Windows.Forms.TextBox
        Me.StatusLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripMultiGameLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.SetupLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.InfoLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.SoundLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.CandidatesLabel = New System.Windows.Forms.ToolStripStatusLabel
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.doubleClickTimer = New System.Windows.Forms.Timer(Me.components)
        Me.RCContextMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EnableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.E1 = New System.Windows.Forms.ToolStripMenuItem
        Me.E2 = New System.Windows.Forms.ToolStripMenuItem
        Me.E3 = New System.Windows.Forms.ToolStripMenuItem
        Me.E4 = New System.Windows.Forms.ToolStripMenuItem
        Me.E5 = New System.Windows.Forms.ToolStripMenuItem
        Me.E6 = New System.Windows.Forms.ToolStripMenuItem
        Me.E7 = New System.Windows.Forms.ToolStripMenuItem
        Me.E8 = New System.Windows.Forms.ToolStripMenuItem
        Me.E9 = New System.Windows.Forms.ToolStripMenuItem
        Me.DisableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.D1 = New System.Windows.Forms.ToolStripMenuItem
        Me.D2 = New System.Windows.Forms.ToolStripMenuItem
        Me.D3 = New System.Windows.Forms.ToolStripMenuItem
        Me.D4 = New System.Windows.Forms.ToolStripMenuItem
        Me.D5 = New System.Windows.Forms.ToolStripMenuItem
        Me.D6 = New System.Windows.Forms.ToolStripMenuItem
        Me.D7 = New System.Windows.Forms.ToolStripMenuItem
        Me.D8 = New System.Windows.Forms.ToolStripMenuItem
        Me.D9 = New System.Windows.Forms.ToolStripMenuItem
        Me.SetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.S1 = New System.Windows.Forms.ToolStripMenuItem
        Me.S2 = New System.Windows.Forms.ToolStripMenuItem
        Me.S3 = New System.Windows.Forms.ToolStripMenuItem
        Me.S4 = New System.Windows.Forms.ToolStripMenuItem
        Me.S5 = New System.Windows.Forms.ToolStripMenuItem
        Me.S6 = New System.Windows.Forms.ToolStripMenuItem
        Me.S7 = New System.Windows.Forms.ToolStripMenuItem
        Me.S8 = New System.Windows.Forms.ToolStripMenuItem
        Me.S9 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator17 = New System.Windows.Forms.ToolStripSeparator
        Me.RevealSolutionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MsgTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CheckBox7 = New System.Windows.Forms.CheckBox
        Me.CheckBox8 = New System.Windows.Forms.CheckBox
        Me.CheckBox9 = New System.Windows.Forms.CheckBox
        Me.MenuStrip = New System.Windows.Forms.MenuStrip
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RCContextMenuStrip.SuspendLayout()
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GameToolStripMenuItem
        '
        Me.GameToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearBoardToolStripMenuItem, Me.ToolStripSeparator6, Me.LoadFromFileToolStripMenuItem, Me.ToolStripSeparator2, Me.PasteToolStripMenuItem, Me.ToolStripSeparator4, Me.CopyCluesToolStripMenuItem, Me.CopyFilledToolStripMenuItem, Me.CopySolutionToolStripMenuItem, Me.ToolStripSeparator3, Me.ExitToolStripMenuItem})
        Me.GameToolStripMenuItem.Name = "GameToolStripMenuItem"
        Me.GameToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.GameToolStripMenuItem.Text = "Game"
        '
        'ClearBoardToolStripMenuItem
        '
        Me.ClearBoardToolStripMenuItem.Name = "ClearBoardToolStripMenuItem"
        Me.ClearBoardToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.ClearBoardToolStripMenuItem.Text = "Clear board"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(171, 6)
        '
        'LoadFromFileToolStripMenuItem
        '
        Me.LoadFromFileToolStripMenuItem.Name = "LoadFromFileToolStripMenuItem"
        Me.LoadFromFileToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.LoadFromFileToolStripMenuItem.Text = "Load from file"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(171, 6)
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.PasteToolStripMenuItem.Text = "Paste"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(171, 6)
        '
        'CopyCluesToolStripMenuItem
        '
        Me.CopyCluesToolStripMenuItem.Name = "CopyCluesToolStripMenuItem"
        Me.CopyCluesToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopyCluesToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.CopyCluesToolStripMenuItem.Text = "Copy clues"
        '
        'CopyFilledToolStripMenuItem
        '
        Me.CopyFilledToolStripMenuItem.Name = "CopyFilledToolStripMenuItem"
        Me.CopyFilledToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.CopyFilledToolStripMenuItem.Text = "Copy filled"
        '
        'CopySolutionToolStripMenuItem
        '
        Me.CopySolutionToolStripMenuItem.Name = "CopySolutionToolStripMenuItem"
        Me.CopySolutionToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.CopySolutionToolStripMenuItem.Text = "Copy solution"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(171, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(174, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'SolveToolStripMenuItem
        '
        Me.SolveToolStripMenuItem.BackColor = System.Drawing.Color.Transparent
        Me.SolveToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HintToolStripMenuItem, Me.NextStepToolStripMenuItem, Me.ShowAllStepsToolStripMenuItem, Me.ToolStripSeparator16, Me.CheckConsistencyToolStripMenuItem, Me.ToolStripSeparator8, Me.SolveUpToToolStripMenuItem, Me.FindOnlyToolStripMenuItem, Me.ShowSolutionToolStripMenuItem, Me.ToolStripSeparator1, Me.UndoToolStripMenuItem, Me.RedoToolStripMenuItem, Me.ToolStripSeparator5})
        Me.SolveToolStripMenuItem.Name = "SolveToolStripMenuItem"
        Me.SolveToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.SolveToolStripMenuItem.Text = "Solve"
        '
        'HintToolStripMenuItem
        '
        Me.HintToolStripMenuItem.Name = "HintToolStripMenuItem"
        Me.HintToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
        Me.HintToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.HintToolStripMenuItem.Text = "Hint"
        '
        'NextStepToolStripMenuItem
        '
        Me.NextStepToolStripMenuItem.Name = "NextStepToolStripMenuItem"
        Me.NextStepToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.NextStepToolStripMenuItem.Text = "Next step"
        '
        'ShowAllStepsToolStripMenuItem
        '
        Me.ShowAllStepsToolStripMenuItem.Name = "ShowAllStepsToolStripMenuItem"
        Me.ShowAllStepsToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.ShowAllStepsToolStripMenuItem.Text = "Show all steps"
        '
        'ToolStripSeparator16
        '
        Me.ToolStripSeparator16.Name = "ToolStripSeparator16"
        Me.ToolStripSeparator16.Size = New System.Drawing.Size(169, 6)
        '
        'CheckConsistencyToolStripMenuItem
        '
        Me.CheckConsistencyToolStripMenuItem.Name = "CheckConsistencyToolStripMenuItem"
        Me.CheckConsistencyToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.CheckConsistencyToolStripMenuItem.Text = "Check consistency"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(169, 6)
        '
        'SolveUpToToolStripMenuItem
        '
        Me.SolveUpToToolStripMenuItem.Name = "SolveUpToToolStripMenuItem"
        Me.SolveUpToToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.SolveUpToToolStripMenuItem.Text = "Solve up to:"
        '
        'FindOnlyToolStripMenuItem
        '
        Me.FindOnlyToolStripMenuItem.Name = "FindOnlyToolStripMenuItem"
        Me.FindOnlyToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.FindOnlyToolStripMenuItem.Text = "Find only:"
        '
        'ShowSolutionToolStripMenuItem
        '
        Me.ShowSolutionToolStripMenuItem.Name = "ShowSolutionToolStripMenuItem"
        Me.ShowSolutionToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.ShowSolutionToolStripMenuItem.Text = "Show solution"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(169, 6)
        '
        'UndoToolStripMenuItem
        '
        Me.UndoToolStripMenuItem.Name = "UndoToolStripMenuItem"
        Me.UndoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.UndoToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.UndoToolStripMenuItem.Text = "Undo"
        '
        'RedoToolStripMenuItem
        '
        Me.RedoToolStripMenuItem.Name = "RedoToolStripMenuItem"
        Me.RedoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
        Me.RedoToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.RedoToolStripMenuItem.Text = "Redo"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(169, 6)
        '
        'OptionToolStripMenuItem
        '
        Me.OptionToolStripMenuItem.Name = "OptionToolStripMenuItem"
        Me.OptionToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.OptionToolStripMenuItem.Text = "Options"
        '
        'ExtraToolStripMenuItem
        '
        Me.ExtraToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PuzzleLibraryToolStripMenuItem, Me.ToolStripSeparator18, Me.OptimiseToolStripMenuItem, Me.ToolStripSeparator7, Me.BatchSolveToolStripMenuItem})
        Me.ExtraToolStripMenuItem.Name = "ExtraToolStripMenuItem"
        Me.ExtraToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.ExtraToolStripMenuItem.Text = "Extra"
        '
        'PuzzleLibraryToolStripMenuItem
        '
        Me.PuzzleLibraryToolStripMenuItem.Name = "PuzzleLibraryToolStripMenuItem"
        Me.PuzzleLibraryToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.PuzzleLibraryToolStripMenuItem.Text = "Puzzle library"
        '
        'ToolStripSeparator18
        '
        Me.ToolStripSeparator18.Name = "ToolStripSeparator18"
        Me.ToolStripSeparator18.Size = New System.Drawing.Size(140, 6)
        '
        'OptimiseToolStripMenuItem
        '
        Me.OptimiseToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OptimisePuzzleToolStripMenuItem, Me.OptimisePuzzleSymmetryToolStripMenuItem, Me.BatchOptimiseToolStripMenuItem})
        Me.OptimiseToolStripMenuItem.Name = "OptimiseToolStripMenuItem"
        Me.OptimiseToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.OptimiseToolStripMenuItem.Text = "Optimise"
        '
        'OptimisePuzzleToolStripMenuItem
        '
        Me.OptimisePuzzleToolStripMenuItem.Name = "OptimisePuzzleToolStripMenuItem"
        Me.OptimisePuzzleToolStripMenuItem.Size = New System.Drawing.Size(272, 22)
        Me.OptimisePuzzleToolStripMenuItem.Text = "Optimise puzzle"
        '
        'OptimisePuzzleSymmetryToolStripMenuItem
        '
        Me.OptimisePuzzleSymmetryToolStripMenuItem.Name = "OptimisePuzzleSymmetryToolStripMenuItem"
        Me.OptimisePuzzleSymmetryToolStripMenuItem.Size = New System.Drawing.Size(272, 22)
        Me.OptimisePuzzleSymmetryToolStripMenuItem.Text = "Optimise puzzle (maintain symmetry)"
        '
        'BatchOptimiseToolStripMenuItem
        '
        Me.BatchOptimiseToolStripMenuItem.Name = "BatchOptimiseToolStripMenuItem"
        Me.BatchOptimiseToolStripMenuItem.Size = New System.Drawing.Size(272, 22)
        Me.BatchOptimiseToolStripMenuItem.Text = "Batch optimise"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(140, 6)
        '
        'BatchSolveToolStripMenuItem
        '
        Me.BatchSolveToolStripMenuItem.Name = "BatchSolveToolStripMenuItem"
        Me.BatchSolveToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.BatchSolveToolStripMenuItem.Text = "Batch solve"
        '
        'FeedbackBox
        '
        Me.FeedbackBox.CausesValidation = False
        Me.FeedbackBox.HideSelection = False
        Me.FeedbackBox.Location = New System.Drawing.Point(220, 27)
        Me.FeedbackBox.Multiline = True
        Me.FeedbackBox.Name = "FeedbackBox"
        Me.FeedbackBox.ReadOnly = True
        Me.FeedbackBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.FeedbackBox.Size = New System.Drawing.Size(126, 242)
        Me.FeedbackBox.TabIndex = 23
        '
        'StatusLabel
        '
        Me.StatusLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.StatusLabel.Name = "StatusLabel"
        Me.StatusLabel.Size = New System.Drawing.Size(203, 17)
        Me.StatusLabel.Spring = True
        Me.StatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripMultiGameLabel
        '
        Me.ToolStripMultiGameLabel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripMultiGameLabel.Name = "ToolStripMultiGameLabel"
        Me.ToolStripMultiGameLabel.Size = New System.Drawing.Size(71, 17)
        Me.ToolStripMultiGameLabel.Tag = "Double click to select a different puzzle"
        Me.ToolStripMultiGameLabel.Text = "Puzzle x of y"
        Me.ToolStripMultiGameLabel.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'SetupLabel
        '
        Me.SetupLabel.DoubleClickEnabled = True
        Me.SetupLabel.Name = "SetupLabel"
        Me.SetupLabel.Size = New System.Drawing.Size(71, 17)
        Me.SetupLabel.Text = "Setup mode"
        '
        'InfoLabel
        '
        Me.InfoLabel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.InfoLabel.Image = Global.SudokuSolver.My.Resources.Resources.info
        Me.InfoLabel.Name = "InfoLabel"
        Me.InfoLabel.Padding = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.InfoLabel.Size = New System.Drawing.Size(20, 17)
        '
        'SoundLabel
        '
        Me.SoundLabel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SoundLabel.DoubleClickEnabled = True
        Me.SoundLabel.Image = Global.SudokuSolver.My.Resources.Resources.sound
        Me.SoundLabel.Name = "SoundLabel"
        Me.SoundLabel.Padding = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.SoundLabel.Size = New System.Drawing.Size(20, 17)
        '
        'CandidatesLabel
        '
        Me.CandidatesLabel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CandidatesLabel.DoubleClickEnabled = True
        Me.CandidatesLabel.Image = Global.SudokuSolver.My.Resources.Resources.candidates_add
        Me.CandidatesLabel.Name = "CandidatesLabel"
        Me.CandidatesLabel.Padding = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.CandidatesLabel.Size = New System.Drawing.Size(20, 17)
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "sdm"
        Me.OpenFileDialog1.Filter = "Sudoku files (*.sdm)|*.sdm|Text files (*.txt)|*.txt"
        Me.OpenFileDialog1.Multiselect = True
        Me.OpenFileDialog1.RestoreDirectory = True
        Me.OpenFileDialog1.Title = "Open sudoku file"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(12, 27)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox1.TabIndex = 25
        Me.CheckBox1.Tag = "1"
        Me.CheckBox1.Text = "1"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox2.Location = New System.Drawing.Point(12, 49)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox2.TabIndex = 26
        Me.CheckBox2.Tag = "2"
        Me.CheckBox2.Text = "2"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox3.Location = New System.Drawing.Point(12, 71)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox3.TabIndex = 27
        Me.CheckBox3.Tag = "3"
        Me.CheckBox3.Text = "3"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox4.Location = New System.Drawing.Point(12, 93)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox4.TabIndex = 28
        Me.CheckBox4.Tag = "4"
        Me.CheckBox4.Text = "4"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox5.Location = New System.Drawing.Point(12, 115)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox5.TabIndex = 29
        Me.CheckBox5.Tag = "5"
        Me.CheckBox5.Text = "5"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox6.Location = New System.Drawing.Point(12, 137)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox6.TabIndex = 30
        Me.CheckBox6.Tag = "6"
        Me.CheckBox6.Text = "6"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'doubleClickTimer
        '
        '
        'RCContextMenuStrip
        '
        Me.RCContextMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EnableToolStripMenuItem, Me.DisableToolStripMenuItem, Me.SetToolStripMenuItem, Me.ToolStripSeparator17, Me.RevealSolutionToolStripMenuItem})
        Me.RCContextMenuStrip.Name = "RCContextMenuStrip"
        Me.RCContextMenuStrip.Size = New System.Drawing.Size(155, 98)
        '
        'EnableToolStripMenuItem
        '
        Me.EnableToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.E1, Me.E2, Me.E3, Me.E4, Me.E5, Me.E6, Me.E7, Me.E8, Me.E9})
        Me.EnableToolStripMenuItem.Name = "EnableToolStripMenuItem"
        Me.EnableToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.EnableToolStripMenuItem.Text = "Enable"
        '
        'E1
        '
        Me.E1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E1.Name = "E1"
        Me.E1.Size = New System.Drawing.Size(80, 22)
        Me.E1.Text = "1"
        '
        'E2
        '
        Me.E2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E2.Name = "E2"
        Me.E2.Size = New System.Drawing.Size(80, 22)
        Me.E2.Text = "2"
        '
        'E3
        '
        Me.E3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E3.Name = "E3"
        Me.E3.Size = New System.Drawing.Size(80, 22)
        Me.E3.Text = "3"
        '
        'E4
        '
        Me.E4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E4.Name = "E4"
        Me.E4.Size = New System.Drawing.Size(80, 22)
        Me.E4.Text = "4"
        '
        'E5
        '
        Me.E5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E5.Name = "E5"
        Me.E5.Size = New System.Drawing.Size(80, 22)
        Me.E5.Text = "5"
        '
        'E6
        '
        Me.E6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E6.Name = "E6"
        Me.E6.Size = New System.Drawing.Size(80, 22)
        Me.E6.Text = "6"
        '
        'E7
        '
        Me.E7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E7.Name = "E7"
        Me.E7.Size = New System.Drawing.Size(80, 22)
        Me.E7.Text = "7"
        '
        'E8
        '
        Me.E8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E8.Name = "E8"
        Me.E8.Size = New System.Drawing.Size(80, 22)
        Me.E8.Text = "8"
        '
        'E9
        '
        Me.E9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.E9.Name = "E9"
        Me.E9.Size = New System.Drawing.Size(80, 22)
        Me.E9.Text = "9"
        '
        'DisableToolStripMenuItem
        '
        Me.DisableToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.D1, Me.D2, Me.D3, Me.D4, Me.D5, Me.D6, Me.D7, Me.D8, Me.D9})
        Me.DisableToolStripMenuItem.Name = "DisableToolStripMenuItem"
        Me.DisableToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.DisableToolStripMenuItem.Text = "Disable"
        '
        'D1
        '
        Me.D1.Name = "D1"
        Me.D1.Size = New System.Drawing.Size(80, 22)
        Me.D1.Text = "1"
        '
        'D2
        '
        Me.D2.Name = "D2"
        Me.D2.Size = New System.Drawing.Size(80, 22)
        Me.D2.Text = "2"
        '
        'D3
        '
        Me.D3.Name = "D3"
        Me.D3.Size = New System.Drawing.Size(80, 22)
        Me.D3.Text = "3"
        '
        'D4
        '
        Me.D4.Name = "D4"
        Me.D4.Size = New System.Drawing.Size(80, 22)
        Me.D4.Text = "4"
        '
        'D5
        '
        Me.D5.Name = "D5"
        Me.D5.Size = New System.Drawing.Size(80, 22)
        Me.D5.Text = "5"
        '
        'D6
        '
        Me.D6.Name = "D6"
        Me.D6.Size = New System.Drawing.Size(80, 22)
        Me.D6.Text = "6"
        '
        'D7
        '
        Me.D7.Name = "D7"
        Me.D7.Size = New System.Drawing.Size(80, 22)
        Me.D7.Text = "7"
        '
        'D8
        '
        Me.D8.Name = "D8"
        Me.D8.Size = New System.Drawing.Size(80, 22)
        Me.D8.Text = "8"
        '
        'D9
        '
        Me.D9.Name = "D9"
        Me.D9.Size = New System.Drawing.Size(80, 22)
        Me.D9.Text = "9"
        '
        'SetToolStripMenuItem
        '
        Me.SetToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.S1, Me.S2, Me.S3, Me.S4, Me.S5, Me.S6, Me.S7, Me.S8, Me.S9})
        Me.SetToolStripMenuItem.Name = "SetToolStripMenuItem"
        Me.SetToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.SetToolStripMenuItem.Text = "Set"
        '
        'S1
        '
        Me.S1.Name = "S1"
        Me.S1.Size = New System.Drawing.Size(80, 22)
        Me.S1.Text = "1"
        '
        'S2
        '
        Me.S2.Name = "S2"
        Me.S2.Size = New System.Drawing.Size(80, 22)
        Me.S2.Text = "2"
        '
        'S3
        '
        Me.S3.Name = "S3"
        Me.S3.Size = New System.Drawing.Size(80, 22)
        Me.S3.Text = "3"
        '
        'S4
        '
        Me.S4.Name = "S4"
        Me.S4.Size = New System.Drawing.Size(80, 22)
        Me.S4.Text = "4"
        '
        'S5
        '
        Me.S5.Name = "S5"
        Me.S5.Size = New System.Drawing.Size(80, 22)
        Me.S5.Text = "5"
        '
        'S6
        '
        Me.S6.Name = "S6"
        Me.S6.Size = New System.Drawing.Size(80, 22)
        Me.S6.Text = "6"
        '
        'S7
        '
        Me.S7.Name = "S7"
        Me.S7.Size = New System.Drawing.Size(80, 22)
        Me.S7.Text = "7"
        '
        'S8
        '
        Me.S8.Name = "S8"
        Me.S8.Size = New System.Drawing.Size(80, 22)
        Me.S8.Text = "8"
        '
        'S9
        '
        Me.S9.Name = "S9"
        Me.S9.Size = New System.Drawing.Size(80, 22)
        Me.S9.Text = "9"
        '
        'ToolStripSeparator17
        '
        Me.ToolStripSeparator17.Name = "ToolStripSeparator17"
        Me.ToolStripSeparator17.Size = New System.Drawing.Size(151, 6)
        '
        'RevealSolutionToolStripMenuItem
        '
        Me.RevealSolutionToolStripMenuItem.Name = "RevealSolutionToolStripMenuItem"
        Me.RevealSolutionToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.RevealSolutionToolStripMenuItem.Text = "Reveal solution"
        '
        'MsgTimer
        '
        Me.MsgTimer.Interval = 4000
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox7.Location = New System.Drawing.Point(12, 159)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox7.TabIndex = 32
        Me.CheckBox7.Tag = "7"
        Me.CheckBox7.Text = "7"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'CheckBox8
        '
        Me.CheckBox8.AutoSize = True
        Me.CheckBox8.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox8.Location = New System.Drawing.Point(12, 181)
        Me.CheckBox8.Name = "CheckBox8"
        Me.CheckBox8.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox8.TabIndex = 33
        Me.CheckBox8.Tag = "8"
        Me.CheckBox8.Text = "8"
        Me.CheckBox8.UseVisualStyleBackColor = True
        '
        'CheckBox9
        '
        Me.CheckBox9.AutoSize = True
        Me.CheckBox9.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CheckBox9.Location = New System.Drawing.Point(12, 203)
        Me.CheckBox9.Name = "CheckBox9"
        Me.CheckBox9.Size = New System.Drawing.Size(32, 17)
        Me.CheckBox9.TabIndex = 34
        Me.CheckBox9.Tag = "9"
        Me.CheckBox9.Text = "9"
        Me.CheckBox9.UseVisualStyleBackColor = True
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GameToolStripMenuItem, Me.SolveToolStripMenuItem, Me.OptionToolStripMenuItem, Me.ExtraToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(420, 24)
        Me.MenuStrip.TabIndex = 35
        Me.MenuStrip.Text = "MenuStrip2"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel, Me.ToolStripMultiGameLabel, Me.SetupLabel, Me.InfoLabel, Me.SoundLabel, Me.CandidatesLabel})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 306)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.ShowItemToolTips = True
        Me.StatusStrip1.Size = New System.Drawing.Size(420, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 36
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(52, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'frmGame
        '
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(420, 328)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.CheckBox9)
        Me.Controls.Add(Me.CheckBox8)
        Me.Controls.Add(Me.CheckBox7)
        Me.Controls.Add(Me.CheckBox6)
        Me.Controls.Add(Me.CheckBox5)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.FeedbackBox)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmGame"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Solver"
        Me.RCContextMenuStrip.ResumeLayout(False)
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GameToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FeedbackBox As System.Windows.Forms.TextBox
    Friend WithEvents ToolStripMultiGameLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents LoadFromFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents doubleClickTimer As System.Windows.Forms.Timer
    Friend WithEvents SetupLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SolveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UndoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RedoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ShowSolutionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents RCContextMenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EnableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents E9 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DisableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents D9 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S9 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MsgTimer As System.Windows.Forms.Timer
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents OptionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckBox7 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox8 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox9 As System.Windows.Forms.CheckBox
    Friend WithEvents HintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CopyCluesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyFilledToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopySolutionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CheckConsistencyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SolveUpToToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SoundLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents InfoLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ExtraToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptimiseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearBoardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BatchSolveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents OptimisePuzzleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptimisePuzzleSymmetryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BatchOptimiseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NextStepToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CandidatesLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator10 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator12 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem9 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem10 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem11 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem12 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator13 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem13 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem14 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator14 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem15 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem16 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem17 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem18 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem19 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem20 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem21 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem22 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator15 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripMenuItem23 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel5 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ShowAllStepsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator16 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents FindOnlyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator17 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents RevealSolutionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PuzzleLibraryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator18 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class

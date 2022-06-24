Option Strict Off
Imports System.ComponentModel
Imports System.Threading
Public Class frmBatch
    '---for reporting on background thread operations--
    Delegate Sub updateProgress()
    Delegate Sub batchComplete()
    Delegate Sub showResults()
    Private batchNames() As String
    Private batchFile As String = ""
    Private gArray(0) As String
    Private rArray(0) As String
    Private intPercent As Integer
    Private strOutput As String
    Private totalTime As Integer
    Private intPuzzles As Integer
    Private strDifficulty As String
    Private intCurrentPuzzle As Integer
    Dim executionTimeWatch As New Stopwatch()
    Dim puzzleTimeWatch As New Stopwatch()
    Public Event BatchFinished()
    Public Sub batch()
        StopThreadWork = False

        '---start the puzzle timer---
        executionTimeWatch.Start()

        Dim intPuzzleTime As Integer
        Dim i As Integer

        intPuzzles = rArray.Length
        intCurrentPuzzle = 0

        For i = 0 To rArray.Length - 1
            Dim Solver As New clsSudokuSolver
            If RadioButton1.Checked = True Then
                Solver.blnClassic = True
            Else
                Solver.blnClassic = False
            End If
            intCurrentPuzzle += 1
            If rArray(i) = "-" Then
                Solver.strGrid = Replace(gArray(i), vbCrLf, "")
                Solver.vsSolvers = My.Settings._UniqueSolvers
                puzzleTimeWatch.Start()
                Solver._vsUnique()
                '---re-run to get solver methods---
                If cb_Methods.Checked AndAlso Solver.intCountSolutions = 1 Then
                    If cb_Default.Checked Then Solver.vsSolvers = My.Settings._defaultSolvers
                    If cb_Enabled.Checked Then Solver.vsSolvers = My.Settings._EnabledSolvers
                    Solver.strInputGameSolution = Solver.strGameSolution
                    Solver._vsUnique()
                    strDifficulty = Solver.strDifficulty
                Else
                    Solver.strInputGameSolution = ""
                End If
                intPuzzleTime = puzzleTimeWatch.ElapsedMilliseconds
                puzzleTimeWatch.Reset()

                Dim strMethods As String = ""
                Dim sm() As String = Solver.solveMethods
                Dim sc() As Integer = Solver.solveCountMethods
                Dim j As Integer
                For j = 0 To UBound(sm)
                    If sm(j) <> "" Then
                        If strMethods = "" Then
                            strMethods = sm(j) & " [" & sc(j) & "]"
                        Else
                            strMethods = strMethods & ", " & sm(j) & " [" & sc(j) & "]"
                        End If
                    End If
                Next
                rArray(i) = Solver.intCountSolutions & "/" & gArray(i) & "/" & Solver.strGameSolution & "/" & strMethods & "/" & intPuzzleTime & "/" & strDifficulty
            Else
                '---input is invalid---
                rArray(i) = 0 & "/" & gArray(i) & "/" & "" & "/" & "" & "/" & "0" & "/"
            End If
            intPercent = 100 * ((i + 1) / rArray.Length)
            '---update progress bar---
            updateBatchProgress()
            If StopThreadWork Then Exit For
        Next

        '---stop the timer---
        executionTimeWatch.Stop()
        totalTime = executionTimeWatch.ElapsedMilliseconds
        executionTimeWatch.Reset()

        'Select Case StopThreadWork
        '   Case True
        'Debug.Print("batch thread cancelled")
        '   Case False
        'Debug.Print("batch thread ended")
        'End Select

        RaiseEvent BatchFinished()

    End Sub
    Private Sub updateBatchProgress()
        If Me.ProgressBar.InvokeRequired Then
            Dim newDelegate As New updateProgress(AddressOf updateBatchProgress)
            Me.ProgressBar.Invoke(newDelegate)
            Exit Sub
        Else
            Me.ProgressBar.Value = intPercent
            Me.FileLabel.Text = batchFile & " (puzzle " & intCurrentPuzzle & " of " & intPuzzles & ")"
        End If
    End Sub
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ProgressBar = New System.Windows.Forms.ProgressBar
        Me.btnLoad = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.cb_Clues = New System.Windows.Forms.CheckBox
        Me.cb_Solution = New System.Windows.Forms.CheckBox
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.btnExit = New System.Windows.Forms.Button
        Me.FileLabel = New System.Windows.Forms.Label
        Me.cb_Methods = New System.Windows.Forms.CheckBox
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.cb_Enabled = New System.Windows.Forms.CheckBox
        Me.cb_Default = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(29, 65)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(200, 23)
        Me.ProgressBar.Step = 1
        Me.ProgressBar.TabIndex = 0
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(13, 4)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 1
        Me.btnLoad.Text = "Load file"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(94, 3)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 2
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(175, 3)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'cb_Clues
        '
        Me.cb_Clues.AutoSize = True
        Me.cb_Clues.Location = New System.Drawing.Point(18, 104)
        Me.cb_Clues.Name = "cb_Clues"
        Me.cb_Clues.Size = New System.Drawing.Size(123, 17)
        Me.cb_Clues.TabIndex = 4
        Me.cb_Clues.Text = "Output starting clues"
        Me.ToolTip.SetToolTip(Me.cb_Clues, "Shows starting clues for each puzzle in the batch")
        Me.cb_Clues.UseVisualStyleBackColor = True
        '
        'cb_Solution
        '
        Me.cb_Solution.AutoSize = True
        Me.cb_Solution.Location = New System.Drawing.Point(18, 129)
        Me.cb_Solution.Name = "cb_Solution"
        Me.cb_Solution.Size = New System.Drawing.Size(97, 17)
        Me.cb_Solution.TabIndex = 5
        Me.cb_Solution.Text = "Output solution"
        Me.ToolTip.SetToolTip(Me.cb_Solution, "Show the solution to puzzles that have a unique solution")
        Me.cb_Solution.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.Filter = "Sudoku files (*.sdm)|*.sdm"
        Me.OpenFileDialog.Multiselect = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(176, 104)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(58, 17)
        Me.RadioButton1.TabIndex = 6
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Classic"
        Me.ToolTip.SetToolTip(Me.RadioButton1, "Select if input is classic 9x9 sudoku")
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(175, 128)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(63, 17)
        Me.RadioButton2.TabIndex = 7
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Samurai"
        Me.ToolTip.SetToolTip(Me.RadioButton2, "Select if input is a samurai sudoku (5 overlapping grids)")
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(165, 193)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'FileLabel
        '
        Me.FileLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.FileLabel.Location = New System.Drawing.Point(13, 34)
        Me.FileLabel.Name = "FileLabel"
        Me.FileLabel.Size = New System.Drawing.Size(232, 23)
        Me.FileLabel.TabIndex = 9
        Me.FileLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_Methods
        '
        Me.cb_Methods.AutoSize = True
        Me.cb_Methods.Location = New System.Drawing.Point(18, 152)
        Me.cb_Methods.Name = "cb_Methods"
        Me.cb_Methods.Size = New System.Drawing.Size(134, 17)
        Me.cb_Methods.TabIndex = 10
        Me.cb_Methods.Text = "Ouput solving methods"
        Me.ToolTip.SetToolTip(Me.cb_Methods, "Output methods used to solve the puzzle - this option may be slow")
        Me.cb_Methods.UseVisualStyleBackColor = True
        '
        'cb_Enabled
        '
        Me.cb_Enabled.AutoSize = True
        Me.cb_Enabled.Location = New System.Drawing.Point(35, 199)
        Me.cb_Enabled.Name = "cb_Enabled"
        Me.cb_Enabled.Size = New System.Drawing.Size(111, 17)
        Me.cb_Enabled.TabIndex = 11
        Me.cb_Enabled.Text = "Enabled solver list"
        Me.ToolTip.SetToolTip(Me.cb_Enabled, "Uses the order of methods currently enabled by the user")
        Me.cb_Enabled.UseVisualStyleBackColor = True
        '
        'cb_Default
        '
        Me.cb_Default.AutoSize = True
        Me.cb_Default.Checked = True
        Me.cb_Default.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cb_Default.Location = New System.Drawing.Point(35, 176)
        Me.cb_Default.Name = "cb_Default"
        Me.cb_Default.Size = New System.Drawing.Size(106, 17)
        Me.cb_Default.TabIndex = 12
        Me.cb_Default.Text = "Default solver list"
        Me.ToolTip.SetToolTip(Me.cb_Default, "Uses the default solver list and order")
        Me.cb_Default.UseVisualStyleBackColor = True
        '
        'frmBatch
        '
        Me.ClientSize = New System.Drawing.Size(258, 222)
        Me.ControlBox = False
        Me.Controls.Add(Me.cb_Default)
        Me.Controls.Add(Me.cb_Enabled)
        Me.Controls.Add(Me.cb_Methods)
        Me.Controls.Add(Me.FileLabel)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.cb_Solution)
        Me.Controls.Add(Me.cb_Clues)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.ProgressBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBatch"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private Sub frmBatch_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        RemoveHandler BatchFinished, AddressOf batchCompleted
    End Sub
    Private Sub frmBatch_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.FileLabel.Text = "Select file to start..."
        Me.btnStart.Enabled = False
        Me.btnCancel.Enabled = False
        Me.btnLoad.Enabled = True
        Me.btnExit.Enabled = True
        Me.ProgressBar.Value = 0
        Me.cb_Clues.Checked = True
        Me.cb_Solution.Checked = True
        Me.cb_Methods.Checked = True
        If blnSamurai Then
            RadioButton1.Checked = False
            RadioButton2.Checked = True
        Else
            RadioButton1.Checked = True
            RadioButton2.Checked = False
        End If
    End Sub
    Private Sub btnLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Dim strSplit() As String
        If OpenFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            batchNames = OpenFileDialog.FileNames
            If UBound(batchNames) > 0 Then
                batchFile = "Multiple files"
            Else
                batchFile = batchNames(0)
                strSplit = Split(batchFile, "\")
                batchFile = strSplit(strSplit.Length - 1)
            End If
            If Not importBatchSudoku() Then Exit Sub
            Me.btnStart.Enabled = True
            Me.FileLabel.Text = batchFile
            Me.ProgressBar.Value = 0
            Me.RadioButton1.Enabled = False
            Me.RadioButton2.Enabled = False
        End If
    End Sub
    Public Function importBatchSudoku() As Boolean
        Dim i As Integer
        Dim strLength As Integer = 0
        Dim blnLength As Boolean = False
        Dim strError As String = ""
        Dim strInput As String = ""
        Dim sg As Integer = 0
        Dim gc As Integer
        Dim blnValid As Boolean = True
        Dim errMsg As String = ""
        Dim tempArr() As String
        ReDim gArray(0)
        ReDim rArray(0)

        Dim f As Integer
        For f = 0 To UBound(batchNames)

            strInput = readinInput(My.Computer.FileSystem.ReadAllText(batchNames(f)))
            strLength = Len(strInput)
            Select Case RadioButton2.Checked
                Case True
                    If strLength >= 413 Then blnLength = True
                Case False
                    If strLength >= 81 Then blnLength = True
            End Select
            If blnLength Then
                '---process into array
                tempArr = Split(strInput, vbCrLf)
                If RadioButton2.Checked Then
                    Dim tempGrid As String = ""
                    For i = 0 To UBound(tempArr)
                        tempGrid += tempArr(i) & vbCrLf
                        sg += 1
                        If blnValid Then blnValid = blnValidGrid(strGrid:=tempArr(i), strErrMsg:=errMsg, intSamuraiGrid:=sg)
                        If (i + 1) Mod 5 = 0 Then
                            If blnValid And Not blnValidOverlaps(Replace(tempGrid, vbCrLf, "")) Then
                                blnValid = False
                                '---Input is invalid as values for overlapping regions don't match---
                                ReDim Preserve gArray(gc)
                                ReDim Preserve rArray(gc)
                                gArray(gc) = tempGrid
                                rArray(gc) = "0/Input is invalid: as entered values for overlapping regions of the samurai grid don't match.//"
                                gc += 1
                                errMsg = ""
                            Else
                                If blnValid Then
                                    '---add to array---
                                    ReDim Preserve gArray(gc)
                                    ReDim Preserve rArray(gc)
                                    gArray(gc) = tempGrid
                                    rArray(gc) = "-"
                                    gc += 1
                                Else
                                    '---error---
                                    ReDim Preserve gArray(gc)
                                    ReDim Preserve rArray(gc)
                                    gArray(gc) = tempGrid
                                    rArray(gc) = "0/Input is invalid: " & errMsg & "//"
                                    gc += 1
                                    errMsg = ""
                                End If
                            End If
                            tempGrid = ""
                            blnValid = True
                            sg = 0
                        End If
                    Next
                Else
                    For i = 0 To UBound(tempArr)
                        blnValid = blnValidGrid(strGrid:=tempArr(i), strErrMsg:=errMsg)
                        If blnValid Then
                            '---add to array---
                            ReDim Preserve gArray(gc)
                            ReDim Preserve rArray(gc)
                            gArray(gc) = tempArr(i)
                            rArray(gc) = "-"
                            gc += 1
                        Else
                            '---error---
                            ReDim Preserve gArray(gc)
                            ReDim Preserve rArray(gc)
                            gArray(gc) = tempArr(i)
                            rArray(gc) = "0/Input is invalid: " & errMsg & "//"
                            gc += 1
                            errMsg = ""
                        End If
                    Next
                End If
                '---end process into array
            End If
        Next

        If gc > 0 Then
            Return True
        End If

    End Function
    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click

        Me.btnCancel.Enabled = True
        Me.btnStart.Enabled = False
        Me.btnLoad.Enabled = False
        Me.btnExit.Enabled = False
        Me.ProgressBar.Value = 0

        Dim b As New Thread(AddressOf batch)
        b.IsBackground = True
        b.Start()

    End Sub
    Private Sub batchCompleted()
        strOutput = outputResults(blnCancel:=StopThreadWork)
        displayoutput()
        strOutput = ""
    End Sub
    Private Sub displayoutput()
        If Me.InvokeRequired Then
            Dim newDelegate As New showResults(AddressOf displayoutput)
            Me.Invoke(newDelegate)
            Exit Sub
        Else
            PlaySound(My.Resources.batch)
            frmBatchOutput.RichTextBox1.Text = strOutput
            frmBatchOutput.Text = "Results for " & batchFile
            frmBatchOutput.ShowDialog(Me)

            Me.btnCancel.Enabled = False
            Me.btnStart.Enabled = False
            Me.btnLoad.Enabled = True
            Me.btnExit.Enabled = True
            Me.RadioButton1.Enabled = True
            Me.RadioButton2.Enabled = True
            Me.FileLabel.Text = "Select file to start..."
            StopThreadWork = False
        End If
    End Sub
    Private Function outputResults(Optional ByVal blnCancel As Boolean = False) As String
        Dim i As Integer
        Dim strClues As String
        Dim strSolution As String
        Dim intCountSolutions As Integer
        Dim strMilliseconds As String
        Dim strMethods As String
        Dim arrSplit() As String
        Dim avgTime As Long
        Dim strMsg As String
        For i = 0 To rArray.Length - 1
            If rArray(i) = "-" Then Exit For
            arrSplit = Split(rArray(i), "/")
            intCountSolutions = arrSplit(0)
            strClues = arrSplit(1)
            strSolution = arrSplit(2)
            strMethods = arrSplit(3)
            strMilliseconds = arrSplit(4)
            strDifficulty = arrSplit(5)
            strOutput += "Puzzle " & i + 1 & " of " & rArray.Length & ": "
            Select Case intCountSolutions
                Case 0
                    strOutput += "is invalid - no solution - took " & strSecondsMilli(strMilliseconds) & " to process./"
                    If cb_Clues.Checked Then
                        strOutput += strClues & "/"
                    End If
                Case 1
                    strOutput += "has a unique solution - took " & strSecondsMilli(strMilliseconds) & " to solve."
                    If cb_Methods.Checked Then
                        strOutput += " Rated " & LCase(strDifficulty) & "/"
                    Else
                        strOutput += "/"
                    End If
                    If cb_Clues.Checked Then
                        strOutput += strClues & "/"
                    End If
                    If cb_Solution.Checked Then
                        strOutput += strSolution & "/"
                    End If
                    If cb_Methods.Checked Then
                        strOutput += strMethods & "/"
                    End If
                Case Else
                    strOutput += "is invalid - multiple solutions - took " & strSecondsMilli(strMilliseconds) & " to process./"
                    If cb_Clues.Checked Then
                        strOutput += strClues & "/"
                    End If
            End Select
        Next

        '---clean-up output---
        While InStr(strOutput, vbCrLf) > 0
            strOutput = Replace(strOutput, vbCrLf, "/")
        End While
        While InStr(strOutput, "//") > 0
            strOutput = Replace(strOutput, "//", "/")
        End While
        strOutput = Replace(strOutput, "/", vbCrLf)

        '---get average time to solve per puzzle---

        If i > 0 Then avgTime = totalTime / intCurrentPuzzle

        If blnCancel Then
            strMsg = "Batch process of " & batchFile & " cancelled (" & intCurrentPuzzle & " of " & rArray.Length & " puzzles processed)"
        Else
            strMsg = "Batch process of " & batchFile & " completed (" & intCurrentPuzzle & " puzzles)."
        End If
        strOutput = strMsg & vbCrLf & "Took " & strSecondsMilli(totalTime) _
        & " at an average time of " & strSecondsMilli(avgTime) & " per puzzle." & vbCrLf & "------------" & vbCrLf & strOutput

        outputResults = strOutput

    End Function
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        StopThreadWork = True
        System.Threading.Thread.Sleep(200)
    End Sub
    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        StopThreadWork = True
        System.Threading.Thread.Sleep(100)
        System.Threading.Thread.Sleep(100)
        Me.Close()
    End Sub
    Private Sub frmBatch_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.None Then
            e.Cancel = True
        End If
    End Sub
    Private Sub frmBatch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        AddHandler BatchFinished, AddressOf batchCompleted
    End Sub
    Private Sub cb_Methods_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_Methods.CheckedChanged
        If cb_Methods.Checked Then
            Me.cb_Default.Enabled = True
            Me.cb_Enabled.Enabled = True
        Else
            Me.cb_Default.Enabled = False
            Me.cb_Enabled.Enabled = False
        End If
    End Sub
    Private Sub cb_Default_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_Default.CheckedChanged
        If Me.cb_Default.Checked = True Then
            Me.cb_Enabled.Checked = False
        Else
            Me.cb_Enabled.Checked = True
        End If
    End Sub
    Private Sub cb_Enabled_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_Enabled.CheckedChanged
        If Me.cb_Enabled.Checked = True Then
            Me.cb_Default.Checked = False
        Else
            Me.cb_Default.Checked = True
        End If
    End Sub
End Class
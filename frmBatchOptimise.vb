Option Strict Off
Imports System.ComponentModel
Imports System.Threading
Public Class frmBatchOptimise
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
    Private intCurrentPuzzle As Integer
    Private strSymmetry As String
    Private intSymmetry As Integer = 0
    Dim executionTimeWatch As New Stopwatch()
    Dim puzzleTimeWatch As New Stopwatch()
    Public Event BatchFinished()
    Public Sub batch()
        StopThreadWork = False

        '---start the puzzle timer---
        executionTimeWatch.Start()

        Dim intPuzzleTime As Integer
        Dim strOptimised As String = ""

        Dim i As Integer
        Dim intRemoved As Integer
        intPuzzles = rArray.Length
        intCurrentPuzzle = 0

        For i = 0 To rArray.Length - 1
            Dim Solver As New clsSudokuSolver
            Solver.blnClassic = True
            intCurrentPuzzle += 1
            If rArray(i) = "-" Then
                Solver.strGrid = Replace(gArray(i), vbCrLf, "")
                Solver.vsSolvers = My.Settings._defaultSolvers
                puzzleTimeWatch.Start()
                Solver._vsUnique()
                '---rerun to optimise---
                strOptimised = ""
                intRemoved = 0
                If Solver.intCountSolutions = 1 Then
                    intSymmetry = 0
                    strSymmetry = ""
                    If Me.cb_MaintainSymmetry.Checked Then
                        '---check if puzzle has symmetry---
                        Dim strGame As String = Solver.strGrid
                        Dim cs As New Symmetry
                        cs.blnSamurai = False
                        cs.strGrid = strGame
                        Dim s() As String
                        s = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
                        intSymmetry = Array.IndexOf(s, cs.checkSymmetry)
                        If intSymmetry > 0 Then strSymmetry = cs.checkSymmetry
                    End If

                    Dim opt As New clsSudokuOptimise
                    opt.strGrid = Replace(gArray(i), vbCrLf, "")
                    opt.intType = 1
                    opt.OptimisePuzzle(intSymmetry:=intSymmetry)
                    If opt.isOptimised Then
                        strOptimised = opt.strOptimised
                        intRemoved = opt.intCluesRemoved
                    End If
                End If
                intPuzzleTime = puzzleTimeWatch.ElapsedMilliseconds
                puzzleTimeWatch.Reset()
                If strOptimised <> "" AndAlso strOptimised <> Replace(gArray(i), vbCrLf, "") Then
                    rArray(i) = "1/" & gArray(i) & "/" & strOptimised & "/" & strMulti(intPuzzleTime, "millisecond", False) & "/" & intRemoved & "/" & strSymmetry
                Else
                    rArray(i) = "0/" & gArray(i) & "/" & strOptimised & "/" & strMulti(intPuzzleTime, "millisecond", False) & "/" & intRemoved & "/" & strSymmetry
                End If
            Else
                '---input is invalid---
                rArray(i) = "-1/" & gArray(i) & "//" & "0 milliseconds/0/"
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
        Me.cb_Optimised = New System.Windows.Forms.CheckBox
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.btnExit = New System.Windows.Forms.Button
        Me.FileLabel = New System.Windows.Forms.Label
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.cb_MaintainSymmetry = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(27, 65)
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
        Me.cb_Clues.Location = New System.Drawing.Point(27, 104)
        Me.cb_Clues.Name = "cb_Clues"
        Me.cb_Clues.Size = New System.Drawing.Size(123, 17)
        Me.cb_Clues.TabIndex = 4
        Me.cb_Clues.Text = "Output starting clues"
        Me.ToolTip.SetToolTip(Me.cb_Clues, "Shows starting clues for each puzzle in the batch")
        Me.cb_Clues.UseVisualStyleBackColor = True
        '
        'cb_Optimised
        '
        Me.cb_Optimised.AutoSize = True
        Me.cb_Optimised.Location = New System.Drawing.Point(27, 129)
        Me.cb_Optimised.Name = "cb_Optimised"
        Me.cb_Optimised.Size = New System.Drawing.Size(133, 17)
        Me.cb_Optimised.TabIndex = 5
        Me.cb_Optimised.Text = "Output optimised clues"
        Me.ToolTip.SetToolTip(Me.cb_Optimised, "Show the solution to puzzles that have a unique solution")
        Me.cb_Optimised.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.Filter = "Sudoku files (*.sdm)|*.sdm"
        Me.OpenFileDialog.Multiselect = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(169, 148)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Close"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'FileLabel
        '
        Me.FileLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.FileLabel.Location = New System.Drawing.Point(15, 34)
        Me.FileLabel.Name = "FileLabel"
        Me.FileLabel.Size = New System.Drawing.Size(232, 23)
        Me.FileLabel.TabIndex = 9
        Me.FileLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cb_MaintainSymmetry
        '
        Me.cb_MaintainSymmetry.AutoSize = True
        Me.cb_MaintainSymmetry.Location = New System.Drawing.Point(27, 152)
        Me.cb_MaintainSymmetry.Name = "cb_MaintainSymmetry"
        Me.cb_MaintainSymmetry.Size = New System.Drawing.Size(112, 17)
        Me.cb_MaintainSymmetry.TabIndex = 10
        Me.cb_MaintainSymmetry.Text = "Maintain symmetry"
        Me.ToolTip.SetToolTip(Me.cb_MaintainSymmetry, "Maintain symmetry (where puzzles exhibit symmetry)")
        Me.cb_MaintainSymmetry.UseVisualStyleBackColor = True
        '
        'frmBatchOptimise
        '
        Me.ClientSize = New System.Drawing.Size(252, 174)
        Me.ControlBox = False
        Me.Controls.Add(Me.cb_MaintainSymmetry)
        Me.Controls.Add(Me.FileLabel)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.cb_Optimised)
        Me.Controls.Add(Me.cb_Clues)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.ProgressBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBatchOptimise"
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
        Me.cb_Optimised.Checked = True
    End Sub
    Private Sub btnLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If OpenFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim strSplit() As String
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
            If strLength >= 81 Then blnLength = True
            If blnLength Then
                '---process into array
                tempArr = Split(strInput, vbCrLf)
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
                        rArray(gc) = "-1/Input is invalid: " & errMsg & "//0/0/"
                        gc += 1
                        errMsg = ""
                    End If
                Next
            End If
            '---end process into array
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
            Me.FileLabel.Text = "Select file to start..."
            StopThreadWork = False
        End If
    End Sub
    Private Function outputResults(Optional ByVal blnCancel As Boolean = False) As String
        Dim i As Integer
        Dim strClues As String
        Dim strOptimised As String
        Dim strMilliseconds As String
        Dim arrSplit() As String
        Dim avgTime As Long
        Dim strMsg As String
        Dim intOptimised As Integer
        Dim intRemoved As Integer
        Dim strSymmetry As String
        For i = 0 To rArray.Length - 1
            If rArray(i) = "-" Then Exit For
            arrSplit = Split(rArray(i), "/")
            intOptimised = CInt(arrSplit(0))
            strClues = arrSplit(1)
            strOptimised = arrSplit(2)
            strMilliseconds = arrSplit(3)
            intRemoved = arrSplit(4)
            strSymmetry = arrSplit(5)
            strOutput += "Puzzle " & i + 1 & " of " & rArray.Length & ": "
            Select Case intOptimised
                Case -1
                    strOutput += "is invalid and could not be optimised. Took " & strMilliseconds & " to process./"
                    If cb_Clues.Checked Then
                        strOutput += strClues & "/"
                    End If
                Case 0
                    If Me.cb_MaintainSymmetry.Checked Then
                        If intSymmetry > 0 Then
                            strOutput += "could not be optimised to maintain " & strSymmetry & " symmetry. Took " & strMilliseconds & " to process./"
                        Else
                            strOutput += "could not be optimised. Took " & strMilliseconds & " to process./"
                        End If
                    Else
                        strOutput += "could not be optimised. Took " & strMilliseconds & " to process./"
                    End If
                    If cb_Clues.Checked Then
                        strOutput += strClues & "/"
                    End If
                Case 1
                    If Me.cb_MaintainSymmetry.Checked Then
                        If intSymmetry > 0 Then
                            strOutput += "has been optimised to maintain " & strSymmetry & " symmetry (" & strMulti(intRemoved, "clue", False) & " removed) - took " & strMilliseconds & " to process./"
                        Else
                            strOutput += "has been optimised (" & strMulti(intRemoved, "clue", False) & " removed) - took " & strMilliseconds & " to process./"
                        End If
                    Else
                        strOutput += "has been optimised (" & strMulti(intRemoved, "clue", False) & " removed) - took " & strMilliseconds & "./"
                    End If
                    If cb_Clues.Checked Then
                        strOutput += strClues & "/"
                    End If
                    If cb_Optimised.Checked Then
                        strOutput += strOptimised & "/"
                    End If
            End Select
        Next

        '---cleanup output---
        While InStr(strOutput, vbCrLf) > 0
            strOutput = Replace(strOutput, vbCrLf, "/")
        End While
        While InStr(strOutput, "//") > 0
            strOutput = Replace(strOutput, "//", "/")
        End While
        strOutput = Replace(strOutput, "/", vbCrLf)

        '---get average time to optimse per puzzle
        If i > 0 Then avgTime = totalTime / intCurrentPuzzle

        If blnCancel Then
            strMsg = "Batch optimise of " & batchFile & " cancelled (" & intCurrentPuzzle & " of " & rArray.Length & " puzzles processed)"
        Else
            strMsg = "Batch optimise of " & batchFile & " completed (" & intCurrentPuzzle & " puzzles)."
        End If
        strOutput = strMsg & vbCrLf & "Took " & strMulti(totalTime / 1000, "second", False) _
        & " at an average time of " & strMulti(avgTime, "millisecond", False) & " per puzzle." & vbCrLf & "------------" & vbCrLf & strOutput

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
End Class
Option Strict Off
Public Class frmGame
    Public ttHeight As Integer
    Public ListMultiGame(0) As List(Of String)
    Public intMultiCell(0) As List(Of Integer)
    Public intMultiGame As Integer = 1
    Public intActiveCell As Integer = 0
    Public strPuzzleSolution As String = ""
    '---track clicks / double clicks---
    Private isFirstClick As Boolean = True
    Private isDoubleClick As Boolean = False
    Private clickms As Integer = 0
    '---stacks to keep track of all the moves
    Public Moves As Stack(Of String)
    Public RedoMoves As Stack(Of String)
    '---track hints---
    Public strHintDetails = ""
    Public intHintCount = 0
    '---setup mode---
    Public blnSetup As Boolean
    Public intAllowSetupInput As Integer
    'TODO: Code more solving methods
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    Private Sub frmGame_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmInitial.Opacity = 100
        frmInitial.Visible = True
        frmInitial.ShowInTaskbar = True
        If blnSamurai Then
            frmInitial.GameType.SelectedIndex = 1
        Else
            frmInitial.GameType.SelectedIndex = 0
        End If
        frmInitial.GameType.Focus()
        Me.Dispose()
    End Sub
    Private Sub FindOnlyClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strMethod As String = sender.ToString
        Dim strUserMethod As String = Trim(strMethod)
        strMethod = "_" & strMethod
        strUserMethod = Microsoft.VisualBasic.LCase(strUserMethod)
        If strPuzzleSolution = "" Then
            ShowMessageLabel(strMsg:="Cannot find " & strUserMethod & " as puzzle has no valid solution", strImage:="error")
            Exit Sub
        End If
        If Not blnConsistent() Then Exit Sub

        StopShowAllSteps = False
        blnAllSteps = True

        If Not My.Settings._blnCandidates Then
            Dim style As MsgBoxStyle
            Dim response As MsgBoxResult
            style = MsgBoxStyle.Information Or MsgBoxStyle.YesNo
            '---display warning message---
            response = MsgBox("It is highly recommended that candidates are displayed. Do you wanr to do this?", style)
            If response = MsgBoxResult.Yes Then
                setCandidates(True)
            End If
        End If

        Dim sMoves As Integer = 0
        Dim eMoves As Integer = 1
        Dim tMoves As Integer = 0

        Dim strClues As String = ""
        Dim strCandidates As String = ""


        While eMoves > sMoves
            If StopShowAllSteps = True Then Exit While

            getGameStr(strFilled:=strClues, strCandidates:=strCandidates)
            If strClues = "" Or strCandidates = "" Then Exit While
            Dim solver As New clsSudokuSolver
            Select Case blnSamurai
                Case True
                    solver.blnClassic = False
                Case False
                    solver.blnClassic = True
            End Select
            sMoves = Moves.Count
            With solver
                .strGrid = strClues
                .blnStep = True
                .strCandidates = strCandidates
                .strInputGameSolution = strPuzzleSolution
                .vsSolvers = sender.tag.ToString
            End With
            Dim blnUnique As Boolean = solver._vsUnique()
            eMoves = Moves.Count
            '---saves the move into the stack---
            If eMoves > sMoves Then
                tMoves += 1
            Else
                StopShowAllSteps = False
                blnAllSteps = False
            End If

        End While

        If tMoves = 0 Then ShowMessageLabel(strMsg:="Attempt to find " & strUserMethod & " resulted in no eliminations.", strImage:="alert")

        Dim t As ToolStripMenuItem = SolveUpToToolStripMenuItem.OwnerItem
        With t
            t.HideDropDown()
        End With

        StopShowAllSteps = False
        blnAllSteps = False


    End Sub
    Private Sub SolveUpToClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strMethod As String = sender.ToString
        Dim strUserMethod As String = Trim(strMethod)
        strMethod = "_" & strMethod
        strUserMethod = Microsoft.VisualBasic.LCase(strUserMethod)
        If strPuzzleSolution = "" Then
            ShowMessageLabel(strMsg:="Cannot solve up to " & strUserMethod & " as puzzle has no valid solution", strImage:="error")
            Exit Sub
        End If
        If Not blnConsistent() Then Exit Sub
        Cursor.Current = Cursors.WaitCursor
        Dim sMoves As Integer
        Dim eMoves As Integer
        Dim strClues As String = ""
        Dim strCandidates As String = ""
        getGameStr(strFilled:=strClues, strCandidates:=strCandidates)
        If strClues = "" Or strCandidates = "" Then Exit Sub
        Dim solver As New clsSudokuSolver
        Select Case blnSamurai
            Case True
                solver.blnClassic = False
            Case False
                solver.blnClassic = True
        End Select
        sMoves = Moves.Count
        With solver
            .strGrid = strClues
            .strCandidates = strCandidates
            .strSolveUpToMethod = strMethod
            .strInputGameSolution = strPuzzleSolution
            .vsSolvers = My.Settings._EnabledSolvers
        End With
        Dim blnUnique As Boolean = solver._vsUnique()
        eMoves = Moves.Count
        '---saves the move into the stack---
        If eMoves > sMoves Then
        Else
            ShowMessageLabel(strMsg:="Attempt to solve up to " & strUserMethod & " resulted in no eliminations.", strImage:="alert")
        End If
        Dim t As ToolStripMenuItem = SolveUpToToolStripMenuItem.OwnerItem
        With t
            t.HideDropDown()
        End With
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub frmGame_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TODO: Uncomment and run once if you add a new solving method
        My.Settings._defaultSolvers = "1:True:Naked Single,1:True:Hidden Single,1:True:Locked Candidates,2:True:Naked Pair,2:True:Naked Triple,2:True:Naked Quad,3:True:Hidden Pair,3:True:Hidden Triple,3:True:Hidden Quad,3:True:X-Wing,3:True:Swordfish,3:True:Jellyfish,4:True:Colour Wrap,4:True:Colour Trap,4:True:XY Wing,4:True:XYZ Wing,4:True:Brute Force"
        'My.Settings._EnabledSolvers = My.Settings._defaultSolvers
        'My.Settings._UniqueSolvers = "1:True:Naked Single,1:True:Hidden Single,1:True:Locked Candidates,2:True:Naked Pair,2:True:Naked Triple,2:True:Naked Quad,3:True:Hidden Pair,3:True:Hidden Triple,3:True:Hidden Quad,4:True:Brute Force"
        Me.AllowDrop = True
        Me.Icon = My.Resources.sudoku
        ListMultiGame(0) = New List(Of String)
        intMultiCell(0) = New List(Of Integer)
        '---initialize the stacks---
        Moves = New Stack(Of String)
        RedoMoves = New Stack(Of String)
        '---setup sudoku cells---
        Dim row, col As Integer
        Dim intCount As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intStartTop As Integer
        Dim intRight As Integer
        Dim intW As Integer
        Dim intCellW As Integer = 14
        Dim intCellH As Integer = 14
        If blnSamurai Then
            intCellW = 10.5
            intCellH = 10.5
        End If
        Dim intGap As Integer = (intCellW * 3) - 1
        Dim intMargin As Integer = 5
        Dim intMax As Integer = 9
        If blnSamurai Then intMax = 21
        intLeft = (intMargin * 2) + Me.CheckBox1.Width
        intTop = intMargin + Me.MenuStrip.Height

        'Me.PictureBox1.Left = intLeft
        'Me.PictureBox1.Top = intTop

        intStartTop = intTop
        For row = 1 To intMax
            For col = 1 To intMax
                intCount += 1
                If Not blnSamurai Or Not blnIgnoreSamurai(intCount) Then
                    Dim sc As New SudokuCell
                    With sc
                        .Name = "SudokuCell" & intCount
                        .Left = intLeft
                        .Top = intTop
                        .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                        .Tag = intCount
                        .sc_IntCandidate = 0
                        .sc_IntBorders = 15
                        .sc_BorderColour = Color.FromName(My.Settings._borderColour)
                        .sc_IntW = intCellW
                        .sc_IntH = intCellH
                        .Parent = Me
                        .ContextMenuStrip = Me.RCContextMenuStrip
                    End With
                    AddHandler sc.MouseEnter, AddressOf Cell_MouseEnter
                    AddHandler sc.MouseLeave, AddressOf Cell_MouseLeave
                    AddHandler sc.MouseDown, AddressOf Cell_MouseClick
                    Me.Controls.Add(sc)
                End If
                intLeft += intGap
                If col Mod 3 = 0 Then
                    intLeft += 3
                End If
            Next

            intW = intLeft
            intRight = intLeft
            intLeft = (intMargin * 2) + Me.CheckBox1.Width
            intTop += intGap

            'Me.PictureBox1.Width = intW - ((intMargin * 3) + Me.CheckBox1.Width)
            'Me.PictureBox1.Height = Me.PictureBox1.Width

            If row Mod 3 = 0 Then
                intTop += 3
            End If
        Next
        '---end setup sudoku cells---

        Me.Height = intTop + (Me.Height - Me.ClientSize.Height) + intMargin + Me.StatusStrip1.Height
        Me.Width = intTop + (Me.Width - Me.ClientSize.Width) + Me.CheckBox1.Width + (intMargin * 3) - Me.MenuStrip.Height + Me.FeedbackBox.Width

        Me.FeedbackBox.Left = intRight + intMargin
        Me.FeedbackBox.Top = intStartTop
        Me.CenterToScreen()

        '---position checkboxes---
        setCheckBoxes(intLeft:=intMargin, intEnabled:=0)

        '---setup everything---
        restoreDefaults()

        '---setup sounds---
        setSounds(blnSwap:=False)

        '---setup candidares---
        setCandidates(blnSwap:=False)

        'Me.PictureBox1.SendToBack()

    End Sub
    Private Sub frmGame_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim i As Integer
            Dim arrFiles() As String
            arrFiles = e.Data.GetData(DataFormats.FileDrop)
            For i = 0 To UBound(arrFiles)
                If blnValidExtension(Microsoft.VisualBasic.Right(arrFiles(i), 3)) Then
                    e.Effect = DragDropEffects.All
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Sub frmGame_DragDrop( _
        ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim arrFiles() As String
            Dim strInput As String = ""
            Dim i As Integer
            '---Assign the files to an array---
            arrFiles = e.Data.GetData(DataFormats.FileDrop)
            For i = 0 To UBound(arrFiles)
                If blnValidExtension(Microsoft.VisualBasic.Right(arrFiles(i), 3)) Then
                    strInput += My.Computer.FileSystem.ReadAllText(arrFiles(i))
                End If
            Next
            inputPuzzle(strInput)
            Me.blnSetup = False
            Me.SetupLabel.Visible = False
        End If
    End Sub
    Private Sub frmGame_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim strError As String = ""
        '---manual input---
        If blnSetup AndAlso intActiveCell > 0 Then
            Select Case e.KeyValue.ToString
                Case 46
                    '---delete key---
                    ManualInput(intActiveCell, 0)
                Case 48 To 57
                    '---numeric keys---
                    ManualInput(intActiveCell, CStr(Chr(e.KeyValue.ToString)))
            End Select
        End If

        If ListMultiGame(0).Count > 0 Then
            Dim blnLoad As Boolean
            If e.KeyValue.ToString = 34 AndAlso intMultiGame < ListMultiGame(0).Count Then
                intMultiGame += 1
                blnLoad = True
            End If
            If e.KeyValue.ToString = 33 AndAlso intMultiGame > 1 Then
                intMultiGame -= 1
                blnLoad = True
            End If
            If blnLoad Then
                Me.ToolStripMultiGameLabel.Text = "Puzzle " & intMultiGame & " of " & ListMultiGame(0).Count
                ShowMessageLabel("")
                inputPuzzle(strInput:=ListMultiGame(0).Item(intMultiGame - 1).ToString, blnMultiImport:=False)
                restoreDefaults(blnDeleteMulti:=False)
                Me.blnSetup = False
                Me.SetupLabel.Visible = False
                Me.MsgTimer.Stop()
            End If
        End If
    End Sub
    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        Dim sMsg As String = Me.StatusLabel.Text
        Dim eMsg As String
        Dim strText As String = My.Computer.Clipboard.GetText
        If Not inputPuzzle(strInput:=strText) Then
            eMsg = Me.StatusLabel.Text
            If eMsg = sMsg Then ShowMessageLabel("No valid game to paste", "alert")
        End If
        If countClues(My.Settings._clueColour) > 0 Then
            Me.blnSetup = False
            Me.SetupLabel.Visible = False
        End If
    End Sub
    Private Sub FeedbackBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles FeedbackBox.Enter
        Me.MenuStrip.Focus()
    End Sub
    Private Sub FeedbackBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FeedbackBox.KeyPress
        e.Handled = False
    End Sub
    Private Sub Cell_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sc As SudokuCell
        sc = CType(sender, SudokuCell)
        intActiveCell = CInt(sc.Tag)
        If intMultiCell(0).IndexOf(sc.Tag) = -1 Then
            If sc.sc_IntSolution = 0 AndAlso IntCellPossibleBits(intActiveCell) > 0 Then
                With sc
                    .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intMO_BGTransparency)
                    .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intM0_BorderTransparency)
                End With
            End If
        End If
    End Sub
    Private Sub Cell_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sc As SudokuCell
        sc = CType(sender, SudokuCell)
        intActiveCell = CInt(sc.Tag)
        If intMultiCell(0).IndexOf(sc.Tag) = -1 Then
            'If sc.sc_IntSolution = 0 AndAlso IntCellPossibleBits(intActiveCell) > 0 Then
            With sc
                .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
            End With
            'End If
        End If
    End Sub
    Sub doubleClickTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles doubleClickTimer.Tick
        clickms += 100
        Dim sc As SudokuCell
        sc = getCtrl("SudokuCell" & intActiveCell)
        '----timer has reached the double click time limit.
        If clickms >= 250 Then
            doubleClickTimer.Stop()
            If isDoubleClick Then
                '---Perform double click action---
                If My.Settings._blnCandidates AndAlso sc.sc_IntSolution = 0 AndAlso sc.sc_IntCandidate > 0 AndAlso blnSingleBit(sc.sc_IntCandidate) Then
                    '---fill in naked single---
                    '---set the background back to normal
                    sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                    sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                    nakedSingle_doubleClick(intCell:=intActiveCell, intValue:=intReverseBit(sc.sc_IntCandidate), strSolvedBy:="U")
                    PlaySound(My.Resources.click)
                Else
                    '---do multiselect---
                    If intMultiCell(0).Count = 1 AndAlso Not My.Settings._blnCandidates Then
                    Else
                        If sc.sc_IntSolution = 0 AndAlso (sc.sc_IntCandidate > 0 Or IntCellPossibleBits(intActiveCell) > 0) Then
                            If intMultiCell(0).IndexOf(sc.Tag) = -1 Then
                                intMultiCell(0).Add(sc.Tag)
                                sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intMO_BGTransparency)
                                sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intM0_BorderTransparency)
                            Else
                                intMultiCell(0).Remove(sc.Tag)
                                sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                                sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                                If intMultiCell(0).Count = 0 Then setCheckBoxes(intEnabled:=0, intChecked:=0)
                            End If
                            enableCheckboxes()
                        End If
                    End If
                End If
            Else
                '---Perform single click action---
                Dim i As Integer
                If sc.sc_IntSolution = 0 AndAlso sc.sc_AlertGraphic = False Then
                    For i = 0 To intMultiCell(0).Count - 1
                        sc = getCtrl("SudokuCell" & intMultiCell(0).Item(i))
                        sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                        sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                    Next
                    setCheckBoxes(intEnabled:=0, intChecked:=0)
                    intMultiCell(0).Clear()
                End If
            End If
            '---allow the MouseDown event handler to process clicks again---
            isFirstClick = True
            isDoubleClick = False
            clickms = 0
        End If
    End Sub
    Sub Cell_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseDown
        Dim typeArr() As String = Split(sender.ToString, ".")
        If Not typeArr(1) = "SudokuCell" Then Exit Sub
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        Dim sc As SudokuCell
        sc = CType(sender, SudokuCell)
        If sc.sc_IntSolution = 0 And strPuzzleSolution <> "" Then
            '---first mouse click---
            If isFirstClick = True Then
                isFirstClick = False
                '---start the double click timer---
                doubleClickTimer.Start()
                '---second mouse click---
            Else
                If clickms < 250 Then
                    isDoubleClick = True
                End If
            End If
        End If
    End Sub
    Private Sub LoadFromFileToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadFromFileToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim strFileContents As String = ""
            Dim i As Integer
            Dim arrFiles() As String
            '---Assign the files to an array---
            arrFiles = OpenFileDialog1.FileNames
            For i = 0 To UBound(arrFiles)
                If blnValidExtension(Microsoft.VisualBasic.Right(arrFiles(i), 3)) Then
                    strFileContents += My.Computer.FileSystem.ReadAllText(arrFiles(i))
                End If
            Next
            inputPuzzle(strInput:=strFileContents)
            Me.blnSetup = False
            Me.SetupLabel.Visible = False
        End If
    End Sub
    Function IntCellPossibleBits(ByVal intCell As Integer) As Integer
        Dim c As Integer
        Dim v As Integer
        Dim ptr As Integer
        Dim arrayPeers(0) As String
        Dim sc As SudokuCell
        Select Case blnSamurai
            Case True
                arrayPeers = ArrSamuraiPeers(intCell)
            Case False
                arrayPeers = arrPeers(intCell)
        End Select
        IntCellPossibleBits = 511
        For c = 0 To UBound(arrayPeers)
            ptr = arrayPeers(c)
            sc = getCtrl("SudokuCell" & ptr)
            v = sc.sc_IntSolution
            If v > 0 Then
                If intGetBit(v) And IntCellPossibleBits Then
                    IntCellPossibleBits -= intGetBit(v)
                End If
            End If
        Next
    End Function
    Sub enableCheckboxes()
        If Not My.Settings._blnCandidates Then Exit Sub
        Dim i As Integer
        Dim v As Integer
        Dim lc As Integer
        Dim cb As CheckBox
        Dim sc As SudokuCell
        Dim intBound As Integer = intMultiCell(0).Count - 1
        Dim intArrayPossible(intBound) As Integer
        Dim intArrayActual(intBound) As Integer
        Dim vp(9) As Integer
        Dim va(9) As Integer
        Dim intBitAvailable As Integer
        Dim intBitPossible As Integer
        Dim ptr As Integer
        For i = 0 To intBound
            ptr = intMultiCell(0).Item(i)
            sc = getCtrl("SudokuCell" & ptr)
            intArrayPossible(i) = IntCellPossibleBits(ptr)
            intArrayActual(i) = sc.sc_IntCandidate
        Next

        For v = 1 To 9
            cb = getCB("Checkbox" & v)
            cb.Enabled = 0
            cb.Checked = 0
        Next

        Dim intBit As Integer
        For v = 1 To 9
            For i = 0 To intBound
                intBit = intGetBit(v)
                If (Not intBitAvailable And intBit) AndAlso (intBit And intArrayActual(i)) Then intBitAvailable += intGetBit(v)
                If (Not intBitPossible And intBit) AndAlso (intBit And intArrayPossible(i)) Then intBitPossible += intGetBit(v)
            Next
        Next

        For v = 1 To 9
            intBit = intGetBit(v)
            cb = getCB("Checkbox" & v)
            If intBitPossible And intBit Then
                cb.Enabled = True
            End If
            If intBitAvailable And intBit Then
                cb.Checked = True
                lc += 1
                i = v
            End If
        Next

        If lc = 1 Then
            cb = getCB("Checkbox" & i)
        End If

    End Sub
    Function setCheckBoxes(Optional ByVal intLeft As Integer = -1, Optional ByVal intEnabled As Integer = -1, Optional ByVal intChecked As Integer = -1) As Boolean
        Dim i As Integer
        Dim cb As CheckBox
        For i = 1 To 9
            cb = getCB("Checkbox" & i)
            If intLeft > -1 Then cb.Left = intLeft
            If intEnabled = 0 Or intEnabled = 1 Then
                cb.Enabled = CBool(intEnabled)
            End If
            If intChecked = 0 Or intChecked = 1 Then
                cb.Checked = CBool(intChecked)
            End If
        Next
    End Function
    Private Sub UndoToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        '---if no previous moves, then exit---
        If Moves.Count = 0 Then
            ShowMessageLabel("Nothing to undo at this time.", "info")
            Exit Sub
        End If
        '---remove from one stack and push into the redo stack---
        Dim str As String = Moves.Pop
        RedoMoves.Push(str)
        '--undo the move---
        UndoRedoMove(str)
    End Sub
    Private Sub RedoToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        '---if no further moves, then exit---
        If RedoMoves.Count = 0 Then
            ShowMessageLabel("Nothing to redo at this time.", "info")
            Exit Sub
        End If
        '---remove from one stack and push into the undo stack---
        Dim str As String = RedoMoves.Pop
        Moves.Push(str)
        '--redo the move---
        UndoRedoMove(str)
    End Sub
    Private Sub CB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click, CheckBox2.Click, CheckBox3.Click, CheckBox4.Click, CheckBox5.Click, CheckBox6.Click, CheckBox7.Click, CheckBox8.Click, CheckBox9.Click
        Dim i As Integer
        Dim ptr As Integer
        Dim cb As CheckBox
        Dim sc As SudokuCell
        Dim blnRemove As Boolean
        Dim intBit As Integer = intGetBit(sender.text)
        Dim intPossibleBits As Integer
        Dim strMoves As String = ""
        Dim strAddRemove As String = ""
        cb = getCB("checkbox" & sender.text)
        Select Case cb.Checked
            Case True
                '---add candidate---
                strAddRemove = "+" & sender.text
                blnRemove = False
            Case False
                '---remove candidate---
                strAddRemove = "-" & sender.text
                blnRemove = True
        End Select
        '---get selected cells---
        For i = 0 To intMultiCell(0).Count - 1
            ptr = intMultiCell(0).Item(i)
            sc = getCtrl("SudokuCell" & ptr)
            If blnRemove Then
                If intBit And sc.sc_IntCandidate Then
                    strMoves += ptr & ":" & strAddRemove & "|"
                    sc.sc_IntCandidate -= intBit
                    If sc.sc_IntCandidate = 0 Then sc.sc_AlertGraphic = True
                End If
            Else
                intPossibleBits = IntCellPossibleBits(ptr)
                If intBit And intPossibleBits Then
                    strMoves += ptr & ":" & strAddRemove & "|"
                    sc.sc_IntCandidate += intBit
                    sc.sc_AlertGraphic = False
                End If
            End If
        Next
        '---saves the move into the stack---
        Moves.Push(strMoves)
        enableCheckboxes()
    End Sub
    Sub UndoRedoMove(ByVal strMove As String)
        Dim i As Integer
        Dim arrayMove() As String
        Dim arraySplit() As String
        Dim arrType() As String
        Dim intBit As Integer
        Dim intValue As Integer
        Dim ptr As Integer
        Dim strLeader As String = ""
        Dim sc As SudokuCell
        Dim strType As String
        arrType = Split(strMove, "~")
        If arrType.Length = 1 Then
            strType = "u"
        Else
            strType = arrType(1)
            strMove = arrType(0)
        End If

        arrayMove = Split(strMove, "|")
        For i = 0 To UBound(arrayMove) - 1
            arraySplit = Split(arrayMove(i), ":")
            ptr = arraySplit(0)
            strLeader = Microsoft.VisualBasic.Left(arraySplit(1), 1)
            intValue = Microsoft.VisualBasic.Left(Replace(arraySplit(1), strLeader, "", 1, 1), 1)
            intBit = 0
            If strLeader = "=" Then
                intBit = CInt(Replace(arraySplit(1), strLeader & intValue & "/", "", 1, 1))
            End If
            sc = getCtrl("SudokuCell" & ptr)
            Select Case strLeader
                Case "="
                    '---set or remove value---
                    If sc.sc_IntSolution = intValue Then
                        '---remove value and restore candidates---
                        sc.sc_IntSolution = 0
                        sc.sc_IntCandidate = intBit
                        sc.sc_TextColour = Color.FromName(My.Settings._clueColour)
                    Else
                        '---set value---
                        sc.sc_IntSolution = intValue
                        sc.sc_IntCandidate = 0
                        Select Case strType
                            Case "U", "u"
                                sc.sc_TextColour = Color.FromName(My.Settings._userColour)
                            Case "P", "p"
                                sc.sc_TextColour = Color.FromName(My.Settings._solveColour)
                        End Select

                    End If
                Case Else

                    intBit = intGetBit(intValue)

                    If sc.sc_IntCandidate And intBit Then
                        If sc.sc_IntCandidate <> intBit Then
                            sc.sc_IntCandidate -= intBit
                            If sc.sc_IntCandidate = 0 Then sc.sc_AlertGraphic = True
                        End If
                    Else
                        sc.sc_IntCandidate += intBit
                        sc.sc_AlertGraphic = False
                    End If
            End Select
        Next
        enableCheckboxes()
    End Sub
    Private Function copyGameStr(ByVal intType As Integer) As String
        copyGameStr = ""
        Dim g As Integer
        Dim i As Integer
        Dim intClue As Integer
        Dim ptr As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim sc As SudokuCell
        intStart = 1
        intEnd = 5
        If Not blnSamurai Then intEnd = 1
        For g = intStart To intEnd
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            ptr = ptr
                        Else
                            ptr = 0
                        End If
                    Case False
                        ptr = i
                End Select
                If ptr > 0 Then
                    sc = getCtrl("SudokuCell" & ptr)
                    intClue = sc.sc_IntSolution
                    If intClue > 0 Then
                        Select Case intType
                            Case 1
                                '---get all filled cells---
                                copyGameStr += CStr(intClue)
                            Case 2
                                '---get clues---
                                If sc.sc_TextColour = Color.FromName(My.Settings._clueColour) Then
                                    copyGameStr += CStr(intClue)
                                Else
                                    copyGameStr += "."
                                End If
                        End Select
                    Else
                        copyGameStr += "."
                    End If
                End If
            Next
            copyGameStr += vbCrLf
        Next
    End Function
    Private Function getSolution() As String
        getSolution = strPuzzleSolution
        If Not blnSamurai Then Exit Function

        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim lstr As String = ""
        Dim _Clues(441) As String
        Dim midStr As String

        getSolution = Replace(getSolution, vbCrLf, "")

        For g = 1 To 5
            For i = 1 To 81
                ptr = intSamuraiOffset(i, g)
                midStr = Mid(getSolution, i + (81 * (g - 1)), 1)
                Select Case Asc(midStr)
                    Case 49 To 57
                        '---numeric, so record in array---
                        _Clues(ptr) = CInt(midStr)
                End Select
            Next
        Next

        For g = 1 To 5
            For i = 1 To 81
                ptr = intSamuraiOffset(i, g)
                If Not blnIgnoreSamurai(ptr) Then
                    If _Clues(ptr) = 0 Then
                        lstr += "."
                    Else
                        lstr += _Clues(ptr)
                    End If
                End If
            Next
            lstr += vbCrLf
        Next

        getSolution = lstr
    End Function
    Public Sub getGameStr(ByRef strFilled As String, ByRef strCandidates As String)
        strCandidates = ""
        strFilled = ""
        If strPuzzleSolution = "" Then Exit Sub
        Dim arrCandidates() As Integer
        ReDim arrCandidates(81)
        If blnSamurai Then ReDim arrCandidates(441)
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim intBit As Integer
        Dim strValue As String = ""
        Dim sc As SudokuCell
        Dim strSolution As String
        strSolution = Replace(strPuzzleSolution, vbCrLf, "")
        strSolution = Replace(strSolution, vbLf, "")
        intStart = 1
        intEnd = 5
        If Not blnSamurai Then intEnd = 1
        For g = intStart To intEnd
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            strValue = Mid(strSolution, i + (81 * (g - 1)), 1)
                        End If
                    Case False
                        ptr = i
                        strValue = Mid(strSolution, ptr, 1)
                End Select
                intBit = intGetBit(strValue)
                sc = getCtrl("SudokuCell" & ptr)
                If sc.sc_IntSolution > 0 Then
                    arrCandidates(ptr) = 0
                    If strFilled = "" Then
                        strFilled = strValue
                    Else
                        strFilled += strValue
                    End If
                Else
                    '---get candidates---
                    If sc.sc_IntCandidate And intBit Then
                        arrCandidates(ptr) = sc.sc_IntCandidate
                        strFilled += "."
                    End If
                End If
            Next
        Next
        For i = 1 To UBound(arrCandidates)
            If i = 1 Then
                strCandidates = arrCandidates(i)
            Else
                strCandidates += arrDivider & arrCandidates(i)
            End If
        Next
    End Sub
    Function blnConsistent(Optional ByVal blnMsg As Boolean = True) As Boolean
        If strPuzzleSolution = "" Then Exit Function
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim intBit As Integer
        Dim sc As SudokuCell
        Dim strValue As String = "0"
        intStart = 1
        intEnd = 5
        Dim strSolution As String = ""
        strSolution = Replace(strPuzzleSolution, vbCrLf, "")
        strSolution = Replace(strSolution, vbLf, "")
        If Not blnSamurai Then intEnd = 1
        For g = intStart To intEnd
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                        strValue = Mid(strSolution, i + (81 * (g - 1)), 1)
                    Case False
                        ptr = i
                        strValue = Mid(strSolution, i, 1)
                End Select
                intBit = intGetBit(strValue)
                sc = getCtrl("SudokuCell" & ptr)
                If sc.sc_IntSolution > 0 Then
                    '---check filled cells---
                    If sc.sc_IntSolution <> CInt(strValue) Then
                        '---error---
                        If blnMsg Then ShowMessageLabel("You've made an invalid move", "error")
                        Exit Function
                    End If
                Else
                    '---check candidates---
                    If sc.sc_IntCandidate And intBit Then
                    Else
                        '---error---
                        If blnMsg Then ShowMessageLabel("You've made an invalid move", "error")
                        Exit Function
                    End If
                End If
            Next
        Next
        blnConsistent = True
    End Function
    Private Sub ShowSolutionToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ShowSolutionToolStripMenuItem.Click
        If strPuzzleSolution = "" Then
            ShowMessageLabel("Cannot show solution as puzzle is not valid", "error")
            Exit Sub
        End If
        If Not blnConsistent() Then Exit Sub
        DisplayfrmSolution()
    End Sub
    Private Function DisplayfrmSolution() As Boolean
        Dim strGame As String
        strGame = strPuzzleSolution
        strGame = Replace(strGame, vbCrLf, "")
        strGame = Replace(strGame, vbLf, "")
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim sColour As Color = Color.FromName(My.Settings._solveColour)
        Dim sc As SudokuCell
        Dim input As String
        intStart = 1
        intEnd = 5
        If Not blnSamurai Then intEnd = 1
        For g = intStart To intEnd
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            sc = getCtrl("SudokuCell" & ptr)
                            If sc.sc_IntSolution = 0 Then
                                With sc
                                    input = Mid(strGame, i + (81 * (g - 1)), 1)
                                    If input <> "0" AndAlso input <> "." Then
                                        Try
                                            .sc_IntCandidate = 0
                                            .sc_TextColour = sColour
                                            .sc_IntSolution = input
                                            .sc_AlertGraphic = False
                                        Catch
                                        End Try
                                    End If
                                End With
                            End If
                        End If
                    Case False
                        ptr = i
                        sc = getCtrl("SudokuCell" & ptr)
                        If sc.sc_IntSolution = 0 Then
                            With sc
                                input = Mid(strGame, ptr, 1)
                                If input <> "0" AndAlso input <> "." Then
                                    .sc_IntCandidate = 0
                                    .sc_TextColour = sColour
                                    .sc_IntSolution = input
                                    .sc_AlertGraphic = False
                                End If
                            End With
                        End If
                End Select
            Next
        Next
        For i = 0 To intMultiCell(0).Count - 1
            sc = getCtrl("SudokuCell" & intMultiCell(0).Item(i))
            sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
            sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
        Next
        setCheckBoxes(intEnabled:=0, intChecked:=0)
        intMultiCell(0).Clear()
    End Function
    Private Sub RCContextMenuStrip_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RCContextMenuStrip.Opening
        Dim rmenu As ContextMenuStrip
        rmenu = Me.RCContextMenuStrip
        Dim eMenu As ToolStripMenuItem
        Dim dMenu As ToolStripMenuItem
        Dim sMenu As ToolStripMenuItem
        eMenu = rmenu.Items(0)
        dMenu = rmenu.Items(1)
        sMenu = rmenu.Items(2)
        sMenu.AutoSize = True
        eMenu.AutoSize = True
        dMenu.AutoSize = True

        Dim i As Integer
        Dim v As Integer
        Dim sc As SudokuCell
        Dim blnRS As Boolean = False
        Dim intBound As Integer = intMultiCell(0).Count - 1
        Dim intArrayPossible(intBound) As Integer
        Dim intArrayActual(intBound) As Integer
        Dim vp(9) As Integer
        Dim va(9) As Integer
        Dim intBitAvailable As Integer
        Dim intBitPossible As Integer
        Dim ptr As Integer
        For i = 0 To intBound
            ptr = intMultiCell(0).Item(i)
            sc = getCtrl("SudokuCell" & ptr)
            intArrayPossible(i) = IntCellPossibleBits(ptr)
            intArrayActual(i) = sc.sc_IntCandidate
        Next

        Me.RevealSolutionToolStripMenuItem.Visible = False

        If intActiveCell > 0 And intBound <= 0 Then
            sc = getCtrl("SudokuCell" & intActiveCell)
            If strPuzzleSolution <> "" And blnSetup = False And sc.sc_IntSolution = 0 And sc.sc_IntCandidate > 0 Then
                blnRS = True
                Me.RevealSolutionToolStripMenuItem.Visible = True
            End If
        End If

        If ptr = 0 Then
            eMenu.Visible = False
            dMenu.Visible = False
            sMenu.Visible = False
        End If

        If ptr = 0 And blnRS = False Then
            e.Cancel = True
            Exit Sub
        End If

        If ptr = 0 Then
            Me.ToolStripSeparator17.Visible = False
        End If

        If blnRS = False Then
            Me.ToolStripSeparator17.Visible = False
        End If

        If ptr > 0 Then

            If blnRS = True Then
                Me.ToolStripSeparator17.Visible = True
            End If

            eMenu.Visible = True
            dMenu.Visible = True
            sMenu.Visible = True

            Dim intBit As Integer
            For v = 1 To 9
                For i = 0 To intBound
                    intBit = intGetBit(v)
                    If (Not intBitAvailable And intBit) AndAlso (intBit And intArrayActual(i)) Then intBitAvailable += intGetBit(v)
                    If (Not intBitPossible And intBit) AndAlso (intBit And intArrayPossible(i)) Then intBitPossible += intGetBit(v)
                Next
                eMenu.DropDownItems.Item(v - 1).Enabled = False
                dMenu.DropDownItems.Item(v - 1).Enabled = False
                sMenu.DropDownItems.Item(v - 1).Enabled = False
            Next

            For v = 1 To 9
                intBit = intGetBit(v)
                If (intBitAvailable And intBit) And (intBitPossible And intBit) Then
                    dMenu.DropDownItems.Item(v - 1).Enabled = True
                    sMenu.DropDownItems.Item(v - 1).Enabled = True
                End If
                If (Not intBitAvailable And intBit) And (intBitPossible And intBit) Then
                    eMenu.DropDownItems.Item(v - 1).Enabled = True
                End If
            Next

            If intBitAvailable = 0 Then dMenu.Visible = False
            If intBitPossible = 0 Then eMenu.Visible = False

            If intBitPossible = intBitAvailable And intBitAvailable > 0 Then eMenu.Visible = False

            If intBound > 0 Then
                sMenu.Visible = False
                If intBitAvailable = intBitPossible And intBitAvailable > 0 Then
                    If blnSingleBit(intBitAvailable) And blnSingleBit(intBitPossible) Then
                        sMenu.Visible = True
                    End If
                End If
            End If

            If Not My.Settings._blnCandidates Then
                dMenu.Visible = False
                eMenu.Visible = False
            End If
        End If

    End Sub
    Private Sub eC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles E1.Click, E2.Click, E3.Click, E4.Click, E5.Click, E6.Click, E7.Click, E8.Click, E9.Click
        Dim i As Integer
        Dim ptr As Integer
        Dim sc As SudokuCell
        Dim intBit As Integer = intGetBit(sender.text)
        Dim intPossibleBits As Integer
        Dim strMoves As String = ""
        Dim strAdd As String = ""
        '---add candidate---
        strAdd = "+" & sender.text
        '---get selected cells---
        For i = 0 To intMultiCell(0).Count - 1
            ptr = intMultiCell(0).Item(i)
            sc = getCtrl("SudokuCell" & ptr)
            intPossibleBits = IntCellPossibleBits(ptr)
            If intBit And intPossibleBits Then
                strMoves += ptr & ":" & strAdd & "|"
                sc.sc_IntCandidate += intBit
                sc.sc_AlertGraphic = False
            End If
        Next
        '---saves the move into the stack---
        Moves.Push(strMoves)
        enableCheckboxes()
    End Sub
    Private Sub dC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles D1.Click, D2.Click, D3.Click, D4.Click, D5.Click, D6.Click, D7.Click, D8.Click, D9.Click
        Dim i As Integer
        Dim ptr As Integer
        Dim sc As SudokuCell
        Dim intBit As Integer = intGetBit(sender.text)
        Dim intPossibleBits As Integer
        Dim strMoves As String = ""
        Dim strRemove As String = ""
        '---remove candidate---
        strRemove = "-" & sender.text
        '---get selected cells---
        For i = 0 To intMultiCell(0).Count - 1
            ptr = intMultiCell(0).Item(i)
            sc = getCtrl("SudokuCell" & ptr)
            intPossibleBits = IntCellPossibleBits(ptr)
            If intBit And sc.sc_IntCandidate Then
                strMoves += ptr & ":" & strRemove & "|"
                sc.sc_IntCandidate -= intBit
                If sc.sc_IntCandidate = 0 Then sc.sc_AlertGraphic = True
            End If
        Next
        '---saves the move into the stack---
        Moves.Push(strMoves)
        enableCheckboxes()
    End Sub
    Private Sub sC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles S1.Click, S2.Click, S3.Click, S4.Click, S5.Click, S6.Click, S7.Click, S8.Click, S9.Click
        Dim i As Integer
        Dim j As Integer
        Dim ptr As Integer
        Dim sc As SudokuCell
        Dim strMoves As String = ""
        Dim intCandidates As Integer
        Dim intBit As Integer
        Dim strEqual As String = ""
        Dim strRemove As String = ""
        '---remove candidate---
        strEqual = "=" & sender.text
        strRemove = "-" & sender.text
        intBit = intGetBit(sender.text)
        '---get selected cells---
        For i = 0 To intMultiCell(0).Count - 1
            ptr = intMultiCell(0).Item(i)
            sc = getCtrl("SudokuCell" & ptr)
            intCandidates = sc.sc_IntCandidate
            strMoves += ptr & ":" & strEqual & "/" & intCandidates & "|"
            sc.sc_IntCandidate = 0
            sc.sc_TextColour = Color.FromName(My.Settings._userColour)
            sc.sc_IntSolution = sender.text
        Next

        Dim arrayPeers(0) As String
        For i = 0 To intMultiCell(0).Count - 1
            ptr = intMultiCell(0).Item(i)
            Select Case blnSamurai
                Case True
                    arrayPeers = ArrSamuraiPeers(ptr)
                Case False
                    arrayPeers = arrPeers(ptr)
            End Select
            For j = 0 To UBound(arrayPeers)
                ptr = arrayPeers(j)
                sc = getCtrl("SudokuCell" & ptr)
                If intBit And sc.sc_IntCandidate Then
                    strMoves += ptr & ":" & strRemove & "|"
                    sc.sc_IntCandidate -= intBit
                End If
            Next
        Next
        strMoves += "~u"
        '---saves the move into the stack---
        Moves.Push(strMoves)
        '---removes multiselect---
        For i = 0 To intMultiCell(0).Count - 1
            sc = getCtrl("SudokuCell" & intMultiCell(0).Item(i))
            sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
            sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
        Next
        setCheckBoxes(intEnabled:=0, intChecked:=0)
        intMultiCell(0).Clear()
        PlaySound(My.Resources.click)
    End Sub
    Public Function nakedSingle_doubleClick(ByVal intCell As Integer, ByVal intValue As Integer, ByVal strSolvedBy As String) As Boolean
        Dim j As Integer
        Dim ptr As Integer
        Dim sc As SudokuCell
        Dim strMoves As String = ""
        Dim strRemove As String = ""
        Dim intCandidates As Integer
        Dim intBit As Integer
        Dim cCol As Color

        Select Case strSolvedBy
            Case "U", "u"
                '---solved by user---
                cCol = Color.FromName(My.Settings._userColour)
            Case "P", "p"
                '---solved by program---
                cCol = Color.FromName(My.Settings._solveColour)
        End Select

        '---remove candidate---
        strMoves = intCell & ":" & "=" & intValue
        strRemove = "-" & intValue
        intBit = intGetBit(intValue)
        '---get selected cells---
        ptr = intCell
        sc = getCtrl("SudokuCell" & ptr)
        intCandidates = sc.sc_IntCandidate
        strMoves += "/" & intCandidates & "|"
        sc.sc_IntCandidate = 0
        sc.sc_TextColour = cCol
        sc.sc_IntSolution = intValue

        Dim arrayPeers(0) As String
        ptr = intCell
        Select Case blnSamurai
            Case True
                arrayPeers = ArrSamuraiPeers(ptr)
            Case False
                arrayPeers = arrPeers(ptr)
        End Select
        For j = 0 To UBound(arrayPeers)
            ptr = arrayPeers(j)
            sc = getCtrl("SudokuCell" & ptr)
            If sc.sc_IntSolution = 0 Then
                If intBit And sc.sc_IntCandidate Then
                    strMoves += ptr & ":" & strRemove & "|"
                    sc.sc_IntCandidate -= intBit
                    If sc.sc_IntCandidate <= 0 Then
                        sc.sc_AlertGraphic = True
                    End If
                End If
            End If
        Next
        '---saves the move into the stack---

        Select Case strSolvedBy
            Case "U", "u"
                '---solved by user---
                strMoves += "~u"
            Case "P", "p"
                '---solved by program---
                strMoves += "~p"
        End Select

        Moves.Push(strMoves)
        '---removes multiselect---
        Dim i As Integer
        For i = 0 To intMultiCell(0).Count - 1
            sc = getCtrl("SudokuCell" & intMultiCell(0).Item(i))
            sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
            sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
        Next
        setCheckBoxes(intEnabled:=0, intChecked:=0)
        intMultiCell(0).Clear()
    End Function
    Public Function ManualInput(ByVal intCell As Integer, ByVal intValue As Integer) As Boolean

        If intAllowSetupInput = 0 And intValue > 0 Then
            ShowMessageLabel("No valid solutions - check and fix input before proceeding", "alert")
        End If

        Dim i As Integer
        Dim ptr As Integer
        Dim sc As SudokuCell
        Dim peersArr(0) As String
        Dim p As Integer
        Dim r As Integer
        Dim cCol As Color = Color.FromName(My.Settings._clueColour)
        Dim blnUndo As Boolean
        Dim blnErrFixed As Boolean
        Dim intSolutions As Integer
        Dim strMsg As String = ""

        '---scan peers for conflicts---
        Dim arrayPeers(0) As String
        ptr = intCell
        Select Case blnSamurai
            Case True
                arrayPeers = ArrSamuraiPeers(ptr)
            Case False
                arrayPeers = arrPeers(ptr)
        End Select
        For i = 0 To UBound(arrayPeers)
            ptr = arrayPeers(i)
            sc = getCtrl("SudokuCell" & ptr)
            If sc.sc_IntSolution = intValue AndAlso intValue > 0 Then
                '---clues already placed---
                ShowMessageLabel("Placing a value of " & intValue & " here would cause a conflict with existing clues.", "error")
                Exit Function
            End If
        Next

reload:

        '---get and set cell---
        ptr = intCell
        sc = getCtrl("SudokuCell" & ptr)
        sc.sc_TextColour = cCol
        If blnUndo Then
            sc.sc_IntSolution = 0
            blnUndo = False
            blnErrFixed = True
        Else
            blnErrFixed = False
            sc.sc_IntSolution = intValue
        End If
        sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)

        '---set initial candidates---
        Dim g As Integer
        Dim intMaxGrid As Integer = 1
        If blnSamurai Then intMaxGrid = 5
        Dim cCount As Integer = countClues(My.Settings._clueColour)
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                    Case False
                        ptr = i
                End Select
                If Not blnSamurai Or (blnSamurai AndAlso Not blnIgnoreSamurai(ptr)) Then
                    sc = getCtrl("SudokuCell" & ptr)
                    If cCount = 0 Then
                        sc.sc_IntCandidate = 0
                    Else
                        sc.sc_IntCandidate = 511
                    End If
                End If
            Next
        Next

        If cCount = 0 Then Exit Function

        '---update initial candidates---
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                    Case False
                        ptr = i
                End Select
                If Not blnSamurai Or (blnSamurai AndAlso Not blnIgnoreSamurai(ptr)) Then
                    sc = getCtrl("SudokuCell" & ptr)
                    If sc.sc_IntSolution > 0 Then
                        r = sc.sc_IntSolution
                        sc.sc_IntCandidate = 0
                        '---update peers---
                        Select Case blnSamurai
                            Case True
                                peersArr = ArrSamuraiPeers(ptr)
                            Case False
                                peersArr = arrPeers(ptr)
                        End Select
                        For p = 0 To UBound(peersArr)
                            ptr = peersArr(p)
                            sc = getCtrl("SudokuCell" & ptr)
                            If blnMatchBit(sc.sc_IntCandidate, r) Then
                                sc.sc_IntCandidate -= intGetBit(r)
                            End If
                        Next
                        '---end update peers---
                    End If
                End If
            Next
        Next

        If blnSamurai Then
            Dim sg() As String = Split(copyGameStr(intType:=2), vbCrLf)
            For g = 1 To intMaxGrid
                If Not blnValidGrid(sg(g - 1), "", g) Then
                    ShowMessageLabel("Cannot progress - this move would result in input being invalid", "error")
                    blnUndo = True
                    Exit For
                End If
            Next
        Else
            If Not blnValidGrid(copyGameStr(intType:=2), "") Then
                ShowMessageLabel("Cannot progress - this move would result in input being invalid", "error")
                blnUndo = True
            End If
        End If

        If blnUndo Then GoTo reload

        If Not blnSamurai Then
            intAllowSetupInput = -1
            If cCount >= 17 Then
                Dim strGame As String = copyGameStr(intType:=2)
                Dim solver As New clsSudokuSolver
                Dim intMax As Integer = 1000
                solver.strGrid = strGame
                solver.blnClassic = True
                solver.intQuit = intMax
                solver.vsSolvers = "1:True:Naked Single" 'My.Settings._UniqueSolvers
                If solver._vsUnique Then
                    intSolutions = solver.intCountSolutions
                    intAllowSetupInput = 1
                    If Not blnErrFixed Then
                        ShowMessageLabel("Input clues result in a valid puzzle.", "info")
                    End If
                Else
                    intSolutions = solver.intCountSolutions
                    If intSolutions = 0 Then intAllowSetupInput = 0

                    If Not blnErrFixed Then
                        Select Case intSolutions
                            Case intMax
                                strMsg = "Input clues result in a puzzle with " & intMax & " or more solutions"
                                ShowMessageLabel(strMsg, "info")
                            Case 0
                                strMsg = "Input clues result in a puzzle with 0 solutions"
                                ShowMessageLabel(strMsg, "alert")
                            Case Else
                                strMsg = "Input clues result in a puzzle with " & strMulti(intSolutions, "solution", False)
                                ShowMessageLabel(strMsg, "info")
                        End Select

                        If My.Settings._blnManualInputSuggest AndAlso solver.intCountSolutions > 1 AndAlso solver.intCountSolutions < solver.intQuit Then
                            Dim strRevised As String = ""
                            Dim blnRevised As Boolean
                            Dim s As Integer
                            Dim c As Integer
                            Dim m(81) As Integer
                            Dim chk(81) As Boolean
                            Dim arrH(0) As Integer
                            Dim mc As Integer
                            Dim curCount As Integer
                            Dim prevCount As Integer
                            Dim chr As String
                            Dim intChr As Integer
                            For c = 1 To 81
                                chk(c) = True
                            Next
                            For s = 0 To UBound(solver.Solutions)
                                If Array.IndexOf(chk, True) = -1 Then Exit For
                                For c = 1 To 81
                                    chr = Mid(solver.Solutions(s), c, 1)
                                    intChr = CInt(chr)
                                    If m(c) = 0 Then
                                        m(c) = intChr
                                    Else
                                        If intChr <> m(c) Then
                                            chk(c) = False
                                            m(c) = -1
                                        End If
                                    End If
                                Next
                            Next
                            For c = 1 To 81
                                chr = Mid(solver.strGrid, c, 1)
                                If chr = "." Then
                                    '---unique value across all solutions---
                                    '---and not found in starting grid---
                                    If m(c) > 0 Then
                                        strRevised += CStr(m(c))
                                        blnRevised = True
                                    Else
                                        strRevised += chr
                                    End If
                                Else
                                    strRevised += chr
                                End If
                            Next
                            If blnRevised Then
                                prevCount = 0
                                curCount = 0
                                mc = 0
                                For c = 1 To 81
                                    chr = Mid(strRevised, c, 1)
                                    If chr = "." Then
                                        sc = getCtrl("SudokuCell" & c)
                                        curCount = intCountBits(sc.sc_IntCandidate)
                                        If curCount > prevCount Then
                                            mc = 0
                                            ReDim arrH(mc)
                                            arrH(mc) = c
                                        End If
                                        If curCount = prevCount Then
                                            mc += 1
                                            ReDim Preserve arrH(mc)
                                            arrH(mc) = c
                                        End If
                                    End If
                                    prevCount = curCount
                                Next
                                For mc = 0 To UBound(arrH)
                                    sc = getCtrl("SudokuCell" & arrH(mc))
                                    With sc
                                        .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intMO_BGTransparency)
                                        .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intM0_BorderTransparency)
                                    End With
                                Next
                                frmHidden.strMsg = strMsg & vbCrLf & vbCrLf & "For the quickest path to finding a unique solution, try entering clues in one of the suggested cells. " & vbCrLf & vbCrLf & "Press any key to continue." & vbCrLf & vbCrLf & "To turn this feature off, select 'Options' from the main menu, then select the 'Manual input' tab"
                                frmHidden.ShowDialog(Me)
                                For c = 1 To 81
                                    chr = Mid(strRevised, c, 1)
                                    If chr = "." Then
                                        sc = getCtrl("SudokuCell" & c)
                                        With sc
                                            .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                                            .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                                        End With
                                    End If
                                Next
                                blnRevised = False
                            End If
                        End If
                    End If

                End If
            End If
        End If
    End Function
    Private Sub MsgTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MsgTimer.Tick
        StatusLabel.Text = ""
        StatusLabel.Image = Nothing
        StatusLabel.ToolTipText = ""
        MsgTimer.Stop()
    End Sub
    Private Sub HintToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles HintToolStripMenuItem.Click
        If strPuzzleSolution = "" Then
            ShowMessageLabel(strMsg:="No hint available as puzzle has no valid solution", strImage:="error")
            Exit Sub
        End If
        If Not blnConsistent() Then Exit Sub
        Dim strClues As String = ""
        Dim strCandidates As String = ""
        getGameStr(strFilled:=strClues, strCandidates:=strCandidates)
        If strClues = "" Or strCandidates = "" Then Exit Sub
        Dim solver As New clsSudokuSolver
        Select Case blnSamurai
            Case True
                solver.blnClassic = False
            Case False
                solver.blnClassic = True
        End Select
        With solver
            .blnHint = True
            .strHintDetails = strHintDetails
            .intHintCount = intHintCount
            .strGrid = strClues
            .strCandidates = strCandidates
            .strInputGameSolution = strPuzzleSolution
            .vsSolvers = My.Settings._EnabledSolvers
        End With
        Dim blnUnique As Boolean = solver._vsUnique()
        strHintDetails = solver.strHintDetails
        intHintCount = solver.intHintCount
    End Sub
    Private Sub CopyCluesToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyCluesToolStripMenuItem.Click
        Dim strText = copyGameStr(intType:=2)
        If strText <> "" Then
            My.Computer.Clipboard.SetText(strText, TextDataFormat.Text)
        Else
            ShowMessageLabel(strMsg:="Nothing to copy", strImage:="alert")
        End If
    End Sub
    Private Sub CopyFilledToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyFilledToolStripMenuItem.Click
        Dim strText = copyGameStr(intType:=1)
        If strText <> "" Then
            My.Computer.Clipboard.SetText(strText, TextDataFormat.Text)
        Else
            ShowMessageLabel(strMsg:="Nothing to copy", strImage:="alert")
        End If
    End Sub
    Private Sub CopySolutionToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopySolutionToolStripMenuItem.Click
        Dim strText = getSolution()
        If strText <> "" Then
            My.Computer.Clipboard.SetText(strText, TextDataFormat.Text)
        Else
            ShowMessageLabel(strMsg:="Nothing to copy", strImage:="alert")
        End If
    End Sub
    Private Sub CheckConsistencyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckConsistencyToolStripMenuItem.Click
        If strPuzzleSolution = "" Then
            ShowMessageLabel(strMsg:="Puzzle has no valid solution", strImage:="error")
            Exit Sub
        End If
        If Not blnConsistent() Then Exit Sub
        ShowMessageLabel(strMsg:="Entered values are consistent with puzzle solution", strImage:="info")
    End Sub
    Private Sub ToolStripMultiGameLabel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMultiGameLabel.Click
        frmPickGame.intMax = ListMultiGame(0).Count
        frmPickGame.intGame = intMultiGame
        frmPickGame.ShowDialog(Me)
        Dim intGame As Integer
        If frmPickGame.DialogResult = Windows.Forms.DialogResult.OK Then
            intGame = CInt(frmPickGame.NoGames.Text)
            If intGame <> intMultiGame Then
                intMultiGame = intGame
                Me.ToolStripMultiGameLabel.Text = "Puzzle " & intMultiGame & " of " & ListMultiGame(0).Count
                ShowMessageLabel("")
                inputPuzzle(strInput:=ListMultiGame(0).Item(intMultiGame - 1).ToString(), blnMultiImport:=False)
                restoreDefaults(blnDeleteMulti:=False)
                Me.blnSetup = False
                Me.SetupLabel.Visible = False
            End If
        End If
    End Sub
    Private Function inputPuzzle(ByVal strInput As String, Optional ByVal blnMultiImport As Boolean = True) As Boolean
        Dim strLength As Integer = 0
        Dim blnLength As Boolean = False
        Dim strError As String = ""
        inputPuzzle = False
        strInput = readinInput(strInput)
        strLength = Len(strInput)
        Select Case blnSamurai
            Case True
                If strLength >= 413 Then blnLength = True
            Case False
                If strLength >= 81 Then blnLength = True
        End Select
        If blnLength Then
            If blnMultiImport Then
                frmClear(blnSamurai)
                blnMultiInput(strInput, blnSamurai)
            End If
            If Not frmLoad(strGrid:=ListMultiGame(0).Item(intMultiGame - 1).ToString, blnSamurai:=blnSamurai, ErrMsg:=strError) Then
                frmClear(blnSamurai)
            Else
                If strError = "" Then
                    '---validate by attempting to solve---
                    Dim blnUnique As Boolean
                    Dim intCount As Integer
                    Dim intSolverQuit As Integer = 1000
                    Dim solver As New clsSudokuSolver
                    solver.intQuit = intSolverQuit
                    Select Case blnSamurai
                        Case False
                            solver.blnClassic = True
                            solver.strGrid = ListMultiGame(0).Item(intMultiGame - 1).ToString
                            solver.vsSolvers = My.Settings._UniqueSolvers
                            blnUnique = solver._vsUnique()
                            intCount = solver.intCountSolutions
                            strPuzzleSolution = solver.strGameSolution
                        Case True
                            frmProgress.strGrid = ListMultiGame(0).Item(intMultiGame - 1).ToString
                            frmProgress.intQuit = 3000 'intSolverQuit
                            frmProgress.ShowDialog(Me)
                            blnUnique = frmProgress.blnUnique
                            intCount = frmProgress.intCountSolutions
                            strPuzzleSolution = frmProgress.strGameSolution
                            Me.Refresh()
                    End Select
                    If Not blnUnique Then
                        If intCount = intSolverQuit Then
                            ShowMessageLabel("Puzzle is not valid (" & intCount & " or more solutions)", "error")
                        Else
                            ShowMessageLabel("Puzzle is not valid (" & intCount & " solutions)", "error")
                        End If
                    Else
                        inputPuzzle = True
                        '---valid puzzle---
                        ShowMessageLabel("")
                    End If
                End If
            End If
        End If
    End Function
    Private Sub setSounds(ByVal blnSwap As Boolean)
        If blnSwap Then
            If My.Settings._PlaySounds Then
                My.Settings._PlaySounds = False
                Me.SoundLabel.Image = My.Resources.sound_mute
                Me.SoundLabel.Tag = "Sounds are off. Double click to enable"
            Else
                My.Settings._PlaySounds = True
                Me.SoundLabel.Image = My.Resources.sound
                Me.SoundLabel.Tag = "Sounds are on. Double click to disable"
            End If
        Else
            '---setup label---
            If My.Settings._PlaySounds Then
                Me.SoundLabel.Image = My.Resources.sound
                Me.SoundLabel.Tag = "Sounds are on. Double click to disable"
            Else
                Me.SoundLabel.Image = My.Resources.sound_mute
                Me.SoundLabel.Tag = "Sounds are off. Double click to enable"
            End If
        End If
    End Sub
    Public Sub setCandidates(ByVal blnSwap As Boolean)
        If blnSwap Then
            If My.Settings._blnCandidates Then
                My.Settings._blnCandidates = False
                Me.CandidatesLabel.Image = My.Resources.candidates_remove
                Me.CandidatesLabel.Tag = "Candidates are off. Double click to enable"
            Else
                My.Settings._blnCandidates = True
                Me.CandidatesLabel.Image = My.Resources.candidates_add
                Me.CandidatesLabel.Tag = "Candidates are on. Double click to disable"
            End If
            '---removes multiselect---
            Dim i As Integer
            Dim sc As SudokuCell
            For i = 0 To intMultiCell(0).Count - 1
                sc = getCtrl("SudokuCell" & intMultiCell(0).Item(i))
                sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
            Next
            setCheckBoxes(intEnabled:=0, intChecked:=0)
            intMultiCell(0).Clear()
        Else
            '---setup label---
            If My.Settings._blnCandidates Then
                Me.CandidatesLabel.Image = My.Resources.candidates_add
                Me.CandidatesLabel.Tag = "Candidates are on. Double click to disable"
            Else
                Me.CandidatesLabel.Image = My.Resources.candidates_remove
                Me.CandidatesLabel.Tag = "Candidates are off. Double click to enable"
            End If
        End If
        Me.Refresh()
    End Sub
    Private Sub soundlabel_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles SoundLabel.DoubleClick
        setSounds(blnSwap:=True)
        Dim activeCtrl As ToolStripStatusLabel = CType(sender, ToolStripStatusLabel)
        '---show tooltip---
        showTT(activeCtrl)
    End Sub
    Private Sub CandidatesLabel_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles CandidatesLabel.DoubleClick
        setCandidates(blnSwap:=True)
        Dim activeCtrl As ToolStripStatusLabel = CType(sender, ToolStripStatusLabel)
        '---show tooltip---
        showTT(activeCtrl)
    End Sub
    Private Function countClues(ByVal strColour As String) As Integer
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim s As Integer
        Dim e As Integer
        Dim mColour As Color = Color.FromName(strColour)
        Dim sc As SudokuCell
        s = 1
        e = 5
        If Not blnSamurai Then
            e = 1
        End If
        For g = s To e
            For i = 1 To 81
                Select Case blnSamurai
                    Case False
                        '---Classic---
                        ptr = i
                        sc = getCtrl("SudokuCell" & ptr)
                        If sc.sc_IntSolution > 0 AndAlso sc.sc_TextColour = mColour Then countClues += 1
                    Case True
                        '---Samurai---
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            sc = getCtrl("SudokuCell" & ptr)
                            If sc.sc_IntSolution > 0 AndAlso sc.sc_TextColour = mColour Then countClues += 1
                        End If
                End Select
            Next
        Next

        If blnSamurai Then
            '---adjust for overlaps---
            For i = 0 To UBound(arrSamuraiOverlap)
                sc = getCtrl("SudokuCell" & arrSamuraiOverlap(i))
                If sc.sc_IntSolution > 0 AndAlso sc.sc_TextColour = mColour Then countClues -= 1
            Next
        End If

    End Function
    Private Sub Label_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles InfoLabel.MouseHover, SoundLabel.MouseHover, ToolStripMultiGameLabel.MouseHover, SetupLabel.MouseHover, CandidatesLabel.MouseHover
        Dim activeCtrl As ToolStripStatusLabel = CType(sender, ToolStripStatusLabel)
        Select Case activeCtrl.Name
            Case "InfoLabel"
                If countClues(My.Settings._clueColour) > 0 Then
                    activeCtrl.Tag = "Puzzle has " & strMulti(countClues(My.Settings._clueColour), "starting clue", False) & "."
                    If countClues(My.Settings._userColour) Then
                        activeCtrl.Tag += vbCrLf & strMulti(countClues(My.Settings._userColour), "clue", False) & " solved by user."
                    End If
                    If countClues(My.Settings._solveColour) Then
                        activeCtrl.Tag += vbCrLf & strMulti(countClues(My.Settings._solveColour), "clue", False) & " solved by the solver."
                    End If
                    Dim cs As New Symmetry
                    cs.blnSamurai = blnSamurai
                    cs.strGrid = copyGameStr(intType:=2)

                    If strPuzzleSolution <> "" And Not blnSolved() Then
                        'TODO: Try to improve solving speed for samurai puzzles
                        Dim solver As New clsSudokuSolver
                        Dim strClues As String = copyGameStr(intType:=2)
                        If strClues <> "" Then
                            solver.blnClassic = True
                            If blnSamurai Then
                                solver.blnClassic = False
                            End If
                            solver.strGrid = Replace(strClues, vbCrLf, "")
                            solver.strCandidates = ""
                            solver.vsSolvers = My.Settings._defaultSolvers
                            solver.strInputGameSolution = strPuzzleSolution
                            If solver._vsUnique() Then
                                activeCtrl.Tag += vbCrLf & solver.strDifficulty
                            End If
                        End If

                        Dim strMethods As String = ""
                        Dim sm() As String = solver.solveMethods
                        Dim sc() As Integer = solver.solveCountMethods
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
                        activeCtrl.Tag += " - " & strMethods
                    End If
                    activeCtrl.Tag += vbCrLf & "Puzzle clues have " & cs.checkSymmetry() & " symmetry."
                Else
                    activeCtrl.Tag = "No puzzle information available as there are no starting clues."
                End If
            Case "SetupLabel"
                    If blnSamurai Then
                        activeCtrl.Tag = "Double click to check if input clues result in a valid puzzle"
                    End If
            Case Else
        End Select
        '---show tooltip---
        showTT(activeCtrl)
    End Sub
    Private Sub Label_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles InfoLabel.MouseLeave, SoundLabel.MouseLeave, ToolStripMultiGameLabel.MouseLeave
        ToolTip1.RemoveAll()
        If Me.ToolTip1.Active Then Me.ToolTip1.Hide(Me)
    End Sub
    Private Sub showTT(ByVal activeCtrl As ToolStripStatusLabel)
        '---get size---
        Dim t() As String = Split(activeCtrl.Tag, vbCrLf)
        Dim i As Integer
        ttHeight = 19
        For i = 1 To UBound(t)
            ttHeight += 15
        Next
        ToolTip1.Show(activeCtrl.Tag, Me.StatusStrip1, New Point(activeCtrl.Bounds.Left, activeCtrl.Bounds.Top - ttHeight), 3000)
    End Sub
    Private Sub ClearBoardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearBoardToolStripMenuItem.Click
        frmClear(blnSamurai:=blnSamurai)
    End Sub
    Private Sub GameToolStripMenuItem_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles GameToolStripMenuItem.DropDownOpening, SolveToolStripMenuItem.DropDownOpening, ExtraToolStripMenuItem.DropDownOpening
        'TODO - update if new menus or menu items added
        Dim i As Integer
        Dim arrSolvers() As String = Split(My.Settings._EnabledSolvers, arrDivider)
        Dim arrMethods() As String
        Me.SolveUpToToolStripMenuItem.DropDownItems.Clear()
        Me.FindOnlyToolStripMenuItem.DropDownItems.Clear()

        For i = 0 To UBound(arrSolvers)
            arrMethods = Split(arrSolvers(i), ":")
            Dim nm As New System.Windows.Forms.ToolStripMenuItem
            Dim fm As New System.Windows.Forms.ToolStripMenuItem
            With nm
                .Text = arrMethods(2)
                .Enabled = arrMethods(1)
            End With
            With fm
                .Text = arrMethods(2)
                .Enabled = arrMethods(1)
                .Tag = arrSolvers(i)
            End With
            AddHandler nm.Click, AddressOf SolveUpToClick
            AddHandler fm.Click, AddressOf FindOnlyClick
            Me.SolveUpToToolStripMenuItem.DropDownItems.Add(nm)
            Me.FindOnlyToolStripMenuItem.DropDownItems.Add(fm)
        Next

        If countClues(My.Settings._clueColour) > 0 Then
            Me.CopyCluesToolStripMenuItem.Enabled = True
        Else
            Me.CopyCluesToolStripMenuItem.Enabled = False
        End If
        If blnSetup Then
            Me.CopyFilledToolStripMenuItem.Enabled = False
            Me.CopySolutionToolStripMenuItem.Enabled = False
        Else
            If strPuzzleSolution <> "" Then
                Me.CopyFilledToolStripMenuItem.Enabled = True
                Me.CopySolutionToolStripMenuItem.Enabled = True
            Else
                Me.CopyFilledToolStripMenuItem.Enabled = False
                Me.CopySolutionToolStripMenuItem.Enabled = False
            End If
        End If

        If strPuzzleSolution <> "" Then
            Me.ShowSolutionToolStripMenuItem.Enabled = True
            Me.SolveUpToToolStripMenuItem.Enabled = True
            Me.FindOnlyToolStripMenuItem.Enabled = True
            Me.CheckConsistencyToolStripMenuItem.Enabled = True
            Me.HintToolStripMenuItem.Enabled = True
            Me.NextStepToolStripMenuItem.Enabled = True
            Me.ShowAllStepsToolStripMenuItem.Enabled = True
        Else
            Me.ShowSolutionToolStripMenuItem.Enabled = False
            Me.SolveUpToToolStripMenuItem.Enabled = False
            Me.FindOnlyToolStripMenuItem.Enabled = False
            Me.CheckConsistencyToolStripMenuItem.Enabled = False
            Me.HintToolStripMenuItem.Enabled = False
            Me.NextStepToolStripMenuItem.Enabled = False
            Me.ShowAllStepsToolStripMenuItem.Enabled = False
        End If

        If blnSamurai Or blnSetup Or strPuzzleSolution = "" Then
            Me.OptimisePuzzleSymmetryToolStripMenuItem.Enabled = False
            Me.OptimisePuzzleToolStripMenuItem.Enabled = False
        Else
            '---check if puzzle has symmetry---
            Dim strGame As String = copyGameStr(intType:=2)
            Dim cs As New Symmetry
            cs.blnSamurai = False
            cs.strGrid = strGame
            Dim intSymmetry As Integer
            Dim s() As String
            s = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
            intSymmetry = Array.IndexOf(s, cs.checkSymmetry)
            If intSymmetry = 0 Then
                Me.OptimisePuzzleSymmetryToolStripMenuItem.Enabled = False
                Me.OptimisePuzzleSymmetryToolStripMenuItem.Text = "Optimise puzzle (maintain symmetry)"
            Else
                Me.OptimisePuzzleSymmetryToolStripMenuItem.Enabled = True
                Me.OptimisePuzzleSymmetryToolStripMenuItem.Text = "Optimise puzzle (" & cs.checkSymmetry & " symmetry)"
                Me.OptimisePuzzleSymmetryToolStripMenuItem.Tag = cs.checkSymmetry
            End If
            Me.OptimisePuzzleToolStripMenuItem.Enabled = True
        End If

        If blnSolved() Then
            Me.ShowSolutionToolStripMenuItem.Enabled = False
            Me.SolveUpToToolStripMenuItem.Enabled = False
            Me.FindOnlyToolStripMenuItem.Enabled = False
            Me.CheckConsistencyToolStripMenuItem.Enabled = False
            Me.HintToolStripMenuItem.Enabled = False
            Me.NextStepToolStripMenuItem.Enabled = False
            Me.ShowAllStepsToolStripMenuItem.Enabled = False
        End If

    End Sub
    Private Sub SetupLabel_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles SetupLabel.DoubleClick
        If Not blnSamurai Then Exit Sub
        Dim blnUnique As Boolean
        Dim intCount As Integer
        frmProgress.strGrid = Replace(copyGameStr(intType:=2), vbCrLf, "")
        frmProgress.intQuit = 2
        frmProgress.ShowDialog(Me)
        Me.Refresh()
        blnUnique = frmProgress.blnUnique
        intCount = frmProgress.intCountSolutions
        strPuzzleSolution = frmProgress.strGameSolution
    End Sub
    Private Sub BatchSolveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BatchSolveToolStripMenuItem.Click
        frmBatch.ShowDialog(Me)
        StopThreadWork = False
    End Sub
    Private Sub BatchOptimiseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BatchOptimiseToolStripMenuItem.Click
        frmBatchOptimise.ShowDialog(Me)
        StopThreadWork = False
    End Sub
    Private Sub OptimisePuzzleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptimisePuzzleToolStripMenuItem.Click
        Cursor.Current = Cursors.WaitCursor
        Dim strGame As String = copyGameStr(intType:=2)
        Dim opt As New clsSudokuOptimise
        opt.strGrid = strGame
        opt.intType = 1
        If blnSamurai Then opt.intType = 2
        opt.OptimisePuzzle(intSymmetry:=0)
        If opt.isOptimised Then
            My.Computer.Clipboard.SetText(opt.strOptimised, TextDataFormat.Text)
            ShowMessageLabel("Puzzle optimised and copied to clipboard (" & strMulti(opt.intCluesRemoved, "clue", False) & ").", "info")
        Else
            ShowMessageLabel("Puzzle cannot be optimised", "alert")
        End If
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub OptimisePuzzleSymmetryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptimisePuzzleSymmetryToolStripMenuItem.Click
        Cursor.Current = Cursors.WaitCursor
        Dim strGame As String = copyGameStr(intType:=2)
        Dim cs As New Symmetry
        cs.blnSamurai = blnSamurai
        cs.strGrid = strGame
        Dim intSymmetry As Integer
        Dim s() As String
        s = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
        intSymmetry = Array.IndexOf(s, cs.checkSymmetry)

        Dim opt As New clsSudokuOptimise
        opt.strGrid = strGame
        opt.intType = 1
        If blnSamurai Then opt.intType = 2
        opt.OptimisePuzzle(intSymmetry:=intSymmetry)
        If opt.isOptimised Then
            My.Computer.Clipboard.SetText(opt.strOptimised, TextDataFormat.Text)
            ShowMessageLabel("Puzzle optimised to maintain " & cs.checkSymmetry & " symmetry and copied to clipboard (" & strMulti(opt.intCluesRemoved, "clue", False) & ").", "info")
        Else
            ShowMessageLabel("Puzzle cannot be optimised to maintain " & cs.checkSymmetry & " symmetry", "alert")
        End If
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub NextStepToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextStepToolStripMenuItem.Click
        If Not blnConsistent() Then Exit Sub

        If Not My.Settings._blnCandidates Then
            Dim style As MsgBoxStyle
            Dim response As MsgBoxResult
            style = MsgBoxStyle.Information Or MsgBoxStyle.YesNo
            '---display warning message---
            response = MsgBox("It is highly recommended that candidates are displayed. Do you wanr to do this?", style)
            If response = MsgBoxResult.Yes Then
                setCandidates(True)
            End If
        End If

        Dim strClues As String = ""
        Dim strCandidates As String = ""
        getGameStr(strFilled:=strClues, strCandidates:=strCandidates)
        If strClues = "" Or strCandidates = "" Then Exit Sub
        Dim solver As New clsSudokuSolver
        Select Case blnSamurai
            Case True
                solver.blnClassic = False
            Case False
                solver.blnClassic = True
        End Select
        With solver
            .blnStep = True
            .strGrid = strClues
            .strCandidates = strCandidates
            .strInputGameSolution = strPuzzleSolution
            .vsSolvers = My.Settings._EnabledSolvers
        End With

        solver._vsUnique()
    End Sub
    Private Sub ShowAllStepsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowAllStepsToolStripMenuItem.Click
        If Not blnConsistent() Then Exit Sub

        StopShowAllSteps = False
        blnAllSteps = True

        If Not My.Settings._blnCandidates Then
            Dim style As MsgBoxStyle
            Dim response As MsgBoxResult
            style = MsgBoxStyle.Information Or MsgBoxStyle.YesNo
            '---display warning message---
            response = MsgBox("It is highly recommended that candidates are displayed. Do you wanr to do this?", style)
            If response = MsgBoxResult.Yes Then
                setCandidates(True)
            End If
        End If

        Dim strClues As String = ""
        Dim strCandidates As String = ""
        Dim solver As New clsSudokuSolver
        Select Case blnSamurai
            Case True
                solver.blnClassic = False
            Case False
                solver.blnClassic = True
        End Select

        While Not blnSolved()
            If StopShowAllSteps = True Then Exit While
            getGameStr(strFilled:=strClues, strCandidates:=strCandidates)
            If strClues = "" Or strCandidates = "" Then Exit While

            With solver
                .blnStep = True
                .strGrid = strClues
                .strCandidates = strCandidates
                .strInputGameSolution = strPuzzleSolution
                .vsSolvers = My.Settings._EnabledSolvers
            End With

            solver._vsUnique()

        End While

        StopShowAllSteps = False
        blnAllSteps = False
    End Sub
    Function blnSolved() As Boolean
        Dim intClues As Integer
        intClues += countClues(My.Settings._solveColour)
        intClues += countClues(My.Settings._clueColour)
        intClues += countClues(My.Settings._userColour)
        Select Case blnSamurai
            Case True
                If intClues = 369 Then Return True
            Case False
                If intClues = 81 Then Return True
        End Select
    End Function
    Private Sub OptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionToolStripMenuItem.Click
        frmOptions.ShowDialog(Me)
    End Sub
    Private Function strReturnSolution(ByVal intCell As Integer, ByVal blnSamurai As Boolean, ByVal strSolution As String) As String
        strReturnSolution = ""
        Dim intGrid As Integer
        Dim intPtr As Integer
        If strSolution = "" Then Exit Function
        strSolution = Replace(strSolution, vbCrLf, "")
        strSolution = Replace(strSolution, vbLf, "")
        Select Case blnSamurai
            Case True
                intGrid = intSamuraiGrid(intSamuraiCell:=intCell)
                intPtr = intSamuraitoClassic(intSamuraiCell:=intCell)
                intPtr = ((intGrid - 1) * 81) + intPtr
                If Len(strSolution) <> 405 Then Exit Function
            Case False
                intPtr = intCell
                If Len(strSolution) <> 81 Then Exit Function
        End Select
        strReturnSolution = Mid(strSolution, intPtr, 1)
    End Function
    Private Sub RevealSolutionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RevealSolutionToolStripMenuItem.Click
        Dim strSolution As String = strReturnSolution(intCell:=intActiveCell, blnSamurai:=blnSamurai, strSolution:=strPuzzleSolution)
        Dim sColour As Color = Color.FromName(My.Settings._solveColour)
        Dim strMoves As String = ""
        nakedSingle_doubleClick(intCell:=intActiveCell, intValue:=CInt(strSolution), strSolvedBy:="P")
    End Sub
    Private Sub ExtraToolStripMenuItem_DropDownOpening(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExtraToolStripMenuItem.DropDownOpening
        If blnSamurai Then
            Me.PuzzleLibraryToolStripMenuItem.Visible = False
        Else
            Dim i As Integer
            Dim j As Integer
            Dim intCount As Integer
            Dim arrDifficulty() = arrDifficultStr
            Dim sym() As String
            sym = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
            Me.PuzzleLibraryToolStripMenuItem.DropDownItems.Clear()
            Me.PuzzleLibraryToolStripMenuItem.Visible = True
            For i = 0 To UBound(arrDifficulty)
                Dim dm As New System.Windows.Forms.ToolStripMenuItem
                With dm
                    intCount = CountPuzzles("*_" & strFirstLetters(arrDifficultStr(i)) & "_*.sdm")
                    .Enabled = False
                    .Text = arrDifficulty(i)
                    If intCount > 0 Then
                        .Enabled = True
                        .Text += " (" & intCount & ")"
                    End If
                    .Tag = "*_" & strFirstLetters(arrDifficultStr(i)) & "_*.sdm"
                End With
                AddHandler dm.Click, AddressOf GetPuzzleLibrary
                Me.PuzzleLibraryToolStripMenuItem.DropDownItems.Add(dm)

                For j = 0 To UBound(sym)
                    Dim sm As New System.Windows.Forms.ToolStripMenuItem
                    With sm
                        intCount = CountPuzzles("*_" & strFirstLetters(arrDifficultStr(i)) & "_" & sym(j) & "_symmetry_*.sdm")
                        .Enabled = False
                        .Text = sym(j) & " symmetry"
                        If intCount > 0 Then
                            .Enabled = True
                            .Text += " (" & intCount & ")"
                        End If
                        .Tag = "*_" & strFirstLetters(arrDifficultStr(i)) & "_" & sym(j) & "_symmetry_*.sdm"
                    End With
                    AddHandler sm.Click, AddressOf GetPuzzleLibrary
                    dm.DropDownItems.Add(sm)
                Next
            Next
        End If
    End Sub
    Private Sub GetPuzzleLibrary(ByVal sender As Object, ByVal e As System.EventArgs)
        strGetPuzzle(sender.tag)
    End Sub
    Private Function strGetPuzzle(ByVal strFilter) As String
        strGetPuzzle = ""
        Dim strContents As String = ""
        If My.Computer.FileSystem.DirectoryExists(dirGames) Then
            Dim fileList As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirGames, FileIO.SearchOption.SearchTopLevelOnly, strFilter)
            If fileList.Count > 0 Then
                strGetPuzzle = fileList(0)
                strContents += My.Computer.FileSystem.ReadAllText(strGetPuzzle)
            End If
        End If
        If strGetPuzzle = "" Or strContents = "" Then
            MsgBox("Cannot find relevant puzzle - please check " & dirGames, MsgBoxStyle.Exclamation, My.Application.Info.Title)
            Exit Function
        End If
        inputPuzzle(strInput:=Microsoft.VisualBasic.Left(strContents, 81))
        Me.blnSetup = False
        Me.SetupLabel.Visible = False
        '---delete puzzle---
        Try
            My.Computer.FileSystem.DeleteFile(strGetPuzzle)
        Catch
        End Try
    End Function
    Private Sub Optimise_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptimiseToolStripMenuItem.DropDownOpening
        Dim intDifficulty As Integer = -1
        If strPuzzleSolution <> "" And Not blnSolved() Then
            Dim solver As New clsSudokuSolver
            Dim strClues As String = copyGameStr(intType:=2)
            If strClues <> "" Then
                solver.blnClassic = True
                If blnSamurai Then
                    solver.blnClassic = False
                End If
                solver.strGrid = Replace(strClues, vbCrLf, "")
                solver.strCandidates = ""
                solver.vsSolvers = My.Settings._defaultSolvers
                solver.strInputGameSolution = strPuzzleSolution
                If solver._vsUnique() Then
                    intDifficulty = solver.intDifficulty
                End If
            End If
        End If

        Dim i As Integer
        Me.OptimisePuzzleToolStripMenuItem.DropDownItems.Clear()
        Me.OptimisePuzzleSymmetryToolStripMenuItem.DropDownItems.Clear()
        For i = 0 To UBound(arrDifficultStr)
            Dim nm As New System.Windows.Forms.ToolStripMenuItem
            With nm
                .Text = arrDifficultStr(i)
                .Tag = i + 1
                .Enabled = False
                If intDifficulty >= -1 AndAlso strDifficult2Int(.Text) >= intDifficulty Then
                    .Enabled = True
                End If
            End With
            Dim nm1 As New System.Windows.Forms.ToolStripMenuItem
            With nm1
                .Text = arrDifficultStr(i)
                .Tag = i + 1
                .Enabled = False
                If intDifficulty >= -1 AndAlso strDifficult2Int(.Text) >= intDifficulty Then
                    .Enabled = True
                End If
            End With
            AddHandler nm.Click, AddressOf OptimseUpTo
            AddHandler nm1.Click, AddressOf OptimseUpToSymmetry
            If Me.OptimisePuzzleToolStripMenuItem.Enabled Then Me.OptimisePuzzleToolStripMenuItem.DropDownItems.Add(nm)
            If Me.OptimisePuzzleSymmetryToolStripMenuItem.Enabled Then Me.OptimisePuzzleSymmetryToolStripMenuItem.DropDownItems.Add(nm1)
        Next
    End Sub
    Private Sub OptimseUpTo(ByVal sender As Object, ByVal e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        Dim strGame As String = copyGameStr(intType:=2)
        Dim opt As New clsSudokuOptimise
        opt.strGrid = strGame
        opt.intType = 1
        If blnSamurai Then opt.intType = 2
        opt.OptimisePuzzle(intSymmetry:=0, intDifficulty:=sender.tag)
        If opt.isOptimised And CInt(sender.tag) = strDifficult2Int(opt.strDifficulty) Then
            My.Computer.Clipboard.SetText(opt.strOptimised, TextDataFormat.Text)
            ShowMessageLabel("Optimised to " & LCase(intDifficult2String(sender.tag)) & " and copied to clipboard (" & strMulti(opt.intCluesRemoved, "clue", False) & ").", "info")
        Else
            ShowMessageLabel("Cannot be optimised to " & LCase(intDifficult2String(sender.tag)), "alert")
        End If
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub OptimseUpToSymmetry(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim p As ToolStripMenuItem = sender
        Dim strSymmetry As String = p.OwnerItem.Tag
        Dim intSymmetry As Integer
        Dim s() As String
        s = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
        intSymmetry = Array.IndexOf(s, strSymmetry)
        Cursor.Current = Cursors.WaitCursor
        Dim strGame As String = copyGameStr(intType:=2)
        Dim opt As New clsSudokuOptimise
        opt.strGrid = strGame
        opt.intType = 1
        If blnSamurai Then opt.intType = 2
        opt.OptimisePuzzle(intSymmetry:=intSymmetry, intDifficulty:=sender.tag)
        If opt.isOptimised And CInt(sender.tag) = strDifficult2Int(opt.strDifficulty) Then
            My.Computer.Clipboard.SetText(opt.strOptimised, TextDataFormat.Text)
            ShowMessageLabel("Optimised to " & LCase(intDifficult2String(sender.tag)) & " and copied to clipboard (" & strMulti(opt.intCluesRemoved, "clue", False) & ").", "info")
        Else
            ShowMessageLabel("Cannot be optimised to " & LCase(intDifficult2String(sender.tag)) & " with " & strSymmetry & " symmetry intact", "alert")
        End If
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox.ShowDialog(Me)
    End Sub
End Class


Option Strict Off
Public Class frmOptions
    Private borderCol As String = My.Settings._borderColour
    Private bgCol As String = My.Settings._bgColour
    Private clueCol As String = My.Settings._clueColour
    Private solveCol As String = My.Settings._solveColour
    Private userCol As String = My.Settings._userColour
    Private nakedCol As String = My.Settings._nakedColour
    Private pCol As String = My.Settings._pColour
    Private nCol As String = My.Settings._nColour
    Private cCol As String = My.Settings._cColour
    Private rCol As String = My.Settings._rColour
    Private blnLoading As Boolean
    Private blnSwap As Boolean
    Private F() As String
    Private arrMethods() As String
    Private arrDetails() As String
    Private arrDifficulty() As String
    Private Sub frmOptions_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        UpdateColours()
        UpdateGenerationOptions()
        Dim i As Integer
        Dim j As Integer
        Dim blnEnabled As Boolean
        Dim strMethods As String = ""
        Dim strMethod As String = ""
        For i = 0 To CheckedListBox1.Items.Count - 1
            blnEnabled = Me.CheckedListBox1.GetItemChecked(i)
            strMethod = Me.CheckedListBox1.Items.Item(i).ToString
            j += 1
            If j = 1 Then
                strMethods += arrDifficulty(i) & ":" & blnEnabled & ":" & strMethod
            Else
                strMethods += "," & arrDifficulty(i) & ":" & blnEnabled & ":" & strMethod
            End If
        Next
        My.Settings._EnabledSolvers = strMethods
        frmInitial.NotifyIconGenerator.Visible = False
        frmInitial.NotifyIconGenerator.Visible = True
    End Sub
    Private Sub frmOptions_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.TabControl1.TabPages.Item(0).Focus()
        blnLoading = True
        ComboBox1.SelectedIndex = 0
        borderCol = My.Settings._borderColour
        bgCol = My.Settings._bgColour
        clueCol = My.Settings._clueColour
        solveCol = My.Settings._solveColour
        userCol = My.Settings._userColour
        nakedCol = My.Settings._nakedColour
        pCol = My.Settings._pColour
        nCol = My.Settings._nColour
        rCol = My.Settings._rColour
        cCol = My.Settings._cColour
        Me.CheckBox1.Checked = My.Settings._blnCandidates
        Me.CheckBox2.Checked = My.Settings._ShowNS
        Me.cbManualInputHelp.Checked = My.Settings._blnManualInputSuggest
        Me.CheckBox2.Enabled = My.Settings._blnCandidates
        Me.ColourCombobox6.Visible = My.Settings._ShowNS
        Me.Label6.Visible = My.Settings._ShowNS
        Me.GroupBox2.Height = 140
        If Not My.Settings._ShowNS Then
            Me.GroupBox2.Height = 107
        End If
        Me.GroupBox4.Top = Me.GroupBox2.Top + Me.GroupBox2.Height + 6
        Me.TextBox1.Text = My.Settings._intMaxGenerate
        blnLoading = False

        Dim a() As String
        Dim i As Integer
        a = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
        Me.clbClassicSymmetry.Items.Clear()
        Me.clbSamuraiSymmetry.Items.Clear()

        For i = 0 To a.Length - 1
            Me.clbClassicSymmetry.Items.Add(a(i))
            Me.clbSamuraiSymmetry.Items.Add(a(i))
            If intGetBit(i + 1) And My.Settings._GenClassicSymmetry Then
                Me.clbClassicSymmetry.SetItemChecked(i, True)
            End If
            If intGetBit(i + 1) And My.Settings._GenSamuraiSymmetry Then
                Me.clbSamuraiSymmetry.SetItemChecked(i, True)
            End If
        Next

        a = arrDifficultStr
        Me.clbClassicDifficulty.Items.Clear()
        Me.clbSamuraiDifficulty.Items.Clear()
        For i = 0 To a.Length - 1
            Me.clbClassicDifficulty.Items.Add(a(i))
            Me.clbSamuraiDifficulty.Items.Add(a(i))
            If intGetBit(i + 1) And My.Settings._GenClassicDifficulty Then
                Me.clbClassicDifficulty.SetItemChecked(i, True)
            End If
            If intGetBit(i + 1) And My.Settings._GenSamuraiDifficulty Then
                Me.clbSamuraiDifficulty.SetItemChecked(i, True)
            End If
        Next

    End Sub
    Private Sub UpdateGenerationOptions()
        Dim i As Integer
        Dim j As Integer
        For i = 0 To Me.clbClassicDifficulty.Items.Count - 1
            If Me.clbClassicDifficulty.GetItemChecked(i) Then
                j += intGetBit(i + 1)
            End If
        Next
        My.Settings._GenClassicDifficulty = j
        j = 0
        For i = 0 To Me.clbSamuraiDifficulty.Items.Count - 1
            If Me.clbSamuraiDifficulty.GetItemChecked(i) Then
                j += intGetBit(i + 1)
            End If
        Next
        My.Settings._GenSamuraiDifficulty = j
        j = 0
        For i = 0 To Me.clbClassicSymmetry.Items.Count - 1
            If Me.clbClassicSymmetry.GetItemChecked(i) Then
                j += intGetBit(i + 1)
            End If
        Next
        My.Settings._GenClassicSymmetry = j
        j = 0
        For i = 0 To Me.clbSamuraiSymmetry.Items.Count - 1
            If Me.clbSamuraiSymmetry.GetItemChecked(i) Then
                j += intGetBit(i + 1)
            End If
        Next
        My.Settings._GenSamuraiSymmetry = j
        j = 0

        My.Settings._intMaxGenerate = Me.TrackBar1.Value

    End Sub
    Private Sub CCB_IndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ColourCombobox1.SelectedIndexChanged, ColourCombobox5.SelectedIndexChanged, ColourCombobox4.SelectedIndexChanged, ColourCombobox3.SelectedIndexChanged, ColourCombobox2.SelectedIndexChanged, ColourCombobox6.SelectedIndexChanged, ColourCombobox7.SelectedIndexChanged, ColourCombobox8.SelectedIndexChanged, ColourCombobox9.SelectedIndexChanged, ColourCombobox10.SelectedIndexChanged
        Dim strCol As String = sender.parent.selectedcolor
        If strCol = "" Then Exit Sub
        Select Case sender.parent.name
            Case "ColourCombobox1"
                My.Settings._borderColour = strCol
            Case "ColourCombobox2"
                My.Settings._bgColour = strCol
            Case "ColourCombobox3"
                My.Settings._clueColour = strCol
            Case "ColourCombobox4"
                My.Settings._solveColour = strCol
            Case "ColourCombobox5"
                My.Settings._userColour = strCol
            Case "ColourCombobox6"
                My.Settings._nakedColour = strCol
            Case "ColourCombobox7"
                My.Settings._cColour = strCol
            Case "ColourCombobox8"
                My.Settings._rColour = strCol
            Case "ColourCombobox9"
                My.Settings._pColour = strCol
            Case "ColourCombobox10"
                My.Settings._nColour = strCol
        End Select
        colourPreview()
    End Sub
    Private Sub UpdateColours()
        Dim sc As SudokuCell
        Dim i As Integer
        Dim g As Integer
        Dim ptr As Integer
        Dim intStart = 1
        Dim intEnd = 1
        If blnSamurai Then
            intEnd = 5
        End If
        For g = intStart To intEnd
            For i = 1 To 81
                If blnSamurai Then
                    ptr = intSamuraiOffset(i, g)
                Else
                    ptr = i
                End If
                sc = getCtrl("SudokuCell" & ptr)
                With sc
                    If borderCol <> My.Settings._borderColour Then
                        If frmGame.intMultiCell(0).IndexOf(i) > -1 Then
                            .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intM0_BorderTransparency)
                        Else
                            .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                        End If
                    End If
                    If bgCol <> My.Settings._bgColour Then
                        If frmGame.intMultiCell(0).IndexOf(i) > -1 Then
                            .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intMO_BGTransparency)
                        Else
                            .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                        End If
                    End If
                    If clueCol <> My.Settings._clueColour Then
                        If .sc_TextColour = Color.FromName(clueCol) Then
                            .sc_TextColour = Color.FromName(My.Settings._clueColour)
                        End If
                    End If
                    If solveCol <> My.Settings._solveColour Then
                        If .sc_TextColour = Color.FromName(solveCol) Then
                            .sc_TextColour = Color.FromName(My.Settings._solveColour)
                        End If
                    End If
                    If userCol <> My.Settings._userColour Then
                        If .sc_TextColour = Color.FromName(userCol) Then
                            .sc_TextColour = Color.FromName(My.Settings._userColour)
                        End If
                    End If
                    If nakedCol <> My.Settings._nakedColour Then
                        If blnSingleBit(sc.sc_IntCandidate) Then
                            .Invalidate()
                        End If
                    End If
                End With
            Next
        Next
    End Sub
    Private Function getPreviewCtrl(ByVal strCtrl As String) As SudokuCell
        getPreviewCtrl = Nothing
        Dim sc() As Control = Me.Controls.Find(strCtrl, True)
        If sc.Length > 0 Then
            getPreviewCtrl = CType(sc(0), SudokuCell)
        End If
    End Function
    Private Sub colourPreview()
        Dim i As Integer
        Dim sc As SudokuCell
        For i = 1 To 9
            sc = getPreviewCtrl("SudokuCell" & i)
            If sc Is Nothing Then
            Else
                sc.sc_TextColour = transparentColor(Color.FromName(My.Settings._clueColour), 1)
                sc.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                sc.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                Select Case i
                    Case 1
                        sc.sc_HighlightColour1 = transparentColor(Color.FromName(My.Settings._pColour), 1)
                        sc.sc_IntHighlightCandidate1 = 256
                    Case 3
                        sc.sc_HighlightColour1 = transparentColor(Color.FromName(My.Settings._rColour), 1)
                        sc.sc_IntHighlightCandidate1 = 272
                        sc.sc_HighlightColour2 = transparentColor(Color.FromName(My.Settings._cColour), 1)
                        sc.sc_IntHighlightCandidate2 = 2
                    Case 4
                        sc.sc_HighlightColour1 = transparentColor(Color.FromName(My.Settings._nColour), 1)
                        sc.sc_IntHighlightCandidate1 = 256
                    Case 6
                        sc.sc_TextColour = transparentColor(Color.FromName(My.Settings._solveColour), 1)
                    Case 7
                        sc.sc_TextColour = transparentColor(Color.FromName(My.Settings._userColour), 1)
                End Select
            End If
        Next
    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click, btnOK2.Click, btnOK3.Click, btnOK4.Click, btnOK5.Click
        Me.Close()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        If blnLoading Then Exit Sub

        If ComboBox1.SelectedIndex = 0 Then Exit Sub

        ColourCombobox1.ForceReload()
        ColourCombobox2.ForceReload()
        ColourCombobox3.ForceReload()
        ColourCombobox4.ForceReload()
        ColourCombobox5.ForceReload()
        ColourCombobox6.ForceReload()
        ColourCombobox7.ForceReload()
        ColourCombobox8.ForceReload()
        ColourCombobox9.ForceReload()
        ColourCombobox10.ForceReload()

        Select Case ComboBox1.SelectedIndex
            Case 1
                '---green---
                '---borderColour---
                ColourCombobox1.SelectedColor = "OliveDrab"
                '---bgColour---
                ColourCombobox2.SelectedColor = "DarkGreen"
                '---clueColour---
                ColourCombobox3.SelectedColor = "Black"
                '---progam solved colour---
                ColourCombobox4.SelectedColor = "DarkOliveGreen"
                '---userColour---
                ColourCombobox5.SelectedColor = "DimGray"
                '---nakedColour---
                ColourCombobox6.SelectedColor = "ForestGreen"
                '---correct colour---
                ColourCombobox7.SelectedColor = "DarkGreen"
                '---remove colour---
                ColourCombobox8.SelectedColor = "DimGray"
                '---positive colour---
                ColourCombobox9.SelectedColor = "DarkGreen"
                '---negative colour---
                ColourCombobox10.SelectedColor = "ForestGreen"
            Case 2
                '---blue---
                '---borderColour---
                ColourCombobox1.SelectedColor = "SteelBlue"
                '---bgColour---
                ColourCombobox2.SelectedColor = "DodgerBlue"
                '---clueColour---
                ColourCombobox3.SelectedColor = "Black"
                '---progam solved colour---
                ColourCombobox4.SelectedColor = "DimGray"
                '---userColour---
                ColourCombobox5.SelectedColor = "DodgerBlue"
                '---nakedColour---
                ColourCombobox6.SelectedColor = "DarkBlue"
                '---correct colour---
                ColourCombobox7.SelectedColor = "DarkBlue"
                '---remove colour---
                ColourCombobox8.SelectedColor = "DodgerBlue"
                '---positive colour---
                ColourCombobox9.SelectedColor = "SteelBlue"
                '---negative colour---
                ColourCombobox10.SelectedColor = "RoyalBlue"
            Case 3
                '---pink---
                '---borderColour---
                ColourCombobox1.SelectedColor = "Magenta"
                '---bgColour---
                ColourCombobox2.SelectedColor = "Orchid"
                '---clueColour---
                ColourCombobox3.SelectedColor = "Black"
                '---progam solved colour---
                ColourCombobox4.SelectedColor = "DimGray"
                '---userColour---
                ColourCombobox5.SelectedColor = "DeepPink"
                '---nakedColour---
                ColourCombobox6.SelectedColor = "Fuchsia"
                '---correct colour---
                ColourCombobox7.SelectedColor = "MediumOrchid"
                '---remove colour---
                ColourCombobox8.SelectedColor = "HotPink"
                '---positive colour---
                ColourCombobox9.SelectedColor = "DarkViolet"
                '---negative colour---
                ColourCombobox10.SelectedColor = "HotPink"
            Case 4
                '---contempory---
                '---borderColour---
                ColourCombobox1.SelectedColor = "SlateGray"
                '---bgColour---
                ColourCombobox2.SelectedColor = "Gray"
                '---clueColour---
                ColourCombobox3.SelectedColor = "Black"
                '---progam solved colour---
                ColourCombobox4.SelectedColor = "Crimson"
                '---userColour---
                ColourCombobox5.SelectedColor = "DimGray"
                '---nakedColour---
                ColourCombobox6.SelectedColor = "DarkRed"
                '---correct colour---
                ColourCombobox7.SelectedColor = "DimGray"
                '---remove colour---
                ColourCombobox8.SelectedColor = "Crimson"
                '---positive colour---
                ColourCombobox9.SelectedColor = "DarkGray"
                '---negative colour---
                ColourCombobox10.SelectedColor = "DimGray"
        End Select
    End Sub
    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        If CheckBox1.Checked = True Then
            CheckBox2.Enabled = True
        Else
            CheckBox2.Enabled = False
        End If
        frmGame.setCandidates(blnSwap:=True)
        My.Settings._blnCandidates = Me.CheckBox1.Checked
    End Sub
    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click
        My.Settings._ShowNS = Me.CheckBox2.Checked
        frmGame.Refresh()
        Me.ColourCombobox6.Visible = My.Settings._ShowNS
        Me.Label6.Visible = My.Settings._ShowNS
        Me.GroupBox2.Height = 140
        If Not My.Settings._ShowNS Then
            Me.GroupBox2.Height = 107
        End If
        Me.GroupBox4.Top = Me.GroupBox2.Top + Me.GroupBox2.Height + 6
        Me.GroupBox5.Top = Me.GroupBox4.Top + Me.GroupBox4.Height + 6
    End Sub
    Private Sub CheckedListBox1_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        If blnSwap Then Exit Sub
        Dim strSolve As String = CheckedListBox1.Items(e.Index).ToString
        Select Case strSolve
            Case "Brute Force"
                If CBool(e.NewValue) = False Then
                    e.NewValue = e.CurrentValue
                End If
            Case "Naked Single"
                If CBool(e.NewValue) = False Then
                    e.NewValue = e.CurrentValue
                End If
            Case Else
                If CBool(e.NewValue) Then
                    setButtons(-1)
                Else
                    setButtons(1)
                End If
        End Select
    End Sub
    Private Sub GenerationListBox_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles clbClassicDifficulty.ItemCheck, clbSamuraiDifficulty.ItemCheck, clbClassicSymmetry.ItemCheck, clbSamuraiSymmetry.ItemCheck
        Dim clb As CheckedListBox = sender
        If clb.CheckedItems.Count = 1 Then
            If CBool(e.NewValue) = False Then
                '---force at least one option to remain checked---
                e.NewValue = CheckState.Checked
            End If
        End If
    End Sub
    Private Sub GenerationListBoxPrompt_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles clbClassicDifficulty.ItemCheck
        Dim clb As CheckedListBox = sender
        If clb.CheckedItems.Count = 2 Then
            If CBool(e.NewValue) = False Then
                MsgBox("It is recommended you select multiple difficulty levels. Selecting only a single difficulty level may mean in some cases the generator takes a long time to find puzzles of the desired difficulty", MsgBoxStyle.Information, My.Application.Info.Title & " - puzzle generator")
            End If
        End If
    End Sub
    Function setButtons(ByVal intOffset As Integer) As Boolean
        If countSolveMethods() = countCheckedMethods() - intOffset Then
            btnSelectAll.Enabled = False
        Else
            btnSelectAll.Enabled = True
        End If
        If countCheckedMethods() - intOffset <= 2 Then
            btnDeselectAll.Enabled = False
        Else
            btnDeselectAll.Enabled = True
        End If
    End Function
    Private Function ToggleAll(ByVal blnValue As Boolean) As Boolean
        Dim i As Integer
        Dim strSolve As String
        For i = 0 To CheckedListBox1.Items.Count - 1
            strSolve = CheckedListBox1.Items.Item(i).ToString()
            CheckedListBox1.SetItemChecked(i, blnValue)
        Next
    End Function
    Private Function countCheckedMethods() As Integer
        countCheckedMethods = CheckedListBox1.CheckedItems.Count
    End Function
    Private Function countSolveMethods() As Integer
        countSolveMethods = CheckedListBox1.Items.Count
    End Function
    Private Sub btnSelectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click
        ToggleAll(True)
    End Sub
    Private Sub btnDeselectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeselectAll.Click
        ToggleAll(False)
    End Sub
    Private Function setupSolveMethods(ByVal strType As String) As Boolean
        Dim strSolvers As String = ""
        Select Case strType
            Case "Enabled"
                strSolvers = My.Settings._EnabledSolvers
            Case "Default"
                strSolvers = My.Settings._defaultSolvers
        End Select
        Me.CheckedListBox1.Items.Clear()
        Dim i As Integer
        arrMethods = Split(strSolvers, ",")
        ReDim arrDifficulty(UBound(arrMethods))
        Me.CheckedListBox1.Height = arrMethods.Length * 15.3
        For i = 0 To UBound(arrMethods)
            arrDetails = Split(arrMethods(i), ":")
            arrDifficulty(i) = arrDetails(0)
            CheckedListBox1.Items.Add(arrDetails(2))
            Me.CheckedListBox1.SetItemChecked(i, False)
            If arrDetails(1) = True Then Me.CheckedListBox1.SetItemChecked(i, True)
        Next
        Me.SolveDifficulty.Items.Clear()
        For i = 1 To 4
            Me.SolveDifficulty.Items.Add(intDifficult2String(i))
        Next
        Me.SolveDifficulty.SelectedIndex = 0
        btnSelectAll.Enabled = False
        btnDeselectAll.Enabled = False
        setButtons(0)
        Me.CheckedListBox1.SelectedIndex = 0
    End Function
    Private Sub frmOptions_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, Me.Shown
        setupSolveMethods(strType:="Enabled")
        Me.LibraryLocationLink.Text = dirGames
    End Sub
    Private Sub btnPlus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPlus.Click
        SwapListItem(-1)
    End Sub
    Private Sub btnMinus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMinus.Click
        SwapListItem(1)
    End Sub
    Private Function SwapListItem(ByVal intMove As Integer) As Boolean
        blnSwap = True
        Dim aName As String
        Dim bName As String
        Dim aDifficulty As Integer
        Dim bDifficulty As Integer
        Dim aChecked As Boolean
        Dim bChecked As Boolean
        Dim intIndex As Integer
        intIndex = Me.CheckedListBox1.SelectedIndex
        If (intMove < 0 And intIndex > 0) Or (intMove > 0 And intIndex >= 0 And intIndex < Me.CheckedListBox1.Items.Count - 1) Then
            aName = Me.CheckedListBox1.Items(intIndex)
            bName = Me.CheckedListBox1.Items(intIndex + intMove)
            aDifficulty = arrDifficulty(intIndex)
            bDifficulty = arrDifficulty(intIndex + intMove)
            '---swap difficulty---
            arrDifficulty(intIndex + intMove) = aDifficulty
            arrDifficulty(intIndex) = bDifficulty
            aChecked = Me.CheckedListBox1.GetItemChecked(intIndex)
            bChecked = Me.CheckedListBox1.GetItemChecked(intIndex + intMove)
            Me.CheckedListBox1.Items(intIndex) = bName
            Me.CheckedListBox1.SetItemChecked(intIndex, bChecked)
            Me.CheckedListBox1.Items(intIndex + intMove) = aName
            Me.CheckedListBox1.SetItemChecked(intIndex + intMove, aChecked)
            '---select item that was swapped---
            Me.CheckedListBox1.SetSelected(intIndex + intMove, True)
        End If
        blnSwap = False
    End Function
    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Me.SolveDifficulty.Enabled = True
        If CheckedListBox1.SelectedItem = "Brute force" Then
            Me.btnMinus.Enabled = False
            Me.btnPlus.Enabled = False
            Me.SolveDifficulty.Enabled = False
            Exit Sub
        End If
        Me.btnMinus.Enabled = True
        Me.btnPlus.Enabled = True
        If CheckedListBox1.SelectedIndex = 0 Then btnPlus.Enabled = False
        If CheckedListBox1.SelectedIndex = CheckedListBox1.Items.Count - 1 Then btnMinus.Enabled = False
        If CheckedListBox1.SelectedIndex > -1 Then Me.SolveDifficulty.SelectedIndex = CInt(arrDifficulty(CheckedListBox1.SelectedIndex) - 1)
    End Sub
    Private Sub SolveDifficulty_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SolveDifficulty.SelectedIndexChanged
        If CheckedListBox1.SelectedItem Is Nothing Then Exit Sub
        If CheckedListBox1.SelectedIndex > -1 Then
            arrDifficulty(CheckedListBox1.SelectedIndex) = SolveDifficulty.SelectedIndex + 1
        End If
    End Sub
    Private Sub btnDefault_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDefault.Click
        setupSolveMethods(strType:="Default")
    End Sub
    Private Sub cbManualInputHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbManualInputHelp.Click
        My.Settings._blnManualInputSuggest = cbManualInputHelp.Checked
    End Sub
    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Me.TextBox1.Text = TrackBar1.Value
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If blnLoading Then Exit Sub
        If Me.TrackBar1.Maximum = 0 Then Exit Sub
        Dim t As Integer
        If IsNumeric(Me.TextBox1.Text) Then
            t = CInt(Me.TextBox1.Text)
        Else
            If Me.TextBox1.Text = "" Then
                Me.TextBox1.Text = TrackBar1.Minimum
                TrackBar1.Value = TrackBar1.Minimum
                Me.TextBox1.Focus()
                Exit Sub
            End If
            MsgBox("Input must be numeric")
            Me.TextBox1.Text = TrackBar1.Value
            Me.TextBox1.Focus()
            Exit Sub
        End If
        If CInt(Me.TextBox1.Text) < TrackBar1.Minimum Or CInt(Me.TextBox1.Text) > TrackBar1.Maximum Then
            MsgBox("You must select a number between " & Me.TrackBar1.Minimum & " and " & Me.TrackBar1.Maximum)
            Me.TextBox1.Text = TrackBar1.Value
            Me.TextBox1.Focus()
            Exit Sub
        End If
        Me.TrackBar1.Value = t

        If t = 0 Then
            With frmInitial.NotifyIconGenerator
                .BalloonTipTitle = My.Application.Info.Title & " - puzzle generator"
                .BalloonTipText = "Puzzle generator is currently disabled as the number of puzzles to generate is set to zero."
                .BalloonTipIcon = ToolTipIcon.Warning
                .ShowBalloonTip(3000)
            End With
        End If


    End Sub
    Private Sub deleteLibrary(ByVal strDirectory As String)
        On Error Resume Next
        '---delete puzzles---
        If Not My.Computer.FileSystem.DirectoryExists(strDirectory) Then Exit Sub
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(strDirectory, FileIO.SearchOption.SearchAllSubDirectories, "*.sdm")
            My.Computer.FileSystem.DeleteFile(foundFile, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
        Next
    End Sub
    Private Sub btnClearLibrary_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearLibrary.Click
        Dim strDirectory = dirGames
        If strDirectory = "" Then Exit Sub
        Dim answer As MsgBoxResult = MsgBox("Are you sure you want to clear the puzzle library? This will delete all puzzle files from " & strDirectory, MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Clear pizzle library?")
        If answer = vbYes Then
            deleteLibrary(strDirectory)
        End If
    End Sub
    Private Sub btnPuzzleStats_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPuzzleStats.Click
        Dim intTotalCount As Integer = 0
        Dim blnFrmShown As Boolean
        intTotalCount = CountPuzzles("*.sdm")
        If intTotalCount < 1 Then GoTo errMsg
        frmPuzzleStats.ShowDialog(Me)
        blnFrmShown = frmPuzzleStats.blnShown
        If Not blnFrmShown Then GoTo errMsg
        Exit Sub
errMsg: MsgBox("No puzzle statistics are available. Please check " & dirGames & " exists and contains files with the '.sdm' extension")
        Exit Sub
    End Sub
    Private Sub LibraryLocationLink_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LibraryLocationLink.LinkClicked
        Dim blnExists As Boolean
        Try
            blnExists = My.Computer.FileSystem.DirectoryExists(LibraryLocationLink.Text)
        Catch
        End Try
        If blnExists Then
            System.Diagnostics.Process.Start(LibraryLocationLink.Text)
        Else
            MsgBox("Error. Cannot open " & LibraryLocationLink.Text & vbCrLf & vbCrLf & "Folder may have been deleted, renamed or moved", MsgBoxStyle.Exclamation, My.Application.Info.Title)
        End If
    End Sub
End Class
Option Strict Off
Class clsSudokuSolver
#Region "threading parameter"
    '---threading events---
    Public Event UI(ByVal Count As Integer, ByRef intSolutions As Integer, ByVal blnFinished As Boolean, ByVal strSolution As String, ByVal strDifficulty As String)
#End Region
#Region "Backtracking parameters"
    '---backtracking data---
    Private _vsSolution(0) As Integer
    Private _vsPeers(0) As String
    Private _vsCandidateCount(0) As Integer
    Private _vsCandidateAvailableBits(0) As Integer
    Private _vsCandidatePtr(0) As Integer
    Private _vsLastGuess(0) As Integer
    Private _vsStoreCandidateBits(0) As Integer
    Private _vsRemovePeers(0) As List(Of Integer) 'an arraylist of generic lists (nested lists)
    Private _vsUnsolvedCells() As List(Of Integer)
    Private _r As Integer
    Private _r2 As Integer
    Private _u As Integer = -1
    Private strFormatClues As String = ""
#End Region
#Region "Game parameter"
    '---solve up to---
    Public strSolveUpToMethod As String
    '---game input---
    Public strGrid As String
    Public strCandidates As String
    '---undo data---
    Private _vsSteps As Integer 'counter
    '---counter---
    Public vsSolvers As String = ""
    '---counter---
    Public vsTried As Integer
    '---samurai or classic---
    Public blnClassic As Boolean
    '---holds solution---
    Public strGameSolution As String
    '---holds solution if already known---
    Public strInputGameSolution As String
    '---holds no of solutions---
    Public intCountSolutions As Integer
    Public Solutions(0) As String
    '---quit after [n] solutions found---
    Public intQuit As Integer = 2
    '---hint details---
    Public blnHint As Boolean = False
    Public strHintDetails = ""
    Public intHintCount = 0
    Private strHintMsg = ""
    '--variables for explaining steps---
    Public blnStep As Boolean = False
    '---returns methods used to solve each puzzle---
    Public solveMethods() As String
    Public solveCountMethods() As Integer
    '---returns difficulty---
    Public strDifficulty As String = ""
    Public intDifficulty As Integer = 0
#End Region
#Region "Solving methods"
#Region "Naked singles"
    Public Function _ns() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'naked singles
        If _vsNakedSingles() Then
            _ns = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
            _u -= 1
        End If
    End Function
    Private Function _vsNakedSingles() As Boolean
        'naked singles
        Dim strHint As String
        Dim ptr As Integer
        Dim i As Integer
        Dim j As Integer
        Dim arrayPeers(0) As String
        Dim intValue As Integer
        Dim intBit As Integer
        Dim strFind As String = ""
        For i = 0 To _vsUnsolvedCells(0).Count - 1
            ptr = _vsUnsolvedCells(0).Item(i)
            If _vsCandidateCount(ptr) = 1 Then
                If _vsSolution(ptr - 1) = 0 Then
                    For j = 1 To 9
                        intBit = intGetBit(j)
                        If _vsCandidateAvailableBits(ptr) And intBit Then
                            intValue = j
                            Exit For
                        End If
                    Next
                    _vsSolution(ptr - 1) = intValue
                    _vsCandidateCount(ptr) = -1
                    _vsUnsolvedCells(0).Remove(ptr)
                    Select Case blnClassic
                        Case True
                            arrayPeers = arrPeers(ptr)
                        Case False
                            arrayPeers = ArrSamuraiPeers(ptr)
                    End Select
                    intBit = intGetBit(intValue)
                    'remove value from peers
                    For j = 0 To UBound(arrayPeers)
                        If _vsSolution(arrayPeers(j) - 1) = 0 AndAlso (_vsCandidateAvailableBits(arrayPeers(j)) And intBit) Then
                            _vsCandidateAvailableBits(arrayPeers(j)) -= intBit
                            _vsCandidateCount(arrayPeers(j)) -= 1
                        End If
                    Next
                    _vsNakedSingles = True

                    If blnStep Then
                        Dim hm(0) As Integer
                        hm(0) = ptr
                        HighlightMove(arrCells:=hm, strStepMsg:="Naked single found. There is only one possible digit (" & intValue & ") that can be placed in this cell.", hc1:=My.Settings._cColour, hBit1:=intBit)
                        frmGame.nakedSingle_doubleClick(intCell:=ptr, intValue:=intValue, strSolvedBy:="P")
                    End If

                    If strSolveUpToMethod <> "" Then frmGame.nakedSingle_doubleClick(intCell:=ptr, intValue:=intValue, strSolvedBy:="P")
                    '---hint code---
                    If blnHint Then
                        strHint = "NS_" & ptr & "_" & intValue
                        If strHint <> strHintDetails Then
                            strHintDetails = strHint
                            intHintCount = 0
                        Else
                            intHintCount += 1
                        End If
                        Select Case blnClassic
                            Case True
                                strFind = "R" & intFindRow(ptr) & "C" & intFindCol(ptr)
                            Case False
                        End Select
                        Select Case intHintCount
                            Case 0
                                strHintMsg = "Look for a naked single"
                            Case 1
                                strHintMsg = "Look for a naked single of digit " & intValue
                            Case Else
                                Select Case blnClassic
                                    Case True
                                        strHintMsg = "Look for a naked single in " & strFind
                                    Case False
                                        strHintMsg = "Look for a naked single of digit " & intValue
                                End Select
                        End Select
                        Exit Function
                    End If
                    '---end hint code---
                    Exit Function
                End If
            End If
        Next
        'end naked singles
    End Function
#End Region
#Region "Hidden singles"
    Public Function _hs() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'hidden singles
        Dim i As Integer
        For i = 1 To 3
            If _vsHiddenSingles(i) Then
                _hs = True
                If blnHint Then
                    ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                    Exit Function
                End If
                If blnStep Then Exit Function
                _u -= 1
            End If
        Next
    End Function
    Private Function _vsHiddenSingles(Optional ByVal intType As Integer = 0) As Boolean
        'hidden singles
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim arrayCells(0) As String
        Dim rcg As Integer
        Dim v As Integer
        Dim i As Integer
        Dim g As Integer
        Dim j As Integer
        Dim cCount As Integer
        Dim ptr As Integer
        Dim aPtr As Integer
        Dim intCell As Integer
        Dim intValue As Integer
        Dim intBit As Integer
        Dim arrayPeers(0) As String
        Dim intMaxGrid As Integer = 5
        Dim strHint As String = ""
        Dim strFind As String = ""
        Dim strLongFind As String = ""
        If blnClassic Then intMaxGrid = 1
        For g = 1 To intMaxGrid
            For rcg = 1 To 9
                Select Case intType
                    Case 1
                        arrayCells = arrRow(rcg)
                        strFind = "R" & rcg
                        strLongFind = "row " & rcg
                    Case 2
                        arrayCells = arrCol(rcg)
                        strFind = "C" & rcg
                        strLongFind = "column " & rcg
                    Case 3
                        arrayCells = arrGrid(rcg)
                        strFind = "G" & rcg
                        strLongFind = "grid " & rcg
                End Select
                'check row, col or grid has one place for candidate or value has been placed
                For v = 1 To 9
                    cCount = 0
                    intBit = intGetBit(v)
                    For i = 0 To 8
                        Select Case blnClassic
                            Case True
                                aPtr = arrayCells(i)
                            Case False
                                aPtr = intSamuraiOffset(arrayCells(i), g)
                        End Select
                        If cCount < 2 AndAlso (_vsCandidateAvailableBits(aPtr) And intBit) AndAlso _vsSolution(aPtr - 1) = 0 Then
                            cCount += 1
                            ptr = aPtr
                            intCell = v
                        End If
                    Next
                    If cCount = 1 Then
                        '---set value---
                        _vsSolution(ptr - 1) = intCell
                        intValue = intCell
                        _vsCandidateCount(ptr) = -1
                        _vsUnsolvedCells(0).Remove(ptr)
                        Select Case blnClassic
                            Case True
                                arrayPeers = arrPeers(ptr)
                            Case False
                                arrayPeers = ArrSamuraiPeers(ptr)
                        End Select
                        '---remove value from peers---
                        intBit = intGetBit(intValue)
                        For j = 0 To UBound(arrayPeers)
                            If _vsSolution(arrayPeers(j) - 1) = 0 AndAlso (_vsCandidateAvailableBits(arrayPeers(j)) And intBit) Then
                                _vsCandidateAvailableBits(arrayPeers(j)) -= intBit
                                _vsCandidateCount(arrayPeers(j)) -= 1
                            End If
                        Next

                        If blnStep Then
                            Dim hm(0) As Integer
                            hm(0) = ptr
                            HighlightMove(arrCells:=hm, strStepMsg:="Hidden single found. This is the only cell in " & strLongFind & " where digit " & intValue & " can be placed.", hc1:=My.Settings._cColour, hBit1:=intBit)
                            frmGame.nakedSingle_doubleClick(intCell:=ptr, intValue:=intValue, strSolvedBy:="P")
                        End If

                        If strSolveUpToMethod <> "" Then frmGame.nakedSingle_doubleClick(intCell:=ptr, intValue:=intValue, strSolvedBy:="P")
                        '---hint code---
                        If blnHint Then
                            strHint = "HS_" & ptr & "_" & intValue
                            If strHint <> strHintDetails Then
                                strHintDetails = strHint
                                intHintCount = 0
                            Else
                                intHintCount += 1
                            End If
                            Select Case intHintCount
                                Case 0
                                    strHintMsg = "Look for a hidden single"
                                Case 1
                                    strHintMsg = "Look for a hidden single of digit " & intValue
                                Case Else
                                    strHintMsg = "Look for a hidden single of digit " & intValue & " in " & strFind
                            End Select
                            If blnSamurai Then
                                strHintMsg += " (in samurai grid & " & g & ")"
                            End If
                        End If
                        '---end hint code---
                        _vsHiddenSingles = True
                        Exit Function
                    End If
                Next
            Next
        Next
        'end hidden singles
    End Function
#End Region
#Region "Locked candidates"
    Public Function _lc() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'locked candidates
        Dim i As Integer
        For i = 1 To 2
            If _vsLC1(i) Then
                _lc = True
                If blnHint Then
                    ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                    Exit Function
                End If
                If blnStep Then Exit Function
            End If
            If _vsLC2(i) Then
                _lc = True
                If blnHint Then
                    ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                    Exit Function
                End If
                If blnStep Then Exit Function
            End If
        Next
    End Function
    Private Function _vsLC2(Optional ByVal intType As Integer = 0) As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim ignoreCells(2) As Integer
        Dim IntRCMatch As Integer
        Dim intCount As Integer = 0
        Dim i, j As Integer ' counters
        Dim g As Integer ' grid
        Dim sg As Integer ' samurai grid
        Dim v As Integer ' value
        Dim k As Integer
        Dim ptr As Integer
        Dim intBit As Integer
        Dim blnIgnore As Boolean
        Dim arrayCells() As String
        Dim arrayGrid() As String
        Dim intMaxGrid As Integer = 5
        Dim strHint As String = ""
        Dim strRC As String = ""
        Dim strSG As String = ""
        Dim strLongRC As String = ""
        Dim strLongSG As String = ""
        Dim strMoves As String = ""
        Dim r As Integer = -1
        Dim rm(0) As Integer

        If blnClassic Then intMaxGrid = 1
        For sg = 1 To intMaxGrid
            For g = 1 To 9
                ReDim arrayCells(0)
                arrayGrid = arrGrid(g)
                strSG = "G" & g
                strLongSG = "grid " & g

                For v = 1 To 9
                    intBit = intGetBit(v)
                    IntRCMatch = 0
                    For i = 1 To 9
                        Select Case blnClassic
                            Case True
                                ptr = arrayGrid(i - 1)
                            Case False
                                'samurai
                                ptr = intSamuraiOffset(arrayGrid(i - 1), sg)
                        End Select
                        If _vsSolution(ptr - 1) = 0 AndAlso (_vsCandidateAvailableBits(ptr) And intBit) Then
                            Select Case intType
                                Case 1
                                    j = intFindRow(arrayGrid(i - 1))
                                    strRC = "R"
                                    strLongRC = "row"
                                Case 2
                                    j = intFindCol(arrayGrid(i - 1))
                                    strRC = "C"
                                    strLongRC = "column"
                            End Select
                            If IntRCMatch = 0 Or IntRCMatch = j Then
                                IntRCMatch = j
                            Else
                                IntRCMatch = -1
                                Exit For
                            End If
                        End If
                    Next
                    If IntRCMatch > 0 Then
                        r = -1
                        k = -1
                        ReDim rm(0)
                        ReDim ignoreCells(2)

                        '---get row or col---
                        Select Case intType
                            Case 1
                                arrayCells = arrRow(IntRCMatch)
                            Case 2
                                arrayCells = arrCol(IntRCMatch)
                        End Select

                        strMoves = ""

                        For i = 0 To 8
                            blnIgnore = False
                            If Array.IndexOf(arrayGrid, arrayCells(i)) > -1 Then blnIgnore = True
                            If blnIgnore Then
                                k += 1
                                ignoreCells(k) = arrayCells(i)
                            Else
                                Select Case blnClassic
                                    Case True
                                        ptr = arrayCells(i)
                                    Case False
                                        ptr = intSamuraiOffset(arrayCells(i), sg)
                                End Select
                                If _vsSolution(ptr - 1) = 0 And (_vsCandidateAvailableBits(ptr) And intBit) Then
                                    _vsCandidateAvailableBits(ptr) -= intBit
                                    _vsCandidateCount(ptr) -= 1
                                    _vsLC2 = True

                                    If strSolveUpToMethod <> "" Then
                                        Dim cCtrl As SudokuCell
                                        cCtrl = getCtrl("SudokuCell" & ptr)
                                        cCtrl.sc_IntCandidate -= intBit
                                    End If

                                    strMoves += ptr & ":" & "-" & v & "|"

                                    '---hint code---
                                    If blnHint Then
                                        strHint = "LC2_" & ptr & "_" & v
                                        If strHint <> strHintDetails Then
                                            strHintDetails = strHint
                                            intHintCount = 0
                                        Else
                                            intHintCount += 1
                                        End If
                                        Select Case intHintCount
                                            Case 0
                                                strHintMsg = "Look for a locked candidate"
                                            Case 1
                                                strHintMsg = "Look for a locked candidate of digit " & v
                                            Case Else
                                                strHintMsg = "Look for a locked candidate of digit " & v & " in " & strRC & j & " and " & strSG
                                        End Select
                                        If blnSamurai Then
                                            strHintMsg += " (in samurai grid & " & sg & ")"
                                        End If
                                        Exit Function
                                    End If
                                    '---end hint code---

                                    '---single step code---
                                    If blnStep Then
                                        r += 1
                                        ReDim Preserve rm(r)
                                        rm(r) = ptr
                                    End If
                                    '---end single step code---

                                End If
                            End If
                        Next

                        If blnStep And r > -1 Then
                            If Not blnClassic Then
                                For i = 0 To UBound(ignoreCells)
                                    ignoreCells(i) = intSamuraiOffset(ignoreCells(i), sg)
                                Next
                            End If
                            HighlightMove(arrCells:=rm, arrCells2:=ignoreCells, strStepMsg:="Locked candidate found. The digit " & v & " can only be placed in one " & strLongRC & " of " & strLongSG & ". This means digit " & v & " can be removed from " & strLongRC & " " & j & " as shown.", hc1:=My.Settings._rColour, hBit1:=intBit, hc2:=My.Settings._cColour, hBit2:=intBit, blnRemove1:=True)
                            Exit Function
                        End If

                        '---saves the move into the stack---
                        If strSolveUpToMethod <> "" And strMoves <> "" Then
                            frmGame.Moves.Push(strMoves)
                        End If
                    End If
                Next
            Next
        Next
        'end locked candidates
    End Function
    Private Function _vsLC1(Optional ByVal intType As Integer = 0) As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'locked candidates type one
        Dim ignoreCells(2) As Integer
        Dim IntGridMatch As Integer
        Dim intCount As Integer = 0
        Dim intBit As Integer
        Dim i, j As Integer ' counters
        Dim g As Integer ' grid
        Dim v As Integer ' value
        Dim rc As Integer
        Dim ptr As Integer
        Dim arrayCells() As String
        Dim arrayGrid() As String
        Dim intMaxGrid As Integer = 5
        Dim strHint As String = ""
        Dim strRC As String = ""
        Dim strSG As String = ""
        Dim strLongRC As String = ""
        Dim strLongSG As String = ""
        Dim strMoves As String = ""
        Dim r As Integer = -1
        Dim rm(0) As Integer

        If blnClassic Then intMaxGrid = 1
        For g = 1 To intMaxGrid
            For rc = 1 To 9
                ReDim arrayCells(0)
                Select Case intType
                    Case 1
                        arrayCells = arrRow(rc)
                        strRC = "R"
                        strLongRC = "row " & rc
                    Case 2
                        arrayCells = arrCol(rc)
                        strRC = "C"
                        strLongRC = "column " & rc
                End Select
                For v = 1 To 9
                    intBit = intGetBit(v)
                    intCount = 0
                    IntGridMatch = 0
                    For i = 1 To 3
                        For j = 1 To 3
                            Select Case blnClassic
                                Case True
                                    ptr = arrayCells(intCount)
                                Case False
                                    ptr = intSamuraiOffset(arrayCells(intCount), g)
                            End Select
                            If _vsSolution(ptr - 1) = 0 AndAlso (_vsCandidateAvailableBits(ptr) And intBit) Then
                                If IntGridMatch = 0 Or IntGridMatch = i Then
                                    IntGridMatch = i
                                Else
                                    IntGridMatch = -1
                                End If
                            End If
                            intCount += 1
                        Next
                    Next
                    intCount = 0
                    If IntGridMatch > 0 Then
                        ReDim rm(0)
                        r = -1
                        IntGridMatch = 1 + (IntGridMatch - 1) * 3
                        For i = 0 To 2
                            ptr = arrayCells(i + IntGridMatch - 1)
                            ignoreCells(i) = ptr
                        Next
                        IntGridMatch = intFindGrid(ignoreCells(0))
                        arrayGrid = arrGrid(IntGridMatch)
                        strMoves = ""
                        For i = 0 To 8
                            If arrayGrid(i) <> ignoreCells(0) And arrayGrid(i) <> ignoreCells(1) And arrayGrid(i) <> ignoreCells(2) Then
                                Select Case blnClassic
                                    Case True
                                        ptr = arrayGrid(i)
                                    Case False
                                        ptr = intSamuraiOffset(arrayGrid(i), g)
                                End Select
                                If _vsSolution(ptr - 1) = 0 And (_vsCandidateAvailableBits(ptr) And intBit) Then
                                    strSG = IntGridMatch
                                    strLongSG = "grid " & IntGridMatch
                                    _vsCandidateAvailableBits(ptr) -= intBit
                                    _vsCandidateCount(ptr) -= 1
                                    _vsLC1 = True

                                    If strSolveUpToMethod <> "" Then
                                        Dim cCtrl As SudokuCell
                                        cCtrl = getCtrl("SudokuCell" & ptr)
                                        cCtrl.sc_IntCandidate -= intBit
                                    End If

                                    strMoves += ptr & ":" & "-" & v & "|"

                                    '---hint code---
                                    If blnHint Then
                                        strHint = "LC1_" & ptr & "_" & v
                                        If strHint <> strHintDetails Then
                                            strHintDetails = strHint
                                            intHintCount = 0
                                        Else
                                            intHintCount += 1
                                        End If
                                        Select Case intHintCount
                                            Case 0
                                                strHintMsg = "Look for a locked candidate"
                                            Case 1
                                                strHintMsg = "Look for a locked candidate of digit " & v
                                            Case Else
                                                strHintMsg = "Look for a locked candidate of digit " & v & " in the intersection of G" & strSG & " and " & strRC & rc
                                        End Select
                                        If blnSamurai Then
                                            strHintMsg += " (in samurai grid & " & g & ")"
                                        End If
                                        Exit Function
                                    End If
                                    '---end hint code---

                                    '---single step code---
                                    If blnStep Then
                                        r += 1
                                        ReDim Preserve rm(r)
                                        rm(r) = ptr
                                    End If
                                    '---end single step code---

                                End If
                            End If
                        Next

                        If blnStep And r > -1 Then
                            If Not blnClassic Then
                                For i = 0 To UBound(ignoreCells)
                                    ignoreCells(i) = intSamuraiOffset(ignoreCells(i), g)
                                Next
                            End If
                            HighlightMove(arrCells:=rm, arrCells2:=ignoreCells, strStepMsg:="Locked candidate found. In " & strLongRC & " the digit " & v & " can only be placed in " & strLongSG & ". This means digit " & v & " can be removed from " & strLongSG & " as shown.", hc1:=My.Settings._rColour, hBit1:=intBit, hc2:=My.Settings._cColour, hBit2:=intBit, blnRemove1:=True)
                            Exit Function
                        End If

                        '---saves the move into the stack---
                        If strSolveUpToMethod <> "" And strMoves <> "" Then
                            frmGame.Moves.Push(strMoves)
                        End If
                    End If
                Next
            Next
        Next
        'end locked candidates
    End Function
#End Region
#Region "Naked pairs, triples, quads"
    Public Function _np() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'naked pairs
        If _vsNakedPairsTriplesQuads(2) Then
            _np = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Public Function _nt() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'naked triples
        If _vsNakedPairsTriplesQuads(3) Then
            _nt = True
            If blnStep Then Exit Function
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Public Function _nq() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        'naked quads
        If _vsNakedPairsTriplesQuads(4) Then
            _nq = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Private Function _vsFindNaked(ByVal arrCells() As String, ByVal intGrid As Integer, ByVal n As Integer, ByRef strCandidates As String, ByVal strRCG As String) As Boolean
        '---generic routine to find naked pairs, triples and quads---
        Dim i As Integer ' counter
        Dim j As Integer ' counter
        Dim b As Integer ' counter
        Dim r As Integer ' counter
        Dim v As Integer ' counter
        Dim m As Integer ' counter
        Dim intCellCount As Integer ' counter
        Dim intEliminate As Integer ' counter
        Dim blnCellMatch As Boolean
        Dim blnEliminate As Boolean
        Dim arrRemove() As Integer
        Dim arrMatch() As Integer
        Dim bitSum As Integer
        Dim ptr As Integer
        Dim inputarray(0) As String
        Dim outputArray(0) As String
        Dim intBit As Integer
        Dim strMoves As String = ""
        Dim strMethod As String = ""
        Select Case n
            Case 2
                strMethod = "Naked pair"
            Case 3
                strMethod = "Naked triple"
            Case 4
                strMethod = "Naked quad"
        End Select
        For i = 0 To 8
            Select Case blnClassic
                Case True
                    ptr = arrCells(i)
                Case False
                    ptr = intSamuraiOffset(arrCells(i), intGrid)
            End Select
            If _vsSolution(ptr - 1) = 0 Then
                For j = 1 To 9
                    intBit = intGetBit(j)
                    If _vsCandidateAvailableBits(ptr) And intBit Then
                        If Not bitSum And intBit Then
                            bitSum += intBit
                        End If
                    End If
                Next
            End If
        Next

        j = -1
        For b = 0 To 8
            If bitSum And intGetBit(b + 1) Then
                j += 1
                ReDim Preserve inputarray(j)
                inputarray(j) = b + 1
            End If
        Next

        If j >= 1 Then
            generateCombinationArray(inputarray, outputArray, n)
            '---scan each cell---
            For b = 0 To UBound(outputArray)
                intCellCount = 0
                intEliminate = 0
                m = -1
                ReDim arrMatch(0)
                ReDim arrRemove(intEliminate)
                For i = 0 To 8
                    Select Case blnClassic
                        Case True
                            ptr = arrCells(i)
                        Case False
                            ptr = intSamuraiOffset(arrCells(i), intGrid)
                    End Select
                    If _vsSolution(ptr - 1) = 0 AndAlso _vsCandidateCount(ptr) > 1 Then
                        blnCellMatch = True
                        For j = 1 To 9
                            intBit = intGetBit(j)
                            If _vsCandidateAvailableBits(ptr) And intBit Then
                                If Not outputArray(b) And intBit Then
                                    '---candidate does not match candidate----                                  
                                    blnCellMatch = False
                                Else
                                    blnEliminate = True
                                End If
                            End If
                        Next
                        If blnCellMatch Then
                            intCellCount += 1
                            m += 1
                            ReDim Preserve arrMatch(m)
                            arrMatch(m) = ptr
                        End If
                        If Not blnCellMatch AndAlso blnEliminate Then
                            ReDim Preserve arrRemove(intEliminate)
                            arrRemove(intEliminate) = ptr
                            intEliminate += 1
                        End If
                        blnEliminate = False
                    End If
                Next
                If intCellCount = n AndAlso intEliminate > 0 Then
                    strMoves = ""
                    For r = 0 To UBound(arrRemove)
                        ptr = arrRemove(r)
                        If _vsCandidateCount(ptr) >= 2 Then
                            For v = 0 To 8
                                intBit = intGetBit(v + 1)
                                If outputArray(b) And intBit Then
                                    If _vsCandidateAvailableBits(ptr) And intBit Then
                                        _vsCandidateAvailableBits(ptr) -= intBit
                                        _vsCandidateCount(ptr) -= 1
                                        strCandidates = bitmask2Str(outputArray(b))
                                        _vsFindNaked = True

                                        If strSolveUpToMethod <> "" Then
                                            Dim cCtrl As SudokuCell
                                            cCtrl = getCtrl("SudokuCell" & ptr)
                                            cCtrl.sc_IntCandidate -= intBit
                                        End If

                                        strMoves += ptr & ":" & "-" & v + 1 & "|"

                                    End If
                                End If
                            Next
                        End If
                    Next

                    If blnStep Then
                        HighlightMove(arrCells:=arrRemove, arrCells2:=arrMatch, strStepMsg:=strMethod & " found. There are " & n & " cells in " & strRCG & " where only the digits " & bitmask2Str(outputArray(b)) & " can be placed exactly. Other instances of the digits " & bitmask2Str(outputArray(b)) & " can be removed from " & strRCG & " as shown.", hc1:=My.Settings._rColour, hc2:=My.Settings._cColour, hBit1:=outputArray(b), hBit2:=outputArray(b), blnRemove1:=True)
                        Exit Function
                    End If

                    '---saves the move into the stack---
                    If strSolveUpToMethod <> "" And strMoves <> "" Then
                        frmGame.Moves.Push(strMoves)
                    End If
                End If
            Next
        End If
        bitSum = 0
    End Function
    Private Function _vsNakedPairsTriplesQuads(ByVal n As Integer) As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim t As Integer
        Dim g As Integer
        Dim rcg As Integer
        Dim arrayCells(0) As String
        Dim strRCG As String = ""
        Dim strLongRCG As String = ""
        Dim intMaxGrid As Integer = 5
        Dim strMethod As String = ""
        Dim strHint As String = ""
        Dim strCandidates As String = ""
        If blnClassic Then intMaxGrid = 1
        For t = 1 To 3
            For g = 1 To intMaxGrid
                For rcg = 1 To 9
                    ReDim arrayCells(0)
                    Select Case t
                        Case 1
                            arrayCells = arrRow(rcg)
                            strRCG = "R" & rcg
                            strLongRCG = "row " & rcg
                        Case 2
                            arrayCells = arrCol(rcg)
                            strRCG = "C" & rcg
                            strLongRCG = "column " & rcg
                        Case 3
                            arrayCells = arrGrid(rcg)
                            strRCG = "G" & rcg
                            strLongRCG = "grid " & rcg
                    End Select
                    Select Case n
                        Case 2
                            strMethod = "naked pair"
                        Case 3
                            strMethod = "naked triple"
                        Case 4
                            strMethod = "naked quad"
                    End Select
                    If _vsFindNaked(arrCells:=arrayCells, intGrid:=g, n:=n, strCandidates:=strCandidates, strRCG:=strLongRCG) Then
                        _vsNakedPairsTriplesQuads = True
                        If blnStep Then Exit Function
                        '---hint code---
                        If blnHint Then
                            strHint = strMethod & "_" & strRCG & "_" & rcg & "_" & strCandidates
                            If strHint <> strHintDetails Then
                                strHintDetails = strHint
                                intHintCount = 0
                            Else
                                intHintCount += 1
                            End If
                            Select Case intHintCount
                                Case 0
                                    strHintMsg = "Look for a " & strMethod
                                Case 1
                                    strHintMsg = "Look for a " & strMethod & " of digits " & strCandidates
                                Case Else
                                    strHintMsg = "Look for a " & strMethod & " of digits " & strCandidates & " in " & strRCG
                            End Select
                            If blnSamurai Then
                                strHintMsg += " (in samurai grid & " & g & ")"
                            End If
                            Exit Function
                        End If
                        '---end hint code---
                    End If
                Next
            Next
        Next
    End Function
#End Region
#Region "Hidden pairs, triples, quads"
    Public Function _hp() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        '---hidden pairs---
        If _vsHiddenPairsTriplesQuads(2) Then
            _hp = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Public Function _ht() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        '---hidden triples---
        If _vsHiddenPairsTriplesQuads(3) Then
            _ht = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Public Function _hq() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        '---hidden quads---
        If _vsHiddenPairsTriplesQuads(4) Then
            _hq = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Private Function _vsHiddenPairsTriplesQuads(ByVal n As Integer) As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim t As Integer
        Dim g As Integer
        Dim rcg As Integer
        Dim arrayCells(0) As String
        Dim strRCG As String = ""
        Dim strLongRCG As String = ""
        Dim intMaxGrid As Integer = 5
        Dim strMethod As String = ""
        Dim strHint As String = ""
        Dim strCandidates As String = ""
        If blnClassic Then intMaxGrid = 1
        For t = 1 To 3
            For g = 1 To intMaxGrid
                For rcg = 1 To 9
                    ReDim arrayCells(0)
                    Select Case t
                        Case 1
                            arrayCells = arrRow(rcg)
                            strRCG = "R" & rcg
                            strLongRCG = "row " & rcg
                        Case 2
                            arrayCells = arrCol(rcg)
                            strRCG = "C" & rcg
                            strLongRCG = "column " & rcg
                        Case 3
                            arrayCells = arrGrid(rcg)
                            strRCG = "G" & rcg
                            strLongRCG = "grid " & rcg
                    End Select
                    Select Case n
                        Case 2
                            strMethod = "hidden pair"
                        Case 3
                            strMethod = "hidden triple"
                        Case 4
                            strMethod = "hidden quad"
                    End Select

                    If _vsFindHidden(arrCells:=arrayCells, intGrid:=g, n:=n, strCandidates:=strCandidates, strRCG:=strLongRCG) Then
                        _vsHiddenPairsTriplesQuads = True
                        If blnStep Then Exit Function

                        '---hint code---
                        If blnHint Then
                            strHint = strMethod & "_" & strRCG & "_" & rcg & "_" & strCandidates
                            If strHint <> strHintDetails Then
                                strHintDetails = strHint
                                intHintCount = 0
                            Else
                                intHintCount += 1
                            End If
                            Select Case intHintCount
                                Case 0
                                    strHintMsg = "Look for a " & strMethod
                                Case 1
                                    strHintMsg = "Look for a " & strMethod & " of digits " & strCandidates
                                Case Else
                                    strHintMsg = "Look for a " & strMethod & " of digits " & strCandidates & " in " & strRCG
                            End Select
                            If blnSamurai Then
                                strHintMsg += " (in samurai grid & " & g & ")"
                            End If
                            Exit Function
                        End If
                        '---end hint code---
                    End If
                Next
            Next
        Next
    End Function
    Private Function _vsFindHidden(ByVal arrCells() As String, ByVal intGrid As Integer, ByVal n As Integer, ByRef strCandidates As String, ByVal strRCG As String) As Boolean
        Dim i As Integer '---counter---
        Dim j As Integer '---counter---
        Dim k As Integer '---counter---
        Dim inputarray(0) As String
        Dim outputArray(0) As String
        Dim posArray(9) As String
        Dim valueArray(0) As String
        Dim replaceArray(0) As String
        Dim rcgArr(0) As String
        Dim hArr(0) As Integer
        Dim intbit As Integer
        Dim rcgBit As Integer
        Dim ptr As Integer
        Dim blnOverlaps As Boolean
        Dim strMoves As String = ""
        Dim strMethod As String = ""

        Select Case n
            Case 2
                strMethod = "Hidden pair"
            Case 3
                strMethod = "Hidden triple"
            Case 4
                strMethod = "Hidden quad"
        End Select

        '---get candidate counts and positions in input array---
        _countCandidatesinRCG(arrCells, intGrid, n, posArray, valueArray)

        If valueArray(0) Is Nothing Then Exit Function
        If UBound(valueArray) + 1 < n Then Exit Function

        generateCombinationArray(valueArray, outputArray, n)

        For i = 0 To UBound(outputArray)
            '---check each combination---
            valueArray = Split(bitmask2Str(outputArray(i), arrDivider), arrDivider)
            rcgBit = 0
            For j = 0 To UBound(valueArray)
                rcgArr = Split(posArray(valueArray(j)), arrDivider)
                For k = 0 To UBound(rcgArr)
                    intbit = intGetBit(CInt(rcgArr(k)))
                    If intbit And rcgBit Then
                    Else
                        rcgBit += intbit
                    End If
                Next
            Next
            If rcgBit > 0 AndAlso intCountBits(rcgBit) = n Then
                rcgArr = Split(bitmask2Str(rcgBit, arrDivider), arrDivider)
                For j = 0 To UBound(rcgArr)
                    ptr = rcgArr(j) - 1
                    Select Case blnClassic
                        Case True
                            ptr = arrCells(ptr)
                        Case False
                            ptr = intSamuraiOffset(arrCells(ptr), intGrid)
                    End Select
                    rcgArr(j) = ptr
                    If Not _overlapBits(outputArray(i), _vsCandidateAvailableBits(ptr)) Then
                        blnOverlaps = False
                        Exit For
                    Else
                        blnOverlaps = True
                    End If
                Next
                If blnOverlaps Then
                    '---remove---
                    _vsFindHidden = True
                    strCandidates = bitmask2Str(outputArray(i))
                    strMoves = ""
                    ReDim hArr(UBound(rcgArr))
                    For j = 0 To UBound(rcgArr)
                        ptr = rcgArr(j)
                        hArr(j) = ptr
                        replaceArray = Split(bitmask2Str(_vsCandidateAvailableBits(ptr), arrDivider), arrDivider)
                        For k = 0 To UBound(replaceArray)
                            If Array.IndexOf(valueArray, replaceArray(k)) = -1 Then
                                intbit = intGetBit(replaceArray(k))
                                _vsCandidateAvailableBits(ptr) -= intbit
                                If strSolveUpToMethod <> "" Then
                                    Dim cCtrl As SudokuCell
                                    cCtrl = getCtrl("SudokuCell" & ptr)
                                    cCtrl.sc_IntCandidate -= intbit
                                End If
                                strMoves += ptr & ":" & "-" & replaceArray(k) & "|"
                            End If
                        Next
                    Next

                    If blnStep Then
                        HighlightMove(arrCells:=hArr, arrCells2:=hArr, strStepMsg:=strMethod & " found. There are " & n & " cells where only the digits " & bitmask2Str(outputArray(i)) & " can be placed in " & strRCG & ". All other digits in these cells can be removed as shown.", hc1:=My.Settings._rColour, hc2:=My.Settings._cColour, hBit1:=511 - outputArray(i), hBit2:=outputArray(i), blnRemove1:=True)
                        Exit Function
                    End If

                    '---saves the move into the stack---
                    If strSolveUpToMethod <> "" And strMoves <> "" Then
                        frmGame.Moves.Push(strMoves)
                    End If
                    'end remove
                End If
            End If
            '---end check each combination---
        Next
    End Function
    Private Function _overlapBits(ByVal intTestBits As Integer, ByVal intLongBit As Integer) As Boolean
        Dim intOppositeBits As Integer = 511 - intTestBits
        Dim intBit As Integer
        Dim blnMatch As Boolean
        Dim blnNoMatch As Boolean
        Dim i As Integer
        For i = 1 To 9
            intBit = intGetBit(i)
            If (intLongBit And intBit) AndAlso (intBit And intTestBits) Then
                blnMatch = True
            End If
            If (intLongBit And intBit) AndAlso (intBit And intOppositeBits) Then
                blnNoMatch = True
            End If
            If blnMatch And blnNoMatch Then
                Return True
            End If
        Next
    End Function
#End Region
#Region "Naked subsets - xwing, swordfish, jellyfish"
    Public Function _xw() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        '---xwing----
        If _vsXWingSwordfishJellyfish(2) Then
            _xw = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Public Function _sf() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        '---swordfish----
        If _vsXWingSwordfishJellyfish(3) Then
            _sf = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Public Function _jf() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        '---jellyfish----
        If _vsXWingSwordfishJellyfish(4) Then
            _jf = True
            If blnHint Then
                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                Exit Function
            End If
            If blnStep Then Exit Function
        End If
    End Function
    Private Function _vsXWingSwordfishJellyfish(ByVal n As Integer) As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim g As Integer
        Dim rc As Integer
        Dim strRC As String = ""
        Dim strLongRC As String = ""
        Dim intMaxGrid As Integer = 5
        Dim strMethod As String = ""
        Dim strHint As String = ""
        Dim strCandidates As String = ""
        If blnClassic Then intMaxGrid = 1
        For g = 1 To intMaxGrid
            For rc = 1 To 2
                Select Case rc
                    Case 1
                        strRC = "R" & rc
                        strLongRC = "row " & rc
                    Case 2
                        strRC = "C" & rc
                        strLongRC = "column " & rc
                End Select
                Select Case n
                    Case 2
                        strMethod = "xwing"
                    Case 3
                        strMethod = "swordfish"
                    Case 4
                        strMethod = "jellyfish"
                End Select
                If _vsFindNakedSubsets(g, n, rc) Then
                    _vsXWingSwordfishJellyfish = True
                    Exit Function
                End If
            Next
        Next
    End Function
    Private Function _vsFindNakedSubsets(ByVal intGrid As Integer, ByVal n As Integer, ByVal t As Integer) As Boolean
        Dim strMethod As String = ""
        Select Case n
            Case 2
                strMethod = "XWing"
            Case 3
                strMethod = "Swordfish"
            Case 4
                strMethod = "Jellyfish"
        End Select

        Dim rc As Integer
        Dim v As Integer
        Dim i As Integer
        Dim m As Integer
        Dim j As Integer
        Dim k As Integer
        Dim h As Integer
        Dim r As Integer
        Dim intMatchBits As Integer
        Dim intTestBit As Integer
        Dim intBit As Integer
        Dim arrCells(0) As String
        Dim posArray(0) As String
        Dim valueArray(0) As String
        Dim d(9, 9) As String
        Dim dCount(9) As Integer
        Dim hArr(0) As Integer
        Dim rArr(0) As Integer
        Dim arrRC(0) As String
        Dim strMatchPos As String = ""
        Dim arrF() As String
        Dim arrRCMatch(0) As String
        Dim arrComboRC(0) As String
        Dim arrTestRC(0) As String
        Dim arrTestMatchRC(0) As String
        Dim strSearch As String = ""
        Dim strFound As String = ""
        Dim strMoves As String = ""
        Dim ptr As Integer
        Dim strHint As String = ""

        For rc = 1 To 9
            Select Case t
                Case 1
                    arrCells = arrRow(rc)
                    strSearch = "row"
                    strFound = "column"
                Case 2
                    arrCells = arrCol(rc)
                    strSearch = "column"
                    strFound = "row"
            End Select
            ReDim posArray(9)
            ReDim valueArray(0)

            _countCandidatesinRCG(arrCells, intGrid, n, posArray, valueArray)
            For v = 0 To UBound(valueArray)
                d(valueArray(v), rc) = posArray(valueArray(v))
                dCount(valueArray(v)) += 1
            Next
        Next

        For v = 1 To 9
            ReDim arrRC(0)
            If dCount(v) >= n Then
                m = 0
                For i = 1 To 9
                    If d(v, i) <> "" Then
                        ReDim Preserve arrRC(m)
                        arrRC(m) = i
                        m += 1
                    End If
                Next
            End If
            If arrRC.Length >= n Then
                If arrRC.Length >= n Then
                    generateCombinationArray(arrRC, arrComboRC, n)
                End If
                For i = 0 To UBound(arrComboRC)
                    arrTestRC = Split(bitmask2Str(arrComboRC(i), arrDivider), arrDivider)
                    intMatchBits = 0
                    m = 0
                    For j = 0 To UBound(arrTestRC)
                        arrTestMatchRC = Split(d(v, arrTestRC(j)), arrDivider)
                        For k = 0 To UBound(arrTestMatchRC)
                            intTestBit = intGetBit(arrTestMatchRC(k))
                            If Not intMatchBits And intTestBit Then
                                intMatchBits += intTestBit
                                m += 1
                                If m > n Then Exit For
                            End If
                        Next
                    Next
                    If m = n Then
                        strMatchPos = bitmask2Str(arrComboRC(i))
                        arrF = Split(bitmask2Str(intMatchBits, arrDivider), arrDivider)
                        strMoves = ""
                        r = -1
                        h = -1
                        ReDim hArr(0)
                        ReDim rArr(0)
                        For j = 0 To UBound(arrF)
                            Select Case t
                                Case 1
                                    arrCells = arrCol(arrF(j))
                                Case 2
                                    arrCells = arrRow(arrF(j))
                            End Select
                            For k = 0 To UBound(arrCells)
                                Select Case blnClassic
                                    Case True
                                        ptr = arrCells(k)
                                    Case False
                                        ptr = intSamuraiOffset(arrCells(k), intGrid)
                                End Select
                                If InStr(strMatchPos, (k + 1)) = 0 Then
                                    '---check for digit---
                                    intBit = intGetBit(v)
                                    If _vsSolution(ptr - 1) = 0 AndAlso (_vsCandidateAvailableBits(ptr) And intBit) Then
                                        '---remove code---
                                        _vsFindNakedSubsets = True
                                        _vsCandidateAvailableBits(ptr) -= intBit
                                        r += 1
                                        ReDim Preserve rArr(r)
                                        rArr(r) = ptr
                                        If strSolveUpToMethod <> "" Then
                                            Dim cCtrl As SudokuCell
                                            cCtrl = getCtrl("SudokuCell" & ptr)
                                            cCtrl.sc_IntCandidate -= intBit
                                        End If
                                        strMoves += ptr & ":" & "-" & v & "|"
                                        '---end remove code---
                                    End If
                                Else
                                    h += 1
                                    ReDim Preserve hArr(h)
                                    hArr(h) = ptr
                                End If
                            Next
                        Next
                        If _vsFindNakedSubsets = True Then

                            '---hint code---
                            If blnHint Then
                                strHint = strMethod & "_" & v & "_" & bitmask2Str(arrComboRC(i), arrDivider) & "_" & bitmask2Str(intMatchBits, arrDivider)
                                If strHint <> strHintDetails Then
                                    strHintDetails = strHint
                                    intHintCount = 0
                                Else
                                    intHintCount += 1
                                End If
                                Select Case intHintCount
                                    Case 0
                                        strHintMsg = "Look for a " & LCase(strMethod) & " pattern"
                                    Case 1
                                        strHintMsg = "Look for a " & LCase(strMethod) & " pattern featuring digit " & v
                                    Case 2
                                        strHintMsg = "Look for a " & LCase(strMethod) & " pattern featuring digit " & v & " in " & strSearch & "s " & bitmask2Str(arrComboRC(i), arrDivider)
                                    Case Else
                                        strHintMsg = "Look for a " & LCase(strMethod) & " pattern featuring digit " & v & " in " & strSearch & "s " & bitmask2Str(arrComboRC(i), arrDivider) & " and " & strFound & "s " & bitmask2Str(intMatchBits, arrDivider)
                                End Select
                                If blnSamurai Then
                                    strHintMsg += " (in samurai grid & " & intGrid & ")"
                                End If
                                Exit Function
                            End If
                            '---end hint code---
                            If blnStep Then
                                HighlightMove(arrCells:=rArr, arrCells2:=hArr, strStepMsg:=strMethod & " pattern found. There are " & n & " " & strSearch & "s (" & bitmask2Str(arrComboRC(i), arrDivider) & ") where digit " & v & " can be found in exactly " & n & " " & strFound & "s (" & bitmask2Str(intMatchBits, arrDivider) & "). The remaining candidates where digit " & v & " is found in these " & strFound & "s can be removed.", hc1:=My.Settings._rColour, hc2:=My.Settings._cColour, hBit1:=intBit, hBit2:=intBit, blnRemove1:=True)
                                Exit Function
                            End If
                            '---saves the move into the stack---
                            If strSolveUpToMethod <> "" And strMoves <> "" Then
                                frmGame.Moves.Push(strMoves)
                            End If
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next


    End Function
#End Region
#Region "XY wing"
    Function _xyw() As Boolean
        '---find XY wing pattern---
        Dim a As Integer
        Dim b As Integer
        Dim i As Integer
        Dim j As Integer
        Dim g As Integer
        Dim p As Integer
        Dim ptr As Integer
        Dim ptr2 As Integer
        Dim maxgrid As Integer = 1
        If blnSamurai Then maxgrid = 5
        Dim arrSplit(0) As String
        Dim arrayPeers(0) As String
        Dim arrayPeers2(0) As String
        Dim arrPos(2) As Integer
        Dim arrXPos(0) As Integer
        Dim arrYPos(0) As Integer
        Dim arrPosTemp(0) As String
        Dim arrPosTest(0) As String
        Dim arrRemove(0) As Integer
        Dim strXY As String = ""
        Dim strXY2 As String = ""
        Dim xyBit As Integer = 0
        Dim xyzBit As Integer = 0
        Dim xBit As Integer = 0
        Dim yBit As Integer = 0
        Dim zBit As Integer = 0
        Dim z1 As Integer
        Dim z2 As Integer
        Dim zBitArr1(0) As Integer
        Dim zBitArr2(0) As Integer
        Dim strMoves As String = ""
        Dim strHint As String = ""
        For g = 1 To maxgrid
            For i = 1 To 81
                ptr = i
                If Not blnClassic Then ptr = intSamuraiOffset(ptr, g)
                If _vsSolution(ptr - 1) = 0 AndAlso _vsCandidateCount(ptr) = 2 Then
                    strXY = bitmask2Str(_vsCandidateAvailableBits(ptr), arrDivider)
                    strXY2 = bitmask2Str(_vsCandidateAvailableBits(ptr), "")
                    xyBit = _vsCandidateAvailableBits(ptr)
                    arrSplit = Split(strXY, arrDivider)
                    If arrSplit.Length = 2 Then
                        xBit = intGetBit(arrSplit(0))
                        yBit = intGetBit(arrSplit(1))
                        arrayPeers = arrPeers(i)
                        arrPos(0) = i
                        z1 = -1
                        z2 = -1
                        a = -1
                        ReDim zBitArr1(0)
                        ReDim zBitArr2(0)
                        ReDim arrXPos(0)
                        ReDim arrYPos(0)
                        ReDim arrPosTemp(0)
                        For p = 0 To UBound(arrayPeers)
                            ptr = arrayPeers(p)
                            If Not blnClassic Then ptr = intSamuraiOffset(ptr, g)
                            If _vsSolution(ptr - 1) = 0 Then
                                If _vsCandidateCount(ptr) = 2 AndAlso (_vsCandidateAvailableBits(ptr) And xBit) Then
                                    z1 += 1
                                    ReDim Preserve zBitArr1(z1)
                                    ReDim Preserve arrXPos(z1)
                                    zBitArr1(z1) = _vsCandidateAvailableBits(ptr) - xBit
                                    arrXPos(z1) = arrayPeers(p)
                                    If z1 > -1 And z2 > -1 Then
                                        '---test to see if we have any matching z values---
                                        If Array.IndexOf(zBitArr2, zBitArr1(z1)) > -1 Then
                                            arrPos(1) = arrXPos(z1)
                                            arrPos(2) = arrYPos(Array.IndexOf(zBitArr2, zBitArr1(z1)))
                                            a += 1
                                            ReDim Preserve arrPosTemp(a)
                                            arrPosTemp(a) = arrPos(0) & arrDivider & arrPos(1) & arrDivider & arrPos(2)
                                        End If
                                    End If
                                End If
                            End If
                            If _vsCandidateCount(ptr) = 2 AndAlso (_vsCandidateAvailableBits(ptr) And yBit) Then
                                z2 += 1
                                ReDim Preserve zBitArr2(z2)
                                ReDim Preserve arrYPos(z2)
                                zBitArr2(z2) = _vsCandidateAvailableBits(ptr) - yBit
                                arrYPos(z2) = arrayPeers(p)
                                If z1 > -1 And z2 > -1 Then
                                    '---test to see if we have any matching z values---
                                    If Array.IndexOf(zBitArr1, zBitArr2(z2)) > -1 Then
                                        arrPos(1) = arrYPos(z2)
                                        arrPos(2) = arrXPos(Array.IndexOf(zBitArr1, zBitArr2(z2)))
                                        a += 1
                                        ReDim Preserve arrPosTemp(a)
                                        arrPosTemp(a) = arrPos(0) & arrDivider & arrPos(1) & arrDivider & arrPos(2)
                                    End If
                                End If
                            End If
                        Next
                        If a > -1 Then
                            For a = 0 To UBound(arrPosTemp)
                                arrPosTest = Split(arrPosTemp(a), arrDivider)
                                For j = 0 To UBound(arrPosTest)
                                    arrPos(j) = arrPosTest(j)
                                Next
                                If (intFindRow(arrPos(1)) <> intFindRow(arrPos(2))) And (intFindCol(arrPos(1)) <> intFindCol(arrPos(2))) And (intFindGrid(arrPos(1)) <> intFindGrid(arrPos(2))) Then
                                    arrayPeers = arrPeers(arrPos(1))
                                    arrayPeers2 = arrPeers(arrPos(2))
                                    strXY = bitmask2Str(xyBit, arrDivider)
                                    strXY2 = bitmask2Str(xyBit)
                                    arrSplit = Split(strXY, arrDivider)
                                    xBit = intGetBit(arrSplit(0))
                                    yBit = intGetBit(arrSplit(1))
                                    ptr2 = arrPos(1)
                                    If Not blnClassic Then ptr2 = intSamuraiOffset(ptr2, g)
                                    If _vsCandidateAvailableBits(ptr2) And xBit Then
                                        zBit = _vsCandidateAvailableBits(ptr2) - xBit
                                    End If
                                    If _vsCandidateAvailableBits(ptr2) And yBit Then
                                        zBit = _vsCandidateAvailableBits(ptr2) - yBit
                                    End If
                                    xyzBit = xBit + yBit + zBit
                                    b = -1
                                    ReDim arrRemove(0)
                                    For j = 0 To UBound(arrayPeers)
                                        ptr = arrayPeers(j)
                                        If Not blnClassic Then ptr = intSamuraiOffset(ptr, g)
                                        If arrayPeers(j) <> arrPos(0) AndAlso Array.IndexOf(arrayPeers2, CStr(arrayPeers(j))) > -1 Then
                                            If _vsSolution(ptr - 1) = 0 Then
                                                If _vsCandidateCount(ptr) >= 2 AndAlso (_vsCandidateAvailableBits(ptr) And zBit) Then
                                                    b += 1
                                                    ReDim Preserve arrRemove(b)
                                                    arrRemove(b) = ptr
                                                    '---remove code---
                                                    _vsCandidateAvailableBits(ptr) -= zBit
                                                    If strSolveUpToMethod <> "" Then
                                                        Dim cCtrl As SudokuCell
                                                        cCtrl = getCtrl("SudokuCell" & ptr)
                                                        cCtrl.sc_IntCandidate -= zBit
                                                    End If
                                                    strMoves += ptr & ":" & "-" & intReverseBit(zBit) & "|"
                                                    '---end remove code---
                                                End If
                                            End If
                                        End If
                                    Next
                                    If b > -1 Then
                                        '---xy wing found---
                                        _xyw = True

                                        If Not blnClassic Then
                                            For j = 0 To UBound(arrPos)
                                                arrPos(j) = intSamuraiOffset(arrPos(j), g)
                                            Next
                                        End If

                                        '---hint code---
                                        If blnHint Then
                                            strHint = "xy_" & zBit & "_" & strMoves
                                            If strHint <> strHintDetails Then
                                                strHintDetails = strHint
                                                intHintCount = 0
                                            Else
                                                intHintCount += 1
                                            End If
                                            Select Case intHintCount
                                                Case 0
                                                    strHintMsg = "Look for an XY wing pattern"
                                                Case Else
                                                    strHintMsg = "Look for an XY wing pattern featuring digits " & strXY
                                            End Select
                                            If blnSamurai Then
                                                strHintMsg += " (in samurai grid & " & g & ")"
                                            End If
                                            ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                                            Exit Function
                                        End If
                                        '---end hint code---
                                        If blnStep Then
                                            HighlightMove(strStepMsg:="XY wing pattern found. Regardless of whether digit " & intReverseBit(xBit) & " or " & intReverseBit(yBit) & " is placed in the cell containing digits " & strXY2 & ", digit " & intReverseBit(zBit) & " is forced to appear in one of the two other cells highlighted. The digit " & intReverseBit(zBit) & " can be removed from any cells which can see both these two cells.", arrCells:=arrPos, hc1:=My.Settings._pColour, hBit1:=xyBit, arrCells2:=arrPos, hc2:=My.Settings._nColour, hBit2:=zBit, arrCells3:=arrRemove, hc3:=My.Settings._rColour, hBit3:=zBit, blnRemove3:=True)
                                            Exit Function
                                        End If
                                        '---saves the move into the stack---
                                        If strSolveUpToMethod <> "" And strMoves <> "" Then
                                            frmGame.Moves.Push(strMoves)
                                        End If
                                        '---end xy wing found---
                                        Exit Function
                                    Else
                                        _xyw = False
                                    End If
                                End If
                            Next
                            b = -1
                        End If
                    End If
                End If
            Next
        Next
    End Function
#End Region
#Region "XYZ wing"
    Function _xyzw() As Boolean
        Dim g As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim p As Integer
        Dim r As Integer
        Dim t1 As Integer
        Dim t2 As Integer
        Dim ptr As Integer
        Dim aCount As Integer
        Dim bCount As Integer
        Dim maxGrid As Integer = 1
        If blnSamurai Then maxGrid = 5
        Dim strXYZ As String
        Dim xyzBit As Integer
        Dim aBit As Integer
        Dim bBit As Integer
        Dim zBit As Integer
        Dim arrSplit() As String
        Dim arrayPeers() As String
        Dim arrPos(2) As Integer
        Dim arrRemove(0) As Integer
        Dim strPos(2) As String
        Dim tPos1() As String
        Dim tPos2() As String
        Dim strMoves As String = ""
        Dim strHint As String
        For g = 1 To maxGrid
            For i = 1 To 81
                ptr = i
                If Not blnClassic Then ptr = intSamuraiOffset(i, g)
                If _vsSolution(ptr - 1) = 0 AndAlso _vsCandidateCount(ptr) = 3 Then
                    strXYZ = bitmask2Str(_vsCandidateAvailableBits(ptr), arrDivider)
                    xyzBit = _vsCandidateAvailableBits(ptr)
                    arrSplit = Split(strXYZ, arrDivider)
                    strXYZ = bitmask2Str(_vsCandidateAvailableBits(ptr))
                    If arrSplit.Length = 3 Then
                        For j = 0 To 2
                            Select Case j
                                Case 0
                                    aBit = intGetBit(arrSplit(0)) + intGetBit(arrSplit(1))
                                    bBit = intGetBit(arrSplit(0)) + intGetBit(arrSplit(2))
                                    zBit = intGetBit(arrSplit(0))
                                Case 1
                                    aBit = intGetBit(arrSplit(0)) + intGetBit(arrSplit(1))
                                    bBit = intGetBit(arrSplit(1)) + intGetBit(arrSplit(2))
                                    zBit = intGetBit(arrSplit(1))
                                Case 2
                                    aBit = intGetBit(arrSplit(0)) + intGetBit(arrSplit(2))
                                    bBit = intGetBit(arrSplit(1)) + intGetBit(arrSplit(2))
                                    zBit = intGetBit(arrSplit(2))
                            End Select
                            ReDim arrPos(2)
                            ReDim strPos(2)
                            arrayPeers = arrPeers(i)
                            aCount = 0
                            bCount = 0
                            strPos(0) = i
                            arrPos(0) = i
                            For p = 0 To UBound(arrayPeers)
                                ptr = arrayPeers(p)
                                If Not blnClassic Then ptr = intSamuraiOffset(ptr, g)
                                If _vsSolution(ptr - 1) = 0 Then
                                    If _vsCandidateAvailableBits(ptr) = aBit AndAlso _vsCandidateCount(ptr) = 2 Then
                                        aCount += 1
                                        strPos(1) += arrayPeers(p) & arrDivider
                                    End If
                                    If _vsCandidateAvailableBits(ptr) = bBit AndAlso _vsCandidateCount(ptr) = 2 Then
                                        bCount += 1
                                        strPos(2) += arrayPeers(p) & arrDivider
                                    End If
                                End If
                            Next
                            If aCount > 0 And bCount > 0 Then
                                '---possible xyz wing found. Check cells that contain---
                                '---digit 'z' and one other candidate and is a peer of---
                                '---all three cells---
                                tPos1 = Split(strPos(1), arrDivider)
                                tPos2 = Split(strPos(2), arrDivider)
                                For t1 = 0 To UBound(tPos1) - 1
                                    For t2 = 0 To UBound(tPos2) - 1
                                        arrPos(0) = strPos(0)
                                        arrPos(1) = tPos1(t1)
                                        arrPos(2) = tPos2(t2)
                                        r = -1
                                        For k = 1 To 81
                                            ptr = k
                                            If Not blnClassic Then ptr = intSamuraiOffset(k, g)
                                            If _vsSolution(ptr - 1) = 0 AndAlso _vsCandidateCount(ptr) > 1 Then
                                                If _vsCandidateAvailableBits(ptr) And zBit Then
                                                    arrayPeers = arrPeers(k)
                                                    If Array.IndexOf(arrayPeers, CStr(arrPos(0))) > -1 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos(1))) > -1 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos(2))) > -1 Then
                                                        _xyzw = True
                                                        r += 1
                                                        ReDim Preserve arrRemove(r)
                                                        If Not blnClassic Then ptr = intSamuraiOffset(k, g)
                                                        arrRemove(r) = ptr
                                                        '---remove code---
                                                        _vsCandidateAvailableBits(ptr) -= zBit
                                                        If strSolveUpToMethod <> "" Then
                                                            Dim cCtrl As SudokuCell
                                                            cCtrl = getCtrl("SudokuCell" & ptr)
                                                            cCtrl.sc_IntCandidate -= zBit
                                                        End If
                                                        strMoves += ptr & ":" & "-" & intReverseBit(zBit) & "|"
                                                        '---end remove code---
                                                    End If
                                                End If
                                            End If
                                        Next

                                        If _xyzw Then

                                            If Not blnClassic Then
                                                For k = 0 To UBound(arrPos)
                                                    arrPos(k) = intSamuraiOffset(arrPos(k), g)
                                                Next
                                            End If

                                            '---hint code---
                                            If blnHint Then
                                                strHint = "xyz_" & intReverseBit(zBit) & "_" & strMoves
                                                If strHint <> strHintDetails Then
                                                    strHintDetails = strHint
                                                    intHintCount = 0
                                                Else
                                                    intHintCount += 1
                                                End If
                                                Select Case intHintCount
                                                    Case 0
                                                        strHintMsg = "Look for an XYZ wing pattern"
                                                    Case Else
                                                        strHintMsg = "Look for an XYZ wing pattern featuring digits " & strXYZ
                                                End Select
                                                If blnSamurai Then
                                                    strHintMsg += " (in samurai grid & " & g & ")"
                                                End If
                                                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                                                Exit Function
                                            End If
                                            '---end hint code---
                                            If blnStep Then
                                                HighlightMove(strStepMsg:="XYZ wing pattern found. The cell containing digits " & strXYZ & " can see two cells containing digits " & bitmask2Str(aBit) & " and " & bitmask2Str(bBit) & ". Therefore the solution to one of the three highlighted cells must be " & intReverseBit(zBit) & " which can be removed from any cells that can see the three highlighted cells as shown.", arrCells:=arrPos, arrCells2:=arrPos, arrCells3:=arrRemove, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hc3:=My.Settings._rColour, hBit1:=xyzBit, hBit2:=zBit, hBit3:=zBit, blnRemove1:=False, blnRemove2:=False, blnRemove3:=True)
                                                Exit Function
                                            End If
                                            '---saves the move into the stack---
                                            If strSolveUpToMethod <> "" And strMoves <> "" Then
                                                frmGame.Moves.Push(strMoves)
                                            End If
                                            Exit Function

                                        End If

                                    Next
                                Next
                            End If
                        Next
                    End If
                End If
            Next
        Next
    End Function
    'Function _xyzw() As Boolean
    '    '---find XYZ wing pattern---
    '    Dim i As Integer
    '    Dim j As Integer
    '    Dim g As Integer
    '    Dim p As Integer
    '    Dim r As Integer
    '    Dim t1 As Integer
    '    Dim t2 As Integer
    '    Dim ptr As Integer
    '    Dim xCount As Integer
    '    Dim yCount As Integer
    '    Dim xCount2 As Integer
    '    Dim yCount2 As Integer
    '    Dim xCount3 As Integer
    '    Dim yCount3 As Integer
    '    Dim maxgrid As Integer = 1
    '    If blnSamurai Then maxgrid = 5
    '    Dim arrSplit(0) As String
    '    Dim arrayPeers(0) As String
    '    Dim arrPos(2) As Integer
    '    Dim arrPos2(2) As Integer
    '    Dim arrPos3(2) As Integer
    '    Dim tPos(0) As String
    '    Dim tPos2(0) As String
    '    Dim arrRemove(0) As Integer
    '    Dim strXYZ As String = ""
    '    Dim xyzBit As Integer = 0
    '    Dim yzBit As Integer
    '    Dim xzBit As Integer
    '    Dim yzBit2 As Integer
    '    Dim xzBit2 As Integer
    '    Dim yzBit3 As Integer
    '    Dim xzBit3 As Integer
    '    Dim zBit As Integer
    '    Dim intType As Integer = 0
    '    Dim blnRemove As Boolean = False
    '    Dim strMoves As String = ""
    '    Dim strHint As String = ""
    '    For g = 1 To maxgrid
    '        For i = 1 To 81
    '            ptr = i
    '            If Not blnClassic Then ptr = intSamuraiOffset(i, g)
    '            If _vsSolution(ptr - 1) = 0 AndAlso _vsCandidateCount(ptr) = 3 Then
    '                strXYZ = bitmask2Str(_vsCandidateAvailableBits(ptr), arrDivider)
    '                xyzBit = _vsCandidateAvailableBits(ptr)
    '                arrSplit = Split(strXYZ, arrDivider)
    '                strXYZ = bitmask2Str(_vsCandidateAvailableBits(ptr))
    '                If arrSplit.Length = 3 Then
    '                    yzBit = intGetBit(arrSplit(1)) + intGetBit(arrSplit(2))
    '                    xzBit = intGetBit(arrSplit(0)) + intGetBit(arrSplit(2))
    '                    yzBit2 = intGetBit(arrSplit(0)) + intGetBit(arrSplit(1))
    '                    xzBit2 = intGetBit(arrSplit(1)) + intGetBit(arrSplit(2))
    '                    yzBit3 = intGetBit(arrSplit(0)) + intGetBit(arrSplit(1))
    '                    xzBit3 = intGetBit(arrSplit(0)) + intGetBit(arrSplit(2))
    '                    zBit = intGetBit(arrSplit(2))
    '                    arrayPeers = arrPeers(i)
    '                    xCount = 0
    '                    yCount = 0
    '                    xCount2 = 0
    '                    yCount2 = 0
    '                    xCount3 = 0
    '                    yCount3 = 0
    '                    arrPos(0) = ptr
    '                    arrPos2(0) = ptr
    '                    arrPos3(0) = ptr
    '                    For p = 0 To UBound(arrayPeers)
    '                        ptr = arrayPeers(p)
    '                        If Not blnClassic Then ptr = intSamuraiOffset(i, g)
    '                        If _vsSolution(ptr - 1) = 0 Then
    '                            If _vsCandidateAvailableBits(ptr) = yzBit AndAlso _vsCandidateCount(ptr) = 2 Then
    '                                yCount += 1
    '                                arrPos(1) += arrayPeers(p) & arrDivider
    '                            End If
    '                            If _vsCandidateAvailableBits(ptr) = xzBit AndAlso _vsCandidateCount(ptr) = 2 Then
    '                                xCount += 1
    '                                arrPos(2) += arrayPeers(p) & arrDivider
    '                            End If
    '                            If _vsCandidateAvailableBits(ptr) = yzBit2 AndAlso _vsCandidateCount(ptr) = 2 Then
    '                                yCount2 += 1
    '                                arrPos2(1) += arrayPeers(p) & arrDivider
    '                            End If
    '                            If _vsCandidateAvailableBits(ptr) = xzBit2 AndAlso _vsCandidateCount(ptr) = 2 Then
    '                                xCount2 += 1
    '                                arrPos2(2) += arrayPeers(p) & arrDivider
    '                            End If
    '                            If _vsCandidateAvailableBits(ptr) = yzBit3 AndAlso _vsCandidateCount(ptr) = 2 Then
    '                                yCount3 += 1
    '                                arrPos3(1) += arrayPeers(p) & arrDivider
    '                            End If
    '                            If _vsCandidateAvailableBits(ptr) = xzBit3 AndAlso _vsCandidateCount(ptr) = 2 Then
    '                                xCount3 += 1
    '                                arrPos3(2) += arrayPeers(p) & arrDivider
    '                            End If
    '                        End If
    '                        '---check---
    '                        If (xCount >= 1 And yCount >= 1) Or (xCount2 >= 1 And yCount2 >= 1) Or (xCount3 >= 1 And yCount3 >= 1) Then
    '                            '---possible xyz wing found. Check cells that contain---
    '                            '---digit 'z' and one other candidate and is a peer of---
    '                            '---all three cells---
    '                            If (xCount >= 1 And yCount >= 1) Then
    '                                zBit = intGetBit(arrSplit(2))
    '                                intType = 1
    '                                tPos = Split(arrPos(1), arrDivider)
    '                                tPos2 = Split(arrPos(2), arrDivider)
    '                            End If
    '                            If (xCount2 >= 1 And yCount2 >= 1) Then
    '                                zBit = intGetBit(arrSplit(1))
    '                                intType = 2
    '                                tPos = Split(arrPos2(1), arrDivider)
    '                                tPos2 = Split(arrPos2(2), arrDivider)
    '                            End If
    '                            If (xCount3 >= 1 And yCount3 >= 1) Then
    '                                zBit = intGetBit(arrSplit(0))
    '                                intType = 3
    '                                tPos = Split(arrPos3(1), arrDivider)
    '                                tPos2 = Split(arrPos3(2), arrDivider)
    '                            End If
    '                            r = -1
    '                            For t1 = 0 To UBound(tPos)
    '                                If _xyzw = True Then Exit For
    '                                For t2 = 0 To UBound(tPos2)
    '                                    If _xyzw = True Then Exit For
    '                                    ReDim arrRemove(0)
    '                                    For j = 1 To 81
    '                                        blnRemove = False
    '                                        ptr = j
    '                                        If Not blnClassic Then ptr = intSamuraiOffset(j, g)
    '                                        If _vsSolution(ptr - 1) = 0 AndAlso _vsCandidateCount(ptr) > 1 Then
    '                                            If _vsCandidateAvailableBits(ptr) And zBit Then
    '                                                arrayPeers = arrPeers(j)
    '                                                If intType = 1 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos(0))) > -1 AndAlso Array.IndexOf(arrayPeers, tPos(t1)) > -1 AndAlso Array.IndexOf(arrayPeers, tPos2(t2)) > -1 Then
    '                                                    _xyzw = True
    '                                                    blnRemove = True
    '                                                End If
    '                                                If intType = 2 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos2(0))) > -1 AndAlso Array.IndexOf(arrayPeers, tPos(t1)) > -1 AndAlso Array.IndexOf(arrayPeers, tPos2(t2)) > -1 Then 'Array.IndexOf(arrayPeers, CStr(arrPos2(1))) > -1 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos2(2))) > -1 Then
    '                                                    _xyzw = True
    '                                                    blnRemove = True
    '                                                End If
    '                                                If intType = 3 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos3(0))) > -1 AndAlso Array.IndexOf(arrayPeers, tPos(t1)) > -1 AndAlso Array.IndexOf(arrayPeers, tPos2(t2)) > -1 Then 'Array.IndexOf(arrayPeers, CStr(arrPos3(1))) > -1 AndAlso Array.IndexOf(arrayPeers, CStr(arrPos3(2))) > -1 Then
    '                                                    _xyzw = True
    '                                                    blnRemove = True
    '                                                End If
    '                                                If blnRemove Then
    '                                                    r += 1
    '                                                    ReDim Preserve arrRemove(r)
    '                                                    If Not blnClassic Then ptr = intSamuraiOffset(j, g)
    '                                                    arrRemove(r) = ptr
    '                                                    '---remove code---
    '                                                    _vsCandidateAvailableBits(ptr) -= zBit
    '                                                    If strSolveUpToMethod <> "" Then
    '                                                        Dim cCtrl As SudokuCell
    '                                                        cCtrl = getCtrl("SudokuCell" & ptr)
    '                                                        cCtrl.sc_IntCandidate -= zBit
    '                                                    End If
    '                                                    strMoves += ptr & ":" & "-" & intReverseBit(zBit) & "|"
    '                                                    '---end remove code---
    '                                                End If
    '                                            End If
    '                                        End If
    '                                    Next
    '                                Next
    '                            Next
    '                            If _xyzw = True Then
    '                                If Not blnClassic Then
    '                                    Select Case intType
    '                                        Case 1
    '                                            For j = 0 To UBound(arrPos)
    '                                                arrPos(j) = intSamuraiOffset(arrPos(j), g)
    '                                            Next
    '                                        Case 2
    '                                            For j = 0 To UBound(arrPos2)
    '                                                arrPos2(j) = intSamuraiOffset(arrPos2(j), g)
    '                                            Next
    '                                        Case 3
    '                                            For j = 0 To UBound(arrPos3)
    '                                                arrPos3(j) = intSamuraiOffset(arrPos3(j), g)
    '                                            Next
    '                                    End Select
    '                                End If
    '                                '---hint code---
    '                                If blnHint Then
    '                                    strHint = "xyz_" & intReverseBit(zBit) & "_" & strMoves
    '                                    If strHint <> strHintDetails Then
    '                                        strHintDetails = strHint
    '                                        intHintCount = 0
    '                                    Else
    '                                        intHintCount += 1
    '                                    End If
    '                                    Select Case intHintCount
    '                                        Case 0
    '                                            strHintMsg = "Look for an XYZ wing pattern"
    '                                        Case Else
    '                                            strHintMsg = "Look for an XYZ wing pattern featuring digits " & strXYZ
    '                                    End Select
    '                                    If blnSamurai Then
    '                                        strHintMsg += " (in samurai grid & " & g & ")"
    '                                    End If
    '                                    ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
    '                                    Exit Function
    '                                End If
    '                                '---end hint code---
    '                                If blnStep Then
    '                                    Select Case intType
    '                                        Case 1
    '                                            HighlightMove(strStepMsg:="XYZ wing pattern found. The cell containing digits " & strXYZ & " can see two cells containing digits " & bitmask2Str(yzBit) & " and " & bitmask2Str(xzBit) & ". Therefore the solution to one of the three highlighted cells must be " & intReverseBit(zBit) & " which can be removed from any cells that can see the three highlighted cells as shown.", arrCells:=arrPos, arrCells2:=arrPos, arrCells3:=arrRemove, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hc3:=My.Settings._rColour, hBit1:=xyzBit, hBit2:=zBit, hBit3:=zBit, blnRemove1:=False, blnRemove2:=False, blnRemove3:=True)
    '                                        Case 2
    '                                            HighlightMove(strStepMsg:="XYZ wing pattern found. The cell containing digits " & strXYZ & " can see two cells containing digits " & bitmask2Str(yzBit2) & " and " & bitmask2Str(xzBit2) & ". Therefore the solution to one of the three highlighted cells must be " & intReverseBit(zBit) & " which can be removed from any cells that can see the three highlighted cells as shown.", arrCells:=arrPos2, arrCells2:=arrPos2, arrCells3:=arrRemove, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hc3:=My.Settings._rColour, hBit1:=xyzBit, hBit2:=zBit, hBit3:=zBit, blnRemove1:=False, blnRemove2:=False, blnRemove3:=True)
    '                                        Case 3
    '                                            HighlightMove(strStepMsg:="XYZ wing pattern found. The cell containing digits " & strXYZ & " can see two cells containing digits " & bitmask2Str(yzBit3) & " and " & bitmask2Str(xzBit3) & ". Therefore the solution to one of the three highlighted cells must be " & intReverseBit(zBit) & " which can be removed from any cells that can see the three highlighted cells as shown.", arrCells:=arrPos3, arrCells2:=arrPos3, arrCells3:=arrRemove, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hc3:=My.Settings._rColour, hBit1:=xyzBit, hBit2:=zBit, hBit3:=zBit, blnRemove1:=False, blnRemove2:=False, blnRemove3:=True)
    '                                    End Select
    '                                    Exit Function
    '                                End If
    '                                '---saves the move into the stack---
    '                                If strSolveUpToMethod <> "" And strMoves <> "" Then
    '                                    frmGame.Moves.Push(strMoves)
    '                                End If
    '                                Exit Function
    '                            End If
    '                        End If
    '                        '---end check---
    '                    Next
    '                End If
    '            End If
    '        Next
    '    Next
    'End Function
#End Region
#Region "Simple colouring"
    Public Function _cw() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim intMaxGrid As Integer = 5
        Dim arrCells(0) As Integer
        Dim arrCells2(0) As Integer
        Dim strType As String = ""
        Dim strHint As String = ""
        Dim strMoves As String = ""
        Dim intBit As Integer
        Dim strMatch As String = ""
        Dim c As Integer
        Dim g As Integer
        Dim v As Integer
        Dim i As Integer ' counter
        Dim j As Integer ' counter
        If blnClassic Then intMaxGrid = 1
        Dim Cells(0) As List(Of Integer)
        Dim pCells(0) As List(Of Integer)
        Dim nCells(0) As List(Of Integer)
        pCells(0) = New List(Of Integer)
        nCells(0) = New List(Of Integer)
        Dim aR(0) As Integer
        Dim aC(0) As Integer
        Dim aG(0) As Integer

        Dim ptr As Integer
        Dim arrayPeers() As String
        For v = 1 To 9
            For g = 1 To intMaxGrid
                Cells(0) = _colourCandidates(g, v, aR, aC, aG)
                While Cells(0).Count > 3
                    _pnCandidates(Cells, pCells, nCells, aR, aC, aG, True)
                    If (pCells(0).Count > 0 AndAlso nCells(0).Count > 0) AndAlso (pCells(0).Count + nCells(0).Count) > 3 Then
                        pCells(0).Sort()
                        nCells(0).Sort()
                        '---check to see if same colours appear in row, column or grid
                        For i = 0 To pCells(0).Count - 1
                            ptr = pCells(0).Item(i)
                            arrayPeers = arrPeers(ptr)
                            For j = 0 To pCells(0).Count - 1
                                If Array.IndexOf(arrayPeers, CStr(pCells(0).Item(j))) > -1 Then
                                    If strMatch = "" AndAlso (intFindRow(ptr) = intFindRow(pCells(0).Item(j))) Then
                                        strMatch = "row " & intFindRow(ptr)
                                    End If
                                    If strMatch = "" AndAlso (intFindCol(ptr) = intFindCol(pCells(0).Item(j))) Then
                                        strMatch = "column " & intFindCol(ptr)
                                    End If
                                    If strMatch = "" AndAlso (intFindGrid(ptr) = intFindGrid(pCells(0).Item(j))) Then
                                        strMatch = "subgrid " & intFindGrid(ptr)
                                    End If
                                    _cw = True
                                    strType = "p"
                                    Exit For
                                End If
                            Next
                            If _cw = True Then Exit For
                        Next

                        For i = 0 To nCells(0).Count - 1
                            ptr = nCells(0).Item(i)
                            arrayPeers = arrPeers(ptr)
                            For j = 0 To nCells(0).Count - 1
                                If Array.IndexOf(arrayPeers, CStr(nCells(0).Item(j))) > -1 Then
                                    If strMatch = "" AndAlso (intFindRow(ptr) = intFindRow(nCells(0).Item(j))) Then
                                        strMatch = "row " & intFindRow(ptr)
                                    End If
                                    If strMatch = "" AndAlso (intFindCol(ptr) = intFindCol(nCells(0).Item(j))) Then
                                        strMatch = "column " & intFindCol(ptr)
                                    End If
                                    If strMatch = "" AndAlso (intFindGrid(ptr) = intFindGrid(nCells(0).Item(j))) Then
                                        strMatch = "subgrid " & intFindGrid(ptr)
                                    End If
                                    _cw = True
                                    strType = "n"
                                    Exit For
                                End If
                            Next
                            If _cw = True Then Exit For
                        Next

                        If _cw Then

                            c = -1
                            ReDim arrCells(0)
                            For i = 0 To pCells(0).Count - 1
                                ptr = pCells(0).Item(i)
                                If Not blnClassic Then ptr = intSamuraiOffset(ptr, g)
                                c += 1
                                ReDim Preserve arrCells(c)
                                arrCells(c) = ptr
                                If strType = "p" Then
                                    '---remove code---
                                    intBit = intGetBit(v)
                                    _vsCandidateAvailableBits(ptr) -= intBit
                                    If strSolveUpToMethod <> "" Then
                                        Dim cCtrl As SudokuCell
                                        cCtrl = getCtrl("SudokuCell" & ptr)
                                        cCtrl.sc_IntCandidate -= intBit
                                    End If
                                    strMoves += ptr & ":" & "-" & v & "|"
                                    '---end remove code---
                                End If
                            Next
                            c = -1
                            ReDim arrCells2(0)
                            For i = 0 To nCells(0).Count - 1
                                ptr = nCells(0).Item(i)
                                If Not blnClassic Then ptr = intSamuraiOffset(ptr, g)
                                c += 1
                                ReDim Preserve arrCells2(c)
                                arrCells2(c) = ptr
                                If strType = "n" Then
                                    '---remove code---
                                    intBit = intGetBit(v)
                                    _vsCandidateAvailableBits(ptr) -= intBit
                                    If strSolveUpToMethod <> "" Then
                                        Dim cCtrl As SudokuCell
                                        cCtrl = getCtrl("SudokuCell" & ptr)
                                        cCtrl.sc_IntCandidate -= intBit
                                    End If
                                    strMoves += ptr & ":" & "-" & v & "|"
                                    '---end remove code---
                                End If
                            Next

                            '---hint code---
                            If blnHint Then
                                strHint = "cw_" & v & "_" & strMoves
                                If strHint <> strHintDetails Then
                                    strHintDetails = strHint
                                    intHintCount = 0
                                Else
                                    intHintCount += 1
                                End If
                                Select Case intHintCount
                                    Case 0
                                        strHintMsg = "Look for a colour wrap pattern"
                                    Case Else
                                        strHintMsg = "Look for a colour wrap pattern featuring digit " & v
                                End Select
                                If blnSamurai Then
                                    strHintMsg += " (in samurai grid & " & g & ")"
                                End If
                                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                                Exit Function
                            End If
                            '---end hint code---
                            If blnStep Then
                                Select Case strType
                                    Case "p"
                                        HighlightMove(strStepMsg:="Colour wrap pattern of digit " & v & " found. Strongly linked pairs of digit " & v & " have been alternately coloured. In " & strMatch & " there are digits with the same colour. This means all digits with this colour can be removed.", arrCells:=arrCells, arrCells2:=arrCells2, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hBit1:=intGetBit(v), hBit2:=intGetBit(v), blnRemove1:=True, blnRemove2:=False)
                                    Case "n"
                                        HighlightMove(strStepMsg:="Colour wrap pattern of digit " & v & " found. Strongly linked pairs of digit " & v & " have been alternately coloured. In " & strMatch & " there are digits with the same colour. This means all digits with this colour can be removed.", arrCells:=arrCells, arrCells2:=arrCells2, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hBit1:=intGetBit(v), hBit2:=intGetBit(v), blnRemove1:=False, blnRemove2:=True)
                                End Select
                                Exit Function
                            End If
                            '---saves the move into the stack---
                            If strSolveUpToMethod <> "" And strMoves <> "" Then
                                frmGame.Moves.Push(strMoves)
                            End If
                            Exit Function
                        End If
                    End If
                    pCells(0).Clear()
                    nCells(0).Clear()
                End While
            Next
        Next
    End Function
    Public Function _ct() As Boolean
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function
        Dim intMaxGrid As Integer = 5
        Dim arrCells(0) As Integer
        Dim arrCells2(0) As Integer
        Dim strType As String = ""
        Dim strHint As String = ""
        Dim strMoves As String = ""
        Dim intBit As Integer
        Dim strMatch As String = ""
        Dim ptr(0) As Integer
        Dim intCell As Integer
        Dim intPtr As Integer
        Dim c As Integer
        Dim g As Integer
        Dim v As Integer
        Dim i As Integer ' counter
        Dim j As Integer ' counter
        Dim p As Integer = -1 ' counter
        Dim item As Integer
        If blnClassic Then intMaxGrid = 1
        Dim Cells(0) As List(Of Integer)
        Dim pCells(0) As List(Of Integer)
        Dim nCells(0) As List(Of Integer)
        pCells(0) = New List(Of Integer)
        nCells(0) = New List(Of Integer)
        Dim aR(0) As Integer
        Dim aC(0) As Integer
        Dim aG(0) As Integer
        Dim blnPosMatch As Boolean
        Dim blnNegMatch As Boolean
        Dim arrayPeers() As String
        For v = 1 To 9
            For g = 1 To intMaxGrid
                Cells(0) = _colourCandidates(g, v, aR, aC, aG)
                While Cells(0).Count >= 3
                    _pnCandidates(Cells, pCells, nCells, aR, aC, aG, True)
                    If (pCells(0).Count > 0 AndAlso nCells(0).Count > 0) AndAlso (pCells(0).Count + nCells(0).Count) >= 3 Then
                        pCells(0).Sort()
                        nCells(0).Sort()
                        '---check to see if cells can see opposite colours---
                        For i = 1 To 81
                            If blnClassic Then
                                intPtr = i
                            Else
                                intPtr = intSamuraiOffset(i, g)
                            End If

                            If _vsSolution(intPtr - 1) = 0 AndAlso (intGetBit(v) And _vsCandidateAvailableBits(intPtr)) Then
                                If pCells(0).IndexOf(i) = -1 AndAlso nCells(0).IndexOf(i) = -1 Then
                                    arrayPeers = arrPeers(i)
                                    blnPosMatch = False
                                    blnNegMatch = False
                                    For Each j In pCells(0)
                                        If Array.IndexOf(arrayPeers, CStr(j)) > -1 Then
                                            blnPosMatch = True
                                            Exit For
                                        End If
                                    Next
                                    For Each j In nCells(0)
                                        If Array.IndexOf(arrayPeers, CStr(j)) > -1 Then
                                            blnNegMatch = True
                                            Exit For
                                        End If
                                    Next
                                    If blnPosMatch And blnNegMatch Then
                                        _ct = True
                                        '---remove code---
                                        p += 1
                                        intBit = intGetBit(v)
                                        ReDim Preserve ptr(p)
                                        ptr(p) = intPtr
                                        _vsCandidateAvailableBits(intPtr) -= intBit
                                        If strSolveUpToMethod <> "" Then
                                            Dim cCtrl As SudokuCell
                                            cCtrl = getCtrl("SudokuCell" & intPtr)
                                            cCtrl.sc_IntCandidate -= intBit
                                        End If
                                        strMoves += ptr(p) & ":" & "-" & v & "|"
                                        '---end remove code---
                                    End If
                                End If
                            End If
                        Next

                        If _ct Then
                            
                            '---hint code---
                            If blnHint Then
                                strHint = "ct_" & v & "_" & strMoves
                                If strHint <> strHintDetails Then
                                    strHintDetails = strHint
                                    intHintCount = 0
                                Else
                                    intHintCount += 1
                                End If
                                Select Case intHintCount
                                    Case 0
                                        strHintMsg = "Look for a colour trap pattern"
                                    Case Else
                                        strHintMsg = "Look for a colour trap pattern featuring digit " & v
                                End Select
                                If blnSamurai Then
                                    strHintMsg += " (in samurai grid & " & g & ")"
                                End If
                                ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
                                Exit Function
                            End If
                            '---end hint code---
                            If blnStep Then
                                ReDim arrCells(0)
                                c = -1
                                For Each item In pCells(0)
                                    intCell = item
                                    If Not blnClassic Then intCell = intSamuraiOffset(item, g)
                                    c += 1
                                    ReDim Preserve arrCells(c)
                                    arrCells(c) = intCell
                                Next
                                ReDim arrCells2(0)
                                c = -1
                                For Each item In nCells(0)
                                    intCell = item
                                    If Not blnClassic Then intCell = intSamuraiOffset(item, g)
                                    c += 1
                                    ReDim Preserve arrCells2(c)
                                    arrCells2(c) = intCell
                                Next
                                HighlightMove(strStepMsg:="Colour trap pattern of digit " & v & " found. Strongly linked pairs of digit " & v & " have been alternately coloured. Digit " & v & " can be removed where it is is visible to alternatiing colours.", arrCells:=arrCells, arrCells2:=arrCells2, arrCells3:=ptr, hc1:=My.Settings._pColour, hc2:=My.Settings._nColour, hc3:=My.Settings._rColour, hBit1:=intBit, hBit2:=intBit, hBit3:=intBit, blnRemove1:=False, blnRemove2:=False, blnRemove3:=True)
                                Exit Function
                            End If
                            '---saves the move into the stack---
                            If strSolveUpToMethod <> "" And strMoves <> "" Then
                                frmGame.Moves.Push(strMoves)
                            End If
                            '---end save move---
                            Exit Function
                        End If

                    End If
                    pCells(0).Clear()
                    nCells(0).Clear()
                End While
            Next
        Next
    End Function
    Private Function _pnCandidates(ByRef Cells() As List(Of Integer), ByRef pCells() As List(Of Integer), ByRef nCells() As List(Of Integer), ByVal r() As Integer, ByVal c() As Integer, ByVal g() As Integer, ByVal isPositive As Boolean, Optional ByVal oPtr As Integer = 0, Optional ByVal cPtr As Integer = 0) As Boolean
        '---colouring - sets cells to positive or negative---
        Dim i As Integer
        Dim j As Integer
        Dim arrayPeers() As String
        Dim ptr As Integer
        Dim rcg As Integer
        Dim pPtr As Integer
        Dim cCount As Integer
        Dim blnRCGMatch As Boolean = False
        If Cells(0).Count = 0 Then Exit Function
        ptr = cPtr
        If ptr = 0 Then ptr = Cells(0).Item(0)
        arrayPeers = arrPeers(ptr)

        If isPositive Then
            '---positive---
            If cPtr > 0 Then
                For j = 1 To 3
                    Select Case j
                        Case 1
                            rcg = intFindRow(oPtr)
                            If (rcg = intFindRow(cPtr)) AndAlso Array.IndexOf(r, rcg) > -1 Then
                                blnRCGMatch = True
                                Exit For
                            End If
                        Case 2
                            rcg = intFindCol(oPtr)
                            If (rcg = intFindCol(cPtr)) AndAlso Array.IndexOf(c, rcg) > -1 Then
                                blnRCGMatch = True
                                Exit For
                            End If
                        Case 3
                            rcg = intFindGrid(oPtr)
                            If (rcg = intFindGrid(cPtr)) AndAlso Array.IndexOf(g, rcg) > -1 Then
                                blnRCGMatch = True
                                Exit For
                            End If
                    End Select
                Next
                If blnRCGMatch Then
                    pCells(0).Add(ptr)
                Else
                    '---no match for row, column or grid---'
                    Exit Function
                End If
            Else
                pCells(0).Add(ptr)
            End If
        Else
            '---negative---
            If cPtr > 0 Then
                For j = 1 To 3
                    Select Case j
                        Case 1
                            rcg = intFindRow(oPtr)
                            If (rcg = intFindRow(cPtr)) AndAlso Array.IndexOf(r, rcg) > -1 Then
                                blnRCGMatch = True
                                Exit For
                            End If
                        Case 2
                            rcg = intFindCol(oPtr)
                            If (rcg = intFindCol(cPtr)) AndAlso Array.IndexOf(c, rcg) > -1 Then
                                blnRCGMatch = True
                                Exit For
                            End If
                        Case 3
                            rcg = intFindGrid(oPtr)
                            If (rcg = intFindGrid(cPtr)) AndAlso Array.IndexOf(g, rcg) > -1 Then
                                blnRCGMatch = True
                                Exit For
                            End If
                    End Select
                Next
                If blnRCGMatch Then
                    nCells(0).Add(ptr)
                Else
                    '---no match for row, column or grid---'
                    Exit Function
                End If
            Else
                nCells(0).Add(ptr)
            End If
        End If
        '---iterate peers---
        Cells(0).Remove(ptr)
        cCount = Cells(0).Count
        For i = 0 To Cells(0).Count - 1

            If cCount > Cells(0).Count AndAlso Cells(0).Count > 0 Then
                'reset counter as the list of cells has changed due to recursion
                i = 0
            End If

            If i <= Cells(0).Count - 1 Then
                pPtr = Cells(0).Item(i)
            Else
                'Catch
                pPtr = 0
            End If

            cCount = Cells(0).Count
            If pPtr > 0 AndAlso Array.IndexOf(arrayPeers, CStr(pPtr)) > -1 Then
                '---recursion---
                If isPositive Then
                    _pnCandidates(Cells, pCells, nCells, r, c, g, False, ptr, pPtr)
                Else
                    _pnCandidates(Cells, pCells, nCells, r, c, g, True, ptr, pPtr)
                    'i = 0
                End If
                '---end recursion---
            End If
        Next

        '---end iterate peers---
    End Function
    Private Function _colourCandidates(ByVal intGrid As Integer, ByVal v As Integer, ByRef r() As Integer, ByRef c() As Integer, ByRef g() As Integer) As List(Of Integer)
        Dim rcg As Integer
        Dim t As Integer
        Dim p As Integer
        Dim i As Integer
        Dim ptr As Integer
        ReDim r(0)
        ReDim c(0)
        ReDim g(0)
        Dim arrCells(0) As String
        Dim posArray(9) As String
        Dim valueArray(0) As String
        Dim Cells(0) As List(Of Integer)
        Cells(0) = New List(Of Integer)
        For t = 1 To 3
            i = -1
            For rcg = 1 To 9
                Select Case t
                    Case 1
                        arrCells = arrRow(rcg)
                    Case 2
                        arrCells = arrCol(rcg)
                    Case 3
                        arrCells = arrGrid(rcg)
                End Select
                ReDim posArray(9)
                ReDim valueArray(0)
                _countCandidatesinRCG(arrCells, intGrid, 2, posArray, valueArray)
                If valueArray Is Nothing Then
                Else
                    If posArray(v) Is Nothing Then
                    Else
                        '---record r/c/g---'
                        i += 1
                        Select Case t
                            Case 1
                                ReDim Preserve r(i)
                                r(i) = rcg
                            Case 2
                                ReDim Preserve c(i)
                                c(i) = rcg
                            Case 3
                                ReDim Preserve g(i)
                                g(i) = rcg
                        End Select
                        '---end record r/c/g---'
                        posArray = Split(posArray(v), arrDivider)
                        For p = 0 To UBound(posArray)
                            ptr = arrCells(posArray(p) - 1)
                            If Cells(0).IndexOf(ptr) = -1 Then
                                Cells(0).Add(ptr)
                            End If
                        Next
                    End If
                End If
            Next
        Next
        Cells(0).Sort()
        _colourCandidates = Cells(0)
    End Function
#End Region
#Region "Brute force"
    Public Function _bf() As Boolean
        If strInputGameSolution = "" Then Exit Function
        If _vsUnsolvedCells(0).Count = 0 Then Exit Function

        Dim ptr As Integer
        Dim i As Integer
        Dim j As Integer
        Dim arrayPeers(0) As String
        Dim intBit As Integer
        Dim strHint As String
        Dim intValue As Integer

        strInputGameSolution = Replace(strInputGameSolution, vbCrLf, "")

        Select Case blnSamurai
            Case True
                '---samurai---
                Dim g As Integer
                Dim aPtr As Integer
                For g = 1 To 5
                    For i = 1 To 81
                        ptr = intSamuraiOffset(i, g)
                        intValue = Mid(strInputGameSolution, i + (81 * (g - 1)), 1)
                        Select Case Asc(intValue)
                            Case 49 To 57
                                '---numeric---
                                aPtr = _vsUnsolvedCells(0).IndexOf(ptr)
                                If aPtr > -1 Then
                                    ptr = _vsUnsolvedCells(0).Item(aPtr)
                                    i = 82
                                    g = 6
                                    Exit For
                                End If
                        End Select
                    Next
                Next
            Case False
                '---classic---
                ptr = _vsUnsolvedCells(0).Item(0)
                intValue = CInt(Mid(strInputGameSolution, ptr, 1))
        End Select

        _vsSolution(ptr - 1) = intValue
        _vsCandidateCount(ptr) = -1
        _vsUnsolvedCells(0).Remove(ptr)
        _u -= 1
        Select Case blnClassic
            Case True
                arrayPeers = arrPeers(ptr)
            Case False
                arrayPeers = ArrSamuraiPeers(ptr)
        End Select
        intBit = intGetBit(intValue)
        'remove value from peers
        For j = 0 To UBound(arrayPeers)
            If _vsSolution(arrayPeers(j) - 1) = 0 AndAlso (_vsCandidateAvailableBits(arrayPeers(j)) And intBit) Then
                _vsCandidateAvailableBits(arrayPeers(j)) -= intBit
                _vsCandidateCount(arrayPeers(j)) -= 1
            End If
        Next
        _bf = True
        If strSolveUpToMethod <> "" Then frmGame.nakedSingle_doubleClick(intCell:=ptr, intValue:=intValue, strSolvedBy:="P")

        If blnStep Then
            Dim hm(0) As Integer
            hm(0) = ptr
            HighlightMove(arrCells:=hm, strStepMsg:="Brute force selects digit " & intValue & " to be placed.", hc1:=My.Settings._cColour, hBit1:=intBit)
            frmGame.nakedSingle_doubleClick(intCell:=ptr, intValue:=intValue, strSolvedBy:="P")
        End If

        '---hint code---
        If blnHint Then
            strHint = "BF_" & ptr & "_" & intValue
            If strHint <> strHintDetails Then
                strHintDetails = strHint
                intHintCount = 0
            Else
                intHintCount += 1
            End If
            strHintMsg = "No logic based methods can progress."
            ShowMessageLabel(strMsg:=strHintMsg, strImage:="info")
            Exit Function
        End If
        '---end hint code---

    End Function
#End Region
#Region "Misc solving code"
    Function HighlightMove(ByVal strStepMsg As String, Optional ByVal arrCells() As Integer = Nothing, Optional ByVal arrCells2() As Integer = Nothing, Optional ByVal arrCells3() As Integer = Nothing, Optional ByVal arrCells4() As Integer = Nothing, Optional ByVal hc1 As String = "", Optional ByVal hc2 As String = "", Optional ByVal hc3 As String = "", Optional ByVal hc4 As String = "", Optional ByVal hBit1 As Integer = 0, Optional ByVal hBit2 As Integer = 0, Optional ByVal hBit3 As Integer = 0, Optional ByVal hBit4 As Integer = 0, Optional ByVal blnRemove1 As Boolean = False, Optional ByVal blnRemove2 As Boolean = False, Optional ByVal blnRemove3 As Boolean = False, Optional ByVal blnRemove4 As Boolean = False) As Boolean
        Dim sc As SudokuCell
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim v As Integer
        Dim ptr As Integer
        Dim arrUpdate(0) As Integer
        Dim intCurBit As Integer = 0
        Dim strCol As String = ""
        Dim blnRemove As Boolean
        Dim arrRemove() As Integer
        Dim strMoves As String = ""
        For j = 1 To 4
            Select Case j
                Case 1
                    arrUpdate = arrCells
                    intCurBit = hBit1
                    strCol = hc1
                Case 2
                    arrUpdate = arrCells2
                    intCurBit = hBit2
                    strCol = hc2
                Case 3
                    arrUpdate = arrCells3
                    intCurBit = hBit3
                    strCol = hc3
                Case 4
                    arrUpdate = arrCells4
                    intCurBit = hBit4
                    strCol = hc4
            End Select

            If arrUpdate Is Nothing Then
            Else
                For i = 0 To UBound(arrUpdate)
                    sc = getCtrl("SudokuCell" & arrUpdate(i))
                    sc.sc_HighlightNaked = False
                    If blnSingleBit(sc.sc_IntCandidate) And My.Settings._ShowNS And My.Settings._blnCandidates Then
                        With sc
                            .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intMO_BGTransparency)
                            .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intM0_BorderTransparency)
                        End With
                    Else
                        Select Case j
                            Case 1
                                If strCol <> "" Then sc.sc_HighlightColour1 = Color.FromName(strCol)
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate1 = intCurBit
                            Case 2
                                If strCol <> "" Then sc.sc_HighlightColour2 = Color.FromName(strCol)
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate2 = intCurBit
                            Case 3
                                If strCol <> "" Then sc.sc_HighlightColour3 = Color.FromName(strCol)
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate3 = intCurBit
                            Case 4
                                If strCol <> "" Then sc.sc_HighlightColour4 = Color.FromName(strCol)
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate4 = intCurBit
                        End Select
                    End If
                Next
            End If
        Next

        If blnAllSteps Then
            frmHidden.strMsg = strStepMsg & vbCrLf & vbCrLf & "Press any key or ESC to exit."
        Else
            frmHidden.strMsg = strStepMsg & vbCrLf & vbCrLf & "Press any key."
        End If
        frmHidden.ShowDialog(frmGame)

        For j = 1 To 4
            Select Case j
                Case 1
                    arrUpdate = arrCells
                    intCurBit = hBit1
                    strCol = hc1
                    blnRemove = blnRemove1
                Case 2
                    arrUpdate = arrCells2
                    intCurBit = hBit2
                    strCol = hc2
                    blnRemove = blnRemove2
                Case 3
                    arrUpdate = arrCells3
                    intCurBit = hBit3
                    strCol = hc3
                    blnRemove = blnRemove3
                Case 4
                    arrUpdate = arrCells4
                    intCurBit = hBit4
                    strCol = hc4
                    blnRemove = blnRemove4
            End Select

            If arrUpdate Is Nothing Then
            Else
                '---undo any colouring---
                For i = 0 To UBound(arrUpdate)
                    ptr = arrUpdate(i)
                    sc = getCtrl("SudokuCell" & ptr)
                    sc.sc_HighlightNaked = True
                    If blnSingleBit(sc.sc_IntCandidate) And My.Settings._ShowNS And My.Settings._blnCandidates Then
                        With sc
                            .sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                            .sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                        End With
                    Else
                        Select Case j
                            Case 1
                                If strCol <> "" Then sc.sc_HighlightColour1 = Nothing
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate1 = 0
                            Case 2
                                If strCol <> "" Then sc.sc_HighlightColour2 = Nothing
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate2 = 0
                            Case 3
                                If strCol <> "" Then sc.sc_HighlightColour3 = Nothing
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate3 = 0
                            Case 4
                                If strCol <> "" Then sc.sc_HighlightColour4 = Nothing
                                If intCurBit <> 0 Then sc.sc_IntHighlightCandidate4 = 0
                        End Select
                        If blnRemove Then
                            arrRemove = bitmask2Arr(intCurBit)
                            For k = 0 To UBound(arrRemove)
                                v = arrRemove(k)
                                If sc.sc_IntSolution = 0 AndAlso (intGetBit(v) And sc.sc_IntCandidate) Then
                                    strMoves += ptr & ":" & "-" & v & "|"
                                    sc.sc_IntCandidate -= intGetBit(v)
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        Next

        If strMoves <> "" Then frmGame.Moves.Push(strMoves)

    End Function
    Private Function _SolveMethods(ByVal blnHint As Boolean) As Boolean
        Dim i As Integer
        Dim blnResult As Boolean
        Dim strMethod As String = ""

        Dim arrMethods() As String = Split(vsSolvers, arrDivider)
        Dim arrDesc() As String
        Dim wordArray(0) As String

        ReDim solveMethods(arrMethods.Length - 1)
        ReDim solveCountMethods(arrMethods.Length - 1)

        strDifficulty = ""
        intDifficulty = 0

        For i = 0 To UBound(arrMethods)
            If _vsUnsolvedCells(0).Count = 0 Then Exit For
            arrDesc = Split(arrMethods(i), ":")
            strMethod = arrDesc(2)

            If InStr(strMethod, "-") > 0 Then
                wordArray = Split(strMethod, "-")
            End If
            If InStr(strMethod, " ") > 0 Then
                wordArray = Split(strMethod, " ")
            End If
            If InStr(strMethod, "fish") > 0 Then
                wordArray = Split(strMethod, "fish")
                wordArray(1) = "fish"
            End If

            strMethod = ""
            Dim w As Integer
            For w = 0 To UBound(wordArray)
                Select Case wordArray(w)
                    Case "XYZ"
                        strMethod += "xyz"
                    Case "XY"
                        strMethod += "xy"
                    Case Else
                        strMethod += Microsoft.VisualBasic.Left(wordArray(w), 1)
                End Select
            Next
            strMethod = LCase("_" & strMethod)

            If StopThreadWork Then
                intCountSolutions = 0
                RaiseEvent UI(0, intCountSolutions, True, "", "")
                Exit Function
            End If

            Try
                Dim callF As System.Reflection.MethodInfo = Me.GetType().GetMethod(strMethod)
                blnResult = callF.Invoke(Me, Nothing)
            Catch
                MsgBox("An error occurred in the method: " & strMethod)
            End Try
            If blnResult = True Then
                If solveMethods(i) = "" Then solveMethods(i) = arrDesc(2)
                solveCountMethods(i) += 1
                If CInt(arrDesc(0)) > intDifficulty Then
                    intDifficulty = arrDesc(0)
                    strDifficulty = intDifficult2String(intDifficulty)
                End If
                i = -1
                If blnHint Then Exit Function
                If blnStep Then Exit Function
            Else
                If strSolveUpToMethod = "_" & arrDesc(2) Then
                    '---exit as only solving up to this method---
                    Exit Function
                End If
            End If
        Next

        If blnHint Then
            ShowMessageLabel(strMsg:="No hints available", strImage:="alert")
            Exit Function
        End If

    End Function
    Private Function _countCandidatesinRCG(ByVal arrCells() As String, ByVal intGrid As Integer, ByVal n As Integer, ByRef posArray() As String, ByRef valueArray() As String) As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim v As Integer
        Dim ptr As Integer
        Dim vCount(9) As Integer
        Dim pArray(9) As Integer
        Dim strCandidates As String
        Dim arrCandidates() As String
        For i = 0 To 8
            Select Case blnClassic
                Case True
                    ptr = arrCells(i)
                Case False
                    ptr = intSamuraiOffset(arrCells(i), intGrid)
            End Select
            If _vsSolution(ptr - 1) = 0 Then
                If _vsCandidateCount(ptr) >= 2 Then
                    strCandidates = bitmask2Str(_vsCandidateAvailableBits(ptr), arrDivider)
                    arrCandidates = Split(strCandidates, arrDivider)
                    For j = 0 To UBound(arrCandidates)
                        vCount(arrCandidates(j)) += 1
                        pArray(arrCandidates(j)) += intGetBit(i + 1)
                    Next
                End If
            End If
        Next
        ReDim valueArray(0)
        For j = 1 To 9
            If vCount(j) >= 2 And vCount(j) <= n Then
                posArray(j) = bitmask2Str(pArray(j), arrDivider)
                ReDim Preserve valueArray(v)
                valueArray(v) = j
                v += 1
            End If
        Next
    End Function
#End Region
#End Region
#Region "Brute force / backtracking code"
    Sub _Unique()
        _vsUnique()
    End Sub
    Public Function _vsUnique() As Boolean
        Dim strSolution As String = ""
        If _vsbackTrack(strGrid:=strGrid, StrSolution:=strSolution, StrCandidates:=strCandidates) Then
            _vsUnique = True
            strGameSolution = strSolution
        End If
    End Function
    Private Function _load(ByVal strGrid As String, Optional ByVal StrCandidates As String = "") As Boolean
        '---load puzzle---
        _vsSteps = 1
        vsTried = 0
        ReDim _vsUnsolvedCells(0)
        Dim i As Integer
        Dim intCellOffset As Integer
        Dim strClues As String = ""
        Dim g As Integer
        Dim j As Integer
        Dim intBit As Integer
        Dim blnCandidates As Boolean = False
        Dim arrCandidates() As String = Split(StrCandidates, arrDivider)
        If arrCandidates.Length >= 81 Then blnCandidates = True
        _u = -1
        _vsCandidateCount(0) = -1
        For i = 1 To _vsCandidateCount.Length - 1
            _vsCandidateAvailableBits(i) = 511
            _vsStoreCandidateBits(i) = 0
            _vsCandidateCount(i) = -1
            If blnClassic = False Then
                If Not blnIgnoreSamurai(i) Then _vsCandidateCount(i) = 9
            Else
                _vsCandidateCount(i) = 9
            End If
            _vsLastGuess(i) = 0
            _vsCandidatePtr(i) = 1
            _vsSolution(i - 1) = 0
            _vsPeers(i) = 0
        Next

        strGrid = Trim(strGrid)
        Dim midStr As String = ""
        Dim ptr As Integer
        Dim arrayPeers(0) As String
        Dim intValue As Integer
        Dim nextGuess As Integer = 0
        Dim nextCandidate As Integer = 0
        _vsUnsolvedCells(0) = New List(Of Integer)
        Dim intMaxGrid As Integer = 5
        If blnClassic Then intMaxGrid = 1
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnClassic
                    Case True
                        midStr = Mid(strGrid, i, 1)
                        intCellOffset = i
                    Case False
                        midStr = Mid(strGrid, i + (81 * (g - 1)), 1)
                        intCellOffset = intSamuraiOffset(i, g)
                End Select
                Select Case Asc(midStr)
                    Case 46, 48
                        '---blank---
                        If (blnClassic Or Not blnIgnoreSamurai(intCellOffset)) AndAlso _vsUnsolvedCells(0).IndexOf(intCellOffset) = -1 Then
                            _u += 1
                            _vsUnsolvedCells(0).Add(intCellOffset)
                            If blnCandidates = True Then
                                '---insert known candidates---
                                _vsCandidateAvailableBits(intCellOffset) = arrCandidates(intCellOffset - 1)
                                _vsCandidateCount(intCellOffset) = intCountBits(arrCandidates(intCellOffset - 1))
                            End If
                        End If
                    Case 49 To 57
                        '---numeric 1 to 9---
                        intValue = CInt(midStr)
                        intBit = intGetBit(intValue)
                        If _vsSolution(intCellOffset - 1) = 0 Then
                            _vsSolution(intCellOffset - 1) = intValue
                            _vsCandidateCount(intCellOffset) = -1
                            If blnCandidates = False Then
                                Select Case blnClassic
                                    Case True
                                        arrayPeers = arrPeers(intCellOffset)
                                    Case False
                                        arrayPeers = ArrSamuraiPeers(intCellOffset)
                                End Select
                                '---remove value from peers---
                                For j = 0 To UBound(arrayPeers)
                                    ptr = arrayPeers(j)
                                    If _vsCandidateAvailableBits(ptr) And intBit Then
                                        _vsCandidateAvailableBits(ptr) -= intBit
                                        _vsCandidateCount(ptr) -= 1
                                    End If
                                Next
                            End If
                        End If
                    Case Else
                        Debug.Print("exiting due to invalid character " & Asc(midStr))
                        _load = False
                        Exit Function
                End Select
                strClues += midStr
            Next
            If Not blnClassic Then strClues += vbCrLf
        Next
        _load = True
        strFormatClues = strClues
    End Function
    Private Function _vsbackTrack(ByVal strGrid As String, ByRef StrSolution As String, Optional ByVal StrCandidates As String = "") As Boolean
        Dim intMax As Integer = 0
        Dim intSolutionMax As Integer = 0
        ReDim Solutions(0)
        Dim i As Integer
        Dim j As Integer
        Dim intSolutions As Integer
        Dim testPeers(0) As String
        Dim tempPeers As String
        Dim nextGuess As Integer = 0
        Dim nextCandidate As Integer = 0
        Select Case blnClassic
            Case True
                intMax = 81
                intSolutionMax = 80
            Case False
                intMax = 441
                intSolutionMax = 440
        End Select
        ReDim _vsSolution(intSolutionMax)
        ReDim _vsPeers(intMax)
        ReDim _vsCandidateCount(intMax)
        ReDim _vsCandidateAvailableBits(intMax)
        ReDim _vsCandidatePtr(intMax)
        ReDim _vsLastGuess(intMax)
        ReDim _vsStoreCandidateBits(intMax)
        ReDim _vsRemovePeers(intMax)

        If Not _load(strGrid:=strGrid, StrCandidates:=StrCandidates) Then
            intCountSolutions = intSolutions
            RaiseEvent UI(vsTried, 0, True, "", "")
            Exit Function
        End If

        _vsUnsolvedCells(0).Sort()

        _SolveMethods(blnHint)

        If blnStep Then Return True

        If strSolveUpToMethod <> "" Then
            Select Case blnClassic
                Case True
                    StrSolution = strWriteSolution(intGrid:=1)
                Case False
                    StrSolution = strWriteSolution()
            End Select
            Return True
        End If

        If _u > -1 Then
            If Not blnClassic Then
                Dim g As Integer
                For g = 1 To 5
                    Dim Solver As New clsSudokuSolver
                    Solver.blnClassic = True
                    Solver.strGrid = strWriteSolution(intGrid:=g)
                    Solver.vsSolvers = My.Settings._UniqueSolvers
                    Solver.intQuit = 100
                    Solver._vsUnique()
                    If Solver.intCountSolutions > 1 AndAlso Solver.intCountSolutions < Solver.intQuit Then
                        Dim s As Integer
                        Dim c As Integer
                        Dim m(81) As Integer
                        Dim chk(81) As Boolean
                        Dim chr As String
                        Dim intChr As Integer
                        For c = 1 To 81
                            chk(c) = True
                        Next
                        For s = 0 To UBound(Solver.Solutions)
                            If Array.IndexOf(chk, True) = -1 Then Exit For
                            For c = 1 To 81
                                chr = Mid(Solver.Solutions(s), c, 1)
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
                        Dim strRevised As String = ""
                        Dim blnRevised As Boolean
                        Dim ptr As Integer
                        Dim arrayPeers() As String
                        Dim intBit As Integer
                        For c = 1 To 81
                            chr = Mid(Solver.strGrid, c, 1)
                            If chr = "." Then
                                '---unique value across all solutions---
                                '---and not found in starting grid---
                                If m(c) > 0 Then
                                    strRevised += CStr(m(c))
                                    blnRevised = True
                                    '---place solution---
                                    ptr = intSamuraiOffset(c, g)
                                    If _vsSolution(ptr - 1) = 0 Then
                                        _vsSolution(ptr - 1) = m(c)
                                        _vsCandidateCount(ptr) = -1
                                        _vsUnsolvedCells(0).Remove(ptr)
                                        arrayPeers = ArrSamuraiPeers(ptr)
                                        intBit = intGetBit(m(c))
                                        'remove value from peers
                                        For j = 0 To UBound(arrayPeers)
                                            If _vsSolution(arrayPeers(j) - 1) = 0 AndAlso (_vsCandidateAvailableBits(arrayPeers(j)) And intBit) Then
                                                _vsCandidateAvailableBits(arrayPeers(j)) -= intBit
                                                _vsCandidateCount(arrayPeers(j)) -= 1
                                            End If
                                        Next
                                        _u -= 1
                                    End If
                                    '--end place solution---
                                Else
                                    strRevised += chr
                                End If
                            Else
                                strRevised += chr
                            End If
                        Next
                        If blnRevised Then
                            'Debug.Print("Grid " & g & " has " & Solver.intCountSolutions & " possible solutions")
                            'Debug.Print(Solver.strGrid)
                            'Debug.Print(strRevised)
                            blnRevised = False
                        End If
                    End If
                Next
            End If
        End If

        If blnHint Then _u = -1

        '---setup peer array---
        For i = 0 To _u
            tempPeers = ""
            Select Case blnClassic
                Case True
                    testPeers = arrPeers(_vsUnsolvedCells(0).Item(i))
                Case False
                    testPeers = ArrSamuraiPeers(_vsUnsolvedCells(0).Item(i))
            End Select
            For j = 0 To UBound(testPeers)
                If _vsUnsolvedCells(0).IndexOf(CInt(testPeers(j))) > -1 Then
                    If tempPeers = "" Then
                        tempPeers = testPeers(j)
                    Else
                        tempPeers += "," & testPeers(j)
                    End If
                End If
            Next
            _vsPeers(_vsUnsolvedCells(0).Item(i)) = tempPeers
        Next
        '---end setup peer array---

        _r = 2
        _r2 = 2
        While _r Mod 2 = 0
            _r = RandomInt(1, 9999)
        End While
        While _r2 Mod 2 = 0
            _r2 = RandomInt(1, 12)
        End While

        If _u = -1 Then
            '---already solved by logic---
            intSolutions = 1
            Select Case blnClassic
                Case True
                    StrSolution = strWriteSolution(intGrid:=1)
                Case False
                    StrSolution = strWriteSolution()
            End Select
            intCountSolutions = intSolutions
            _vsbackTrack = True
            RaiseEvent UI(vsTried, intCountSolutions, True, StrSolution, strDifficulty)
            Exit Function
        End If

        While _vsSteps <= _u + 1 AndAlso _vsSteps > 0
            '---update for background threading---
            If vsTried Mod (50000 + _r) = 0 Then
                intCountSolutions = intSolutions
                RaiseEvent UI(vsTried, intCountSolutions, False, "", "")
            End If
            If StopThreadWork Then
                StrSolution = ""
                intCountSolutions = 0
                RaiseEvent UI(vsTried, intCountSolutions, True, "", "")
                Exit Function
            End If
            If nextGuess = 0 Then nextGuess = intFindCell()
            If nextGuess > 0 Then
                nextCandidate = IntNextCandidate(nextGuess)
                If nextCandidate > 0 Then
                    vsTried += 1
                    MakeGuess(nextGuess, nextCandidate)
                    nextGuess = 0
                Else
                    If _vsSteps <= 1 Then
                        Select Case intSolutions
                            Case 0
                                '---invalid puzzle - no solution---
                                _vsbackTrack = False
                                intCountSolutions = 0
                                RaiseEvent UI(vsTried, intCountSolutions, True, "", "")
                                Exit Function
                            Case 1
                                '---single solution---
                                _vsbackTrack = True
                                intCountSolutions = 1
                                RaiseEvent UI(vsTried, intCountSolutions, True, StrSolution, strDifficulty)
                                Exit Function
                            Case Else
                                '---multiple solutions---
                                _vsbackTrack = False
                                intCountSolutions = intSolutions
                                RaiseEvent UI(vsTried, intCountSolutions, True, "", "")
                                Exit Function
                        End Select
                    Else
                        '---need to go back...no remaining candidates for this cell---
                        UndoGuess(nextGuess)
                    End If
                End If
            Else
                If _vsSteps = 0 Then
                    _vsbackTrack = False
                    '---invalid puzzle...---
                    intCountSolutions = intSolutions
                    RaiseEvent UI(vsTried, intCountSolutions, True, "", "")
                    Exit Function
                Else
                    '---cannot go further...so need to go back---
                    UndoGuess()
                End If
            End If

            If _vsSteps > _u + 1 Then

                intSolutions += 1
                If intSolutions Mod ((intQuit / 10) + _r2) = 0 Then
                    RaiseEvent UI(vsTried, intSolutions, False, "", "")
                End If
                ReDim Preserve Solutions(intSolutions - 1)
                Select Case blnClassic
                    Case True
                        StrSolution = strWriteSolution(intGrid:=1)
                    Case False
                        StrSolution = strWriteSolution()
                End Select
                Solutions(intSolutions - 1) = StrSolution

                If intSolutions = intQuit Then
                    '---quit---
                    _vsbackTrack = False
                    intCountSolutions = intSolutions
                    RaiseEvent UI(vsTried, intCountSolutions, True, "", "")
                    Exit Function
                End If
                '---solution found so backtrack
                UndoGuess()
            End If
        End While
    End Function
    Private Function MakeGuess(ByVal intCell As Integer, ByVal intCandidate As Integer) As Boolean
        Dim arrayPeers() As String
        Dim j As Integer
        Dim ptr As Integer
        Dim intBit As Integer
        _vsSolution(intCell - 1) = intCandidate
        _vsCandidateCount(intCell) = -1
        _vsLastGuess(_vsSteps) = intCell
        '----remove from unsolved cells list---
        _vsUnsolvedCells(0).Remove(intCell)
        setCandidates(intCell, intCandidate)
        _vsSteps += 1
        arrayPeers = Split(_vsPeers(intCell), ",")
        '---remove value from peers---
        _vsRemovePeers(intCell) = New List(Of Integer)
        intBit = intGetBit(intCandidate)
        For j = 0 To UBound(arrayPeers)
            ptr = arrayPeers(j)
            If _vsSolution(ptr - 1) = 0 AndAlso (_vsCandidateAvailableBits(ptr) And intBit) Then
                _vsCandidateAvailableBits(ptr) -= intBit
                _vsCandidateCount(ptr) -= 1
                _vsRemovePeers(intCell).Add(ptr)
                If _vsCandidateCount(ptr) = 0 Then Exit Function
            End If
        Next
    End Function
    Private Function UndoGuess(Optional ByRef nextGuess As Integer = 0) As Boolean
        Dim intCell As Integer = 0
        Dim intCandidate As Integer = 0
        Dim blnLoop As Boolean = True
        _vsCandidatePtr(nextGuess) = 1
        _vsSteps -= 1
        If _vsSteps = 0 Then Exit Function
        intCell = _vsLastGuess(_vsSteps)
        intCandidate = _vsSolution(intCell - 1)
        '---restore to unsolved list---
        _vsUnsolvedCells(0).Add(intCell)
        '---sort unsolved cells---
        _vsUnsolvedCells(0).Sort()
        Dim j As Integer
        Dim i As Integer = 1
        Dim c As Integer
        Dim tC As Integer
        Dim intBit As Integer = intGetBit(intCandidate)
        Dim lbit As Integer = 0
        '---restore candidates in this cell---
        If intCell > 0 Then
            If Not (_vsStoreCandidateBits(intCell) And intBit) Then
                _vsStoreCandidateBits(intCell) += intBit
            End If
        End If

        lbit = _vsStoreCandidateBits(intCell)
        _vsCandidateAvailableBits(intCell) = 0
        For c = 1 To 9
            intBit = intGetBit(c)
            If lbit And intBit Then
                _vsCandidateAvailableBits(intCell) += intBit
                tC += 1
            End If
        Next

        nextGuess = intCell
        _vsSolution(intCell - 1) = 0
        _vsCandidateCount(intCell) = tC

        If intCell = 0 Then
            '---no valid solution found---
            Exit Function
        End If

        '---restore value to peers---
        Dim pCell As Integer
        For j = 0 To _vsRemovePeers(intCell).Count - 1
            pCell = _vsRemovePeers(intCell).Item(j)
            _vsCandidateAvailableBits(pCell) += intGetBit(intCandidate)
            _vsCandidateCount(pCell) += 1
        Next
        '---end restore values to peers---
    End Function
    Private Function IntNextCandidate(ByVal intCell As Integer, Optional ByVal blnLookup As Boolean = False) As Integer
        Dim c As Integer
        Dim intBit As Integer
        For c = _vsCandidatePtr(intCell) To 9
            intBit = intGetBit(c)
            If _vsCandidateAvailableBits(intCell) And intBit Then
                IntNextCandidate = c
                If blnLookup = False Then _vsCandidatePtr(intCell) = c + 1
                Exit Function
            End If
        Next
    End Function
    Private Function setCandidates(ByVal intCell As Integer, ByVal intCandidate As Integer) As Integer
        Dim c As Integer
        Dim intBit As Integer
        Dim intCBit As Integer = intGetBit(intCandidate)
        _vsStoreCandidateBits(intCell) = 0
        For c = 1 To 9
            intBit = intGetBit(c)
            If (intBit And _vsCandidateAvailableBits(intCell)) AndAlso (intCBit <> intBit) Then
                _vsStoreCandidateBits(intCell) += intBit
            End If
        Next
    End Function
    Private Function intFindCell() As Integer
        Dim i As Integer
        Dim j As Integer
        Dim ptr As Integer
        Dim ptr2 As Integer
        Dim arrPeers() As String
        Dim intCell As Integer
        Dim intCount As Integer
        Dim intPeerCount As Integer

        For i = 0 To 9
            ptr = Array.IndexOf(_vsCandidateCount, i)
            If ptr > -1 Then
                intFindCell = ptr
                If i = 0 Then
                    intFindCell = 0
                End If
                If i = 1 Then Exit Function

                While ptr2 > -1
                    ptr2 = Array.IndexOf(_vsCandidateCount, i, ptr2)
                    If ptr2 > -1 Then
                        arrPeers = Split(_vsPeers(ptr2), arrDivider)
                        intPeerCount = 0
                        For j = 0 To UBound(arrPeers)
                            'If arrPeers(j) <> "" AndAlso _vsUnsolvedCells(0).IndexOf(arrPeers(j)) > -1 AndAlso (_vsCandidateAvailableBits(arrPeers(j)) And intBit) Then
                            If arrPeers(j) <> "" AndAlso _vsUnsolvedCells(0).IndexOf(arrPeers(j)) > -1 Then
                                intPeerCount += 1
                            End If
                        Next
                        'intPeerCount = 100 * (intPeerCount / arrPeers.Length)
                        If intPeerCount >= intCount Then
                            intCount = intPeerCount
                            intCell = ptr2
                        End If
                        ptr2 += 1
                    End If
                End While
                intFindCell = intCell
                Exit For
            End If
        Next
    End Function
    Private Function strWriteSolution(Optional ByVal intGrid As Integer = 0) As String
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim lstr As String = ""
        Dim s As Integer
        Dim e As Integer
        s = 1
        e = 5
        If intGrid <> 0 Then
            s = intGrid
            e = intGrid
        End If
        For g = s To e
            For i = 1 To 81
                Select Case blnClassic
                    Case True
                        ptr = i
                        If _vsSolution(ptr - 1) = 0 Then
                            lstr += "."
                        Else
                            lstr += CStr(_vsSolution(ptr - 1))
                        End If
                    Case False
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            If _vsSolution(ptr - 1) = 0 Then
                                lstr += "."
                            Else
                                lstr += CStr(_vsSolution(ptr - 1))
                            End If
                        End If
                End Select
            Next
            lstr += vbCrLf
        Next
        strWriteSolution = lstr
    End Function
#End Region
End Class

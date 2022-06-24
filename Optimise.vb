Class clsSudokuOptimise
    Private _Clues(0) As Integer
    Public strGrid As String
    Public strOptimised As String
    Public isOptimised As Boolean
    Public intCluesRemoved As Integer
    Public intType As Integer
    Public strDifficulty As String
    Public intMaxDifficulty As Integer
    Sub New()
        isOptimised = False
        intCluesRemoved = 0
        strOptimised = ""
    End Sub
    Function OptimisePuzzle(Optional ByVal intSymmetry As Integer = 0, Optional ByVal intDifficulty As Integer = -1, Optional ByVal intPrevDifficulty As Integer = 1) As Boolean
        Select Case intType
            Case 1
                ReDim _Clues(81)
            Case 2
                ReDim _Clues(441)
        End Select
        strGrid = Replace(strGrid, vbCrLf, "")
        Dim i As Integer
        Dim g As Integer
        Dim s As Integer
        Dim e As Integer
        Dim r As Integer
        Dim c As Integer = 0
        Dim cellPtr As Integer
        Dim ptr As Integer
        Dim midStr As String = ""
        Dim _restorePtrs(0) As String
        Dim strStartGrid As String = strGrid
        Dim _restoreVals(0) As Integer
        Dim blnRemove As Boolean
        Dim blnUndo As Boolean
        Dim solver As New clsSudokuSolver
        Dim dsolver As New clsSudokuSolver
        s = 1
        e = 5
        If intType = 1 Then e = 1
        For g = s To e
            For i = 1 To 81
                Select Case intType
                    Case 1
                        ptr = i
                        midStr = Mid(strGrid, ptr, 1)
                    Case 2
                        ptr = intSamuraiOffset(i, g)
                        midStr = Mid(strGrid, i + (81 * (g - 1)), 1)
                End Select
                Select Case Asc(midStr)
                    Case 49 To 57
                        '---numeric, so record in array---
                        _Clues(ptr) = CInt(midStr)
                        c += 1
                End Select
            Next
        Next

        '---check there are some clues---
        If c = 0 Then Exit Function

        c = 0
        For ptr = 1 To UBound(_Clues)
            If _Clues(ptr) > 0 Then
                '---store clue/s---
                Dim cs As New Symmetry
                Select Case intType
                    Case 1
                        cs.blnSamurai = False
                    Case 2
                        cs.blnSamurai = True
                End Select
                cs.getCellSymmetry(_restorePtrs, ptr, intSymmetry)
                ReDim _restoreVals(_restorePtrs.Length - 1)
                For r = 1 To UBound(_restorePtrs)
                    If _Clues(_restorePtrs(r)) > 0 Then
                        _restoreVals(r) = _Clues(_restorePtrs(r))
                        blnRemove = True
                    Else
                        blnRemove = False
                        Exit For
                    End If
                Next
                blnUndo = False
                If blnRemove Then
                    '---remove clue/s----
                    For r = 1 To UBound(_restorePtrs)
                        _Clues(_restorePtrs(r)) = 0
                    Next
                    Select Case intType
                        Case 1
                            strGrid = getClues(1)
                        Case 2
                            strGrid = getClues()
                            strGrid = Replace(strGrid, vbCrLf, "")
                    End Select
                    '---check if valid---
                    Select Case intType
                        Case 1
                            solver.blnClassic = True
                        Case 2
                            solver.blnClassic = False
                    End Select
                    solver.strGrid = strGrid
                    solver.intQuit = 2
                    solver.vsSolvers = My.Settings._UniqueSolvers
                    Dim blnUnique As Boolean = solver._vsUnique()
                    If Not blnUnique Then
                        '---not unique, so restore value/s---
                        For r = 1 To UBound(_restorePtrs)
                            cellPtr = _restorePtrs(r)
                            _Clues(cellPtr) = _restoreVals(r)
                        Next
                        strDifficulty = intDifficult2String(intPrevDifficulty)
                        blnUndo = True
                    Else
                        '---result is unique so continue---
                        '---check difficulty---
                        If intDifficulty > -1 Then
                            Select Case intType
                                Case 1
                                    dsolver.blnClassic = True
                                Case 2
                                    dsolver.blnClassic = False
                            End Select
                            dsolver.strDifficulty = ""
                            dsolver.strInputGameSolution = solver.strGameSolution
                            dsolver.vsSolvers = My.Settings._defaultSolvers
                            dsolver.strGrid = strGrid
                            dsolver.intQuit = 2
                            dsolver.strCandidates = ""
                            dsolver._vsUnique()
                            If dsolver.intDifficulty >= intMaxDifficulty Then intMaxDifficulty = dsolver.intDifficulty
                            If dsolver.intDifficulty > intDifficulty Then
                                '---difficulty greater than target so restore value/s---
                                For r = 1 To UBound(_restorePtrs)
                                    cellPtr = _restorePtrs(r)
                                    _Clues(cellPtr) = _restoreVals(r)
                                Next
                                strDifficulty = intDifficult2String(intPrevDifficulty)
                                blnUndo = True
                            End If
                        End If
                        '---end check difficulty---
                    End If
                    If blnUndo = False Then
                        c += _restorePtrs.Length - 1
                        isOptimised = True
                        If dsolver.intDifficulty > intPrevDifficulty Then
                            intPrevDifficulty = dsolver.intDifficulty
                            strDifficulty = intDifficult2String(intPrevDifficulty)
                        End If
                    End If
                    blnUndo = False
                End If
            End If
        Next

        If isOptimised Then
            Select Case intType
                Case 1
                    strOptimised = getClues(1)
                Case 2
                    strOptimised = getClues()
            End Select
        Else
            'Debug.Print("Not optimised: " & strStartGrid)
        End If
        intCluesRemoved = c
    End Function
    Private Function getClues(Optional ByVal intGrid As Integer = 0) As String
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
                Select Case intType
                    Case 1
                        ptr = i
                        If _Clues(ptr) = 0 Then
                            lstr += "."
                        Else
                            lstr += CStr(_Clues(ptr))
                        End If
                    Case 2
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            If _Clues(ptr) = 0 Then
                                lstr += "."
                            Else
                                lstr += CStr(_Clues(ptr))
                            End If
                        End If
                End Select
            Next
            lstr += vbCrLf
        Next
        getClues = lstr
    End Function
End Class

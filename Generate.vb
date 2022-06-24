Module Generate
    Public dirGames As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\Sudoku"
    Private Function fillSamuraiGrids() As String
        Dim i As Integer
        Dim arrGrids(5) As String
        Dim cg As String = fillGrid()
        arrGrids(3) = cg
        For i = 1 To 5
            Select Case i
                Case 1, 2, 4, 5
                    arrGrids(i) = fillGrid(i, arrGrids(3))
            End Select
        Next
        fillSamuraiGrids = ""
        For i = 1 To 5
            fillSamuraiGrids += arrGrids(i) & vbCrLf
        Next
    End Function
    Private Function fillGrid(Optional ByVal intGrid As Integer = 0, Optional ByVal strGrid As String = "") As String
        fillGrid = ""
        Dim c As Integer = 1
        Dim cPtr As Integer
        Dim v As Integer
        Dim rv As String
        Dim strCandidates As String
        Dim arrGame(81) As String
        Dim arrCandidates(81) As String
        Dim blnCollision As Boolean
        Dim arrayPeers() As String
        Dim arrFind() As String
        Dim arrReplace() As String
        Dim arrFixed(8) As String
        Dim ptr As Integer
        Dim ptr2 As Integer
        Dim uPtr As Integer
        Dim j As Integer
        Dim i As Integer
        Dim intPeer As Integer
        Dim sGrid As Integer
        Dim oGrid As Integer
        Dim intMax As Integer = 82
        ReDim arrReplace(0)
        arrReplace(0) = "0"

        Dim avail(0) As List(Of Integer)
        avail(0) = New List(Of Integer)
        For i = 0 To 80
            avail(0).Add(i + 1)
            arrGame(i + 1) = 0
        Next

        If intGrid > 0 Then
            Select Case intGrid
                Case 1
                    sGrid = 1
                Case 2
                    sGrid = 3
                Case 4
                    sGrid = 7
                Case 5
                    sGrid = 9
            End Select
            oGrid = 10 - sGrid
            arrFind = arrGrid(sGrid)
            arrReplace = arrGrid(oGrid)
            For i = 0 To 8
                ptr = arrFind(i)
                ptr2 = arrReplace(i)
                arrGame(ptr2) = Mid(strGrid, ptr, 1)
                avail(0).Remove(ptr2)
            Next
        End If

        'fill grid

        While avail(0).Count > 0
            cPtr = avail(0).Item(0)
            strCandidates = arrCandidates(cPtr)
            If strCandidates = "" Then
                strCandidates = GenerateRandomStr()
            End If

            If arrGame(cPtr) > 0 Then strCandidates = Replace(strCandidates, arrGame(cPtr), "")

            v = 1
            rv = Mid(strCandidates, v, 1)
            'check for conflict
            blnCollision = False
            arrayPeers = arrPeers(cPtr)
            For j = 0 To UBound(arrayPeers)
                intPeer = CInt(arrayPeers(j))
                If arrGame(intPeer) = rv Then
                    blnCollision = True
                    Exit For
                End If
            Next
            strCandidates = Replace(strCandidates, rv, "")
            arrCandidates(cPtr) = strCandidates
            If blnCollision = False Then
                arrGame(cPtr) = rv
                avail(0).Remove(cPtr)
                c += 1
            Else
                If strCandidates = "" Then
                    'no more candidates, so go back
                    arrGame(cPtr) = 0
                    If intGrid = 0 Then
                        avail(0).Add(cPtr - 1)
                    Else
                        uPtr = cPtr - 1
                        Do Until Array.IndexOf(arrReplace, CStr(uPtr)) = -1
                            uPtr -= 1
                        Loop
                        avail(0).Add(uPtr)
                    End If
                    avail(0).Sort()
                    c -= 1
                End If
            End If
        End While

        For i = 1 To 81
            fillGrid += arrGame(i)
        Next

    End Function
    Function GenerateRandomStr(Optional ByVal StrDivider As String = "") As String
        GenerateRandomStr = ""
        Dim avail(0) As List(Of Integer)
        avail(0) = New List(Of Integer)
        Dim i As Integer
        Dim intTest As Integer
        Dim strNumber As String
        For i = 0 To 8
            avail(0).Add(i + 1)
        Next
        While avail(0).Count > 0
            intTest = RandomInt(0, avail(0).Count - 1)
            strNumber = avail(0).Item(intTest)
            If GenerateRandomStr <> "" Then GenerateRandomStr += StrDivider
            GenerateRandomStr += strNumber
            avail(0).Remove(strNumber)
        End While
    End Function
    Function RemoveCellsNoSymmetry(ByVal strGrid As String) As String
        Dim fp As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim p As Integer
        Dim r As Integer
        Dim r2 As Integer
        Dim intRemoved As Integer
        'Dim strGeneratorSeed As String = "0032211000"
        Dim strGeneratorSeed As String = "0122211000"
        Dim randomArr() As String = Split(GenerateRandomStr(arrDivider), arrDivider)
        Dim randomArr2() As String
        Dim ptr As Integer
        Dim arrGame(0) As Integer
        Dim arrPos(0) As Integer
        Dim midStr As String = ""
        strGrid = Replace(strGrid, vbCrLf, "")
        ReDim arrGame(81)

        '---load game into array---
        For p = 1 To 81
            midStr = Mid(strGrid, p, 1)
            ptr = p
            If midStr <> "" AndAlso CInt(midStr) > 0 Then
                arrGame(ptr) = CInt(midStr)
            End If
        Next
        '---finish load game into array---

        For i = 0 To 9
            r = Mid(strGeneratorSeed, i + 1, 1)
            For j = 1 To CInt(r)
                'Debug.Print(randomArr(k) & " will be found " & i & " times so delete " & 9 - i & " instances")
                '---start delete---
                fp = -1
                For p = 1 To 81
                    If arrGame(p) = randomArr(k) Then
                        fp += 1
                        ReDim Preserve arrPos(fp)
                        '---save position---
                        arrPos(fp) = p
                    End If
                Next

                '---randomly remove from array of cell positions 
                intRemoved = 0
                randomArr2 = Split(GenerateRandomStr(arrDivider), arrDivider)
                For r2 = 0 To UBound(randomArr2)
                    If intRemoved >= (9 - i) Then Exit For
                    arrGame(arrPos(randomArr2(r2) - 1)) = 0
                    intRemoved += 1
                Next

                '---end delete---
                k += 1
            Next
        Next

        RemoveCellsNoSymmetry = ""
        For p = 1 To 81
            ptr = p
            If arrGame(ptr) <> "0" Then
                RemoveCellsNoSymmetry += CStr(arrGame(ptr))
            Else
                RemoveCellsNoSymmetry += "."
            End If
        Next

    End Function
    Public Function RemoveCells(ByVal strGrid As String, ByVal intSymmetry As Integer, Optional ByVal blnSamurai As Boolean = False) As String
        Dim cs As New Symmetry
        cs.blnSamurai = blnSamurai
        Dim i As Integer
        Dim j As Integer
        Dim r As Integer
        Dim g As Integer
        Dim ptr As Integer
        Dim arrCells(0) As String
        Dim arrGame(0) As Integer
        Dim intMax As Integer
        Dim intMaxGrid As Integer
        Dim intRemove As Integer
        Dim midStr As String = ""
        If Not blnSamurai Then intMax = 81 Else intMax = 441
        If Not blnSamurai Then intRemove = 49 Else intRemove = 225
        If Not blnSamurai Then intMaxGrid = 1 Else intMaxGrid = 5
        strGrid = Replace(strGrid, vbCrLf, "")
        ReDim arrGame(intMax)

        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case False
                        midStr = Mid(strGrid, i, 1)
                        ptr = i
                    Case True
                        midStr = Mid(strGrid, i + (81 * (g - 1)), 1)
                        ptr = intSamuraiOffset(i, g)
                End Select
                If midStr <> "" AndAlso CInt(midStr) > 0 Then
                    arrGame(ptr) = CInt(midStr)
                End If
            Next
        Next

        i = 0
        While i < intRemove
            ptr = RandomInt(1, 81)
            g = RandomInt(1, 5)
            If blnSamurai Then ptr = intSamuraiOffset(ptr, g)
            r = 0

            cs.getCellSymmetry(arrCells, ptr, intSymmetry)
            For j = 1 To UBound(arrCells)
                If arrGame(arrCells(j)) > 0 Then
                    r += 1
                End If
            Next
            If r = arrCells.Length - 1 Then
                For j = 1 To UBound(arrCells)
                    arrGame(arrCells(j)) = 0
                    i += 1
                Next
            End If
        End While

        RemoveCells = ""
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case False
                        ptr = i
                    Case True
                        ptr = intSamuraiOffset(i, g)
                End Select
                If arrGame(ptr) <> "0" Then
                    RemoveCells += CStr(arrGame(ptr))
                Else
                    RemoveCells += "."
                End If

            Next
            If blnSamurai Then RemoveCells += vbCrLf
        Next
    End Function
    Public Function GenerateNewGrid(ByVal intSymmetry As Integer, ByVal intDifficulty As Integer, ByRef arrPuzzles() As String, ByVal intMaxAttempts As Integer) As Boolean
        GenerateNewGrid = False
        Dim IntGenDifficulty As Integer
        Dim strDifficulty As String = ""
        Dim Solver As New clsSudokuSolver
        Dim dSolver As New clsSudokuSolver
        Dim oGrid As String = ""
        Dim rGrid As String = ""
        Dim sGrid As String = ""
        Dim strSolution As String = ""
        Dim isUnique As Boolean = False
        Dim tries As Integer = 0
        Dim totalTries As Integer
        Dim intRetries As Integer = 10
        Dim intMaxDifficulty As Integer = 0
        Dim intOptDifficulty As Integer = 0

        isUnique = False
        oGrid = fillGrid()
        Solver.blnClassic = True
        Solver.vsSolvers = My.Settings._UniqueSolvers
        Solver.intQuit = 2
        While Not isUnique
            If tries = intRetries Then
                oGrid = fillGrid()
                tries = 0
                totalTries += intRetries
            End If
            If totalTries >= intMaxAttempts Then
                'Debug.Print("Timed out")
                totalTries = 0
                Exit Function
            End If

            If intSymmetry = Symmetry.symmetryTypes.no Then
                rGrid = RemoveCellsNoSymmetry(strGrid:=oGrid)
            Else
                rGrid = RemoveCells(strGrid:=oGrid, intSymmetry:=intSymmetry, blnSamurai:=False)
            End If

            Solver.strGrid = rGrid
            isUnique = Solver._vsUnique()
            strSolution = Solver.strGameSolution
            If Not isUnique Then tries += 1
        End While

        dSolver.blnClassic = True
        dSolver.strInputGameSolution = strSolution
        dSolver.vsSolvers = My.Settings._defaultSolvers
        dSolver.strGrid = rGrid
        dSolver.intQuit = 2
        dSolver.strCandidates = ""
        dSolver._vsUnique()
        IntGenDifficulty = dSolver.intDifficulty
        strDifficulty = LCase(dSolver.strDifficulty)

        Dim opt As New clsSudokuOptimise

        Dim m As Integer = 1
        Dim d() = arrDifficultStr
        intMaxDifficulty = UBound(d) + 1
        sGrid = rGrid

        Dim i As Integer
        For i = 0 To UBound(arrPuzzles)
            arrPuzzles(i) = ""
        Next
        arrPuzzles(IntGenDifficulty) = sGrid
        arrPuzzles(0) = strSolution

        While m <= intMaxDifficulty
            opt.strGrid = sGrid
            opt.intType = 1
            opt.isOptimised = False
            opt.strDifficulty = ""
            opt.intMaxDifficulty = 1
            opt.OptimisePuzzle(intSymmetry:=intSymmetry, intDifficulty:=m, intPrevDifficulty:=IntGenDifficulty)
            If opt.isOptimised Then
                rGrid = opt.strOptimised
                strDifficulty = LCase(opt.strDifficulty)
                intOptDifficulty = strDifficult2Int(strDifficulty)
                If intOptDifficulty = m Then
                    'Debug.Print("Optimised to: " & intDifficult2String(m) & ": " & Replace(rGrid, vbCrLf, ""))
                    arrPuzzles(m) = Replace(rGrid, vbCrLf, "")
                Else
                    'Debug.Print("Could not optimise to: " & intDifficult2String(m))
                    arrPuzzles(m) = ""
                End If
                '---set maximum difficulty---'
                intMaxDifficulty = opt.intMaxDifficulty
            Else
                'Debug.Print("Could not optimise to: " & intDifficult2String(m))
                arrPuzzles(m) = ""
                If m = 1 Then
                    intMaxDifficulty = 1
                    arrPuzzles(m) = Replace(rGrid, vbCrLf, "")
                End If
            End If
            m += 1
        End While

        For i = 1 To UBound(arrPuzzles)
            If Not My.Settings._GenClassicDifficulty And intGetBit(i) Then
                arrPuzzles(IntGenDifficulty) = ""
            End If
        Next

        GenerateNewGrid = True

    End Function
End Module

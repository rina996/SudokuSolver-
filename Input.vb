Module Input
    Function frmLoad(ByVal strGrid As String, Optional ByRef ErrMsg As String = "", Optional ByVal blnSamurai As Boolean = False) As Boolean
        Dim g As Integer ' counter
        Dim i As Integer ' counter
        Dim p As Integer
        Dim c As Integer ' count total input
        Dim ptr As Integer
        Dim u(0) As List(Of Integer)
        u(0) = New List(Of Integer)
        Dim intMaxGrid As Integer = 1
        Dim intMaxTotal As Integer = 81
        Dim intCountTotal As Integer = 81
        If blnSamurai Then intMaxGrid = 5
        If blnSamurai Then intMaxTotal = 441
        If blnSamurai Then intCountTotal = 405
        Dim midStr As String = ""
        Dim intCell As Integer = 0
        Dim intValue As Integer = 0
        Dim intBitArr(intMaxTotal) As Integer
        Dim PeersArr(0) As String
        Dim SC As SudokuCell
        Dim ec As Integer
        Dim tColour As Color = Color.FromName(My.Settings._clueColour)

        If Len(strGrid) > intCountTotal Then
            errMsg = Microsoft.VisualBasic.Right(strGrid, Len(strGrid) - intCountTotal)
        End If

        '---fill bit array---
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        intCell = intSamuraiOffset(i, g)
                    Case False
                        intCell = i
                End Select
                If Not blnSamurai Or (blnSamurai AndAlso Not blnIgnoreSamurai(intCell)) Then
                    intBitArr(intCell) = 511
                End If
            Next
        Next
        '---end fill bit array---

        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        intCell = intSamuraiOffset(i, g)
                        midStr = Mid(strGrid, i + (81 * (g - 1)), 1)
                    Case False
                        intCell = i
                        midStr = Mid(strGrid, i, 1)
                End Select
                If Not blnSamurai Or (blnSamurai AndAlso Not blnIgnoreSamurai(intCell)) Then
                    Select Case Asc(midStr)
                        Case 46, 48
                            '---blank (0 or full stop)---
                            u(0).Add(intCell)
                            c += 1
                        Case 49 To 57
                            '---numeric 1 to 9---
                            intValue = CInt(midStr)
                            c += 1
                            Select Case blnSamurai
                                Case True
                                    PeersArr = ArrSamuraiPeers(intCell)
                                Case False
                                    PeersArr = arrPeers(intCell)
                            End Select
                            For p = 0 To UBound(PeersArr)
                                ptr = PeersArr(p)
                                If blnMatchBit(intBitArr(ptr), intValue) Then
                                    intBitArr(ptr) -= intGetBit(intValue)
                                End If
                            Next
                        Case Else
                            '---invalid input - ignore---
                    End Select
                End If
            Next
        Next

        If c <> intCountTotal Then
            frmGame.FeedbackBox.Text = "Invalid input"
            Return False
            Exit Function
        End If

        '---draw clues---
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        intCell = intSamuraiOffset(i, g)
                        midStr = Mid(strGrid, i + (81 * (g - 1)), 1)
                    Case False
                        intCell = i
                        midStr = Mid(strGrid, i, 1)
                End Select
                If Not blnSamurai Or (blnSamurai AndAlso Not blnIgnoreSamurai(intCell)) Then
                    SC = getCtrl("SudokuCell" & intCell)
                    SC.sc_BackgroundColour = transparentColor(Color.FromName(My.Settings._bgColour), intBGTransparency)
                    SC.sc_BorderColour = transparentColor(Color.FromName(My.Settings._borderColour), intBorderTransparency)
                    SC.sc_TextColour = tColour
                    Select Case Asc(midStr)
                        Case 46, 48
                            '---blank (0 or full stop)---
                            SC.sc_IntSolution = 0
                            SC.sc_IntCandidate = 0
                        Case 49 To 57
                            '---numeric 1 to 9---
                            intValue = CInt(midStr)
                            SC.sc_IntSolution = intValue
                            SC.sc_AlertGraphic = False
                        Case Else
                            '---invalid input - ignore---
                    End Select
                End If
            Next
        Next

        '---fill candidates---
        Dim errArr(0) As Integer
        For i = 0 To u(0).Count - 1
            ptr = u(0).Item(i)
            SC = getCtrl("SudokuCell" & ptr)
            SC.sc_TextColour = tColour
            If intBitArr(ptr) > 0 Then
                SC.sc_IntCandidate = intBitArr(ptr)
                SC.sc_AlertGraphic = False
            Else
                If Array.IndexOf(errArr, ptr) = -1 Then
                    ReDim Preserve errArr(ec)
                    errArr(ec) = ptr
                    ec += 1
                    SC.sc_AlertGraphic = True
                End If
            End If
        Next

        If ec > 0 Then
            Select Case ec
                Case 1
                    ErrMsg = "Error - one cell has no candidates available."
                Case Else
                    ErrMsg = "Error - " & ec & "  cells have no candidates available."
            End Select
        End If

        '---show error msg---
        frmGame.FeedbackBox.Text = errMsg
        Return True
    End Function
    Function frmClear(Optional ByVal blnSamurai As Boolean = False) As Boolean
        Dim g As Integer ' counter
        Dim i As Integer ' counter
        Dim tColour As Color = Color.FromName(My.Settings._clueColour)
        Dim intMaxGrid As Integer = 1
        If blnSamurai Then intMaxGrid = 5
        Dim intCell As Integer
        For g = 1 To intMaxGrid
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        intCell = intSamuraiOffset(i, g)
                    Case False
                        intCell = i
                End Select
                If Not blnSamurai Or (blnSamurai AndAlso Not blnIgnoreSamurai(intCell)) Then
                    Try
                        With getCtrl("SudokuCell" & intCell)
                            .sc_TextColour = tColour
                            .sc_IntSolution = 0
                            .sc_IntCandidate = 0
                            .sc_AlertGraphic = False
                        End With
                    Catch
                    End Try
                End If
            Next
        Next
        restoreDefaults()
        frmGame.FeedbackBox.Text = ""
    End Function
    Function blnMultiInput(ByVal strGrid As String, Optional ByVal blnSamurai As Boolean = False) As Boolean
        frmGame.ListMultiGame(0).Clear()
        Dim i As Integer
        Dim sg As Integer = 0
        Dim gc As Integer = 0
        Dim blnValid As Boolean = True
        Dim errMsg As String = ""
        If blnSamurai Then
            Dim tempArr() As String
            Dim tempGrid As String = ""
            tempArr = Split(strGrid, vbCrLf)
            For i = 0 To UBound(tempArr)
                tempGrid += tempArr(i)
                sg += 1
                If blnValid Then blnValid = blnValidGrid(strGrid:=tempArr(i), strErrMsg:=errMsg, intSamuraiGrid:=sg)
                If (i + 1) Mod 5 = 0 Then
                    If blnValid And Not blnValidOverlaps(tempGrid) Then
                        blnValid = False
                        errMsg = "Input is invalid as values for overlapping regions don't match."
                    End If
                    If Not blnValid Then
                        frmGame.ListMultiGame(0).Add(tempGrid & errMsg)
                    Else
                        frmGame.ListMultiGame(0).Add(tempGrid)
                    End If
                    tempGrid = ""
                    blnValid = True
                    sg = 0
                End If
            Next
        Else
            frmGame.ListMultiGame(0).AddRange(Split(strGrid, vbCrLf))
            For i = 0 To frmGame.ListMultiGame(0).Count - 1
                blnValid = blnValidGrid(strGrid:=frmGame.ListMultiGame(0).Item(i), strErrMsg:=errMsg)
                If Not blnValid Then
                    frmGame.ListMultiGame(0).Item(i) = frmGame.ListMultiGame(0).Item(i) & errMsg
                End If
            Next
        End If
        gc = frmGame.ListMultiGame(0).Count
        If gc > 1 Then
            frmGame.ToolStripMultiGameLabel.Visible = True
            frmGame.ToolStripMultiGameLabel.Text = "Puzzle 1 of " & gc
            frmGame.intMultiGame = 1
            Return True
        Else
            Return False
        End If
    End Function
    Function blnValidOverlaps(ByVal strGrid As String) As Boolean
        Dim g As Integer
        Dim i As Integer
        Dim intCell As Integer
        Dim arrOverlaps() As String = arrSamuraiOverlap()
        Dim arrCheck(441) As String
        Dim strMid As String
        For g = 1 To 5
            For i = 1 To 81
                intCell = intSamuraiOffset(i, g)
                If Not blnIgnoreSamurai(intCell) Then
                    If Array.IndexOf(arrOverlaps, CStr(intCell)) > -1 Then
                        strMid = Mid(strGrid, i + (81 * (g - 1)), 1)
                        If arrCheck(intCell) <> "" Then
                            If arrCheck(intCell) <> strMid Then
                                '---overlaps between samurai grids don't match---
                                Return False
                            End If
                        Else
                            arrCheck(intCell) = strMid
                        End If
                    End If
                End If
            Next
        Next
        Return True
    End Function
    Function blnValidGrid(ByVal strGrid As String, ByRef strErrMsg As String, Optional ByVal intSamuraiGrid As Integer = 0) As Boolean
        strErrMsg = ""
        Dim rowBitarr(8) As Integer
        Dim colBitarr(8) As Integer
        Dim gridBitarr(8) As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim p As Integer
        Dim rcg As Integer
        Dim intValue As Integer
        Dim strValue As String
        Dim ascValue As Integer
        Dim intBit As Integer
        Dim intBitCandidates As Integer
        Dim intBitFilled As Integer
        Dim intBitArr(81) As Integer
        Dim PeersArr(0) As String
        Dim rcgArr(0) As String
        Dim strRCG As String = ""
        Dim blnValid As Boolean = True
        Dim ptr As Integer
        Dim intCell As Integer

        '---fill bit array---
        For i = 1 To 81
            intCell = i
            intBitArr(intCell) = 511
        Next
        '---end fill bit array---

        For i = 1 To 81
            intCell = i
            strValue = Mid(strGrid, i, 1)
            Select Case Asc(strValue)
                Case 46, 48
                    '---blank (0 or full stop)---
                Case 49 To 57
                    '---numeric 1 to 9---
                    intValue = CInt(strValue)
                    PeersArr = arrPeers(intCell)
                    For p = 0 To UBound(PeersArr)
                        ptr = PeersArr(p)
                        If blnMatchBit(intBitArr(ptr), intValue) Then
                            intBitArr(ptr) -= intGetBit(intValue)
                            If intBitArr(ptr) <= 0 Then
                                '---cell has no candidates---
                                strErrMsg = "puzzle has a cell where no candidates can be placed."
                                If intSamuraiGrid > 0 Then
                                    strErrMsg = "Error - in samurai grid " & intSamuraiGrid & ", " & strErrMsg
                                Else
                                    strErrMsg = "Error - " & strErrMsg
                                End If
                                blnValidGrid = False
                                Exit Function
                            End If
                        End If
                    Next
                Case Else
                    '---invalid input - ignore---
            End Select
        Next

        Dim scArr(8) As Integer
        For i = 1 To 3
            For j = 1 To 9
                ReDim scArr(8)
                Select Case i
                    Case 1
                        rcgArr = arrRow(j)
                        strRCG = "row"
                    Case 2
                        rcgArr = arrCol(j)
                        strRCG = "column"
                    Case 3
                        rcgArr = arrGrid(j)
                        strRCG = "subgrid"
                End Select
                intBitCandidates = 0
                intBitFilled = 0
                For k = 0 To 8
                    ptr = rcgArr(k)
                    strValue = Mid(strGrid, ptr, 1)
                    Select Case Asc(strValue)
                        Case 46, 48
                            '---blank (0 or full stop)---
                            For p = 1 To 9
                                intBit = intGetBit(p)
                                If intBit And intBitArr(ptr) Then
                                    If Not blnMatchBit(intBitCandidates, p) Then
                                        intBitCandidates += intBit
                                    End If
                                End If
                            Next
                            If blnSingleBit(intBitArr(ptr)) Then
                                If Array.IndexOf(scArr, intBitArr(ptr)) < 0 Then
                                    scArr(k) = intBitArr(ptr)
                                Else
                                    '---duplicate single candidates---
                                    strErrMsg = strRCG & " " & j & " has more than one cell where only a single candidate (" & bitmask2Str(intBitArr(ptr)) & ") could be placed."
                                    If intSamuraiGrid > 0 Then
                                        strErrMsg = "Error - in grid " & intSamuraiGrid & ", " & strErrMsg
                                    Else
                                        strErrMsg = "Error - " & strErrMsg
                                    End If
                                blnValidGrid = False
                                Exit Function
                            End If
                            End If
                        Case 49 To 57
                            '---numeric---
                            intValue = CInt(strValue)
                            intBit = intGetBit(intValue)
                            If Not blnMatchBit(intBitFilled, CInt(strValue)) Then
                                intBitFilled += intBit
                            End If
                    End Select
                Next
                If intBitFilled + intBitCandidates <> 511 Then
                    strErrMsg = strRCG & " " & j & " does not have sufficient candidates to result in a valid solution. (There are no cells where candidate " & bitmask2Str(intBit:=511 - (intBitFilled + intBitCandidates), strDivider:=" or ") & " can be placed in this " & strRCG & ")."
                    If intSamuraiGrid > 0 Then
                        strErrMsg = "Error - in grid " & intSamuraiGrid & ", " & strErrMsg
                    Else
                        strErrMsg = "Error - " & strErrMsg
                    End If
                    blnValidGrid = False
                    Exit Function
                End If
            Next
        Next

        '---check for duplicates---
        For i = 1 To Len(strGrid)
            strValue = Mid(strGrid, i, 1)
            ascValue = Asc(strValue)
            Select Case ascValue
                Case 46, 48
                    '---blank (0 or full stop)---
                Case 49 To 57
                    '---1 to 9---
                    intValue = CInt(strValue)
                    intBit = intGetBit(intValue)
                    For j = 1 To 3
                        Select Case j
                            Case 1
                                rcg = intFindRow(i)
                                strRCG = "row"
                                If intBit And rowBitarr(rcg - 1) Then
                                    blnValid = False
                                Else
                                    rowBitarr(rcg - 1) += intBit
                                End If
                            Case 2
                                rcg = intFindCol(i)
                                strRCG = "column"
                                If intBit And colBitarr(rcg - 1) Then
                                    blnValid = False
                                Else
                                    colBitarr(rcg - 1) += intBit
                                End If
                            Case 3
                                rcg = intFindGrid(i)
                                strRCG = "subgrid"
                                If intBit And gridBitarr(rcg - 1) Then
                                    blnValid = False
                                Else
                                    gridBitarr(rcg - 1) += intBit
                                End If
                        End Select
                        If blnValid = False Then
                            strErrMsg = strRCG & " " & rcg & " contains a duplicate value (" & intValue & ")"
                            If intSamuraiGrid > 0 Then
                                strErrMsg = "Error - In grid " & intSamuraiGrid & ", " & strErrMsg
                            Else
                                strErrMsg = "Error - " & strErrMsg
                            End If
                            Exit For
                        End If
                    Next
            End Select
            If blnValid = False Then Exit For
        Next

        If Not blnValid Then
            '---input has duplicates---
            blnValidGrid = False
            Exit Function
        End If

        Return True
    End Function
    Function readinInput(ByVal strInput As String) As String
        Dim strTemp As String = ""
        Dim strMid As String
        Dim strGames As String = ""
        Dim i As Integer ' counter
        Dim c As Integer ' counter
        For i = 1 To Len(strInput)
            strMid = Mid(strInput, i, 1)
            Select Case Asc(strMid)
                Case 46, 48
                    '---blank (0 or full stop)---
                    c += 1
                    strTemp += "."
                Case 49 To 57
                    '---numeric 1 to 9---
                    strTemp += Chr(Asc(strMid))
                    c += 1
                Case Else
                    '---ignore---
            End Select
            If c = 81 Then
                If strGames = "" Then
                    strGames = strTemp
                Else
                    strGames += vbCrLf & strTemp
                End If
                strTemp = ""
                c = 0
            End If
        Next
        readinInput = strGames
    End Function
    Sub restoreDefaults(Optional ByVal blnDeleteMulti As Boolean = True)
        If blnDeleteMulti Then
            frmGame.ToolStripMultiGameLabel.Visible = False
            frmGame.ToolStripMultiGameLabel.Text = ""
            frmGame.intMultiGame = 1
            frmGame.intActiveCell = 0
            frmGame.strPuzzleSolution = ""
        End If
        frmGame.setCheckBoxes(intEnabled:=0, intChecked:=0)
        frmGame.intMultiCell(0).Clear()
        frmGame.Moves.Clear()
        frmGame.RedoMoves.Clear()
        frmGame.intHintCount = 0
        frmGame.strHintDetails = ""
        frmGame.blnSetup = True
        frmGame.intAllowSetupInput = -1
        frmGame.SetupLabel.Visible = True
    End Sub
End Module

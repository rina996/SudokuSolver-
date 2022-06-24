Class Symmetry
    Private _grid(0) As Integer
    Public blnSamurai As Boolean = False
    Public strGrid As String = ""
    Public Enum symmetryTypes
        no = 0
        full = 1
        vertical = 2
        horizontal = 3
        diagonal = 4
    End Enum
    Private Function RCtoInt(ByVal r As Integer, ByVal c As Integer) As Integer
        Dim cLength As Integer
        If blnSamurai Then
            '--samurai---
            cLength = 21
        Else
            '--classic--
            cLength = 9
        End If
        RCtoInt = (cLength * (r - 1)) + c
    End Function
    Public Function checkSymmetry() As String

        Select Case blnSamurai
            Case False
                ReDim _grid(81)
            Case True
                ReDim _grid(441)
        End Select

        Dim s As Integer
        Dim e As Integer
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim midStr As String = ""
        s = 1
        e = 5
        If Not blnSamurai Then
            e = 1
        End If
        For g = s To e
            For i = 1 To 81
                Select Case blnSamurai
                    Case True
                        ptr = intSamuraiOffset(i, g)
                        midStr = Mid(strGrid, i + (81 * (g - 1)), 1)
                    Case False
                        ptr = i
                        midStr = Mid(strGrid, ptr, 1)
                End Select
                Select Case Asc(midStr)
                    Case 49 To 57
                        '---numeric, so record in array---
                        _grid(ptr) = CInt(midStr)
                End Select
            Next
        Next

        Dim a() As String
        a = [Enum].GetNames(GetType(symmetryTypes))
        For i = 0 To a.Length - 1
            If scanSymmetry(i) Then
                checkSymmetry = a(i)
                Exit Function
            End If
        Next
        checkSymmetry = a(0)
    End Function
    Public Function scanSymmetry(ByVal intType As Integer) As Boolean
        Dim g As Integer
        Dim i As Integer
        Dim ptr As Integer
        Dim s As Integer
        Dim e As Integer
        Dim v As String
        Dim arrCells(0) As String
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
                        v = _grid(ptr)
                        If v > 0 Then
                            CheckCellSymmetry(arrCells:=arrCells, intCell:=ptr)
                            If Not testSymmetry(arrCells, intType) Then Exit Function
                        End If
                    Case True
                        '---Samurai---
                        ptr = intSamuraiOffset(i, g)
                        If Not blnIgnoreSamurai(ptr) Then
                            v = _grid(ptr)
                            If v > 0 Then
                                CheckCellSymmetry(arrCells:=arrCells, intCell:=i)
                                If Not testSymmetry(arrCells, intType) Then Exit Function
                            End If
                        End If
                End Select
            Next
        Next
        scanSymmetry = True
    End Function
    Private Sub CheckCellSymmetry(ByRef arrCells() As String, ByVal intCell As Integer)
        Dim intRow As Integer
        Dim intCol As Integer

        ReDim arrCells(4)

        intRow = intFindRow(intCell)
        intCol = intFindCol(intCell)

        If Not blnSamurai Then
            arrCells(1) = RCtoInt(intRow, intCol)
            arrCells(2) = RCtoInt(intRow, 10 - intCol)
            arrCells(3) = RCtoInt(10 - intRow, intCol)
            arrCells(4) = RCtoInt(10 - intRow, 10 - intCol)
        Else
            arrCells(1) = RCtoInt(intRow, intCol)
            arrCells(2) = RCtoInt(intRow, 22 - intCol)
            arrCells(3) = RCtoInt(22 - intRow, intCol)
            arrCells(4) = RCtoInt(22 - intRow, 22 - intCol)
        End If
    End Sub
    Public Sub getCellSymmetry(ByRef getCells() As String, ByVal intCell As Integer, ByVal intType As Integer)
        Dim intRow As Integer
        Dim intCol As Integer
        intRow = intFindRow(intCell)
        intCol = intFindCol(intCell)

        If blnSamurai Then
            If blnIgnoreSamurai(intCell) Then Exit Sub
            intRow = intFindSamuraiRow(intCell)
            intCol = intFindSamuraiCol(intCell)
        End If

        Dim arrCells(4)

        If Not blnSamurai Then
            arrCells(1) = RCtoInt(intRow, intCol)
            arrCells(2) = RCtoInt(intRow, 10 - intCol)
            arrCells(3) = RCtoInt(10 - intRow, intCol)
            arrCells(4) = RCtoInt(10 - intRow, 10 - intCol)
        Else
            arrCells(1) = RCtoInt(intRow, intCol)
            arrCells(2) = RCtoInt(intRow, 22 - intCol)
            arrCells(3) = RCtoInt(22 - intRow, intCol)
            arrCells(4) = RCtoInt(22 - intRow, 22 - intCol)
        End If

        Select Case intType

            Case symmetryTypes.full
                ReDim getCells(4)
                getCells(1) = arrCells(1)
                getCells(2) = arrCells(2)
                getCells(3) = arrCells(3)
                getCells(4) = arrCells(4)

            Case symmetryTypes.vertical
                ReDim getCells(2)
                getCells(1) = arrCells(1)
                getCells(2) = arrCells(2)

            Case symmetryTypes.horizontal
                ReDim getCells(2)
                getCells(1) = arrCells(1)
                getCells(2) = arrCells(3)

            Case symmetryTypes.diagonal
                ReDim getCells(2)
                getCells(1) = arrCells(1)
                getCells(2) = arrCells(4)

            Case Else
                '---no symmetry---
                ReDim getCells(1)
                getCells(1) = arrCells(1)

        End Select
    End Sub
    Private Function testSymmetry(ByVal arrCells() As String, ByVal intType As Integer) As Boolean
        Dim i As Integer
        Dim tc(4) As Integer
        For i = 1 To 4
            tc(i) = _grid(arrCells(i))
        Next

        Select Case intType
            Case symmetryTypes.full
                If tc(1) > 0 AndAlso tc(2) > 0 AndAlso tc(3) > 0 AndAlso tc(4) > 0 Then
                    '---full symmetrical---
                    testSymmetry = True
                End If

            Case symmetryTypes.vertical
                If tc(1) > 0 AndAlso tc(2) > 0 Then
                    '---vertical symmetrical---
                    testSymmetry = True
                    Exit Function
                End If

            Case symmetryTypes.horizontal
                If tc(1) > 0 AndAlso tc(3) > 0 Then
                    '---horizontal symmetrical---
                    testSymmetry = True
                    Exit Function
                End If

            Case symmetryTypes.diagonal
                If tc(1) > 0 AndAlso tc(4) > 0 Then
                    '---diagonal symmetrical---
                    testSymmetry = True
                    Exit Function
                End If

        End Select
    End Function
End Class

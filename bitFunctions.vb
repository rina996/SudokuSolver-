Module bitFunctions
    Public Function bitmask2Str(ByVal intBit As Integer, Optional ByVal strDivider As String = "") As String
        '---turn bitmask into text---
        bitmask2Str = ""
        Dim tBit As Integer
        Dim i As Integer
        For i = 0 To 8
            tBit = intGetBit(i + 1)
            If tBit > intBit Then Exit For
            If tBit And intBit Then
                If bitmask2Str = "" Then
                    bitmask2Str += CStr(i + 1)
                Else
                    bitmask2Str += strDivider & CStr(i + 1)
                End If
            End If
        Next
    End Function
    Public Function bitmask2Arr(ByVal intBit As Integer) As Integer()
        '---turn bitmask into array---
        Dim rArr(0) As Integer
        Dim i As Integer
        Dim j As Integer = -1
        Dim tBit As Integer
        For i = 0 To 8
            tBit = intGetBit(i + 1)
            If tBit > intBit Then Exit For
            If tBit And intBit Then
                j += 1
                ReDim Preserve rArr(j)
                rArr(j) = (i + 1)
            End If
        Next
        bitmask2Arr = rArr
    End Function
    Public Function blnMatchBit(ByVal intLongBit As Integer, ByVal intCandidate As Integer) As Boolean
        Select Case intCandidate
            Case 1 To 9
            Case Else
                Exit Function
        End Select
        If intGetBit(intCandidate) And intLongBit Then
            blnMatchBit = True
        End If
    End Function
    Public Function intCountBits(ByVal intLongBit) As Integer
        Dim i As Integer
        Dim tBit As Integer
        For i = 1 To 9
            tBit = intGetBit(i)
            If tBit > intLongBit Then Exit For
            If tBit And intLongBit Then
                intCountBits += 1
            End If
        Next
    End Function
    Public Function intGetBit(ByVal intCandidate As Integer) As Integer
        Select Case intCandidate
            Case 1
                intGetBit = 1
            Case 2
                intGetBit = 2
            Case 3
                intGetBit = 4
            Case 4
                intGetBit = 8
            Case 5
                intGetBit = 16
            Case 6
                intGetBit = 32
            Case 7
                intGetBit = 64
            Case 8
                intGetBit = 128
            Case 9
                intGetBit = 256
            Case Else
                '---error---
        End Select
    End Function
    Public Function intReverseBit(ByVal intBit As Integer) As Integer
        '---turns bit into integer---
        Select Case intBit
            Case 1
                intReverseBit = 1
            Case 2
                intReverseBit = 2
            Case 4
                intReverseBit = 3
            Case 8
                intReverseBit = 4
            Case 16
                intReverseBit = 5
            Case 32
                intReverseBit = 6
            Case 64
                intReverseBit = 7
            Case 128
                intReverseBit = 8
            Case 256
                intReverseBit = 9
        End Select
    End Function
    Public Function blnSingleBit(ByVal intBit As Integer) As Boolean
        '---returns true if input is a single bit
        blnSingleBit = True
        Select Case intBit
            Case 1
            Case 2
            Case 4
            Case 8
            Case 16
            Case 32
            Case 64
            Case 128
            Case 256
            Case Else
                blnSingleBit = False
        End Select
    End Function
End Module

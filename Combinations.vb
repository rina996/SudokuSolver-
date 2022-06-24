Module Combinations
    Public Function generateCombinationArray(ByRef inputArray() As String, ByRef outputArray() As String, ByVal intMax As Integer) As Boolean
        Dim i As Integer
        Dim b As Integer
        Dim bArray() As String
        Dim bitSum As Integer
        IterateCombinations(inputArray, inputArray, outputArray)
        For i = 3 To intMax
            IterateCombinations(inputArray, outputArray, outputArray)
        Next
        '---turn combination into bitmask---
        For i = 0 To UBound(outputArray)
            bArray = Split(outputArray(i), arrDivider)
            For b = 0 To UBound(bArray)
                bitSum += 2 ^ (bArray(b) - 1)
            Next
            outputArray(i) = bitSum
            bitSum = 0
        Next
    End Function
    Public Function IterateCombinations(ByRef originalArr() As String, ByRef inputArr() As String, ByRef fArr() As String) As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim intCount As Integer
        intCount = -1
        Dim intRight As String
        Dim nArr() As String
        Dim nArr2() As String
        Dim tempArr() As String
        nArr = inputArr
        nArr2 = originalArr
        For i = 0 To UBound(nArr)
            tempArr = Split(nArr(i), arrDivider)
            intRight = tempArr(UBound(tempArr))
            For j = 0 To UBound(nArr2)
                If CInt(nArr2(j)) > CInt(intRight) Then
                    intCount = intCount + 1
                    ReDim Preserve fArr(intCount)
                    fArr(intCount) = nArr(i) & arrDivider & nArr2(j)
                End If
            Next
        Next
    End Function
End Module

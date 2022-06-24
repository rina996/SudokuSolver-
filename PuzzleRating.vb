Module PuzzleRating
    Public arrDifficultStr() As String = Split("Easy,Moderate,Hard,Very Hard", ",")
    Public Function intDifficult2String(ByVal intDifficulty As Integer) As String
        intDifficult2String = arrDifficultStr(intDifficulty - 1)
    End Function
    Public Function strDifficult2Int(ByVal strDifficulty As String) As Integer
        If Array.IndexOf(arrDifficultStr, StrConv(strDifficulty, VbStrConv.ProperCase)) > -1 Then
            strDifficult2Int = Array.IndexOf(arrDifficultStr, StrConv(strDifficulty, VbStrConv.ProperCase)) + 1
        End If
    End Function
End Module

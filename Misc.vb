Module Misc
    Public Enum filetype
        txt = 1
        sdm = 2
    End Enum
    Public Function blnValidExtension(ByVal fileEx As String) As Boolean
        If Len(fileEx) <> 3 Then Exit Function
        Dim a() As String
        a = [Enum].GetNames(GetType(filetype))
        If Array.IndexOf(a, fileEx) > -1 Then
            Return True
        End If
    End Function
    Public Function getCtrl(ByVal strCtrl As String) As SudokuCell
        getCtrl = Nothing
        Dim sc() As Control = frmGame.Controls.Find(strCtrl, True)
        If sc.Length > 0 Then
            getCtrl = CType(sc(0), SudokuCell)
        End If
    End Function
    Public Function getCB(ByVal strCtrl As String) As CheckBox
        getCB = Nothing
        Dim cb() As Control = frmGame.Controls.Find(strCtrl, True)
        If cb.Length > 0 Then
            getCB = CType(cb(0), CheckBox)
        End If
    End Function
    Public Function transparentColor(ByVal inputColor As Color, ByVal intTransparent As Double) As Color
        transparentColor = Color.FromArgb(intTransparent * 255, inputColor.R, inputColor.G, inputColor.B)
    End Function
    Public Function RandomInt(ByVal lowerBound As Integer, ByVal upperBound As Integer) As Integer
        '---Return a random integer---
        If upperBound - lowerBound <= 0 Then
            RandomInt = lowerBound
            Exit Function
        End If
        Randomize()
        RandomInt = Int(Rnd() * (upperBound - lowerBound + 1)) + lowerBound
    End Function
    Public Function ShowMessageLabel(ByVal strMsg As String, Optional ByVal strImage As String = "") As Boolean
        Select Case strImage
            Case "info"
                frmGame.StatusLabel.Image = My.Resources.info
            Case "error"
                frmGame.StatusLabel.Image = My.Resources.err
            Case "alert"
                frmGame.StatusLabel.Image = My.Resources.alert
            Case Else
                frmGame.StatusLabel.Image = Nothing
        End Select
        If frmGame.StatusLabel.Text = "" Or frmGame.MsgTimer.Enabled = False Then
            frmGame.MsgTimer.Start()
            frmGame.StatusLabel.Text = strMsg
            frmGame.StatusLabel.ToolTipText = strMsg
        Else
            If frmGame.StatusLabel.Text <> strMsg Then
                frmGame.MsgTimer.Stop()
                frmGame.MsgTimer.Start()
                frmGame.StatusLabel.Text = strMsg
                frmGame.StatusLabel.ToolTipText = strMsg
            End If
        End If
    End Function
    Public Function strMulti(ByVal intValue As Long, ByVal strText As String, ByVal blnPrefix As Boolean) As String
        '---displays string to reflect entered input---
        '---"1" and "cat" will return " 1 cat "
        '---"3" and "cat" will return " 3 cats"
        strMulti = ""
        If blnPrefix Then
            Select Case intValue
                Case 1
                    strMulti = "is "
                Case Else
                    strMulti = "are "
            End Select
        End If
        strMulti += intValue & " " & strText
        Select Case intValue
            Case 1
            Case Else
                strMulti += "s"
        End Select
    End Function
    Public Function PlaySound(ByVal soundFile As Object) As Boolean
        '---plays a sound if sounds are enabled---
        If My.Settings._PlaySounds Then My.Computer.Audio.Play(soundFile, AudioPlayMode.Background)
    End Function
    Public Function strSecondsMilli(ByVal intValue As Long, Optional ByVal intRound As Integer = 2) As String
        Dim intSeconds As Double
        Select Case intValue
            Case Is < 1000
                Select Case intValue
                    Case 1
                        strSecondsMilli = intValue & " millisecond"
                    Case Else
                        strSecondsMilli = intValue & " milliseconds"
                End Select
            Case Else
                intSeconds = Math.Round(intValue / 1000, intRound)
                Select Case intSeconds
                    Case 1
                        strSecondsMilli = intSeconds & " second"
                    Case Else
                        strSecondsMilli = intSeconds & " seconds"
                End Select
        End Select
    End Function
    Public Function intSymmetrytoString(ByVal intSymmetry As Integer) As String
        Dim sym() As String
        sym = [Enum].GetNames(GetType(Symmetry.symmetryTypes))
        intSymmetrytoString = sym(intSymmetry)
    End Function
    Public Function strFirstLetters(ByVal strText) As String
        strFirstLetters = ""
        Dim w As Integer
        Dim arrWords() As String = Split(strText, " ")
        For w = 0 To UBound(arrWords)
            strFirstLetters += UCase(Left(arrWords(w), 1))
        Next
    End Function
    Public Function CountPuzzles(ByVal strFilter) As Integer
        Try
            CountPuzzles = My.Computer.FileSystem.GetDirectoryInfo(dirGames).GetFiles(strFilter).Length()
        Catch
            CountPuzzles = -1
        End Try
    End Function
End Module

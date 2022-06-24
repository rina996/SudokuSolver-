Option Strict Off
Imports System
Imports System.ComponentModel
Imports System.Threading
Imports Microsoft.Win32
' TODO: Code samurai puzzle generation
' TODO: Add support for jigsaw puzzles?
' TODO: Support for custom sounds?
Public Class frmInitial
    Public blnGenerateBatch As Boolean
    Public intAnimStep As Integer
    Public batchCount As Integer = 0
    Private blnAniPause As Boolean = False
    Dim blnDisplay As Boolean = False
    '---for animated icon---
    Dim myBitmap1 As Bitmap = My.Resources.arrow_f01
    Dim myBitmap2 As Bitmap = My.Resources.arrow_f02
    Dim myBitmap3 As Bitmap = My.Resources.arrow_f03
    Dim myBitmap4 As Bitmap = My.Resources.arrow_f04
    Dim myBitmap5 As Bitmap = My.Resources.arrow_f05
    Dim myBitmap6 As Bitmap = My.Resources.arrow_f06
    Dim myBitmap7 As Bitmap = My.Resources.arrow_f07
    Dim myBitmap8 As Bitmap = My.Resources.arrow_f08
    Dim Icon1 = Icon.FromHandle(myBitmap1.GetHicon)
    Dim Icon2 = Icon.FromHandle(myBitmap2.GetHicon)
    Dim Icon3 = Icon.FromHandle(myBitmap3.GetHicon)
    Dim Icon4 = Icon.FromHandle(myBitmap4.GetHicon)
    Dim Icon5 = Icon.FromHandle(myBitmap5.GetHicon)
    Dim Icon6 = Icon.FromHandle(myBitmap6.GetHicon)
    Dim Icon7 = Icon.FromHandle(myBitmap7.GetHicon)
    Dim Icon8 = Icon.FromHandle(myBitmap8.GetHicon)
    Private Sub RunBatchGenerator()
        '---start batch generator---
        If batchCount = 0 Then
            Dim b As New Thread(AddressOf BatchGenerate)
            b.IsBackground = True
            b.Name = "Generator"
            b.Start()
            batchCount += 1
        End If
        '---end start batch generator---
    End Sub
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Me.GameType.SelectedIndex >= 0 Then
            Select Case Me.GameType.SelectedIndex
                Case 0
                    blnSamurai = False
                    frmGame.Text = "Classic sudoku solver"
                Case 1
                    blnSamurai = True
                    frmGame.Text = "Samurai sudoku solver"
            End Select
            Me.Opacity = 0
            Me.Visible = False
            Me.ShowInTaskbar = False
            Me.Refresh()

            If My.Settings._blnShowTips Then frmTipOfTheDay.ShowDialog(Me)
            frmGame.ShowDialog(Me)
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub frmInitial_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.NotifyIconProgram.Visible = False
        Me.NotifyIconGenerator.Visible = False
    End Sub
    Private Sub Initial_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = My.Application.Info.Title & " - select game type"
        Dim bm As Bitmap = My.Resources.alert
        Dim icon As IntPtr = bm.GetHicon()
        Dim newIcon As Icon = System.Drawing.Icon.FromHandle(icon)
        Me.Icon = newIcon
        Me.Icon = My.Resources.sudoku
        Me.NotifyIconProgram.Icon = My.Resources.sudoku
        Me.NotifyIconProgram.Text = My.Application.Info.Title
        Me.NotifyIconGenerator.Visible = False
        Me.GameType.SelectedIndex = 0
        Me.GameType.Focus()
        Me.NotifyIconProgram.ContextMenuStrip = Me.ContextMenuStrip1
        Me.PuzzleGenerationOptionsToolStripMenuItem.Visible = False
        '---batch generator---
        If blnGenerateBatch Then
            Me.Visible = False
            Me.ShowInTaskbar = False
            Me.PuzzleGenerationOptionsToolStripMenuItem.Visible = True
        Else
        End If
        Me.AniTimer.Enabled = True
        blnAniPause = False
        RunBatchGenerator()
        '---end batch generator---
    End Sub
    Private Sub frmInitial_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.GameType.SelectedIndex = 0
        Me.GameType.Focus()
    End Sub
    Private Function MakeFolder(ByVal strDirectory As String) As Boolean
        If Not My.Computer.FileSystem.DirectoryExists(strDirectory) Then
            Try
                My.Computer.FileSystem.CreateDirectory(strDirectory)
            Catch
                With Me.NotifyIconProgram
                    .BalloonTipTitle = My.Application.Info.Title & " - puzzle generator"
                    .BalloonTipText = "Error. Exiting puzzle generation as could not create directory " & strDirectory
                    .BalloonTipIcon = ToolTipIcon.Error
                    .ShowBalloonTip(10000)
                End With
                Exit Function
            End Try
        End If
        MakeFolder = True
    End Function
    Private Function WriteFile(ByVal strFile As String, ByVal strContent As String, ByVal blnAppend As Boolean) As Boolean
        Try
            My.Computer.FileSystem.WriteAllText(strFile, strContent, blnAppend)
        Catch
            With Me.NotifyIconProgram
                .BalloonTipTitle = My.Application.Info.Title & " - puzzle generator"
                .BalloonTipText = "Error. Exiting puzzle generation as could not create " & strFile
                .BalloonTipIcon = ToolTipIcon.Error
                .ShowBalloonTip(10000)
            End With
            Exit Function
        End Try
        WriteFile = True
    End Function
    Private Sub BatchGenerate()
        Dim s As Integer
        Dim sPtr As Integer
        Dim blnGenerated As Boolean = False
        Dim strFile As String = ""
        Dim intAllDifficulty As Integer = 0
        Dim arrGeneratedPuzzles(UBound(arrDifficultStr) + 1) As String
        Dim arrSymmetry(0) As Integer
        Dim strSavePuzzle As String = ""
        Dim strSaveSolution As String = ""

        '---get total sum of bits for all difficulties---
        Dim j As Integer
        For j = 0 To UBound(arrDifficultStr)
            intAllDifficulty += intGetBit(j + 1)
        Next

        Me.NotifyIconGenerator.Icon = Icon1
        Me.NotifyIconGenerator.Visible = True
        Me.NotifyIconGenerator.Text = My.Application.Info.Title & " puzzle generator. Hold mouse down for info"

        If Not MakeFolder(dirGames) Then GoTo StepOut

        If My.Settings._intMaxGenerate = 0 Then
            If blnGenerateBatch Then
                With Me.NotifyIconGenerator
                    .BalloonTipTitle = My.Application.Info.Title & " - puzzle generator"
                    .BalloonTipText = "Puzzle generator is currently disabled as the number of puzzles to generate is set to zero. Right click on the " & My.Application.Info.Title & " icon to view or change puzzle generation options."
                    .BalloonTipIcon = ToolTipIcon.Warning
                    .ShowBalloonTip(5000)
                End With
            End If
        End If

restart:

        '---generate classic puzzles---
        sPtr = 0
        While CountPuzzles("*.sdm") < My.Settings._intMaxGenerate
            blnAniPause = False
            If Thread.CurrentThread.Name = "Generator" Then
                If Thread.CurrentThread.Priority <> My.Settings._GeneratePriority - 1 Then
                    Thread.CurrentThread.Priority = My.Settings._GeneratePriority - 1
                    Debug.Print(Thread.CurrentThread.Name & " set to priority " & Thread.CurrentThread.Priority)
                End If
            End If
            arrSymmetry = bitmask2Arr(My.Settings._GenClassicSymmetry)
            If sPtr > UBound(arrSymmetry) Then sPtr = 0
            s = arrSymmetry(sPtr) - 1
            While blnGenerated = 0
                blnGenerated = GenerateNewGrid(intSymmetry:=s, intDifficulty:=intAllDifficulty, arrPuzzles:=arrGeneratedPuzzles, intMaxAttempts:=250)
            End While
            sPtr += 1
            For j = UBound(arrGeneratedPuzzles) To 1 Step -1
                If arrGeneratedPuzzles(j) <> "" Then
                    If intGetBit(j) And My.Settings._GenClassicDifficulty Then
                        strSavePuzzle = arrGeneratedPuzzles(j)
                        strSaveSolution = arrGeneratedPuzzles(0)
                        Dim ic() As String = Split(strSavePuzzle, ".")
                        Dim cc As Integer = 81 - ic.GetUpperBound(0)
                        '---save puzzle---
                        If Not MakeFolder(dirGames) Then GoTo StepOut
                        strFile = dirGames & "\" & strDateStamp() & "_" & strFirstLetters(intDifficult2String(j)) & "_" & LCase(intSymmetrytoString(s)) & "_symmetry_" & cc & "_clues" & ".sdm"
                        If CountPuzzles("*.sdm") >= My.Settings._intMaxGenerate Then Exit For
                        If My.Settings._blnGeneratePuzzleSaveSolution Then
                            strSavePuzzle = strSavePuzzle & vbCrLf & strSaveSolution
                        End If
                        If Not WriteFile(strFile:=strFile, strContent:=strSavePuzzle, blnAppend:=False) Then GoTo StepOut
                        '---end save puzzle---
                    End If
                End If
            Next
            blnGenerated = False
        End While
        '---end classic---

        blnAniPause = True
        Thread.Sleep(3000)
        '---batch generator paused---
        GoTo restart

stepout:

        Me.AniTimer.Stop()
        Me.NotifyIconGenerator.Visible = False
        batchCount -= 1

        If blnGenerateBatch Then
            Thread.Sleep(10000)
            Me.Close()
        End If

    End Sub
    Private Function strDateStamp() As String
        Return DateTime.Now.ToString("yyyyMMddHHmmss")
    End Function
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub
    Private Sub PuzzleGenerationOptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PuzzleGenerationOptionsToolStripMenuItem.Click
        Dim t As TabPage
        For Each t In frmOptions.TabControl1.TabPages
            If t.Text <> "Puzzle generator" Then
                frmOptions.TabControl1.TabPages.Remove(t)
            End If
        Next
        frmOptions.Show()
    End Sub
    Private Sub AniTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AniTimer.Tick
        If blnAniPause Then Exit Sub
        Dim curIcon As Icon = Icon1
        Dim prevIcon As Icon
        Select Case intAnimStep
            Case 1
                curIcon = Icon1
            Case 2
                curIcon = Icon2
            Case 3
                curIcon = Icon3
            Case 4
                curIcon = Icon4
            Case 5
                curIcon = Icon5
            Case 6
                curIcon = Icon6
            Case 7
                curIcon = Icon7
            Case 8
                curIcon = Icon8
            Case Else
        End Select
        '---set icon---
        NotifyIconGenerator.Icon = curIcon
        prevIcon = NotifyIconGenerator.Icon
        '---cleanup---
        Try
            prevIcon.Dispose()
            curIcon.Dispose()
        Catch ex As Exception
        End Try
        '---iterate counter---
        intAnimStep += 1
        If intAnimStep > 8 Then intAnimStep = 1
    End Sub
    Private Sub NotifyIcon_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIconGenerator.BalloonTipClicked
        blnDisplay = False
    End Sub
    Private Sub NotifyIcon_BalloonTipClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIconGenerator.BalloonTipClosed
        blnDisplay = False
    End Sub
    Private Sub NotifyIcon_BalloonTipShown(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIconGenerator.BalloonTipShown
        blnDisplay = True
    End Sub
    Private Sub NotifyIcon_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIconGenerator.MouseDown
        Dim c As Integer = 0
        Dim intCount As Integer = 0
        intCount = CountPuzzles("*.sdm")
        If intCount = -1 Then Exit Sub

        Dim strMsg As String = ""
        If Not blnAniPause Then
            strMsg += "Running: "
        Else
            strMsg += "Paused: "
        End If
        strMsg += "generator is set to find " & strMulti(My.Settings._intMaxGenerate, "puzzle", False) & ". "

        If intCount > 0 Then

            strMsg += "There "
            strMsg += strMulti(intValue:=intCount, strText:="puzzle", blnPrefix:=True)

            Dim j As Integer
            For j = 0 To UBound(arrDifficultStr)
                intCount = CountPuzzles("*_" & strFirstLetters(arrDifficultStr(j)) & "_*.sdm")
                If intCount = -1 Then Exit Sub
                If intCount > 0 Then
                    If c = 0 Then
                        strMsg += " ("
                    Else
                        strMsg += "; "
                    End If
                    strMsg += CStr(intCount) & " " & LCase(arrDifficultStr(j))
                    c += 1
                End If
            Next
            If c > 0 Then strMsg += "). "

        End If

        If blnGenerateBatch Then
            strMsg += "Right click on the " & My.Application.Info.Title & " icon"
        Else
            strMsg += "Select 'Options' from main menu"
        End If
        strMsg += " to view or change puzzle generation options"

        Me.NotifyIconGenerator.BalloonTipText = strMsg
        Me.NotifyIconGenerator.BalloonTipIcon = ToolTipIcon.Info
        Me.NotifyIconGenerator.BalloonTipTitle = My.Application.Info.Title & " - puzzle generator"
        If blnDisplay = False Then Me.NotifyIconGenerator.ShowBalloonTip(0)
    End Sub
    Private Sub NotifyIcon_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIconGenerator.MouseUp
        blnDisplay = False
        Me.NotifyIconGenerator.Visible = False
        Me.NotifyIconGenerator.Visible = True
    End Sub
End Class
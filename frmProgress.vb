Imports System
Imports System.ComponentModel
Imports System.Threading
Public Class frmProgress
    '---for reporting on background thread operations---
    Delegate Sub UI(ByVal Count As Integer, ByRef intSolutions As Integer, ByVal blnFinished As Boolean, ByVal strSolution As String, ByVal strDifficulty As String)
    Public strGrid As String = ""
    Public blnUnique As Boolean = False
    Public intCountSolutions As Integer = 0
    Public strGameSolution As String = ""
    Public intQuit As Integer
    Dim bgSolver As New clsSudokuSolver
    Dim bgThread As Thread
    Dim executionTime As New Stopwatch()
    Private Sub updateInterface(ByVal Count As Integer, ByRef intSolutions As Integer, ByVal blnFinished As Boolean, ByVal strSolution As String, ByVal strDifficulty As String)
        If Count = 0 And Not blnFinished Then intSolutions = 0
        If Me.Label2.InvokeRequired Then
            Dim newDelegate As New UI(AddressOf updateInterface)
            Me.Label2.Invoke(newDelegate, Count, intSolutions, blnFinished, strSolution, strDifficulty)
            Exit Sub
        Else
            If Count > 0 Then Me.Label3.Text = "Tested " & Count & " steps so far"
            Select Case intSolutions
                Case 0
                    Me.Label2.Text = "No solutions found yet..."
                Case 1
                    Me.Label2.Text = "One solution found, searching for more..."
                Case Else
                    Me.Label2.Text = intSolutions & " solutions found"
                    System.Threading.Thread.Sleep(10)
            End Select
            If blnFinished Then
                Me.Label3.Text = "Search completed after " & Count & " steps"
                executionTime.Stop()
                If intSolutions = 1 Then
                    blnUnique = True
                Else
                    blnUnique = False
                End If
                intCountSolutions = intSolutions
                strGameSolution = strSolution
            End If
        End If
    End Sub
    Private Sub frmProgress_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Label2.Text = ""
        Me.Label3.Text = ""
    End Sub
    Private Sub frmProgress_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If bgThread.IsAlive Then
            StopThreadWork = True
            System.Threading.Thread.Sleep(1000)
        End If
    End Sub
    Private Sub frmProgress_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Label2.Text = ""
        Me.Label3.Text = ""
        AddHandler bgSolver.UI, AddressOf updateInterface
    End Sub
    Private Sub frmProgress_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Label2.Text = ""
        Me.Label3.Text = ""
        bgSolver.strGameSolution = ""
        bgSolver.strDifficulty = ""
        bgSolver.strGrid = strGrid
        bgSolver.intQuit = intQuit
        bgSolver.blnClassic = False
        bgSolver.intCountSolutions = 0
        bgSolver.vsSolvers = My.Settings._defaultSolvers
        StopThreadWork = False
        intCountSolutions = 0
        strGameSolution = ""
        executionTime.Reset()
        executionTime.Start()
        Dim MyThread1 As New Thread(AddressOf bgSolver._Unique)
        MyThread1.IsBackground = True
        MyThread1.Priority = ThreadPriority.Highest
        MyThread1.Start()
        Timer1.Interval = 1000
        Timer1.Start()
        bgThread = MyThread1
    End Sub
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Not bgThread.IsAlive() Then
            Timer1.Dispose()
            Me.Close()
        End If
    End Sub
End Class
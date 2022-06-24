Public Class frmHidden
    Public strMsg As String
    Private Sub frmHidden_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        strMsg = ""
        frmGame.FeedbackBox.Text = strMsg
    End Sub
    Private Sub frmHidden_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Me.ExitTimer.Start()
        If e.KeyCode = 27 Then
            '---ESC key---
            StopShowAllSteps = True
        End If
    End Sub
    Private Sub frmHidden_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        Me.ExitTimer.Start()
    End Sub
    Private Sub frmHidden_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.Height = My.Computer.Screen.WorkingArea.Height
        Me.CenterToScreen()
        Me.Visible = True
        Me.Opacity = 0.01
        Me.BringToFront()
        Me.ExitTimer.Stop()
        ShowMessageLabel("", "")
        frmGame.FeedbackBox.Text = strMsg
    End Sub
    Private Sub ExitTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitTimer.Tick
        ExitTimer.Dispose()
        Me.Close()
    End Sub
End Class
Public Class frmPickGame
    Public intMax As Integer
    Public intGame As Integer
    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OKButton.Click
        If IsNumeric(Me.NoGames.Text) Then
        Else
            MsgBox("Input must be numeric")
            Me.NoGames.Text = intGame
            Me.NoGames.Focus()
            Exit Sub
        End If
        If CInt(Me.NoGames.Text) < 1 Or CInt(Me.NoGames.Text) > intMax Then
            MsgBox("You must select a number between 1 and " & intMax)
            Me.NoGames.Text = intGame
            Me.NoGames.Focus()
            Exit Sub
        End If
        Me.Hide()
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
    Private Sub frmPickGame_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.LblMaxGames.Text = " of " & intMax
        Me.NoGames.Text = intGame
        Me.NoGames.MaxLength = Len(CStr(intMax))
    End Sub
End Class
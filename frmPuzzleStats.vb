Public Class frmPuzzleStats
    Public blnShown As Boolean = False
    Private Sub frmPuzzleStats_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel1.Top = 15
        Me.TableLayoutPanel1.Left = 15
        Dim r As Integer
        Dim c As Integer
        Dim i As Integer
        Dim j As Integer

        Dim intCount As Integer = 0
        Dim intTotalCount As Integer = 0
        Dim intPercent As Integer
        intTotalCount = CountPuzzles("*.sdm")
        If intTotalCount < 1 Then
            Me.Close()
            Exit Sub
        End If

        Dim sym() As String
        sym = [Enum].GetNames(GetType(Symmetry.symmetryTypes))

        Me.TableLayoutPanel1.Controls.Clear()
        Me.TableLayoutPanel1.AutoSize = True
        c = UBound(arrDifficultStr) + 3
        r = UBound(sym) + 3
        Me.TableLayoutPanel1.ColumnCount = c
        Me.TableLayoutPanel1.RowCount = r

        For i = 0 To UBound(sym) + 1
            For j = 0 To UBound(arrDifficultStr) + 1
                If i > 0 And j > 0 And j < c And i < r Then
                    intCount = CountPuzzles("*_" & strFirstLetters(arrDifficultStr(j - 1)) & "_" & sym(i - 1) & "_symmetry_*.sdm")
                    If intCount = -1 Then Me.Close()
                    If intCount > 0 Then
                        intPercent = (intCount / intTotalCount) * 100
                    End If
                    Dim lbl As New Label
                    lbl.AutoSize = True
                    lbl.Text = intCount
                    Me.TableLayoutPanel1.Controls.Add(lbl, j, i)
                End If
                '---first column---
                If j = 0 And i > 0 Then
                    Dim lbl As New Label
                    lbl.AutoSize = True
                    lbl.Text = sym(i - 1) & " symmetry"
                    Me.TableLayoutPanel1.Controls.Add(lbl, j, i)
                End If
                '---first row---
                If i = 0 And j > 0 Then
                    Dim lbl As New Label
                    lbl.AutoSize = True
                    lbl.Text = arrDifficultStr(j - 1)
                    Me.TableLayoutPanel1.Controls.Add(lbl, j, i)
                End If
            Next
        Next

        j = 0
        For i = 0 To UBound(sym)
            '---last column---
            intCount = CountPuzzles("*_" & sym(i) & "_symmetry_*.sdm")
            If intCount = -1 Then Me.Close()
            Dim lbl As New Label
            lbl.AutoSize = True
            lbl.Text = intCount
            lbl.Font = New Font(lbl.Font, FontStyle.Bold)
            Me.TableLayoutPanel1.Controls.Add(lbl, c, i + 1)
        Next

        i = 0
        For j = 0 To UBound(arrDifficultStr)
            '---last row---
            intCount = CountPuzzles("*_" & strFirstLetters(arrDifficultStr(j)) & "_*.sdm")
            If intCount = -1 Then Me.Close()
            Dim lbl As New Label
            lbl.AutoSize = True
            lbl.Text = intCount
            lbl.Font = New Font(lbl.Font, FontStyle.Bold)
            Me.TableLayoutPanel1.Controls.Add(lbl, j + 1, r)
        Next

        '---last row/col---
        Dim tlbl As New Label
        tlbl.AutoSize = True
        tlbl.Text = intTotalCount
        tlbl.Font = New Font(tlbl.Font, FontStyle.Bold)
        Me.TableLayoutPanel1.Controls.Add(tlbl, c, r)

        '---first row/last col---
        Dim tlbl1 As New Label
        tlbl1.AutoSize = True
        tlbl1.Text = "Total"
        tlbl1.Font = New Font(tlbl1.Font, FontStyle.Bold)

        Me.TableLayoutPanel1.Controls.Add(tlbl1, c, 0)

        '---first col/last row---
        Dim tlbl2 As New Label
        tlbl2.AutoSize = True
        tlbl2.Text = "Total"
        tlbl2.Font = New Font(tlbl2.Font, FontStyle.Bold)
        Me.TableLayoutPanel1.Controls.Add(tlbl2, 0, r)

        Me.TableLayoutPanel1.ResumeLayout()

        Dim frmSize As New Size(Me.TableLayoutPanel1.Width + 30, Me.TableLayoutPanel1.Height + 30)
        Me.ClientSize = frmSize
        blnShown = True

    End Sub
End Class
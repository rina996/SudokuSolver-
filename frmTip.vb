Public Class frmTipOfTheDay
    Dim arrStrTip(19) As String
    Dim Icon1 = Icon.FromHandle(My.Resources.lightbulb.GetHicon)
    Function LoadTips() As Boolean
        'TODO: Add more tips?
        arrStrTip(0) = "If there is only a single candidate (naked single) remaining for a cell, you can double mouse click to insert this value as the solution for the cell (only applicable when pencilmarks are shown)."
        arrStrTip(1) = "To select multiple cells, double click each cell that you want to select. The checkboxes at the left of the screen will then be enabled, and can then be used to enable or disable candidates across all selected cells (only applicable when pencilmarks are shown). To deselect multiple cells, simply do a single click."
        arrStrTip(2) = "For information on the puzzle, hover your mouse over the information icon at the bottom right hand corner of the main screen."
        arrStrTip(3) = "To manually input clues for a puzzle (e.g if entering a puzzle from a book or newspaper), starting with a blank game, place the mouse cursor over the relevant cell and use the numeric keys from 1 to 9 to enter clues. Use DELETE or BACKSPACE to remove incorrect values. You will receive messages regarding the validity of the puzzle. Once the clues entered result in a unique solution, you can then copy and paste the puzzle into the program."
        arrStrTip(4) = "The program has a puzzle generator (an arrow icon in the windows taskbar will animate when puzzles are being generated). If the program runs slowly, you can try reducing the priority or turn off this feature (Select 'Options' and then the 'Puzzle generator' tab)"
        arrStrTip(5) = "To turn sounds on or off click on the small speaker icon in the bottom right hand corner of the main screen."
        arrStrTip(6) = "If you open a file with multiple puzzles, use the PAGE UP or PAGE DOWN keys to navigate through these puzzles. Alternatively, click on the status bar where the text 'Puzzle x of y' appears to go directly to a particular puzzle."
        arrStrTip(7) = "You can right click on cells to enable or disable candidates or to input the value for a particular cell."
        arrStrTip(8) = "If you get stuck while solving a puzzle, press CTRL+H to get a hint (for most methods, you can hit CTRL+H more than once to get progressively more detailed hints)."
        arrStrTip(9) = "If you are manually solving a puzzle, you can select the 'check consistency' option from the 'Solve' menu to double check that the moves you have made so far are valid."
        arrStrTip(10) = "Use CTRL+C or CTRL+V to copy and paste puzzles. Puzzles are stored as a string with numbers from 1 to 9 showing solved cells/clues and zeros or full stops to show unfilled cells."
        arrStrTip(11) = "Use the optimise or optimise (maintain symmetry) options under the 'Extras' menu to remove clues while still ensuring a unique solution."
        arrStrTip(12) = "Use the batch solve option under the 'Extra' menu to process a number of puzzles stored in a single file. Each puzzle within the file should be on an individual line."
        arrStrTip(13) = "Use CTRL+Y and CTRL+Z to undo/redo moves."
        arrStrTip(14) = "To make naked singles stand out use the 'Options' menu and select 'highlight naked singles' from the 'display' tab."
        arrStrTip(15) = "Selecting 'next step' from the 'solve' menu will complete the next valid move and explain the reason for this move."
        arrStrTip(16) = "You can enable/disable the display of candidates/pencilmarks by clicking on the icon in the bottom right hand corner of the main screen."
        arrStrTip(17) = "You can open a puzzle by dragging and dropping a file or files with a .txt or .sdm extension."
        arrStrTip(18) = "Selecting 'show all steps' from the 'solve' menu will step through the puzzle and detail the reason for each move."
        arrStrTip(19) = "To customise text colours etc, select the 'Options' menu and select the 'display' tab."
    End Function
    Function getCurrentTip() As String
        Dim intTip As Integer = My.Settings._intTipoftheDay + 1
        If intTip >= arrStrTip.Length Then intTip = 0
        My.Settings._intTipoftheDay = intTip
        getCurrentTip = "Tip #" & intTip + 1 & ": " & arrStrTip(intTip)
    End Function
    Function getNextTip() As String
        Dim intTip As Integer = My.Settings._intTipoftheDay
        If intTip = arrStrTip.Length - 1 Then intTip = -1
        intTip += 1
        getNextTip = "Tip #" & intTip + 1 & ": " & arrStrTip(intTip)
        My.Settings._intTipoftheDay = intTip
        Me.TextBox1.Text = getNextTip
    End Function
    Function getprevTip() As String
        Dim intTip As Integer = My.Settings._intTipoftheDay
        If intTip = 0 Then intTip = arrStrTip.Length
        intTip -= 1
        getprevTip = "Tip #" & intTip + 1 & ": " & arrStrTip(intTip)
        My.Settings._intTipoftheDay = intTip
        Me.TextBox1.Text = getprevTip
    End Function
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        getprevTip()
    End Sub
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        getNextTip()
    End Sub
    Private Sub frmTipOfTheDay_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadTips()
        Me.TextBox1.Text = getCurrentTip()
        Me.Icon = Icon1
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If Me.TextBox1.Text <> "" Then
            Dim sz = New Size(TextBox1.ClientSize.Width, TextBox1.ClientSize.Height)
            Dim flags As TextFormatFlags = TextFormatFlags.WordBreak
            Dim Padding As Integer = 5
            Dim Borders As Integer = TextBox1.Height - TextBox1.ClientSize.Height
            sz = TextRenderer.MeasureText(TextBox1.Text, TextBox1.Font, sz, flags)
            Dim h As Integer = sz.Height + Borders + Padding
            TextBox1.Height = h
            Dim intMenu As Integer = Me.Height - Me.ClientSize.Height
            Me.CheckBox1.Top = Me.TextBox1.Height + Me.TextBox1.Top + 1
            h = intLowestCtrl() + intMenu
            Me.Height = h + 3
        End If
    End Sub
    Function intLowestCtrl() As Integer
        Dim i As Integer
        Dim lc As Integer
        Dim oControl As Control
        For Each oControl In Me.Controls
            i = oControl.Top + oControl.Height
            If i > lc Then lc = i
        Next
        intLowestCtrl = lc
    End Function
End Class
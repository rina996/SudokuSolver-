Imports System.Drawing.Drawing2D
Imports System.Windows.Forms.Design
Public Class SudokuCell
    Private Const WS_EX_TRANSPARENT As Int32 = &H20
#Region "Variables and Enumerations"
    Private textCol As System.Drawing.Color = Color.Black
    Private BGCol As System.Drawing.Color = Color.Transparent
    Private BorderCol As System.Drawing.Color = Color.Black
    Private blnAlertGraphic As Boolean
    Private intText As Integer = 511 'sum of bits from 1 to 9 (each bit is n^2)
    Private intSol As Integer = 0
    Private IntBorder As Integer = 0
    Private intHighlight1 As Integer = 0
    Private intHighlight2 As Integer = 0
    Private intHighlight3 As Integer = 0
    Private intHighlight4 As Integer = 0
    Private intCellH As Integer = 16
    Private intCellW As Integer = 16
    Private blnHighlightNS As Boolean
    Private HighlightCol1 As System.Drawing.Color = Color.Transparent
    Private HighlightCol2 As System.Drawing.Color = Color.Transparent
    Private HighlightCol3 As System.Drawing.Color = Color.Transparent
    Private HighlightCol4 As System.Drawing.Color = Color.Transparent
    Private e As PaintEventArgs = Nothing
#End Region
#Region "Control Properties"
    Public Property sc_HighlightNaked() As Boolean
        Get
            Return blnHighlightNS
        End Get
        Set(ByVal value As Boolean)
            blnHighlightNS = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_AlertGraphic() As Boolean
        Get
            Return blnAlertGraphic
        End Get
        Set(ByVal value As Boolean)
            blnAlertGraphic = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntW() As Integer
        Get
            Return intCellW
        End Get
        Set(ByVal value As Integer)
            intCellW = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntH() As Integer
        Get
            Return intCellH
        End Get
        Set(ByVal value As Integer)
            intCellH = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntBorders() As Integer
        Get
            Return IntBorder
        End Get
        Set(ByVal value As Integer)
            IntBorder = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_TextColour() As System.Drawing.Color
        Get
            Return textCol
        End Get
        Set(ByVal value As System.Drawing.Color)
            textCol = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_BorderColour() As System.Drawing.Color
        Get
            Return BorderCol
        End Get
        Set(ByVal value As System.Drawing.Color)
            BorderCol = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_BackgroundColour() As System.Drawing.Color
        Get
            Return BGCol
        End Get
        Set(ByVal value As System.Drawing.Color)
            BGCol = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntSolution() As Integer
        Get
            Return intSol
        End Get
        Set(ByVal value As Integer)
            intSol = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntHighlightCandidate1() As Integer
        Get
            Return intHighlight1
        End Get
        Set(ByVal value As Integer)
            intHighlight1 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntHighlightCandidate2() As Integer
        Get
            Return intHighlight2
        End Get
        Set(ByVal value As Integer)
            intHighlight2 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntHighlightCandidate3() As Integer
        Get
            Return intHighlight3
        End Get
        Set(ByVal value As Integer)
            intHighlight3 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntHighlightCandidate4() As Integer
        Get
            Return intHighlight4
        End Get
        Set(ByVal value As Integer)
            intHighlight4 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_IntCandidate() As Integer
        Get
            Return intText
        End Get
        Set(ByVal value As Integer)
            intText = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_HighlightColour1() As System.Drawing.Color
        Get
            Return HighlightCol1
        End Get
        Set(ByVal value As System.Drawing.Color)
            HighlightCol1 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_HighlightColour2() As System.Drawing.Color
        Get
            Return HighlightCol2
        End Get
        Set(ByVal value As System.Drawing.Color)
            HighlightCol2 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_HighlightColour3() As System.Drawing.Color
        Get
            Return HighlightCol3
        End Get
        Set(ByVal value As System.Drawing.Color)
            HighlightCol3 = value
            Me.Invalidate()
        End Set
    End Property
    Public Property sc_HighlightColour4() As System.Drawing.Color
        Get
            Return HighlightCol4
        End Get
        Set(ByVal value As System.Drawing.Color)
            HighlightCol4 = value
            Me.Invalidate()
        End Set
    End Property
#End Region
#Region "Drawing Functions"
    Protected Overrides Sub OnPaint(ByVal ea As System.Windows.Forms.PaintEventArgs)
        e = ea
        e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Low ' or NearestNeighbour
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
        e.Graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.None
        e.Graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed
        e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault
        DrawText(intText, intHighlight1, intHighlight2, intHighlight3, intHighlight4)
    End Sub
    Protected Sub InvalidateEx()
        If Parent Is Nothing Then
            Return
        End If
        Dim rc As New Rectangle(Me.Location, Me.Size)
        Parent.Invalidate(rc, True)
    End Sub
    '---Draws text in the specified position---
    Private Sub DrawText(ByVal intText As Integer, ByVal intH1 As Integer, ByVal intH2 As Integer, ByVal intH3 As Integer, ByVal intH4 As Integer)
        Dim i As Integer
        Dim c As Integer
        Dim h As Integer
        Dim intBit As Integer
        Dim intOnBits As Integer
        Dim intOBit As Integer = intText

        drawBackgroundandBorders()

        If blnAlertGraphic Then
            drawAlert()
            Exit Sub
        End If

        If intSol >= 1 And intSol <= 9 Then
            '---draw solution---
            drawSolution(intSol)
        Else
            '---make sure highligted candidates are removed from list of overall candidates---
            For i = 1 To 4
                Select Case i
                    Case 1
                        h = intH1
                    Case 2
                        h = intH2
                    Case 3
                        h = intH3
                    Case 4
                        h = intH4
                End Select
                For c = 0 To 8
                    intBit = intGetBit(c + 1)
                    If CBool(intBit And h) And CBool(intBit And intText) Then
                        intText -= intBit
                        intOnBits += intBit
                    End If
                Next
            Next
            If blnSingleBit(intOBit) Then
                '---draw single candidate---
                drawCandidate(intCandidate:=intReverseBit(intOBit), blnSingleCandidate:=True)
            Else
                For c = 0 To 8
                    intBit = intGetBit(c + 1)
                    If CBool(intBit And intText) Then
                        '---draw normal candidate---
                        drawCandidate(intCandidate:=c + 1)
                    End If
                Next
            End If
            '---draw highlighted candidates---
            For i = 1 To 4
                Select Case i
                    Case 1
                        h = intH1
                    Case 2
                        h = intH2
                    Case 3
                        h = intH3
                    Case 4
                        h = intH4
                End Select
                For c = 0 To 8
                    intBit = intGetBit(c + 1)
                    If CBool(intBit And h) And CBool(intBit And intOnBits) Then
                        '---draw highlight---
                        drawCandidate(intCandidate:=c + 1, highlightColour:=i)
                    End If
                Next
            Next
        End If
    End Sub
    Private Sub drawAlert()
        Dim alertBitmap As New Bitmap(My.Resources.alert)
        Dim bmH As Integer = alertBitmap.Height
        Dim bmW As Integer = alertBitmap.Width
        Dim xPos As Integer = ((intCellW * 3) - bmW) / 2
        Dim yPos As Integer = ((intCellH * 3) - bmH) / 2
        e.Graphics.DrawImage(alertBitmap, xPos, xPos)
    End Sub
    Private Sub drawBackgroundandBorders()
        '---draw bg---
        e.Graphics.Clear(Me.Parent.BackColor)
        Dim fillBrush As New System.Drawing.SolidBrush(BGCol)
        e.Graphics.FillRectangle(fillBrush, New Rectangle(0, 0, Me.Width, Me.Height))
        fillBrush.Dispose()
        '---draw borders---
        Dim borderWidth As Integer = 1.0F
        Dim borderPen As New Pen(sc_BorderColour, borderWidth)
        borderPen.DashStyle = DashStyle.Dot
        borderPen.DashPattern = New Single() {1.3F, 1.3F}
        borderPen.LineJoin = LineJoin.Round
        Dim i As Integer
        Dim intBit As Integer
        Dim StartPoint As Point
        Dim EndPoint As Point
        For i = 1 To 4
            intBit = intGetBit(i)
            If intBit And sc_IntBorders Then
                '---draw line---
                Select Case intBit
                    Case 1
                        '---left---
                        StartPoint.X = 0
                        StartPoint.Y = 0
                        EndPoint.X = 0
                        EndPoint.Y = -0 + intCellH * 3
                    Case 2
                        '---right---
                        StartPoint.X = -0 + (intCellW * 3) - borderWidth
                        StartPoint.Y = 0
                        EndPoint.X = -0 + (intCellW * 3) - borderWidth
                        EndPoint.Y = -0 + (intCellH * 3)
                    Case 4
                        '---top---
                        StartPoint.X = 0
                        StartPoint.Y = 0
                        EndPoint.X = -0 + intCellW * 3
                        EndPoint.Y = 0
                    Case 8
                        '---bottom---
                        StartPoint.X = 0
                        StartPoint.Y = -0 + (intCellH * 3) - borderWidth
                        EndPoint.X = -0 + intCellW * 3
                        EndPoint.Y = -0 + (intCellH * 3) - borderWidth
                End Select
                e.Graphics.DrawLine(borderPen, StartPoint, EndPoint)
            End If
        Next
        borderPen.Dispose()
    End Sub
    Private Sub drawCandidate(ByVal intCandidate As Integer, Optional ByVal highlightColour As Integer = 0, Optional ByVal blnSingleCandidate As Boolean = False)
        If Not My.Settings._blnCandidates Then Exit Sub
        '---size and position of candidate---
        Dim ctrlBounds As New Rectangle(intXOffset(intCandidate), intYOffset(intCandidate), intCellH, intCellW)
        '---text to display---
        Dim strFontText As String = CStr(intCandidate)
        '---font---
        Dim fSize As Double = My.Settings._CandidateFontSize
        Dim fName As String = My.Settings._CandidateFont
        If (Me.ParentForm.Name <> "frmGame" Or frmGame.strPuzzleSolution <> "") AndAlso blnSingleCandidate AndAlso My.Settings._ShowNS AndAlso My.Settings._blnCandidates Then
            '---naked single---
            fSize = My.Settings._SolvedFontSize
            fName = My.Settings._SolvedFont
        End If
        Dim tColour As Color
        Dim textFont As New Font(fName, fSize, FontStyle.Regular, GraphicsUnit.Point)
        '---draw transparent bounding box---
        Dim rectanglePen As New Pen(Color.Transparent, 0)
        '---brush for drawing text---
        tColour = textCol
        If (Me.ParentForm.Name <> "frmGame" Or frmGame.strPuzzleSolution <> "") AndAlso blnSingleCandidate AndAlso My.Settings._ShowNS AndAlso My.Settings._blnCandidates Then
            tColour = Color.FromName(My.Settings._nakedColour)
        End If
        Dim textBrush As New System.Drawing.SolidBrush(tColour)
        '---set text colour to white if highlighting---
        If highlightColour > 0 Then textBrush.Color = Color.White
        '---highlight colour---
        Dim highlightBrush As New System.Drawing.SolidBrush(Color.Black)
        Select Case highlightColour
            Case 1
                highlightBrush.Color = HighlightCol1
            Case 2
                highlightBrush.Color = HighlightCol2
            Case 3
                highlightBrush.Color = HighlightCol3
            Case 4
                highlightBrush.Color = HighlightCol4
            Case Else
        End Select

        Dim stringFormat As New StringFormat()
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center

        '---Calculate the size of the supplied text---
        Dim MySize As SizeF = e.Graphics.MeasureString(strFontText, textFont)
        '---Calculate the location. Start with the center--- 
        '---of the client area. Then offset the point---
        '---half the width and height of the text.---
        Dim MyPoint As Point = New Point(intCellW / 2, intCellH / 2)
        If (Me.ParentForm.Name <> "frmGame" Or frmGame.strPuzzleSolution <> "") AndAlso blnSingleCandidate AndAlso My.Settings._ShowNS AndAlso My.Settings._blnCandidates Then
            MyPoint = New Point((intCellW * 3) / 2, (intCellH * 3) / 2)
        End If
        MyPoint.Offset(-MySize.Width / 2, -MySize.Height / 2)
        Try
            '---draw background highlight if required---
            If highlightColour > 0 Then e.Graphics.FillEllipse(highlightBrush, ctrlBounds)
            '---Draw the text---
            e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
            If (Me.ParentForm.Name <> "frmGame" Or frmGame.strPuzzleSolution <> "") AndAlso blnSingleCandidate AndAlso My.Settings._ShowNS AndAlso My.Settings._blnCandidates Then
                e.Graphics.DrawString(strFontText, textFont, textBrush, MyPoint.X, MyPoint.Y)
            Else
                e.Graphics.DrawString(strFontText, textFont, textBrush, intXOffset(intCandidate) + MyPoint.X, intYOffset(intCandidate) + MyPoint.Y)
            End If
            '---draw the bounding rectangle---
            'e.graphics.DrawRectangle(rectanglePen, ctrlBounds)
        Finally
            textFont.Dispose()
            textBrush.Dispose()
            rectanglePen.Dispose()
            highlightBrush.Dispose()
        End Try
    End Sub
    Private Sub drawSolution(ByVal intSolution As Integer, Optional ByVal textColour As String = "")
        '---size and position of control---
        Dim ctrlBounds As New Rectangle(0, 0, intCellH * 3, intCellW * 3)
        '---text to display---
        Dim strFontText As String = CStr(intSolution)
        '---font---
        Dim textFont As New Font(My.Settings._SolvedFont, My.Settings._SolvedFontSize, FontStyle.Regular, GraphicsUnit.Point)
        '---draw transparent bounding box---
        Dim rectanglePen As New Pen(Color.Transparent, 0)
        '---brush for drawing text---
        Dim textBrush As New System.Drawing.SolidBrush(textCol)
        Dim stringFormat As New StringFormat(StringFormatFlags.FitBlackBox)
        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center
        Try
            '---Draw the text---
            e.Graphics.DrawString(strFontText, textFont, textBrush, ctrlBounds, stringFormat)
            '---draw the bounding rectangle--
            'e.graphics.DrawRectangle(rectanglePen, ctrlBounds)
        Finally
            textFont.Dispose()
            textBrush.Dispose()
            rectanglePen.Dispose()
        End Try
    End Sub
    Private Function intYOffset(ByVal intCandidate As Integer) As Integer
        Select Case intCandidate
            Case 1 To 3
                intYOffset = 0
            Case 4 To 6
                intYOffset = 1
            Case 7 To 9
                intYOffset = 2
        End Select
        intYOffset = intCellW * intYOffset
    End Function
    Private Function intXOffset(ByVal intCandidate As Integer) As Integer
        Select Case intCandidate
            Case 1, 4, 7
                intXOffset = 0
            Case 2, 5, 8
                intXOffset = 1
            Case 3, 6, 9
                intXOffset = 2
        End Select
        intXOffset = intCellH * intXOffset
    End Function
#End Region
    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
    End Sub
    Public Sub New()
        '---This call is required by the Windows Form Designer---
        InitializeComponent()
        '---Add any initialization after the InitializeComponent() call---
        'Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        Me.UpdateStyles()
        intCellH = 16
        intCellW = 16
        Me.sc_BackgroundColour = Color.Blue
        Me.sc_BorderColour = Color.Black
        Me.sc_IntBorders = 15
        Me.Height = intCellH * 3
        Me.Width = intCellW * 3
    End Sub
    Private Sub SudokuCell_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, Me.SizeChanged
        Me.Height = intCellH * 3
        Me.Width = intCellW * 3
    End Sub
    Private Sub SudokuCell_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        If Me.ParentForm.Name <> "frmGame" Then Exit Sub
        If frmGame.strPuzzleSolution <> "" AndAlso blnSingleBit(Me.sc_IntCandidate) AndAlso My.Settings._ShowNS AndAlso My.Settings._blnCandidates Then
            Me.ToolTip1.SetToolTip(Me, "Naked single - double click to set this cell to " & intReverseBit(Me.sc_IntCandidate))
        End If
        If frmGame.blnSetup And Me.sc_IntSolution = 0 Then
            Me.ToolTip1.SetToolTip(Me, "Manual input mode:" & vbCrLf & "Press numeric keys to input new clue")
        End If
        If frmGame.blnSetup And Me.sc_IntSolution > 0 Then
            Me.ToolTip1.SetToolTip(Me, "Manual input mode:" & vbCrLf & "Press numeric keys to change clue or zero to delete this clue")
        End If
    End Sub
    Private Sub SudokuCell_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
        Me.ToolTip1.RemoveAll()
    End Sub
End Class

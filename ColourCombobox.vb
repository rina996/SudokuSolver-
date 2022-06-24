'=====================================================================================
'  mkc_ColorCombobox
'  combobox color picker control
'=====================================================================================
'  Created By: Marc Cramer
'  Published Date: 1/28/2003
'  Legal Copyright: Marc Cramer © 1/28/2003
'=====================================================================================
' Notes: GetIDEBackgroundColor is based on the function CalculateColor found in ColorUtil.cs written by Carlos H. Perez
'=====================================================================================
Option Strict Off
Imports System
Imports System.ComponentModel

Public Class ColourCombobox
    Inherits System.Windows.Forms.UserControl

#Region " Variable Declaration "
    '=====================================================================================
    ' VARIABLE DECLARATION
    '=====================================================================================
    Dim Known_Color() As String = Split("Black,DimGray,Gray,DarkGray,Silver,LightGray,Gainsboro,RosyBrown,IndianRed,Brown,Firebrick,LightCoral,Maroon,DarkRed,Red,Salmon,Tomato,DarkSalmon,OrangeRed,Coral,LightSalmon,Sienna,Chocolate,SaddleBrown,SandyBrown,Peru,DarkOrange,Orange,Goldenrod,DarkGoldenrod,Gold,Khaki,PaleGoldenrod,DarkKhaki,Olive,Yellow,OliveDrab,YellowGreen,DarkOliveGreen,GreenYellow,LawnGreen,Chartreuse,DarkSeaGreen,ForestGreen,LimeGreen,LightGreen,PaleGreen,DarkGreen,Green,Lime,SeaGreen,MediumSeaGreen,SpringGreen,MediumSpringGreen,MediumAquamarine,Aquamarine,Turquoise,LightSeaGreen,MediumTurquoise,DarkSlateGray,PaleTurquoise,Teal,DarkCyan,Aqua,Cyan,DarkTurquoise,CadetBlue,PowderBlue,LightBlue,DeepSkyBlue,SkyBlue,LightSkyBlue,SteelBlue,SlateGray,LightSlateGray,DodgerBlue,LightSteelBlue,CornflowerBlue,RoyalBlue,MidnightBlue,Navy,DarkBlue,MediumBlue,Blue,DarkSlateBlue,SlateBlue,MediumSlateBlue,MediumPurple,BlueViolet,Indigo,DarkOrchid,DarkViolet,MediumOrchid,Thistle,Plum,Violet,Purple,DarkMagenta,Fuchsia,Magenta,Orchid,MediumVioletRed,DeepPink,HotPink,PaleVioletRed,Crimson,Pink,LightPink", ",")
    Dim System_Color() As String = Split("ActiveBorder,ActiveCaption,ActiveCaptionText,AppWorkspace,Control,ControlDark,ControlDarkDark,ControlLight,ControlLightLight,ControlText,Desktop,GrayText,HighLight,HighLightText,HotTrack,InactiveBorder,InactiveCaption,InactiveCaptionText,Info,InfoText,Menu,MenuText,ScrollBar,Window,WindowFrame,WindowText", ",")
    Dim IDEBackgroundColor As System.Drawing.Color = CalculateIDEBackgroundColor()

    '=====================================================================================
    ' PROPERTY VARIABLES
    '=====================================================================================
    Dim m_ColorEnum As ColorEnum = ColorEnum.KnownColor
    Dim m_ColourExclude As String
    Dim m_ColourExclude2 As String
    Dim m_ColourExclude3 As String
    Dim m_ColourExclude4 As String
    Dim m_SelectedColor As String
    Dim m_FocusStyle As FocusStyleEnum = FocusStyleEnum.IDE

    '=====================================================================================
    ' ENUM DECLARATION
    '=====================================================================================
    Public Enum ColorEnum As Byte
        KnownColor = 0
        SystemColor = 1
    End Enum
    Public Enum FocusStyleEnum As Byte
        IDE = 0
        Normal = 1
    End Enum
    '=====================================================================================
#End Region

#Region " Windows Form Designer generated code "
    '=====================================================================================
    ' WINDOWS FORM DESIGNER GENERATED CODE
    '=====================================================================================
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub   'New
    '=====================================================================================
    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub   'Dispose(ByVal disposing As Boolean)
    '=====================================================================================
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cboColor As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cboColor = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'cboColor
        '
        Me.cboColor.BackColor = System.Drawing.Color.White
        Me.cboColor.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cboColor.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.cboColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboColor.Location = New System.Drawing.Point(0, 0)
        Me.cboColor.Name = "cboColor"
        Me.cboColor.Size = New System.Drawing.Size(150, 21)
        Me.cboColor.TabIndex = 0
        '
        'ColourCombobox
        '
        Me.Controls.Add(Me.cboColor)
        Me.Name = "ColourCombobox"
        Me.Size = New System.Drawing.Size(150, 24)
        Me.ResumeLayout(False)

    End Sub   'InitializeComponent()
    '=====================================================================================
#End Region

#Region " Events "
    '=====================================================================================
    'EVENTS
    '=====================================================================================
    Public Event SelectedColorChanged(ByVal SelectedColor As System.Drawing.Color, ByVal sender As Object)
    Public Event SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Private Sub cboColor_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboColor.DropDown
        Dim sColor As String
        sColor = m_SelectedColor
        LoadColors(ColorEnum.KnownColor)
        cboColor.SelectedIndex = cboColor.Items.IndexOf(sColor)
    End Sub
    '=====================================================================================
    Private Sub cboColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboColor.SelectedIndexChanged
        ' a new value was picked so set variable and fire event
        m_SelectedColor = cboColor.Text
        RaiseEvent SelectedColorChanged(System.Drawing.Color.FromName(cboColor.Text), sender)
        RaiseEvent SelectedIndexChanged(sender, e)
    End Sub   'cboColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboColor.SelectedIndexChanged
    '=====================================================================================
    Private Sub mkc_ColorPicker_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FontChanged
        ' update our font to selected values
        cboColor.Font = Me.Font
    End Sub   'mkc_ColorPicker_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FontChanged
    '=====================================================================================
#End Region

#Region " Properties "
    '=====================================================================================
    ' PROPERTIES
    '=====================================================================================
    <Description("Gets and sets the type of color shown in list."), Category("Display")> _
    Public Property ColorType() As ColorEnum
        Get
            Return m_ColorEnum
        End Get
        Set(ByVal Value As ColorEnum)
            m_ColorEnum = Value
            LoadColors(Value)
        End Set
    End Property
    '=====================================================================================
    Public Property ColourExclude() As String
        Get
            Return m_ColourExclude
        End Get
        Set(ByVal Value As String)
            m_ColourExclude = Value
        End Set
    End Property
    '=====================================================================================
    Public Property ColourExclude2() As String
        Get
            Return m_ColourExclude2
        End Get
        Set(ByVal Value As String)
            m_ColourExclude2 = Value
        End Set
    End Property
    Public Property ColourExclude3() As String
        Get
            Return m_ColourExclude3
        End Get
        Set(ByVal Value As String)
            m_ColourExclude3 = Value
        End Set
    End Property
    Public Property ColourExclude4() As String
        Get
            Return m_ColourExclude4
        End Get
        Set(ByVal Value As String)
            m_ColourExclude4 = Value
        End Set
    End Property
    '=====================================================================================
    <Description("Gets and sets the currently selected color."), Category("Display")> _
    Public Property SelectedColor() As String
        Get
            Return m_SelectedColor
        End Get
        Set(ByVal Value As String)
            m_SelectedColor = Value
            Dim ColorIndex As Integer
            ColorIndex = cboColor.FindStringExact(m_SelectedColor)
            cboColor.SelectedIndex = ColorIndex
        End Set
    End Property
    '=====================================================================================
    <Description("Gets and sets the focus style."), Category("Display")> _
    Public Property FocusStyle() As FocusStyleEnum
        Get
            Return m_FocusStyle
        End Get
        Set(ByVal Value As FocusStyleEnum)
            m_FocusStyle = Value
            ' we want IDE look so calculate the background color
            If m_FocusStyle = FocusStyleEnum.IDE Then IDEBackgroundColor = CalculateIDEBackgroundColor()
        End Set
    End Property      'FocusStyle() As FocusStyleEnum
    '=====================================================================================
#End Region

#Region " Methods "
    '=====================================================================================
    ' METHODS
    '=====================================================================================
    Private Sub cboColor_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles cboColor.DrawItem
        ' if a valid value then draw the color entry
        If e.Index <> -1 Then

            'MsgBox(cboColor.Items.Item(e.Index).ToString)
            'If cboColor.Items.Item(e.Index).ToString = ColorExclude.Name Then
            'cboColor.Items.Remove(e.Index)
            'Exit Sub
            'End If

            Dim HasFocusFlag As Boolean
            If e.State And DrawItemState.ComboBoxEdit Then
                DrawHighlight(False, e.Graphics, e.Bounds)
                HasFocusFlag = False
            ElseIf e.State And DrawItemState.Focus Then
                DrawHighlight(True, e.Graphics, e.Bounds)
                HasFocusFlag = True
            Else
                DrawHighlight(False, e.Graphics, e.Bounds)
                HasFocusFlag = False
            End If
            ColorItem(e.Graphics, e.Bounds, e.Index, m_ColorEnum, HasFocusFlag)
        End If
    End Sub   'cboColor_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles cboColor.DrawItem
    '=====================================================================================
#End Region

#Region " Helper Routines "
    '=====================================================================================
    ' HELPER ROUTINES
    '=====================================================================================
    Private Sub LoadColors(ByVal ColorEnum As ColorEnum)
        ' load our color entry list into the combobox
        Dim Counter As Integer
        Dim ColorArray() As String

        If ColorEnum = ColorEnum.KnownColor Then
            ColorArray = Known_Color
        Else
            ColorArray = System_Color
        End If

        Dim ArrayCount As Integer = ColorArray.GetUpperBound(0)
        cboColor.Items.Clear()

        For Counter = 0 To ArrayCount
            If ColorArray(Counter) = ColourExclude Or ColorArray(Counter) = ColourExclude2 Or ColorArray(Counter) = ColourExclude3 Or ColorArray(Counter) = ColourExclude4 Then
            Else
                cboColor.Items.Add(ColorArray(Counter))
            End If
        Next
    End Sub
    '=====================================================================================
    Private Sub ColorItem(ByVal ItemGraphics As Graphics, ByVal ItemRectangle As Rectangle, ByVal ItemIndex As Integer, ByVal ColorEnum As ColorEnum, ByVal HasFocus As Boolean)
        ' draw the color entries
        Dim ItemText As String = ""
        Dim TextBrush As New SolidBrush(System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.MenuText))
        ' if we are normal style focus and we our drawing our item with focus change text color to the right color
        If m_FocusStyle = FocusStyleEnum.Normal And HasFocus = True Then TextBrush.Color = System.Drawing.Color.FromKnownColor(KnownColor.HighlightText)
        Select Case ColorEnum
            Case ColorEnum.KnownColor
                ItemText = cboColor.Items.Item(ItemIndex).ToString
            Case ColorEnum.SystemColor
                ItemText = System_Color(ItemIndex)
        End Select
        Dim ColorBrush As New SolidBrush(System.Drawing.Color.FromName(ItemText))
        Dim Pen As New Pen(System.Drawing.Color.Black, 1)
        With ItemGraphics
            .FillRectangle(ColorBrush, ItemRectangle.Left + 2, ItemRectangle.Top + 2, 20, ItemRectangle.Height - 4)
            .DrawRectangle(Pen, New Rectangle(ItemRectangle.Left + 1, ItemRectangle.Top + 1, 21, ItemRectangle.Height - 3))
            .DrawString(ItemText, cboColor.Font, TextBrush, ItemRectangle.Left + 28, ItemRectangle.Top + (ItemRectangle.Height - cboColor.Font.GetHeight()) / 2)
        End With
        TextBrush.Dispose()
        ColorBrush.Dispose()
        Pen.Dispose()
    End Sub   'ColorItem(ByVal ItemGraphics As Graphics, ByVal ItemRectangle As Rectangle, ByVal ItemIndex As Integer, ByVal ColorEnum As ColorEnum, ByVal HasFocus As Boolean)
    '=====================================================================================
    Private Sub DrawHighlight(ByVal IsSelected As Boolean, ByVal ItemGraphics As Graphics, ByVal ItemRectangle As Rectangle)
        If IsSelected = False Then
            ' draw without the highlight rectangle
            ItemGraphics.FillRectangle(New SolidBrush(SystemColors.Window), ItemRectangle.Left, ItemRectangle.Top, ItemRectangle.Width, ItemRectangle.Height)
        Else
            ' draw with the highlight rectangle
            Dim BorderPen As New Pen(System.Drawing.Color.FromKnownColor(KnownColor.Highlight))
            Dim BackgroundBrush As New SolidBrush(System.Drawing.Color.FromKnownColor(KnownColor.Highlight))
            If m_FocusStyle = FocusStyleEnum.IDE Then BackgroundBrush.Color = IDEBackgroundColor
            ItemGraphics.FillRectangle(BackgroundBrush, ItemRectangle.Left, ItemRectangle.Top, ItemRectangle.Width, ItemRectangle.Height)
            ItemGraphics.DrawRectangle(BorderPen, ItemRectangle.Left, ItemRectangle.Top, ItemRectangle.Width - 1, ItemRectangle.Height - 1)
            BorderPen.Dispose()
            BackgroundBrush.Dispose()
        End If
    End Sub   'DrawHighlight(ByVal IsSelected As Boolean, ByVal ItemGraphics As Graphics, ByVal ItemRectangle As Rectangle)
    '=====================================================================================
    Private Function CalculateIDEBackgroundColor() As System.Drawing.Color
        ' create the background color
        Dim AlphaBlend As Integer = 70
        Dim BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, System.Drawing.Color.FromKnownColor(KnownColor.Highlight))
        Dim WindowColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, System.Drawing.Color.FromKnownColor(KnownColor.Window))

        Dim RedValue As Byte = CByte(CSng(BorderColor.R * AlphaBlend / 255 + WindowColor.R * (CSng((255 - AlphaBlend) / 255))))
        Dim BlueValue As Byte = CByte(CSng(BorderColor.B * AlphaBlend / 255 + WindowColor.B * (CSng((255 - AlphaBlend) / 255))))
        Dim GreenValue As Byte = CByte(CSng(BorderColor.G * AlphaBlend / 255 + WindowColor.G * (CSng((255 - AlphaBlend) / 255))))

        CalculateIDEBackgroundColor = System.Drawing.Color.FromArgb(255, RedValue, GreenValue, BlueValue)
    End Function      'CalculateIDEBackgroundColor() As System.Drawing.Color
    '=====================================================================================
#End Region
    Private Sub ColourCombobox_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadColors(ColorEnum.KnownColor)
    End Sub
    Public Sub ForceReload()
        ColourExclude = Nothing
        ColourExclude2 = Nothing
        ColourExclude3 = Nothing
        ColourExclude4 = Nothing
        LoadColors(ColorEnum.KnownColor)
    End Sub

End Class

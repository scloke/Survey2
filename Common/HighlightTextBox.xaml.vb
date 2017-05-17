Imports System.Windows

Public Class HighlightTextBox
    Private Const InnerMarginScale As Double = 0.15
    Private Const InnerHeightScale As Double = 0.5
    Private Const InnerFontSizeScale As Double = 0.25
    Private m_WidthMultiplier As Double = 1
    Private m_FontMultiplier As Double = 1
    Public Event Click(sender As Object, e As EventArgs)
    Public Event TextChanged(sender As Object, e As Controls.TextChangedEventArgs)
    Private bTextChanging As Boolean = False

    Public Shared ReadOnly HTBReferenceHeightProperty As DependencyProperty = DependencyProperty.Register("HTBReferenceHeight", GetType(Double), GetType(HighlightTextBox), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHTBReferenceHeightChanged)))
    Public Shared ReadOnly HTBReferenceWidthProperty As DependencyProperty = DependencyProperty.Register("HTBReferenceWidth", GetType(Double), GetType(HighlightTextBox), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHTBReferenceWidthChanged)))
    Public Shared ReadOnly HTBTextProperty As DependencyProperty = DependencyProperty.Register("HTBText", GetType(String), GetType(HighlightTextBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf OnHTBTextChanged)))
    Public Shared ReadOnly HTBInnerToolTipProperty As DependencyProperty = DependencyProperty.Register("HTBInnerToolTip", GetType(String), GetType(HighlightTextBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.None, New PropertyChangedCallback(AddressOf HTBInnerToolTipChanged)))

    Public Property HTBReferenceHeight As Double
        Get
            Return CDbl(GetValue(HTBReferenceHeightProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HTBReferenceHeightProperty, value)
        End Set
    End Property
    Public Property HTBReferenceWidth As Double
        Get
            Return CDbl(GetValue(HTBReferenceWidthProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HTBReferenceWidthProperty, value)
        End Set
    End Property
    Public Property HTBText As String
        Get
            Return CDbl(GetValue(HTBTextProperty))
        End Get
        Set(ByVal value As String)
            SetValue(HTBTextProperty, value)
        End Set
    End Property
    Public Property AcceptsReturn As Boolean
        Get
            Return TextBoxMain.AcceptsReturn
        End Get
        Set(ByVal value As Boolean)
            TextBoxMain.AcceptsReturn = value
        End Set
    End Property
    Public Property WidthMultiplier As Double
        Get
            Return m_WidthMultiplier
        End Get
        Set(value As Double)
            m_WidthMultiplier = value
        End Set
    End Property
    Public Property FontMultiplier As Double
        Get
            Return m_FontMultiplier
        End Get
        Set(value As Double)
            m_FontMultiplier = value
        End Set
    End Property
    Public WriteOnly Property BorderWidth As Double
        Set(value As Double)
            TextBoxMain.BorderThickness = New Thickness(value)
        End Set
    End Property
    Public WriteOnly Property TransparentBackground As Boolean
        Set(value As Boolean)
            If value Then
                RectangleBackground.Fill = Media.Brushes.Transparent
            End If
        End Set
    End Property
    Public Property TextWrapping As TextWrapping
        Get
            Return TextBoxMain.TextWrapping
        End Get
        Set(value As TextWrapping)
            TextBoxMain.TextWrapping = value
        End Set
    End Property
    Public Property HTBInnerToolTip As String
        Get
            Return TextBoxMain.ToolTip
        End Get
        Set(value As String)
            TextBoxMain.ToolTip = value
        End Set
    End Property
    Private Shared Sub OnHTBReferenceHeightChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightTextBox As HighlightTextBox = CType(d, HighlightTextBox)
        Dim oGrid As Controls.Grid = CType(oHighlightTextBox.Content, Controls.Grid)
        Dim oRectangleBackground As Shapes.Rectangle = CType(oGrid.FindName("RectangleBackground"), Shapes.Rectangle)
        oRectangleBackground.Margin = New Thickness(e.NewValue * InnerMarginScale)

        Dim oTextBoxMain As Controls.TextBox = CType(oGrid.FindName("TextBoxMain"), Controls.TextBox)
        oTextBoxMain.Margin = New Thickness(e.NewValue * InnerMarginScale)
        If e.NewValue > 0 Then
            oTextBoxMain.FontSize = e.NewValue * oHighlightTextBox.FontMultiplier * InnerFontSizeScale
        End If
        oTextBoxMain.MinHeight = e.NewValue * InnerHeightScale
        oTextBoxMain.MinWidth = Math.Max(e.NewValue, oHighlightTextBox.HTBReferenceWidth) * InnerHeightScale * oHighlightTextBox.WidthMultiplier
    End Sub
    Private Shared Sub OnHTBReferenceWidthChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightTextBox As HighlightTextBox = CType(d, HighlightTextBox)
        Dim oGrid As Controls.Grid = CType(oHighlightTextBox.Content, Controls.Grid)
        Dim oTextBoxMain As Controls.TextBox = CType(oGrid.FindName("TextBoxMain"), Controls.TextBox)
        oTextBoxMain.MinWidth = Math.Max(oHighlightTextBox.HTBReferenceHeight, e.NewValue) * InnerHeightScale * oHighlightTextBox.WidthMultiplier
    End Sub
    Private Shared Sub HTBInnerToolTipChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightTextBox As HighlightTextBox = CType(d, HighlightTextBox)
        Dim oGrid As Controls.Grid = CType(oHighlightTextBox.Content, Controls.Grid)
        Dim oTextBoxMain As Controls.TextBox = CType(oGrid.FindName("TextBoxMain"), Controls.TextBox)
        oTextBoxMain.ToolTip = e.NewValue
    End Sub
    Private Shared Sub OnHTBTextChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        If Not IsNothing(e.NewValue) Then
            Dim sText As String = e.NewValue
            Dim oHighlightTextBox As HighlightTextBox = CType(d, HighlightTextBox)
            Dim oGrid As Controls.Grid = CType(oHighlightTextBox.Content, Controls.Grid)
            Dim oTextBoxMain As Controls.TextBox = CType(oGrid.FindName("TextBoxMain"), Controls.TextBox)

            oHighlightTextBox.bTextChanging = True
            oTextBoxMain.Text = sText
            oHighlightTextBox.bTextChanging = False
        End If
    End Sub
    Private Sub OnTextBoxTextChanged(sender As Object, e As Controls.TextChangedEventArgs) Handles TextBoxMain.TextChanged
        If Not bTextChanging Then
            Dim oTextBoxMain As Controls.TextBox = CType(sender, Controls.TextBox)
            HTBText = oTextBoxMain.Text
            RaiseEvent TextChanged(sender, e)
        End If
    End Sub
    Private Sub TextBoxMainHandler(sender As Object, e As EventArgs) Handles TextBoxMain.MouseLeftButtonDown, TextBoxMain.TouchDown
        RaiseEvent Click(sender, e)
    End Sub
    Private Sub TextBoxMainMouseMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles TextBoxMain.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            TextBoxMain.Background = New Media.SolidColorBrush(Media.Color.FromArgb(&H33, &H0, &HFF, &H0))
        End If
    End Sub
    Private Sub TextBoxMainMouseLeaveHandler(sender As Object, e As EventArgs) Handles TextBoxMain.MouseLeave, TextBoxMain.TouchLeave
        TextBoxMain.Background = New Media.SolidColorBrush(Media.Colors.Transparent)
    End Sub
End Class
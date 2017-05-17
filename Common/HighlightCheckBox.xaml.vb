Imports System.Windows
Imports System.Windows.Media

Public Class HighlightCheckBox
    Private Const ColumnWidthScale As Double = 0.3
    Private Const ImageHeightScale As Double = 0.3
    Private Const TextBlockFontSizeScale As Double = 0.4
    Public Event Click(sender As Object, e As EventArgs)

    Private Shared m_CCMCheckboxChecked As ImageSource = Common.Converter.XamlToDrawingImage(My.Resources.CCMCheckboxChecked)
    Private Shared m_CCMCheckboxUnchecked As ImageSource = Common.Converter.XamlToDrawingImage(My.Resources.CCMCheckboxUnchecked)
    Public Shared ReadOnly HCBReferenceHeightProperty As DependencyProperty = DependencyProperty.Register("HCBReferenceHeight", GetType(Double), GetType(HighlightCheckBox), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHCBReferenceHeightChanged)))
    Public Shared ReadOnly HCBCheckedProperty As DependencyProperty = DependencyProperty.Register("HCBChecked", GetType(Boolean), GetType(HighlightCheckBox), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf OnHCBCheckedChanged)))

    Public Property HCBReferenceHeight() As Double
        Get
            Return CDbl(GetValue(HCBReferenceHeightProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HCBReferenceHeightProperty, value)
        End Set
    End Property
    Public Property HCBChecked() As Boolean
        Get
            Return CDbl(GetValue(HCBCheckedProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(HCBCheckedProperty, value)
        End Set
    End Property
    Public WriteOnly Property LabelText As String
        Set(ByVal value As String)
            TextBlockMain.Text = value
            If HCBChecked Then
                ImageMain.Source = m_CCMCheckboxChecked
            Else
                ImageMain.Source = m_CCMCheckboxUnchecked
            End If
        End Set
    End Property
    Public WriteOnly Property InnerToolTip As String
        Set(value As String)
            RectangleMain.ToolTip = value
        End Set
    End Property
    Private Shared Sub OnHCBReferenceHeightChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightCheckBox As HighlightCheckBox = CType(d, HighlightCheckBox)
        Dim oGrid As Controls.Grid = CType(oHighlightCheckBox.Content, Controls.Grid)
        Dim oGridMain As Controls.Grid = CType(oGrid.FindName("GridMain"), Controls.Grid)
        oGridMain.ColumnDefinitions(0).Width = New GridLength(e.NewValue * ColumnWidthScale)
        oGridMain.ColumnDefinitions(3).Width = New GridLength(e.NewValue * ColumnWidthScale)

        Dim oImageMain As Controls.Image = CType(oGridMain.FindName("ImageMain"), Controls.Image)
        oImageMain.Height = e.NewValue * ImageHeightScale

        Dim oTextBlockMain As Controls.TextBlock = CType(oGridMain.FindName("TextBlockMain"), Controls.TextBlock)
        If e.NewValue > 0 Then
            oTextBlockMain.FontSize = e.NewValue * TextBlockFontSizeScale
        End If
    End Sub
    Private Shared Sub OnHCBCheckedChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightCheckBox As HighlightCheckBox = CType(d, HighlightCheckBox)
        Dim oGrid As Controls.Grid = CType(oHighlightCheckBox.Content, Controls.Grid)
        Dim oGridMain As Controls.Grid = CType(oGrid.FindName("GridMain"), Controls.Grid)
        Dim oImageMain As Controls.Image = CType(oGridMain.FindName("ImageMain"), Controls.Image)
        If e.NewValue Then
            oImageMain.Source = m_CCMCheckboxChecked
        Else
            oImageMain.Source = m_CCMCheckboxUnchecked
        End If
    End Sub
    Private Sub RectangleMainHandler(sender As Object, e As EventArgs) Handles RectangleMain.MouseLeftButtonDown, RectangleMain.TouchDown
        RaiseEvent Click(sender, e)
    End Sub
    Private Sub RectangleMainMouseMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles RectangleMain.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            RectangleMain.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            RectangleBackground.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub RectangleMainMainMouseLeaveHandler(sender As Object, e As EventArgs) Handles RectangleMain.MouseLeave, RectangleMain.TouchLeave
        RectangleMain.Fill = New SolidColorBrush(Colors.Transparent)
        RectangleBackground.Visibility = Visibility.Hidden
    End Sub
End Class

Imports System.Windows
Imports System.Windows.Media

Public Class HighlightButton
    Private Const InnerMarginScale As Double = 0.15
    Private Const InnerHeightScale As Double = 0.5
    Private Const InnerWidthScale As Double = 0.5
    Public Event Click(sender As Object, e As EventArgs)
    Public Event RightClick(sender As Object, e As EventArgs)
    Private Shared Event SelectedChanged(sender As Object, e As EventArgs)
    Public Shadows Event ToolTipOpening(sender As Object, e As Controls.ToolTipEventArgs)
    Public Shared ReadOnly HBReferenceHeightProperty As DependencyProperty = DependencyProperty.Register("HBReferenceHeight", GetType(Double), GetType(HighlightButton), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHBReferenceHeightChanged)))
    Public Shared ReadOnly HBReferenceWidthProperty As DependencyProperty = DependencyProperty.Register("HBReferenceWidth", GetType(Double), GetType(HighlightButton), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHBReferenceWidthChanged)))
    Public Shared ReadOnly HBReferenceProperty As DependencyProperty = DependencyProperty.Register("HBReference", GetType(Double), GetType(HighlightButton), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHBReferenceChanged)))
    Public Shared ReadOnly HBSelectedProperty As DependencyProperty = DependencyProperty.Register("HBSelected", GetType(Boolean), GetType(HighlightButton), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf OnHBSelectedChanged)))
    Public Shared ReadOnly HBSourceProperty As DependencyProperty = DependencyProperty.Register("HBSource", GetType(ImageSource), GetType(HighlightButton), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf OnHBSourceChanged)))
    Public Shared ReadOnly HBTagProperty As DependencyProperty = DependencyProperty.Register("HBTag", GetType(Object), GetType(HighlightButton), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.None, New PropertyChangedCallback(AddressOf OnHBTagChanged)))

    Public Property HBReferenceHeight() As Double
        Get
            Return CDbl(GetValue(HBReferenceHeightProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HBReferenceHeightProperty, value)
        End Set
    End Property
    Public Property HBReferenceWidth As Double
        Get
            Return CDbl(GetValue(HBReferenceWidthProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HBReferenceWidthProperty, value)
        End Set
    End Property
    Public Property HBReference As Double
        Get
            Return CDbl(GetValue(HBReferenceProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HBReferenceProperty, value)
        End Set
    End Property
    Public Property HBSelected As Boolean
        Get
            Return CDbl(GetValue(HBSelectedProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(HBSelectedProperty, value)
        End Set
    End Property
    Public Property HBSource As ImageSource
        Get
            Return CType(GetValue(HBSourceProperty), ImageSource)
        End Get
        Set(ByVal value As ImageSource)
            SetValue(HBSourceProperty, value)
        End Set
    End Property
    Public Property HBTag As Object
        Get
            Return GetValue(HBTagProperty)
        End Get
        Set(value As Object)
            SetValue(HBTagProperty, value)
        End Set
    End Property
    Public WriteOnly Property InnerToolTip As String
        Set(value As String)
            RectangleMain.ToolTip = value
        End Set
    End Property
    Private Shared Sub OnHBReferenceHeightChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightButton As HighlightButton = CType(d, HighlightButton)
        Dim oGrid As Controls.Grid = CType(oHighlightButton.Content, Controls.Grid)
        Dim oImageMain As Controls.Image = CType(oGrid.FindName("ImageMain"), Controls.Image)
        oImageMain.Margin = New Thickness(e.NewValue * InnerMarginScale)
        oImageMain.Height = e.NewValue * InnerHeightScale
    End Sub
    Private Shared Sub OnHBReferenceWidthChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightButton As HighlightButton = CType(d, HighlightButton)
        Dim oGrid As Controls.Grid = CType(oHighlightButton.Content, Controls.Grid)
        Dim oImageMain As Controls.Image = CType(oGrid.FindName("ImageMain"), Controls.Image)
        oImageMain.Margin = New Thickness(e.NewValue * InnerMarginScale)
        oImageMain.Width = e.NewValue * InnerWidthScale
    End Sub
    Private Shared Sub OnHBReferenceChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightButton As HighlightButton = CType(d, HighlightButton)
        Dim oGrid As Controls.Grid = CType(oHighlightButton.Content, Controls.Grid)
        Dim oImageMain As Controls.Image = CType(oGrid.FindName("ImageMain"), Controls.Image)
        oImageMain.Margin = New Thickness(e.NewValue * InnerMarginScale)
        oImageMain.Height = e.NewValue * InnerHeightScale
        oImageMain.Width = e.NewValue * InnerWidthScale
    End Sub
    Private Shared Sub OnHBSelectedChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        RaiseEvent SelectedChanged(Nothing, Nothing)
    End Sub
    Private Shared Sub OnHBSourceChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightButton As HighlightButton = CType(d, HighlightButton)
        Dim oGrid As Controls.Grid = CType(oHighlightButton.Content, Controls.Grid)
        Dim oImageMain As Controls.Image = CType(oGrid.FindName("ImageMain"), Controls.Image)
        oImageMain.Source = e.NewValue
    End Sub
    Private Shared Sub OnHBTagChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightButton As HighlightButton = CType(d, HighlightButton)
        Dim oGrid As Controls.Grid = CType(oHighlightButton.Content, Controls.Grid)
        Dim oRectangleMain As Shapes.Rectangle = CType(oGrid.FindName("RectangleMain"), Shapes.Rectangle)
        oRectangleMain.Tag = e.NewValue
    End Sub
    Private Sub RectangleMainLeftClickHandler(sender As Object, e As EventArgs) Handles RectangleMain.MouseLeftButtonUp, RectangleMain.TouchUp
        Select Case e.GetType
            Case GetType(Input.MouseButtonEventArgs)
                CType(e, Input.MouseButtonEventArgs).Handled = True
            Case GetType(Input.TouchEventArgs)
                CType(e, Input.TouchEventArgs).Handled = True
        End Select
        RaiseEvent Click(sender, e)
    End Sub
    Private Sub RectangleMainRightClickHandler(sender As Object, e As EventArgs) Handles RectangleMain.MouseRightButtonUp
        Select Case e.GetType
            Case GetType(Input.MouseButtonEventArgs)
                CType(e, Input.MouseButtonEventArgs).Handled = True
        End Select
        RaiseEvent RightClick(sender, e)
    End Sub
    Private Sub RectangleMainMouseMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles RectangleMain.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            RectangleMain.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            RectangleBackground.Visibility = Windows.Visibility.Visible
        End If
    End Sub
    Private Sub RectangleMainMouseLeaveHandler(sender As Object, e As EventArgs) Handles RectangleMain.MouseLeave, RectangleMain.TouchLeave, Me.SelectedChanged
        If HBSelected Then
            RectangleMain.Fill = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
            RectangleBackground.Visibility = Windows.Visibility.Visible
        Else
            RectangleMain.Fill = New SolidColorBrush(Colors.Transparent)
            RectangleBackground.Visibility = Windows.Visibility.Hidden
        End If
    End Sub
End Class

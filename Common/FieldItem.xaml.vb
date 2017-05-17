Imports System.Windows.Media
Imports System.Windows

Public Class FieldItem
    Public Shared ReadOnly FIContentProperty As DependencyProperty = DependencyProperty.Register("FIContent", GetType(Tuple(Of String, Double, ImageSource, ImageSource)), GetType(FieldItem), New FrameworkPropertyMetadata(New Tuple(Of String, Double, ImageSource, ImageSource)([String].Empty, 1, Nothing, Nothing), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnFIContentChanged)))
    Public Shared ReadOnly FITitleProperty As DependencyProperty = DependencyProperty.Register("FITitle", GetType(String), GetType(FieldItem), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnFITitleChanged)))
    Public Shared ReadOnly FIRichTextBoxRightProperty As DependencyProperty = DependencyProperty.Register("FIRichTextBoxRight", GetType(String), GetType(FieldItem), New FrameworkPropertyMetadata([String].Empty, FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnFIRichTextBoxRightChanged)))
    Public Shared ReadOnly FIReferenceProperty As DependencyProperty = DependencyProperty.Register("FIReference", GetType(Tuple(Of Double, Double)), GetType(FieldItem), New FrameworkPropertyMetadata(New Tuple(Of Double, Double)(1, 1), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnFIReferenceChanged)))

    Private m_Multiplier As Double = 1
    Private Const TitleMultiplier As Double = 0.015
    Private Const MarginMultiplierWidth As Double = 0.001
    Private Const MarginMultiplierHeight As Double = 0.002
    Private Const RichTextBoxMultiplier As Double = 0.0125

    Public Property FIContent As Tuple(Of String, Double, ImageSource, ImageSource)
        Get
            Return CType(GetValue(FIContentProperty), Tuple(Of String, Double, ImageSource, ImageSource))
        End Get
        Set(ByVal value As Tuple(Of String, Double, ImageSource, ImageSource))
            SetValue(FIContentProperty, value)
        End Set
    End Property
    Public Property FITitle As String
        Get
            Return CType(GetValue(FITitleProperty), String)
        End Get
        Set(ByVal value As String)
            SetValue(FITitleProperty, value)
        End Set
    End Property
    Public Property FIRichTextBoxRight As String
        Get
            Return CStr(GetValue(FIRichTextBoxRightProperty))
        End Get
        Set(ByVal value As String)
            SetValue(FIRichTextBoxRightProperty, value)
        End Set
    End Property
    Public Property FIReference As Tuple(Of Double, Double)
        Get
            Return CType(GetValue(FIReferenceProperty), Tuple(Of Double, Double))
        End Get
        Set(ByVal value As Tuple(Of Double, Double))
            SetValue(FIReferenceProperty, value)
        End Set
    End Property
    Private Shared Sub OnFIContentChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oNewValue As Tuple(Of String, Double, ImageSource, ImageSource) = CType(e.NewValue, Tuple(Of String, Double, ImageSource, ImageSource))
        Dim oFieldItem As FieldItem = CType(d, FieldItem)
        Dim oGrid As Controls.Grid = CType(oFieldItem.Content, Controls.Grid)
        Dim oGridRight As Controls.Grid = CType(oGrid.FindName("GridRight"), Controls.Grid)

        Dim oTextBlockTitle As Controls.TextBlock = CType(oGrid.FindName("Title"), Controls.TextBlock)
        oTextBlockTitle.Text = oNewValue.Item1

        oFieldItem.m_Multiplier = oNewValue.Item2

        Dim oViewBox As Controls.Viewbox = CType(oGrid.FindName("ViewBoxLeft"), Controls.Viewbox)
        Dim oImageLeft As Controls.Image = CType(oViewBox.FindName("ImageLeft"), Controls.Image)
        oImageLeft.Source = oNewValue.Item3

        Dim oImageRight As Controls.Image = CType(oGridRight.FindName("ImageRight"), Controls.Image)
        If IsNothing(oNewValue.Item4) Then
            oImageRight.Visibility = Visibility.Collapsed
        Else
            oImageRight.Visibility = Visibility.Visible
            oImageRight.Source = oNewValue.Item4
        End If
    End Sub
    Private Shared Sub OnFITitleChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim sNewValue As String = CType(e.NewValue, String)
        Dim oFieldItem As FieldItem = CType(d, FieldItem)
        Dim oGrid As Controls.Grid = CType(oFieldItem.Content, Controls.Grid)
        Dim oGridRight As Controls.Grid = CType(oGrid.FindName("GridRight"), Controls.Grid)

        Dim oRichTextBoxRight As Controls.RichTextBox = CType(oGridRight.FindName("RichTextBoxRight"), Controls.RichTextBox)
        If sNewValue = String.Empty Then
            oRichTextBoxRight.Visibility = Visibility.Collapsed
        Else
            oRichTextBoxRight.Visibility = Visibility.Visible
            oRichTextBoxRight.SelectAll()
            oRichTextBoxRight.Selection.Text = sNewValue
        End If
    End Sub
    Private Shared Sub OnFIRichTextBoxRightChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oFieldItem As FieldItem = CType(d, FieldItem)
        Dim oGrid As Controls.Grid = CType(oFieldItem.Content, Controls.Grid)
        Dim oGridRight As Controls.Grid = CType(oGrid.FindName("GridRight"), Controls.Grid)
        Dim oRichTextBoxRight As Controls.RichTextBox = CType(oGridRight.FindName("RichTextBoxRight"), Controls.RichTextBox)

        If e.NewValue = String.Empty Then
            oRichTextBoxRight.Visibility = Visibility.Collapsed
        Else
            oRichTextBoxRight.Visibility = Visibility.Visible
            oRichTextBoxRight.SelectAll()
            oRichTextBoxRight.Selection.Text = e.NewValue
        End If
    End Sub
    Private Shared Sub OnFIReferenceChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oFieldItem As FieldItem = CType(d, FieldItem)
        If oFieldItem.m_Multiplier > 0 And CType(e.NewValue, Tuple(Of Double, Double)).Item1 > 0 And CType(e.NewValue, Tuple(Of Double, Double)).Item2 > 0 Then
            Dim oReference As Tuple(Of Double, Double) = e.NewValue

            Dim fMarginWidth As Double = oReference.Item1 * MarginMultiplierWidth
            Dim fMarginHeight As Double = oReference.Item2 * MarginMultiplierHeight

            Dim oGrid As Controls.Grid = CType(oFieldItem.Content, Controls.Grid)
            Dim oViewBoxLeft As Controls.Viewbox = CType(oGrid.FindName("ViewBoxLeft"), Controls.Viewbox)
            oViewBoxLeft.Margin = New Thickness(5 * fMarginWidth, 5 * fMarginHeight, 0, 5 * fMarginHeight)

            Dim oTextBlockTitle As Controls.TextBlock = CType(oGrid.FindName("Title"), Controls.TextBlock)
            oTextBlockTitle.Margin = New Thickness(5 * fMarginWidth, 5 * fMarginHeight, 5 * fMarginWidth, 2 * fMarginHeight)
            oTextBlockTitle.FontSize = oReference.Item2 * TitleMultiplier * oFieldItem.m_Multiplier

            Dim oGridRight As Controls.Grid = CType(oGrid.FindName("GridRight"), Controls.Grid)
            oGridRight.Margin = New Thickness(5 * fMarginWidth, 5 * fMarginHeight, 5 * fMarginWidth, 10 * fMarginHeight)

            Dim oRichTextBoxRight As Controls.RichTextBox = CType(oGridRight.FindName("RichTextBoxRight"), Controls.RichTextBox)
            oRichTextBoxRight.FontSize = oReference.Item2 * RichTextBoxMultiplier * oFieldItem.m_Multiplier
        End If
    End Sub
End Class

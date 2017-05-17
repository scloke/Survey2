Imports System.Collections.ObjectModel
Imports System.Windows

Public Class HighlightComboBox
    Private Const InnerMarginScale As Double = 0.15
    Private Const InnerHeightScale As Double = 0.5
    Private Const InnerFontSizeScale As Double = 0.25
    Private m_WidthMultiplier As Double = 1
    Private m_FontMultiplier As Double = 1
    Public Event SelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs)
    Public Shared ReadOnly HCBReferenceHeightProperty As DependencyProperty = DependencyProperty.Register("HCBReferenceHeight", GetType(Double), GetType(HighlightComboBox), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHCBReferenceHeightChanged)))
    Public Shared ReadOnly HCBReferenceWidthProperty As DependencyProperty = DependencyProperty.Register("HCBReferenceWidth", GetType(Double), GetType(HighlightComboBox), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure, New PropertyChangedCallback(AddressOf OnHCBReferenceWidthChanged)))
    Public Shared ReadOnly HCBTextProperty As DependencyProperty = DependencyProperty.Register("HCBText", GetType(HCBDisplay), GetType(HighlightComboBox), New FrameworkPropertyMetadata(New HCBDisplay, FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf OnHCBTextChanged)))
    Public Shared ReadOnly HCBContentProperty As DependencyProperty = DependencyProperty.Register("HCBContent", GetType(ObservableCollection(Of HCBDisplay)), GetType(HighlightComboBox), New FrameworkPropertyMetadata(New ObservableCollection(Of HCBDisplay), FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf OnHCBContentChanged)))

    Public Property HCBReferenceHeight As Double
        Get
            Return CDbl(GetValue(HCBReferenceHeightProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HCBReferenceHeightProperty, value)
        End Set
    End Property
    Public Property HCBReferenceWidth As Double
        Get
            Return CDbl(GetValue(HCBReferenceWidthProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HCBReferenceWidthProperty, value)
        End Set
    End Property
    Public WriteOnly Property HCBText As HCBDisplay
        Set(ByVal value As HCBDisplay)
        End Set
    End Property
    Public WriteOnly Property HCBContent As ObservableCollection(Of HCBDisplay)
        Set(value As ObservableCollection(Of HCBDisplay))
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
    Public Property IsEditable As Boolean
        Get
            Return ComboBoxMain.IsEditable
        End Get
        Set(value As Boolean)
            ComboBoxMain.IsEditable = value
        End Set
    End Property
    Public Property IsReadOnly As Boolean
        Get
            Return ComboBoxMain.IsReadOnly
        End Get
        Set(value As Boolean)
            ComboBoxMain.IsReadOnly = value
        End Set
    End Property
    Public WriteOnly Property InnerToolTip As String
        Set(value As String)
            ComboBoxMain.ToolTip = value
        End Set
    End Property
    Private Shared Sub OnHCBReferenceHeightChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightComboBox As HighlightComboBox = CType(d, HighlightComboBox)
        Dim oGrid As Controls.Grid = CType(oHighlightComboBox.Content, Controls.Grid)
        Dim oRectangleBackground As Shapes.Rectangle = CType(oGrid.FindName("RectangleBackground"), Shapes.Rectangle)
        oRectangleBackground.Margin = New Thickness(e.NewValue * InnerMarginScale)

        Dim oComboBoxMain As Controls.ComboBox = CType(oGrid.FindName("ComboBoxMain"), Controls.ComboBox)
        oComboBoxMain.Margin = New Thickness(e.NewValue * InnerMarginScale)
        If e.NewValue > 0 Then
            oComboBoxMain.FontSize = e.NewValue * oHighlightComboBox.FontMultiplier * InnerFontSizeScale
        End If
        oComboBoxMain.MinHeight = e.NewValue * InnerHeightScale
        oComboBoxMain.MinWidth = Math.Max(e.NewValue, oHighlightComboBox.HCBReferenceWidth) * InnerHeightScale * oHighlightComboBox.WidthMultiplier
    End Sub
    Private Shared Sub OnHCBReferenceWidthChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightComboBox As HighlightComboBox = CType(d, HighlightComboBox)
        Dim oGrid As Controls.Grid = CType(oHighlightComboBox.Content, Controls.Grid)
        Dim oComboBoxMain As Controls.ComboBox = CType(oGrid.FindName("ComboBoxMain"), Controls.ComboBox)
        oComboBoxMain.MinWidth = Math.Max(oHighlightComboBox.HCBReferenceHeight, e.NewValue) * InnerHeightScale * oHighlightComboBox.WidthMultiplier
    End Sub
    Private Shared Sub OnHCBTextChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        If Not IsNothing(e.NewValue) Then
            Dim oHCBDisplay As HCBDisplay = e.NewValue
            Dim oHighlightComboBox As HighlightComboBox = CType(d, HighlightComboBox)
            Dim oGrid As Controls.Grid = CType(oHighlightComboBox.Content, Controls.Grid)
            Dim oComboBoxMain As Controls.ComboBox = CType(oGrid.FindName("ComboBoxMain"), Controls.ComboBox)
            If (Not IsNothing(oComboBoxMain.ItemsSource)) AndAlso (IsNothing(oComboBoxMain.SelectedItem) OrElse IsNothing(oComboBoxMain.SelectedItem) OrElse (Not oComboBoxMain.SelectedItem.Equals(e.NewValue))) Then
                Dim oHCBDisplayCollection As ObservableCollection(Of HCBDisplay) = oComboBoxMain.ItemsSource
                Dim iIndexList As List(Of Integer) = (From iIndex As Integer In Enumerable.Range(0, oHCBDisplayCollection.Count) Where oHCBDisplayCollection(iIndex).Name = oHCBDisplay.Name Select iIndex).ToList
                If iIndexList.Count > 0 Then
                    oComboBoxMain.SelectedIndex = iIndexList.First
                End If
            End If
        End If
    End Sub
    Private Shared Sub OnHCBContentChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oHighlightComboBox As HighlightComboBox = CType(d, HighlightComboBox)
        Dim oGrid As Controls.Grid = CType(oHighlightComboBox.Content, Controls.Grid)
        Dim oComboBoxMain As Controls.ComboBox = CType(oGrid.FindName("ComboBoxMain"), Controls.ComboBox)
        If IsNothing(oComboBoxMain.ItemsSource) OrElse IsNothing(e.NewValue) OrElse oComboBoxMain.ItemsSource.GetHashCode <> e.NewValue.GetHashCode Then
            oComboBoxMain.ItemsSource = e.NewValue
        End If
    End Sub
    Private Sub OnHCBSelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs) Handles ComboBoxMain.SelectionChanged
        Dim oComboBoxMain As Controls.ComboBox = CType(sender, Controls.ComboBox)
        If Not IsNothing(oComboBoxMain.SelectedItem) AndAlso CType(oComboBoxMain.SelectedItem, HCBDisplay).Name <> oComboBoxMain.Text Then
            RaiseEvent SelectionChanged(sender, e)
        End If
    End Sub
    Private Sub ComboBoxMainMouseMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles ComboBoxMain.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            ComboBoxMain.Background = New Media.SolidColorBrush(Media.Color.FromArgb(&H33, &H0, &HFF, &H0))
        End If
    End Sub
    Private Sub ComboBoxMainMouseLeaveHandler(sender As Object, e As EventArgs) Handles ComboBoxMain.MouseLeave, ComboBoxMain.TouchLeave
        ComboBoxMain.Background = New Media.SolidColorBrush(Media.Colors.Transparent)
    End Sub
    Public Structure HCBDisplay
        Private m_Name As String
        Private m_Highlight As Boolean
        Private m_Colour As String
        Private m_GUID As Guid

        Sub New(ByVal sName As String, ByVal oGUID As Guid, ByVal bHighlight As Boolean, Optional ByVal sColour As String = "Red")
            m_Name = sName
            m_GUID = oGUID
            m_Highlight = bHighlight
            m_Colour = sColour
        End Sub
        Public Property Name As String
            Get
                Return m_Name
            End Get
            Set(value As String)
                m_Name = value
            End Set
        End Property
        Public Property GUID As Guid
            Get
                Return m_GUID
            End Get
            Set(value As Guid)
                m_GUID = value
            End Set
        End Property
        Public Property Highlight As Boolean
            Get
                Return m_Highlight
            End Get
            Set(value As Boolean)
                m_Highlight = value
            End Set
        End Property
        Public Property Colour As String
            Get
                Return m_Colour
            End Get
            Set(value As String)
                m_Colour = value
            End Set
        End Property
    End Structure
End Class

Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports Common.Common
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Media

Public Class PDFViewer
    Implements INotifyPropertyChanged

    Public Const InitialViewWidth As Single = 0.8
    Public Shared ReadOnly PVReferenceHeightProperty As DependencyProperty = DependencyProperty.Register("PVReferenceHeight", GetType(Double), GetType(PDFViewer), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure))
    Public Shared ReadOnly PVReferenceWidthProperty As DependencyProperty = DependencyProperty.Register("PVReferenceWidth", GetType(Double), GetType(PDFViewer), New FrameworkPropertyMetadata(CDbl(0), FrameworkPropertyMetadataOptions.AffectsMeasure))
    Public Shared ReadOnly PVTagProperty As DependencyProperty = DependencyProperty.Register("PVTag", GetType(Object), GetType(PDFViewer), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.None, New PropertyChangedCallback(AddressOf OnPVTagChanged)))
    Private ElementDictionary As Dictionary(Of String, FrameworkElement)
    Private Shared LastTouched As Date = Date.MaxValue
    Private Shared LastTouchedID As Integer = Integer.MinValue
    Private Shared MinTouchInterval As New TimeSpan(0, 0, 0, 0, 250)
    Private m_PDFDocument As PdfSharp.Pdf.PdfDocument
    Private m_Page As Integer = 0
    Private m_SubjectNameText As String = String.Empty
    Private m_SubjectCurrentText As String = String.Empty
    Private m_SubjectCountText As String = String.Empty
    Private m_DefaultSave As String = String.Empty
    Private m_PDFDocumentBytes As Byte()
    Private m_CurrentPage As Integer = 0
    Private m_Images As New ObservableCollection(Of ImageSource)
    Private m_ActiveRectangles As New Dictionary(Of Guid, Tuple(Of Integer, Rect))
    Private m_RectangleStore As New Dictionary(Of Guid, Tuple(Of Integer, Rect))
    Public Event ReturnMessage(ByVal oColour As Media.Color, ByVal oDateTime As Date, ByVal sMessage As String)
    Public Event DirectSelect(ByVal oGUID As Guid)

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChangedLocal(ByVal sName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
    End Sub
    Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ElementDictionary = New Dictionary(Of String, FrameworkElement)
        CommonFunctions.IterateElementDictionary(ElementDictionary, GridMain)

        ScrollViewerPDFHost.Tag = InitialViewWidth
        SetIcons()

        PDFPage.DataContext = Me
        PDFHost.DataContext = Me

        Dim oBindingPDF1 As New Data.Binding
        oBindingPDF1.Path = New PropertyPath("CurrentPageText")
        oBindingPDF1.Mode = Data.BindingMode.TwoWay
        oBindingPDF1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
        PDFPage.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingPDF1)

        PDFPrint.InnerToolTip = "Print PDF File - " + (New System.Drawing.Printing.PrinterSettings).PrinterName
    End Sub
    Public Property PDFDocument As PdfSharp.Pdf.PdfDocument
        Get
            Return m_PDFDocument
        End Get
        Set(value As PdfSharp.Pdf.PdfDocument)
            If Not IsNothing(m_PDFDocument) Then
                m_PDFDocument.Dispose()
            End If
            m_PDFDocument = value
        End Set
    End Property
    Public Property DefaultSave As String
        Get
            Return m_DefaultSave
        End Get
        Set(ByVal value As String)
            m_DefaultSave = value
        End Set
    End Property
    Public Property CurrentPageText As String
        Get
            If m_CurrentPage = -1 Then
                Return String.Empty
            Else
                Return (m_CurrentPage + 1).ToString
            End If
        End Get
        Set(value As String)
            CurrentPage = CInt(Val(value)) - 1
        End Set
    End Property
    Public Property CurrentPage As Integer
        Get
            Return m_CurrentPage
        End Get
        Set(value As Integer)
            If IsNothing(PDFDocument) Then
                m_CurrentPage = -1
            Else
                m_CurrentPage = value
                m_CurrentPage = Math.Max(m_CurrentPage, 0)
                m_CurrentPage = Math.Min(m_CurrentPage, PDFDocument.PageCount - 1)
            End If

            ' bring into focus
            If (Not IsNothing(m_Images)) AndAlso m_Images.Count > 0 Then
                Dim oImageList As New List(Of Controls.Image)

                For i = 0 To PDFHost.Items.Count - 1
                    Dim oContentPresenter As Controls.ContentPresenter = PDFHost.ItemContainerGenerator.ContainerFromIndex(i)
                    oContentPresenter.ApplyTemplate()
                    Dim oImage As Controls.Image = oContentPresenter.ContentTemplate.FindName("ImagePage", oContentPresenter)
                    oImageList.Add(oImage)

                    ' activates the overlay rectangles for each page
                    Dim oCanvas As Controls.Canvas = oContentPresenter.ContentTemplate.FindName("CanvasPage", oContentPresenter)
                    oCanvas.Children.Clear()

                    Dim iCurrentPage As Integer = i
                    Dim oPageRectangles As List(Of Rect) = (From oRectangle In m_ActiveRectangles Where oRectangle.Value.Item1 = iCurrentPage Select oRectangle.Value.Item2).ToList
                    If oPageRectangles.Count > 0 Then
                        oImage.InvalidateArrange()
                        oImage.UpdateLayout()

                        Dim fScaleX As Double = oImage.ActualWidth
                        Dim fScaleY As Double = oImage.ActualHeight
                        Dim oRectangle As New Shapes.Rectangle
                        With oRectangle
                            .Width = (oPageRectangles(0).Width * fScaleX) + (SpacingLarge * 2)
                            .Height = (oPageRectangles(0).Height * fScaleY) + (SpacingLarge * 2)
                            .Stroke = New SolidColorBrush(Color.FromArgb(&HFF, &HFF, &H0, &H0))
                            .StrokeThickness = 2
                            .StrokeDashArray = New DoubleCollection From {4}
                            .Fill = Brushes.Transparent
                            .Visibility = Visibility.Visible
                        End With
                        Controls.Canvas.SetLeft(oRectangle, oImage.Margin.Left + ((PDFHost.ActualWidth - oImage.ActualWidth) / 2) + (oPageRectangles(0).X * fScaleX) - SpacingLarge)
                        Controls.Canvas.SetTop(oRectangle, oImage.Margin.Top + (oPageRectangles(0).Y * fScaleY) - SpacingLarge)
                        Controls.Canvas.SetZIndex(oRectangle, 3)
                        oCanvas.Children.Add(oRectangle)
                    End If
                Next

                If oImageList.Count > 0 Then
                    oImageList(Math.Min(m_CurrentPage, oImageList.Count - 1)).BringIntoView()
                End If
            End If
            OnPropertyChangedLocal("CurrentPageText")
            OnPropertyChangedLocal("CurrentPage")
        End Set
    End Property
    Public ReadOnly Property PrintDocument As System.Drawing.Printing.PrintDocument
        Get
            Dim oReturnDocument As System.Drawing.Printing.PrintDocument = Nothing
            If Not IsNothing(m_PDFDocumentBytes) Then
                Using oMemoryStream As New IO.MemoryStream(m_PDFDocumentBytes)
                    Dim oDocument = PdfiumViewer.PdfDocument.Load(oMemoryStream)
                    If Not IsNothing(oDocument) Then
                        oReturnDocument = oDocument.CreatePrintDocument()
                    End If
                End Using
            End If
            Return oReturnDocument
        End Get
    End Property
    Public ReadOnly Property Images As ObservableCollection(Of ImageSource)
        Get
            Return m_Images
        End Get
    End Property
    Public Property PVReferenceHeight As Double
        Get
            Return CDbl(GetValue(PVReferenceHeightProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(PVReferenceHeightProperty, value)
        End Set
    End Property
    Public Property PVReferenceWidth As Double
        Get
            Return CDbl(GetValue(PVReferenceWidthProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(PVReferenceWidthProperty, value)
        End Set
    End Property
    Public Property PVTag As Object
        Get
            Return GetValue(PVTagProperty)
        End Get
        Set(value As Object)
            SetValue(PVTagProperty, value)
        End Set
    End Property
    Public Sub SetPDFDocument(ByVal oPDFDocument As PdfSharp.Pdf.PdfDocument, Optional ByVal oDocumentRectangles As List(Of Tuple(Of Guid, Integer, Rect)) = Nothing, Optional ByVal oPageSize As PdfSharp.PageSize = PdfSharp.PageSize.A4)
        If (Not IsNothing(PDFDocument)) AndAlso (Not IsNothing(oPDFDocument)) AndAlso PDFDocument.Guid.Equals(oPDFDocument.Guid) Then
            PDFDocument.Close()
            PDFDocument.Dispose()
        End If

        PDFDocument = oPDFDocument
        If IsNothing(PDFDocument) Then
            m_CurrentPage = -1
        Else
            m_CurrentPage = 0
        End If

        m_ActiveRectangles.Clear()
        m_RectangleStore.Clear()
        If Not IsNothing(oDocumentRectangles) Then
            For Each oRectangle As Tuple(Of Guid, Integer, Rect) In oDocumentRectangles
                m_RectangleStore.Add(oRectangle.Item1, New Tuple(Of Integer, Rect)(oRectangle.Item2, ConvertRectFraction(oRectangle.Item3, oPageSize)))
            Next
        End If

        SetImages()
    End Sub
    Public Sub SetActiveRectangles(ByVal oActiveRectangles As List(Of Tuple(Of Guid, Integer, Rect)), Optional ByVal oPageSize As PdfSharp.PageSize = PdfSharp.PageSize.A4)
        ' sets the active rectangles for each page
        m_ActiveRectangles.Clear()
        For Each oRectangle As Tuple(Of Guid, Integer, Rect) In oActiveRectangles
            If (Not m_ActiveRectangles.ContainsKey(oRectangle.Item1)) AndAlso m_RectangleStore.ContainsKey(oRectangle.Item1) Then
                m_ActiveRectangles.Add(oRectangle.Item1, New Tuple(Of Integer, Rect)(oRectangle.Item2, ConvertRectFraction(oRectangle.Item3, oPageSize)))
            End If
        Next
    End Sub
    Public Sub Update()
        m_PDFDocumentBytes = Nothing
        m_Images.Clear()

        If IsNothing(PDFDocument) Then
            m_CurrentPage = -1
        Else
            m_CurrentPage = 0
        End If
        SetImages()

        OnPropertyChangedLocal("CurrentPageText")
        OnPropertyChangedLocal("CurrentPage")
        OnPropertyChangedLocal("Images")
    End Sub
    Public Sub FullPageView()
        ' reset to full page view
        If Images.Count > 0 Then
            ScrollViewerPDFHost.Tag = (ScrollViewerPDFHost.ActualHeight * Images(0).Width) / (ScrollViewerPDFHost.ActualWidth * Images(0).Height)
        End If
    End Sub
    Public Sub SetVisibility(ByVal oControlList As List(Of String), ByVal oVisibility As Visibility)
        ' sets the visibility of named controls
        For Each sControl As String In oControlList
            If ElementDictionary.ContainsKey(sControl) Then
                ElementDictionary(sControl).Visibility = oVisibility
            End If
        Next
    End Sub
    Private Sub SetIcons()
        PDFSave.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMPDF)
        PDFPrint.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMPrint)
        PDFPrintSelect.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMPrintSelect)
        PDFZoomIn.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMZoomIn)
        PDFZoomOut.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMZoomOut)
        PDFZoom100.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMZoom100)
        PDFZoomPage.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMZoomPage)
        PDFStart.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMStart)
        PDFBack.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMBack)
        PDFForward.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMForward)
        PDFEnd.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMEnd)
    End Sub
    Private Sub SetImages()
        m_Images.Clear()

        If (Not IsNothing(PDFDocument)) AndAlso PDFDocument.PageCount > 0 Then
            Dim oDocumentImages As Tuple(Of Byte(), List(Of System.Drawing.Bitmap)) = DocumentToImages(PDFDocument, ViewResolution150)

            If Not IsNothing(oDocumentImages) Then
                m_PDFDocumentBytes = oDocumentImages.Item1

                For Each oBitmap As System.Drawing.Bitmap In oDocumentImages.Item2
                    Dim oBitmapSource As Imaging.BitmapSource = Converter.BitmapToBitmapSource(oBitmap, True)
                    m_Images.Add(oBitmapSource)
                    oBitmap.Dispose()
                    oBitmap = Nothing
                Next
            End If
        Else
            m_PDFDocumentBytes = Nothing
        End If
    End Sub
    Private Shared Function ConvertRectFraction(ByVal oRect As Rect, ByVal oPageSize As PdfSharp.PageSize) As Rect
        ' converts a rect in points to a fraction of one
        Dim oXPageSize As PdfSharp.Drawing.XSize = PdfSharp.PageSizeConverter.ToSize(oPageSize)
        Return New Rect(oRect.X / oXPageSize.Width, oRect.Y / oXPageSize.Height, oRect.Width / oXPageSize.Width, oRect.Height / oXPageSize.Height)
    End Function
    Private Shared Sub CheckDoubleTouch(sender As Object, e As RoutedEventArgs, ByVal oAction As Action, ByVal oMinInterval As TimeSpan)
        ' ignore stylus events as touch triggers both
        If (Not e.RoutedEvent.Equals(Input.Stylus.StylusDownEvent)) Then
            Dim eTouch As Input.TouchEventArgs = e
            If e.RoutedEvent.Equals(TouchDownEvent) Then
                If LastTouchedID = Integer.MinValue Then
                    ' first touchdown
                    LastTouchedID = eTouch.TouchDevice.Id
                    LastTouched = Date.MaxValue
                ElseIf LastTouchedID <> Integer.MinValue AndAlso LastTouched <> Date.MaxValue Then
                    ' second touchdown
                    If TimeSpan.Compare(Date.Now - LastTouched, oMinInterval) < 0 Then
                        oAction.Invoke()
                    End If

                    ' reset values
                    LastTouchedID = eTouch.TouchDevice.Id
                    LastTouched = Date.MaxValue
                End If
            ElseIf e.RoutedEvent.Equals(TouchUpEvent) Then
                If LastTouchedID = eTouch.TouchDevice.Id Then
                    ' first touchup
                    LastTouched = Date.Now
                End If
            End If
        End If
    End Sub
    Public Shared Function DocumentToImages(ByRef oPDFDocument As PdfSharp.Pdf.PdfDocument, ByVal fResolution As Single) As Tuple(Of Byte(), List(Of System.Drawing.Bitmap))
        ' renders a PDF document into a list of bitmaps
        If oPDFDocument.PageCount > 0 Then
            Dim oPDFDocumentBytes As Byte()
            Using oMemoryStream As New IO.MemoryStream
                oPDFDocument.Save(oMemoryStream)
                oPDFDocumentBytes = oMemoryStream.ToArray()
                oPDFDocument = PdfSharp.Pdf.IO.PdfReader.Open(New IO.MemoryStream(oPDFDocumentBytes))
            End Using

            ' repeat the process until it is complete or three successive tries unsuccesful
            Dim oBitmaps As New List(Of System.Drawing.Bitmap)
            Dim iRepeatCount As Integer = 0
            Do Until iRepeatCount > 3
                Try
                    For Each oBitmap In oBitmaps
                        oBitmap.Dispose()
                    Next
                    oBitmaps.Clear()
                    CommonFunctions.ClearMemory()

                    Using oMemoryStream As New IO.MemoryStream(oPDFDocumentBytes)
                        Using oDocument = PdfiumViewer.PdfDocument.Load(oMemoryStream)
                            If Not IsNothing(oDocument) Then
                                For i = 0 To oDocument.PageCount - 1
                                    Dim iWidth As Integer = CInt(Math.Ceiling(PdfSharp.Drawing.XUnit.FromPoint(oDocument.PageSizes(i).Width).Inch * fResolution))
                                    Dim iHeight As Integer = CInt(Math.Ceiling(PdfSharp.Drawing.XUnit.FromPoint(oDocument.PageSizes(i).Height).Inch * fResolution))

                                    Dim oBitmap As System.Drawing.Bitmap = oDocument.Render(i, iWidth, iHeight, fResolution, fResolution, False)
                                    oBitmaps.Add(oBitmap)
                                Next
                            End If
                        End Using
                    End Using

                    Return New Tuple(Of Byte(), List(Of System.Drawing.Bitmap))(oPDFDocumentBytes, oBitmaps)
                Catch ex As Exception
                    ' no success
                    iRepeatCount += 1
                End Try
            Loop

            ' if not successful, clean up and return nothing
            For Each oBitmap In oBitmaps
                oBitmap.Dispose()
            Next
            oBitmaps.Clear()
            CommonFunctions.ClearMemory()

            Return Nothing
        Else
            Return Nothing
        End If
    End Function
    Private Function DirectSelectSub(ByVal e As EventArgs) As Boolean
        ' checks for selection of an input field and returns true if one is found
        Dim bReturn As Boolean = False
        Dim oLocation As Tuple(Of Integer, Point) = Nothing

        For i = 0 To PDFHost.Items.Count - 1
            Dim oContentPresenter As Controls.ContentPresenter = PDFHost.ItemContainerGenerator.ContainerFromIndex(i)
            oContentPresenter.ApplyTemplate()
            Dim oImage As Controls.Image = oContentPresenter.ContentTemplate.FindName("ImagePage", oContentPresenter)

            Dim oPoint As Point = Nothing
            If e.GetType.Equals(GetType(Input.MouseButtonEventArgs)) Then
                oPoint = CType(e, Input.MouseButtonEventArgs).GetPosition(oImage)
            Else
                oPoint = CType(e, Input.TouchEventArgs).GetTouchPoint(oImage).Position
            End If

            If oPoint.Y < 0 Then
                Exit For
            Else
                oLocation = New Tuple(Of Integer, Point)(i + 1, New Point(oPoint.X / oImage.ActualWidth, oPoint.Y / oImage.ActualHeight))
            End If
        Next

        If Not IsNothing(oLocation) Then
            Dim oGUIDList As List(Of Guid) = (From oGUID In m_RectangleStore.Keys Where oLocation.Item1 = m_RectangleStore(oGUID).Item1 AndAlso oLocation.Item2.X >= m_RectangleStore(oGUID).Item2.Left AndAlso oLocation.Item2.X <= m_RectangleStore(oGUID).Item2.Right AndAlso oLocation.Item2.Y >= m_RectangleStore(oGUID).Item2.Top AndAlso oLocation.Item2.Y <= m_RectangleStore(oGUID).Item2.Bottom Select oGUID).ToList
            If oGUIDList.Count > 0 Then
                bReturn = True
                RaiseEvent DirectSelect(oGUIDList(0))
            End If
        End If
        Return bReturn
    End Function
    Private Shared Sub OnPVTagChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim oPDFViewer As PDFViewer = CType(d, PDFViewer)
        Dim oGrid As Controls.Grid = CType(oPDFViewer.Content, Controls.Grid)
        Dim oRectanglePDFHost As Shapes.Rectangle = CType(oGrid.FindName("RectanglePDFHost"), Shapes.Rectangle)
        oRectanglePDFHost.Tag = e.NewValue
    End Sub
#Region "Buttons"
    Private Sub PDFSaveHandler(sender As Object, e As EventArgs) Handles PDFSave.Click
        Dim oAction As Action = Sub()
                                    Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
                                    oSaveFileDialog.FileName = String.Empty
                                    oSaveFileDialog.DefaultExt = ".pdf"
                                    oSaveFileDialog.Filter = "PDF File|*.pdf|All Files|*.*"
                                    oSaveFileDialog.Title = "Save PDF File"
                                    oSaveFileDialog.InitialDirectory = DefaultSave
                                    Dim result? As Boolean = oSaveFileDialog.ShowDialog()
                                    If result = True Then
                                        If IO.File.Exists(oSaveFileDialog.FileName) Then
                                            IO.File.Delete(oSaveFileDialog.FileName)
                                        End If

                                        PDFDocument.Save(oSaveFileDialog.FileName)
                                        PDFDocument.Close()
                                        PDFDocument.Dispose()
                                        PDFDocument = Nothing
                                        PDFDocument = PdfSharp.Pdf.IO.PdfReader.Open(oSaveFileDialog.FileName)

                                        RaiseEvent ReturnMessage(Colors.Green, Date.Now, "PDF file saved.")
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFPrintHandler(sender As Object, e As EventArgs) Handles PDFPrint.Click
        Dim oAction As Action = Sub()
                                    Dim oPrintDocument As System.Drawing.Printing.PrintDocument = PrintDocument
                                    oPrintDocument.DocumentName = ModuleName + " - " + PDFDocument.Info.Title
                                    oPrintDocument.Print()

                                    RaiseEvent ReturnMessage(Colors.Green, Date.Now, "PDF file printed to " + oPrintDocument.PrinterSettings.PrinterName + ".")
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFPrintSelectHandler(sender As Object, e As EventArgs) Handles PDFPrintSelect.Click
        Dim oAction As Action = Sub()
                                    Dim oPrintDialog As New Controls.PrintDialog()
                                    oPrintDialog.PageRangeSelection = Controls.PageRangeSelection.SelectedPages
                                    oPrintDialog.MinPage = 1
                                    oPrintDialog.MaxPage = PDFDocument.PageCount
                                    oPrintDialog.UserPageRangeEnabled = True

                                    Dim result? As Boolean = oPrintDialog.ShowDialog()
                                    If result = True Then
                                        Dim oPrintDocument As System.Drawing.Printing.PrintDocument = PrintDocument
                                        oPrintDocument.DocumentName = ModuleName + " - " + PDFDocument.Info.Title
                                        oPrintDocument.PrinterSettings.PrinterName = oPrintDialog.PrintQueue.FullName
                                        oPrintDocument.PrinterSettings.Copies = oPrintDialog.PrintTicket.CopyCount
                                        oPrintDocument.PrinterSettings.FromPage = oPrintDialog.PageRange.PageFrom
                                        oPrintDocument.PrinterSettings.ToPage = oPrintDialog.PageRange.PageTo
                                        oPrintDocument.Print()

                                        RaiseEvent ReturnMessage(Colors.Green, Date.Now, "PDF file printed to " + oPrintDocument.PrinterSettings.PrinterName + ".")
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFZoomInHandler(sender As Object, e As EventArgs) Handles PDFZoomIn.Click
        Dim oAction As Action = Sub()
                                    ScrollViewerPDFHost.Tag = CDbl(Val(ScrollViewerPDFHost.Tag)) * 1.25
                                    CurrentPage = CurrentPage
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFZoomOutHandler(sender As Object, e As EventArgs) Handles PDFZoomOut.Click
        Dim oAction As Action = Sub()
                                    ScrollViewerPDFHost.Tag = CDbl(Val(ScrollViewerPDFHost.Tag)) / 1.25
                                    CurrentPage = CurrentPage
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFZoom100Handler(sender As Object, e As EventArgs) Handles PDFZoom100.Click
        Dim oAction As Action = Sub()
                                    ScrollViewerPDFHost.Tag = InitialViewWidth
                                    CurrentPage = CurrentPage
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFZoomPageHandler(sender As Object, e As EventArgs) Handles PDFZoomPage.Click
        Dim oAction As Action = Sub()
                                    If Images.Count > 0 Then
                                        FullPageView()
                                        CurrentPage = CurrentPage
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFStartHandler(sender As Object, e As EventArgs) Handles PDFStart.Click
        Dim oAction As Action = Sub()
                                    If IsNothing(PDFDocument) Then
                                        CurrentPage = -1
                                    Else
                                        CurrentPage = 0
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFBackHandler(sender As Object, e As EventArgs) Handles PDFBack.Click
        Dim oAction As Action = Sub()
                                    If IsNothing(PDFDocument) Then
                                        CurrentPage = -1
                                    Else
                                        CurrentPage -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFForwardHandler(sender As Object, e As EventArgs) Handles PDFForward.Click
        Dim oAction As Action = Sub()
                                    If IsNothing(PDFDocument) Then
                                        CurrentPage = -1
                                    Else
                                        CurrentPage += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFEndHandler(sender As Object, e As EventArgs) Handles PDFEnd.Click
        Dim oAction As Action = Sub()
                                    If IsNothing(PDFDocument) Then
                                        CurrentPage = -1
                                    Else
                                        CurrentPage = PDFDocument.PageCount
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PDFPrintToolTipHandler(sender As Object, e As Controls.ToolTipEventArgs) Handles PDFPrint.ToolTipOpening
        PDFPrint.InnerToolTip = "Print PDF File - " + (New System.Drawing.Printing.PrinterSettings).PrinterName
    End Sub
#End Region
#Region "UI"
    Private Sub RectanglePDFHost_ZoomHandler(sender As Object, e As Input.MouseWheelEventArgs) Handles RectanglePDFHost.MouseWheel
        If e.MiddleButton = Input.MouseButtonState.Pressed Then
            If e.Delta > 0 Then
                ScrollViewerPDFHost.Tag = CDbl(Val(ScrollViewerPDFHost.Tag)) * 1.25
            ElseIf e.Delta < 0 Then
                ScrollViewerPDFHost.Tag = CDbl(Val(ScrollViewerPDFHost.Tag)) / 1.25
            End If
            CurrentPage = CurrentPage
        End If
    End Sub
    Private Sub RectanglePDFHost_ManipulationStarting(sender As Object, e As Input.ManipulationStartingEventArgs) Handles RectanglePDFHost.ManipulationStarting
        e.ManipulationContainer = Me
        e.Handled = True
    End Sub
    Private Sub RectanglePDFHost_ManipulationDelta(ByVal sender As Object, ByVal e As Input.ManipulationDeltaEventArgs) Handles RectanglePDFHost.ManipulationDelta
        ScrollViewerPDFHost.Tag *= Math.Sqrt((e.DeltaManipulation.Scale.X + e.DeltaManipulation.Scale.Y) / 2)

        ScrollViewerPDFHost.ScrollToHorizontalOffset(ScrollViewerPDFHost.HorizontalOffset - e.DeltaManipulation.Translation.X)
        ScrollViewerPDFHost.ScrollToVerticalOffset(ScrollViewerPDFHost.VerticalOffset - e.DeltaManipulation.Translation.Y)
        CurrentPage = CurrentPage
    End Sub
    Private Sub RectanglePDFHost_InertiaStarting(ByVal sender As Object, ByVal e As Input.ManipulationInertiaStartingEventArgs) Handles RectanglePDFHost.ManipulationInertiaStarting
        e.TranslationBehavior.DesiredDeceleration = 10.0 * 96.0 / (1000.0 * 1000.0)
        e.Handled = True
    End Sub
    Private Sub RectanglePDFHost_Zoom100Handler(sender As Object, e As Input.TouchEventArgs) Handles RectanglePDFHost.TouchDown, RectanglePDFHost.TouchUp
        ' check for double touch only if a field is not selected
        If Not DirectSelectSub(e) Then
            Dim oAction As Action = Sub()
                                        ScrollViewerPDFHost.Tag = 0.8
                                        CurrentPage = CurrentPage
                                    End Sub
            CheckDoubleTouch(sender, e, oAction, MinTouchInterval)
        End If
    End Sub
    Private Sub RectanglePDFHost_LeftMouseDownHandler(sender As Object, e As Input.MouseButtonEventArgs) Handles RectanglePDFHost.MouseLeftButtonDown
        ' selects the input field if present
        DirectSelectSub(e)
    End Sub
#End Region
End Class

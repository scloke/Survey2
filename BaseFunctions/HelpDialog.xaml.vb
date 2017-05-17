Imports PdfSharp
Imports PdfSharp.Drawing
Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Media

Public Class HelpDialog
    Inherits Window
    Implements INotifyPropertyChanged

#Region "Variables"
    Private m_GUID As Guid
    Private m_Icon As IconType = IconType.Help
    Private m_Image As ImageSource
    Private m_Resized As Boolean = False
    Public Const InitialViewWidth As Single = 1
#End Region
#Region "INotifyPropertyChanged"
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private Sub OnPropertyChangedLocal(ByVal sName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
    End Sub
#End Region
    Public Sub New(ByVal fMaxWidth As Double, ByVal fMaxHeight As Double, ByVal fColumnWidth As Double, ByVal fRowHeight As Double)
        Me.InitializeComponent()

        MaxWidth = Math.Floor(fMaxWidth)
        MaxHeight = Math.Floor(fMaxHeight)

        ' set button sizes
        GridButtons.ColumnDefinitions(0).Width = New GridLength(fColumnWidth * 2)
        GridButtons.ColumnDefinitions(1).Width = New GridLength(fColumnWidth)
        GridMain.RowDefinitions(1).Height = New GridLength(fRowHeight)

        SetDialogIcon()

        ScrollViewerPDFHost.Tag = InitialViewWidth
        ImagePage.DataContext = Me
    End Sub
    Private Sub ResizeImage()
        ' resize the image
        If m_Resized Then
            m_Resized = False
            Me.Top = (SystemParameters.WorkArea.Height - Me.ActualHeight) / 2
            Me.Left = (SystemParameters.WorkArea.Width - Me.ActualWidth) / 2
        Else
            If Not IsNothing(Image) Then
                Dim oScrollViewerBlockContent As Controls.ScrollViewer = ScrollViewerPDFHost
                Dim oScrollBar As Controls.Primitives.ScrollBar = CType(oScrollViewerBlockContent.Template.FindName("PART_VerticalScrollBar", oScrollViewerBlockContent), Controls.Primitives.ScrollBar)

                If Not IsNothing(oScrollBar) Then
                    Me.SizeToContent = SizeToContent.WidthAndHeight
                    oScrollViewerBlockContent.UpdateLayout()

                    Dim fScaleFactor As Double = Math.Min(Math.Min(MaxWidth, oScrollViewerBlockContent.ActualWidth - oScrollBar.ActualWidth) / Image.PixelWidth, 1)
                    Dim fImageWidth As Double = Image.PixelWidth * fScaleFactor
                    Dim fImageHeight As Double = Image.PixelHeight * fScaleFactor

                    ' set image
                    With RectanglePDFHost
                        .Width = fImageWidth
                        .Height = fImageHeight
                    End With

                    With ImagePage
                        .Width = fImageWidth
                        .Height = fImageHeight
                    End With
                End If

                m_Resized = True
            End If
        End If
    End Sub
    Public Property GUID As Guid
        Get
            Return m_GUID
        End Get
        Set(value As Guid)
            m_GUID = value

            ' set image
            If IsNothing(oSettings.HelpFile) Then
                Image = Nothing
            Else
                Dim oPDFDocument As Pdf.PdfDocument = oSettings.HelpFile.GetPDFDocument(m_GUID)
                If IsNothing(oPDFDocument) Then
                    Image = Nothing
                Else
                    Image = DocumentToImage(oPDFDocument, RenderResolution300)
                End If
            End If
        End Set
    End Property
    Public Property IconLabel As IconType
        Get
            Return m_Icon
        End Get
        Set(value As IconType)
            m_Icon = value
            SetDialogIcon()
        End Set
    End Property
    Public Enum IconType As Integer
        Help = 0
        Info
    End Enum
    Private Sub SetDialogIcon()
        ' sets the dialog icon and label
        Select Case m_Icon
            Case IconType.Help
                DialogIcon.Source = BitmapToBitmapSource(My.Resources.IconHelp.ToBitmap)
                DialogLabel.Content = "Help"
            Case IconType.Info
                DialogIcon.Source = BitmapToBitmapSource(My.Resources.IconInfo.ToBitmap)
                DialogLabel.Content = "Info"
            Case Else
                DialogIcon.Source = Nothing
                DialogLabel.Content = String.Empty
        End Select
    End Sub
    Public Property Image As Imaging.BitmapSource
        Get
            Return m_Image
        End Get
        Set(value As Imaging.BitmapSource)
            m_Image = value
            OnPropertyChangedLocal("Image")
        End Set
    End Property
    Private Sub HelpDialog_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        ' when there is a change in size, fix the image size
        ResizeImage()
    End Sub
    Private Shared Function DocumentToImage(ByVal oPDFDocument As Pdf.PdfDocument, ByVal iResolution As Integer) As Imaging.BitmapSource
        ' renders a PDF document into a list of bitmaps
        If oPDFDocument.PageCount > 0 Then
            Dim oPDFDocumentBytes As Byte()
            Using oMemoryStream As New IO.MemoryStream
                oPDFDocument.Save(oMemoryStream)
                oPDFDocumentBytes = oMemoryStream.ToArray()
                oPDFDocument = Pdf.IO.PdfReader.Open(New IO.MemoryStream(oPDFDocumentBytes))
            End Using

            Dim oBitmapSource As Imaging.BitmapSource = Nothing
            Using oMemoryStream As New IO.MemoryStream(oPDFDocumentBytes)
                Using oDocument = PdfiumViewer.PdfDocument.Load(oMemoryStream)
                    If Not IsNothing(oDocument) Then
                        For i = 0 To oDocument.PageCount - 1
                            Dim iWidth As Integer = Math.Ceiling(XUnit.FromPoint(oDocument.PageSizes(i).Width).Inch * iResolution)
                            Dim iHeight As Integer = Math.Ceiling(XUnit.FromPoint(oDocument.PageSizes(i).Height).Inch * iResolution)

                            Using oBitmap As System.Drawing.Bitmap = oDocument.Render(i, iWidth, iHeight, iResolution, iResolution, False)
                                oBitmapSource = BitmapToBitmapSource(oBitmap)
                            End Using
                        Next
                    End If
                End Using
            End Using
            Return oBitmapSource
        Else
            Return Nothing
        End If
    End Function
    Private Shared Function BitmapToBitmapSource(ByVal oBitmap As System.Drawing.Bitmap) As Imaging.BitmapSource
        If oBitmap.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppArgb Then
            Dim oRect As New System.Drawing.Rectangle(0, 0, oBitmap.Width, oBitmap.Height)
            Dim oBitmapData As System.Drawing.Imaging.BitmapData = oBitmap.LockBits(oRect, System.Drawing.Imaging.ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Dim iBufferSize As Integer = oBitmapData.Stride * oBitmap.Height
            Dim oWriteableBitmap As New Imaging.WriteableBitmap(oBitmap.Width, oBitmap.Height, oBitmap.HorizontalResolution, oBitmap.VerticalResolution, PixelFormats.Bgra32, Nothing)
            oWriteableBitmap.WritePixels(New Int32Rect(0, 0, oBitmap.Width, oBitmap.Height), oBitmapData.Scan0, iBufferSize, oBitmapData.Stride)
            oBitmap.UnlockBits(oBitmapData)
            Return oWriteableBitmap
        Else
            Return Nothing
        End If
    End Function
#Region "Buttons"
    ' add processing for the buttons here
    Private Sub HelpOK_Button_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
        Me.Close()
    End Sub
#End Region
#Region "UI"
    Private Sub RectanglePDFHost_ManipulationStarting(sender As Object, e As Input.ManipulationStartingEventArgs) Handles RectanglePDFHost.ManipulationStarting
        e.ManipulationContainer = Me
        e.Handled = True
    End Sub
    Private Sub RectanglePDFHost_ManipulationDelta(ByVal sender As Object, ByVal e As Input.ManipulationDeltaEventArgs) Handles RectanglePDFHost.ManipulationDelta
        ScrollViewerPDFHost.ScrollToVerticalOffset(ScrollViewerPDFHost.VerticalOffset - e.DeltaManipulation.Translation.Y)
    End Sub
    Private Sub RectanglePDFHost_InertiaStarting(ByVal sender As Object, ByVal e As Input.ManipulationInertiaStartingEventArgs) Handles RectanglePDFHost.ManipulationInertiaStarting
        e.TranslationBehavior.DesiredDeceleration = 10.0 * 96.0 / (1000.0 * 1000.0)
        e.Handled = True
    End Sub
#End Region
End Class
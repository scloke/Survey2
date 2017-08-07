Imports PdfSharp
Imports PdfSharp.Drawing
Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports Common.Common
Imports Twain32Shared.Twain32Shared
Imports System.IO.IsolatedStorage
Imports System.Collections.Concurrent
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Threading
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Media

Public Class Scanner
    Implements FunctionInterface

#Region "Variables"
    ' PluginName defines the friendly name of the plugin
    ' Priority determines the order in which the buttons are arranged on the main page. The lower the number, the earlier it is placed
    Private Const PluginName As String = "Scanner"
    Private Const Priority As Integer = 2
    Private Const IsolatedImageDirectory As String = "Image"
    Private Shared m_Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer))
    Private Shared m_Icons As Dictionary(Of String, ImageSource)
    Private Shared ButtonDictionary As New Dictionary(Of String, FrameworkElement)
    Private Shared ChoiceRectangles As New Dictionary(Of Guid, List(Of Tuple(Of Rect, Integer, Integer)))
    Private Shared BoxRectangles As New Dictionary(Of Guid, List(Of Tuple(Of Rect, Integer, Integer)))
    Private Shared FilterState As Enumerations.FilterData = Enumerations.FilterData.None
    Public Shared Root As Scanner
    Public Shared MarkState As Integer = 0
    Public Shared ViewerState As Boolean = False
    Private WithEvents m_CommonScanner As CommonScanner
    Private ImageStore As New BlockingCollection(Of Tuple(Of String, Integer, Integer, Integer, Integer))
    Private ImageIsolatedStorage As IsolatedStorageFile = Nothing
    Private m_SelectedScannerSource As Tuple(Of Twain32Enumerations.ScannerSource, String, String) = Nothing
    Private RawBarCodesHash As Integer = 0
    Private ScanPageProgress As Integer = 0
    Private Shared oHOGDescriptor As Emgu.CV.HOGDescriptor = Nothing
    Private Shared oRansacLine As Accord.MachineLearning.Geometry.RansacLine
    Private Shared oStructuringElement As Emgu.CV.Matrix(Of Byte)
    Private ctMainSource As CancellationTokenSource
    Private ctMain As CancellationToken
    Public Shared UIDispatcher As Threading.Dispatcher
    Public Const MaxChoiceField As Integer = 10
    Private Const fDetectionThreshold As Double = 0.2

    ' insert your own GUIDs here
    Private Shared oRecogniserStoreGUID As New Guid("{a94ec857-90d5-4e5f-8633-3ec9f8053cc0}")
#End Region
#Region "FunctionInterface"
    Public Function GetDataTypes() As List(Of Tuple(Of Guid, Type)) Implements FunctionInterface.GetDataTypes
        ' returns a list of GUIDs and variable type representing the data types that the plug-in creates
        Dim oDataTypes As New List(Of Tuple(Of Guid, Type))
        oDataTypes.Add(New Tuple(Of Guid, Type)(oRecogniserStoreGUID, GetType(Recognisers)))
        Return oDataTypes
    End Function
    Public Function CheckDataTypes(ByRef oCommonVariables As CommonVariables) As Boolean Implements FunctionInterface.CheckDataTypes
        ' checks the commonvariable data store to see if the variables objects required have been properly initialised
        Dim bCheck As Boolean = True

        ' check variables
        If Not oCommonVariables.DataStore(oRecogniserStoreGUID).GetType.Equals(GetType(Recognisers)) Then
            bCheck = False
        End If

        Return bCheck
    End Function
    Public Function GetIdentifier() As Tuple(Of Guid, ImageSource, String, Integer) Implements FunctionInterface.GetIdentifier
        ' returns the identifiers: GUID, icon, friendly name, and priority for the plugin
        Return New Tuple(Of Guid, ImageSource, String, Integer)(Guid.NewGuid, Converter.BitmapToBitmapSource(My.Resources.IconScanner1), PluginName, Priority)
    End Function
    WriteOnly Property Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)) Implements FunctionInterface.Identifiers
        Set(value As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)))
            ' set buttons to link to the other plugins
            m_Identifiers = value
            Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
            For i = 0 To oFilteredNames.Count - 1
                Dim oButton As Controls.Button = ButtonDictionary("Button" + i.ToString)
                Dim oImage As Controls.Image = ButtonDictionary("ButtonImage" + i.ToString)
                oButton.ToolTip = "Go To " + m_Identifiers(oFilteredNames(i)).Item3
                oButton.IsEnabled = True
                oImage.Source = m_Identifiers(oFilteredNames(i)).Item2
            Next
        End Set
    End Property
    ' the status message event handler to update the main program status window
    Public Event StatusMessage(oMessage As Messages.Message) Implements FunctionInterface.StatusMessage
    ' propogates the exit message to the parent page
    Event ExitButtonClick(oActivatePluginGUID As Guid) Implements FunctionInterface.ExitButtonClick
#End Region
#Region "Main"
    Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CommonFunctions.IterateElementDictionary(ButtonDictionary, GridMain, "Button")
        Root = Me

        ' set cancellation tokens and run processing loops
        ctMainSource = New CancellationTokenSource()
        ctMain = ctMainSource.Token
        If IsNothing(UIDispatcher) Then
            UIDispatcher = Threading.Dispatcher.CurrentDispatcher
        End If

        Task.Factory.StartNew(Sub() ProcessImage(ctMain), ctMain, TaskCreationOptions.LongRunning)

        ' initialise isolated storage
        ImageIsolatedStorage = IsolatedStorageFile.GetUserStoreForAssembly
        ImageIsolatedStorage.CreateDirectory(IsolatedImageDirectory)

        If IsNothing(FieldDocumentStore.Field.InteractionEvent) Then
            FieldDocumentStore.Field.InteractionEvent = New RelayCommand
        End If

        ' initially set all buttons to disabled and clear tooltips
        For i = 0 To 2
            Dim oButton As Controls.Button = ButtonDictionary("Button" + i.ToString)
            Dim oImage As Controls.Image = ButtonDictionary("ButtonImage" + i.ToString)
            oButton.ToolTip = String.Empty
            oButton.IsEnabled = False
            oImage.Source = Nothing
        Next

        ' initialise shared variables
        oRansacLine = New Accord.MachineLearning.Geometry.RansacLine(1.5, 0.99)
        oStructuringElement = New Emgu.CV.Matrix(Of Byte)(3, 3)
        oStructuringElement.SetValue(1)
        oStructuringElement(1, 1) = 0

        SetIcons()
        SetScanIcons()
    End Sub
    Protected Overrides Sub Finalize()
        ' runs when application is closed
        If Not IsNothing(ImageIsolatedStorage) Then
            ImageIsolatedStorage.Close()
            ImageIsolatedStorage.Dispose()
            ImageIsolatedStorage = Nothing
        End If
        If Not IsNothing(m_CommonScanner) Then
            m_CommonScanner.Close()
        End If

        ctMainSource.Cancel()

        MyBase.Finalize()
    End Sub
    Private Sub Page_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles Me.Loaded
        ' checks to perform when navigating to this page
        If IsNothing(m_CommonScanner) Then
            m_CommonScanner = New CommonScanner
            SetScanner()
        End If

        ' refresh the display to account for a display resolution change in the configuration
        Using oUpdateSuspender As New UpdateSuspender(False)
            Dim oSelectedField As FieldDocumentStore.Field = DataGridScanner.SelectedItem
            DataGridScannerChanged(False, True)

            ' if the selected field is a choice field and a checkbox is not selected, then select the first checkbox
            If (Not IsNothing(oSelectedField)) AndAlso (oSelectedField.FieldType = Enumerations.FieldTypeEnum.Choice Or oSelectedField.FieldType = Enumerations.FieldTypeEnum.ChoiceVertical Or oSelectedField.FieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ) AndAlso (Not IsNothing(Keyboard.FocusedElement)) AndAlso (Not Keyboard.FocusedElement.GetType.Equals(GetType(Controls.CheckBox))) Then
                UpdateSuspender.MoveDataGridFocus(DataGridScanner.SelectedIndex, MarkState, 0, 0, False)
            Else
                UpdateSuspender.MoveDataGridFocus(DataGridScanner.SelectedIndex, MarkState, 0, 0, False, False)
            End If
        End Using
    End Sub
    Private Sub SetIcons()
        m_Icons = New Dictionary(Of String, ImageSource)
        m_Icons.Add("ScannerSpreadsheet", Converter.XamlToDrawingImage(My.Resources.ScannerSpreadsheet))
        m_Icons.Add("CCMPDF", Converter.XamlToDrawingImage(My.Resources.CCMPDF))
        m_Icons.Add("CCMScan", Converter.XamlToDrawingImage(My.Resources.CCMScan))
        m_Icons.Add("CCMScanSelect", Converter.XamlToDrawingImage(My.Resources.CCMScanSelect))
        m_Icons.Add("ScannerConfigure", Converter.XamlToDrawingImage(My.Resources.ScannerConfigure))
        m_Icons.Add("CCMDefinitions", Converter.XamlToDrawingImage(My.Resources.CCMDefinitions))
        m_Icons.Add("CCMSave", Converter.XamlToDrawingImage(My.Resources.CCMSave))
        m_Icons.Add("CCMFilter", Converter.XamlToDrawingImage(My.Resources.CCMFilter))
        m_Icons.Add("CCMFilterGreen", Converter.XamlToDrawingImage(My.Resources.CCMFilterGreen))
        m_Icons.Add("CCMFilterOrange", Converter.XamlToDrawingImage(My.Resources.CCMFilterOrange))
        m_Icons.Add("CCMMark0", Converter.XamlToDrawingImage(My.Resources.CCMMark0))
        m_Icons.Add("CCMMark1", Converter.XamlToDrawingImage(My.Resources.CCMMark1))
        m_Icons.Add("CCMMark2", Converter.XamlToDrawingImage(My.Resources.CCMMark2))
        m_Icons.Add("CCMBack", Converter.XamlToDrawingImage(My.Resources.CCMBack))
        m_Icons.Add("CCMForward", Converter.XamlToDrawingImage(My.Resources.CCMForward))
        m_Icons.Add("CCMBackPink", Converter.XamlToDrawingImage(My.Resources.CCMBackPink))
        m_Icons.Add("CCMForwardPink", Converter.XamlToDrawingImage(My.Resources.CCMForwardPink))
        m_Icons.Add("CCMDoubleGears", Converter.XamlToDrawingImage(My.Resources.CCMDoubleGears))
        m_Icons.Add("CCMVisible", Converter.XamlToDrawingImage(My.Resources.CCMVisible))
        m_Icons.Add("CCMNotVisible", Converter.XamlToDrawingImage(My.Resources.CCMNotVisible))

        ScannerLoad.HBSource = GetIcon("CCMDefinitions")
        ScannerSave.HBSource = GetIcon("CCMSave")
        ScannerScan.HBSource = GetIcon("CCMScan")
        ScannerChoose.HBSource = GetIcon("CCMScanSelect")
        ScannerConfigure.HBSource = GetIcon("ScannerConfigure")
        ScannerPreviousSubject.HBSource = GetIcon("CCMBack")
        ScannerNextSubject.HBSource = GetIcon("CCMForward")
        ScannerPreviousData.HBSource = GetIcon("CCMBackPink")
        ScannerNextData.HBSource = GetIcon("CCMForwardPink")
        ScannerProcessed.HBSource = GetIcon("CCMDoubleGears")
        ScannerHideProcessed.HBSource = GetIcon("CCMVisible")
    End Sub
    Private Sub SetScanIcons()
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        If oScannerCollection.Count = 0 Then
            ScannerScan.IsEnabled = False
            ScannerSave.IsEnabled = False
            ScannerFilter.IsEnabled = False
            ScannerMark.IsEnabled = False
            ScannerView.IsEnabled = False
        Else
            If IsNothing(SelectedScannerSource) Then
                ScannerScan.IsEnabled = False
            Else
                ScannerScan.IsEnabled = True
            End If
            ScannerSave.IsEnabled = True
            ScannerFilter.IsEnabled = True
            ScannerMark.IsEnabled = True
            ScannerView.IsEnabled = True
        End If

        ' sets icons and tooltips for scanner icons
        Select Case FilterState
            Case Enumerations.FilterData.None
                ScannerFilter.HBSource = GetIcon("CCMFilter")
                ScannerFilter.InnerToolTip = "Filter: None"
            Case Enumerations.FilterData.DataMissing
                ScannerFilter.HBSource = GetIcon("CCMFilterOrange")
                ScannerFilter.InnerToolTip = "Filter: Missing Data"
            Case Enumerations.FilterData.DataPresent
                ScannerFilter.HBSource = GetIcon("CCMFilterGreen")
                ScannerFilter.InnerToolTip = "Filter: Data Present"
        End Select

        Select Case MarkState
            Case 0
                ScannerMark.HBSource = GetIcon("CCMMark0")
                ScannerMark.InnerToolTip = "Mark: Detected"
                ScannerDetectedMarks.Text = "Detected Marks"
            Case 1
                ScannerMark.HBSource = GetIcon("CCMMark1")
                ScannerMark.InnerToolTip = "Mark: Verified"
                ScannerDetectedMarks.Text = "Verified Marks"
            Case 2
                ScannerMark.HBSource = GetIcon("CCMMark2")
                ScannerMark.InnerToolTip = "Mark: Final"
                ScannerDetectedMarks.Text = "Final Marks"
        End Select

        If ViewerState Then
            ' if bViewerState is true, then show PDF mode, otherwise show table mode
            ScannerView.HBSource = GetIcon("CCMPDF")
            ScannerView.InnerToolTip = "Viewer: PDF"
            DataGridScanner.Visibility = Visibility.Hidden
            PDFViewerControl.Visibility = Visibility.Visible
        Else
            ScannerView.HBSource = GetIcon("ScannerSpreadsheet")
            ScannerView.InnerToolTip = "Viewer: Table"
            DataGridScanner.Visibility = Visibility.Visible
            PDFViewerControl.Visibility = Visibility.Hidden
        End If
    End Sub
#End Region
#Region "Properties"
    Private Shared ReadOnly Property GetIcon(ByVal sIconName As String) As ImageSource
        Get
            If m_Icons.ContainsKey(sIconName) Then
                Return m_Icons(sIconName)
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region
#Region "Functions"
    Public Shared Sub DataGridScannerChanged(Optional ByVal bScannedChanged As Boolean = True, Optional ByVal bDetectedChanged As Boolean = True)
        ' refreshes the images shown in the right sided panels
        Dim oDataGridScanner As Controls.DataGrid = Root.DataGridScanner
        If Not UpdateSuspender.GlobalSuspendProcessing Then
            If IsNothing(oDataGridScanner.SelectedItem) Then
                If oDataGridScanner.Items.Count > 0 Then
                    oDataGridScanner.SelectedIndex = 0
                    UpdateSuspender.ShowImages(0, bScannedChanged, bDetectedChanged)
                Else
                    UpdateSuspender.ShowImages(-1, bScannedChanged, bDetectedChanged)
                End If
            Else
                UpdateSuspender.ShowImages(oDataGridScanner.SelectedIndex, bScannedChanged, bDetectedChanged)
            End If
        End If
    End Sub
    Private Sub DataGridScanner_SelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs) Handles DataGridScanner.SelectionChanged
        ' refreshes the images shown in the right sided panels
        Using oUpdateSuspender As New UpdateSuspender(False)
            If Not UpdateSuspender.GlobalSuspendProcessing Then
                For Each oField As FieldDocumentStore.Field In e.RemovedItems
                    If oField.FieldType = Enumerations.FieldTypeEnum.Free Then
                        oField.MarkFree0 = Trim(oField.MarkFree0)
                        oField.MarkFree1 = Trim(oField.MarkFree1)
                        oField.MarkFree2 = Trim(oField.MarkFree2)
                    End If
                Next

                Dim oSelectedField As FieldDocumentStore.Field = DataGridScanner.SelectedItem
                If (Not IsNothing(oSelectedField)) AndAlso CType(Root.GridMain.Resources("cvsScannnerCollection"), Data.CollectionViewSource).View.Groups.Count > 0 Then
                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                    Dim oDocumentLocationList As List(Of Tuple(Of Guid, Integer, Rect)) = (From oField As FieldDocumentStore.Field In oScannerCollection Select New Tuple(Of Guid, Integer, Rect)(oField.GUID, oField.PageNumber, oField.Location)).ToList
                    SetPDFDocument(oSelectedField.RawBarCodes, oDocumentLocationList)
                    PDFViewerControl.SetActiveRectangles(New List(Of Tuple(Of Guid, Integer, Rect)) From {New Tuple(Of Guid, Integer, Rect)(oSelectedField.GUID, oSelectedField.PageNumber - 1, oSelectedField.Location)})

                    PDFViewerControl.FullPageView()
                    PDFViewerControl.CurrentPage = oSelectedField.PageNumber - 1
                Else
                    SetPDFDocument()
                    PDFViewerControl.CurrentPage = -1
                End If

                DataGridScannerChanged()

                ' if the selected field is a choice field and a checkbox is not selected, then select the first checkbox
                If (Not IsNothing(oSelectedField)) AndAlso (oSelectedField.FieldType = Enumerations.FieldTypeEnum.Choice Or oSelectedField.FieldType = Enumerations.FieldTypeEnum.ChoiceVertical Or oSelectedField.FieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ) AndAlso (Not IsNothing(Keyboard.FocusedElement)) AndAlso (Not Keyboard.FocusedElement.GetType.Equals(GetType(Controls.CheckBox))) Then
                    UpdateSuspender.MoveDataGridFocus(DataGridScanner.SelectedIndex, MarkState, 0, 0, False)
                Else
                    UpdateSuspender.MoveDataGridFocus(DataGridScanner.SelectedIndex, MarkState, 0, 0, False, False)
                End If
            End If
        End Using
    End Sub
    Private Sub DataGridScanner_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles DataGridScanner.SizeChanged
        ' refreshes the images shown in the right sided panels
        DataGridScannerChanged()

        If MarkState = 2 Then
            Dim oDataGridScanner As Controls.DataGrid = Root.DataGridScanner
            oDataGridScanner.CommitEdit()
            oDataGridScanner.CommitEdit()

            Dim oCollectionViewSource As Data.CollectionViewSource = Root.GridMain.Resources("cvsScannnerCollection")
            oCollectionViewSource.View.Refresh()
        End If
    End Sub
    Private Sub SetPDFDocument(Optional ByVal oRawBarCodes As List(Of Tuple(Of String, String, String)) = Nothing, Optional oDocumentLocationList As List(Of Tuple(Of Guid, Integer, Rect)) = Nothing)
        ' adds barcodes to the PDF template and save to the PDF viewer
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        If IsNothing(oScannerCollection) Then
            PDFViewerControl.SetPDFDocument(Nothing)
        Else
            Dim oPDFDocument As Pdf.PdfDocument = Nothing
            If Not IsNothing(oScannerCollection.FieldDocumentStore.PDFTemplate) Then
                Using oMemoryStream As New IO.MemoryStream(oScannerCollection.FieldDocumentStore.PDFTemplate)
                    oPDFDocument = Pdf.IO.PdfReader.Open(oMemoryStream)
                End Using
            End If

            If IsNothing(oPDFDocument) Then
                PDFViewerControl.SetPDFDocument(Nothing)
            Else
                If IsNothing(oRawBarCodes) Then
                    If RawBarCodesHash <> 0 Then
                        ' set blank template
                        PDFViewerControl.SetPDFDocument(oPDFDocument)
                        RawBarCodesHash = 0
                    End If
                Else
                    ' change PDF document if bar codes are different
                    If RawBarCodesHash <> oRawBarCodes.GetHashCode Then
                        For i = 0 To oPDFDocument.Pages.Count - 1
                            Dim oPage As Pdf.PdfPage = oPDFDocument.Pages(i)
                            If i < oRawBarCodes.Count AndAlso oRawBarCodes(i).Item1 <> String.Empty Then
                                ' if the barcode data is empty, then do not draw the barcode
                                PDFHelper.DrawBarcode(oPage, oRawBarCodes(i).Item1, oRawBarCodes(i).Item2, oRawBarCodes(i).Item3)
                            End If
                        Next
                        PDFViewerControl.SetPDFDocument(oPDFDocument, oDocumentLocationList)
                        RawBarCodesHash = oRawBarCodes.GetHashCode
                    End If
                End If
            End If
        End If
    End Sub
    Private Shared Sub ShowImages(ByVal iSelectedIndex As Integer, Optional ByVal bScannedChanged As Boolean = True, Optional ByVal bDetectedChanged As Boolean = True)
        ' code for refreshing the images
        Dim oDataGridScanner As Controls.DataGrid = Root.DataGridScanner
        Dim oImageScannedImageContent As Controls.Image = Root.ImageScannedImageContent
        Dim oImageDetectedMarksContent As Controls.Image = Root.ImageDetectedMarksContent
        Dim oRectangleScannedImageBackground As Shapes.Rectangle = Root.RectangleScannedImageBackground

        Dim fResolution As Single = oSettings.RenderResolutionValueMax
        If iSelectedIndex = -1 Then
            oImageScannedImageContent.Source = Nothing
            oImageScannedImageContent.Visibility = Visibility.Hidden
            oImageDetectedMarksContent.Source = Nothing
            oImageDetectedMarksContent.Visibility = Visibility.Hidden
        Else
            Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")

            Dim oMatrixStore As FieldDocumentStore.MatrixStore = oScannerCollection.FieldDocumentStore.FieldMatrixStore
            oImageScannedImageContent.Visibility = Visibility.Visible
            oImageDetectedMarksContent.Visibility = Visibility.Visible

            Const fFontSize As Double = 10
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontComicSansMS, fFontSize, XFontStyle.Regular, oFontOptions)

            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
            Dim oScannedBitmap As System.Drawing.Bitmap = Nothing
            Dim oDetectedBitmap As System.Drawing.Bitmap = Nothing

            ChoiceRectangles.Clear()
            BoxRectangles.Clear()
            Dim oField As FieldDocumentStore.Field = oDataGridScanner.SelectedItem
            Dim fFieldMargin As Double = SpacingSmall * 72 / fResolution
            If Not IsNothing(oField) Then
                Select Case oField.FieldType
                    Case Enumerations.FieldTypeEnum.Choice
                        Dim XWidth As New XUnit(oField.Location.Width + (fFieldMargin * 2))
                        Dim XHeight As New XUnit(oField.Location.Height + (fFieldMargin * 2))
                        Dim XAdjustedWidth As XUnit = Nothing
                        Dim XAdjustedHeight As XUnit = Nothing

                        Dim iTabletColumns As Integer = Math.Ceiling(oField.MarkCount / oField.TabletGroups)
                        Dim fDescriptionHeight As Double = (oField.Location.Height - (fSingleBlockWidth * oField.TabletGroups)) / 2
                        If oRectangleScannedImageBackground.ActualWidth > 0 And oRectangleScannedImageBackground.ActualHeight > 0 Then
                            Dim fDisplayRatio As Double = oRectangleScannedImageBackground.ActualWidth / oRectangleScannedImageBackground.ActualHeight
                            Dim fBitmapRatio As Double = XWidth.Point / XHeight.Point
                            If fDisplayRatio > fBitmapRatio Then
                                ' the display box has a wider aspect than the bitmap
                                ' adjusted the width
                                XAdjustedWidth = New XUnit(XWidth.Point * fDisplayRatio / fBitmapRatio)
                                XAdjustedHeight = XHeight
                            Else
                                ' the display box has a narrower aspect than the bitmap
                                ' adjust the height
                                XAdjustedWidth = XWidth
                                XAdjustedHeight = New XUnit(XHeight.Point * fBitmapRatio / fDisplayRatio)
                            End If
                        Else
                            XAdjustedWidth = XWidth
                            XAdjustedHeight = XHeight
                        End If
                        Dim XDisplacement As New XPoint((XAdjustedWidth.Point + fFieldMargin - (fSingleBlockWidth * iTabletColumns)) / 2, XAdjustedHeight.Point / 2)

                        Dim XSize As New XSize(XAdjustedWidth.Point, XAdjustedHeight.Point)
                        Dim iBitmapWidth As Integer = Math.Ceiling(XAdjustedWidth.Inch * fResolution)
                        Dim iBitmapHeight As Integer = Math.Ceiling(XAdjustedHeight.Inch * fResolution)
                        If bScannedChanged Then
                            oScannedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oScannedBitmap.SetResolution(fResolution, fResolution)

                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oScannedBitmap)
                                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                    Dim oTabletBitmaps As New List(Of System.Drawing.Bitmap)
                                    Dim oTabletImages As New List(Of XImage)
                                    Dim bImagePresent As Boolean = False
                                    For i = 0 To oField.ImageCount - 1
                                        Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oField.Images(i).Item2), CInt(oField.Images(0).Item7))
                                            If IsNothing(oBitmap) Then
                                                oTabletImages.Add(Nothing)
                                            Else
                                                bImagePresent = True
                                                Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                                Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                                Dim oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                oResampledBitmap.SetResolution(fResolution, fResolution)
                                                oTabletBitmaps.Add(oResampledBitmap)
                                                Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                    oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                    oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                    oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                End Using
                                                Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                oTabletImages.Add(oXResampledBitmap)
                                            End If
                                        End Using
                                    Next

                                    ' overlay tablet descriptions
                                    If bImagePresent Then
                                        PDFHelper.DrawFieldTablets(oXGraphics, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, New FieldDocumentStore.Field, oTabletImages, oField.TabletDescriptionTop, oField.TabletDescriptionBottom)
                                    End If

                                    ' clean up
                                    For Each oBitmap In oTabletBitmaps
                                        oBitmap.Dispose()
                                    Next
                                End Using
                            End Using
                        End If
                        If bDetectedChanged Then
                            oDetectedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oDetectedBitmap.SetResolution(fResolution, fResolution)

                            Dim XDimension As New XUnit(fSingleBlockWidth * 0.8)
                            Dim fDimension As Double = XDimension.Inch * fResolution
                            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
                            Using oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, fDimension, fDimension, Enumerations.StretchEnum.Uniform), fResolution)
                                Dim oXCheckMarkImage As XImage = XImage.FromGdiPlusImage(oCheckMarkBitmap)

                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oDetectedBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                        Dim oInputField As New FieldDocumentStore.Field
                                        PDFHelper.DrawFieldTablets(oXGraphics, XDisplacement, oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, oInputField, Nothing, oField.TabletDescriptionTop, oField.TabletDescriptionBottom)

                                        ' add rect as a fraction of the detected bitmap dimensions (ie. 0-1)
                                        If Not ChoiceRectangles.ContainsKey(oField.GUID) Then
                                            ChoiceRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                        End If
                                        For i = 0 To oInputField.Images.Count - 1
                                            Dim oImageRect As Rect = oInputField.Images(i).Item1
                                            ChoiceRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(oImageRect.X / XSize.Width, oImageRect.Y / XSize.Height, oImageRect.Width / XSize.Width, oImageRect.Height / XSize.Height), oInputField.Images(i).Item4, oInputField.Images(i).Item5))
                                        Next

                                        oXGraphics.SmoothingMode = XSmoothingMode.None
                                        For i = 0 To oField.MarkCount - 1
                                            Select Case MarkState
                                                Case 0
                                                    If oField.MarkChoice0(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oInputField.Images(i).Item1.X + (oInputField.Images(i).Item1.Width / 2) - (oXCheckMarkImage.PointWidth / 2), oInputField.Images(i).Item1.Y + (oInputField.Images(i).Item1.Height / 2) - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                                Case 1
                                                    If oField.MarkChoice1(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oInputField.Images(i).Item1.X + (oInputField.Images(i).Item1.Width / 2) - (oXCheckMarkImage.PointWidth / 2), oInputField.Images(i).Item1.Y + (oInputField.Images(i).Item1.Height / 2) - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                                Case 2
                                                    If oField.MarkChoice2(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oInputField.Images(i).Item1.X + (oInputField.Images(i).Item1.Width / 2) - (oXCheckMarkImage.PointWidth / 2), oInputField.Images(i).Item1.Y + (oInputField.Images(i).Item1.Height / 2) - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                            End Select
                                        Next
                                    End Using
                                End Using
                            End Using
                        End If
                    Case Enumerations.FieldTypeEnum.ChoiceVertical
                        Dim XWidth As New XUnit(oField.Location.Width + (fFieldMargin * 2))
                        Dim XHeight As New XUnit(oField.Location.Height + (fFieldMargin * 2))
                        Dim XAdjustedWidth As XUnit = Nothing
                        Dim XAdjustedHeight As XUnit = Nothing

                        If oRectangleScannedImageBackground.ActualWidth > 0 And oRectangleScannedImageBackground.ActualHeight > 0 Then
                            Dim fDisplayRatio As Double = oRectangleScannedImageBackground.ActualWidth / oRectangleScannedImageBackground.ActualHeight
                            Dim fBitmapRatio As Double = XWidth.Point / XHeight.Point
                            If fDisplayRatio > fBitmapRatio Then
                                ' the display box has a wider aspect than the bitmap
                                ' adjusted the width
                                XAdjustedWidth = New XUnit(XWidth.Point * fDisplayRatio / fBitmapRatio)
                                XAdjustedHeight = XHeight
                            Else
                                ' the display box has a narrower aspect than the bitmap
                                ' adjust the height
                                XAdjustedWidth = XWidth
                                XAdjustedHeight = New XUnit(XHeight.Point * fBitmapRatio / fDisplayRatio)
                            End If
                        Else
                            XAdjustedWidth = XWidth
                            XAdjustedHeight = XHeight
                        End If
                        Dim XDisplacement As New XPoint(fFieldMargin + (XAdjustedWidth.Point - XWidth.Point) / 2, fFieldMargin + (XAdjustedHeight.Point - XHeight.Point) / 2)

                        Dim XSize As New XSize(XAdjustedWidth.Point, XAdjustedHeight.Point)
                        Dim iBitmapWidth As Integer = Math.Ceiling(XAdjustedWidth.Inch * fResolution)
                        Dim iBitmapHeight As Integer = Math.Ceiling(XAdjustedHeight.Inch * fResolution)

                        Dim oTabletImage As List(Of XPoint) = (From oImage In oField.Images Select New XPoint(XDisplacement.X + oImage.Item1.X + oImage.Item1.Width / 2 - oField.Location.Left, XDisplacement.Y + oImage.Item1.Y + oImage.Item1.Height / 2 - oField.Location.Top)).ToList
                        Dim oTabletRectTop As List(Of XRect) = (From oDescriptionTop In oField.TabletDescriptionTop Select If(oDescriptionTop.Item1.IsEmpty, XRect.Empty, New XRect(XDisplacement.X + oDescriptionTop.Item1.X - oField.Location.Left, XDisplacement.Y + oDescriptionTop.Item1.Y - oField.Location.Top, oDescriptionTop.Item1.Width, oDescriptionTop.Item1.Height))).ToList
                        Dim oTabletRectBottom As List(Of XRect) = (From oDescriptionBottom In oField.TabletDescriptionBottom Select If(oDescriptionBottom.Item1.IsEmpty, XRect.Empty, New XRect(XDisplacement.X + oDescriptionBottom.Item1.X - oField.Location.Left, XDisplacement.Y + oDescriptionBottom.Item1.Y - oField.Location.Top, oDescriptionBottom.Item1.Width, oDescriptionBottom.Item1.Height))).ToList
                        Dim oTabletDisplacements As List(Of Tuple(Of XPoint, XRect, XRect)) = (From iIndex In Enumerable.Range(0, oTabletImage.Count) Select New Tuple(Of XPoint, XRect, XRect)(oTabletImage(iIndex), If(iIndex > oTabletRectTop.Count - 1, Nothing, oTabletRectTop(iIndex)), If(iIndex > oTabletRectBottom.Count - 1, Nothing, oTabletRectBottom(iIndex)))).ToList

                        If bScannedChanged Then
                            oScannedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oScannedBitmap.SetResolution(fResolution, fResolution)

                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oScannedBitmap)
                                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                    Dim oTabletBitmaps As New List(Of System.Drawing.Bitmap)
                                    Dim oTabletImages As New List(Of XImage)
                                    Dim bImagePresent As Boolean = False
                                    For i = 0 To oField.ImageCount - 1
                                        Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oField.Images(i).Item2), CInt(oField.Images(0).Item7))
                                            If IsNothing(oBitmap) Then
                                                oTabletImages.Add(Nothing)
                                            Else
                                                bImagePresent = True
                                                Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                                Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                                Dim oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                oResampledBitmap.SetResolution(fResolution, fResolution)
                                                oTabletBitmaps.Add(oResampledBitmap)
                                                Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                    oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                    oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                    oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                End Using
                                                Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                oTabletImages.Add(oXResampledBitmap)
                                            End If
                                        End Using
                                    Next

                                    ' overlay tablet descriptions
                                    If bImagePresent Then
                                        Select Case oField.TabletAlignment
                                            Case Enumerations.AlignmentEnum.Left
                                                PDFHelper.DrawFieldTabletsVertical(oXGraphics, XWidth, oField.TabletLimit, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, New FieldDocumentStore.Field, oField.TabletAlignment, oTabletImages, Nothing, oField.TabletDescriptionBottom, oTabletDisplacements)
                                            Case Enumerations.AlignmentEnum.Center
                                                PDFHelper.DrawFieldTabletsVertical(oXGraphics, XWidth, oField.TabletLimit, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, New FieldDocumentStore.Field, oField.TabletAlignment, oTabletImages, oField.TabletDescriptionTop, oField.TabletDescriptionBottom, oTabletDisplacements)
                                            Case Enumerations.AlignmentEnum.Right
                                                PDFHelper.DrawFieldTabletsVertical(oXGraphics, XWidth, oField.TabletLimit, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, New FieldDocumentStore.Field, oField.TabletAlignment, oTabletImages, oField.TabletDescriptionTop, Nothing, oTabletDisplacements)
                                        End Select
                                    End If

                                    ' clean up
                                    For Each oBitmap In oTabletBitmaps
                                        oBitmap.Dispose()
                                    Next
                                End Using
                            End Using
                        End If
                        If bDetectedChanged Then
                            oDetectedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oDetectedBitmap.SetResolution(fResolution, fResolution)

                            Dim XDimension As New XUnit(fSingleBlockWidth * 0.8)
                            Dim fDimension As Double = XDimension.Inch * fResolution
                            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
                            Using oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, fDimension, fDimension, Enumerations.StretchEnum.Uniform), fResolution)
                                Dim oXCheckMarkImage As XImage = XImage.FromGdiPlusImage(oCheckMarkBitmap)
                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oDetectedBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                        Dim oInputField As New FieldDocumentStore.Field
                                        Select Case oField.TabletAlignment
                                            Case Enumerations.AlignmentEnum.Left
                                                PDFHelper.DrawFieldTabletsVertical(oXGraphics, XWidth, oField.TabletLimit, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, oInputField, oField.TabletAlignment, Nothing, Nothing, oField.TabletDescriptionBottom, oTabletDisplacements)
                                            Case Enumerations.AlignmentEnum.Center
                                                PDFHelper.DrawFieldTabletsVertical(oXGraphics, XWidth, oField.TabletLimit, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, oInputField, oField.TabletAlignment, Nothing, oField.TabletDescriptionTop, oField.TabletDescriptionBottom, oTabletDisplacements)
                                            Case Enumerations.AlignmentEnum.Right
                                                PDFHelper.DrawFieldTabletsVertical(oXGraphics, XWidth, oField.TabletLimit, New XPoint(XDisplacement.X, XDisplacement.Y), oField.MarkCount, oField.TabletStart, oField.TabletGroups, oField.TabletContent, oInputField, oField.TabletAlignment, Nothing, oField.TabletDescriptionTop, Nothing, oTabletDisplacements)
                                        End Select

                                        ' add rect as a fraction of the detected bitmap dimensions (ie. 0-1)
                                        If Not ChoiceRectangles.ContainsKey(oField.GUID) Then
                                            ChoiceRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                        End If
                                        For i = 0 To oInputField.Images.Count - 1
                                            Dim oImageRect As Rect = oInputField.Images(i).Item1
                                            ChoiceRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(oImageRect.X / XSize.Width, oImageRect.Y / XSize.Height, oImageRect.Width / XSize.Width, oImageRect.Height / XSize.Height), oInputField.Images(i).Item4, oInputField.Images(i).Item5))
                                        Next

                                        oXGraphics.SmoothingMode = XSmoothingMode.None
                                        For i = 0 To oField.MarkCount - 1
                                            Select Case MarkState
                                                Case 0
                                                    If oField.MarkChoice0(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oInputField.Images(i).Item1.X + (oInputField.Images(i).Item1.Width / 2) - (oXCheckMarkImage.PointWidth / 2), oInputField.Images(i).Item1.Y + (oInputField.Images(i).Item1.Height / 2) - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                                Case 1
                                                    If oField.MarkChoice1(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oInputField.Images(i).Item1.X + (oInputField.Images(i).Item1.Width / 2) - (oXCheckMarkImage.PointWidth / 2), oInputField.Images(i).Item1.Y + (oInputField.Images(i).Item1.Height / 2) - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                                Case 2
                                                    If oField.MarkChoice2(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oInputField.Images(i).Item1.X + (oInputField.Images(i).Item1.Width / 2) - (oXCheckMarkImage.PointWidth / 2), oInputField.Images(i).Item1.Y + (oInputField.Images(i).Item1.Height / 2) - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                            End Select
                                        Next
                                    End Using
                                End Using
                            End Using
                        End If
                    Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        Dim XWidth As New XUnit(oField.Location.Width + (fFieldMargin * 2))
                        Dim XHeight As New XUnit(oField.Location.Height + (fFieldMargin * 2))
                        Dim XAdjustedWidth As XUnit = Nothing
                        Dim XAdjustedHeight As XUnit = Nothing

                        If oRectangleScannedImageBackground.ActualWidth > 0 And oRectangleScannedImageBackground.ActualHeight > 0 Then
                            Dim fDisplayRatio As Double = oRectangleScannedImageBackground.ActualWidth / oRectangleScannedImageBackground.ActualHeight
                            Dim fBitmapRatio As Double = XWidth.Point / XHeight.Point
                            If fDisplayRatio > fBitmapRatio Then
                                ' the display box has a wider aspect than the bitmap
                                ' adjusted the width
                                XAdjustedWidth = New XUnit(XWidth.Point * fDisplayRatio / fBitmapRatio)
                                XAdjustedHeight = XHeight
                            Else
                                ' the display box has a narrower aspect than the bitmap
                                ' adjust the height
                                XAdjustedWidth = XWidth
                                XAdjustedHeight = New XUnit(XHeight.Point * fBitmapRatio / fDisplayRatio)
                            End If
                        Else
                            XAdjustedWidth = XWidth
                            XAdjustedHeight = XHeight
                        End If
                        Dim XDisplacement As New XPoint(fFieldMargin + (XAdjustedWidth.Point - XWidth.Point) / 2, fFieldMargin + (XAdjustedHeight.Point - XHeight.Point) / 2)

                        Dim XSize As New XSize(XAdjustedWidth.Point, XAdjustedHeight.Point)
                        Dim iBitmapWidth As Integer = Math.Ceiling(XAdjustedWidth.Inch * fResolution)
                        Dim iBitmapHeight As Integer = Math.Ceiling(XAdjustedHeight.Inch * fResolution)

                        Dim oTabletImage As List(Of XRect) = (From oImage In oField.Images Select New XRect(XDisplacement.X + oImage.Item1.X - oField.Location.Left, XDisplacement.Y + oImage.Item1.Y - oField.Location.Top, oImage.Item1.Width, oImage.Item1.Height)).ToList
                        Dim oTabletRectMCQ As List(Of Tuple(Of Integer, Integer, XRect, List(Of ElementStruc))) = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Select New Tuple(Of Integer, Integer, XRect, List(Of ElementStruc))(oField.TabletDescriptionMCQ(iIndex).Item2, oField.TabletDescriptionMCQ(iIndex).Item3, If(oField.TabletDescriptionMCQ(iIndex).Item1.IsEmpty, XRect.Empty, New XRect(XDisplacement.X + oField.TabletDescriptionMCQ(iIndex).Item1.X - oField.Location.Left, XDisplacement.Y + oField.TabletDescriptionMCQ(iIndex).Item1.Y - oField.Location.Top, oField.TabletDescriptionMCQ(iIndex).Item1.Width, oField.TabletDescriptionMCQ(iIndex).Item1.Height)), oField.TabletDescriptionMCQ(iIndex).Item4)).ToList
                        Dim oTabletDisplacements As List(Of Tuple(Of XRect, Integer, Integer, XRect, List(Of ElementStruc))) = (From iIndex In Enumerable.Range(0, oTabletImage.Count) Select New Tuple(Of XRect, Integer, Integer, XRect, List(Of ElementStruc))(New XRect(oTabletImage(iIndex).X, oTabletImage(iIndex).Y, oTabletImage(iIndex).Width, oTabletImage(iIndex).Height), If(iIndex > oTabletRectMCQ.Count - 1, -1, oTabletRectMCQ(iIndex).Item1), If(iIndex > oTabletRectMCQ.Count - 1, -1, oTabletRectMCQ(iIndex).Item2), If(iIndex > oTabletRectMCQ.Count - 1, Nothing, oTabletRectMCQ(iIndex).Item3), If(iIndex > oTabletRectMCQ.Count - 1, -1, oTabletRectMCQ(iIndex).Item4))).ToList

                        If bScannedChanged Then
                            oScannedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oScannedBitmap.SetResolution(fResolution, fResolution)

                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oScannedBitmap)
                                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                    Dim oTabletBitmaps As New List(Of System.Drawing.Bitmap)
                                    Dim oTabletImages As New List(Of XImage)
                                    Dim bImagePresent As Boolean = False
                                    For i = 0 To oField.ImageCount - 1
                                        Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oField.Images(i).Item2), CInt(oField.Images(0).Item7))
                                            If IsNothing(oBitmap) Then
                                                oTabletImages.Add(Nothing)
                                            Else
                                                bImagePresent = True
                                                Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                                Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                                Dim oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                oResampledBitmap.SetResolution(fResolution, fResolution)
                                                oTabletBitmaps.Add(oResampledBitmap)
                                                Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                    oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                    oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                    oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                End Using
                                                Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                oTabletImages.Add(oXResampledBitmap)
                                            End If
                                        End Using
                                    Next

                                    ' overlay tablet descriptions
                                    If bImagePresent Then
                                        Dim oInputField As New FieldDocumentStore.Field
                                        oInputField.TabletGroups = oField.TabletGroups
                                        oInputField.TabletContent = oField.TabletContent
                                        oInputField.TabletDescriptionMCQ.AddRange(oField.TabletDescriptionMCQ)
                                        PDFHelper.DrawFieldTabletsMCQ(oXGraphics, New XUnit(oField.TabletMCQParams.Item1), New XUnit(oField.TabletMCQParams.Item2), XDisplacement, oField.TabletMCQParams.Item4, oField.TabletMCQParams.Item5, oField.TabletMCQParams.Item6, oField.TabletMCQParams.Item7, oInputField, oTabletImages, oTabletDisplacements)
                                    End If

                                    ' clean up
                                    For Each oBitmap In oTabletBitmaps
                                        oBitmap.Dispose()
                                    Next
                                End Using
                            End Using
                        End If
                        If bDetectedChanged Then
                            oDetectedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oDetectedBitmap.SetResolution(fResolution, fResolution)

                            Dim XDimension As New XUnit(fSingleBlockWidth * 0.8)
                            Dim fDimension As Double = XDimension.Inch * fResolution
                            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
                            Using oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, fDimension, fDimension, Enumerations.StretchEnum.Uniform), fResolution)
                                Dim oXCheckMarkImage As XImage = XImage.FromGdiPlusImage(oCheckMarkBitmap)

                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oDetectedBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                        Dim oInputField As New FieldDocumentStore.Field
                                        oInputField.TabletGroups = oField.TabletGroups
                                        oInputField.TabletContent = oField.TabletContent
                                        oInputField.TabletDescriptionMCQ.AddRange(oField.TabletDescriptionMCQ)
                                        PDFHelper.DrawFieldTabletsMCQ(oXGraphics, New XUnit(oField.TabletMCQParams.Item1), New XUnit(oField.TabletMCQParams.Item2), XDisplacement, oField.TabletMCQParams.Item4, oField.TabletMCQParams.Item5, oField.TabletMCQParams.Item6, oField.TabletMCQParams.Item7, oInputField, Nothing, oTabletDisplacements)

                                        ' add rect as a fraction of the detected bitmap dimensions (ie. 0-1)
                                        If Not ChoiceRectangles.ContainsKey(oField.GUID) Then
                                            ChoiceRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                        End If
                                        For i = 0 To oInputField.Images.Count - 1
                                            Dim oImageRect As Rect = oInputField.Images(i).Item1
                                            ChoiceRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(oImageRect.X / XSize.Width, oImageRect.Y / XSize.Height, oImageRect.Width / XSize.Width, oImageRect.Height / XSize.Height), oInputField.Images(i).Item4, oInputField.Images(i).Item5))
                                        Next

                                        oXGraphics.SmoothingMode = XSmoothingMode.None
                                        For i = 0 To oField.MarkCount - 1
                                            Select Case MarkState
                                                Case 0
                                                    If oField.MarkChoice0(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oTabletImage(i).X + oTabletImage(i).Width / 2 - (oXCheckMarkImage.PointWidth / 2), oTabletImage(i).Y + oTabletImage(i).Height / 2 - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                                Case 1
                                                    If oField.MarkChoice1(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oTabletImage(i).X + oTabletImage(i).Width / 2 - (oXCheckMarkImage.PointWidth / 2), oTabletImage(i).Y + oTabletImage(i).Height / 2 - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                                Case 2
                                                    If oField.MarkChoice2(i) Then
                                                        oXGraphics.DrawImage(oXCheckMarkImage, oTabletImage(i).X + oTabletImage(i).Width / 2 - (oXCheckMarkImage.PointWidth / 2), oTabletImage(i).Y + oTabletImage(i).Height / 2 - (oXCheckMarkImage.PointHeight / 2))
                                                    End If
                                            End Select
                                        Next
                                    End Using
                                End Using
                            End Using
                        End If
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        Dim oBoxImageList As List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))) = (From oImage In oField.Images Where oImage.Item5 = -1 Select oImage).ToList
                        Dim oBoxTabletDictionary As Dictionary(Of Integer, List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))) = (From oBoxImage In oBoxImageList Select New KeyValuePair(Of Integer, List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))))(oBoxImage.Item4, (From oImage In oField.Images Where oImage.Item5 >= 0 And oImage.Item4 = oBoxImage.Item4 Order By oImage.Item5 Ascending Select oImage).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                        Dim fLabelHeight As Double = fSingleBlockWidth / 2
                        Dim fBoxChoiceWidth As Double = fLabelHeight + fSingleBlockWidth + (10 * fSingleBlockWidth)

                        Dim XWidth As New XUnit(fBoxChoiceWidth + (fFieldMargin * 2))
                        Dim XHeight As New XUnit((fSingleBlockWidth * oBoxImageList.Count) + (fFieldMargin * 2))

                        Dim XSize As New XSize(XWidth.Point, XHeight.Point)
                        Dim iBitmapWidth As Integer = Math.Ceiling(XWidth.Inch * fResolution)
                        Dim iBitmapHeight As Integer = Math.Ceiling(XHeight.Inch * fResolution)
                        Dim XDimension As New XUnit(fSingleBlockWidth * 0.8)
                        Dim XSingleBlockDimension As New XUnit(fSingleBlockWidth)

                        If bScannedChanged Then
                            oScannedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oScannedBitmap.SetResolution(fResolution, fResolution)

                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oScannedBitmap)
                                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                    Dim fScaledFontSize As Double = fFontSize * XSingleBlockDimension.Point / oXGraphics.MeasureString("XXX", oTestFont, oStringFormat).Height
                                    Dim oArielFont As New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)

                                    For i = 0 To oBoxImageList.Count - 1
                                        Dim XLeft As New XUnit(fFieldMargin)
                                        Dim XTop As New XUnit(fFieldMargin + (oBoxImageList(i).Item4 * fSingleBlockWidth))

                                        ' draws image or fixed text
                                        If oBoxImageList(i).Item6 Then
                                            If oField.ImagesPresent Then
                                                Dim sChar As String = oField.MarkBoxChoiceRow(oBoxImageList(i).Item4, 0)
                                                If sChar <> String.Empty Then
                                                    PDFHelper.DrawStringRotated(oXGraphics, New XPoint(XLeft.Point + XSingleBlockDimension.Point / 2, XTop.Point + XSingleBlockDimension.Point / 2), oArielFont, oStringFormat, sChar, -90)
                                                End If
                                            End If
                                        Else
                                            Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oBoxImageList(i).Item2), CInt(oField.Images(0).Item7))
                                                If Not IsNothing(oBitmap) Then
                                                    Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                                    Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                                    Using oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                        oResampledBitmap.SetResolution(fResolution, fResolution)
                                                        Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                            oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                            oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                            oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                        End Using
                                                        Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                        Dim fResampledBitmapWidth As Double = oResampledBitmap.Width * 72 / fResolution
                                                        Dim fResampledBitmapHeight As Double = oResampledBitmap.Height * 72 / fResolution

                                                        oXGraphics.SmoothingMode = XSmoothingMode.None
                                                        oXGraphics.DrawImage(oXResampledBitmap, XLeft.Point + (fSingleBlockWidth - fResampledBitmapWidth) / 2, XTop.Point + (fSingleBlockWidth - fResampledBitmapHeight) / 2)
                                                    End Using
                                                End If
                                            End Using
                                        End If

                                        oXGraphics.SmoothingMode = XSmoothingMode.None
                                        For j = 0 To oBoxTabletDictionary(i).Count - 1
                                            Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oBoxTabletDictionary(i)(j).Item2), CInt(oBoxTabletDictionary(i)(0).Item7))
                                                If Not IsNothing(oBitmap) Then
                                                    Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                                    Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                                    Using oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                        oResampledBitmap.SetResolution(fResolution, fResolution)
                                                        Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                            oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                            oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                            oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                        End Using
                                                        Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                        Dim fResampledBitmapWidth As Double = oResampledBitmap.Width * 72 / fResolution
                                                        Dim fResampledBitmapHeight As Double = oResampledBitmap.Height * 72 / fResolution

                                                        Dim XDisplacementLeft As New XUnit(XLeft.Point + (fSingleBlockWidth * 3 / 2) + (fSingleBlockWidth - oXResampledBitmap.PointWidth) / 2 + (j * fSingleBlockWidth))
                                                        Dim XDisplacementTop As New XUnit(XTop.Point + (fSingleBlockWidth - oXResampledBitmap.PointHeight) / 2)

                                                        oXGraphics.DrawImage(oXResampledBitmap, XDisplacementLeft.Point, XDisplacementTop.Point)
                                                    End Using
                                                End If
                                            End Using
                                        Next
                                    Next
                                End Using
                            End Using
                        End If
                        If bDetectedChanged Then
                            oDetectedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oDetectedBitmap.SetResolution(fResolution, fResolution)

                            Dim fDimension As Double = XDimension.Inch * fResolution
                            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
                            Using oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, fDimension, fDimension, Enumerations.StretchEnum.Uniform), fResolution)
                                Dim oXCheckMarkImage As XImage = XImage.FromGdiPlusImage(oCheckMarkBitmap)

                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oDetectedBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                                        Dim fScaledFontSize As Double = fFontSize * XSingleBlockDimension.Point / oXGraphics.MeasureString("XXX", oTestFont, oStringFormat).Height
                                        Dim oArielFont As New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)
                                        Dim oComicSansFont As New XFont(FontComicSansMS, fScaledFontSize, XFontStyle.Regular, oFontOptions)

                                        For i = 0 To oBoxImageList.Count - 1
                                            Dim XLeft As New XUnit(fFieldMargin)
                                            Dim XTop As New XUnit(fFieldMargin + (oBoxImageList(i).Item4 * fSingleBlockWidth))

                                            ' draws border or fixed text
                                            Dim sChar As String = oField.MarkBoxChoiceRow(oBoxImageList(i).Item4, 1)
                                            If oBoxImageList(i).Item6 Then
                                                If sChar <> String.Empty Then
                                                    PDFHelper.DrawStringRotated(oXGraphics, New XPoint(XLeft.Point + XSingleBlockDimension.Point / 2, XTop.Point + XSingleBlockDimension.Point / 2), oArielFont, oStringFormat, sChar, -90)
                                                End If
                                            Else
                                                Dim oInputField As New FieldDocumentStore.Field
                                                PDFHelper.DrawFieldTablets(oXGraphics, New XPoint(XLeft.Point + (fSingleBlockWidth * 3 / 2), XTop.Point), oBoxTabletDictionary(i).Count, oField.TabletStart, 1, oField.TabletContent, oInputField, Nothing)

                                                ' add rect as a fraction of the detected bitmap dimensions (ie. 0-1)
                                                If Not ChoiceRectangles.ContainsKey(oField.GUID) Then
                                                    ChoiceRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                                End If
                                                For j = 0 To oInputField.Images.Count - 1
                                                    Dim oImageRect As Rect = oInputField.Images(j).Item1
                                                    ChoiceRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(oImageRect.X / XSize.Width, oImageRect.Y / XSize.Height, oImageRect.Width / XSize.Width, oImageRect.Height / XSize.Height), i, oInputField.Images(j).Item5))
                                                Next

                                                If Not BoxRectangles.ContainsKey(oField.GUID) Then
                                                    BoxRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                                End If
                                                BoxRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(XLeft.Point / XSize.Width, XTop.Point / XSize.Height, XSingleBlockDimension.Point / XSize.Width, XSingleBlockDimension.Point / XSize.Height), i, -1))
                                                PDFHelper.DrawFieldBorder(oXGraphics, XSingleBlockDimension, XSingleBlockDimension, New XPoint(XLeft, XTop), 0.5)

                                                If sChar <> String.Empty Then
                                                    oXGraphics.DrawString(sChar, oComicSansFont, XBrushes.Black, XLeft.Point + fSingleBlockWidth / 2, XTop.Point + fSingleBlockWidth / 2, oStringFormat)
                                                End If
                                            End If

                                            ' sets the mark
                                            Dim sMarkChar As String = oField.MarkBoxChoiceRow(i, 0)
                                            If sMarkChar <> String.Empty AndAlso IsNumeric(sMarkChar) Then
                                                Dim j As Integer = Val(sMarkChar)
                                                Dim XDisplacementLeft As New XUnit(XLeft.Point + (fSingleBlockWidth * 3 / 2) + (fSingleBlockWidth - oXCheckMarkImage.PointWidth) / 2 + (j * fSingleBlockWidth))
                                                Dim XDisplacementTop As New XUnit(XTop.Point + (fSingleBlockWidth - oXCheckMarkImage.PointHeight) / 2)
                                                oXGraphics.DrawImage(oXCheckMarkImage, XDisplacementLeft.Point, XDisplacementTop.Point)
                                            End If
                                        Next
                                    End Using
                                End Using
                            End Using
                        End If
                    Case Enumerations.FieldTypeEnum.Handwriting
                        If oField.MarkCount > 0 Then
                            Dim xMax As Integer = 0
                            Dim yMax As Integer = 0
                            For i = 0 To oField.MarkCount - 1
                                xMax = Math.Max(xMax, oField.Images(i).Item5)
                                yMax = Math.Max(yMax, oField.Images(i).Item4)
                            Next

                            Dim XWidth As New XUnit((fSingleBlockWidth * (xMax + 1)) + (fFieldMargin * 2))
                            Dim XHeight As New XUnit((fSingleBlockWidth * (yMax + 1)) + (fFieldMargin * 2))

                            Dim XSize As New XSize(XWidth.Point, XHeight.Point)
                            Dim iBitmapWidth As Integer = Math.Ceiling(XWidth.Inch * fResolution)
                            Dim iBitmapHeight As Integer = Math.Ceiling(XHeight.Inch * fResolution)
                            Dim XDimension As New XUnit(fSingleBlockWidth * 0.8)
                            Dim XSingleBlockDimension As New XUnit(fSingleBlockWidth)

                            If bScannedChanged Then
                                oScannedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                                oScannedBitmap.SetResolution(fResolution, fResolution)

                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oScannedBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                        For i = 0 To oField.ImageCount - 1
                                            Dim XLeft As New XUnit(fFieldMargin + (oField.Images(i).Item5 * fSingleBlockWidth))
                                            Dim XTop As New XUnit(fFieldMargin + (oField.Images(i).Item4 * fSingleBlockWidth))

                                            ' draws image or fixed text
                                            If oField.Images(i).Item6 Then
                                                If oField.ImagesPresent Then
                                                    Dim sChar As String = oField.MarkHandwritingRowCol(oField.Images(i).Item4, oField.Images(i).Item5, 0)
                                                    If sChar <> String.Empty Then
                                                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                                                        PDFHelper.DrawText(oXGraphics, XSingleBlockDimension, XUnit.Zero, XSingleBlockDimension, New XPoint(XLeft, XTop), sChar, False, False, False, MigraDoc.DocumentObjectModel.ParagraphAlignment.Center, MigraDoc.DocumentObjectModel.Colors.Black)
                                                    End If
                                                End If
                                            Else
                                                Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oField.Images(i).Item2), CInt(oField.Images(0).Item7))
                                                    If Not IsNothing(oBitmap) Then
                                                        Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                                        Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                                        Using oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                            oResampledBitmap.SetResolution(fResolution, fResolution)
                                                            Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                                oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                                oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                                oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                            End Using
                                                            Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                            Dim fResampledBitmapWidth As Double = oResampledBitmap.Width * 72 / fResolution
                                                            Dim fResampledBitmapHeight As Double = oResampledBitmap.Height * 72 / fResolution

                                                            oXGraphics.SmoothingMode = XSmoothingMode.None
                                                            oXGraphics.DrawImage(oXResampledBitmap, XLeft.Point + (fSingleBlockWidth - fResampledBitmapWidth) / 2, XTop.Point + (fSingleBlockWidth - fResampledBitmapHeight) / 2)
                                                        End Using
                                                    End If
                                                End Using
                                            End If
                                        Next
                                    End Using
                                End Using
                            End If
                            If bDetectedChanged Then
                                oDetectedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                                oDetectedBitmap.SetResolution(fResolution, fResolution)

                                ' draw letters
                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oDetectedBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                        Dim fScaledFontSize As Double = fFontSize * XDimension.Point / oXGraphics.MeasureString("XXX", oTestFont, oStringFormat).Height
                                        Dim oComicSansFont As New XFont(FontComicSansMS, fScaledFontSize, XFontStyle.Regular, oFontOptions)

                                        For i = 0 To oField.MarkCount - 1
                                            Dim XLeft As New XUnit(fFieldMargin + (oField.Images(i).Item5 * fSingleBlockWidth))
                                            Dim XTop As New XUnit(fFieldMargin + (oField.Images(i).Item4 * fSingleBlockWidth))

                                            ' draws border or fixed text
                                            Dim sChar As String = oField.MarkHandwritingRowCol(oField.Images(i).Item4, oField.Images(i).Item5, MarkState)
                                            If oField.Images(i).Item6 Then
                                                If sChar <> String.Empty Then
                                                    PDFHelper.DrawText(oXGraphics, XSingleBlockDimension, XUnit.Zero, XSingleBlockDimension, New XPoint(XLeft, XTop), sChar, False, False, False, MigraDoc.DocumentObjectModel.ParagraphAlignment.Center, MigraDoc.DocumentObjectModel.Colors.Black)
                                                End If
                                            Else
                                                If Not BoxRectangles.ContainsKey(oField.GUID) Then
                                                    BoxRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                                End If
                                                BoxRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(XLeft.Point / XSize.Width, XTop.Point / XSize.Height, XSingleBlockDimension.Point / XSize.Width, XSingleBlockDimension.Point / XSize.Height), oField.Images(i).Item4, oField.Images(i).Item5))
                                                PDFHelper.DrawFieldBorder(oXGraphics, XSingleBlockDimension, XSingleBlockDimension, New XPoint(XLeft, XTop), 0.5)

                                                If sChar <> String.Empty Then
                                                    oXGraphics.DrawString(sChar, oComicSansFont, XBrushes.Black, XLeft.Point + fSingleBlockWidth / 2, XTop.Point + fSingleBlockWidth / 2, oStringFormat)
                                                End If
                                            End If
                                        Next
                                    End Using
                                End Using
                            End If
                        End If
                    Case Enumerations.FieldTypeEnum.Free
                        Dim XWidth As New XUnit(oField.Images(0).Item1.Width + (fFieldMargin * 2))
                        Dim XHeight As New XUnit(oField.Images(0).Item1.Height + (fFieldMargin * 2))
                        Dim XSize As New XSize(XWidth.Point, XHeight.Point)
                        Dim iBitmapWidth As Integer = Math.Ceiling(XWidth.Inch * fResolution)
                        Dim iBitmapHeight As Integer = Math.Ceiling(XHeight.Inch * fResolution)
                        If bScannedChanged Then
                            oScannedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oScannedBitmap.SetResolution(fResolution, fResolution)

                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oScannedBitmap)
                                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                                    Using oBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oMatrixStore.GetMat(oField.Images(0).Item2), CInt(oField.Images(0).Item7))
                                        If Not IsNothing(oBitmap) Then
                                            Dim iActualResampledBitmapWidth As Integer = oBitmap.Width * fResolution / oBitmap.HorizontalResolution
                                            Dim iActualResampledBitmapHeight As Integer = oBitmap.Height * fResolution / oBitmap.VerticalResolution

                                            Using oResampledBitmap As New System.Drawing.Bitmap(iActualResampledBitmapWidth, iActualResampledBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                                oResampledBitmap.SetResolution(fResolution, fResolution)
                                                Using oResampledGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oResampledBitmap)
                                                    oResampledGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                                                    oResampledGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                                    oResampledGraphics.DrawImage(oBitmap, 0, 0, oResampledBitmap.Width, oResampledBitmap.Height)
                                                End Using
                                                Dim oXResampledBitmap As XImage = XImage.FromGdiPlusImage(oResampledBitmap)
                                                Dim fResampledBitmapWidth As Double = oResampledBitmap.Width * 72 / fResolution
                                                Dim fResampledBitmapHeight As Double = oResampledBitmap.Height * 72 / fResolution

                                                oXGraphics.DrawImage(oXResampledBitmap, fFieldMargin + (oField.Images(0).Item1.Width - fResampledBitmapWidth) / 2, fFieldMargin + (oField.Images(0).Item1.Height - fResampledBitmapHeight) / 2)
                                            End Using
                                        End If
                                    End Using
                                End Using
                            End Using
                        End If
                        If bDetectedChanged Then
                            oDetectedBitmap = New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                            oDetectedBitmap.SetResolution(fResolution, fResolution)

                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oDetectedBitmap)
                                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSize, XGraphicsUnit.Point)
                                    If Not BoxRectangles.ContainsKey(oField.GUID) Then
                                        BoxRectangles.Add(oField.GUID, New List(Of Tuple(Of Rect, Integer, Integer)))
                                    End If
                                    BoxRectangles(oField.GUID).Add(New Tuple(Of Rect, Integer, Integer)(New Rect(fFieldMargin / XSize.Width, fFieldMargin / XSize.Height, oField.Images(0).Item1.Width / XSize.Width, oField.Images(0).Item1.Height / XSize.Height), -1, -1))
                                    PDFHelper.DrawFieldBorder(oXGraphics, oField.Images(0).Item1.Width, oField.Images(0).Item1.Height, New XPoint(fFieldMargin, fFieldMargin), 0.5)

                                    ' draw text
                                    Dim sFieldString As String = String.Empty
                                    Select Case MarkState
                                        Case 0
                                            sFieldString = Trim(oField.MarkFree0)
                                        Case 1
                                            sFieldString = Trim(oField.MarkFree1)
                                        Case 2
                                            sFieldString = Trim(oField.MarkFree2)
                                    End Select
                                    If sFieldString <> String.Empty Then
                                        Dim oTextSize As XSize = oXGraphics.MeasureString(sFieldString, oTestFont, oStringFormat)
                                        Dim fScale As Double = Math.Min(oField.Images(0).Item1.Width * 0.8 / oTextSize.Width, oField.Images(0).Item1.Height * 0.8 / oTextSize.Height)
                                        Dim fScaledFontSize As Double = fFontSize * fScale
                                        Dim oComicSansFont As New XFont(FontComicSansMS, fScaledFontSize, XFontStyle.Regular, oFontOptions)
                                        oXGraphics.DrawString(sFieldString, oComicSansFont, XBrushes.Black, XWidth.Point / 2, XHeight.Point / 2, oStringFormat)
                                    End If
                                End Using
                            End Using
                        End If
                End Select
            End If

            If bScannedChanged Then
                If IsNothing(oScannedBitmap) Then
                    oImageScannedImageContent.Source = Nothing
                Else
                    oImageScannedImageContent.Source = Converter.BitmapToBitmapSource(oScannedBitmap.Clone)
                    oScannedBitmap.Dispose()
                End If
            End If
            If bDetectedChanged Then
                If IsNothing(oDetectedBitmap) Then
                    oImageDetectedMarksContent.Source = Nothing
                Else
                    oImageDetectedMarksContent.Source = Converter.BitmapToBitmapSource(oDetectedBitmap.Clone)
                    oDetectedBitmap.Dispose()
                End If
            End If
        End If
    End Sub
    Public Shared Sub MoveDataGridFocus(ByVal iIndex As Integer, ByVal iOrder As Integer, ByVal iColumn As Integer, ByVal iRow As Integer, ByVal bLast As Boolean, Optional ByVal bMove As Boolean = True)
        ' moves focus to the specified field in the datagrid
        Dim oDataGrid As Controls.DataGrid = Root.DataGridScanner
        If iIndex >= 0 And iIndex < oDataGrid.Items.Count Then
            Dim oField As FieldDocumentStore.Field = oDataGrid.Items(iIndex)
            Dim iItem4 As Integer = -1
            Dim iItem5 As Integer = -1

            Dim sControlName As String = String.Empty
            If bLast Then
                Select Case oField.FieldType
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        Dim iMaxColumn As Integer = Aggregate iCurrentIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iCurrentIndex).Item5 = -1 And Not oField.Images(iCurrentIndex).Item6 Into Max(oField.Images(iCurrentIndex).Item4)
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + iOrder.ToString + "_" + iMaxColumn.ToString
                        iItem4 = iMaxColumn
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        Dim iGroups As Integer = Math.Ceiling(oField.MarkCount / Scanner.MaxChoiceField)
                        Dim iRowLength As Integer = Math.Ceiling(oField.MarkCount / iGroups)
                        Dim iSetCol As Integer = (oField.Images.Count - 1) Mod iRowLength
                        Dim iSetRow As Integer = iGroups - 1
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + iOrder.ToString + "_" + iSetCol.ToString + "_" + iSetRow.ToString
                        iItem5 = (iSetRow * iRowLength) + iSetCol
                    Case Enumerations.FieldTypeEnum.Handwriting
                        Dim iMaxColumn As Integer = Aggregate iCurrentIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where Not oField.Images(iCurrentIndex).Item6 Into Max(oField.Images(iCurrentIndex).Item5)
                        Dim iMaxRow As Integer = Aggregate iCurrentIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where Not oField.Images(iCurrentIndex).Item6 Into Max(oField.Images(iCurrentIndex).Item4)
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + iOrder.ToString + "_" + iMaxColumn.ToString + "_" + iMaxRow.ToString
                        iItem4 = iMaxRow
                        iItem5 = iMaxColumn
                    Case Enumerations.FieldTypeEnum.Free
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + iOrder.ToString
                End Select
            Else
                Select Case oField.FieldType
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + iOrder.ToString + "_" + iColumn.ToString
                        iItem4 = iColumn
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        Dim iGroups As Integer = Math.Ceiling(oField.MarkCount / Scanner.MaxChoiceField)
                        Dim iRowLength As Integer = Math.Ceiling(oField.MarkCount / iGroups)
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + iOrder.ToString + "_" + iColumn.ToString + "_" + iRow.ToString
                        iItem5 = (iRow * iRowLength) + iColumn
                    Case Enumerations.FieldTypeEnum.Handwriting
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + iOrder.ToString + "_" + iColumn.ToString + "_" + iRow.ToString
                        iItem4 = iRow
                        iItem5 = iColumn
                    Case Enumerations.FieldTypeEnum.Free
                        sControlName = "Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + iOrder.ToString
                End Select
            End If

            If sControlName <> String.Empty Then
                If bMove Then
                    Dim oTextBoxList As List(Of Controls.TextBox) = CommonFunctions.GetChildObjects(Of Controls.TextBox)(oDataGrid, sControlName)
                    If oTextBoxList.Count > 0 Then
                        oDataGrid.SelectedIndex = iIndex
                        oTextBoxList.First.Focus()
                    Else
                        Dim oCheckBoxList As List(Of Controls.CheckBox) = CommonFunctions.GetChildObjects(Of Controls.CheckBox)(oDataGrid, sControlName)
                        If oCheckBoxList.Count > 0 Then
                            oDataGrid.SelectedIndex = iIndex
                            Dim oAction As Action = Function() oCheckBoxList.First.Focus()
                            oCheckBoxList.First.Dispatcher.BeginInvoke(Threading.DispatcherPriority.Background, oAction)
                        End If
                    End If
                End If

                ' add red rectangle
                Dim oImageDetectedMarksContent As Controls.Image = Root.ImageDetectedMarksContent
                Dim oCanvasDetectedMarksContent As Controls.Canvas = Root.CanvasDetectedMarksContent
                oCanvasDetectedMarksContent.Children.Clear()
                Select Case oField.FieldType
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        SetRectangle(oImageDetectedMarksContent, oCanvasDetectedMarksContent, oField, iItem4, iItem5, True)
                    Case Enumerations.FieldTypeEnum.BoxChoice, Enumerations.FieldTypeEnum.Handwriting, Enumerations.FieldTypeEnum.Free
                        SetRectangle(oImageDetectedMarksContent, oCanvasDetectedMarksContent, oField, iItem4, iItem5, False)
                End Select
            End If
        End If
    End Sub
    Private Shared Sub SetRectangle(ByRef oImageDetectedMarksContent As Controls.Image, ByRef oCanvasDetectedMarksContent As Controls.Canvas, ByVal oField As FieldDocumentStore.Field, ByVal iItem4 As Integer, ByVal iItem5 As Integer, ByVal bChoice As Boolean)
        Const fMarginFraction As Double = 0.05
        Dim oFocusRectangleList As List(Of Tuple(Of Rect, Integer, Integer)) = Nothing
        If bChoice Then
            If ChoiceRectangles.ContainsKey(oField.GUID) Then
                oFocusRectangleList = ChoiceRectangles(oField.GUID)
            End If
        Else
            If BoxRectangles.ContainsKey(oField.GUID) Then
                oFocusRectangleList = BoxRectangles(oField.GUID)
            End If
        End If

        If Not IsNothing(oFocusRectangleList) Then
            Dim oSelectedRectangleList As List(Of Rect) = (From oFocusRectangle As Tuple(Of Rect, Integer, Integer) In oFocusRectangleList Where If(iItem4 = -1, True, oFocusRectangle.Item2 = iItem4) And If(iItem5 = -1, True, oFocusRectangle.Item3 = iItem5) Select oFocusRectangle.Item1).ToList
            If oSelectedRectangleList.Count > 0 Then
                oImageDetectedMarksContent.InvalidateMeasure()
                oImageDetectedMarksContent.UpdateLayout()

                Dim fMargin As Double = Math.Min(oSelectedRectangleList.First.Width * oImageDetectedMarksContent.ActualWidth, oSelectedRectangleList.First.Height * oImageDetectedMarksContent.ActualHeight) * fMarginFraction
                Dim oRectangle As New Shapes.Rectangle
                With oRectangle
                    .Width = oSelectedRectangleList.First.Width * oImageDetectedMarksContent.ActualWidth + (2 * fMargin)
                    .Height = oSelectedRectangleList.First.Height * oImageDetectedMarksContent.ActualHeight + (2 * fMargin)
                    .Stroke = New SolidColorBrush(Color.FromArgb(&HFF, &HFF, &H0, &H0))
                    .StrokeThickness = 2
                    .StrokeDashArray = New DoubleCollection From {4}
                    .Fill = Brushes.Transparent
                    .Visibility = Visibility.Visible
                End With

                Dim oTransform As GeneralTransform = oImageDetectedMarksContent.TransformToVisual(oCanvasDetectedMarksContent)
                Dim oTransformedPoint As Point = oTransform.Transform(New Point((oSelectedRectangleList.First.Left * oImageDetectedMarksContent.ActualWidth) - fMargin, (oSelectedRectangleList.First.Top * oImageDetectedMarksContent.ActualHeight) - fMargin))
                Controls.Canvas.SetLeft(oRectangle, oTransformedPoint.X)
                Controls.Canvas.SetTop(oRectangle, oTransformedPoint.Y)
                Controls.Canvas.SetZIndex(oRectangle, 4)
                oCanvasDetectedMarksContent.Children.Add(oRectangle)
            End If
        End If
    End Sub
    Private Sub CollectionViewSource_Filter(ByVal sender As Object, ByVal e As Data.FilterEventArgs)
        Dim oField As FieldDocumentStore.Field = e.Item
        If Not IsNothing(oField) Then
            ' If filter is turned on, filter completed items.
            Select Case FilterState
                Case Enumerations.FilterData.None
                    e.Accepted = True
                Case Enumerations.FilterData.DataMissing
                    If oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataNone Or oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataPartial Then
                        e.Accepted = True
                    Else
                        e.Accepted = False
                    End If
                Case Enumerations.FilterData.DataPresent
                    If oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataFull Then
                        e.Accepted = True
                    Else
                        e.Accepted = False
                    End If
            End Select
        End If
    End Sub
    Private Sub ScannerHeader_SelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs) Handles ScannerHeader.SelectionChanged
        ' updates the data grid
        Dim oComboBoxMain As Controls.ComboBox = sender
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Dim iSelectedIndex As Integer = oComboBoxMain.SelectedIndex

        Dim oHCBDisplay As Common.HighlightComboBox.HCBDisplay = oComboBoxMain.ItemsSource(iSelectedIndex)
        If oHCBDisplay.Name = oScannerCollection.FieldDocumentStore.FieldCollectionStore(iSelectedIndex).SubjectName Then
            oScannerCollection.SelectedCollection = iSelectedIndex
        End If
    End Sub
    Private Sub PDFViewerControl_DirectSelect(ByVal oGUID As Guid) Handles PDFViewerControl.DirectSelect
        ' handles a direct selection by clicking on the PDF viewer
        Dim oFieldList As List(Of FieldDocumentStore.Field) = (From oField As FieldDocumentStore.Field In DataGridScanner.Items Where oField.GUID.Equals(oGUID) Select oField).ToList
        If oFieldList.Count > 0 Then
            DataGridScanner.SelectedItem = oFieldList.First
        End If
    End Sub
    Private Function LocalKnownTypes() As List(Of Type)
        ' gets a list of known local class types
        Dim oKnownTypes As New List(Of Type)

        oKnownTypes.Add(GetType(FieldDocumentStore))
        oKnownTypes.Add(GetType(FieldDocumentStore.Field))
        oKnownTypes.Add(GetType(FieldDocumentStore.FieldCollection))
        oKnownTypes.Add(GetType(FieldDocumentStore.MatrixStore))
        oKnownTypes.Add(GetType(System.Drawing.Bitmap))

        Return oKnownTypes
    End Function
    Private Shared Sub ExportData(ByVal sFileName As String)
        ' exports data to spreadsheet
        Const WorksheetName As String = "Data"
        Const CommentsName As String = "Comments"
        Const StatsName As String = "Statistics"
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")

        Dim oItemTypeList As List(Of Tuple(Of Enumerations.FieldTypeEnum, String)) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Select New Tuple(Of Enumerations.FieldTypeEnum, String)(oField.FieldType, oField.Numbering)).Distinct.ToList
        Dim oModifiedItemList As New List(Of Tuple(Of Enumerations.FieldTypeEnum, String, String))
        For Each oItem In oItemTypeList
            Select Case oItem.Item1
                Case Enumerations.FieldTypeEnum.Choice
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, "C"))
                Case Enumerations.FieldTypeEnum.ChoiceVertical
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, "CV"))
                Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, "MCQ"))
                Case Enumerations.FieldTypeEnum.Handwriting
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, "H"))
                Case Enumerations.FieldTypeEnum.BoxChoice
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, "B"))
                Case Enumerations.FieldTypeEnum.Free
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, "F"))
                Case Else
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String)(oItem.Item1, oItem.Item2, String.Empty))
            End Select
        Next

        Dim oExcelDocument As New ClosedXML.Excel.XLWorkbook

        oExcelDocument.AddWorksheet(WorksheetName)
        For Each oCurrentWorksheet As ClosedXML.Excel.IXLWorksheet In oExcelDocument.Worksheets
            If oCurrentWorksheet.Name <> WorksheetName Then
                oExcelDocument.Worksheets.Delete(oCurrentWorksheet.Name)
            End If
        Next

        Dim oWorksheet As ClosedXML.Excel.IXLWorksheet = oExcelDocument.Worksheet(WorksheetName)

        ' set headers
        oWorksheet.Cell(1, 1).Value = "No."
        oWorksheet.Cell(1, 1).Style.Font.Bold = True
        oWorksheet.Cell(1, 2).Value = "Subject"
        oWorksheet.Cell(1, 2).Style.Font.Bold = True

        Dim oComments As New List(Of Tuple(Of String, Boolean))
        For i = 0 To oModifiedItemList.Count - 1
            Dim sNumbering As String = oModifiedItemList(i).Item2
            oWorksheet.Cell(1, 3 + i).Value = "'" + sNumbering + oModifiedItemList(i).Item3
            oWorksheet.Cell(1, 3 + i).Style.Font.Bold = True

            If oModifiedItemList(i).Item1 = Enumerations.FieldTypeEnum.Choice Or oModifiedItemList(i).Item1 = Enumerations.FieldTypeEnum.ChoiceVertical Or (oModifiedItemList(i).Item1 = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ) Then
                Select Case oModifiedItemList(i).Item1
                    Case Enumerations.FieldTypeEnum.Choice
                        oComments.Add(New Tuple(Of String, Boolean)("Choice Field: " + sNumbering + oModifiedItemList(i).Item3, True))
                    Case Enumerations.FieldTypeEnum.ChoiceVertical
                        oComments.Add(New Tuple(Of String, Boolean)("Choice Vertical Field: " + sNumbering + oModifiedItemList(i).Item3, True))
                    Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        oComments.Add(New Tuple(Of String, Boolean)("Choice MCQ Field: " + oModifiedItemList(i).Item3 + oModifiedItemList(i).Item3, True))
                End Select

                Dim iCurrentIndex As Integer = 0
                Dim oFieldCollection As FieldDocumentStore.FieldCollection = oScannerCollection.FieldDocumentStore.FieldCollectionStore.First
                Dim oField As FieldDocumentStore.Field = (From oCurrentField In oFieldCollection.Fields Where oCurrentField.Numbering = sNumbering Select oCurrentField).First
                Select Case oField.FieldType
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical
                        Dim iMaxCount As Integer = Math.Min(Math.Max(oField.TabletDescriptionTop.Count, oField.TabletDescriptionBottom.Count), oField.ImageCount)
                        For j = 0 To iMaxCount - 1
                            Dim sTabletDescriptionTop As String = If(IsNothing(oField.TabletDescriptionTop(j).Item2), String.Empty, oField.TabletDescriptionTop(j).Item2)
                            Dim sTabletDescriptionBottom As String = If(IsNothing(oField.TabletDescriptionBottom(j).Item2), String.Empty, oField.TabletDescriptionBottom(j).Item2)

                            Dim sContentText As String = String.Empty
                            Select Case oField.TabletContent
                                Case Enumerations.TabletContentEnum.Number
                                    sContentText = (iCurrentIndex + If(oField.TabletStart = -2, 0, oField.TabletStart + 1)).ToString
                                Case Enumerations.TabletContentEnum.Letter
                                    sContentText = Converter.ConvertNumberToLetter(iCurrentIndex + Math.Max(oField.TabletStart, 0), True)
                            End Select

                            Dim sTabletDescription As String = sContentText + ": " + sTabletDescriptionTop + If(sTabletDescriptionTop <> String.Empty And sTabletDescriptionBottom <> String.Empty, " / ", String.Empty) + sTabletDescriptionBottom
                            oComments.Add(New Tuple(Of String, Boolean)(sTabletDescription, False))
                            iCurrentIndex += 1
                        Next
                    Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        For j = 0 To oField.TabletDescriptionMCQ.Count - 1
                            Dim sContentText As String = String.Empty
                            Select Case oField.TabletContent
                                Case Enumerations.TabletContentEnum.Number
                                    sContentText = (j + If(oField.TabletStart = -2, 0, oField.TabletStart + 1)).ToString
                                Case Enumerations.TabletContentEnum.Letter
                                    sContentText = Converter.ConvertNumberToLetter(j + Math.Max(oField.TabletStart, 0), True)
                            End Select

                            Dim sDescription As String = String.Join(" ", (From oElement As ElementStruc In oField.TabletDescriptionMCQ(j).Item4 Select oElement.Text))
                            Dim sTabletDescription As String = sContentText + ": " + sDescription
                            oComments.Add(New Tuple(Of String, Boolean)(sTabletDescription, False))
                        Next
                End Select

                oComments.Add(New Tuple(Of String, Boolean)(String.Empty, False))
            End If
        Next

        For i = 0 To oScannerCollection.FieldDocumentStore.FieldCollectionStore.Count - 1
            Dim oFieldCollection As FieldDocumentStore.FieldCollection = oScannerCollection.FieldDocumentStore.FieldCollectionStore(i)

            oWorksheet.Cell(2 + i, 1).Value = (1 + i).ToString
            oWorksheet.Cell(2 + i, 2).Value = oFieldCollection.SubjectName

            For j = 0 To oModifiedItemList.Count - 1
                Dim sNumbering As String = oModifiedItemList(j).Item2
                Dim oField As FieldDocumentStore.Field = (From oCurrentField In oFieldCollection.Fields Where oCurrentField.Numbering = sNumbering Select oCurrentField).First
                Select Case oModifiedItemList(j).Item1
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        If oField.TabletSingleChoiceOnly Then
                            Dim oMarkTrue As List(Of Integer) = (From iIndex As Integer In Enumerable.Range(0, oField.MarkCount) Where oField.MarkChoice2(iIndex) Select iIndex).ToList
                            If oMarkTrue.Count > 0 Then
                                Dim sContentText As String = String.Empty
                                Select Case oField.TabletContent
                                    Case Enumerations.TabletContentEnum.Letter
                                        sContentText = Converter.ConvertNumberToLetter(oMarkTrue.First + Math.Max(oField.TabletStart, 0), True)
                                    Case Else
                                        sContentText = (oMarkTrue.First + If(oField.TabletStart = -2, 0, oField.TabletStart + 1)).ToString
                                End Select
                                oWorksheet.Cell(2 + i, 3 + j).Value = sContentText
                            End If
                        Else
                            oWorksheet.Cell(2 + i, 3 + j).Value = oField.MarkChoiceCombined2
                        End If
                    Case Enumerations.FieldTypeEnum.Handwriting
                        oWorksheet.Cell(2 + i, 3 + j).Value = oField.MarkHandwritingCombined2
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        oWorksheet.Cell(2 + i, 3 + j).Value = oField.MarkBoxChoiceCombined2
                    Case Enumerations.FieldTypeEnum.Free
                        oWorksheet.Cell(2 + i, 3 + j).Value = oField.MarkFree2
                End Select
                oWorksheet.Cell(2 + i, 3 + j).DataType = ClosedXML.Excel.XLCellValues.Text
            Next
        Next

        ' autofit columns
        For i = 0 To oModifiedItemList.Count + 1
            oWorksheet.Column(1 + i).AdjustToContents()
        Next

        ' set comments
        oExcelDocument.AddWorksheet(CommentsName)
        Dim oCommentWorksheet As ClosedXML.Excel.IXLWorksheet = oExcelDocument.Worksheet(CommentsName)
        For i = 0 To oComments.Count - 1
            oCommentWorksheet.Cell(1 + i, 1).Value = oComments(i).Item1
            If oComments(i).Item2 Then
                oCommentWorksheet.Cell(1 + i, 1).Style.Font.Bold = True
            End If
        Next

        ' set statistics
        ' 1) choice fields
        ' 2) choice vertical fields
        ' 3) box choice fields
        ' 4) MCQ fields
        ' 5) handwriting fields
        ' look into total fields, critical and non-critical fields, fields where detected differs from verified (critical only), or fields where detected differs from final (box choice only)
        ' 6) total fields
        ' 7) fields with no data
        ' 8) fields with partial data
        ' 9) fields with full data
        oExcelDocument.AddWorksheet(StatsName)
        Dim oStatsWorksheet As ClosedXML.Excel.IXLWorksheet = oExcelDocument.Worksheet(StatsName)
        oStatsWorksheet.Cell(1, 1).Value = "Statistics"
        oStatsWorksheet.Cell(1, 1).Style.Font.Bold = True

        Dim iCurrentRow As Integer = 2
        Dim oCurrentFieldType As Enumerations.FieldTypeEnum = Enumerations.FieldTypeEnum.Undefined
        Dim sCurrentFieldText As String = String.Empty
        Dim oCurrentFieldList As New List(Of FieldDocumentStore.Field)
        Dim oCriticalFieldList As New List(Of FieldDocumentStore.Field)
        Dim oNonCriticalFieldList As New List(Of FieldDocumentStore.Field)
        For i = 0 To 4
            Select Case i
                Case 0
                    oCurrentFieldType = Enumerations.FieldTypeEnum.Choice
                    sCurrentFieldText = "Choice"
                Case 1
                    oCurrentFieldType = Enumerations.FieldTypeEnum.ChoiceVertical
                    sCurrentFieldText = "ChoiceVertical"
                Case 2
                    oCurrentFieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    sCurrentFieldText = "ChoiceVerticalMCQ"
                Case 3
                    oCurrentFieldType = Enumerations.FieldTypeEnum.BoxChoice
                    sCurrentFieldText = "BoxChoice"
                Case 4
                    oCurrentFieldType = Enumerations.FieldTypeEnum.Handwriting
                    sCurrentFieldText = "Handwriting"
            End Select

            oCurrentFieldList.Clear()
            oCurrentFieldList = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = oCurrentFieldType Select oField).ToList
            oCriticalFieldList.Clear()
            oCriticalFieldList = (From oField In oCurrentFieldList Where oField.Critical Select oField).ToList
            oNonCriticalFieldList.Clear()
            oNonCriticalFieldList = (From oField In oCurrentFieldList Where Not oField.Critical Select oField).ToList

            iCurrentRow += 1
            oStatsWorksheet.Cell(iCurrentRow, 1).Value = sCurrentFieldText
            oStatsWorksheet.Cell(iCurrentRow, 1).Style.Font.SetUnderline(ClosedXML.Excel.XLFontUnderlineValues.Single)

            Select Case oCurrentFieldType
                Case Enumerations.FieldTypeEnum.BoxChoice
                    iCurrentRow += 1
                    Dim iCurrentTotal As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = -1 Into Count()
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Total Fields: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iCurrentTotal.ToString

                    iCurrentRow += 1
                    Dim iCriticalTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = -1 Into Count()
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Critical Fields: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iCriticalTotal.ToString

                    iCurrentRow += 1
                    Dim iNonCriticalTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = -1 Into Count()
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Non-Critical Fields: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iNonCriticalTotal.ToString

                    iCurrentRow += 1
                    Dim oBoxList As Dictionary(Of Guid, List(Of Integer)) = (From oField In oCurrentFieldList Select New KeyValuePair(Of Guid, List(Of Integer))(oField.GUID, (From iIndex In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = -1 Select iIndex).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    Dim iDFTotal As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oBoxList(oField.GUID).Count) Where (Not oField.Images(oBoxList(oField.GUID)(iIndex)).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    Dim iDFNotEqual As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oBoxList(oField.GUID).Count) Where (Not oField.Images(oBoxList(oField.GUID)(iIndex)).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Detected <> Final (BoxChoice Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iDFTotal > 0, (iDFNotEqual / iDFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iVFTotal As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oBoxList(oField.GUID).Count) Where (Not oField.Images(oBoxList(oField.GUID)(iIndex)).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    Dim iVFNotEqual As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oBoxList(oField.GUID).Count) Where (Not oField.Images(oBoxList(oField.GUID)(iIndex)).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Verified <> Final (BoxChoice Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iVFTotal > 0, (iVFNotEqual / iVFTotal).ToString("P2"), "N/A")

                    ' item by item list
                    iCurrentRow += 1
                    Dim oFieldList As List(Of FieldDocumentStore.Field) = (From oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oBoxList(oField.GUID).Count) Where (Not oField.Images(oBoxList(oField.GUID)(iIndex)).Item5) AndAlso ((Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3)) Or (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3))) Select oField).ToList
                    Dim oSubjectNameList As List(Of String) = (From oField In oFieldList Select oField.SubjectName Distinct).ToList
                    Dim oFieldDictionary As Dictionary(Of String, List(Of FieldDocumentStore.Field)) = (From sSubjectName As String In oSubjectNameList Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sSubjectName, (From oField In oFieldList Where oField.SubjectName = sSubjectName Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    For Each sSubjectName As String In oFieldDictionary.Keys
                        iCurrentRow += 1
                        Dim sCurrentText As String = sSubjectName + ":"
                        For Each oField As FieldDocumentStore.Field In oFieldDictionary(sSubjectName)
                            sCurrentText += "[" + oField.Numbering + "]"
                        Next
                        oStatsWorksheet.Cell(iCurrentRow, 1).Value = sCurrentText
                    Next
                Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    iCurrentRow += 1
                    Dim iCurrentTotal As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Total Tablets: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iCurrentTotal.ToString

                    iCurrentRow += 1
                    Dim iCriticalTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Critical Tablets: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iCriticalTotal.ToString

                    iCurrentRow += 1
                    Dim iNonCriticalTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Non-Critical Tablets: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iNonCriticalTotal.ToString

                    iCurrentRow += 1
                    Dim iDFTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    Dim iDFNotEqual As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where ((Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice0(iIndex) <> oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Detected <> Final (Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iDFTotal > 0, (iDFNotEqual / iDFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iVFTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    Dim iVFNotEqual As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where ((Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice1(iIndex) <> oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Verified <> Final (Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iVFTotal > 0, (iVFNotEqual / iVFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iNDFTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    Dim iNDFNotEqual As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where ((Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice0(iIndex) <> oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Detected <> Final (Non-Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iNDFTotal > 0, (iNDFNotEqual / iNDFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iNVFTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex)) Into Count
                    Dim iNVFNotEqual As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where ((Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice1(iIndex) <> oField.MarkChoice2(iIndex)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Verified <> Final (Non-Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iNVFTotal > 0, (iNVFNotEqual / iNVFTotal).ToString("P2"), "N/A")

                    ' item by item list
                    iCurrentRow += 1
                    Dim oFieldList As List(Of FieldDocumentStore.Field) = (From oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (((Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice0(iIndex) <> oField.MarkChoice2(iIndex))) Or (((Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice1(iIndex) <> oField.MarkChoice2(iIndex))) Select oField).ToList
                    Dim oSubjectNameList As List(Of String) = (From oField In oFieldList Select oField.SubjectName Distinct).ToList
                    Dim oFieldDictionary As Dictionary(Of String, List(Of FieldDocumentStore.Field)) = (From sSubjectName As String In oSubjectNameList Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sSubjectName, (From oField In oFieldList Where oField.SubjectName = sSubjectName Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    For Each sSubjectName As String In oFieldDictionary.Keys
                        iCurrentRow += 1
                        Dim sCurrentText As String = sSubjectName + ":"
                        For Each oField As FieldDocumentStore.Field In oFieldDictionary(sSubjectName)
                            sCurrentText += "[" + oField.Numbering + "]"
                        Next
                        oStatsWorksheet.Cell(iCurrentRow, 1).Value = sCurrentText
                    Next
                Case Enumerations.FieldTypeEnum.Handwriting
                    iCurrentRow += 1
                    Dim iCurrentTotal As Integer = Aggregate oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Trim(oField.MarkText(iIndex).Item1) <> String.Empty Or Trim(oField.MarkText(iIndex).Item2) <> String.Empty Or Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Total Fields: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iCurrentTotal.ToString

                    iCurrentRow += 1
                    Dim iCriticalTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Trim(oField.MarkText(iIndex).Item1) <> String.Empty Or Trim(oField.MarkText(iIndex).Item2) <> String.Empty Or Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Critical Fields: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iCriticalTotal.ToString

                    iCurrentRow += 1
                    Dim iNonCriticalTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Trim(oField.MarkText(iIndex).Item1) <> String.Empty Or Trim(oField.MarkText(iIndex).Item2) <> String.Empty Or Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Non-Critical Fields: "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = iNonCriticalTotal.ToString

                    iCurrentRow += 1
                    Dim iDFTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    Dim iDFNotEqual As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Detected <> Final (Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iDFTotal > 0, (iDFNotEqual / iDFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iVFTotal As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    Dim iVFNotEqual As Integer = Aggregate oField In oCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Verified <> Final (Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iVFTotal > 0, (iVFNotEqual / iVFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iNDFTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    Dim iNDFNotEqual As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Detected <> Final (Non-Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iNDFTotal > 0, (iNDFNotEqual / iNDFTotal).ToString("P2"), "N/A")

                    iCurrentRow += 1
                    Dim iNVFTotal As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) Into Count
                    Dim iNVFNotEqual As Integer = Aggregate oField In oNonCriticalFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3)) Into Count
                    oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Verified <> Final (Non-Critical Fields): "
                    oStatsWorksheet.Cell(iCurrentRow, 2).Value = If(iNVFTotal > 0, (iNVFNotEqual / iNVFTotal).ToString("P2"), "N/A")

                    ' item by item list
                    iCurrentRow += 1
                    Dim oFieldList As List(Of FieldDocumentStore.Field) = (From oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where ((Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3))) Or ((Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3))) Select oField).ToList
                    Dim oSubjectNameList As List(Of String) = (From oField In oFieldList Select oField.SubjectName Distinct).ToList
                    Dim oFieldDictionary As Dictionary(Of String, List(Of FieldDocumentStore.Field)) = (From sSubjectName As String In oSubjectNameList Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sSubjectName, (From oField In oFieldList Where oField.SubjectName = sSubjectName Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    For Each sSubjectName As String In oFieldDictionary.Keys
                        iCurrentRow += 1
                        Dim sCurrentText As String = sSubjectName + ":"
                        For Each oField As FieldDocumentStore.Field In oFieldDictionary(sSubjectName)
                            sCurrentText += "[" + oField.Numbering + "]"
                        Next
                        oStatsWorksheet.Cell(iCurrentRow, 1).Value = sCurrentText
                    Next
            End Select

            iCurrentRow += 1
        Next

        Dim iNoData As Integer = Aggregate oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataNone Into Count
        Dim iPartialData As Integer = Aggregate oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataPartial Into Count
        Dim iFullData As Integer = Aggregate oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataFull Into Count

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Summary"
        oStatsWorksheet.Cell(iCurrentRow, 1).Style.Font.SetUnderline(ClosedXML.Excel.XLFontUnderlineValues.Single)

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Total Fields: "
        oStatsWorksheet.Cell(iCurrentRow, 2).Value = (Aggregate oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Into Count).ToString

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "No Data: "
        oStatsWorksheet.Cell(iCurrentRow, 2).Value = iNoData.ToString

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Partial Data: "
        oStatsWorksheet.Cell(iCurrentRow, 2).Value = iPartialData.ToString

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Full Data: "
        oStatsWorksheet.Cell(iCurrentRow, 2).Value = iFullData.ToString

        ' autofit columns
        For i = 0 To 1
            oStatsWorksheet.Column(1 + i).AdjustToContents()
        Next

        oExcelDocument.SaveAs(sFileName)

        Dim oFileInfo As New IO.FileInfo(sFileName)
        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Data saved to file " + oFileInfo.Name + "."))
    End Sub
#End Region
#Region "Scan"
    Private Sub SetScanner()
        ' sets default scanner to nothing
        m_CommonScanner.SelectScannerSource(String.Empty)
        SelectedScannerSource = Nothing
    End Sub
    Private Function GetScannerSources() As List(Of Tuple(Of Twain32Enumerations.ScannerSource, String, String))
        ' gets list of active scanners
        Dim oScannerSources As New List(Of Tuple(Of Twain32Enumerations.ScannerSource, String, String))

        ' gets TWAIN scanners
        Dim oTWAINScannerNameList As New List(Of String)
        Dim oTwainScanners As List(Of String) = m_CommonScanner.InitTwain
        For Each sTwainScanner As String In oTwainScanners
            Dim sScannerName As String = Trim(Replace(Replace(sTwainScanner, "TWAIN ", String.Empty), " TWAIN", String.Empty))
            oScannerSources.Add(New Tuple(Of Twain32Enumerations.ScannerSource, String, String)(Twain32Enumerations.ScannerSource.TWAIN, sTwainScanner, sScannerName))
            oTWAINScannerNameList.Add(sScannerName)
        Next

        ' gets WIA scanners
        Dim oWIADeviceManager As New WIA.DeviceManager
        For Each oDeviceInfo As WIA.DeviceInfo In oWIADeviceManager.DeviceInfos
            If oDeviceInfo.Type = WIA.WiaDeviceType.ScannerDeviceType Then
                Dim sScannerName As String = oDeviceInfo.DeviceID
                For Each oProperty As WIA.Property In oDeviceInfo.Properties
                    If oProperty.Name = "Name" Then
                        sScannerName = Trim(Replace(Replace(oProperty.Value.ToString, "WIA ", String.Empty), " WIA", String.Empty))
                    End If
                Next

                ' exclude TWAIN scanners with WIA interfaces
                Dim iContainsName As Integer = Aggregate sTWAINScannerName As String In oTWAINScannerNameList Into Count(sScannerName.Contains(sTWAINScannerName))
                If iContainsName = 0 Then
                    oScannerSources.Add(New Tuple(Of Twain32Enumerations.ScannerSource, String, String)(Twain32Enumerations.ScannerSource.WIA, oDeviceInfo.DeviceID, sScannerName))
                End If
            End If
        Next

        Return oScannerSources
    End Function
    Private Property SelectedScannerSource() As Tuple(Of Twain32Enumerations.ScannerSource, String, String)
        Get
            Return m_SelectedScannerSource
        End Get
        Set(value As Tuple(Of Twain32Enumerations.ScannerSource, String, String))
            m_SelectedScannerSource = value
            If IsNothing(m_SelectedScannerSource) Then
                ScannerScan.InnerToolTip = String.Empty

                ' set WIA and TWAIN scanner sources
                m_CommonScanner.SelectScannerSource(String.Empty)
            Else
                ScannerScan.InnerToolTip = m_SelectedScannerSource.Item3 + " (" + [Enum].GetName(GetType(Twain32Enumerations.ScannerSource), value.Item1) + ")"

                ' set WIA and TWAIN scanner sources
                If m_SelectedScannerSource.Item1 = Twain32Enumerations.ScannerSource.TWAIN Then
                    m_CommonScanner.SelectScannerSource(m_SelectedScannerSource.Item2)
                Else
                    m_CommonScanner.SelectScannerSource(String.Empty)
                End If
            End If
        End Set
    End Property
    Private Sub ReturnScannedImage(ByVal oReturnMessage As Tuple(Of Twain32Enumerations.ScanProgress, Object)) Handles m_CommonScanner.ReturnScannedImage
        ' processes scanner images
        Select Case oReturnMessage.Item1
            Case Twain32Enumerations.ScanProgress.Image
                ScanPageProgress += 1

                ' store bitmap to be retrieved for later processing
                StoreMatrix(Converter.BitmapToMatrix(oReturnMessage.Item2), CType(oReturnMessage.Item2, System.Drawing.Bitmap).HorizontalResolution, CType(oReturnMessage.Item2, System.Drawing.Bitmap).VerticalResolution)
                oReturnMessage = Nothing

                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Yellow, Date.Now, "Page " + ScanPageProgress.ToString + " scanned."))
            Case Twain32Enumerations.ScanProgress.ScanError
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Scan error: " + oReturnMessage.Item2.ToString))
            Case Twain32Enumerations.ScanProgress.Complete
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Scan complete. " + ScanPageProgress.ToString + " pages scanned."))
                ScanPageProgress = 0
        End Select
    End Sub
    Private Sub StoreMatrix(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal iHorizontalResolution As Integer, ByVal iVerticalResolution As Integer)
        ' stores bitmap in isolated storage or temporary file and leaves the path in the image store
        If Not IsNothing(oMatrix) Then
            Dim sTempFile As String = String.Empty
            If IsNothing(ImageIsolatedStorage) OrElse (Not IsolatedStorageFile.IsEnabled) Then
                ' use temporary file
                Do
                    sTempFile = IO.Path.GetTempPath + "\" + CommonFunctions.ReplaceExtension(IO.Path.GetRandomFileName, "tif")
                Loop While IO.File.Exists(sTempFile)
                IO.File.WriteAllBytes(sTempFile, Converter.MatrixToArray(oMatrix))
            Else
                ' use isolated storage
                Do
                    sTempFile = IsolatedImageDirectory + "\" + CommonFunctions.ReplaceExtension(IO.Path.GetRandomFileName, "tif")
                Loop While ImageIsolatedStorage.FileExists(sTempFile)

                Using oFileStream As New IsolatedStorageFileStream(sTempFile, IO.FileMode.Create, ImageIsolatedStorage)
                    Using oBinaryWriter As New IO.BinaryWriter(oFileStream)
                        oBinaryWriter.Write(Converter.MatrixToArray(oMatrix))
                        oBinaryWriter.Flush()
                        oBinaryWriter.Close()
                    End Using
                End Using
            End If

            ImageStore.Add(New Tuple(Of String, Integer, Integer, Integer, Integer)(sTempFile, oMatrix.Width, oMatrix.Height, iHorizontalResolution, iVerticalResolution))
        End If
    End Sub
    Private Function RetrieveMatrix() As Tuple(Of Emgu.CV.Matrix(Of Byte), Integer, Integer)
        ' gets bitmap from isolated storage 
        ' also returns horizontal and vertical resolutions
        Dim oMatrix As Emgu.CV.Matrix(Of Byte) = Nothing
        Dim oFileInfo As Tuple(Of String, Integer, Integer, Integer, Integer) = ImageStore.Take()
        If IsNothing(ImageIsolatedStorage) OrElse (Not IsolatedStorageFile.IsEnabled) Then
            ' use temporary file
            Do
                If IO.File.Exists(oFileInfo.Item1) Then
                    oMatrix = Converter.ArrayToMatrix(IO.File.ReadAllBytes(oFileInfo.Item1), oFileInfo.Item2, oFileInfo.Item3)
                    IO.File.Delete(oFileInfo.Item1)
                Else
                    oFileInfo = ImageStore.Take()
                End If
            Loop Until Not IsNothing(oMatrix)
        Else
            ' use isolated storage
            Do
                If ImageIsolatedStorage.FileExists(oFileInfo.Item1) Then
                    Using oFileStream As New IsolatedStorageFileStream(oFileInfo.Item1, IO.FileMode.Open, ImageIsolatedStorage)
                        Using oMemoryStream As New IO.MemoryStream
                            oFileStream.CopyTo(oMemoryStream)
                            oMatrix = Converter.ArrayToMatrix(oMemoryStream.ToArray(), oFileInfo.Item2, oFileInfo.Item3)
                        End Using
                    End Using
                    ImageIsolatedStorage.DeleteFile(oFileInfo.Item1)
                Else
                    oFileInfo = ImageStore.Take()
                End If
            Loop Until Not IsNothing(oMatrix)
        End If

        Return New Tuple(Of Emgu.CV.Matrix(Of Byte), Integer, Integer)(oMatrix, oFileInfo.Item4, oFileInfo.Item5)
    End Function
#End Region
#Region "Processing"
    Private Sub LoadImagesFunction(ByVal oFileNames As String())
        ' loads images for processing
        If SetRecognisers() Then
            Dim iPreAligned As Integer = 0
            Dim iUnaligned As Integer = 0

            Dim oSelectedField As Tuple(Of Integer, String) = If(IsNothing(Root.DataGridScanner.SelectedItem), Nothing, New Tuple(Of Integer, String)(CType(Root.DataGridScanner.SelectedItem, FieldDocumentStore.Field).RawBarCodes.GetHashCode, CType(Root.DataGridScanner.SelectedItem, FieldDocumentStore.Field).OrderSort))
            For i = 0 To oFileNames.Count - 1
                Dim sFileName As String = oFileNames(i)
                Dim oFileInfo As New IO.FileInfo(sFileName)
                Using oBitmap As New System.Drawing.Bitmap(sFileName)
                    oBitmap.SetResolution(CInt(oBitmap.HorizontalResolution), CInt(oBitmap.VerticalResolution))
                    Using oMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(oBitmap)
                        Dim oRectifiedImage As Tuple(Of Double, PointDouble, PointDouble, Emgu.CV.Matrix(Of Byte), PointDouble, Tuple(Of String, String, String, String)) = Aligner(oMatrix, oBitmap.HorizontalResolution, oBitmap.VerticalResolution)
                        If Not IsNothing(oRectifiedImage) AndAlso (Math.Abs(oRectifiedImage.Item1) < 0.005 And oRectifiedImage.Item2.X < 1.001 And oRectifiedImage.Item2.Y < 1.001 And oRectifiedImage.Item2.X > 0.999 And oRectifiedImage.Item2.Y > 0.999 And oRectifiedImage.Item3.X < 1 And oRectifiedImage.Item3.Y < 1 And oRectifiedImage.Item5.X < SpacingSmall And oRectifiedImage.Item5.Y < SpacingSmall) Then
                            ' prealigned
                            StoreMatrix(oMatrix, oBitmap.HorizontalResolution, oBitmap.VerticalResolution)
                            iPreAligned += 1
                        Else
                            ' not aligned
                            StoreMatrix(oRectifiedImage.Item4, oBitmap.HorizontalResolution, oBitmap.VerticalResolution)
                            iUnaligned += 1
                        End If

                        oRectifiedImage.Item4.Dispose()
                        oRectifiedImage = Nothing
                    End Using
                End Using

                CommonFunctions.ClearMemory()
            Next

            ' restore location
            If IsNothing(oSelectedField) Then
                ' deselect
                Root.DataGridScanner.SelectedItem = Nothing
            Else
                Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                Dim oFields As List(Of FieldDocumentStore.Field) = (From oField In oScannerCollection Where oField.RawBarCodes.GetHashCode = oSelectedField.Item1 AndAlso oField.OrderSort = oSelectedField.Item2 Select oField).ToList
                If oFields.Count > 0 Then
                    Root.DataGridScanner.SelectedItem = oFields.First
                Else
                    ' deselect
                    Root.DataGridScanner.SelectedItem = Nothing
                End If
            End If

            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, iPreAligned.ToString + " pre-aligned and " + iUnaligned.ToString + " unaligned images processed."))
        Else
            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Unable to load recognisers."))
        End If
    End Sub
    Private Sub ProcessImage(ByVal ct As CancellationToken)
        ' background process to process scanned images from image store
        ' continue looping until cancelled
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Do While True
            If ct.IsCancellationRequested Then
                Exit Do
            Else
                ' add plug-in background processing here
                ' retrieve matrix blocks until an image is present in the store
                Dim oMatrixInfo As Tuple(Of Emgu.CV.Matrix(Of Byte), Integer, Integer) = RetrieveMatrix()
                Dim oRectifiedResult As Tuple(Of Double, PointDouble, PointDouble, Emgu.CV.Matrix(Of Byte), PointDouble, Tuple(Of String, String, String, String)) = Aligner(oMatrixInfo.Item1, oMatrixInfo.Item2, oMatrixInfo.Item3)
                If Not IsNothing(oRectifiedResult) Then
                    Dim oRectifiedMatrix As Emgu.CV.Matrix(Of Byte) = oRectifiedResult.Item4
                    Dim oProcessedBarcode As Tuple(Of String, String, String, String) = oRectifiedResult.Item6

                    Dim oIndexList As List(Of Integer) = (From iIndex As Integer In Enumerable.Range(0, oScannerCollection.FieldDocumentStore.FieldCollectionStore.Count) Where oScannerCollection.FieldDocumentStore.FieldCollectionStore(iIndex).FormTitle.Replace(SeparatorChar, String.Empty).GetHashCode.ToString.Equals(oProcessedBarcode.Item1) AndAlso oScannerCollection.FieldDocumentStore.FieldCollectionStore(iIndex).SubjectName.GetHashCode.ToString.Equals(oProcessedBarcode.Item3) Select iIndex Distinct).ToList
                    Dim sFormTitle As String = oScannerCollection.FieldDocumentStore.FieldCollectionStore(oIndexList.First).FormTitle + " - Page " + oProcessedBarcode.Item2 + "-" + oScannerCollection.FieldDocumentStore.FieldCollectionStore(oIndexList.First).BarCodes.Count.ToString
                    Dim sSubjectName As String = oScannerCollection.FieldDocumentStore.FieldCollectionStore(oIndexList.First).SubjectName

                    If oSettings.DefaultSave <> String.Empty Then
                        Dim sNumberingText As String = String.Empty
                        If oIndexList.Count = 1 Then
                            sNumberingText = oIndexList.First.ToString.PadLeft(oScannerCollection.FieldDocumentStore.FieldCollectionStore.Count.ToString.Length, "0") + "_"
                        ElseIf oIndexList.Count > 1 Then
                            sNumberingText = oIndexList(Val(oProcessedBarcode.Item4)).ToString.PadLeft(oScannerCollection.FieldDocumentStore.FieldCollectionStore.Count.ToString.Length, "0") + "_"
                        End If

                        Dim sFileName As String = oSettings.DefaultSave + "\" + sNumberingText + sSubjectName + If(oProcessedBarcode.Item4 = String.Empty, String.Empty, " (" + oProcessedBarcode.Item4 + ")")
                        If sFormTitle <> String.Empty Then
                            sFileName += "_" + sFormTitle
                        End If
                        sFileName += ".tif"

                        ' save image to default save directory
                        CommonFunctions.SaveBitmap(sFileName, Converter.MatToBitmap(oRectifiedMatrix.Mat, oMatrixInfo.Item2, oMatrixInfo.Item3), False)
                    End If

                    ' detect marks
                    Using oUpdateSuspender As New UpdateSuspender(True)
                        Using oDeferRefresh As New DeferRefresh()
                            ProcessFields(oRectifiedMatrix, oProcessedBarcode, oMatrixInfo.Item2)
                            DetectMarks()
                        End Using
                    End Using

                    ' refresh only if the loaded page is from the selected subject
                    If oProcessedBarcode.Item3 = oScannerCollection.FieldDocumentStore.FieldCollectionStore(oScannerCollection.SelectedCollection).SubjectName.GetHashCode.ToString Then
                        Dim oAction As Action = Sub()
                                                    DataGridScannerChanged()
                                                End Sub
                        CommonFunctions.DispatcherInvoke(UIDispatcher, oAction)
                    End If

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, sFormTitle + ", " + sSubjectName + " processed."))
                End If
            End If
        Loop
    End Sub
    Private Sub ProcessFields(ByVal oRectifiedMatrix As Emgu.CV.Matrix(Of Byte), ByVal oProcessedBarcode As Tuple(Of String, String, String, String), ByVal iResolution As Integer)
        ' processes images according to the input field collection and cuts out the relevant bitmaps
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Dim oMatrixStore As FieldDocumentStore.MatrixStore = oScannerCollection.FieldDocumentStore.FieldMatrixStore

        For Each oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore
            Dim sReplacedTitle As String = oFieldCollection.FormTitle.Replace(SeparatorChar, Nothing)
            Dim sSubjectName As String = oFieldCollection.SubjectName
            Dim sAppendText As String = oFieldCollection.AppendText

            For Each oField In oFieldCollection.Fields
                Dim sCurrentPage As String = oField.PageNumber.ToString
                If oProcessedBarcode.Item2 = sCurrentPage AndAlso ((oProcessedBarcode.Item1 = sReplacedTitle.GetHashCode.ToString And oProcessedBarcode.Item3 = sSubjectName.GetHashCode.ToString And oProcessedBarcode.Item4 = sAppendText) OrElse (oProcessedBarcode.Item1 = sReplacedTitle And oProcessedBarcode.Item3 = sSubjectName And oProcessedBarcode.Item4 = sAppendText)) Then
                    Using oTabletFilled As Emgu.CV.Matrix(Of Byte) = PDFHelper.GetCheckMatrix(iResolution, Enumerations.TabletType.Filled)
                        For i = 0 To oField.Images.Count - 1
                            If oField.Images(i).Item2.Equals(Guid.Empty) Then
                                Dim oMatchMatrix As Emgu.CV.Matrix(Of Byte) = Nothing
                                Dim oMaskMatrix As Emgu.CV.Matrix(Of Byte) = Nothing

                                Select Case oField.FieldType
                                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                                        Dim sContentText As String = String.Empty
                                        Select Case oField.TabletContent
                                            Case Enumerations.TabletContentEnum.Number
                                                sContentText = (i + 1 + oField.TabletStart).ToString
                                            Case Enumerations.TabletContentEnum.Letter
                                                sContentText = Converter.ConvertNumberToLetter((i + oField.TabletStart), True)
                                        End Select
                                        oMatchMatrix = PDFHelper.GetChoiceMatrix(sContentText, iResolution, Enumerations.TabletType.Empty)
                                    Case Enumerations.FieldTypeEnum.BoxChoice
                                        Select Case oField.Images(i).Item5
                                            Case -1
                                                Dim oReturnTuple As Tuple(Of Emgu.CV.Matrix(Of Byte), Emgu.CV.Matrix(Of Byte)) = PDFHelper.GetHandwritingMatrix(String.Empty, iResolution, SpacingSmall, 3)
                                                oMatchMatrix = oReturnTuple.Item1
                                                oMaskMatrix = oReturnTuple.Item2
                                            Case Else
                                                Dim sContentText As String = oField.Images(i).Item5.ToString
                                                oMatchMatrix = PDFHelper.GetChoiceMatrix(sContentText, iResolution, Enumerations.TabletType.Empty)
                                        End Select
                                    Case Enumerations.FieldTypeEnum.Handwriting
                                        Dim oReturnTuple As Tuple(Of Emgu.CV.Matrix(Of Byte), Emgu.CV.Matrix(Of Byte)) = PDFHelper.GetHandwritingMatrix(If(oField.Images(i).Item6, oField.Images(i).Item3, String.Empty), iResolution, SpacingSmall, 3)
                                        oMatchMatrix = oReturnTuple.Item1
                                        oMaskMatrix = oReturnTuple.Item2
                                    Case Enumerations.FieldTypeEnum.Free
                                        oMatchMatrix = PDFHelper.GetFreeMatrix(New Size(oField.Images(i).Item1.Width, oField.Images(i).Item1.Height), iResolution, SpacingSmall)
                                End Select

                                If oField.Images(i).Item6 Then
                                    oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, Nothing, oField.Images(i).Item3, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, 0, New Tuple(Of Single)(oField.Images(i).Rest.Item1)), Nothing, oMatrixStore)
                                Else
                                    Using oFieldMatrix As Emgu.CV.Matrix(Of Byte) = ExtractFieldMatrix(oRectifiedMatrix, iResolution, oField.Images(i).Item1, oMatchMatrix, oMaskMatrix)
                                        oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, Guid.NewGuid, String.Empty, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, iResolution, New Tuple(Of Single)(oField.Images(i).Rest.Item1)), oFieldMatrix, oMatrixStore)
                                    End Using
                                End If

                                ' clean up
                                oMatchMatrix.Dispose()
                            End If
                        Next

                        ' remove extension marks
                        ScreenAdjacentRect(oField)
                    End Using
                End If
            Next
        Next
    End Sub
    Private Sub DetectMarks()
        ' routine to detect marks from the stored bitmaps
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Dim oMatrixStore As FieldDocumentStore.MatrixStore = oScannerCollection.FieldDocumentStore.FieldMatrixStore

        For Each oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore
            For Each oField As FieldDocumentStore.Field In oFieldCollection.Fields
                If oField.Images.Count > 0 Then
                    Select Case oField.FieldType
                        Case Enumerations.FieldTypeEnum.BoxChoice
                            ' check to see if detected marks already present
                            Dim oBoxImageList As List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))) = (From oImage In oField.Images Where oImage.Item5 = -1 Select oImage).ToList
                            Dim iEmpty As Integer = Aggregate oBoxImage In oBoxImageList From oImage In oField.Images Where oBoxImage.Item3 = String.Empty AndAlso oBoxImage.Item4 = oImage.Item4 AndAlso (Not oImage.Item2.Equals(Guid.Empty)) Into Count
                            If iEmpty > 0 Then
                                Dim oBoxTabletDictionary As Dictionary(Of Integer, List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))) = (From oBoxImage In oBoxImageList Select New KeyValuePair(Of Integer, List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))))(oBoxImage.Item4, (From oImage In oField.Images Where oImage.Item5 >= 0 And oImage.Item4 = oBoxImage.Item4 Order By oImage.Item5 Ascending Select oImage).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)

                                For i = 0 To oBoxTabletDictionary.Count - 1
                                    If oBoxTabletDictionary.Values(i).Count > 0 Then
                                        ' check for images to be present for all tablets
                                        Dim iMissingImageCount As Integer = Aggregate oImage In oBoxTabletDictionary.Values(i) Into Count(oImage.Item2.Equals(Guid.Empty))
                                        If iMissingImageCount = 0 Then
                                            Dim oDetectedChoiceList As New List(Of Tuple(Of String, Integer))
                                            For j = 0 To oBoxTabletDictionary.Values(i).Count - 1
                                                oDetectedChoiceList.Add(DetectChoice(oField.TabletContent, oBoxTabletDictionary.Values(i), j, -1))
                                            Next

                                            Dim oOrderedChoiceList As List(Of Tuple(Of Integer, String, Integer)) = (From iIndex As Integer In Enumerable.Range(0, oDetectedChoiceList.Count) Order By oDetectedChoiceList(iIndex).Item2 Descending Select New Tuple(Of Integer, String, Integer)(iIndex, oDetectedChoiceList(iIndex).Item1, oDetectedChoiceList(iIndex).Item2)).ToList
                                            Dim iCurrent As Integer = i
                                            Dim iSelectedImageIndex As Integer = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iCurrent And oField.Images(iIndex).Item5 = -1 Select iIndex).First
                                            oField.SetImage(iSelectedImageIndex, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(iSelectedImageIndex).Item1, oField.Images(iSelectedImageIndex).Item2, If(Trim(oOrderedChoiceList.First.Item2) = String.Empty, " ", oOrderedChoiceList.First.Item1.ToString), oField.Images(iSelectedImageIndex).Item4, oField.Images(iSelectedImageIndex).Item5, oField.Images(iSelectedImageIndex).Item6, oField.Images(iSelectedImageIndex).Item7, New Tuple(Of Single)(0)))
                                        End If
                                    End If
                                Next
                            End If
                        Case Else
                            For i = 0 To oField.Images.Count - 1
                                If oField.Images(i).Item2.Equals(Guid.Empty) Then
                                    If oField.Images(i).Item6 Then
                                        oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, Nothing, oField.Images(i).Item3, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, 0, New Tuple(Of Single)(oField.Images(i).Rest.Item1)), Nothing, oMatrixStore)
                                    Else
                                        oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, Nothing, String.Empty, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, 0, New Tuple(Of Single)(oField.Images(i).Rest.Item1)), Nothing, oMatrixStore)
                                    End If
                                Else
                                    Select Case oField.FieldType
                                        Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                                            If oField.Images(i).Item3 = String.Empty Then
                                                Dim oDetectedChoice As Tuple(Of String, Integer) = DetectChoice(oField.TabletContent, oField.Images, i, oField.TabletStart)
                                                oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, oField.Images(i).Item2, oDetectedChoice.Item1, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, oField.Images(i).Item7, New Tuple(Of Single)(oDetectedChoice.Item2)))
                                            End If
                                        Case Enumerations.FieldTypeEnum.Handwriting
                                            If (Not oField.Images(i).Item6) AndAlso oField.Images(i).Item3 = String.Empty Then
                                                Dim oDetectedChar As List(Of Tuple(Of String, Double)) = DetectHandwriting(oField.Images(i), oField.CharacterASCII, oHOGDescriptor)
                                                Dim sDetectedText As String = " "
                                                If oDetectedChar.Count = 1 Then
                                                    sDetectedText = oDetectedChar.First.Item1
                                                ElseIf oDetectedChar.Count > 1 Then
                                                    Const fDetectedThreshold As Double = 0.125
                                                    sDetectedText = String.Join(String.Empty, From oChar As Tuple(Of String, Double) In oDetectedChar Where oChar.Item2 > oDetectedChar.First.Item2 * fDetectedThreshold Select oChar.Item1)
                                                End If

                                                oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, oField.Images(i).Item2, sDetectedText, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, oField.Images(i).Item7, New Tuple(Of Single)(oField.Images(i).Rest.Item1)))
                                            End If
                                        Case Enumerations.FieldTypeEnum.Free
                                            ' no detection routine set
                                            oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, oField.Images(i).Item2, String.Empty, oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, oField.Images(i).Item7, New Tuple(Of Single)(oField.Images(i).Rest.Item1)))
                                    End Select
                                End If
                            Next
                    End Select

                    ' post-processing of fields
                    ' scans through a handwriting field row by row to process the content
                    If oField.FieldType = Enumerations.FieldTypeEnum.Handwriting Then
                        Dim oRowSegmentList As New List(Of List(Of Integer))
                        Dim oRowList As List(Of Integer) = (From oImage In oField.Images Where Not oImage.Item6 Select oImage.Item4 Distinct).ToList
                        For Each iRow As Integer In oRowList
                            Dim iCurrentRow As Integer = iRow
                            Dim oCurrentColMaxList As List(Of Integer) = (From oImage In oField.Images Where oImage.Item4 = iCurrentRow And Not oImage.Item6 Select oImage.Item5).ToList
                            If oCurrentColMaxList.Count > 0 Then
                                Dim iCurrentColMax As Integer = oCurrentColMaxList.Max
                                Dim oSeparatorList As List(Of Integer) = (From oImage In oField.Images Where oImage.Item4 = iCurrentRow And oImage.Item6 Select oImage.Item5).ToList

                                oRowSegmentList.Clear()
                                Dim oRange As List(Of Integer) = Enumerable.Range(0, iCurrentColMax + 1).ToList
                                Dim oCurrentList As New List(Of Integer)
                                For i = 0 To iCurrentColMax
                                    ' check for separator
                                    If oSeparatorList.Contains(i) Then
                                        oRowSegmentList.Add(New List(Of Integer)(oCurrentList))
                                        oCurrentList.Clear()
                                    Else
                                        ' add the index of the image
                                        Dim iCurrentCol As Integer = i
                                        oCurrentList.Add((From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iCurrentRow AndAlso oField.Images(iIndex).Item5 = iCurrentCol Select iIndex).First)
                                    End If
                                Next
                                oRowSegmentList.Add(oCurrentList)
                            End If

                            ' run through segment by segment
                            Dim oCurrentWord As New List(Of Integer)
                            For Each oRowSegment As List(Of Integer) In oRowSegmentList
                                oCurrentWord.Clear()
                                For Each iImageIndex As Integer In oRowSegment
                                    If oField.Images(iImageIndex).Item3 = String.Empty Then
                                        ProcessCurrentWord(oCurrentWord, oField)
                                    Else
                                        oCurrentWord.Add(iImageIndex)
                                    End If
                                Next
                                ProcessCurrentWord(oCurrentWord, oField)
                            Next
                        Next
                    End If

                    ' extract a dictionary of related choice fields
                    If oField.TabletSingleChoiceOnly AndAlso (oField.FieldType = Enumerations.FieldTypeEnum.Choice Or oField.FieldType = Enumerations.FieldTypeEnum.ChoiceVertical Or oField.FieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ) Then
                        Dim fMaxDiagonal As Single = (From oImage In oField.Images Select oImage.Rest.Item1).Max
                        Dim iMaxDiagonalIndex As Integer = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Rest.Item1 = fMaxDiagonal Select iIndex).First
                        For i = 0 To oField.Images.Count - 1
                            If oField.Images(i).Item2.Equals(Guid.Empty) OrElse i <> iMaxDiagonalIndex Then
                                oField.SetImage(i, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(i).Item1, oField.Images(i).Item2, " ", oField.Images(i).Item4, oField.Images(i).Item5, oField.Images(i).Item6, oField.Images(i).Item7, New Tuple(Of Single)(0)))
                            End If
                        Next
                    End If

                    ' set marks
                    Select Case oField.FieldType
                        Case Enumerations.FieldTypeEnum.BoxChoice
                            Dim oBoxImageList As List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))) = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item5 = -1 Select oField.Images(iIndex)).ToList
                            oField.MarkText = (From iKeyIndex As Integer In Enumerable.Range(0, oBoxImageList.Count) Select New Tuple(Of String, String, String)(oBoxImageList(iKeyIndex).Item3, String.Empty, String.Empty)).ToList
                        Case Else
                            oField.MarkText = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Select New Tuple(Of String, String, String)(oField.Images(iIndex).Item3, If(oField.Images(iIndex).Item6, oField.Images(iIndex).Item3, String.Empty), If(oField.Images(iIndex).Item6, oField.Images(iIndex).Item3, String.Empty))).ToList
                    End Select

                    ' for non-critical fields, copy the first mark to the second and third
                    Select Case oField.FieldType
                        Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                            For j = 0 To oField.MarkCount - 1
                                If oField.Critical Then
                                    If oField.MarkChoice1(j) = oField.MarkChoice0(j) Then
                                        oField.MarkChoice2(j) = oField.MarkChoice0(j)
                                    Else
                                        oField.MarkChoice2(j) = False
                                    End If
                                Else
                                    oField.MarkChoice1(j) = oField.MarkChoice0(j)
                                    oField.MarkChoice2(j) = oField.MarkChoice0(j)
                                End If
                            Next
                        Case Enumerations.FieldTypeEnum.Handwriting
                            For j = 0 To oField.MarkCount - 1
                                If oField.Critical Then
                                    If oField.MarkHandwriting1(j) = oField.MarkHandwriting0(j) Then
                                        oField.MarkHandwriting2(j) = oField.MarkHandwriting0(j)
                                    Else
                                        oField.MarkHandwriting2(j) = String.Empty
                                    End If
                                Else
                                    oField.MarkHandwriting1(j) = oField.MarkHandwriting0(j)
                                    oField.MarkHandwriting2(j) = oField.MarkHandwriting0(j)
                                End If
                            Next
                        Case Enumerations.FieldTypeEnum.BoxChoice
                            For j = 0 To oField.MarkCount - 1
                                If oField.Critical Then
                                    If oField.MarkBoxChoice1(j) = oField.MarkBoxChoice0(j) Then
                                        oField.MarkBoxChoice2(j) = oField.MarkBoxChoice0(j)
                                    Else
                                        oField.MarkBoxChoice2(j) = String.Empty
                                    End If
                                Else
                                    oField.MarkBoxChoice1(j) = oField.MarkBoxChoice0(j)
                                    oField.MarkBoxChoice2(j) = oField.MarkBoxChoice0(j)
                                End If
                            Next
                        Case Enumerations.FieldTypeEnum.Free
                        Case Else
                    End Select
                End If
            Next
        Next
    End Sub
    Private Shared Function DetectChoice(ByVal oTabletContent As Enumerations.TabletContentEnum, ByVal oImages As List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))), ByVal i As Integer, ByVal iTabletStart As Integer) As Tuple(Of String, Integer)
#Disable Warning BC42099 ' Unused local constant
        Const PixelSumMax300 As Integer = 727033
        Const DotCountMax300 As Integer = 427
#Enable Warning BC42099 ' Unused local constant
        Const iTabletThresholdPixel As Integer = 2500
        Const iDotCountThreshold As Integer = 45
        Const iLowCountThreshold As Integer = 20
        Const iMediumCountThreshold As Integer = 30
        Const iHighCountThreshold As Integer = 80
        Const fLowCountFraction As Single = 0.8
        Const fMediumCountFraction As Single = 0.5
        Const fHighCountFraction As Single = 0.4
        Const fSlopeThreshold As Single = 0.088

        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Dim oMatrixStore As FieldDocumentStore.MatrixStore = oScannerCollection.FieldDocumentStore.FieldMatrixStore
        Dim sContentText As String = String.Empty
        Select Case oTabletContent
            Case Enumerations.TabletContentEnum.Number
                sContentText = (i + 1 + iTabletStart).ToString
            Case Enumerations.TabletContentEnum.Letter
                sContentText = Converter.ConvertNumberToLetter(i, True)
        End Select

        Using oChoiceMaskMatrix As Emgu.CV.Matrix(Of Byte) = PDFHelper.GetChoiceMatrix(sContentText, oImages(i).Item7, Enumerations.TabletType.EmptyBlack,, False)
            ' erode the template
            Emgu.CV.CvInvoke.Erode(oChoiceMaskMatrix, oChoiceMaskMatrix, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Rectangle, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MaxValue))

            ' invert template
            oChoiceMaskMatrix.SubR(Byte.MaxValue).CopyTo(oChoiceMaskMatrix)

            Using oImageMatrix As Emgu.CV.Matrix(Of Byte) = oMatrixStore.GetMatrix(oImages(i).Item2)
                ' overlay the template on the image
                Emgu.CV.CvInvoke.Normalize(oImageMatrix, oImageMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                Emgu.CV.CvInvoke.Threshold(oChoiceMaskMatrix, oChoiceMaskMatrix, Byte.MinValue + 1, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary)
                Emgu.CV.CvInvoke.Max(oImageMatrix, oChoiceMaskMatrix, oImageMatrix)
                Emgu.CV.CvInvoke.Dilate(oImageMatrix, oImageMatrix, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                Dim iPixelSum As Integer = oImageMatrix.SubR(Byte.MaxValue).Sum

                ' get dot count
                Emgu.CV.CvInvoke.Normalize(oImageMatrix, oImageMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                Emgu.CV.CvInvoke.AdaptiveThreshold(oImageMatrix, oImageMatrix, Byte.MaxValue, Emgu.CV.CvEnum.AdaptiveThresholdType.GaussianC, Emgu.CV.CvEnum.ThresholdType.BinaryInv, 11, 56)
                Using oMaskMatrix As New Emgu.CV.Matrix(Of Byte)(oImageMatrix.Size)
                    Emgu.CV.CvInvoke.Normalize(oImageMatrix, oImageMatrix, 0, 1, Emgu.CV.CvEnum.NormType.MinMax)
                    Emgu.CV.CvInvoke.Filter2D(oImageMatrix, oMaskMatrix, oStructuringElement, New System.Drawing.Point(-1, -1))
                    Emgu.CV.CvInvoke.Threshold(oMaskMatrix, oMaskMatrix, 0, 1, Emgu.CV.CvEnum.ThresholdType.BinaryInv)
                    Emgu.CV.CvInvoke.BitwiseAnd(oImageMatrix, oMaskMatrix, oMaskMatrix)
                    Emgu.CV.CvInvoke.Subtract(oImageMatrix, oMaskMatrix, oImageMatrix)
                    Emgu.CV.CvInvoke.Normalize(oImageMatrix, oImageMatrix, 0, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                End Using
                Dim iDotCount As Integer = Emgu.CV.CvInvoke.CountNonZero(oImageMatrix)

                If iPixelSum > iTabletThresholdPixel AndAlso iDotCount > iDotCountThreshold Then
                    Return New Tuple(Of String, Integer)("X", iPixelSum)
                ElseIf iPixelSum > (iTabletThresholdPixel / 2) AndAlso iDotCount > (iDotCountThreshold / 3) Then
                    ' use ransac line estimator to detect lines
                    Using oLineMatrix As Emgu.CV.Matrix(Of Byte) = oMatrixStore.GetMatrix(oImages(i).Item2)
                        Emgu.CV.CvInvoke.Normalize(oLineMatrix, oLineMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                        Emgu.CV.CvInvoke.Max(oLineMatrix, oChoiceMaskMatrix, oLineMatrix)
                        Emgu.CV.CvInvoke.Dilate(oLineMatrix, oLineMatrix, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                        Emgu.CV.CvInvoke.AdaptiveThreshold(oLineMatrix, oLineMatrix, Byte.MaxValue, Emgu.CV.CvEnum.AdaptiveThresholdType.GaussianC, Emgu.CV.CvEnum.ThresholdType.BinaryInv, 11, 8)

                        Dim oPointBag As New ConcurrentBag(Of Accord.IntPoint)

                        Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                                    Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                                    Dim oImageData As Byte(,) = CType(oParam.Item4, Emgu.CV.Matrix(Of Byte)).Data
                                                                    Dim localwidth As Integer = CType(oParam.Item5, Integer)

                                                                    For y = 0 To oParam.Item2 - 1
                                                                        Dim localY As Integer = y + oParam.Item1
                                                                        For x = 0 To localwidth - 1
                                                                            If oImageData(localY, x) <> 0 Then
                                                                                oPointBag.Add(New Accord.IntPoint(x, localY))
                                                                            End If
                                                                        Next
                                                                    Next
                                                                End Sub
                        CommonFunctions.ParallelRun(oLineMatrix.Height, oLineMatrix, oLineMatrix.Width, Nothing, Nothing, TaskDelegate)

                        oRansacLine.Estimate(oPointBag)

                        Select Case oRansacLine.Inliers.Count
                            Case < iLowCountThreshold
                                ' reject
                                Return New Tuple(Of String, Integer)(" ", 0)
                            Case > iHighCountThreshold
                                If oRansacLine.Inliers.Count / oPointBag.Count > fHighCountFraction AndAlso GetAxisDifferential(oPointBag) > fSlopeThreshold Then
                                    Return New Tuple(Of String, Integer)("X", iPixelSum)
                                Else
                                    Return New Tuple(Of String, Integer)(" ", 0)
                                End If
                            Case iLowCountThreshold To iMediumCountThreshold
                                Dim fInlinerThreshold As Single = fLowCountFraction + ((oRansacLine.Inliers.Count - iLowCountThreshold) * (fMediumCountFraction - fLowCountFraction) / (iMediumCountThreshold - iLowCountThreshold))
                                If oRansacLine.Inliers.Count / oPointBag.Count > fInlinerThreshold AndAlso GetAxisDifferential(oPointBag) > fSlopeThreshold Then
                                    Return New Tuple(Of String, Integer)("X", iPixelSum)
                                Else
                                    Return New Tuple(Of String, Integer)(" ", 0)
                                End If
                            Case Else
                                Dim fInlinerThreshold As Single = fMediumCountFraction + ((oRansacLine.Inliers.Count - iMediumCountThreshold) * (fHighCountFraction - fMediumCountFraction) / (iHighCountThreshold - iMediumCountThreshold))
                                If oRansacLine.Inliers.Count / oPointBag.Count > fInlinerThreshold AndAlso GetAxisDifferential(oPointBag) > fSlopeThreshold Then
                                    Return New Tuple(Of String, Integer)("X", iPixelSum)
                                Else
                                    Return New Tuple(Of String, Integer)(" ", 0)
                                End If
                        End Select
                    End Using
                Else
                    Return New Tuple(Of String, Integer)(" ", 0)
                End If
            End Using
        End Using
    End Function
    Private Shared Function GetAxisDifferential(ByVal oPointBag As ConcurrentBag(Of Accord.IntPoint)) As Double
        ' gets the slope of a filtered set of points, and compare it with the horizontal and vertical axes
        ' return the minimum slope between the line and one of the axes
        Dim oInputX As New List(Of Double)
        Dim oInputY As New List(Of Double)
        For Each oPoint As Accord.IntPoint In oPointBag
            oInputX.Add(oPoint.X)
            oInputY.Add(oPoint.Y)
        Next

        Dim oOLS As New Accord.Statistics.Models.Regression.Linear.OrdinaryLeastSquares
        Dim oRegression As Accord.Statistics.Models.Regression.Linear.SimpleLinearRegression = Nothing
        Try
            oRegression = oOLS.Learn(oInputX.ToArray, oInputY.ToArray)
        Catch ex As InvalidOperationException
            ' infinite slope
            Return 0
        End Try

        If Math.Abs(oRegression.Slope) > 1 Then
            Return 1 / Math.Abs(oRegression.Slope)
        Else
            Return Math.Abs(oRegression.Slope)
        End If
    End Function
    Public Shared Function ExtractFieldMatrix(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal iResolution As Integer, ByVal oRect As Rect, ByVal oMatchMatrix As Emgu.CV.Matrix(Of Byte), Optional ByVal oMaskMatrix As Emgu.CV.Matrix(Of Byte) = Nothing) As Emgu.CV.Matrix(Of Byte)
        ' extracts input field bitmap based on supplied boundaries
        ' the matched bitmap contains the bitmap to match, while the mask bitmap is overlayed on the extracted image before matching (Max)
        ' if this is present, it is used to refine the location of the extracted field bitmap
        Dim fPixelMargin As Double = SpacingSmall
        Dim fPixelLeft As Double = oRect.Left * iResolution / 72
        Dim fPixelTop As Double = oRect.Top * iResolution / 72
        Dim fPixelWidth As Double = oRect.Width * iResolution / 72
        Dim fPixelHeight As Double = oRect.Height * iResolution / 72
        Dim fLeft As Double = fPixelLeft - fPixelMargin
        Dim fTop As Double = fPixelTop - fPixelMargin
        Dim fWidth As Double = fPixelWidth + (fPixelMargin * 2)
        Dim fHeight As Double = fPixelHeight + (fPixelMargin * 2)

        If IsNothing(oMatchMatrix) Then
            Dim oExtractedMatrix As New Emgu.CV.Matrix(Of Byte)(fHeight, fWidth)
            Dim oRectangle As New System.Drawing.RectangleF(fLeft, fTop, oExtractedMatrix.Width, oExtractedMatrix.Height)
            DrawMatrixSubPix(oMatrix, oExtractedMatrix, oRectangle)
            Return oExtractedMatrix
        Else
            Dim oMatchLocationList As New List(Of Tuple(Of Double, Double, Double, Emgu.CV.Matrix(Of Byte)))
            Using oExtractedMatchMatrix As New Emgu.CV.Matrix(Of Byte)(fHeight + (SpacingSmall * 2), fWidth + (SpacingSmall * 2))
                Dim oRectangle As New System.Drawing.RectangleF(fPixelLeft - fPixelMargin - SpacingSmall, fPixelTop - fPixelMargin - SpacingSmall, oExtractedMatchMatrix.Width, oExtractedMatchMatrix.Height)
                DrawMatrixSubPix(oMatrix, oExtractedMatchMatrix, oRectangle)

                If (Not IsNothing(oMaskMatrix)) AndAlso (oMaskMatrix.Width < oExtractedMatchMatrix.Width And oMaskMatrix.Height < oExtractedMatchMatrix.Height) Then
                    Using oExtractedSubMatrix As Emgu.CV.Matrix(Of Byte) = oExtractedMatchMatrix.GetSubRect(New System.Drawing.Rectangle((oExtractedMatchMatrix.Width - oMaskMatrix.Width) / 2, (oExtractedMatchMatrix.Height - oMaskMatrix.Height) / 2, oMaskMatrix.Width, oMaskMatrix.Height))
                        Emgu.CV.CvInvoke.Max(oExtractedSubMatrix, oMaskMatrix, oExtractedSubMatrix)
                    End Using
                End If

                Dim oLocatedImage As List(Of Tuple(Of Double, PointDouble)) = SimpleLocateImage(oExtractedMatchMatrix, oMatchMatrix, 1)
                If oLocatedImage.Count > 0 Then
                    oMatchLocationList.Add(New Tuple(Of Double, Double, Double, Emgu.CV.Matrix(Of Byte))(oLocatedImage.First.Item1, oLocatedImage.First.Item2.X - oExtractedMatchMatrix.Width / 2, oLocatedImage.First.Item2.Y - oExtractedMatchMatrix.Height / 2, oMatchMatrix))
                End If
            End Using

            Dim oExtractedMatrix As New Emgu.CV.Matrix(Of Byte)(fHeight, fWidth)
            If oMatchLocationList.Count > 0 Then
                Dim oTopMatch As Tuple(Of Double, Double, Double, Emgu.CV.Matrix(Of Byte)) = (From oMatch As Tuple(Of Double, Double, Double, Emgu.CV.Matrix(Of Byte)) In oMatchLocationList Order By oMatch.Item1 Descending Select oMatch).First
                Dim oRectangle As New System.Drawing.RectangleF(fLeft + oTopMatch.Item2, fTop + oTopMatch.Item3, oExtractedMatrix.Width, oExtractedMatrix.Height)
                DrawMatrixSubPix(oMatrix, oExtractedMatrix, oRectangle)
            Else
                Dim oRectangle As New System.Drawing.RectangleF(fLeft, fTop, oExtractedMatrix.Width, oExtractedMatrix.Height)
                DrawMatrixSubPix(oMatrix, oExtractedMatrix, oRectangle)
            End If

            Return oExtractedMatrix
        End If
    End Function
    Public Shared Sub DrawMatrixSubPix(ByVal oSourceMatrix As Emgu.CV.Matrix(Of Byte), ByRef oDestinationMatrix As Emgu.CV.Matrix(Of Byte), ByVal oRectangle As System.Drawing.RectangleF)
        ' copies the source matrix to the destination matrix rectangle with subpixel accuracy
        Emgu.CV.CvInvoke.GetRectSubPix(oSourceMatrix, New System.Drawing.Size(oRectangle.Width, oRectangle.Height), New System.Drawing.PointF(oRectangle.Left + (oRectangle.Width - 1) / 2, oRectangle.Top + (oRectangle.Height - 1) / 2), oDestinationMatrix)
    End Sub
    Public Class DeferRefresh
        Implements IDisposable

        Sub New()
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Shadows Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    Dim oAction As Action = Sub()
                                                Dim oCollectionViewSource As Data.CollectionViewSource = Root.GridMain.Resources("cvsScannnerCollection")
                                                oCollectionViewSource.DeferRefresh.Dispose()
                                            End Sub
                    CommonFunctions.DispatcherInvoke(UIDispatcher, oAction)
                End If
            End If
            disposedValue = True
        End Sub
        Public Shadows Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
#End Region
#Region "ScreenAdjacentRect"
    Public Shared Sub ScreenAdjacentRect(ByRef oField As FieldDocumentStore.Field)
        ' screens through a handwriting field to remove extension marks from adjacent fields
        ' iStep is zero-based and denotes in order left, right, top, bottom
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Dim oMatrixStore As FieldDocumentStore.MatrixStore = oScannerCollection.FieldDocumentStore.FieldMatrixStore
        Const iBorderWidth As Integer = 9
        Const iScaleFactor As Integer = 5

        If (oField.FieldType = Enumerations.FieldTypeEnum.Handwriting) AndAlso oField.Images.Count >= 2 Then
            Dim iImagesPresent As Integer = Aggregate oImage In oField.Images Where Not IsNothing(oImage.Item2) Into Count
            Dim iImagesTotal As Integer = Aggregate oImage In oField.Images Where Not oImage.Item6 Into Count
            If iImagesPresent = iImagesTotal Then
                Dim iWidth As Integer = 0
                Dim iHeight As Integer = 0
                If oField.FieldType = Enumerations.FieldTypeEnum.Handwriting Then
                    iWidth = oMatrixStore.GetMatrixWidth(oField.Images.First.Item2)
                    iHeight = oMatrixStore.GetMatrixHeight(oField.Images.First.Item2)
                Else
                    iWidth = (From oImage In oField.Images Where (Not IsNothing(oImage.Item2)) AndAlso oImage.Item5 = -1 Select oMatrixStore.GetMatrixWidth(oImage.Item2)).First
                    iHeight = (From oImage In oField.Images Where (Not IsNothing(oImage.Item2)) AndAlso oImage.Item5 = -1 Select oMatrixStore.GetMatrixHeight(oImage.Item2)).First
                End If
                Using oBlobDetector As New Emgu.CV.Cvb.CvBlobDetector
                    Using oBorderMatrixLeft As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                        oBorderMatrixLeft.SetZero()
                        For i = 0 To iBorderWidth - 1
                            oBorderMatrixLeft.GetCol(i).SetValue(Byte.MaxValue)
                        Next

                        Using oBorderMatrixRight As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                            oBorderMatrixRight.SetZero()
                            For i = 0 To iBorderWidth - 1
                                oBorderMatrixRight.GetCol(iWidth - 1 - i).SetValue(Byte.MaxValue)
                            Next

                            Using oBorderMatrixTop As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                oBorderMatrixTop.SetZero()
                                For i = 0 To iBorderWidth - 1
                                    oBorderMatrixTop.GetRow(i).SetValue(Byte.MaxValue)
                                Next

                                Using oBorderMatrixBottom As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                    oBorderMatrixBottom.SetZero()
                                    For i = 0 To iBorderWidth - 1
                                        oBorderMatrixBottom.GetRow(iHeight - 1 - i).SetValue(Byte.MaxValue)
                                    Next

                                    Using oBorderMatrixMargin As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                        oBorderMatrixMargin.SetZero()
                                        Emgu.CV.CvInvoke.Max(oBorderMatrixLeft, oBorderMatrixMargin, oBorderMatrixMargin)
                                        Emgu.CV.CvInvoke.Max(oBorderMatrixRight, oBorderMatrixMargin, oBorderMatrixMargin)
                                        Emgu.CV.CvInvoke.Max(oBorderMatrixTop, oBorderMatrixMargin, oBorderMatrixMargin)
                                        Emgu.CV.CvInvoke.Max(oBorderMatrixBottom, oBorderMatrixMargin, oBorderMatrixMargin)
                                        oBorderMatrixMargin.SubR(Byte.MaxValue).CopyTo(oBorderMatrixMargin)

                                        Using oBorderMatrixRim As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                            oBorderMatrixRim.SetZero()
                                            oBorderMatrixRim.GetCol(0).SetValue(Byte.MaxValue)
                                            oBorderMatrixRim.GetCol(iWidth - 1).SetValue(Byte.MaxValue)
                                            oBorderMatrixRim.GetRow(0).SetValue(Byte.MaxValue)
                                            oBorderMatrixRim.GetRow(iHeight - 1).SetValue(Byte.MaxValue)

                                            Using oDetectMatrixOriginal As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                                Using oDetectMatrixComparison As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                                    Using oDetectMatrixOriginalMargin As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                                        Using oDetectMatrixComparisonMargin As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                                            For iIndex = 0 To oField.Images.Count - 1
                                                                If (Not IsNothing(oField.Images(iIndex).Item2)) AndAlso If(oField.FieldType = Enumerations.FieldTypeEnum.Handwriting, True, oField.Images(iIndex).Item5 = -1) Then
                                                                    ' set up list of adjacent rectangles for a handwriting field
                                                                    Dim oAdjacentRect As List(Of Integer) = ExtractAdjacentRect(oField, iIndex)
                                                                    If oAdjacentRect.Max > -1 Then
                                                                        Emgu.CV.CvInvoke.Threshold(oMatrixStore.GetMatrix(oField.Images(iIndex).Item2), oDetectMatrixOriginal, Byte.MaxValue * 0.9, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.BinaryInv)
                                                                        For iStep = 0 To oAdjacentRect.Count - 1
                                                                            If oAdjacentRect(iStep) <> -1 Then
                                                                                Emgu.CV.CvInvoke.Threshold(oMatrixStore.GetMatrix(oField.Images(oAdjacentRect(iStep)).Item2), oDetectMatrixComparison, Byte.MaxValue * 0.9, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.BinaryInv)
                                                                                Emgu.CV.CvInvoke.Min(oDetectMatrixOriginal, oBorderMatrixMargin, oDetectMatrixOriginalMargin)
                                                                                Emgu.CV.CvInvoke.Min(oDetectMatrixComparison, oBorderMatrixMargin, oDetectMatrixComparisonMargin)

                                                                                Using oBlobsOriginal As New Emgu.CV.Cvb.CvBlobs
                                                                                    Using oBlobsOriginalMargin As New Emgu.CV.Cvb.CvBlobs
                                                                                        ' get list of blobs with outer frame removed
                                                                                        oBlobDetector.Detect(oDetectMatrixOriginalMargin.Mat.ToImage(Of Emgu.CV.Structure.Gray, Byte), oBlobsOriginal)
                                                                                        Dim oBlobListOriginal As List(Of Emgu.CV.Cvb.CvBlob) = (From oBlob In oBlobsOriginal Select oBlob.Value).ToList

                                                                                        Select Case iStep
                                                                                            Case 0
                                                                                                ' left
                                                                                                Emgu.CV.CvInvoke.Max(oDetectMatrixOriginalMargin, oBorderMatrixLeft, oDetectMatrixOriginalMargin)
                                                                                            Case 1
                                                                                                ' right
                                                                                                Emgu.CV.CvInvoke.Max(oDetectMatrixOriginalMargin, oBorderMatrixRight, oDetectMatrixOriginalMargin)
                                                                                            Case 2
                                                                                                ' top
                                                                                                Emgu.CV.CvInvoke.Max(oDetectMatrixOriginalMargin, oBorderMatrixTop, oDetectMatrixOriginalMargin)
                                                                                            Case Else
                                                                                                ' bottom
                                                                                                Emgu.CV.CvInvoke.Max(oDetectMatrixOriginalMargin, oBorderMatrixBottom, oDetectMatrixOriginalMargin)
                                                                                        End Select
                                                                                        Emgu.CV.CvInvoke.Max(oDetectMatrixOriginalMargin, oBorderMatrixRim, oDetectMatrixOriginalMargin)

                                                                                        ' get list of blobs that are not joined to the outer frame
                                                                                        oBlobDetector.Detect(oDetectMatrixOriginalMargin.Mat.ToImage(Of Emgu.CV.Structure.Gray, Byte), oBlobsOriginalMargin)
                                                                                        Dim oBlobListOriginalMargin As List(Of Emgu.CV.Cvb.CvBlob) = (From oBlob In oBlobsOriginalMargin Where oBlob.Value.BoundingBox.Width <> iWidth Or oBlob.Value.BoundingBox.Height <> iHeight Select oBlob.Value).ToList

                                                                                        ' subtract the second list of blobs from the first to leave just the blobs connected to the outer frame
                                                                                        For i = oBlobListOriginal.Count - 1 To 0 Step -1
                                                                                            For j = 0 To oBlobListOriginalMargin.Count - 1
                                                                                                If oBlobListOriginal(i).BoundingBox.Width = oBlobListOriginalMargin(j).BoundingBox.Width AndAlso oBlobListOriginal(i).BoundingBox.Height = oBlobListOriginalMargin(j).BoundingBox.Height AndAlso oBlobListOriginal(i).Centroid.X = oBlobListOriginalMargin(j).Centroid.X AndAlso oBlobListOriginal(i).Centroid.Y = oBlobListOriginalMargin(j).Centroid.Y Then
                                                                                                    oBlobListOriginal.RemoveAt(i)
                                                                                                    Exit For
                                                                                                End If
                                                                                            Next
                                                                                        Next

                                                                                        ' check comparison
                                                                                        If oBlobListOriginal.Count > 0 Then
                                                                                            Using oBlobsComparison As New Emgu.CV.Cvb.CvBlobs
                                                                                                Using oBlobsComparisonMargin As New Emgu.CV.Cvb.CvBlobs
                                                                                                    ' get list of blobs with outer frame removed
                                                                                                    oBlobDetector.Detect(oDetectMatrixComparisonMargin.Mat.ToImage(Of Emgu.CV.Structure.Gray, Byte), oBlobsComparison)
                                                                                                    Dim oBlobListComparison As List(Of Emgu.CV.Cvb.CvBlob) = (From oBlob In oBlobsComparison Select oBlob.Value).ToList

                                                                                                    Select Case iStep
                                                                                                        Case 0
                                                                                                            ' left
                                                                                                            Emgu.CV.CvInvoke.Max(oDetectMatrixComparisonMargin, oBorderMatrixRight, oDetectMatrixComparisonMargin)
                                                                                                        Case 1
                                                                                                            ' right
                                                                                                            Emgu.CV.CvInvoke.Max(oDetectMatrixComparisonMargin, oBorderMatrixLeft, oDetectMatrixComparisonMargin)
                                                                                                        Case 2
                                                                                                            ' top
                                                                                                            Emgu.CV.CvInvoke.Max(oDetectMatrixComparisonMargin, oBorderMatrixBottom, oDetectMatrixComparisonMargin)
                                                                                                        Case Else
                                                                                                            ' bottom
                                                                                                            Emgu.CV.CvInvoke.Max(oDetectMatrixComparisonMargin, oBorderMatrixTop, oDetectMatrixComparisonMargin)
                                                                                                    End Select
                                                                                                    Emgu.CV.CvInvoke.Max(oDetectMatrixComparisonMargin, oBorderMatrixRim, oDetectMatrixComparisonMargin)

                                                                                                    ' get list of blobs that are not joined to the outer frame
                                                                                                    oBlobDetector.Detect(oDetectMatrixComparisonMargin.Mat.ToImage(Of Emgu.CV.Structure.Gray, Byte), oBlobsComparisonMargin)
                                                                                                    Dim oBlobListComparisonMargin As List(Of Emgu.CV.Cvb.CvBlob) = (From oBlob In oBlobsComparisonMargin Where oBlob.Value.BoundingBox.Width <> iWidth Or oBlob.Value.BoundingBox.Height <> iHeight Select oBlob.Value).ToList

                                                                                                    ' subtract the second list of blobs from the first to leave just the blobs connected to the outer frame
                                                                                                    For i = oBlobListComparison.Count - 1 To 0 Step -1
                                                                                                        For j = 0 To oBlobListComparisonMargin.Count - 1
                                                                                                            If oBlobListComparison(i).BoundingBox.Width = oBlobListComparisonMargin(j).BoundingBox.Width AndAlso oBlobListComparison(i).BoundingBox.Height = oBlobListComparisonMargin(j).BoundingBox.Height AndAlso oBlobListComparison(i).Centroid.X = oBlobListComparisonMargin(j).Centroid.X AndAlso oBlobListComparison(i).Centroid.Y = oBlobListComparisonMargin(j).Centroid.Y Then
                                                                                                                oBlobListComparison.RemoveAt(i)
                                                                                                                Exit For
                                                                                                            End If
                                                                                                        Next
                                                                                                    Next

                                                                                                    ' adjacent blocks have extensions touching the common border
                                                                                                    If oBlobListComparison.Count > 0 Then
                                                                                                        Dim iCroppedWidth As Integer = iWidth - iBorderWidth * 2
                                                                                                        Dim iCroppedHeight As Integer = iHeight - iBorderWidth * 2
                                                                                                        Dim oCropRect As New System.Drawing.Rectangle(iBorderWidth, iBorderWidth, iCroppedWidth, iCroppedHeight)
                                                                                                        Dim oScaleSize As New System.Drawing.Size((iScaleFactor * 2) + 1, (iScaleFactor * 2) + 1)

                                                                                                        ' create an upscaled matrix for stroke analysis
                                                                                                        Using oScaledMatrixOriginal As New Emgu.CV.Matrix(Of Byte)(iCroppedHeight * iScaleFactor, iCroppedWidth * iScaleFactor)
                                                                                                            Using oScaledMatrixComparison As New Emgu.CV.Matrix(Of Byte)(iCroppedHeight * iScaleFactor, iCroppedWidth * iScaleFactor)
                                                                                                                Emgu.CV.CvInvoke.Resize(oDetectMatrixOriginalMargin.GetSubRect(oCropRect), oScaledMatrixOriginal, oScaledMatrixOriginal.Size, 0, 0, Emgu.CV.CvEnum.Inter.Cubic)
                                                                                                                Emgu.CV.CvInvoke.Resize(oDetectMatrixComparisonMargin.GetSubRect(oCropRect), oScaledMatrixComparison, oScaledMatrixComparison.Size, 0, 0, Emgu.CV.CvEnum.Inter.Cubic)

                                                                                                                Dim iReverseStep As Integer = (Math.Floor(iStep / 2) * 2) + (1 - (iStep Mod 2))
                                                                                                                Dim oPointListOriginal As PointList = GetPoints(oScaledMatrixOriginal)
                                                                                                                Dim oNearestPointListOriginal As List(Of Tuple(Of Integer, PointList)) = GetNearestPoint(oBlobListOriginal, oPointListOriginal, iCroppedWidth * iScaleFactor, iCroppedHeight * iScaleFactor, iStep, iBorderWidth, iScaleFactor)
                                                                                                                Dim oPointListComparison As PointList = GetPoints(oScaledMatrixComparison)
                                                                                                                Dim oNearestPointListComparison As List(Of Tuple(Of Integer, PointList)) = GetNearestPoint(oBlobListComparison, oPointListComparison, iCroppedWidth * iScaleFactor, iCroppedHeight * iScaleFactor, iReverseStep, iBorderWidth, iScaleFactor)

                                                                                                                Dim oRemoveOriginalList As New List(Of Tuple(Of Integer, PointList))
                                                                                                                Dim oRemoveComparisonList As New List(Of Tuple(Of Integer, PointList))
                                                                                                                For i = 0 To oNearestPointListOriginal.Count - 1
                                                                                                                    ' the remove list is a tuple with the first value as the index on the blob list, the second value as the stroke to remove, the third value as the angle deviance (smaller is better), the fourth value showing whether to remove the original stroke (false) or the comparison stroke (true)
                                                                                                                    Dim oRemoveList As New List(Of Tuple(Of Integer, PointList, Double, Boolean))
                                                                                                                    For j = 0 To oNearestPointListComparison.Count - 1
                                                                                                                        Dim oConnectedStrokes As Tuple(Of Boolean, Double, Boolean) = ConnectedStrokes(oNearestPointListOriginal(i).Item2, oNearestPointListComparison(j).Item2, iWidth, iHeight, iStep, iBorderWidth, iScaleFactor)
                                                                                                                        If oConnectedStrokes.Item1 Then
                                                                                                                            oRemoveList.Add(New Tuple(Of Integer, PointList, Double, Boolean)(If(oConnectedStrokes.Item3, oNearestPointListComparison(j).Item1, oNearestPointListOriginal(i).Item1), If(oConnectedStrokes.Item3, oNearestPointListComparison(j).Item2, oNearestPointListOriginal(i).Item2), oConnectedStrokes.Item2, oConnectedStrokes.Item3))
                                                                                                                        End If
                                                                                                                    Next

                                                                                                                    ' order by ascending angle deviance 
                                                                                                                    If oRemoveList.Count > 0 Then
                                                                                                                        Dim oRemoveActual As Tuple(Of Integer, PointList, Double, Boolean) = (From oRemove In oRemoveList Order By oRemove.Item3 Select oRemove).First
                                                                                                                        If oRemoveActual.Item4 Then
                                                                                                                            oRemoveComparisonList.Add(New Tuple(Of Integer, PointList)(oRemoveActual.Item1, oRemoveActual.Item2))
                                                                                                                        Else
                                                                                                                            oRemoveOriginalList.Add(New Tuple(Of Integer, PointList)(oRemoveActual.Item1, oRemoveActual.Item2))
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                Next

                                                                                                                Using oMask As New Emgu.CV.Matrix(Of Byte)(iHeight + 2, iWidth + 2)
                                                                                                                    If oRemoveOriginalList.Count > 0 Then
                                                                                                                        Using oModifiedMatrixOriginal As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                                                                                                            oDetectMatrixOriginal.CopyTo(oModifiedMatrixOriginal)

                                                                                                                            For Each oRemoveStroke In oRemoveOriginalList
                                                                                                                                Dim oScaledBoundingBox As System.Drawing.Rectangle = oRemoveStroke.Item2.BoundingBox
                                                                                                                                Dim oBoundingBox As New System.Drawing.Rectangle(iBorderWidth + (oScaledBoundingBox.X / iScaleFactor), iBorderWidth + (oScaledBoundingBox.Y / iScaleFactor), oScaledBoundingBox.Width / iScaleFactor, oScaledBoundingBox.Height / iScaleFactor)
                                                                                                                                Dim oRect As New System.Drawing.Rectangle(oBoundingBox.X + 1, oBoundingBox.Y + 1, oBoundingBox.Width, oBoundingBox.Height)
                                                                                                                                Dim oNearestPointScaled As PointList.SinglePoint = oRemoveStroke.Item2.First
                                                                                                                                Dim oNearestPoint As New System.Drawing.Point(iBorderWidth + (oNearestPointScaled.Point.X / iScaleFactor), iBorderWidth + (oNearestPointScaled.Point.Y / iScaleFactor))

                                                                                                                                oMask.SetValue(Byte.MaxValue)
                                                                                                                                Emgu.CV.CvInvoke.Rectangle(oMask, oRect, New Emgu.CV.Structure.MCvScalar(0), -1)
                                                                                                                                Using oDuplicateMask As Emgu.CV.Matrix(Of Byte) = oMask.Clone
                                                                                                                                    Emgu.CV.CvInvoke.FloodFill(oModifiedMatrixOriginal, oMask, oNearestPoint, New Emgu.CV.Structure.MCvScalar(0), Nothing, New Emgu.CV.Structure.MCvScalar(0), New Emgu.CV.Structure.MCvScalar(0), Emgu.CV.CvEnum.Connectivity.EightConnected, Emgu.CV.CvEnum.FloodFillType.FixedRange Or (255 << 8) Or Emgu.CV.CvEnum.FloodFillType.MaskOnly)
                                                                                                                                    oMask.SetValue(0, oDuplicateMask)
                                                                                                                                End Using
                                                                                                                                Emgu.CV.CvInvoke.Dilate(oMask, oMask, Nothing, Nothing, 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(0))

                                                                                                                                Using oNewMatrix As Emgu.CV.Matrix(Of Byte) = oMatrixStore.GetMatrix(oField.Images(iIndex).Item2)
                                                                                                                                    oNewMatrix.SetValue(New Emgu.CV.Structure.MCvScalar(Byte.MaxValue), oMask.GetSubRect(New System.Drawing.Rectangle(1, 1, iWidth, iHeight)))
                                                                                                                                    oMatrixStore.ReplaceMatrix(oField.Images(iIndex).Item2, oNewMatrix)
                                                                                                                                End Using
                                                                                                                            Next
                                                                                                                        End Using
                                                                                                                    End If
                                                                                                                    If oRemoveComparisonList.Count > 0 Then
                                                                                                                        Using oModifiedMatrixComparison As New Emgu.CV.Matrix(Of Byte)(iHeight, iWidth)
                                                                                                                            oDetectMatrixComparison.CopyTo(oModifiedMatrixComparison)

                                                                                                                            For Each oRemoveStroke In oRemoveComparisonList
                                                                                                                                Dim oScaledBoundingBox As System.Drawing.Rectangle = oRemoveStroke.Item2.BoundingBox
                                                                                                                                Dim oBoundingBox As New System.Drawing.Rectangle(iBorderWidth + (oScaledBoundingBox.X / iScaleFactor), iBorderWidth + (oScaledBoundingBox.Y / iScaleFactor), oScaledBoundingBox.Width / iScaleFactor, oScaledBoundingBox.Height / iScaleFactor)
                                                                                                                                Dim oRect As New System.Drawing.Rectangle(oBoundingBox.X + 1, oBoundingBox.Y + 1, oBoundingBox.Width, oBoundingBox.Height)
                                                                                                                                Dim oNearestPointScaled As PointList.SinglePoint = oRemoveStroke.Item2.First
                                                                                                                                Dim oNearestPoint As New System.Drawing.Point(iBorderWidth + (oNearestPointScaled.Point.X / iScaleFactor), iBorderWidth + (oNearestPointScaled.Point.Y / iScaleFactor))

                                                                                                                                oMask.SetValue(Byte.MaxValue)
                                                                                                                                Emgu.CV.CvInvoke.Rectangle(oMask, oRect, New Emgu.CV.Structure.MCvScalar(0), -1)
                                                                                                                                Using oDuplicateMask As Emgu.CV.Matrix(Of Byte) = oMask.Clone
                                                                                                                                    Emgu.CV.CvInvoke.FloodFill(oModifiedMatrixComparison, oMask, oNearestPoint, New Emgu.CV.Structure.MCvScalar(0), Nothing, New Emgu.CV.Structure.MCvScalar(0), New Emgu.CV.Structure.MCvScalar(0), Emgu.CV.CvEnum.Connectivity.EightConnected, Emgu.CV.CvEnum.FloodFillType.FixedRange Or (255 << 8) Or Emgu.CV.CvEnum.FloodFillType.MaskOnly)
                                                                                                                                    oMask.SetValue(0, oDuplicateMask)
                                                                                                                                End Using
                                                                                                                                Emgu.CV.CvInvoke.Dilate(oMask, oMask, Nothing, Nothing, 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(0))

                                                                                                                                Using oNewMatrix As Emgu.CV.Matrix(Of Byte) = oMatrixStore.GetMatrix(oField.Images(oAdjacentRect(iStep)).Item2)
                                                                                                                                    oNewMatrix.SetValue(New Emgu.CV.Structure.MCvScalar(Byte.MaxValue), oMask.GetSubRect(New System.Drawing.Rectangle(1, 1, iWidth, iHeight)))
                                                                                                                                    oMatrixStore.ReplaceMatrix(oField.Images(oAdjacentRect(iStep)).Item2, oNewMatrix)
                                                                                                                                End Using
                                                                                                                            Next
                                                                                                                        End Using
                                                                                                                    End If
                                                                                                                End Using
                                                                                                            End Using
                                                                                                        End Using
                                                                                                    End If
                                                                                                End Using
                                                                                            End Using
                                                                                        End If
                                                                                    End Using
                                                                                End Using
                                                                            End If
                                                                        Next
                                                                    End If
                                                                End If
                                                            Next
                                                        End Using
                                                    End Using
                                                End Using
                                            End Using
                                        End Using
                                    End Using
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End If
        End If
    End Sub
    Public Shared Function ExtractAdjacentRect(ByVal oField As FieldDocumentStore.Field, ByVal iIndex As Integer) As List(Of Integer)
        ' extracts the adjacent rectangles for handwriting fields
        ' adjacent rectangles are in order left, right, top, bottom
        Dim oAdjacentRect As New List(Of Integer)
        Dim oFoundRect As List(Of Integer) = Nothing

        Dim iRow As Integer = oField.Images(iIndex).Item4
        Dim iColumn As Integer = oField.Images(iIndex).Item5
        If oField.FieldType = Enumerations.FieldTypeEnum.Handwriting Then
            If iRow >= 0 And iColumn >= 0 Then
                ' left
                oAdjacentRect.Add(GetRect(oField, iRow, iColumn - 1))

                ' right
                oAdjacentRect.Add(GetRect(oField, iRow, iColumn + 1))

                ' top
                oAdjacentRect.Add(GetRect(oField, iRow - 1, iColumn))

                ' right
                oAdjacentRect.Add(GetRect(oField, iRow + 1, iColumn))
            End If
        Else
            ' left
            oAdjacentRect.Add(-1)

            ' right
            oAdjacentRect.Add(-1)

            ' top
            oAdjacentRect.Add(GetRect(oField, iRow - 1, iColumn))

            ' right
            oAdjacentRect.Add(GetRect(oField, iRow + 1, iColumn))
        End If

        Return oAdjacentRect
    End Function
    Private Shared Function GetRect(ByVal oField As FieldDocumentStore.Field, ByVal iRow As Integer, ByVal iColumn As Integer) As Integer
        ' extracts the adjacent rectangles for handwriting fields
        If If(oField.FieldType = Enumerations.FieldTypeEnum.Handwriting, iRow < 0 Or iColumn < 0, False) Then
            Return -1
        Else
            Dim oFoundRect As List(Of Integer) = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iRow And oField.Images(iIndex).Item5 = iColumn And (Not oField.Images(iIndex).Item6) Select iIndex).ToList
            If oFoundRect.Count = 0 Then
                Return -1
            Else
                Return oFoundRect.First
            End If
        End If
    End Function
    Private Shared Function GetPoints(ByVal oScaledMatrix As Emgu.CV.Matrix(Of Byte)) As PointList
        ' get the list of single points
        Const fMinFraction As Single = 0.25
        Dim oPointCircleList As New PointList
        Using oImageMatrix As New Emgu.CV.Matrix(Of Byte)(oScaledMatrix.Size)
            Emgu.CV.CvInvoke.Threshold(oScaledMatrix, oImageMatrix, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)

            ' get contours
            Dim oContours As Tuple(Of Emgu.CV.Util.VectorOfVectorOfPoint, Dictionary(Of Integer, ContourType)) = GetContours(oImageMatrix)
            If Not IsNothing(oContours) Then
                ' simplify contour
                Using oContourMat As New Emgu.CV.Image(Of Emgu.CV.Structure.Gray, Single)(oImageMatrix.Cols, oImageMatrix.Rows, New Emgu.CV.Structure.Gray(0))
                    ' create a distance map for the contour
                    Using oContourFilledMat As New Emgu.CV.Image(Of Emgu.CV.Structure.Gray, Byte)(oImageMatrix.Cols, oImageMatrix.Rows, New Emgu.CV.Structure.Gray(0))
                        ' close contours
                        For iIndex = 0 To oContours.Item1.Size - 1
                            Emgu.CV.CvInvoke.ApproxPolyDP(oContours.Item1(iIndex), oContours.Item1(iIndex), 0, True)
                        Next

                        For iIndex = 0 To oContours.Item2.Count - 1
                            If oContours.Item2(oContours.Item2.Keys(iIndex)) = ContourType.External Then
                                Emgu.CV.CvInvoke.DrawContours(oContourFilledMat, oContours.Item1, oContours.Item2.Keys(iIndex), New Emgu.CV.Structure.MCvScalar(Byte.MaxValue), -1)
                            End If
                        Next
                        For iIndex = 0 To oContours.Item2.Count - 1
                            If oContours.Item2(oContours.Item2.Keys(iIndex)) = ContourType.Hole Then
                                Emgu.CV.CvInvoke.DrawContours(oContourFilledMat, oContours.Item1, oContours.Item2.Keys(iIndex), New Emgu.CV.Structure.MCvScalar(Byte.MinValue), -1)
                            End If
                        Next

                        Emgu.CV.CvInvoke.DistanceTransform(oImageMatrix, oContourMat, Nothing, Emgu.CV.CvEnum.DistType.L2, 0)
                    End Using

                    ' create point circle list from skeleton
                    Using oSkeletonMatrix As Emgu.CV.Matrix(Of Byte) = Skeletonisation(oImageMatrix)
                        Dim oPixelBag As New ConcurrentBag(Of PointList.SinglePoint)
                        Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                                    Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                                    Dim oBufferIn As Emgu.CV.Matrix(Of Byte) = CType(oParam.Item4, Emgu.CV.Matrix(Of Byte))
                                                                    Dim oBufferOut As ConcurrentBag(Of PointList.SinglePoint) = CType(oParam.Item5, ConcurrentBag(Of PointList.SinglePoint))
                                                                    Dim localwidth As Integer = CType(oParam.Item6, Integer)
                                                                    Dim ContourMatLocal As Emgu.CV.Image(Of Emgu.CV.Structure.Gray, Single) = CType(oParam.Item7, Emgu.CV.Image(Of Emgu.CV.Structure.Gray, Single))

                                                                    For y = 0 To oParam.Item2 - 1
                                                                        Dim localY As Integer = y + oParam.Item1
                                                                        For x = 0 To localwidth - 1
                                                                            If oBufferIn(localY, x) <> Byte.MinValue Then
                                                                                oBufferOut.Add(New PointList.SinglePoint(ContourMatLocal(localY, x).Intensity, New System.Drawing.Point(x, localY), 0, 0))
                                                                            End If
                                                                        Next
                                                                    Next
                                                                End Sub
                        CommonFunctions.ParallelRun(oSkeletonMatrix.Height, oSkeletonMatrix, oPixelBag, oSkeletonMatrix.Width, oContourMat, TaskDelegate)

                        Dim oPixelList As List(Of PointList.SinglePoint) = oPixelBag.OrderByDescending(Function(x) x.Radius).ToList
                        If oPixelList.Count > 0 Then
                            Dim fSetMax As Double = oPixelList.First.Radius * fMinFraction
                            Do Until oPixelList.Count = 0 OrElse oPixelList.First.Radius < fSetMax
                                oPointCircleList.Add(oPixelList.First)
                                For iIndex = oPixelList.Count - 1 To 1 Step -1
                                    If PointList.InCircle(oPixelList.First, oPixelList(iIndex)) Then
                                        oPixelList.RemoveAt(iIndex)
                                    End If
                                Next
                                oPixelList.RemoveAt(0)
                            Loop
                        End If
                    End Using
                End Using
            End If
        End Using
        Return oPointCircleList
    End Function
    Private Shared Function GetNearestPoint(ByVal oBlobList As List(Of Emgu.CV.Cvb.CvBlob), ByVal oPointList As PointList, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal iStep As Integer, ByVal iBorderWidth As Integer, ByVal iScaleFactor As Integer) As List(Of Tuple(Of Integer, PointList))
        ' checks which points are nearest to the edge for each stated blob
        ' the return value gives a tuple with the first item being the index in the blob list, and the second being the index in the point list
        Dim oReturnList As New List(Of Tuple(Of Integer, PointList))
        If oBlobList.Count > 0 And oPointList.Count > 0 Then
            For i = 0 To oBlobList.Count - 1
                Dim iCurrent As Integer = i
                Dim oSelectedPointList As New PointList((From oPoint As PointList.SinglePoint In oPointList Where oBlobList(iCurrent).BoundingBox.Contains(iBorderWidth + oPoint.Point.X / iScaleFactor, iBorderWidth + oPoint.Point.Y / iScaleFactor) Select oPoint).ToList)

                ' get a list of points that belong to the blob
                Dim oBorderPointList As List(Of PointList.SinglePoint) = Nothing
                If oSelectedPointList.Count > 0 Then
                    Select Case iStep
                        Case 0
                            ' left
                            oBorderPointList = (From iIndex As Integer In Enumerable.Range(0, oSelectedPointList.Count) Where oSelectedPointList(iIndex).Point.X - oSelectedPointList(iIndex).Radius <= iScaleFactor / 2 Order By oSelectedPointList(iIndex).Radius Descending Select oSelectedPointList(iIndex)).ToList
                        Case 1
                            ' right
                            oBorderPointList = (From iIndex As Integer In Enumerable.Range(0, oSelectedPointList.Count) Where oSelectedPointList(iIndex).Point.X + oSelectedPointList(iIndex).Radius >= iWidth - 1 - iScaleFactor / 2 Order By oSelectedPointList(iIndex).Radius Descending Select oSelectedPointList(iIndex)).ToList
                        Case 2
                            ' top
                            oBorderPointList = (From iIndex As Integer In Enumerable.Range(0, oSelectedPointList.Count) Where oSelectedPointList(iIndex).Point.Y - oSelectedPointList(iIndex).Radius <= iScaleFactor / 2 Order By oSelectedPointList(iIndex).Radius Descending Select oSelectedPointList(iIndex)).ToList
                        Case Else
                            ' bottom
                            oBorderPointList = (From iIndex As Integer In Enumerable.Range(0, oSelectedPointList.Count) Where oSelectedPointList(iIndex).Point.Y + oSelectedPointList(iIndex).Radius >= iHeight - 1 - iScaleFactor / 2 Order By oSelectedPointList(iIndex).Radius Descending Select oSelectedPointList(iIndex)).ToList
                    End Select

                    ' remove smaller points
                    If oBorderPointList.Count > 1 Then
                        For j = 1 To oBorderPointList.Count - 1
                            oSelectedPointList.Remove(oBorderPointList(j))
                        Next
                    End If

                    If oBorderPointList.Count > 0 Then
                        Dim iPointIndex As Integer = oSelectedPointList.IndexOf(oBorderPointList.First)

                        ' build stroke
                        Dim oStroke As New PointList({oSelectedPointList(iPointIndex)}.ToList)
                        oSelectedPointList.Remove(oStroke.First)
                        Do Until oSelectedPointList.Count = 0
                            Using oNeighbouringPointsList As PointList = PointList.GetNeighbouringPoints(oStroke.Last, oSelectedPointList, True, 0)
                                If oNeighbouringPointsList.Count = 1 Then
                                    oStroke.Add(oNeighbouringPointsList.First)
                                    oSelectedPointList.Remove(oNeighbouringPointsList.First)
                                ElseIf oNeighbouringPointsList.Count > 1 AndAlso GetAngleSpread(oStroke.Last, oNeighbouringPointsList) < Math.PI / 4 Then
                                    ' all points are within 90 degrees
                                    Dim oNearestPoint As PointList.SinglePoint = (From oPoint In oNeighbouringPointsList Order By oPoint.AForgePoint.DistanceTo(oStroke.Last.AForgePoint) Ascending).First
                                    oStroke.Add(oNearestPoint)
                                    oSelectedPointList.Remove(oNearestPoint)
                                Else
                                    ' remove the last point
                                    If oStroke.Count > 1 Then
                                        oStroke.RemoveAt(oStroke.Count - 1)
                                    End If
                                    Exit Do
                                End If
                            End Using
                        Loop

                        oReturnList.Add(New Tuple(Of Integer, PointList)(i, oStroke))
                    End If
                End If
            Next
        End If
        Return oReturnList
    End Function
    Private Shared Function GetAngleSpread(ByVal oStartPoint As PointList.SinglePoint, ByVal oNeighbouringPointsList As PointList) As Double
        ' gets the angle spread of all neighbouring points from the start point
        Dim fMinAngle As Double = Double.MaxValue
        Dim fMaxAngle As Double = Double.MinValue
        For Each oPoint In oNeighbouringPointsList
            Dim fAngle As Double = PointList.GetAngle(oStartPoint, oPoint)
            fMinAngle = Math.Min(fMinAngle, fAngle)
            fMaxAngle = Math.Max(fMaxAngle, fAngle)
        Next
        Return Math.Abs(fMaxAngle - fMinAngle)
    End Function
    Private Shared Function ConnectedStrokes(ByVal oOriginalStroke As PointList, ByVal oComparisonStroke As PointList, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal iStep As Integer, ByVal iBorderWidth As Integer, ByVal iScaleFactor As Integer) As Tuple(Of Boolean, Double, Boolean)
        ' checks two strokes from adjacent handwriting boxes to see if they are connected
        ' this is using an angle as well as a location check
        ' the nearest points are the index on the stroke which is nearest to the common border
        ' the return value is a tuple with the first value showing whether the two points are in range, and the second value shows the angle deviance between the strokes (smaller is better), and the third value showing whether to remove the original stroke (false) or the comparison stroke (true)
        ' check to see if the nearest points are either at the beginning or end of the respective strokes, and if not then the strokes are not connected
        ' determine the displacement to add to the second stroke to convert it to the same coordinates as the first stroke
        Dim iDisplacementX As Integer = 0
        Dim iDisplacementY As Integer = 0
        Select Case iStep
            Case 0
                ' left
                iDisplacementX = ((iBorderWidth - iWidth) * iScaleFactor)
            Case 1
                ' right
                iDisplacementX = ((iWidth - iBorderWidth) * iScaleFactor)
            Case 2
                ' top
                iDisplacementY = ((iBorderWidth - iHeight) * iScaleFactor)
            Case Else
                ' bottom
                iDisplacementY = ((iHeight - iBorderWidth) * iScaleFactor)
        End Select

        Dim oComparisonStrokeDisplaced As New PointList((From oPoint In oComparisonStroke Select New PointList.SinglePoint(oPoint.Radius, New System.Drawing.Point(oPoint.Point.X + iDisplacementX, oPoint.Point.Y + iDisplacementY), oPoint.Count, oPoint.Distance, oPoint.GUID)).ToList)
        Dim oPointComparisonDisplaced As PointList.SinglePoint = oComparisonStrokeDisplaced.First

        ' get a copy of the original stroke with the nearest point last
        If oOriginalStroke.Count >= 3 Or oComparisonStrokeDisplaced.Count >= 3 Then
            Dim oOriginalStrokeCopy As New PointList
            oOriginalStrokeCopy.AddRange(oOriginalStroke)
            oOriginalStrokeCopy.Reverse()

            Dim oNeighbouringPointsListOriginal As New PointList({oPointComparisonDisplaced}.ToList)
            Dim oNeighbouringPointsListComparison As New PointList({oOriginalStrokeCopy(oOriginalStrokeCopy.Count - 1)}.ToList)

            Dim oCompareAngleListOriginal As List(Of Double) = Nothing
            Dim oCompareAngleListComparison As List(Of Double) = Nothing
            If oOriginalStroke.Count >= 3 Then
                oCompareAngleListOriginal = GetFollowingPoints(oOriginalStrokeCopy, oNeighbouringPointsListOriginal, Math.PI / 4)
            End If

            ' if the comparison stroke has enough points to get a reliable running angle, then get the reverse comparison
            If oComparisonStrokeDisplaced.Count >= 3 Then
                ' get a copy of the comparison stroke with the nearest point last
                oComparisonStrokeDisplaced.Reverse()
                oCompareAngleListComparison = GetFollowingPoints(oComparisonStrokeDisplaced, oNeighbouringPointsListComparison, Math.PI / 4)
            End If

            Dim fCompareAngle As Double = 0
            If oNeighbouringPointsListOriginal.Count > 0 AndAlso (Not IsNothing(oCompareAngleListOriginal)) AndAlso oNeighbouringPointsListComparison.Count > 0 AndAlso (Not IsNothing(oCompareAngleListComparison)) Then
                ' both present
                fCompareAngle = (Math.Abs(oCompareAngleListOriginal.First) + Math.Abs(oCompareAngleListComparison.First)) / 2
            ElseIf oNeighbouringPointsListOriginal.Count > 0 AndAlso (Not IsNothing(oCompareAngleListOriginal)) Then
                ' only original present
                fCompareAngle = Math.Abs(oCompareAngleListOriginal.First)
            ElseIf oNeighbouringPointsListComparison.Count > 0 AndAlso (Not IsNothing(oCompareAngleListComparison)) Then
                ' only comparison present
                fCompareAngle = Math.Abs(oCompareAngleListComparison.First)
            Else
                ' not connected
                Return New Tuple(Of Boolean, Double, Boolean)(False, 0, False)
            End If

            ' the smaller stroke is the one to delete
            Dim fOriginalArea As Single = Aggregate oPoint In oOriginalStroke Into Sum(oPoint.Radius * oPoint.Radius)
            Dim fComparisonArea As Single = Aggregate oPoint In oComparisonStroke Into Sum(oPoint.Radius * oPoint.Radius)

            Return New Tuple(Of Boolean, Double, Boolean)(True, fCompareAngle, fOriginalArea >= fComparisonArea)
        Else
            ' strokes not long enough to determine connection
            Return New Tuple(Of Boolean, Double, Boolean)(False, 0, False)
        End If
    End Function
    Private Shared Function GetFollowingPoints(ByVal oCurrentStroke As PointList, ByRef oNeighbouringPointsList As PointList, ByVal fAngleRange As Double) As List(Of Double)
        ' filters out points on the same side of the current stroke
        ' returns a list of compare angles corresponding to the residual neighbouring points list
        Dim oCompareAngleList As New List(Of Tuple(Of Guid, Double))
        If oNeighbouringPointsList.Count > 0 Then
            For i = oNeighbouringPointsList.Count - 1 To 0 Step -1
                Dim fNewAngle As Double = PointList.GetAngle(oCurrentStroke(oCurrentStroke.Count - 1), oNeighbouringPointsList(i))
                Dim fCompareAngle As Double = PointList.CompareAngleValue(GetRunningAngle(oCurrentStroke, 6), fNewAngle)
                If fCompareAngle <= fAngleRange Then
                    oCompareAngleList.Add(New Tuple(Of Guid, Double)(oNeighbouringPointsList(i).GUID, fCompareAngle))
                Else
                    oNeighbouringPointsList.RemoveAt(i)
                End If
            Next
        End If

        Dim oQueryPointList As New PointList(oNeighbouringPointsList)
        Dim oReturnCompareAngleList As List(Of Double) = (From oCompareAngle In oCompareAngleList From iIndex As Integer In Enumerable.Range(0, oQueryPointList.Count) Where oCompareAngle.Item1.Equals(oQueryPointList(iIndex).GUID) Select New Tuple(Of Integer, Double)(iIndex, oCompareAngle.Item2)).OrderBy(Function(x) x.Item1).Select(Function(x) x.Item2).ToList
        Return oReturnCompareAngleList
    End Function
    Private Shared Function GetRunningAngle(ByVal oCurrentStroke As PointList, ByVal iPointCount As Integer) As Double
        ' gets running average of RunningAngle
        ' take only the last iPointCount strokes, 6 of which should account for >90% of the RunningAngle
        Dim oRunningAngleList As New List(Of Double)
        For i = Math.Max(1, oCurrentStroke.Count - iPointCount) To oCurrentStroke.Count - 1
            oRunningAngleList.Add(PointList.GetAngle(oCurrentStroke(i - 1), oCurrentStroke(i)))
        Next

        ' get running average of RunningAngle
        ' the running deviance is used to predict the final running angle
        Dim fRunningDeviance As Double = 0
        Dim fRunningAngle As Double = oRunningAngleList.First
        For i = 1 To oRunningAngleList.Count - 1
            Dim fOldRunningAngle As Double = fRunningAngle
            fRunningAngle = (fRunningAngle * 2 / 3) + (PointList.MatchAngle(oRunningAngleList(i), fRunningAngle) / 3)
            fRunningDeviance = (fRunningDeviance * 2 / 3) + ((fRunningAngle - fOldRunningAngle) / 3)
        Next
        Return fRunningAngle + fRunningDeviance
    End Function
    Private Shared Function GetContours(ByVal oMatrix As Emgu.CV.Matrix(Of Byte)) As Tuple(Of Emgu.CV.Util.VectorOfVectorOfPoint, Dictionary(Of Integer, ContourType))
        ' gets a list of contours from an image
        Using oContourMat As Emgu.CV.Matrix(Of Byte) = oMatrix.Clone
            Dim oContours As New Emgu.CV.Util.VectorOfVectorOfPoint
            Dim oHierarchy As Integer(,) = Emgu.CV.CvInvoke.FindContourTree(oContourMat, oContours, Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxTc89L1)

            Dim oContourTypeList As New Dictionary(Of Integer, ContourType)
            SetContourType(oContours(0), oContours, oHierarchy, oContourTypeList, ContourType.External, 0)

            Return New Tuple(Of Emgu.CV.Util.VectorOfVectorOfPoint, Dictionary(Of Integer, ContourType))(oContours, oContourTypeList)
        End Using
    End Function
    Private Shared Sub SetContourType(ByRef oContour As Emgu.CV.Util.VectorOfPoint, ByRef oContours As Emgu.CV.Util.VectorOfVectorOfPoint, ByRef oHierarchy As Integer(,), ByRef oContourTypeDictionary As Dictionary(Of Integer, ContourType), ByVal oContourType As ContourType, ByVal iContourIndex As Integer)
        ' recursively steps through contours and sets their contour type
        ' sets the contourtype
        oContourTypeDictionary.Add(iContourIndex, oContourType)

        ' go forwards
        If oHierarchy(iContourIndex, 0) <> -1 AndAlso (Not oContourTypeDictionary.ContainsKey(oHierarchy(iContourIndex, 0))) Then
            SetContourType(oContours(oHierarchy(iContourIndex, 0)), oContours, oHierarchy, oContourTypeDictionary, oContourType, oHierarchy(iContourIndex, 0))
        End If

        ' go backwards
        If oHierarchy(iContourIndex, 1) <> -1 AndAlso (Not oContourTypeDictionary.ContainsKey(oHierarchy(iContourIndex, 1))) Then
            SetContourType(oContours(oHierarchy(iContourIndex, 1)), oContours, oHierarchy, oContourTypeDictionary, oContourType, oHierarchy(iContourIndex, 1))
        End If

        ' go down
        If oHierarchy(iContourIndex, 2) <> -1 AndAlso (Not oContourTypeDictionary.ContainsKey(oHierarchy(iContourIndex, 2))) Then
            Dim oNewContourType As ContourType = ContourType.Undefined
            Select Case oContourType
                Case ContourType.External
                    oNewContourType = ContourType.ExternalNone
                Case ContourType.ExternalNone
                    oNewContourType = ContourType.Hole
                Case ContourType.Hole
                    oNewContourType = ContourType.HoleNone
                Case ContourType.HoleNone
                    oNewContourType = ContourType.External
            End Select
            SetContourType(oContours(oHierarchy(iContourIndex, 2)), oContours, oHierarchy, oContourTypeDictionary, oNewContourType, oHierarchy(iContourIndex, 2))
        End If
    End Sub
    Private Shared Function Skeletonisation(ByVal oMatrix As Emgu.CV.Matrix(Of Byte)) As Emgu.CV.Matrix(Of Byte)
        Dim oSkeletonMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
        Using oTempImageMat As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
            Using oErodedMat As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                Using oTempMat As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                    oSkeletonMatrix.SetValue(0)
                    Emgu.CV.CvInvoke.Threshold(oMatrix, oTempImageMat, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)

                    Dim done As Boolean
                    Do
                        Emgu.CV.CvInvoke.Erode(oTempImageMat, oErodedMat, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                        Emgu.CV.CvInvoke.Dilate(oErodedMat, oTempMat, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                        Emgu.CV.CvInvoke.Subtract(oTempImageMat, oTempMat, oTempImageMat)
                        Emgu.CV.CvInvoke.BitwiseOr(oSkeletonMatrix, oTempImageMat, oSkeletonMatrix)
                        oErodedMat.CopyTo(oTempImageMat)
                        done = (Emgu.CV.CvInvoke.CountNonZero(oTempImageMat) = 0)
                    Loop While Not done
                End Using
            End Using
        End Using
        Return oSkeletonMatrix
    End Function
    Public Enum ContourType As Integer
        Undefined = 0
        External
        ExternalNone
        Hole
        HoleNone
    End Enum
#End Region
#Region "Handwriting Detectors"
    Private Function SetRecognisers() As Boolean
        Const KeyMapStore As String = "KeyMapStore"
        Const FileStore As String = "HandwritingStore"
        Const FileKNN As String = "_knn"
        Const FileSVM As String = "_svm"
        Const FileDeep As String = "_DeepLearning"
        Const FileIntensity As String = "_Intensity"
        Const FileHOG As String = "_HoG"
        Const FileStroke As String = "_Stroke"
        Const SearchPatternRightGZ As String = "].gz"

        If IsNothing(oHOGDescriptor) Then
            oHOGDescriptor = New Emgu.CV.HOGDescriptor(New System.Drawing.Size(BoxSize, BoxSize), New System.Drawing.Size(BoxSize / 4, BoxSize / 4), New System.Drawing.Size(BoxSize / 4, BoxSize / 4), New System.Drawing.Size(BoxSize / 4, BoxSize / 4), 16)
        End If

        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")

        Dim oCharacterTypeList As New List(Of Enumerations.CharacterASCII)
        oCharacterTypeList.AddRange(From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.Handwriting Select oField.CharacterASCII Distinct)
        If oCharacterTypeList.Contains(Enumerations.CharacterASCII.None) Then
            oCharacterTypeList.Remove(Enumerations.CharacterASCII.None)
        End If
        If Not oCharacterTypeList.Contains(Enumerations.CharacterASCII.Numbers) Then
            oCharacterTypeList.Add(Enumerations.CharacterASCII.Numbers)
        End If

        Dim sDetectorPath As String = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location) + "\Detectors"
        Dim sBinariesPath As String = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location) + "\Detectors\Binaries"
        If Not IO.Directory.Exists(sBinariesPath) Then
            IO.Directory.CreateDirectory(sBinariesPath)
        End If

        If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap) Then
            CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap = New Dictionary(Of Enumerations.CharacterASCII, Dictionary(Of Integer, Integer))

            Dim KeyMapFileFull As String = sDetectorPath + "\" + KeyMapStore + ".gz"
            If IO.File.Exists(KeyMapFileFull) Then
                CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap.Clear()

                ' creates individualised key maps for each character type
                Dim oKeyMap As Dictionary(Of Integer, Integer) = CommonFunctions.DeserializeDataContractFile(Of Dictionary(Of Integer, Integer))(KeyMapFileFull)
                For Each oCharacterType In oCharacterTypeList
                    Dim oNewKeymap As New Dictionary(Of Integer, Integer)
                    For Each iChar In oKeyMap.Keys
                        If CommonFunctions.CheckCharacterASCII(iChar, oCharacterType) Then
                            oNewKeymap.Add(iChar, oKeyMap(iChar))
                        End If
                    Next
                    CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap.Add(oCharacterType, oNewKeymap)
                Next
            Else
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading key map file."))
                Return False
            End If
        End If

#Region "LoadRecognisers"
        Dim TaskDelegate1 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnIntensity) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnIntensity = New Dictionary(Of Enumerations.CharacterASCII, Accord.MachineLearning.KNearestNeighbors)

                                              Dim knnIntensityFileSearchPatternLeft As String = FileStore + FileKNN + FileIntensity + "["
                                              Dim knnIntensityFileSearchPatternGZ As String = knnIntensityFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, knnIntensityFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + knnIntensityFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - knnIntensityFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              Dim oFileInfo As New IO.FileInfo(oSelectedFile(0))
                                                              Dim sBinaryFile As String = oFileInfo.DirectoryName + "\Binaries\" + CommonFunctions.ReplaceExtension(oFileInfo.Name, "bin")

                                                              ' convert to binary file on first deserialisation
                                                              If Not IO.File.Exists(sBinaryFile) Then
                                                                  CommonFunctions.SerializeBinaryFile(sBinaryFile, CommonFunctions.DeserializeDataContractFile(Of Tuple(Of List(Of Double()), List(Of Integer)))(oSelectedFile(0)))
                                                              End If

                                                              Dim oTrainingAccumulator As Tuple(Of List(Of Double()), List(Of Integer)) = CommonFunctions.DeserializeBinaryFile(Of Tuple(Of List(Of Double()), List(Of Integer)))(sBinaryFile)
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnIntensity.Add(oCharacterType, New Accord.MachineLearning.KNearestNeighbors(DetectorFunctions.GetKnnNearestNeighbourCount(DetectorFunctions.Detectors.knnIntensity), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType).Count, oTrainingAccumulator.Item1.ToArray, oTrainingAccumulator.Item2.ToArray))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub

        Dim TaskDelegate2 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnHog) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnHog = New Dictionary(Of Enumerations.CharacterASCII, Accord.MachineLearning.KNearestNeighbors)

                                              Dim knnHogFileSearchPatternLeft As String = FileStore + FileKNN + FileHOG + "["
                                              Dim knnHogFileSearchPatternGZ As String = knnHogFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, knnHogFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + knnHogFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - knnHogFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              Dim oFileInfo As New IO.FileInfo(oSelectedFile(0))
                                                              Dim sBinaryFile As String = oFileInfo.DirectoryName + "\Binaries\" + CommonFunctions.ReplaceExtension(oFileInfo.Name, "bin")

                                                              ' convert to binary file on first deserialisation
                                                              If Not IO.File.Exists(sBinaryFile) Then
                                                                  CommonFunctions.SerializeBinaryFile(sBinaryFile, CommonFunctions.DeserializeDataContractFile(Of Tuple(Of List(Of Double()), List(Of Integer)))(oSelectedFile(0)))
                                                              End If

                                                              Dim oTrainingAccumulator As Tuple(Of List(Of Double()), List(Of Integer)) = CommonFunctions.DeserializeBinaryFile(Of Tuple(Of List(Of Double()), List(Of Integer)))(sBinaryFile)
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnHog.Add(oCharacterType, New Accord.MachineLearning.KNearestNeighbors(DetectorFunctions.GetKnnNearestNeighbourCount(DetectorFunctions.Detectors.knnHog), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType).Count, oTrainingAccumulator.Item1.ToArray, oTrainingAccumulator.Item2.ToArray))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub

        Dim TaskDelegate3 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmIntensityMachine) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmIntensityMachine = New Dictionary(Of Enumerations.CharacterASCII, Object)

                                              Dim SvmIntensityFileSearchPatternLeft As String = FileStore + FileSVM + FileIntensity + "["
                                              Dim SvmIntensityFileSearchPatternGZ As String = SvmIntensityFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, SvmIntensityFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + SvmIntensityFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - SvmIntensityFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ThenByDescending(Function(x) Val(x.Value(1))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmIntensityMachine.Add(oCharacterType, DetectorFunctions.LoadMulticlassSupportVectorMachine(oSelectedFile(0), DetectorFunctions.Detectors.SVMIntensity))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub

        Dim TaskDelegate4 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmHogMachine) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmHogMachine = New Dictionary(Of Enumerations.CharacterASCII, Object)

                                              Dim SvmHogFileSearchPatternLeft As String = FileStore + FileSVM + FileHOG + "["
                                              Dim SvmHogFileSearchPatternGZ As String = SvmHogFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, SvmHogFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + SvmHogFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - SvmHogFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ThenByDescending(Function(x) Val(x.Value(1))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmHogMachine.Add(oCharacterType, DetectorFunctions.LoadMulticlassSupportVectorMachine(oSelectedFile(0), DetectorFunctions.Detectors.SVMHog))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub

        Dim TaskDelegate5 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmStrokeMachine) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmStrokeMachine = New Dictionary(Of Enumerations.CharacterASCII, Object)

                                              Dim SvmStrokeFileSearchPatternLeft As String = FileStore + FileSVM + FileStroke + "["
                                              Dim SvmStrokeFileSearchPatternGZ As String = SvmStrokeFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, SvmStrokeFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + SvmStrokeFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - SvmStrokeFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ThenByDescending(Function(x) Val(x.Value(1))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmStrokeMachine.Add(oCharacterType, DetectorFunctions.LoadMulticlassSupportVectorMachine(oSelectedFile(0), DetectorFunctions.Detectors.SVMStroke))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub

        Dim TaskDelegate6 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepHogNetwork) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepHogNetwork = New Dictionary(Of Enumerations.CharacterASCII, Accord.Neuro.Networks.DeepBeliefNetwork)

                                              Dim DeepHogFileSearchPatternLeft As String = FileStore + FileDeep + FileHOG + "["
                                              Dim DeepHogFileSearchPatternGZ As String = DeepHogFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, DeepHogFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + DeepHogFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - DeepHogFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ThenByDescending(Function(x) Val(x.Value(1))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepHogNetwork.Add(oCharacterType, DetectorFunctions.LoadDeepBeliefNetwork(oSelectedFile(0)))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub

        Dim TaskDelegate7 As Action = Sub()
                                          If IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepStrokeNetwork) Then
                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepStrokeNetwork = New Dictionary(Of Enumerations.CharacterASCII, Accord.Neuro.Networks.DeepBeliefNetwork)

                                              Dim DeepStrokeFileSearchPatternLeft As String = FileStore + FileDeep + FileStroke + "["
                                              Dim DeepStrokeFileSearchPatternGZ As String = DeepStrokeFileSearchPatternLeft + "*" + SearchPatternRightGZ
                                              Dim oFiles As IEnumerable(Of String) = IO.Directory.EnumerateFiles(sDetectorPath, DeepStrokeFileSearchPatternGZ)
                                              If oFiles.Count > 0 Then
                                                  Dim oFilesDictionary As Dictionary(Of String, String()) = (From sFile As String In oFiles Select New KeyValuePair(Of String, String())(sFile, Mid(sFile, sDetectorPath.Length + 2 + DeepStrokeFileSearchPatternLeft.Length, sFile.Length - 1 - sDetectorPath.Length - DeepStrokeFileSearchPatternLeft.Length - SearchPatternRightGZ.Length).Split({"]", "["}, StringSplitOptions.RemoveEmptyEntries))).Where(Function(x) x.Value.Count = 3).OrderByDescending(Function(x) Val(x.Value(0))).ThenByDescending(Function(x) Val(x.Value(1))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                                                  If oFilesDictionary.Count > 0 Then
                                                      For Each oCharacterType In oCharacterTypeList
                                                          Dim oSelectedFile As List(Of String) = (From sFile In oFilesDictionary.Keys Where CType(Val(oFilesDictionary(sFile)(0)), Enumerations.CharacterASCII) = oCharacterType Select sFile).ToList
                                                          If oSelectedFile.Count > 0 Then
                                                              CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepStrokeNetwork.Add(oCharacterType, DetectorFunctions.LoadDeepBeliefNetwork(oSelectedFile(0)))
                                                          End If
                                                      Next
                                                  End If
                                              End If
                                          End If
                                      End Sub
#End Region

        Dim oTaskList As New List(Of Task)
        oTaskList.Add(New Task(TaskDelegate1))
        oTaskList.Add(New Task(TaskDelegate2))
        oTaskList.Add(New Task(TaskDelegate3))
        oTaskList.Add(New Task(TaskDelegate4))
        oTaskList.Add(New Task(TaskDelegate5))
        oTaskList.Add(New Task(TaskDelegate6))
        oTaskList.Add(New Task(TaskDelegate7))

        For Each oTask In oTaskList
            oTask.Start()
        Next
        Task.WaitAll(oTaskList.ToArray)
        oTaskList.Clear()

        Return True
    End Function
    Private Shared Sub ProcessCurrentWord(ByVal oCurrentWord As List(Of Integer), ByRef oField As FieldDocumentStore.Field)
        ' processes the current word
        ' if a word starts and ends with a number or character, then the inner contents are likely to contain numbers or characters respectively
        If oCurrentWord.Count <= 2 Then
            ' word too short to have contents
            For Each iImageIndex As Integer In oCurrentWord
                oField.SetImage(iImageIndex, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(iImageIndex).Item1, oField.Images(iImageIndex).Item2, Left(oField.Images(iImageIndex).Item3, 1), oField.Images(iImageIndex).Item4, oField.Images(iImageIndex).Item5, oField.Images(iImageIndex).Item6, oField.Images(iImageIndex).Item7, New Tuple(Of Single)(oField.Images(iImageIndex).Rest.Item1)))
            Next
        Else
            If CommonFunctions.CheckCharacterASCII(Left(oField.Images(oCurrentWord.First).Item3, 1), Enumerations.CharacterASCII.Numbers) And CommonFunctions.CheckCharacterASCII(Left(oField.Images(oCurrentWord.Last).Item3, 1), Enumerations.CharacterASCII.Numbers) Then
                ' both start and end are numbers
                For Each iImageIndex As Integer In oCurrentWord
                    oField.SetImage(iImageIndex, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(iImageIndex).Item1, oField.Images(iImageIndex).Item2, GetFirstChar(oField.Images(iImageIndex).Item3, Enumerations.CharacterASCII.Numbers), oField.Images(iImageIndex).Item4, oField.Images(iImageIndex).Item5, oField.Images(iImageIndex).Item6, oField.Images(iImageIndex).Item7, New Tuple(Of Single)(oField.Images(iImageIndex).Rest.Item1)))
                Next
            ElseIf CommonFunctions.CheckCharacterASCII(Left(oField.Images(oCurrentWord.First).Item3, 1), Enumerations.CharacterASCII.Lowercase Or Enumerations.CharacterASCII.Uppercase) And CommonFunctions.CheckCharacterASCII(Left(oField.Images(oCurrentWord.Last).Item3, 1), Enumerations.CharacterASCII.Lowercase Or Enumerations.CharacterASCII.Uppercase) Then
                ' both start and end are characters
                For Each iImageIndex As Integer In oCurrentWord
                    oField.SetImage(iImageIndex, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(iImageIndex).Item1, oField.Images(iImageIndex).Item2, GetFirstChar(oField.Images(iImageIndex).Item3, Enumerations.CharacterASCII.Lowercase Or Enumerations.CharacterASCII.Uppercase), oField.Images(iImageIndex).Item4, oField.Images(iImageIndex).Item5, oField.Images(iImageIndex).Item6, oField.Images(iImageIndex).Item7, New Tuple(Of Single)(oField.Images(iImageIndex).Rest.Item1)))
                Next
            Else
                ' standard processing
                For Each iImageIndex As Integer In oCurrentWord
                    oField.SetImage(iImageIndex, New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oField.Images(iImageIndex).Item1, oField.Images(iImageIndex).Item2, Left(oField.Images(iImageIndex).Item3, 1), oField.Images(iImageIndex).Item4, oField.Images(iImageIndex).Item5, oField.Images(iImageIndex).Item6, oField.Images(iImageIndex).Item7, New Tuple(Of Single)(oField.Images(iImageIndex).Rest.Item1)))
                Next
            End If
        End If
    End Sub
    Private Shared Function GetFirstChar(ByVal sText As String, ByVal oCharacterASCII As Enumerations.CharacterASCII) As String
        ' gets the first character in the string of the correct type
        ' if no character is of the correct type, then return the first character as default
        Dim sChar As String = Left(sText, 1)
        Dim oCharArray As Char() = sText.ToCharArray
        For Each sCurrentChar As Char In oCharArray
            If CommonFunctions.CheckCharacterASCII(sCurrentChar.ToString, oCharacterASCII) Then
                sChar = sCurrentChar.ToString
                Exit For
            End If
        Next
        Return sChar
    End Function
    Private Shared Function DetectHandwriting(ByVal oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)), ByVal oCharacterType As Enumerations.CharacterASCII, ByVal oHOGDescriptor As Emgu.CV.HOGDescriptor) As List(Of Tuple(Of String, Double))
        ' detects character
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
        Dim oMatrixStore As FieldDocumentStore.MatrixStore = oScannerCollection.FieldDocumentStore.FieldMatrixStore
        Dim oDetectedCharOrdered As New List(Of Tuple(Of String, Double))
        Using oUnscaledMatrix As Emgu.CV.Matrix(Of Byte) = oMatrixStore.GetMatrix(oImage.Item2)
            Dim oCropRectangle As New System.Drawing.Rectangle(SpacingSmall * 2, SpacingSmall * 2, oUnscaledMatrix.Width - (SpacingSmall * 4), oUnscaledMatrix.Height - (SpacingSmall * 4))
            Using oItemMatrix As Emgu.CV.Matrix(Of Byte) = oUnscaledMatrix.GetSubRect(oCropRectangle)
                ' check for marks
                Dim fCoveredFraction As Double = DetectorFunctions.GetCoveredFraction(oItemMatrix, False)
                If fCoveredFraction > fDetectionThreshold Then
                    ' the first item is the detector name, the third item is a dictionary of the ascii number of the characters detected together with their likelihood (from 0-1)
                    Dim oDetectedCandidateList As New ConcurrentBag(Of Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double)))

                    Dim oActionList As New List(Of Action)

                    ' runs tasks
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnIntensity)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnIntensity.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorknn(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnIntensity(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.knnIntensity))
                    End If
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnHog)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnHog.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorknn(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).knnHog(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.knnHog))
                    End If
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmIntensityMachine)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmIntensityMachine.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorSVM(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmIntensityMachine(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.SVMIntensity))
                    End If
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmHogMachine)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmHogMachine.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorSVM(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmHogMachine(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.SVMHog))
                    End If
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmStrokeMachine)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmStrokeMachine.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorSVM(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).SvmStrokeMachine(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.SVMStroke))
                    End If
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepHogNetwork)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepHogNetwork.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorDeep(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepHogNetwork(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.DeepHog))
                    End If
                    If (Not IsNothing(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepStrokeNetwork)) AndAlso CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepStrokeNetwork.ContainsKey(oCharacterType) Then
                        oActionList.Add(Sub() HandwritingDetectorDeep(oItemMatrix, oDetectedCandidateList, CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType), CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).DeepStrokeNetwork(oCharacterType), oHOGDescriptor, DetectorFunctions.Detectors.DeepStroke))
                    End If

                    CommonFunctions.ProtectedRunTasks(oActionList)

                    ' combines into a common candidate dictionary
                    Dim oCandidateDictionary As New Dictionary(Of Integer, Double)
                    For Each oCandidate As Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double)) In oDetectedCandidateList
                        For Each iCharAsc As Integer In oCandidate.Item2.Keys
                            If Not oCandidateDictionary.ContainsKey(iCharAsc) Then
                                oCandidateDictionary.Add(iCharAsc, 0)
                            End If

                            oCandidateDictionary(iCharAsc) += oCandidate.Item2(iCharAsc)
                        Next
                    Next

                    If oCandidateDictionary.Count > 0 Then
                        oDetectedCharOrdered = (From oKeyValuePair In oCandidateDictionary.OrderByDescending(Function(y) y.Value) Select New Tuple(Of String, Double)(Chr(CType(oCommonVariables.DataStore(oRecogniserStoreGUID), Recognisers).KeyMap(oCharacterType).Keys(oKeyValuePair.Key)), oKeyValuePair.Value)).ToList
                    End If
                End If
            End Using
        End Using

        Return oDetectedCharOrdered
    End Function
    Private Shared Sub HandwritingDetectorknn(ByVal oItemMatrix As Emgu.CV.Matrix(Of Byte), ByRef oDetectedCandidateList As ConcurrentBag(Of Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))), ByVal oKeyMapLocal As Dictionary(Of Integer, Integer), ByVal oknnLocal As Accord.MachineLearning.KNearestNeighbors, ByVal oHOGDescriptorLocal As Emgu.CV.HOGDescriptor, ByVal oDetector As DetectorFunctions.Detectors)
        ' processing function for knn
        Dim oDoubleImage As Double() = GetDoubleImage(oItemMatrix, oHOGDescriptorLocal, oDetector)
        Dim oLabels(DetectorFunctions.GetKnnNearestNeighbourCount(oDetector) - 1) As Integer
        Dim oNearestNeighbours As Double()() = oknnLocal.GetNearestNeighbors(oDoubleImage, oLabels)
        Dim fMinDistance As Double = oknnLocal.Distance.Distance(oDoubleImage, oNearestNeighbours(0))
        Dim oNearestNeighboursList As List(Of Tuple(Of Integer, Double)) = (From iIndex In Enumerable.Range(0, DetectorFunctions.GetKnnNearestNeighbourCount(oDetector)) Select New Tuple(Of Integer, Double)(oLabels(iIndex), 1 / Math.Exp(oknnLocal.Distance.Distance(oDoubleImage, oNearestNeighbours(iIndex)) / fMinDistance))).ToList

        Dim oLabelDictionary As Dictionary(Of Integer, Double) = (From iIndex As Integer In Enumerable.Range(0, oKeyMapLocal.Count) Select New Tuple(Of Integer, Double)(iIndex, 0)).ToDictionary(Function(x) x.Item1, Function(x) x.Item2)
        For j = 0 To oNearestNeighboursList.Count - 1
            oLabelDictionary(oNearestNeighboursList(j).Item1) += oNearestNeighboursList(j).Item2
        Next

        Dim fMaxResponse As Double = oLabelDictionary.Values.Max
        If fMaxResponse > 0 Then
            Dim oResponseDictionary As Dictionary(Of Integer, Double) = (From oLabel As KeyValuePair(Of Integer, Double) In oLabelDictionary Where oLabel.Value > 0 Order By oLabel.Value Descending Select New KeyValuePair(Of Integer, Double)(oLabel.Key, oLabel.Value)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
            oDetectedCandidateList.Add(New Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))(oDetector, DetectorFunctions.NormaliseDictionary(oResponseDictionary)))
        End If
    End Sub
    Private Shared Sub HandwritingDetectorSVM(ByVal oItemMatrix As Emgu.CV.Matrix(Of Byte), ByRef oDetectedCandidateList As ConcurrentBag(Of Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))), ByVal oKeyMapLocal As Dictionary(Of Integer, Integer), ByVal oSvmMachineLocal As Object, ByVal oHOGDescriptorLocal As Emgu.CV.HOGDescriptor, ByVal oDetector As DetectorFunctions.Detectors)
        ' processing function for svm
        Dim oKernelType = DetectorFunctions.GetSVMType(oDetector)
        Select Case oKernelType
            Case DetectorFunctions.KernelType.Linear
                HandwritingDetectorSVM(Of Accord.Statistics.Kernels.Linear)(oItemMatrix, oDetectedCandidateList, oKeyMapLocal, oSvmMachineLocal, oHOGDescriptorLocal, oDetector)
            Case DetectorFunctions.KernelType.Polynomial2, DetectorFunctions.KernelType.Polynomial5
                HandwritingDetectorSVM(Of Accord.Statistics.Kernels.Polynomial)(oItemMatrix, oDetectedCandidateList, oKeyMapLocal, oSvmMachineLocal, oHOGDescriptorLocal, oDetector)
            Case DetectorFunctions.KernelType.Gaussian1, DetectorFunctions.KernelType.Gaussian2, DetectorFunctions.KernelType.Gaussian3
                HandwritingDetectorSVM(Of Accord.Statistics.Kernels.Gaussian)(oItemMatrix, oDetectedCandidateList, oKeyMapLocal, oSvmMachineLocal, oHOGDescriptorLocal, oDetector)
            Case DetectorFunctions.KernelType.Sigmoid
                HandwritingDetectorSVM(Of Accord.Statistics.Kernels.Sigmoid)(oItemMatrix, oDetectedCandidateList, oKeyMapLocal, oSvmMachineLocal, oHOGDescriptorLocal, oDetector)
            Case DetectorFunctions.KernelType.HistogramIntersection
                HandwritingDetectorSVM(Of Accord.Statistics.Kernels.HistogramIntersection)(oItemMatrix, oDetectedCandidateList, oKeyMapLocal, oSvmMachineLocal, oHOGDescriptorLocal, oDetector)
        End Select
    End Sub
    Private Shared Sub HandwritingDetectorSVM(Of TKernel As Accord.Statistics.Kernels.IKernel(Of Double()))(ByVal oItemMatrix As Emgu.CV.Matrix(Of Byte), ByRef oDetectedCandidateList As ConcurrentBag(Of Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))), ByVal oKeyMapLocal As Dictionary(Of Integer, Integer), ByVal oSvmMachineLocal As Accord.MachineLearning.VectorMachines.MulticlassSupportVectorMachine(Of TKernel), ByVal oHOGDescriptorLocal As Emgu.CV.HOGDescriptor, ByVal oDetector As DetectorFunctions.Detectors)
        ' processing function for svm
        Dim oDoubleImage As Double() = GetDoubleImage(oItemMatrix, oHOGDescriptorLocal, oDetector)
        oSvmMachineLocal.Method = Accord.MachineLearning.VectorMachines.MulticlassComputeMethod.Voting
        Dim oResponses As Double() = oSvmMachineLocal.Probabilities(oDoubleImage)

        If oResponses.Max > 0 Then
            Dim oResponseDictionary As Dictionary(Of Integer, Double) = (From iIndex As Integer In Enumerable.Range(0, oKeyMapLocal.Count) Where oResponses(iIndex) > 0 Select New KeyValuePair(Of Integer, Double)(iIndex, oResponses(iIndex))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
            oDetectedCandidateList.Add(New Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))(oDetector, DetectorFunctions.NormaliseDictionary(oResponseDictionary)))
        End If
    End Sub
    Private Shared Sub HandwritingDetectorDeep(ByVal oItemMatrix As Emgu.CV.Matrix(Of Byte), ByRef oDetectedCandidateList As ConcurrentBag(Of Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))), ByVal oKeyMapLocal As Dictionary(Of Integer, Integer), ByVal oDeepMachineLocal As Accord.Neuro.Networks.DeepBeliefNetwork, ByVal oHOGDescriptorLocal As Emgu.CV.HOGDescriptor, ByVal oDetector As DetectorFunctions.Detectors)
        ' processing function for deep
        Dim oDoubleImage As Double() = GetDoubleImage(oItemMatrix, oHOGDescriptorLocal, oDetector)
        Dim oResponses As Double() = oDeepMachineLocal.Compute(oDoubleImage)
        If oResponses.Max > 0 Then
            Dim oResponseDictionary As Dictionary(Of Integer, Double) = (From iIndex As Integer In Enumerable.Range(0, oKeyMapLocal.Count) Where oResponses(iIndex) > 0 Select New KeyValuePair(Of Integer, Double)(iIndex, oResponses(iIndex))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
            oDetectedCandidateList.Add(New Tuple(Of DetectorFunctions.Detectors, Dictionary(Of Integer, Double))(oDetector, DetectorFunctions.NormaliseDictionary(oResponseDictionary)))
        End If
    End Sub
    Private Shared Function GetDoubleImage(ByVal oItemMatrix As Emgu.CV.Matrix(Of Byte), ByVal oHOGDescriptorLocal As Emgu.CV.HOGDescriptor, ByVal oDetector As DetectorFunctions.Detectors) As Double()
        ' initial pre-processing
        Dim oDoubleImage As Double() = Nothing
        Select Case oDetector
            Case DetectorFunctions.Detectors.knnIntensity, DetectorFunctions.Detectors.SVMIntensity, DetectorFunctions.Detectors.DeepIntensity
                oDoubleImage = Converter.MatrixToDoubleArray(DetectorFunctions.ProcessMat(DetectorFunctions.DeskewMat(oItemMatrix), BoxSize, True))
            Case DetectorFunctions.Detectors.knnHog, DetectorFunctions.Detectors.SVMHog, DetectorFunctions.Detectors.DeepHog
                oDoubleImage = Array.ConvertAll(oHOGDescriptorLocal.Compute(DetectorFunctions.ProcessMat(DetectorFunctions.DeskewMat(oItemMatrix), BoxSize, True)), Function(x) Convert.ToDouble(x))
            Case DetectorFunctions.Detectors.knnStroke, DetectorFunctions.Detectors.SVMStroke, DetectorFunctions.Detectors.DeepStroke
                Dim iStrokeBoxSize As Integer = Math.Max(oItemMatrix.Width, oItemMatrix.Height)
                Dim oSegmentList As Tuple(Of List(Of PointList.Segment), List(Of PointList.SinglePoint)) = DetectorFunctions.SegmentExtractor(DetectorFunctions.ProcessMat(DetectorFunctions.DeskewMat(oItemMatrix), iStrokeBoxSize, True))
                oDoubleImage = DetectorFunctions.ProcessSegment(oItemMatrix.Cols * ScaleFactor, oItemMatrix.Rows * ScaleFactor, oSegmentList)
        End Select
        Return oDoubleImage
    End Function
    Public Class Recognisers
        Public knnIntensity As Dictionary(Of Enumerations.CharacterASCII, Accord.MachineLearning.KNearestNeighbors) = Nothing
        Public knnHog As Dictionary(Of Enumerations.CharacterASCII, Accord.MachineLearning.KNearestNeighbors) = Nothing
        Public SvmIntensityMachine As Dictionary(Of Enumerations.CharacterASCII, Object) = Nothing
        Public SvmHogMachine As Dictionary(Of Enumerations.CharacterASCII, Object) = Nothing
        Public SvmStrokeMachine As Dictionary(Of Enumerations.CharacterASCII, Object) = Nothing
        Public DeepHogNetwork As Dictionary(Of Enumerations.CharacterASCII, Accord.Neuro.Networks.DeepBeliefNetwork) = Nothing
        Public DeepStrokeNetwork As Dictionary(Of Enumerations.CharacterASCII, Accord.Neuro.Networks.DeepBeliefNetwork) = Nothing
        Public KeyMap As Dictionary(Of Enumerations.CharacterASCII, Dictionary(Of Integer, Integer)) = Nothing
    End Class
#End Region
#Region "Alignment"
    Private Function Aligner(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal iHorizontalResolution As Integer, ByVal iVerticalResolution As Integer) As Tuple(Of Double, PointDouble, PointDouble, Emgu.CV.Matrix(Of Byte), PointDouble, Tuple(Of String, String, String, String))
        Dim oLocatedCornerstoneList As List(Of PointDouble) = LocateCornerstones(oMatrix, iHorizontalResolution, iVerticalResolution)
        Dim oPresetCornerstoneList As List(Of Tuple(Of Integer, Integer, PointDouble)) = GetCornerstonePixelLocations(iHorizontalResolution, iVerticalResolution)
        Dim oPresetCornerstonePointList As List(Of PointDouble) = (From oPoint As Tuple(Of Integer, Integer, PointDouble) In oPresetCornerstoneList Order By oPoint.Item1 Ascending Select oPoint.Item3).ToList
        Dim oDistinctIndices As List(Of Tuple(Of Integer, Integer)) = GetMatchedList(oPresetCornerstonePointList, oLocatedCornerstoneList)
        Dim oMatchedPoints As New List(Of Tuple(Of Tuple(Of Integer, Integer, PointDouble), PointDouble))
        For Each oTuple As Tuple(Of Integer, Integer) In oDistinctIndices
            oMatchedPoints.Add(New Tuple(Of Tuple(Of Integer, Integer, PointDouble), PointDouble)(oPresetCornerstoneList(oTuple.Item1), oLocatedCornerstoneList(oTuple.Item2)))
        Next

        If oMatchedPoints.Count >= 4 Then
            Dim oReturnMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
            Dim oPointList1 As New List(Of System.Drawing.PointF)
            Dim oPointList2 As New List(Of System.Drawing.PointF)
            Dim oLabelList As New List(Of Tuple(Of Integer, Integer))
            For i = 0 To oMatchedPoints.Count - 1
                oPointList1.Add(New System.Drawing.PointF(oMatchedPoints(i).Item1.Item3.X, oMatchedPoints(i).Item1.Item3.Y))
                oPointList2.Add(New System.Drawing.PointF(oMatchedPoints(i).Item2.X, oMatchedPoints(i).Item2.Y))
                oLabelList.Add(New Tuple(Of Integer, Integer)(oMatchedPoints(i).Item1.Item1, oMatchedPoints(i).Item1.Item2))
            Next

            ' calculate ransac inliers
            Dim oRansacPointArray As System.Drawing.PointF()() = New System.Drawing.PointF()() {oPointList1.ToArray, oPointList2.ToArray}
            Dim oRansacHomographyEstimator As New Accord.Imaging.RansacHomographyEstimator(0.001, 0.99)
            Dim oMatrixH As Accord.Imaging.MatrixH = oRansacHomographyEstimator.Estimate(oRansacPointArray)

            ' calculate maximum displacement of matched points
            Dim fMaxDisplacementX As Double = Aggregate iIndex As Integer In oRansacHomographyEstimator.Inliers Into Max(Math.Abs(oMatchedPoints(iIndex).Item1.Item3.X - oMatchedPoints(iIndex).Item2.X))
            Dim fMaxDisplacementY As Double = Aggregate iIndex As Integer In oRansacHomographyEstimator.Inliers Into Max(Math.Abs(oMatchedPoints(iIndex).Item1.Item3.Y - oMatchedPoints(iIndex).Item2.Y))

            If oRansacHomographyEstimator.Inliers.Count < 4 Then
                Return Nothing
            Else
                ' trim matched points by removing outliers
                Dim oNewPointList1A As New List(Of System.Drawing.PointF)
                Dim oNewPointList1B As New List(Of System.Drawing.PointF)
                Dim oNewPointList2 As New List(Of System.Drawing.PointF)
                For Each iIndex In oRansacHomographyEstimator.Inliers
                    oNewPointList1A.Add(oPointList1(iIndex))
                    oNewPointList2.Add(oPointList2(iIndex))

                    ' get rotated match
                    Dim iTranslatedIndex As Integer = (From iCurrentIndex In Enumerable.Range(0, oLabelList.Count) Where oLabelList(iCurrentIndex).Item2 = oLabelList(iIndex).Item1 Select iCurrentIndex).First
                    oNewPointList1B.Add(oPointList1(iTranslatedIndex))
                Next

                ' set barcodes
                Dim oBarcodeRect As New System.Drawing.Rectangle(PDFHelper.BarcodeBoundaries.X * iHorizontalResolution / 72, PDFHelper.BarcodeBoundaries.Y * iVerticalResolution / 72, PDFHelper.BarcodeBoundaries.Width * iHorizontalResolution / 72, PDFHelper.BarcodeBoundaries.Height * iVerticalResolution / 72)
                Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                Dim oBarcodeDataList As New List(Of String)
                For Each oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore
                    oBarcodeDataList.AddRange(oFieldCollection.BarCodes)
                Next

                ' use emgu homography perspective correction
                Dim oProcessedBarcode As Tuple(Of String, String, String, String) = Nothing
                Dim oNewPointList As System.Drawing.PointF() = Nothing
                Using oHomographyMat As New Emgu.CV.Mat(3, 3, Emgu.CV.CvEnum.DepthType.Cv64F, 1)
                    Emgu.CV.CvInvoke.FindHomography(oNewPointList1A.ToArray, oNewPointList2.ToArray, oHomographyMat, Emgu.CV.CvEnum.HomographyMethod.Default)
                    Emgu.CV.CvInvoke.WarpPerspective(oMatrix, oReturnMatrix, oHomographyMat, oReturnMatrix.Size, Emgu.CV.CvEnum.Inter.Lanczos4, Emgu.CV.CvEnum.Warp.FillOutliers Or Emgu.CV.CvEnum.Warp.InverseMap, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MaxValue))

                    oProcessedBarcode = PDFHelper.ProcessBarcode(m_CommonScanner, oReturnMatrix.GetSubRect(oBarcodeRect), oBarcodeDataList)

                    ' rotate page
                    If IsNothing(oProcessedBarcode) Then
                        Emgu.CV.CvInvoke.FindHomography(oNewPointList1B.ToArray, oNewPointList2.ToArray, oHomographyMat, Emgu.CV.CvEnum.HomographyMethod.Default)
                        Emgu.CV.CvInvoke.WarpPerspective(oMatrix, oReturnMatrix, oHomographyMat, oReturnMatrix.Size, Emgu.CV.CvEnum.Inter.Lanczos4, Emgu.CV.CvEnum.Warp.FillOutliers Or Emgu.CV.CvEnum.Warp.InverseMap, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MaxValue))

                        oProcessedBarcode = PDFHelper.ProcessBarcode(m_CommonScanner, oReturnMatrix.GetSubRect(oBarcodeRect), oBarcodeDataList)

                        If IsNothing(oProcessedBarcode) Then
                            Return Nothing
                        Else
                            oNewPointList = oNewPointList1B.ToArray
                        End If
                    Else
                        oNewPointList = oNewPointList1A.ToArray
                    End If
                End Using

                ' determine rotation and scale
                Using oRotationMat As Emgu.CV.Mat = Emgu.CV.CvInvoke.EstimateRigidTransform(oNewPointList, oNewPointList2.ToArray, False)
                    Using oInverseRotationMat As Emgu.CV.Mat = Emgu.CV.CvInvoke.EstimateRigidTransform(oNewPointList2.ToArray, oNewPointList, False)
                        If IsNothing(oRotationMat) And IsNothing(oInverseRotationMat) Then
                            Return New Tuple(Of Double, PointDouble, PointDouble, Emgu.CV.Matrix(Of Byte), PointDouble, Tuple(Of String, String, String, String))(Double.NaN, PointDouble.NaN, PointDouble.NaN, oReturnMatrix, PointDouble.NaN, New Tuple(Of String, String, String, String)(String.Empty, String.Empty, String.Empty, String.Empty))
                        Else
                            Using oRotationMatrix As New Emgu.CV.Matrix(Of Double)(2, 3)
                                oRotationMat.CopyTo(oRotationMatrix)
                                Using oInverseRotationMatrix As New Emgu.CV.Matrix(Of Double)(2, 3)
                                    oInverseRotationMat.CopyTo(oInverseRotationMatrix)

                                    Dim a = oRotationMatrix(0, 0)
                                    Dim b = oRotationMatrix(0, 1)
                                    Dim tx = oRotationMatrix(0, 2)
                                    Dim c = oRotationMatrix(1, 0)
                                    Dim d = oRotationMatrix(1, 1)
                                    Dim ty = oRotationMatrix(1, 2)

                                    Dim a1 = oInverseRotationMatrix(0, 0)
                                    Dim b1 = oInverseRotationMatrix(0, 1)
                                    Dim tx1 = oInverseRotationMatrix(0, 2)
                                    Dim c1 = oInverseRotationMatrix(1, 0)
                                    Dim d1 = oInverseRotationMatrix(1, 1)
                                    Dim ty1 = oInverseRotationMatrix(1, 2)

                                    Dim fAngle As Double = ((Math.Atan2(-b1, a1) + Math.Atan2(c1, d1) - Math.Atan2(-b, a) - Math.Atan2(c, d)) / 4) * 180 / Math.PI
                                    Dim oScale As PointDouble = (New PointDouble(Math.Sign(a) * Math.Sqrt((a * a) + (b * b)), Math.Sign(d) * Math.Sqrt((c * c) + (d * d)))).Add(New PointDouble(1 / (Math.Sign(a1) * Math.Sqrt((a1 * a1) + (b1 * b1))), 1 / (Math.Sign(d1) * Math.Sqrt((c1 * c1) + (d1 * d1))))).Divide(2)

                                    oScale = oScale.Abs
                                    fAngle = ((fAngle + 180) Mod 360) - 180

                                    ' calculated drift needs to be derotated & descaled to get original drift distance
                                    Dim oCenter As New PointDouble((oMatrix.Width - 1) / 2, (oMatrix.Height - 1) / 2)
                                    Dim fULDistance As Double = oCenter.DistanceTo(PointDouble.Zero)
                                    Dim fULAngle As Double = oCenter.AngleTo(PointDouble.Zero)
                                    Dim fULRotatedAngle As Double = fULAngle - (fAngle * Math.PI / 180)
                                    Dim oULDisp As New PointDouble(oCenter.X + fULDistance * Math.Cos(fULRotatedAngle), oCenter.Y + fULDistance * Math.Sin(fULRotatedAngle))

                                    Dim oUnrotatedDisp As PointDouble = (New PointDouble(oULDisp.X * a1 + oULDisp.Y * b1 + tx1, oULDisp.X * c1 + oULDisp.Y * d1 + ty1))
                                    Dim fUnrotatedDistance As Double = PointDouble.Zero.DistanceTo(oUnrotatedDisp)
                                    Dim fUnrotatedAngle As Double = PointDouble.Zero.AngleTo(oUnrotatedDisp)
                                    Dim fRotatedAngle As Double = fUnrotatedAngle - (fAngle * Math.PI / 180)
                                    Dim oDrift As PointDouble = (New PointDouble(fUnrotatedDistance * Math.Cos(fRotatedAngle), fUnrotatedDistance * Math.Sin(fRotatedAngle))).Invert.Multiply(oScale)

                                    Return New Tuple(Of Double, PointDouble, PointDouble, Emgu.CV.Matrix(Of Byte), PointDouble, Tuple(Of String, String, String, String))(fAngle, oScale, oDrift, oReturnMatrix, New PointDouble(fMaxDisplacementX, fMaxDisplacementY), oProcessedBarcode)
                                End Using
                            End Using
                        End If
                    End Using
                End Using
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function LocateCornerstones(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal iHorizontalResolution As Integer, ByVal iVerticalResolution As Integer) As List(Of PointDouble)
        Using oCornerstoneMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(Converter.BitmapConvertGrayscale(PDFHelper.GetCornerstoneImage(SpacingSmall, iHorizontalResolution, iVerticalResolution, True)))
            Dim oCornerstoneLocations As List(Of Tuple(Of Double, PointDouble)) = LocateImage(oMatrix, oCornerstoneMatrix, PDFHelper.Cornerstone.Count)
            Return (From oCornerstoneLocation As Tuple(Of Double, PointDouble) In oCornerstoneLocations Select oCornerstoneLocation.Item2).ToList
        End Using
    End Function
    Private Shared Function GetCornerstonePixelLocations(ByVal iHorizontalResolution As Integer, ByVal iVerticalResolution As Integer) As List(Of Tuple(Of Integer, Integer, PointDouble))
        ' converts preset cornerstone locations in points to pixels based on the supplied resolution in DPI
        Dim oReturnList As New List(Of Tuple(Of Integer, Integer, PointDouble))
        For Each oCornerstone As Tuple(Of Integer, Integer, XPoint) In PDFHelper.Cornerstone
            oReturnList.Add(New Tuple(Of Integer, Integer, PointDouble)(oCornerstone.Item1, oCornerstone.Item2, New PointDouble(oCornerstone.Item3.X * iHorizontalResolution / 72, oCornerstone.Item3.Y * iVerticalResolution / 72)))
        Next
        Return oReturnList
    End Function
    Private Shared Function GetMatchedList(ByVal oPresetCornerstoneList As List(Of PointDouble), ByVal oLocatedCornerstoneList As List(Of PointDouble)) As List(Of Tuple(Of Integer, Integer))
        ' matches the list of points and returns a list of matching indices in from the supplied lists
        Const fMatchThreshold As Double = 0.005
        Const fModeThreshold As Double = 0.02
        Dim oMatchedIndices As New List(Of Tuple(Of Integer, Integer))
        Dim oTriangleList1 As New List(Of Tuple(Of Integer, Integer, Integer))
        Dim oTriangleList2 As New List(Of Tuple(Of Integer, Integer, Integer))
        Dim oMatchedPointsList1 As New List(Of MatchedPoints)
        Dim oMatchedPointsList2 As New List(Of MatchedPoints)
        Dim oBestMatchList As New List(Of Tuple(Of Integer, Integer, Double, Double))
        Dim TaskDelegate As Action(Of Object) = Nothing

        If oPresetCornerstoneList.Count >= 3 And oLocatedCornerstoneList.Count >= 3 Then
            ' create all combinations of triangles from these points
            oTriangleList1.AddRange(GetTriangleList(oPresetCornerstoneList.Count))
            If oPresetCornerstoneList.Count = oLocatedCornerstoneList.Count Then
                oTriangleList2.AddRange(oTriangleList1)
            Else
                oTriangleList2.AddRange(GetTriangleList(oLocatedCornerstoneList.Count))
            End If

            ' for each triangle, initialise with the actual points
            Dim oMatchedPointsArray(TaskCount - 1) As List(Of MatchedPoints)
            TaskDelegate = Sub(oParamIn As Object)
                               Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                               Dim oTriangleListLocal As List(Of Tuple(Of Integer, Integer, Integer)) = CType(oParam.Item4, List(Of Tuple(Of Integer, Integer, Integer)))
                               Dim oPointsLocal As List(Of PointDouble) = CType(oParam.Item5, List(Of PointDouble))
                               Dim oMatchedPointsArrayLocal As New List(Of MatchedPoints)
                               For y = 0 To oParam.Item2 - 1
                                   Dim localY As Integer = y + oParam.Item1
                                   oMatchedPointsArrayLocal.Add(New MatchedPoints(oPointsLocal(oTriangleListLocal(localY).Item1), oPointsLocal(oTriangleListLocal(localY).Item2), oPointsLocal(oTriangleListLocal(localY).Item3), oTriangleListLocal(localY).Item1, oTriangleListLocal(localY).Item2, oTriangleListLocal(localY).Item3))
                               Next
                               CType(oParam.Item6, List(Of MatchedPoints)())(oParam.Item3) = oMatchedPointsArrayLocal
                           End Sub
            CommonFunctions.ParallelRun(oTriangleList1.Count, oTriangleList1, oPresetCornerstoneList, oMatchedPointsArray, Nothing, TaskDelegate)

            For i = 0 To TaskCount - 1
                If Not IsNothing(oMatchedPointsArray(i)) Then
                    oMatchedPointsList1.AddRange(oMatchedPointsArray(i))
                    oMatchedPointsArray(i) = Nothing
                End If
            Next

            CommonFunctions.ParallelRun(oTriangleList2.Count, oTriangleList2, oLocatedCornerstoneList, oMatchedPointsArray, Nothing, TaskDelegate)

            For i = 0 To TaskCount - 1
                If Not IsNothing(oMatchedPointsArray(i)) Then
                    oMatchedPointsList2.AddRange(oMatchedPointsArray(i))
                    oMatchedPointsArray(i) = Nothing
                End If
            Next

            ' item1=index1, item2=index2, item3=compare
            Dim oBestMatchArray(TaskCount - 1) As List(Of Tuple(Of Integer, Integer, Double, Double))
            TaskDelegate = Sub(oParamIn As Object)
                               Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                               Dim oMatchedPointsListLocal1 As List(Of MatchedPoints) = CType(oParam.Item4, List(Of MatchedPoints))
                               Dim oMatchedPointsListLocal2 As List(Of MatchedPoints) = CType(oParam.Item5, List(Of MatchedPoints))
                               Dim oBestMatchArrayLocal As New List(Of Tuple(Of Integer, Integer, Double, Double))
                               Dim oCurrentMatchList As New List(Of Tuple(Of Integer, Double, Double))
                               For y = 0 To oParam.Item2 - 1
                                   Dim localY As Integer = y + oParam.Item1
                                   oCurrentMatchList.Clear()
                                   For x = 0 To oMatchedPointsListLocal2.Count - 1
                                       oCurrentMatchList.Add(New Tuple(Of Integer, Double, Double)(x, MatchedPoints.Compare(oMatchedPointsListLocal1(localY), oMatchedPointsListLocal2(x)), MatchedPoints.AngleDifference(oMatchedPointsListLocal1(localY), oMatchedPointsListLocal2(x))))
                                   Next
                                   If oCurrentMatchList.Count > 0 Then
                                       Dim oBestMatch = oCurrentMatchList.OrderBy(Of Double)(Function(x) x.Item2).First
                                       oBestMatchArrayLocal.Add(New Tuple(Of Integer, Integer, Double, Double)(localY, oBestMatch.Item1, oBestMatch.Item2, oBestMatch.Item3))
                                   End If
                               Next
                               CType(oParam.Item6, List(Of Tuple(Of Integer, Integer, Double, Double))())(oParam.Item3) = oBestMatchArrayLocal
                           End Sub
            CommonFunctions.ParallelRun(oMatchedPointsList1.Count, oMatchedPointsList1, oMatchedPointsList2, oBestMatchArray, Nothing, TaskDelegate)

            For i = 0 To TaskCount - 1
                If Not IsNothing(oBestMatchArray(i)) Then
                    oBestMatchList.AddRange(oBestMatchArray(i))
                    oBestMatchArray(i) = Nothing
                End If
            Next
            oBestMatchList = (From oTuple As Tuple(Of Integer, Integer, Double, Double) In oBestMatchList Where oTuple.Item3 < fMatchThreshold Order By oTuple.Item3 Select oTuple).ToList

            If oBestMatchList.Count > 0 Then
                ' use meanshift algorithm to select the mode
                Dim oUniformKernel As New Accord.Statistics.Distributions.DensityKernels.UniformKernel
                Dim oMeanShift As New Accord.MachineLearning.MeanShift(1, oUniformKernel, 0.01)
                Dim oInput As List(Of Double()) = (From oTuple As Tuple(Of Integer, Integer, Double, Double) In oBestMatchList Select New Double() {oTuple.Item4}).ToList
                Dim oClustering As Accord.MachineLearning.MeanShiftClusterCollection = oMeanShift.Learn(oInput.ToArray)
                Dim oResult As Dictionary(Of Integer, List(Of Double)) = (From iIndex As Integer In Enumerable.Range(0, oClustering.Count) Select New KeyValuePair(Of Integer, List(Of Double))(iIndex, (From oPoint In oInput Where iIndex = oClustering.Decide(oPoint) Select oPoint.First).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value).OrderByDescending(Function(x) x.Value.Count).ToDictionary(Function(x) x.Key, Function(x) x.Value)

                Dim fMode As Double = oResult.Values.First.Average
                oBestMatchList = (From oTuple In oBestMatchList Where oTuple.Item4 > fMode - fModeThreshold And oTuple.Item4 < fMode + fModeThreshold Select oTuple).ToList

                ' from these triangles extract unique blob pairs
                For Each oTuple As Tuple(Of Integer, Integer, Double, Double) In oBestMatchList
                    oMatchedIndices.Add(New Tuple(Of Integer, Integer)(oMatchedPointsList1(oTuple.Item1).Index1, oMatchedPointsList2(oTuple.Item2).Index1))
                    oMatchedIndices.Add(New Tuple(Of Integer, Integer)(oMatchedPointsList1(oTuple.Item1).Index2, oMatchedPointsList2(oTuple.Item2).Index2))
                    oMatchedIndices.Add(New Tuple(Of Integer, Integer)(oMatchedPointsList1(oTuple.Item1).Index3, oMatchedPointsList2(oTuple.Item2).Index3))
                Next
            End If
        End If

        Dim oReturnIndices As New List(Of Tuple(Of Integer, Integer))
        oReturnIndices.AddRange(oMatchedIndices.Distinct)

        ' clean up
        oMatchedIndices = Nothing
        oTriangleList1 = Nothing
        oTriangleList2 = Nothing
        oMatchedPointsList1 = Nothing
        oMatchedPointsList2 = Nothing
        oBestMatchList = Nothing
        TaskDelegate = Nothing

        Return oReturnIndices
    End Function
    Private Shared Function GetTriangleList(ByVal iCount As Integer) As List(Of Tuple(Of Integer, Integer, Integer))
        ' generates point triplets from list (only index counts generated)
        Dim oTriangleList As New List(Of Tuple(Of Integer, Integer, Integer))
        If iCount >= 3 Then
            Dim oTriangleArray(TaskCount - 1) As List(Of Tuple(Of Integer, Integer, Integer))
            Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                        Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                        Dim iCountLocal As Integer = CType(oParam.Item4, Integer)
                                                        Dim oTriangleArrayLocal As New List(Of Tuple(Of Integer, Integer, Integer))
                                                        For x = 0 To oParam.Item2 - 1
                                                            Dim localX As Integer = x + oParam.Item1
                                                            For y = 1 To iCountLocal - 1
                                                                For z = 2 To iCountLocal - 1
                                                                    If localX <> y And localX <> z And y <> z Then
                                                                        Dim oNumArray As Integer() = New Integer() {localX, y, z}
                                                                        Array.Sort(oNumArray)
                                                                        oTriangleArrayLocal.Add(New Tuple(Of Integer, Integer, Integer)(oNumArray.First, oNumArray(1), oNumArray(2)))
                                                                    End If
                                                                Next
                                                            Next
                                                        Next
                                                        CType(oParam.Item5, List(Of Tuple(Of Integer, Integer, Integer))())(oParam.Item3) = oTriangleArrayLocal
                                                    End Sub
            CommonFunctions.ParallelRun(iCount, iCount, oTriangleArray, Nothing, Nothing, TaskDelegate)

            For i = 0 To TaskCount - 1
                If Not IsNothing(oTriangleArray(i)) Then
                    oTriangleList.AddRange(oTriangleArray(i))
                    oTriangleArray(i) = Nothing
                End If
            Next

            ' clean up
            TaskDelegate = Nothing
        End If
        Return oTriangleList.Distinct.ToList
    End Function
    Private Structure MatchedPoints
        Private m_Point() As PointDouble
        Private m_Index() As Integer
        Private m_Side() As Double
        Private m_Angle() As Double

        Public Sub New(ByVal oPoint1 As PointDouble, ByVal oPoint2 As PointDouble, ByVal oPoint3 As PointDouble, ByVal iIndex1 As Integer, ByVal iIndex2 As Integer, ByVal iIndex3 As Integer)
            ReDim m_Point(2)
            ReDim m_Index(2)
            ReDim m_Side(2)
            ReDim m_Angle(2)
            m_Point(0) = oPoint1
            m_Point(1) = oPoint2
            m_Point(2) = oPoint3
            m_Index(0) = iIndex1
            m_Index(1) = iIndex2
            m_Index(2) = iIndex3
            Recalculate()
        End Sub
        Public ReadOnly Property Point1 As PointDouble
            Get
                Return m_Point(0)
            End Get
        End Property
        Public ReadOnly Property Point2 As PointDouble
            Get
                Return m_Point(1)
            End Get
        End Property
        Public ReadOnly Property Point3 As PointDouble
            Get
                Return m_Point(2)
            End Get
        End Property
        Public ReadOnly Property Index1 As Integer
            Get
                Return m_Index(0)
            End Get
        End Property
        Public ReadOnly Property Index2 As Integer
            Get
                Return m_Index(1)
            End Get
        End Property
        Public ReadOnly Property Index3 As Integer
            Get
                Return m_Index(2)
            End Get
        End Property
        Public ReadOnly Property Side1 As Double
            Get
                Return m_Side(0)
            End Get
        End Property
        Public ReadOnly Property Side2 As Double
            Get
                Return m_Side(1)
            End Get
        End Property
        Public ReadOnly Property Side3 As Double
            Get
                Return m_Side(2)
            End Get
        End Property
        Public ReadOnly Property Angle1 As Double
            Get
                Return m_Angle(0)
            End Get
        End Property
        Public ReadOnly Property Angle2 As Double
            Get
                Return m_Angle(1)
            End Get
        End Property
        Public ReadOnly Property Angle3 As Double
            Get
                Return m_Angle(2)
            End Get
        End Property
        Private Function OrderedPoints() As Integer()
            ' return the points ordered by the angle opposite the point
            Dim oOrderedPoints(2) As Tuple(Of Integer, Double)
            oOrderedPoints(0) = New Tuple(Of Integer, Double)(0, m_Angle(0))
            oOrderedPoints(1) = New Tuple(Of Integer, Double)(1, m_Angle(1))
            oOrderedPoints(2) = New Tuple(Of Integer, Double)(2, m_Angle(2))
            Dim oSortedPoints() As Tuple(Of Integer, Double) = oOrderedPoints.OrderBy(Function(x) x.Item2).ToArray
            Return New Integer(2) {oSortedPoints(0).Item1, oSortedPoints(1).Item1, oSortedPoints(2).Item1}
        End Function
        Public Shared Function Compare(ByVal oMatchedPoints1 As MatchedPoints, ByVal oMatchedPoints2 As MatchedPoints) As Double
            ' returns a number showing the sum of the normalised residuals between the angles and sides of the two triangles
            Dim oSide1 As Double() = {oMatchedPoints1.Side1, oMatchedPoints1.Side2, oMatchedPoints1.Side3}
            Dim oSide2 As Double() = {oMatchedPoints2.Side1, oMatchedPoints2.Side2, oMatchedPoints2.Side3}
            Dim oAngle1 As Double() = {oMatchedPoints1.Angle1, oMatchedPoints1.Angle2, oMatchedPoints1.Angle3}
            Dim oAngle2 As Double() = {oMatchedPoints2.Angle1, oMatchedPoints2.Angle2, oMatchedPoints2.Angle3}

            Normalise(oSide1)
            Normalise(oSide2)
            Normalise(oAngle1)
            Normalise(oAngle2)

            Return Difference(oSide1, oSide2) + Difference(oAngle1, oAngle2)
        End Function
        Public Shared Function AngleDifference(ByVal oMatchedPoints1 As MatchedPoints, ByVal oMatchedPoints2 As MatchedPoints) As Double
            ' computes the sum of the angles of all the sides and finds the difference between the two sets
            Dim fAngle1 As Double() = {GetLineAngle(oMatchedPoints1.Point1, oMatchedPoints1.Point2), GetLineAngle(oMatchedPoints1.Point2, oMatchedPoints1.Point3), GetLineAngle(oMatchedPoints1.Point3, oMatchedPoints1.Point1)}
            Dim fAngle2 As Double() = {GetLineAngle(oMatchedPoints2.Point1, oMatchedPoints2.Point2), GetLineAngle(oMatchedPoints2.Point2, oMatchedPoints2.Point3), GetLineAngle(oMatchedPoints2.Point3, oMatchedPoints2.Point1)}

            Dim fAngle As Double = 0
            For i = 0 To 2
                fAngle += fAngle1(i) - fAngle2(i)
            Next
            Return Math.Round(fAngle, 3)
        End Function
        Private Shared Function GetLineAngle(ByVal oPoint1 As PointDouble, ByVal oPoint2 As PointDouble) As Double
            Return Math.Atan2(oPoint1.Y - oPoint2.Y, oPoint1.X - oPoint2.X)
        End Function
        Private Shared Sub Normalise(ByRef oArray() As Double)
            Dim fSum As Double = 0
            For i = 0 To oArray.Length - 1
                fSum += Math.Abs(oArray(i))
            Next
            For i = 0 To oArray.Length - 1
                oArray(i) = Math.Abs(oArray(i)) / fSum
            Next
        End Sub
        Private Shared Function Difference(ByVal oArray1() As Double, ByVal oArray2() As Double) As Double
            Dim fSum As Double = 0
            For i = 0 To Math.Min(oArray1.Length, oArray2.Length) - 1
                fSum += Math.Abs(oArray1(i) - oArray2(i))
            Next
            Return fSum
        End Function
        Private Sub Recalculate()
            m_Angle(0) = GetAngle(m_Point(2), m_Point(0), m_Point(1))
            m_Angle(1) = GetAngle(m_Point(0), m_Point(1), m_Point(2))
            m_Angle(2) = GetAngle(m_Point(1), m_Point(2), m_Point(0))

            ' reorder according to angle size
            Dim oOrderedPoints(2) As PointDouble
            Dim oOrderedIndex(2) As Integer
            Dim oOrder As Integer() = OrderedPoints()

            For i = 0 To 2
                oOrderedPoints(i) = m_Point(oOrder(i))
                oOrderedIndex(i) = m_Index(oOrder(i))
            Next
            For i = 0 To 2
                m_Point(i) = oOrderedPoints(i)
                m_Index(i) = oOrderedIndex(i)
            Next

            m_Side(0) = DistanceTo(m_Point(0), m_Point(1))
            m_Side(1) = DistanceTo(m_Point(1), m_Point(2))
            m_Side(2) = DistanceTo(m_Point(2), m_Point(0))

            m_Angle(0) = GetAngle(m_Point(2), m_Point(0), m_Point(1))
            m_Angle(1) = GetAngle(m_Point(0), m_Point(1), m_Point(2))
            m_Angle(2) = GetAngle(m_Point(1), m_Point(2), m_Point(0))
        End Sub
        Private Function GetAngle(ByVal oPoint1 As PointDouble, ByVal oPoint2 As PointDouble, ByVal oPoint3 As PointDouble) As Double
            Dim fAngle1 As Double = Math.Atan2(oPoint1.Y - oPoint2.Y, oPoint1.X - oPoint2.X)
            Dim fAngle2 As Double = Math.Atan2(oPoint3.Y - oPoint2.Y, oPoint3.X - oPoint2.X)
            Return Math.Abs(fAngle1 - fAngle2)
        End Function
        Private Function DistanceTo(ByVal oPoint1 As PointDouble, ByVal oPoint2 As PointDouble) As Double
            Return Math.Sqrt(((oPoint2.X - oPoint1.X) * (oPoint2.X - oPoint1.X)) + ((oPoint2.Y - oPoint1.Y) * (oPoint2.Y - oPoint1.Y)))
        End Function
    End Structure
    Public Shared Function LocateImage(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal oLocateMatrix As Emgu.CV.Matrix(Of Byte), ByVal iMaxCount As Integer) As List(Of Tuple(Of Double, PointDouble))
        ' extracts a list of cornerstone weighted centroids from a scanned image
        ' the first element of the tuple represents the max value within the result map of that cornerstone (ie. indicating of the strength of the match)
        Dim oReturnBag As New ConcurrentBag(Of Tuple(Of Double, PointDouble))
        Dim oReturnList As List(Of Tuple(Of Double, PointDouble)) = Nothing

        Dim iVerticalStep As Integer = oLocateMatrix.Height * 2
        Dim iHorizontalStep As Integer = oLocateMatrix.Width * 2
        Dim iRadius As Integer = Math.Ceiling(Math.Max(oLocateMatrix.Width, oLocateMatrix.Height) / 2)
        If Not IsNothing(oMatrix) Then
            Emgu.CV.CvInvoke.Normalize(oMatrix, oMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
            Emgu.CV.CvInvoke.Normalize(oLocateMatrix, oLocateMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)

            Dim iHeight As Integer = Math.Floor(oMatrix.Height / oLocateMatrix.Height)
            Dim iWidth As Integer = Math.Floor(oMatrix.Width / oLocateMatrix.Width)

            Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                        Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                        Dim oSourceMatrixLocal As Emgu.CV.Matrix(Of Byte) = CType(oParam.Item4, Emgu.CV.Matrix(Of Byte))
                                                        Dim oLocateMatrixLocal As Emgu.CV.Matrix(Of Byte) = CType(oParam.Item5, Emgu.CV.Matrix(Of Byte))
                                                        Dim localwidth As Integer = CType(oParam.Item6, Tuple(Of Integer, Integer, Integer, Integer)).Item1
                                                        Dim localHorizontalStep As Integer = CType(oParam.Item6, Tuple(Of Integer, Integer, Integer, Integer)).Item2
                                                        Dim localVerticalStep As Integer = CType(oParam.Item6, Tuple(Of Integer, Integer, Integer, Integer)).Item3
                                                        Dim localMaxCount As Integer = CType(oParam.Item6, Tuple(Of Integer, Integer, Integer, Integer)).Item4
                                                        Dim oReturnBagLocal As ConcurrentBag(Of Tuple(Of Double, PointDouble)) = CType(oParam.Item7, ConcurrentBag(Of Tuple(Of Double, PointDouble)))
                                                        Dim iCornerstoneWidth As Integer = oLocateMatrixLocal.Width
                                                        Dim iCornerstoneHeight As Integer = oLocateMatrixLocal.Height
                                                        Dim iSourceMatrixWidth As Integer = oSourceMatrixLocal.Width
                                                        Dim iSourceMatrixHeight As Integer = oSourceMatrixLocal.Height

                                                        For y = 0 To oParam.Item2 - 1
                                                            Dim localY As Integer = y + oParam.Item1
                                                            For x = 0 To localwidth - 1
                                                                Dim iHorizontalDisplacement As Integer = Math.Min(x * iCornerstoneWidth, iSourceMatrixWidth - localHorizontalStep)
                                                                Dim iVerticalDisplacement As Integer = Math.Min(localY * iCornerstoneHeight, iSourceMatrixHeight - localVerticalStep)

                                                                Dim oReturnPoints As List(Of Tuple(Of Double, PointDouble)) = LocatePoint(oSourceMatrixLocal, oLocateMatrixLocal, localMaxCount, iHorizontalDisplacement, iVerticalDisplacement, localHorizontalStep, localVerticalStep)
                                                                For Each oPoint In oReturnPoints
                                                                    oReturnBagLocal.Add(oPoint)
                                                                Next
                                                            Next
                                                        Next
                                                    End Sub
            CommonFunctions.ParallelRun(iHeight, oMatrix, oLocateMatrix, New Tuple(Of Integer, Integer, Integer, Integer)(iWidth, iHorizontalStep, iVerticalStep, iMaxCount), oReturnBag, TaskDelegate)

            oReturnList = oReturnBag.ToList
            EliminateDuplicatePoints(oMatrix, oLocateMatrix, oReturnList, iRadius)
        End If

        If oReturnList.Count > iMaxCount Then
            oReturnList = oReturnList.Take(iMaxCount).ToList
        End If
        Return oReturnList
    End Function
    Public Shared Function SimpleLocateImage(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal oLocateMatrix As Emgu.CV.Matrix(Of Byte), ByVal iMaxCount As Integer) As List(Of Tuple(Of Double, PointDouble))
        ' extracts a list of cornerstone weighted centroids from a scanned image
        ' the first element of the tuple represents the max value within the result map of that cornerstone (ie. indicating of the strength of the match)
        ' resize the cornerstone to match the scan resolution of the image
        Using oSourceMatrix As Emgu.CV.Matrix(Of Byte) = oMatrix.Clone
            Emgu.CV.CvInvoke.Normalize(oSourceMatrix, oSourceMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
            Dim oCornerstoneLocations As New List(Of Tuple(Of Double, PointDouble))

            Using oCornerstone As Emgu.CV.Matrix(Of Byte) = oLocateMatrix.Clone
                Emgu.CV.CvInvoke.Normalize(oCornerstone, oCornerstone, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                oCornerstoneLocations.AddRange(LocatePoint(oSourceMatrix, oCornerstone, iMaxCount, 0, 0, oSourceMatrix.Width, oSourceMatrix.Height))
                Return oCornerstoneLocations
            End Using
        End Using
    End Function
    Private Shared Function LocatePoint(ByVal oSourceMatrix As Emgu.CV.Matrix(Of Byte), ByVal oCornerstone As Emgu.CV.Matrix(Of Byte), ByVal iMaxCount As Integer, ByVal iHorizontalDisplacement As Integer, ByVal iVerticalDisplacement As Integer, ByVal iHorizontalStep As Integer, ByVal iVerticalStep As Integer) As List(Of Tuple(Of Double, PointDouble))
        ' locates the point in the suppplied matrix
        Dim oReturnList As New List(Of Tuple(Of Double, PointDouble))
        Dim oRect As New System.Drawing.Rectangle(iHorizontalDisplacement, iVerticalDisplacement, iHorizontalStep, iVerticalStep)
        Using oSubMatrix As Emgu.CV.Matrix(Of Byte) = oSourceMatrix.GetSubRect(oRect)
            ' perform template matching of the cornerstone to the submatrix
            Using oResultMap As New Emgu.CV.Matrix(Of Single)(oSubMatrix.Rows - oCornerstone.Rows + 1, oSubMatrix.Cols - oCornerstone.Cols + 1, 1)
                Emgu.CV.CvInvoke.MatchTemplate(oSubMatrix, oCornerstone, oResultMap, Emgu.CV.CvEnum.TemplateMatchingType.CcoeffNormed)

                ' extract a mask containing the cornerstone blobs
                Using oMaskMatrix As New Emgu.CV.Matrix(Of Byte)(oResultMap.Size)
                    Emgu.CV.CvInvoke.Threshold(oResultMap, oMaskMatrix, TemplateMatchCutoff, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary)

                    ' count the number of blobs in the mask and take the largest iMaxCount blobs to be the location of the images
                    Using oBlobDetector As New Emgu.CV.Cvb.CvBlobDetector
                        Using oBlobs As New Emgu.CV.Cvb.CvBlobs
                            oBlobDetector.Detect(oMaskMatrix.Mat.ToImage(Of Emgu.CV.Structure.Gray, Byte), oBlobs)
                            Dim oBlobList As List(Of Emgu.CV.Cvb.CvBlob) = (From oBlob In oBlobs Order By oBlob.Value.Area Descending Take iMaxCount Select oBlob.Value).ToList

                            ' determine the weighted centroids of each blob using moments
                            Dim fMaxVal As Double
                            For Each oBlob As Emgu.CV.Cvb.CvBlob In oBlobList
                                Using oSubRect As Emgu.CV.Matrix(Of Single) = oResultMap.GetSubRect(New System.Drawing.Rectangle(oBlob.BoundingBox.X, oBlob.BoundingBox.Y, oBlob.BoundingBox.Width, oBlob.BoundingBox.Height))
                                    ' get max value
                                    oSubRect.MinMax(Nothing, fMaxVal, Nothing, Nothing)

                                    Dim oMoments As Emgu.CV.Structure.MCvMoments = Emgu.CV.CvInvoke.Moments(oSubRect)
                                    oReturnList.Add(New Tuple(Of Double, PointDouble)(fMaxVal, New PointDouble(iHorizontalDisplacement + (CSng(oCornerstone.Cols) / 2) + If(oMoments.M00 = 0, oBlob.Centroid.X, oBlob.BoundingBox.X + (oMoments.M10 / oMoments.M00)), iVerticalDisplacement + (CSng(oCornerstone.Rows) / 2) + If(oMoments.M00 = 0, oBlob.Centroid.Y, oBlob.BoundingBox.Y + (oMoments.M01 / oMoments.M00)))))
                                End Using
                            Next
                        End Using
                    End Using
                End Using
            End Using
        End Using
        Return oReturnList.OrderByDescending(Function(x) x.Item1).ToList
    End Function
    Private Shared Sub EliminateDuplicatePoints(ByVal oSourceMatrix As Emgu.CV.Matrix(Of Byte), ByVal oCornerstone As Emgu.CV.Matrix(Of Byte), ByRef oReturnList As List(Of Tuple(Of Double, PointDouble)), ByVal iRadius As Integer)
        ' eliminates duplicate points
        Dim iVerticalStep As Integer = (oCornerstone.Height + iRadius) * 2
        Dim iHorizontalStep As Integer = (oCornerstone.Width + iRadius) * 2

        If oReturnList.Count > 1 Then
            For i = oReturnList.Count - 1 To 1 Step -1
                For j = i - 1 To 0 Step -1
                    Dim oPoint1 As PointDouble = oReturnList(i).Item2
                    Dim oPoint2 As PointDouble = oReturnList(j).Item2

                    ' check if the points are within the radius
                    If oPoint1.X >= oPoint2.X - iRadius AndAlso oPoint1.X <= oPoint2.X + iRadius AndAlso oPoint1.Y >= oPoint2.Y - iRadius AndAlso oPoint1.Y <= oPoint2.Y + iRadius Then
                        Dim iVerticalDisplacement As Integer = Math.Max(Math.Min((oPoint1.Y + oPoint2.Y) / 2 - (iVerticalStep / 2), oSourceMatrix.Height - iVerticalStep), 0)
                        Dim iHorizontalDisplacement As Integer = Math.Max(Math.Min((oPoint1.X + oPoint2.X) / 2 - (iHorizontalStep / 2), oSourceMatrix.Width - iHorizontalStep), 0)

                        Dim oReturnPoints As List(Of Tuple(Of Double, PointDouble)) = LocatePoint(oSourceMatrix, oCornerstone, 1, iHorizontalDisplacement, iVerticalDisplacement, iHorizontalStep, iVerticalStep)
                        If oReturnPoints.Count > 0 Then
                            oReturnList.RemoveAt(i)
                            oReturnList(j) = oReturnPoints.First
                        End If
                        Exit For
                    End If
                Next
            Next
        End If

        oReturnList = oReturnList.OrderByDescending(Function(x) x.Item1).ToList
    End Sub
#End Region
#Region "Command"
    Private Class RelayCommand
        Implements ICommand

        Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
        Public Sub Execute(parameter As Object) Implements ICommand.Execute
            Using oUpdateSuspender As New UpdateSuspender(False)
                If Not UpdateSuspender.GlobalSuspendProcessing Then
                    Dim sParameter As String() = CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item2.Split("|")
                    Dim sender As FrameworkElement = CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item3
                    Dim oField As FieldDocumentStore.Field = CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item4
                    Dim oImages As New List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))

                    Dim oDataGrid As Controls.DataGrid = Root.DataGridScanner
                    Dim iCurrentIndex As Integer = oDataGrid.Items.IndexOf(oField)
                    If CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item1.GetType.Equals(GetType(KeyEventArgs)) Then
                        Dim e As KeyEventArgs = CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item1
                        Select Case oField.FieldType
                            Case Enumerations.FieldTypeEnum.BoxChoice
                                Dim oTextBox As Controls.TextBox = sender
                                Dim iColumn As Integer = Val(sParameter(0))
                                Dim iOrder As Integer = Val(sParameter(1))

                                If e.Key = Key.Left Or e.Key = Key.Right Then
                                    Do
                                        If e.Key = Key.Left AndAlso oTextBox.CaretIndex = 0 Then
                                            iColumn -= 1
                                        ElseIf e.Key = Key.Right AndAlso oTextBox.CaretIndex = oTextBox.Text.Length Then
                                            iColumn += 1
                                        End If

                                        oImages = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iColumn And oField.Images(iIndex).Item5 = -1 Select oField.Images(iIndex)).ToList
                                    Loop Until oImages.Count = 0 OrElse Not oImages.First.Item6

                                    If oImages.Count = 0 Then
                                        If iColumn < 0 Then
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, True)
                                        Else
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                        End If
                                    Else
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, 0, False)
                                    End If
                                ElseIf e.Key = Key.Up Then
                                    If iCurrentIndex > 0 Then
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, False)
                                    End If
                                ElseIf e.Key = Key.Down Then
                                    If iCurrentIndex < oDataGrid.Items.Count - 1 Then
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                    End If
                                End If
                            Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                                Dim oCheckBox As Controls.CheckBox = sender
                                Dim iColumn As Integer = Val(sParameter(0))
                                Dim iRow As Integer = Val(sParameter(1))
                                Dim iOrder As Integer = Val(sParameter(2))
                                If e.Key = Key.Left Or e.Key = Key.Right Then
                                    Dim iGroups As Integer = Math.Ceiling(oField.MarkCount / Scanner.MaxChoiceField)
                                    Dim iRowLength As Integer = Math.Ceiling(oField.MarkCount / iGroups)

                                    e.Handled = True
                                    If e.Key = Key.Left Then
                                        If iColumn = 0 Then
                                            If iRow = 0 Then
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, True)
                                            Else
                                                iColumn = iRowLength - 1
                                                iRow -= 1
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                            End If
                                        Else
                                            iColumn -= 1
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                        End If
                                    Else
                                        If (iRow * iRowLength) + iColumn = oField.Images.Count - 1 Then
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                        Else
                                            If iColumn = iRowLength - 1 Then
                                                iColumn = 0
                                                iRow += 1
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                            Else
                                                iColumn += 1
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                            End If
                                        End If
                                    End If
                                ElseIf e.Key = Key.Up Then
                                    e.Handled = True
                                    If iRow > 0 Then
                                        iRow -= 1
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                    Else
                                        If iCurrentIndex > 0 Then
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, False)
                                        End If
                                    End If
                                ElseIf e.Key = Key.Down Then
                                    e.Handled = True
                                    Dim iMaxRow As Integer = Math.Ceiling(oField.MarkCount / MaxChoiceField) - 1
                                    If iRow < iMaxRow Then
                                        iRow += 1
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                    Else
                                        If iCurrentIndex < oDataGrid.Items.Count - 1 Then
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                        End If
                                    End If
                                End If
                            Case Enumerations.FieldTypeEnum.Handwriting
                                Dim oTextBox As Controls.TextBox = sender
                                Dim iColumn As Integer = Val(sParameter(0))
                                Dim iRow As Integer = Val(sParameter(1))
                                Dim iOrder As Integer = Val(sParameter(2))

                                If e.Key = Key.Left Or e.Key = Key.Right Then
                                    Do
                                        If e.Key = Key.Left AndAlso oTextBox.CaretIndex = 0 Then
                                            iColumn -= 1
                                        ElseIf e.Key = Key.Right AndAlso oTextBox.CaretIndex = oTextBox.Text.Length Then
                                            iColumn += 1
                                        End If

                                        oImages = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iRow And oField.Images(iIndex).Item5 = iColumn Select oField.Images(iIndex)).ToList
                                    Loop Until oImages.Count = 0 OrElse Not oImages.First.Item6

                                    If oImages.Count = 0 Then
                                        If iColumn < 0 Then
                                            If iRow = 0 Then
                                                ' move to previous field
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, True)
                                            Else
                                                ' move focus to previous row, last column
                                                Dim iMaxColumn As Integer = Aggregate iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where Not oField.Images(iIndex).Item6 Into Max(oField.Images(iIndex).Item5)
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iMaxColumn, iRow - 1, False)
                                            End If
                                        Else
                                            Dim iMaxRow As Integer = Aggregate iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where Not oField.Images(iIndex).Item6 Into Max(oField.Images(iIndex).Item4)
                                            If iRow = iMaxRow Then
                                                ' move to next field
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                            Else
                                                ' move focus to next row, first column
                                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, 0, iRow + 1, False)
                                            End If
                                        End If
                                    Else
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                    End If
                                ElseIf e.Key = Key.Up Then
                                    If iRow > 0 Then
                                        iRow -= 1
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                    Else
                                        If iCurrentIndex > 0 Then
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, False)
                                        End If
                                    End If
                                ElseIf e.Key = Key.Down Then
                                    Dim iMaxRow As Integer = Aggregate oImage In oField.Images Into Max(oImage.Item4)
                                    If iRow < iMaxRow Then
                                        iRow += 1
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                    Else
                                        If iCurrentIndex < oDataGrid.Items.Count - 1 Then
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                        End If
                                    End If
                                End If
                            Case Enumerations.FieldTypeEnum.Free
                                Dim oTextBox As Controls.TextBox = sender
                                Dim iOrder As Integer = Val(sParameter(0))
                                If (e.Key = Key.Left AndAlso oTextBox.CaretIndex = 0) OrElse e.Key = Key.Up Then
                                    If iCurrentIndex > 0 Then
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex - 1, iOrder, 0, 0, False)
                                    End If
                                ElseIf (e.Key = Key.Right AndAlso oTextBox.CaretIndex = oTextBox.Text.Length) OrElse e.Key = Key.Down Then
                                    If iCurrentIndex < oDataGrid.Items.Count - 1 Then
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                    End If
                                End If
                        End Select
                    ElseIf CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item1.GetType.Equals(GetType(Controls.TextChangedEventArgs)) Then
                        Dim e As Controls.TextChangedEventArgs = CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item1
                        If e.Changes.Count > 0 AndAlso Not (e.Changes.First.AddedLength = 0 And e.Changes.First.RemovedLength = 1) Then
                            Select Case oField.FieldType
                                Case Enumerations.FieldTypeEnum.BoxChoice
                                    Dim oTextBox As Controls.TextBox = sender
                                    Dim iColumn As Integer = Val(sParameter(0))
                                    Dim iOrder As Integer = Val(sParameter(1))
                                    Do
                                        iColumn += 1
                                        oImages = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iColumn And oField.Images(iIndex).Item5 = -1 Select oField.Images(iIndex)).ToList
                                    Loop Until oImages.Count = 0 OrElse Not oImages.First.Item6

                                    If oImages.Count = 0 Then
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                    Else
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, 0, False)
                                    End If
                                Case Enumerations.FieldTypeEnum.Handwriting
                                    Dim oTextBox As Controls.TextBox = sender
                                    Dim iColumn As Integer = Val(sParameter(0))
                                    Dim iRow As Integer = Val(sParameter(1))
                                    Dim iOrder As Integer = Val(sParameter(2))
                                    Do
                                        iColumn += 1
                                        oImages = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item4 = iRow And oField.Images(iIndex).Item5 = iColumn Select oField.Images(iIndex)).ToList
                                    Loop Until oImages.Count = 0 OrElse Not oImages.First.Item6

                                    If oImages.Count = 0 Then
                                        Dim iMaxRow As Integer = Aggregate iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where Not oField.Images(iIndex).Item6 Into Max(oField.Images(iIndex).Item4)
                                        If iRow = iMaxRow Then
                                            ' move to next field
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex + 1, iOrder, 0, 0, False)
                                        Else
                                            ' move focus to next row, first column
                                            UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, 0, iRow + 1, False)
                                        End If
                                    Else
                                        UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                                    End If
                            End Select
                        End If
                    ElseIf CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item1.GetType.Equals(GetType(MouseButtonEventArgs)) Then
                        Select Case oField.FieldType
                            Case Enumerations.FieldTypeEnum.BoxChoice
                                Dim oTextBox As Controls.TextBox = sender
                                Dim iColumn As Integer = Val(sParameter(0))
                                Dim iOrder As Integer = Val(sParameter(1))
                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, 0, False)
                            Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                                Dim oCheckBox As Controls.CheckBox = sender
                                Dim iColumn As Integer = Val(sParameter(0))
                                Dim iRow As Integer = Val(sParameter(1))
                                Dim iOrder As Integer = Val(sParameter(2))
                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                            Case Enumerations.FieldTypeEnum.Handwriting
                                Dim oTextBox As Controls.TextBox = sender
                                Dim iColumn As Integer = Val(sParameter(0))
                                Dim iRow As Integer = Val(sParameter(1))
                                Dim iOrder As Integer = Val(sParameter(2))
                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, iColumn, iRow, False)
                            Case Enumerations.FieldTypeEnum.Free
                                Dim iOrder As Integer = Val(sParameter(0))
                                UpdateSuspender.MoveDataGridFocus(iCurrentIndex, iOrder, 0, 0, False)
                        End Select
                    End If
                End If
            End Using
        End Sub
        Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
            If parameter.GetType.Equals(GetType(Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field))) Then
                Dim oField As FieldDocumentStore.Field = CType(parameter, Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)).Item4
                Select Case oField.FieldType
                    Case Enumerations.FieldTypeEnum.BoxChoice, Enumerations.FieldTypeEnum.Handwriting
                        Return True
                    Case Enumerations.FieldTypeEnum.Free
                        Return True
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        Return True
                    Case Else
                        Return False
                End Select
            Else
                Return False
            End If
        End Function
    End Class
#End Region
#Region "Classes"
    Public Class UpdateSuspender
        ' suspends image
        Implements IDisposable

        Private Shared m_GlobalSuspendState As Boolean = False
        Private Shared m_GlobalProcessingState As Boolean = False
        Private Shared m_PreventReentrancy As Boolean = False
        Private m_PreviousGlobalSuspendState As Boolean = False
        Private m_PreviousGlobalProcessingState As Boolean = False

        Private Shared m_ShowImagesSet As Boolean = False
        Private Shared m_ShowImages_SelectedIndex As Integer = -1
        Private Shared m_ShowImages_ScannedChanged As Boolean = True
        Private Shared m_ShowImages_DetectedChanged As Boolean = True

        Private Shared m_MoveDataGridFocusSet As Boolean = False
        Private Shared m_MoveDataGridFocusSet_Index As Integer = -1
        Private Shared m_MoveDataGridFocusSet_Order As Integer = -1
        Private Shared m_MoveDataGridFocusSet_Column As Integer = -1
        Private Shared m_MoveDataGridFocusSet_Row As Integer = -1
        Private Shared m_MoveDataGridFocusSet_Last As Boolean = False
        Private Shared m_MoveDataGridFocusSet_Move As Boolean = False

        Private Shared m_UpdateScannerCollection As Boolean = False

        Sub New(ByVal bSuspendProcessing As Boolean)
            If Not m_PreventReentrancy Then
                m_PreviousGlobalSuspendState = GlobalSuspendUpdates
                GlobalSuspendUpdates = True

                m_PreviousGlobalProcessingState = GlobalSuspendProcessing
                If bSuspendProcessing Then
                    m_GlobalProcessingState = True
                End If
            End If
        End Sub
        Public Shared Sub ShowImages(ByVal iSelectedIndex As Integer, Optional ByVal bScannedChanged As Boolean = True, Optional ByVal bDetectedChanged As Boolean = True)
            m_ShowImagesSet = True
            m_ShowImages_SelectedIndex = iSelectedIndex
            m_ShowImages_ScannedChanged = bScannedChanged
            m_ShowImages_DetectedChanged = bDetectedChanged
        End Sub
        Public Shared Sub MoveDataGridFocus(ByVal iIndex As Integer, ByVal iOrder As Integer, ByVal iColumn As Integer, ByVal iRow As Integer, ByVal bLast As Boolean, Optional ByVal bMove As Boolean = True)
            m_MoveDataGridFocusSet = True
            m_MoveDataGridFocusSet_Index = iIndex
            m_MoveDataGridFocusSet_Order = iOrder
            m_MoveDataGridFocusSet_Column = iColumn
            m_MoveDataGridFocusSet_Row = iRow
            m_MoveDataGridFocusSet_Last = bLast
            m_MoveDataGridFocusSet_Move = bMove
        End Sub
        Public Shared Sub UpdateScannerCollection()
            m_UpdateScannerCollection = True
        End Sub
        Public Shared Property GlobalSuspendUpdates As Boolean
            Get
                Return m_GlobalSuspendState
            End Get
            Set(value As Boolean)
                m_GlobalSuspendState = value
                If (Not m_GlobalSuspendState) Then
                    ' update all pending content changes
                    Dim oAction As Action = Sub()
                                                If m_ShowImagesSet Then
                                                    Scanner.ShowImages(m_ShowImages_SelectedIndex, m_ShowImages_ScannedChanged, m_ShowImages_DetectedChanged)
                                                    m_ShowImagesSet = False
                                                End If

                                                ' update move focus changes
                                                If m_MoveDataGridFocusSet Then
                                                    Scanner.MoveDataGridFocus(m_MoveDataGridFocusSet_Index, m_MoveDataGridFocusSet_Order, m_MoveDataGridFocusSet_Column, m_MoveDataGridFocusSet_Row, m_MoveDataGridFocusSet_Last, m_MoveDataGridFocusSet_Move)
                                                    m_MoveDataGridFocusSet = False
                                                End If

                                                ' update scanner collection
                                                If m_UpdateScannerCollection Then
                                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                                    oScannerCollection.Update()
                                                    m_UpdateScannerCollection = False
                                                End If
                                            End Sub
                    CommonFunctions.DispatcherInvoke(UIDispatcher, oAction)
                End If
            End Set
        End Property
        Public Shared ReadOnly Property GlobalSuspendProcessing As Boolean
            Get
                Return m_GlobalProcessingState
            End Get
        End Property
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Shadows Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If Not m_PreventReentrancy Then
                        ' global suspend
                        GlobalSuspendUpdates = m_PreviousGlobalSuspendState

                        ' global processing
                        m_GlobalProcessingState = m_PreviousGlobalProcessingState
                    End If
                End If
            End If
            disposedValue = True
        End Sub
        Public Shadows Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
#End Region
#Region "Buttons"
    ' add processing for the plugin buttons here
    Private Sub LoadImages_Button_Click(sender As Object, e As RoutedEventArgs)
        ' loads images for processing
        Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
        oOpenFileDialog.FileName = String.Empty
        oOpenFileDialog.DefaultExt = "*"
        oOpenFileDialog.Multiselect = True
        oOpenFileDialog.Filter = "All Files|*.*|JPEG Images|*.jpg;*.jpeg|TIFF Images|*.tif;*.tiff|PNG Images|*.png"
        oOpenFileDialog.Title = "Load Images For Processing"
        Dim result? As Boolean = oOpenFileDialog.ShowDialog()
        If result = True Then
            LoadImagesFunction(oOpenFileDialog.FileNames)
        End If
    End Sub
    Private Sub ExportData_Button_Click(sender As Object, e As RoutedEventArgs)
        ' exports data to spreadsheet
        Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
        oSaveFileDialog.FileName = String.Empty
        oSaveFileDialog.DefaultExt = "*.xlsx"
        oSaveFileDialog.Filter = "Excel Spreadsheet|*.xlsx"
        oSaveFileDialog.Title = "Export Data To File"
        oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
        Dim result? As Boolean = oSaveFileDialog.ShowDialog()
        If result = True Then
            Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
            If oScannerCollection.FieldDocumentStore.FieldCollectionStore.Count > 0 Then
                ExportData(oSaveFileDialog.FileName)
            Else
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "No data. File not saved."))
            End If
        End If
    End Sub
    Private Sub Exit_Button_Click(sender As Object, e As RoutedEventArgs)
        RaiseEvent ExitButtonClick(Guid.Empty)
    End Sub
    Private Sub LoadLink_Button_Click(sender As Object, e As RoutedEventArgs)
        ' loads linked plugin
        Dim oButton As Controls.Button = sender
        Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
        Dim iButton As Integer = Val(Right(oButton.Name, Len(oButton.Name) - Len("Button")))

        RaiseEvent ExitButtonClick(m_Identifiers(oFilteredNames(iButton)).Item1)
    End Sub
    Private Sub Help_Button_Click(sender As Object, e As RoutedEventArgs)
        ' help button

    End Sub
#End Region
#Region "ButtonHelp"
    Public Const DragHelp As String = "Help"
    Public Shared Sub HelpMouseMoveHandler(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        ' allows drag and drop
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim oInputElement As IInputElement = Mouse.DirectlyOver
            If (Not IsNothing(oInputElement)) AndAlso oInputElement.GetType.IsSubclassOf(GetType(FrameworkElement)) Then
                Dim oFrameworkElement As FrameworkElement = oInputElement
                Dim oGUID As Guid = Guid.Empty

                Dim oDataGridScanner As Controls.DataGrid = CommonFunctions.GetParentObject(Of Controls.DataGrid)(oFrameworkElement, "DataGridScanner")
                If IsNothing(oDataGridScanner) Then
                    IterateElements(oFrameworkElement, oGUID)
                Else
                    Guid.TryParse(oDataGridScanner.Tag, oGUID)
                End If

                If Not oGUID.Equals(Guid.Empty) Then
                    ' drag and drop operation move
                    DragDrop.DoDragDrop(oFrameworkElement, New DataObject(DragHelp, oGUID), DragDropEffects.Move)
                End If
            End If
        End If
    End Sub
    Private Shared Sub IterateElements(ByVal oFrameworkElement As FrameworkElement, ByRef oGUID As Guid)
        ' iterates through the child elements and finds the first element to have a GUID tag
        ' process only if guid not initialised
        If oGUID.Equals(Guid.Empty) Then
            Dim oParseGUID As Guid = Guid.Empty
            If IsMouseOverElement(oFrameworkElement) AndAlso (Not IsNothing(oFrameworkElement.Tag)) AndAlso oFrameworkElement.Tag.GetType.Equals(GetType(String)) AndAlso Guid.TryParse(oFrameworkElement.Tag, oParseGUID) Then
                oGUID = oParseGUID
            Else
                ' iterate through children
                Dim oChildren As IEnumerable = LogicalTreeHelper.GetChildren(oFrameworkElement)
                For Each oChild In oChildren
                    Dim oChildFrameworkElement As FrameworkElement = TryCast(oChild, FrameworkElement)
                    If Not IsNothing(oChildFrameworkElement) Then
                        IterateElements(oChildFrameworkElement, oGUID)
                    End If
                Next
            End If
        End If
    End Sub
    Private Shared Function IsMouseOverElement(ByVal oFrameworkElement As FrameworkElement) As Boolean
        ' checks if mouse is over an element
        Dim oPoint As Point = Mouse.GetPosition(oFrameworkElement)
        If oPoint.X >= 0 AndAlso oPoint.Y >= 0 AndAlso oPoint.X < oFrameworkElement.ActualWidth AndAlso oPoint.Y < oFrameworkElement.ActualHeight Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub ButtonHelpMoveHandler(sender As Object, e As MouseEventArgs) Handles ButtonHelp.MouseMove
        If Not e.LeftButton = MouseButtonState.Pressed Then
            ButtonHelp.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
        End If
    End Sub
    Private Sub ButtonHelpLeaveHandler(sender As Object, e As EventArgs) Handles ButtonHelp.MouseLeave, ButtonHelp.TouchLeave
        ButtonHelp.Background = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub ButtonHelp_DragEnter(sender As Object, e As DragEventArgs) Handles ButtonHelp.DragEnter
        If e.Data.GetDataPresent(DragHelp) Then
            ButtonHelp.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
        Else
            ButtonHelp.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
        End If
    End Sub
    Private Sub ButtonHelp_DragLeave(sender As Object, e As DragEventArgs) Handles ButtonHelp.DragLeave
        ButtonHelp.Background = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub ButtonHelp_Drop(sender As Object, e As DragEventArgs) Handles ButtonHelp.Drop
        ' handles drop event
        Dim oDragGUID As Guid = Guid.Empty
        If e.Data.GetDataPresent(DragHelp) Then
            oDragGUID = e.Data.GetData(DragHelp)
        End If

        If Not oDragGUID.Equals(Guid.Empty) Then
            ' shows help dialog
            Dim oHelpDialog As New HelpDialog(GridMain.ActualWidth * 3 / 5, GridMain.ActualHeight * 22 / 23, GridButtons.ColumnDefinitions(0).ActualWidth / 2, GridMain.RowDefinitions(1).ActualHeight)
            oHelpDialog.GUID = oDragGUID
            oHelpDialog.ShowDialog()

            ButtonHelp.Background = New SolidColorBrush(Colors.Transparent)
        End If
    End Sub
#End Region
#Region "UI"
    Private Sub ScannerLoadHandler(sender As Object, e As EventArgs) Handles ScannerLoad.Click
        Dim oAction As Action = Sub()
                                    Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
                                    oOpenFileDialog.FileName = String.Empty
                                    oOpenFileDialog.DefaultExt = "*.gz"
                                    oOpenFileDialog.Multiselect = False
                                    oOpenFileDialog.Filter = "GZip Files|*.gz"
                                    oOpenFileDialog.Title = "Load Field Definitions From File"
                                    Dim result? As Boolean = oOpenFileDialog.ShowDialog()
                                    If result = True Then
                                        ' loads the scanner collection from a file
                                        Dim oNewFieldDocumentStore As FieldDocumentStore = CommonFunctions.DeserializeDataContractFile(Of FieldDocumentStore)(oOpenFileDialog.FileName)
                                        If Not IsNothing(oNewFieldDocumentStore) Then
                                            Using oUpdateSuspender As New UpdateSuspender(False)
                                                Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")

                                                Using oDeferRefresh As New DeferRefresh()
                                                    oScannerCollection.SetContent(oNewFieldDocumentStore)
                                                End Using

                                                ' select the first item
                                                DataGridScanner.InvalidateArrange()
                                                DataGridScanner.UpdateLayout()
                                                UpdateSuspender.MoveDataGridFocus(0, MarkState, 0, 0, False)
                                                UpdateSuspender.ShowImages(DataGridScanner.SelectedIndex)

                                                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Field definitions loaded from file."))

                                                SetScanIcons()
                                            End Using
                                        Else
                                            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Unable to load field definitions from file."))
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerClearHandler(sender As Object, e As EventArgs) Handles ScannerLoad.RightClick
        Dim oAction As Action = Sub()
                                    If MessageBox.Show("Clear field definitions?", ModuleName, MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                                        Using oUpdateSuspender As New UpdateSuspender(False)
                                            Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")

                                            Using oDeferRefresh As New DeferRefresh()
                                                oScannerCollection.ClearAll()
                                            End Using

                                            DataGridScanner.InvalidateArrange()
                                            DataGridScanner.UpdateLayout()

                                            SetPDFDocument()
                                            PDFViewerControl.CurrentPage = -1
                                            UpdateSuspender.ShowImages(-1)
                                            Dim oCanvasDetectedMarksContent As Controls.Canvas = Root.CanvasDetectedMarksContent
                                            oCanvasDetectedMarksContent.Children.Clear()

                                            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Field definitions cleared."))

                                            SetScanIcons()
                                        End Using
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerSaveHandler(sender As Object, e As EventArgs) Handles ScannerSave.Click
        Dim oAction As Action = Sub()
                                    ' saves data to file
                                    Const FieldDataName As String = "FieldData.gz"
                                    Dim oFolderBrowserDialog As New Forms.FolderBrowserDialog
                                    oFolderBrowserDialog.Description = "Save Data to File"
                                    oFolderBrowserDialog.ShowNewFolderButton = True
                                    oFolderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop
                                    If oFolderBrowserDialog.ShowDialog = Forms.DialogResult.OK Then
                                        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                        If oScannerCollection.FieldDocumentStore.FieldCollectionStore.Count > 0 Then
                                            Dim sFileName As String = oFolderBrowserDialog.SelectedPath + "\" + FieldDataName
                                            If IO.File.Exists(sFileName) Then
                                                IO.File.Delete(sFileName)
                                            End If

                                            CommonFunctions.SerializeDataContractFile(sFileName, oScannerCollection.FieldDocumentStore, CommonFunctions.GetKnownTypes(LocalKnownTypes), False)

                                            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Data saved to '" + FieldDataName + "'."))
                                        Else
                                            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "No data. File not saved."))
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerScanHandler(sender As Object, e As EventArgs) Handles ScannerScan.Click
        Dim oAction As Action = Sub()
                                    If SetRecognisers() Then
                                        ScanPageProgress = 0
                                        Select Case SelectedScannerSource.Item1
                                            Case Twain32Enumerations.ScannerSource.TWAIN
                                                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Scan started."))
                                                m_CommonScanner.ScannerScan()
                                            Case Twain32Enumerations.ScannerSource.WIA
                                                ' configure
                                                Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                                If IsNothing(oScannerCollection) Then
                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Field definitions not loaded."))
                                                Else
                                                    Dim oWIADeviceManager As New WIA.DeviceManager
                                                    For Each oDeviceInfo As WIA.DeviceInfo In oWIADeviceManager.DeviceInfos
                                                        If oDeviceInfo.DeviceID = SelectedScannerSource.Item2 Then
                                                            Dim oDevice As WIA.Device = oDeviceInfo.Connect()
                                                            If IsNothing(oDevice) Then
                                                                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Unable to connect scanner."))
                                                            Else
                                                                Dim oCommonDialog As New WIA.CommonDialog
                                                                Dim oImageFile As WIA.ImageFile = oCommonDialog.ShowAcquireImage(WIA.WiaDeviceType.ScannerDeviceType, WIA.WiaImageIntent.GrayscaleIntent, WIA.WiaImageBias.MaximizeQuality, WIA.FormatID.wiaFormatTIFF, False, False)
                                                                If IsNothing(oImageFile) Then
                                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Scan error."))
                                                                Else
                                                                    Using oImageFileMemoryStream As New IO.MemoryStream(CType(oImageFile.FileData.BinaryData, Byte()))
                                                                        Using oBitmap As System.Drawing.Bitmap = System.Drawing.Image.FromStream(oImageFileMemoryStream)
                                                                            ReturnScannedImage(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.Image, Converter.BitmapConvertGrayscale(oBitmap)))
                                                                            ReturnScannedImage(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.Complete, Nothing))
                                                                        End Using
                                                                    End Using
                                                                End If
                                                            End If
                                                            Exit For
                                                        End If
                                                    Next
                                                End If
                                        End Select
                                    Else
                                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Unable to load recognisers."))
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerChooseHandler(sender As Object, e As EventArgs) Handles ScannerChoose.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerDialog As New ScannerDialog
                                    oScannerDialog.Width = 600
                                    oScannerDialog.Height = 600
                                    oScannerDialog.ScannerSources.Clear()

                                    Dim oScannerSources As List(Of Tuple(Of Twain32Enumerations.ScannerSource, String, String)) = GetScannerSources()

                                    For Each oScanner In oScannerSources
                                        oScannerDialog.ScannerSources.Add(oScanner.Item3 + " (" + [Enum].GetName(GetType(Twain32Enumerations.ScannerSource), oScanner.Item1) + ")")
                                    Next

                                    If oScannerDialog.ShowDialog Then
                                        If oScannerSources.Count > 0 AndAlso (oScannerDialog.SelectedIndex >= 0 And oScannerDialog.SelectedIndex <= oScannerSources.Count - 1) Then
                                            If oScannerDialog.SelectedIndex = -1 Then
                                                SelectedScannerSource = Nothing
                                                ScannerScan.InnerToolTip = String.Empty
                                            Else
                                                SelectedScannerSource = oScannerSources(oScannerDialog.SelectedIndex)
                                                ScannerScan.InnerToolTip = SelectedScannerSource.Item3 + " (" + [Enum].GetName(GetType(Twain32Enumerations.ScannerSource), SelectedScannerSource.Item1) + ")"
                                            End If
                                        End If
                                    End If

                                    SetScanIcons()

                                    If IsNothing(SelectedScannerSource) Then
                                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Scanner not chosen."))
                                    Else
                                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Scanner " + SelectedScannerSource.Item3 + " chosen."))
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerConfigureHandler(sender As Object, e As EventArgs) Handles ScannerConfigure.Click
        Dim oAction As Action = Sub()
                                    If IsNothing(SelectedScannerSource) Then
                                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Scanner not selected."))
                                    Else
                                        Select Case SelectedScannerSource.Item1
                                            Case Twain32Enumerations.ScannerSource.TWAIN
                                                m_CommonScanner.ScannerConfigure()
                                            Case Twain32Enumerations.ScannerSource.WIA
                                                Dim bScannerFound As Boolean = False
                                                Dim oWIADeviceManager As New WIA.DeviceManager
                                                For Each oDeviceInfo As WIA.DeviceInfo In oWIADeviceManager.DeviceInfos
                                                    If oDeviceInfo.DeviceID = SelectedScannerSource.Item2 Then
                                                        Dim oDevice As WIA.Device = oDeviceInfo.Connect()
                                                        Dim oCommonDialog As New WIA.CommonDialog
                                                        oCommonDialog.ShowItemProperties(oDevice.Items(1))
                                                        bScannerFound = True
                                                        Exit For
                                                    End If
                                                Next

                                                If Not bScannerFound Then
                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Scanner " + SelectedScannerSource.Item3 + " not found."))
                                                End If
                                        End Select
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerViewHandler(sender As Object, e As EventArgs) Handles ScannerView.Click
        Dim oAction As Action = Sub()
                                    ViewerState = Not ViewerState
                                    SetScanIcons()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerFilterHandler(sender As Object, e As EventArgs) Handles ScannerFilter.Click
        Dim oAction As Action = Sub()
                                    Select Case FilterState
                                        Case Enumerations.FilterData.None
                                            FilterState = Enumerations.FilterData.DataMissing
                                        Case Enumerations.FilterData.DataMissing
                                            FilterState = Enumerations.FilterData.DataPresent
                                        Case Enumerations.FilterData.DataPresent
                                            FilterState = Enumerations.FilterData.None
                                    End Select
                                    SetScanIcons()

                                    Dim oDataGridScanner As Controls.DataGrid = Root.DataGridScanner
                                    oDataGridScanner.CommitEdit()
                                    oDataGridScanner.CommitEdit()

                                    Dim oCollectionViewSource As Data.CollectionViewSource = Root.GridMain.Resources("cvsScannnerCollection")
                                    oCollectionViewSource.View.Refresh()

                                    If Not IsNothing(oDataGridScanner.SelectedItem) Then
                                        oDataGridScanner.ScrollIntoView(oDataGridScanner.SelectedItem)
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerMarkHandler(sender As Object, e As EventArgs) Handles ScannerMark.Click
        Dim oAction As Action = Sub()
                                    MarkState = (MarkState + 1) Mod 3
                                    SetScanIcons()

                                    Dim oDataGridScanner As Controls.DataGrid = Root.DataGridScanner
                                    oDataGridScanner.CommitEdit()
                                    oDataGridScanner.CommitEdit()

                                    Dim oCollectionViewSource As Data.CollectionViewSource = Root.GridMain.Resources("cvsScannnerCollection")
                                    oCollectionViewSource.View.Refresh()

                                    DataGridScannerChanged(False, True)
                                    If Not IsNothing(oDataGridScanner.SelectedItem) Then
                                        oDataGridScanner.ScrollIntoView(oDataGridScanner.SelectedItem)
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerPreviousSubjectHandler(sender As Object, e As EventArgs) Handles ScannerPreviousSubject.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                    oScannerCollection.PreviousItem()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerNextSubjectHandler(sender As Object, e As EventArgs) Handles ScannerNextSubject.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                    oScannerCollection.NextItem()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerPreviousDataHandler(sender As Object, e As EventArgs) Handles ScannerPreviousData.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                    oScannerCollection.PreviousItemData()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerNextDataHandler(sender As Object, e As EventArgs) Handles ScannerNextData.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                    oScannerCollection.NextItemData()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerProcessedHandler(sender As Object, e As EventArgs) Handles ScannerProcessed.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                    oScannerCollection.Processed = Not oScannerCollection.Processed
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ScannerHideProcessedHandler(sender As Object, e As EventArgs) Handles ScannerHideProcessed.Click
        Dim oAction As Action = Sub()
                                    Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")
                                    oScannerCollection.HideProcessed = Not oScannerCollection.HideProcessed
                                    If oScannerCollection.HideProcessed Then
                                        ScannerHideProcessed.HBSource = GetIcon("CCMNotVisible")
                                        ScannerHideProcessed.InnerToolTip = "Processed Subjects Hidden"
                                    Else
                                        ScannerHideProcessed.HBSource = GetIcon("CCMVisible")
                                        ScannerHideProcessed.InnerToolTip = "Processed Subjects Visible"
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub CanvasDetectedMarksContent_MouseDownHandler(sender As Object, e As MouseButtonEventArgs) Handles CanvasDetectedMarksContent.MouseLeftButtonDown, CanvasDetectedMarksContent.MouseRightButtonDown
        ' changes the mark for the choice field
        Using oUpdateSuspender As New UpdateSuspender(False)
            If Not UpdateSuspender.GlobalSuspendProcessing Then
                Dim oPosition As Point = e.GetPosition(ImageDetectedMarksContent)
                Dim oAdjustedPosition As New Point(oPosition.X / ImageDetectedMarksContent.ActualWidth, oPosition.Y / ImageDetectedMarksContent.ActualHeight)
                Dim oSelectedRectangles As List(Of Tuple(Of Guid, Integer, Integer)) = (From oGUID As Guid In ChoiceRectangles.Keys From iIndex As Integer In Enumerable.Range(0, ChoiceRectangles(oGUID).Count) Where oAdjustedPosition.X >= ChoiceRectangles(oGUID)(iIndex).Item1.Left And oAdjustedPosition.X <= ChoiceRectangles(oGUID)(iIndex).Item1.Right And oAdjustedPosition.Y >= ChoiceRectangles(oGUID)(iIndex).Item1.Top And oAdjustedPosition.Y <= ChoiceRectangles(oGUID)(iIndex).Item1.Bottom Select New Tuple(Of Guid, Integer, Integer)(oGUID, ChoiceRectangles(oGUID)(iIndex).Item2, ChoiceRectangles(oGUID)(iIndex).Item3)).ToList
                If oSelectedRectangles.Count > 0 Then
                    Dim oFieldList As List(Of FieldDocumentStore.Field) = (From oField As FieldDocumentStore.Field In DataGridScanner.Items Where oField.GUID.Equals(oSelectedRectangles.First.Item1) Select oField).ToList
                    If oFieldList.Count > 0 Then
                        Select Case oFieldList.First.FieldType
                            Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                                Select Case MarkState
                                    Case 0
                                        oFieldList.First.MarkChoice0(oSelectedRectangles.First.Item3) = Not oFieldList.First.MarkChoice0(oSelectedRectangles.First.Item3)
                                    Case 1
                                        oFieldList.First.MarkChoice1(oSelectedRectangles.First.Item3) = Not oFieldList.First.MarkChoice1(oSelectedRectangles.First.Item3)
                                    Case 2
                                        oFieldList.First.MarkChoice2(oSelectedRectangles.First.Item3) = Not oFieldList.First.MarkChoice2(oSelectedRectangles.First.Item3)
                                End Select
                            Case Enumerations.FieldTypeEnum.BoxChoice
                                If e.ChangedButton = MouseButton.Left Then
                                    oFieldList.First.MarkBoxChoiceRow(oSelectedRectangles.First.Item2, 0) = oSelectedRectangles.First.Item3.ToString
                                ElseIf e.ChangedButton = MouseButton.Right Then
                                    oFieldList.First.MarkBoxChoiceRow(oSelectedRectangles.First.Item2, 0) = String.Empty
                                End If
                        End Select
                    End If
                End If
            End If
        End Using
    End Sub
#End Region
End Class
#Region "Fields"
Public Class ScannerCollection
    Inherits TrueObservableCollection(Of FieldDocumentStore.Field)
    Implements INotifyPropertyChanged

    Private m_FieldDocumentStore As New FieldDocumentStore
    Private m_SelectedCollection As Integer = -1
    Private m_HideProcessed As Boolean = False
    Private Shared m_PreventReentrancy As Boolean = False

#Region "INotifyPropertyChanged"
    Public Shadows Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private Sub OnPropertyChangedLocal(ByVal sName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
    End Sub
#End Region
    Public Property FieldDocumentStore As FieldDocumentStore
        Get
            Return m_FieldDocumentStore
        End Get
        Set(value As FieldDocumentStore)
            m_FieldDocumentStore = value
        End Set
    End Property
    Public Property SelectedCollection As Integer
        Get
            Return m_SelectedCollection
        End Get
        Set(value As Integer)
            Using oUpdateSuspender As New Scanner.UpdateSuspender(True)
                Using oDeferRefresh As New Scanner.DeferRefresh()
                    Dim iOldSelectedCollection As Integer = m_SelectedCollection
                    If FieldDocumentStore.FieldCollectionStore.Count = 0 Then
                        m_SelectedCollection = -1
                    ElseIf value < 0 Or value > FieldDocumentStore.FieldCollectionStore.Count - 1 Then
                        m_SelectedCollection = 0
                    Else
                        m_SelectedCollection = value
                    End If

                    ' update old collection and replace with new one
                    If iOldSelectedCollection <> m_SelectedCollection Then
                        ' update old collection
                        If iOldSelectedCollection >= 0 And iOldSelectedCollection < FieldDocumentStore.FieldCollectionStore.Count Then
                            FieldDocumentStore.FieldCollectionStore(iOldSelectedCollection).Fields.Reset(Me)
                        End If

                        ' update new collection
                        If m_SelectedCollection >= 0 And m_SelectedCollection < FieldDocumentStore.FieldCollectionStore.Count Then
                            Reset(FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Fields)
                        Else
                            Clear()
                        End If

                        Update()
                    End If
                End Using
            End Using
        End Set
    End Property
    Public Property SelectedSubjectText As String
        Get
            If FieldDocumentStore.FieldCollectionStore.Count > 0 AndAlso m_SelectedCollection >= 0 Then
                Return (m_SelectedCollection + 1).ToString
            Else
                Return String.Empty
            End If
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property SubjectCountText As String
        Get
            If FieldDocumentStore.FieldCollectionStore.Count > 0 AndAlso m_SelectedCollection >= 0 Then
                Return FieldDocumentStore.FieldCollectionStore.Count.ToString
            Else
                Return String.Empty
            End If
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property SelectedItem As Common.HighlightComboBox.HCBDisplay
        Get
            If FieldDocumentStore.FieldCollectionStore.Count > 0 AndAlso m_SelectedCollection >= 0 Then
                Return New Common.HighlightComboBox.HCBDisplay(FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).HCBDisplay.Name, Guid.NewGuid, FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).HCBDisplay.Highlight)
            Else
                Return New Common.HighlightComboBox.HCBDisplay
            End If
        End Get
        Set(value As Common.HighlightComboBox.HCBDisplay)
        End Set
    End Property
    Public Property Processed As Boolean
        Get
            If FieldDocumentStore.FieldCollectionStore.Count > 0 AndAlso m_SelectedCollection >= 0 Then
                Return FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Processed
            Else
                Return False
            End If
        End Get
        Set(value As Boolean)
            If FieldDocumentStore.FieldCollectionStore.Count > 0 AndAlso m_SelectedCollection >= 0 Then
                FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Processed = Not FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Processed
                OnPropertyChangedLocal("Processed")
            End If
        End Set
    End Property
    Public Sub Finalise()
        ' stores the current fields prior to saving
        If m_SelectedCollection >= 0 And m_SelectedCollection < FieldDocumentStore.FieldCollectionStore.Count - 1 Then
            FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Fields.Clear()
            For Each oField In Me
                FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Fields.Add(oField)
            Next
        End If
    End Sub
    Public Sub ClearAll()
        ' clears store
        FieldDocumentStore.FieldCollectionStore.Clear()
        FieldDocumentStore.PDFTemplate = Nothing
        FieldDocumentStore.FieldMatrixStore.ClearAll()

        ' force update
        m_SelectedCollection = -2
        SelectedCollection = -1
    End Sub
    Public Sub SetAll(ByVal oScannerCollection As ScannerCollection)
        ' clears store
        FieldDocumentStore = oScannerCollection.FieldDocumentStore

        ' force update
        m_SelectedCollection = -2
        SelectedCollection = -1
    End Sub
    Public Sub PreviousItem()
        If FieldDocumentStore.FieldCollectionStore.Count > 0 Then
            Dim iNewIndex As Integer = -1
            For i = m_SelectedCollection - 1 To 0 Step -1
                If (Not FieldDocumentStore.FieldCollectionStore(i).Processed) Or (Not m_HideProcessed) Then
                    iNewIndex = i
                    Exit For
                End If
            Next
            If iNewIndex <> -1 Then
                SelectedCollection = iNewIndex
            End If
        End If
    End Sub
    Public Sub NextItem()
        If FieldDocumentStore.FieldCollectionStore.Count > 0 Then
            Dim iNewIndex As Integer = -1
            For i = m_SelectedCollection + 1 To FieldDocumentStore.FieldCollectionStore.Count - 1
                If (Not FieldDocumentStore.FieldCollectionStore(i).Processed) Or (Not m_HideProcessed) Then
                    iNewIndex = i
                    Exit For
                End If
            Next
            If iNewIndex <> -1 Then
                SelectedCollection = iNewIndex
            End If
        End If
    End Sub
    Public Sub PreviousItemData()
        If FieldDocumentStore.FieldCollectionStore.Count > 0 Then
            Dim iNewIndex As Integer = -1
            For i = m_SelectedCollection - 1 To 0 Step -1
                If (Not FieldDocumentStore.FieldCollectionStore(i).Processed) Or (Not m_HideProcessed) Then
                    Dim iCount As Integer = Aggregate oField In FieldDocumentStore.FieldCollectionStore(i).Fields Where oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataPartial Into Count
                    If iCount > 0 Then
                        iNewIndex = i
                        Exit For
                    End If
                End If
            Next
            If iNewIndex <> -1 Then
                SelectedCollection = iNewIndex
            End If
        End If
    End Sub
    Public Sub NextItemData()
        If FieldDocumentStore.FieldCollectionStore.Count > 0 Then
            Dim iNewIndex As Integer = -1
            For i = m_SelectedCollection + 1 To FieldDocumentStore.FieldCollectionStore.Count - 1
                If (Not FieldDocumentStore.FieldCollectionStore(i).Processed) Or (Not m_HideProcessed) Then
                    Dim iCount As Integer = Aggregate oField In FieldDocumentStore.FieldCollectionStore(i).Fields Where oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataPartial Into Count
                    If iCount > 0 Then
                        iNewIndex = i
                        Exit For
                    End If
                End If
            Next
            If iNewIndex <> -1 Then
                SelectedCollection = iNewIndex
            End If
        End If
    End Sub
    Public Property HideProcessed As Boolean
        Get
            Return m_HideProcessed
        End Get
        Set(value As Boolean)
            m_HideProcessed = value

            ' currently selected item is hidden
            If m_HideProcessed AndAlso m_SelectedCollection > -1 AndAlso FieldDocumentStore.FieldCollectionStore(m_SelectedCollection).Processed Then
                ' go to next higher index
                Dim iNewIndex As Integer = -1
                For i = m_SelectedCollection + 1 To FieldDocumentStore.FieldCollectionStore.Count - 1
                    If (Not FieldDocumentStore.FieldCollectionStore(i).Processed) Or (Not m_HideProcessed) Then
                        iNewIndex = i
                        Exit For
                    End If
                Next

                ' go to next lower index
                For i = m_SelectedCollection - 1 To 0 Step -1
                    If (Not FieldDocumentStore.FieldCollectionStore(i).Processed) Or (Not m_HideProcessed) Then
                        Dim iCount As Integer = Aggregate oField In FieldDocumentStore.FieldCollectionStore(i).Fields Where oField.DataPresent = FieldDocumentStore.Field.DataPresentEnum.DataPartial Into Count
                        If iCount > 0 Then
                            iNewIndex = i
                            Exit For
                        End If
                    End If
                Next

                If iNewIndex = -1 Then
                    SelectedCollection = -1
                Else
                    SelectedCollection = iNewIndex
                End If
            End If

            OnPropertyChangedLocal("HideProcessed")
        End Set
    End Property
    Public Sub SetContent(ByVal oFieldDocumentStore As FieldDocumentStore)
        ' sets store
        FieldDocumentStore.FieldCollectionStore.Clear()
        FieldDocumentStore.FieldCollectionStore.AddRange(oFieldDocumentStore.FieldCollectionStore)
        FieldDocumentStore.PDFTemplate = If(IsNothing(oFieldDocumentStore.PDFTemplate), Nothing, oFieldDocumentStore.PDFTemplate.Clone)
        oFieldDocumentStore.FieldMatrixStore.Transfer(FieldDocumentStore.FieldMatrixStore)

        ' force update
        m_SelectedCollection = -2
        If FieldDocumentStore.FieldCollectionStore.Count > 0 Then
            SelectedCollection = 0
        Else
            SelectedCollection = -1
        End If
    End Sub
    Public Sub Update()
        If Scanner.UpdateSuspender.GlobalSuspendUpdates Then
            Scanner.UpdateSuspender.UpdateScannerCollection()
        Else
            If Not m_PreventReentrancy Then
                m_PreventReentrancy = True
                Scanner.DataGridScannerChanged(False, True)
                OnPropertyChangedLocal("SelectedCollection")
                OnPropertyChangedLocal("SelectedSubjectText")
                OnPropertyChangedLocal("SubjectCountText")
                OnPropertyChangedLocal("SelectedItem")
                OnPropertyChangedLocal("FieldDocumentStore.FieldContent")
                OnPropertyChangedLocal("Processed")
                m_PreventReentrancy = False
            End If
        End If
    End Sub
    Private Shadows Sub Changed(sender As Object, e As NotifyCollectionChangedEventArgs) Handles MyBase.CollectionChanged
        If Not IsNothing(e.OldItems) Then
            For Each oField As FieldDocumentStore.Field In e.OldItems
                RemoveHandler oField.UpdateEvent, AddressOf Update
            Next
        End If
        If Not IsNothing(e.NewItems) Then
            For Each oField As FieldDocumentStore.Field In e.NewItems
                AddHandler oField.UpdateEvent, AddressOf Update
            Next
        End If
    End Sub
End Class
Public Class FieldTemplateSelector
    Inherits Controls.DataTemplateSelector

    Private m_ChoiceTemplate As String
    Private m_BoxChoiceTemplate As String
    Private m_HandwritingTemplate As String
    Private m_FreeTemplate As String

    Public WriteOnly Property ChoiceTemplate() As DataTemplate
        Set
            m_ChoiceTemplate = Markup.XamlWriter.Save(Value)
        End Set
    End Property
    Public WriteOnly Property BoxChoiceTemplate() As DataTemplate
        Set
            m_BoxChoiceTemplate = Markup.XamlWriter.Save(Value)
        End Set
    End Property
    Public WriteOnly Property HandwritingTemplate() As DataTemplate
        Set
            m_HandwritingTemplate = Markup.XamlWriter.Save(Value)
        End Set
    End Property
    Public WriteOnly Property FreeTemplate() As DataTemplate
        Set
            m_FreeTemplate = Markup.XamlWriter.Save(Value)
        End Set
    End Property
    Public ReadOnly Property ChoiceTemplateText As String
        Get
            Return m_ChoiceTemplate
        End Get
    End Property
    Public ReadOnly Property BoxChoiceTemplateText As String
        Get
            Return m_BoxChoiceTemplate
        End Get
    End Property
    Public ReadOnly Property HandwritingTemplateText As String
        Get
            Return m_HandwritingTemplate
        End Get
    End Property
    Public ReadOnly Property FreeTemplateText As String
        Get
            Return m_FreeTemplate
        End Get
    End Property
    Public Overrides Function SelectTemplate(item As Object, container As DependencyObject) As DataTemplate
        Const sColumnText As String = "<ColumnDefinition Width=""*"" />"
        Dim sFixedColumnText As String = "<ColumnDefinition Width=""XX"" />"
        Const sDVFText As String = "<Label Content=""D"" Margin=""0"" Padding=""0"" Grid.Row=""0"" Grid.Column=""0"" HorizontalAlignment=""Center"" VerticalContentAlignment=""Center""/><Label Content=""V"" Margin=""0"" Padding=""0"" Grid.Row=""1"" Grid.Column=""0"" HorizontalAlignment=""Center"" VerticalContentAlignment=""Center""/><Label Content=""F"" Margin=""0"" Padding=""0"" Grid.Row=""2"" Grid.Column=""0"" HorizontalAlignment=""Center"" VerticalContentAlignment=""Center""/>"
        Const sRowText As String = "<RowDefinition Height=""*"" />"
        Const sCheckboxText As String = "<CheckBox IsChecked=""False"" VerticalContentAlignment=""Center"" Margin=""0,0,0,0"" HorizontalAlignment=""Center"" Grid.Column=""0"" Grid.Row=""0"">X</CheckBox>"
        Const sTextboxText As String = "<TextBox HorizontalContentAlignment=""Center"" Margin=""0,0,0,0"" Grid.Column=""0"" Grid.Row=""0"">X</TextBox>"

        Dim oField As FieldDocumentStore.Field = TryCast(item, FieldDocumentStore.Field)

        ' custom logic to select appropriate data template and return
        If IsNothing(oField) Then
            Return MyBase.SelectTemplate(item, container)
        Else
            Dim iFixedWidth As Integer = Scanner.Root.DataGridScanner.ActualWidth / 40
            sFixedColumnText = sFixedColumnText.Replace("XX", iFixedWidth.ToString)
            Select Case oField.FieldType
                Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    Dim sChoiceTemplateText As String = ChoiceTemplateText

                    Dim iGroups As Integer = Math.Ceiling(oField.MarkCount / Scanner.MaxChoiceField)
                    Dim iRowLength As Integer = Math.Ceiling(oField.MarkCount / iGroups)

                    ' set column definition
                    Dim sCombinedColumnText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        sCombinedColumnText += sFixedColumnText
                    End If
                    For i = 0 To iRowLength - 1
                        sCombinedColumnText += sColumnText
                    Next
                    sChoiceTemplateText = sChoiceTemplateText.Replace(sColumnText, sCombinedColumnText)

                    ' set row definition
                    Dim iRows As Integer = If(Scanner.MarkState = 2, iGroups * 3, iGroups)
                    Dim sCombinedRowText As String = String.Empty
                    For i = 1 To iRows
                        sCombinedRowText += sRowText
                    Next
                    sChoiceTemplateText = sChoiceTemplateText.Replace(sRowText, sCombinedRowText)

                    ' set checkbox
                    Dim sCombinedCheckboxText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        Dim sNewDVFText As String = sDVFText
                        sNewDVFText = sNewDVFText.Replace("Grid.Row=""2""", "Grid.Row=""" + (2 * iGroups).ToString + """")
                        sNewDVFText = sNewDVFText.Replace("Grid.Row=""1""", "Grid.Row=""" + iGroups.ToString + """")
                        sCombinedCheckboxText += sNewDVFText
                    End If
                    For i = 0 To oField.MarkCount - 1
                        Dim iCol As Integer = -1
                        Dim iRow As Integer = -1
                        If Scanner.MarkState = 2 Then
                            For j = 0 To 2
                                Dim sCurrentCheckboxText As String = sCheckboxText
                                iCol = i Mod iRowLength
                                iRow = Math.Floor(i / iRowLength)
                                sCurrentCheckboxText = sCurrentCheckboxText.Replace("Grid.Column=""0""", "Grid.Column=""" + (iCol + 1).ToString + """")
                                sCurrentCheckboxText = sCurrentCheckboxText.Replace("Grid.Row=""0""", "Grid.Row=""" + (iRow + (j * iGroups)).ToString + """")
                                If j = 2 Then
                                    sCurrentCheckboxText = sCurrentCheckboxText.Replace("IsChecked=""False""", "IsChecked=""{Binding Path=MarkChoice" + j.ToString + "[" + i.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}""")
                                Else
                                    sCurrentCheckboxText = sCurrentCheckboxText.Replace("IsChecked=""False""", "IsEnabled=""False"" IsChecked=""{Binding Path=MarkChoice" + j.ToString + "[" + i.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}""")
                                End If

                                Dim sTabletText As String = String.Empty
                                Select Case oField.TabletContent
                                    Case Enumerations.TabletContentEnum.Letter
                                        sTabletText = Converter.ConvertNumberToLetter(oField.TabletStart + i, True)
                                    Case Enumerations.TabletContentEnum.Number
                                        sTabletText = (oField.TabletStart + i + 1).ToString
                                End Select
                                sCurrentCheckboxText = sCurrentCheckboxText.Replace(">X</CheckBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + j.ToString + "_" + iCol.ToString + "_" + iRow.ToString + """ Content=""" + sTabletText + """ ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></CheckBox>")
                                sCombinedCheckboxText += sCurrentCheckboxText
                            Next
                        Else
                            Dim sCurrentCheckboxText As String = sCheckboxText
                            iCol = i Mod iRowLength
                            iRow = Math.Floor(i / iRowLength)
                            sCurrentCheckboxText = sCurrentCheckboxText.Replace("Grid.Column=""0""", "Grid.Column=""" + iCol.ToString + """")
                            sCurrentCheckboxText = sCurrentCheckboxText.Replace("Grid.Row=""0""", "Grid.Row=""" + iRow.ToString + """")
                            sCurrentCheckboxText = sCurrentCheckboxText.Replace("IsChecked=""False""", "IsChecked=""{Binding Path=MarkChoice" + Scanner.MarkState.ToString + "[" + i.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}""")

                            Dim sTabletText As String = String.Empty
                            Select Case oField.TabletContent
                                Case Enumerations.TabletContentEnum.Letter
                                    sTabletText = Converter.ConvertNumberToLetter(oField.TabletStart + i, True)
                                Case Enumerations.TabletContentEnum.Number
                                    sTabletText = (oField.TabletStart + i + 1).ToString
                            End Select
                            sCurrentCheckboxText = sCurrentCheckboxText.Replace(">X</CheckBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + Scanner.MarkState.ToString + "_" + iCol.ToString + "_" + iRow.ToString + """ Content=""" + sTabletText + """ ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></CheckBox>")
                            sCombinedCheckboxText += sCurrentCheckboxText
                        End If
                    Next
                    sChoiceTemplateText = sChoiceTemplateText.Replace(sCheckboxText, sCombinedCheckboxText)
                    sChoiceTemplateText = sChoiceTemplateText.Replace("<DataTemplate ", "<DataTemplate xmlns:i=""clr-namespace:System.Windows.Interactivity;assembly=System.Windows.Interactivity"" xmlns:local=""clr-namespace:Scanner;assembly=Scanner"" ")

                    Dim oChoiceTemplate As DataTemplate = Nothing
                    Using oMemoryStream As New IO.MemoryStream(Text.Encoding.ASCII.GetBytes(sChoiceTemplateText))
                        Dim oParserContext = New Markup.ParserContext()
                        oChoiceTemplate = Markup.XamlReader.Load(oMemoryStream, oParserContext)
                    End Using

                    Return oChoiceTemplate
                Case Enumerations.FieldTypeEnum.BoxChoice
                    Dim sBoxChoiceTemplateText As String = BoxChoiceTemplateText
                    Dim oBoxImageList As Dictionary(Of Integer, Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))) = (From iIndex As Integer In Enumerable.Range(0, oField.Images.Count) Where oField.Images(iIndex).Item5 = -1 Select New KeyValuePair(Of Integer, Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))(iIndex, oField.Images(iIndex))).ToDictionary(Function(x) x.Key, Function(x) x.Value)

                    ' set column definition
                    Dim sCombinedColumnText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        sCombinedColumnText += sFixedColumnText
                    End If
                    For i = 0 To oField.MarkCount - 1
                        sCombinedColumnText += sColumnText
                    Next
                    sBoxChoiceTemplateText = sBoxChoiceTemplateText.Replace(sColumnText, sCombinedColumnText)

                    ' set row definition
                    If Scanner.MarkState = 2 Then
                        Dim sCombinedRowText As String = sRowText + sRowText + sRowText
                        sBoxChoiceTemplateText = sBoxChoiceTemplateText.Replace(sRowText, sCombinedRowText)
                    End If

                    ' set textbox
                    Dim sCombinedTextboxText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        sCombinedTextboxText += sDVFText
                    End If
                    For i = 0 To oField.MarkCount - 1
                        If Scanner.MarkState = 2 Then
                            For j = 0 To 2
                                Dim sCurrentTextboxText As String = sTextboxText
                                If oBoxImageList.Values(i).Item6 Then
                                    sCurrentTextboxText = sCurrentTextboxText.Replace("TextBox ", "TextBox IsEnabled=""False"" ")
                                End If

                                sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Column=""0""", "Grid.Column=""" + (i + 1).ToString + """")
                                sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Row=""0""", "Grid.Row=""" + j.ToString + """")
                                If j = 2 Or sCurrentTextboxText.Contains("IsEnabled=") Then
                                    sCurrentTextboxText = sCurrentTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + j.ToString + "_" + i.ToString + """ Text=""{Binding Path=MarkBoxChoice" + j.ToString + "[" + i.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                                Else
                                    sCurrentTextboxText = sCurrentTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + j.ToString + "_" + i.ToString + """ IsEnabled=""False"" Text=""{Binding Path=MarkBoxChoice" + j.ToString + "[" + i.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                                End If

                                sCombinedTextboxText += sCurrentTextboxText
                            Next
                        Else
                            Dim sCurrentTextboxText As String = sTextboxText
                            If oBoxImageList.Values(i).Item6 Then
                                sCurrentTextboxText = sCurrentTextboxText.Replace("TextBox ", "TextBox IsEnabled=""False"" ")
                            End If

                            sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Column=""0""", "Grid.Column=""" + i.ToString + """")
                            sCurrentTextboxText = sCurrentTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + Scanner.MarkState.ToString + "_" + i.ToString + """ Text=""{Binding Path=MarkBoxChoice" + Scanner.MarkState.ToString + "[" + i.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + i.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                            sCombinedTextboxText += sCurrentTextboxText
                        End If
                    Next
                    sBoxChoiceTemplateText = sBoxChoiceTemplateText.Replace(sTextboxText, sCombinedTextboxText)
                    sBoxChoiceTemplateText = sBoxChoiceTemplateText.Replace("<DataTemplate ", "<DataTemplate xmlns:i=""clr-namespace:System.Windows.Interactivity;assembly=System.Windows.Interactivity"" xmlns:local=""clr-namespace:Scanner;assembly=Scanner"" ")

                    Dim oBoxChoiceTemplate As DataTemplate = Nothing
                    Using oMemoryStream As New IO.MemoryStream(Text.Encoding.ASCII.GetBytes(sBoxChoiceTemplateText))
                        Dim oParserContext = New Markup.ParserContext()
                        oBoxChoiceTemplate = Markup.XamlReader.Load(oMemoryStream, oParserContext)
                    End Using

                    Return oBoxChoiceTemplate
                Case Enumerations.FieldTypeEnum.Handwriting
                    Dim sHandwritingTemplateText As String = HandwritingTemplateText

                    Dim iRowCount As Integer = (Aggregate oImage In oField.Images Into Max(oImage.Item4)) + 1
                    Dim iColCount As Integer = (Aggregate oImage In oField.Images Into Max(oImage.Item5)) + 1

                    ' set column definition
                    Dim sCombinedColumnText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        sCombinedColumnText += sFixedColumnText
                    End If
                    For i = 0 To iColCount - 1
                        sCombinedColumnText += sColumnText
                    Next
                    sHandwritingTemplateText = sHandwritingTemplateText.Replace(sColumnText, sCombinedColumnText)

                    ' set row definition
                    Dim sCombinedRowText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        For i = 0 To iRowCount - 1
                            sCombinedRowText += sRowText + sRowText + sRowText
                        Next
                        sHandwritingTemplateText = sHandwritingTemplateText.Replace(sRowText, sCombinedRowText)
                    Else
                        For i = 0 To iRowCount - 1
                            sCombinedRowText += sRowText
                        Next
                        sHandwritingTemplateText = sHandwritingTemplateText.Replace(sRowText, sCombinedRowText)
                    End If

                    Dim oRowColList As List(Of Tuple(Of Integer, Integer, Integer)) = (From iIndex In Enumerable.Range(0, oField.Images.Count) Select New Tuple(Of Integer, Integer, Integer)(iIndex, oField.Images(iIndex).Item4, oField.Images(iIndex).Item5)).ToList

                    ' set textbox
                    Dim sCombinedTextboxText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        Dim sModifiedDVFText As String = sDVFText
                        sModifiedDVFText = sModifiedDVFText.Replace("Content=""V"" Margin=""0"" Padding=""0"" Grid.Row=""1"" Grid", "Content=""V"" Margin=""0"" Padding=""0"" Grid.Row=""" + iRowCount.ToString + """ Grid")
                        sModifiedDVFText = sModifiedDVFText.Replace("Content=""F"" Margin=""0"" Padding=""0"" Grid.Row=""2"" Grid", "Content=""F"" Margin=""0"" Padding=""0"" Grid.Row=""" + (iRowCount * 2).ToString + """ Grid")
                        sCombinedTextboxText += sModifiedDVFText
                    End If
                    For iCol = 0 To iColCount - 1
                        For iRow = 0 To iRowCount - 1
                            Dim iCurrentRow As Integer = iRow
                            Dim iCurrentCol As Integer = iCol
                            Dim iIndex As Integer = (From oRowCol In oRowColList Where oRowCol.Item2 = iCurrentRow And oRowCol.Item3 = iCurrentCol Select oRowCol.Item1).First

                            If Scanner.MarkState = 2 Then
                                For j = 0 To 2
                                    Dim sCurrentTextboxText As String = sTextboxText
                                    If oField.Images(iIndex).Item6 Then
                                        sCurrentTextboxText = sCurrentTextboxText.Replace("TextBox ", "TextBox IsEnabled=""False"" ")
                                    End If

                                    sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Column=""0""", "Grid.Column=""" + (iCol + 1).ToString + """")
                                    sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Row=""0""", "Grid.Row=""" + (j * iRowCount + iRow).ToString + """")
                                    If j = 2 Or sCurrentTextboxText.Contains("IsEnabled=") Then
                                        sCurrentTextboxText = sCurrentTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + j.ToString + "_" + iCol.ToString + "_" + iRow.ToString + """ Text=""{Binding Path=MarkHandwritingRowColText[" + iRow.ToString + "." + iCol.ToString + "." + j.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                                    Else
                                        sCurrentTextboxText = sCurrentTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + j.ToString + "_" + iCol.ToString + "_" + iRow.ToString + """ IsEnabled=""False"" Text=""{Binding Path=MarkHandwritingRowColText[" + iRow.ToString + "." + iCol.ToString + "." + j.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                                    End If

                                    sCombinedTextboxText += sCurrentTextboxText
                                Next
                            Else
                                Dim sCurrentTextboxText As String = sTextboxText
                                If oField.Images(iIndex).Item6 Then
                                    sCurrentTextboxText = sCurrentTextboxText.Replace("TextBox ", "TextBox IsEnabled=""False"" ")
                                End If

                                sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Column=""0""", "Grid.Column=""" + iCol.ToString + """")
                                sCurrentTextboxText = sCurrentTextboxText.Replace("Grid.Row=""0""", "Grid.Row=""" + iRow.ToString + """")
                                sCurrentTextboxText = sCurrentTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + Scanner.MarkState.ToString + "_" + iCol.ToString + "_" + iRow.ToString + """ Text=""{Binding Path=MarkHandwritingRowColText[" + iRow.ToString + "." + iCol.ToString + "." + Scanner.MarkState.ToString + "], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + iCol.ToString + "|" + iRow.ToString + "|" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                                sCombinedTextboxText += sCurrentTextboxText
                            End If
                        Next
                    Next
                    sHandwritingTemplateText = sHandwritingTemplateText.Replace(sTextboxText, sCombinedTextboxText)
                    sHandwritingTemplateText = sHandwritingTemplateText.Replace("<DataTemplate ", "<DataTemplate xmlns:i=""clr-namespace:System.Windows.Interactivity;assembly=System.Windows.Interactivity"" xmlns:local=""clr-namespace:Scanner;assembly=Scanner"" ")

                    Dim oHandwritingTemplate As DataTemplate = Nothing
                    Using oMemoryStream As New IO.MemoryStream(Text.Encoding.ASCII.GetBytes(sHandwritingTemplateText))
                        Dim oParserContext = New Markup.ParserContext()
                        oHandwritingTemplate = Markup.XamlReader.Load(oMemoryStream, oParserContext)
                    End Using

                    Return oHandwritingTemplate
                Case Enumerations.FieldTypeEnum.Free
                    Dim sFreeTemplateText As String = FreeTemplateText

                    ' set column definition
                    Dim sCombinedColumnText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        sCombinedColumnText += sFixedColumnText
                    End If
                    sCombinedColumnText += sColumnText
                    sFreeTemplateText = sFreeTemplateText.Replace(sColumnText, sCombinedColumnText)

                    ' set row definition
                    If Scanner.MarkState = 2 Then
                        Dim sCombinedRowText As String = sRowText + sRowText + sRowText
                        sFreeTemplateText = sFreeTemplateText.Replace(sRowText, sCombinedRowText)
                    End If

                    Dim sCombinedFreeText As String = String.Empty
                    If Scanner.MarkState = 2 Then
                        sCombinedFreeText += sDVFText
                    End If
                    If Scanner.MarkState = 2 Then
                        For j = 0 To 2
                            Dim sCurrentFreeTextboxText As String = sTextboxText
                            sCurrentFreeTextboxText = sCurrentFreeTextboxText.Replace("Grid.Column=""0""", "Grid.Column=""1""")
                            sCurrentFreeTextboxText = sCurrentFreeTextboxText.Replace("Grid.Row=""0""", "Grid.Row=""" + j.ToString + """")
                            If j = 2 Then
                                sCurrentFreeTextboxText = sCurrentFreeTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + j.ToString + """ Text=""{Binding Path=MarkFree" + j.ToString + ", Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                            Else
                                sCurrentFreeTextboxText = sCurrentFreeTextboxText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + j.ToString + """ IsEnabled=""False"" Text=""{Binding Path=MarkFree" + j.ToString + ", Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + j.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                            End If

                            sCombinedFreeText += sCurrentFreeTextboxText
                        Next
                    Else
                        sCombinedFreeText = sTextboxText
                        sCombinedFreeText = sCombinedFreeText.Replace(">X</TextBox>", " Name=""Control" + CommonFunctions.SafeFrameworkName(oField.OrderSort) + "_" + CommonFunctions.SafeFrameworkName(oField.Numbering) + "_" + Scanner.MarkState.ToString + """ Text=""{Binding Path=MarkFree" + Scanner.MarkState.ToString + ", Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"" ><i:Interaction.Triggers><i:EventTrigger EventName=""PreviewKeyDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""TextChanged""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger><i:EventTrigger EventName=""PreviewMouseDown""><local:EventCommand Command=""{Binding InteractionEvent}"" CommandParameter=""{Binding RelativeSource={RelativeSource Self}, Path=InvokeParameter}"" CommandArgument=""" + Scanner.MarkState.ToString + """ FieldObject=""{Binding}""/></i:EventTrigger></i:Interaction.Triggers></TextBox>")
                    End If
                    sFreeTemplateText = sFreeTemplateText.Replace(sTextboxText, sCombinedFreeText)
                    sFreeTemplateText = sFreeTemplateText.Replace("<DataTemplate ", "<DataTemplate xmlns:i=""clr-namespace:System.Windows.Interactivity;assembly=System.Windows.Interactivity"" xmlns:local=""clr-namespace:Scanner;assembly=Scanner"" ")

                    Dim oFreeTemplate As DataTemplate = Nothing
                    Using oMemoryStream As New IO.MemoryStream(Text.Encoding.ASCII.GetBytes(sFreeTemplateText))
                        Dim oParserContext = New Markup.ParserContext()
                        oFreeTemplate = Markup.XamlReader.Load(oMemoryStream, oParserContext)
                    End Using

                    Return oFreeTemplate
                Case Else
                    Return MyBase.SelectTemplate(item, container)
            End Select
        End If
    End Function
End Class
Public NotInheritable Class EventCommand
    Inherits Interactivity.TriggerAction(Of DependencyObject)
    Public Shared ReadOnly InvokeParameterProperty As DependencyProperty = DependencyProperty.Register("InvokeParameter", GetType(Object), GetType(EventCommand), Nothing)
    Public Shared ReadOnly CommandProperty As DependencyProperty = DependencyProperty.Register("Command", GetType(ICommand), GetType(EventCommand), Nothing)
    Public Shared ReadOnly CommandParameterProperty As DependencyProperty = DependencyProperty.Register("CommandParameter", GetType(EventArgs), GetType(EventCommand), Nothing)
    Public Shared ReadOnly CommandArgumentProperty As DependencyProperty = DependencyProperty.Register("CommandArgument", GetType(String), GetType(EventCommand), New PropertyMetadata(String.Empty))
    Public Shared ReadOnly FieldObjectProperty As DependencyProperty = DependencyProperty.Register("FieldObject", GetType(FieldDocumentStore.Field), GetType(EventCommand), Nothing)

    Public Property InvokeParameter() As Object
        Get
            Return GetValue(InvokeParameterProperty)
        End Get
        Set
            SetValue(InvokeParameterProperty, Value)
        End Set
    End Property
    Public Property Command() As ICommand
        Get
            Return DirectCast(GetValue(CommandProperty), ICommand)
        End Get
        Set
            SetValue(CommandProperty, Value)
        End Set
    End Property
    Public Property CommandParameter() As EventArgs
        Get
            Return GetValue(CommandParameterProperty)
        End Get
        Set
            SetValue(CommandParameterProperty, Value)
        End Set
    End Property
    Public Property CommandArgument() As String
        Get
            Return GetValue(CommandArgumentProperty)
        End Get
        Set
            SetValue(CommandArgumentProperty, Value)
        End Set
    End Property
    Public Property FieldObject() As FieldDocumentStore.Field
        Get
            Return GetValue(FieldObjectProperty)
        End Get
        Set
            SetValue(FieldObjectProperty, Value)
        End Set
    End Property
    Protected Overrides Sub Invoke(parameter As Object)
        InvokeParameter = parameter
        If Not IsNothing(AssociatedObject) Then
            Dim command As ICommand = Me.Command
            If (Not IsNothing(command)) AndAlso command.CanExecute(New Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)(CommandParameter, CommandArgument, AssociatedObject, FieldObject)) Then
                command.Execute(New Tuple(Of EventArgs, String, FrameworkElement, FieldDocumentStore.Field)(CommandParameter, CommandArgument, AssociatedObject, FieldObject))
            End If
        End If
    End Sub
End Class
#End Region
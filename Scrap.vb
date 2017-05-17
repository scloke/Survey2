Imports Microsoft.VisualBasic

Public Class Scrap
    Private Sub ExpanderShow(ByVal bExpanded As Boolean, ByVal oGUID As Guid)
        ' expands or collapses rows in the grid after clicking the expander icon
        Dim oBorderFormItem As List(Of Controls.Border) = (From oControls In m_Grid.Children Where oControls.GetType.Equals(GetType(Controls.Border)) AndAlso CType(CType(oControls, Controls.Border).Tag, Guid).Equals(oGUID) Select CType(oControls, Controls.Border)).ToList

        If oBorderFormItem.Count > 0 Then
            If bExpanded Then
                oBorderFormItem(0).Visibility = Visibility.Visible
            Else
                oBorderFormItem(0).Visibility = Visibility.Collapsed
            End If

            Dim oFormItem As BaseFormItem = BaseFormItem.Root.FindChild(oBorderFormItem(0).Tag)
            If Not IsNothing(oBorderFormItem(0).Tag) Then
                Dim oImageExpander As ImageExpander = oFormItem.ImageExpander
                Dim oBorderSpacer As BorderSpacer = oFormItem.BorderSpacer
                If Not IsNothing(oImageExpander) Then
                    If bExpanded Then
                        oImageExpander.Visibility = Visibility.Visible
                    Else
                        oImageExpander.Visibility = Visibility.Collapsed
                    End If
                End If
                If Not IsNothing(oBorderSpacer) Then
                    If bExpanded Then
                        oBorderSpacer.Visibility = Visibility.Visible
                    Else
                        oBorderSpacer.Visibility = Visibility.Collapsed
                    End If
                End If
            End If

            For Each oChildGUID As Guid In oFormItem.Children
                ExpanderShow(bExpanded And oFormItem.Expanded, oChildGUID)
            Next
        End If
    End Sub
    Private Shared Sub SetOccupiedRectangle()
        ' checks through the fields in the selected block to indicate if the underlying grid rectangles are occupied by 1) any field, or 2) a restricted field
        Dim oFormField As FormField = TryCast(SelectedItem, FormField)
        Dim oBlock As FormBlock = If(IsNothing(oFormField), Nothing, TryCast(oFormField.Parent, FormBlock))
        If IsNothing(oFormField) And IsNothing(oBlock) Then
            RectangleOccupied = Nothing
        Else
            Dim oFieldNumberingList As List(Of FormField) = (From oChild In oBlock.Children Where (Not IsNothing(TryCast(FormMain.FindChild(oChild.Key), FormField))) AndAlso (Not oChild.Value.Equals(GetType(FieldBorder))) AndAlso (Not oChild.Key.Equals(oFormField.GUID)) Select CType(FormMain.FindChild(oChild.Key), FormField)).ToList
            ReDim RectangleOccupied(oBlock.GridWidth - 1, oBlock.GridHeight - 1)
            For x = 0 To oBlock.GridWidth - 1
                For y = 0 To oBlock.GridHeight - 1
                    Dim iXGridLocation As Integer = x
                    Dim iYGridLocation As Integer = y
                    Dim bOccupiedAnyField As Boolean = If((Aggregate oFormItem In oFieldNumberingList Where iXGridLocation >= oFormItem.GridRect.X AndAlso iXGridLocation < oFormItem.GridRect.X + oFormItem.GridRect.Width AndAlso iYGridLocation >= oFormItem.GridRect.Y AndAlso iYGridLocation < oFormItem.GridRect.Y + oFormItem.GridRect.Height Into Count()) > 0, True, False)
                    Dim bOccupiedRestrictedField As Boolean = If((Aggregate oFormItem In oFieldNumberingList Where oFormItem.IsRestricted AndAlso iXGridLocation >= oFormItem.GridRect.X AndAlso iXGridLocation < oFormItem.GridRect.X + oFormItem.GridRect.Width AndAlso iYGridLocation >= oFormItem.GridRect.Y AndAlso iYGridLocation < oFormItem.GridRect.Y + oFormItem.GridRect.Height Into Count()) > 0, True, False)
                    RectangleOccupied(x, y) = New Tuple(Of Boolean, Boolean)(bOccupiedAnyField, bOccupiedRestrictedField)
                Next
            Next
        End If
    End Sub
    Private Function GetBlockBitmapSource(ByVal bRender As Boolean, ByVal oScaleDirection As Enumerations.ScaleDirection) As Tuple(Of Imaging.BitmapSource, XUnit, XUnit, Integer)
        ' gets the block bitmap with margins as indicated by scale direction
        Dim oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = Nothing
        If bRender Then
            oBlockDimensions = GetBlockDimensions(True)
        Else
            oBlockDimensions = GetBlockDimensions(False)
        End If
        Dim XBlockWidth As XUnit = oBlockDimensions.Item1
        Dim XBlockHeight As XUnit = oBlockDimensions.Item2
        Dim XExpandedBlockWidth As XUnit = XUnit.Zero
        Dim XExpandedBlockHeight As XUnit = XUnit.Zero
        Dim iStart As Integer = -1

        ' get expanded block dimensions
        If oScaleDirection And Enumerations.ScaleDirection.Horizontal = Enumerations.ScaleDirection.Horizontal Then
            XExpandedBlockWidth = New XUnit(XBlockWidth.Point + (2 * PDFHelper.BlockSpacer.Point))
        Else
            XExpandedBlockWidth = XBlockWidth
        End If
        If oScaleDirection And Enumerations.ScaleDirection.Vertical = Enumerations.ScaleDirection.Vertical Then
            XExpandedBlockHeight = New XUnit(XBlockHeight.Point + (2 * PDFHelper.BlockSpacer.Point))
        Else
            XExpandedBlockHeight = XBlockHeight
        End If
        Dim XExpandedBlockSize As New XSize(XExpandedBlockWidth.Point, XExpandedBlockHeight.Point)

        Dim width As Integer = Math.Ceiling(XExpandedBlockWidth.Inch * RenderResolution300)
        Dim height As Integer = Math.Ceiling(XExpandedBlockHeight.Inch * RenderResolution300)

        Dim oBitmapSource As Imaging.BitmapSource = Nothing
        Using oBitmap As New System.Drawing.Bitmap(width, height)
            oBitmap.SetResolution(RenderResolution300, RenderResolution300)

            Dim oParamList As New Dictionary(Of String, Object)
            oParamList.Add(FormPDF.KeySubjectName, StringSubject)

            Dim iIterator As Integer = 0
            Dim bReturn As Boolean = False
            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XExpandedBlockSize, XGraphicsUnit.Point)
                    Do Until bReturn
                        Dim oContainer = oXGraphics.BeginContainer()
                        Dim oBlockReturn = DrawBlockDirect(bRender, iIterator, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, New XPoint(0, 0), XBlockWidth, XBlockHeight, New XPoint(PDFHelper.BlockSpacer.Point, PDFHelper.BlockSpacer.Point), oParamList, oBlockDimensions.Item3)
                        bReturn = oBlockReturn.Item1

                        ' get numbering for block
                        If oBlockReturn.Item4 <> -1 Then
                            iStart = oBlockReturn.Item4
                        End If
                        oXGraphics.EndContainer(oContainer)
                    Loop
                End Using
            End Using

            oBitmapSource = Converter.BitmapToBitmapSource(oBitmap)
        End Using

        Return New Tuple(Of Imaging.BitmapSource, XUnit, XUnit, Integer)(oBitmapSource, oBlockDimensions.Item1, oBlockDimensions.Item2, iStart)
    End Function
    Public Function DrawBlockDirect(ByVal bRender As Boolean, ByRef iIterator As Integer, ByRef oXGraphics As XGraphics, ByVal XBlockWidth As XUnit, XBlockHeight As XUnit, ByVal XDisplacement As XPoint, ByVal XContentWidth As XUnit, XContentHeight As XUnit, ByVal XContentDisplacement As XPoint, ByVal oParamList As Dictionary(Of String, Object), Optional ByVal oOverflowRows As List(Of Integer) = Nothing) As Tuple(Of Boolean, XUnit, XUnit, Integer)
        ' draw directly on the supplied XGraphics with the specified displacement
        ' iInitialItems represents the number of items to draw before the form fields (in this case, only the background)
        Const iInitialItems As Integer = 2

        Dim bReturn As Boolean = False
        Dim iStart As Integer = -1
        Dim oFormFields As List(Of FormField) = (From oChild In Children Where oChild.Value.IsSubclassOf(GetType(FormField)) Select CType(FormMain.FindChild(oChild.Key), FormField)).ToList
        Dim oPlacedFormFields As List(Of FormField) = (From oFormField In oFormFields Where oFormField.ContentField AndAlso (Not oFormField.GridRect.IsEmpty) AndAlso oFormField.Placed Select oFormField).ToList
        If iIterator = 0 Then
            ' draw background
            Dim oFieldBackgroundWholeBlockList As List(Of FieldBackground) = (From iIndex In Enumerable.Range(0, Children.Count) Where Children.Values(iIndex).Equals(GetType(FieldBackground)) AndAlso CType(FormMain.FindChild(Children.Keys(iIndex)), FormField).WholeBlock Select CType(FormMain.FindChild(Children.Keys(iIndex)), FieldBackground)).ToList
            Dim oFieldBackgroundList As List(Of FieldBackground) = (From iIndex In Enumerable.Range(0, Children.Count) Where Children.Values(iIndex).Equals(GetType(FieldBackground)) AndAlso (Not CType(FormMain.FindChild(Children.Keys(iIndex)), FormField).WholeBlock) Select CType(FormMain.FindChild(Children.Keys(iIndex)), FieldBackground)).ToList

            If oFieldBackgroundWholeBlockList.Count > 0 Then
                Select Case oFieldBackgroundWholeBlockList.First.Lightness
                    Case 1
                        DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.WhiteSmoke)
                    Case 2
                        DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.LightGray)
                    Case 3
                        DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.Silver)
                    Case 4
                        DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.DarkGray)
                End Select
            Else
                DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.White)
            End If
            For Each oBackground In oFieldBackgroundList
                Select Case oBackground.Lightness
                    Case 1
                        DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XContentDisplacement, XBrushes.WhiteSmoke, GridWidth, GridHeight, oBackground.GridRect)
                    Case 2
                        DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XContentDisplacement, XBrushes.LightGray, GridWidth, GridHeight, oBackground.GridRect)
                    Case 3
                        DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XContentDisplacement, XBrushes.Silver, GridWidth, GridHeight, oBackground.GridRect)
                    Case 4
                        DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XContentDisplacement, XBrushes.DarkGray, GridWidth, GridHeight, oBackground.GridRect)
                End Select
            Next
        ElseIf iIterator = 1 Then
            ' draw numbering
            If Children.Values.Contains(GetType(FieldNumbering)) Then
                Dim oFieldNumbering As FieldNumbering = (From iIndex In Enumerable.Range(0, Children.Count) Where Children.Values(iIndex).Equals(GetType(FieldNumbering)) Select CType(FormMain.FindChild(Children.Keys(iIndex)), FieldNumbering)).First
                If Not oFieldNumbering.Spacer Then
                    ' set background tint
                    Dim oFieldBackgroundWholeBlockList As List(Of FieldBackground) = (From iIndex In Enumerable.Range(0, Children.Count) Where Children.Values(iIndex).Equals(GetType(FieldBackground)) AndAlso CType(FormMain.FindChild(Children.Keys(iIndex)), FormField).WholeBlock Select CType(FormMain.FindChild(Children.Keys(iIndex)), FieldBackground)).ToList
                    Dim oBorderBackground As XBrush = Nothing
                    If oFieldBackgroundWholeBlockList.Count > 0 Then
                        Select Case oFieldBackgroundWholeBlockList.First.Lightness
                            Case 1
                                oBorderBackground = XBrushes.WhiteSmoke
                            Case 2
                                oBorderBackground = XBrushes.LightGray
                            Case 3
                                oBorderBackground = XBrushes.Silver
                            Case 4
                                oBorderBackground = XBrushes.DarkGray
                        End Select
                    Else
                        oBorderBackground = XBrushes.White
                    End If

                    ' set section parameters
                    Dim oSection As FormSection = FindParent(GetType(FormSection))
                    If Not IsNothing(oSection) Then
                        With oSection
                            iStart = .Start
                            Dim oBlockList As List(Of FormBlock) = (From oChild In oSection.Children Where oChild.Value.Equals(GetType(FormBlock)) Select CType(FormMain.FindChild(oChild.Key), FormBlock)).ToList
                            For Each oBlock In oBlockList
                                If oBlock.Children.Values.Contains(GetType(FieldNumbering)) Then
                                    Dim oBlockNumbering As FieldNumbering = (From iIndex In Enumerable.Range(0, oBlock.Children.Count) Where oBlock.Children.Values(iIndex).Equals(GetType(FieldNumbering)) Select CType(FormMain.FindChild(oBlock.Children.Keys(iIndex)), FieldNumbering)).First
                                    If oBlockNumbering.GUID.Equals(oFieldNumbering.GUID) Then
                                        Exit For
                                    End If
                                    If Not oBlockNumbering.Spacer Then
                                        iStart += 1
                                    End If
                                End If
                            Next
                            DrawFieldNumbering(oXGraphics, XBlockWidth, XBlockHeight, New XPoint(XDisplacement.X + XContentDisplacement.X, XDisplacement.Y + XContentDisplacement.Y), iStart, .NumberingType, .NumberingBackground, .NumberingBorder, oBorderBackground)
                        End With
                    End If
                End If
            End If
        ElseIf iIterator = oPlacedFormFields.Count + iInitialItems Then
            ' draw border
            ' last item to draw
            Dim oFieldBorderWholeBlockList As List(Of FieldBorder) = (From iIndex In Enumerable.Range(0, Children.Count) Where Children.Values(iIndex).Equals(GetType(FieldBorder)) AndAlso CType(FormMain.FindChild(Children.Keys(iIndex)), FormField).WholeBlock Select CType(FormMain.FindChild(Children.Keys(iIndex)), FieldBorder)).ToList
            Dim oFieldBorderList As List(Of FieldBorder) = (From iIndex In Enumerable.Range(0, Children.Count) Where Children.Values(iIndex).Equals(GetType(FieldBorder)) AndAlso (Not CType(FormMain.FindChild(Children.Keys(iIndex)), FormField).WholeBlock) Select CType(FormMain.FindChild(Children.Keys(iIndex)), FieldBorder)).ToList

            If oFieldBorderWholeBlockList.Count > 0 Then
                DrawFieldBorder(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, oFieldBorderWholeBlockList.First.BorderWidth, 0, 0, Nothing)
            End If
            For Each oBorder In oFieldBorderList
                DrawFieldBorder(oXGraphics, XContentWidth, XContentHeight, XContentDisplacement, oBorder.BorderWidth, GridWidth, GridHeight, oBorder.GridRect)
            Next

            bReturn = True
        Else
            ' runs through each form field that has been placed
            Dim oFormField As FormField = oPlacedFormFields(iIterator - iInitialItems)

            Dim XCombinedDisplacement As New XPoint(XDisplacement.X + XContentDisplacement.X, XDisplacement.Y + XContentDisplacement.Y)
            Dim fLeftIndent As Double = LeftIndent()
            Dim XBlockSize As New XSize(XContentWidth.Point, XContentHeight.Point)
            Dim XImageWidth As XUnit = XUnit.FromPoint(((XBlockSize.Width - fLeftIndent) * oFormField.GridRect.Width / GridWidth))
            Dim XImageHeight As XUnit = Nothing
            Dim XFieldDisplacement As XPoint = Nothing
            If IsNothing(oOverflowRows) Then
                XImageHeight = XUnit.FromPoint(XBlockSize.Height * oFormField.GridRect.Height / GridHeight)
                XFieldDisplacement = New XPoint(XUnit.FromPoint(((XBlockSize.Width - fLeftIndent) * oFormField.GridRect.X / GridWidth) + XCombinedDisplacement.X + fLeftIndent), XUnit.FromPoint((XBlockSize.Height * oFormField.GridRect.Y / GridHeight) + XCombinedDisplacement.Y))
            Else
                Dim iPreviousOverflows As Integer = 0
                For i = 0 To oFormField.GridRect.Y - 1
                    iPreviousOverflows += oOverflowRows(i)
                Next

                XImageHeight = XUnit.FromPoint(XBlockSize.Height * oFormField.GridAdjustedHeight / (GridHeight + oOverflowRows.Sum))
                XFieldDisplacement = New XPoint(XUnit.FromPoint(((XBlockSize.Width - fLeftIndent) * oFormField.GridRect.X / GridWidth) + XCombinedDisplacement.X + fLeftIndent), XUnit.FromPoint((XBlockSize.Height * (oFormField.GridRect.Y + iPreviousOverflows) / (GridHeight + oOverflowRows.Sum)) + XCombinedDisplacement.Y))
            End If

            oFormField.DrawFieldContents(bRender, oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, oParamList, oOverflowRows)
        End If

        iIterator += 1
        Return New Tuple(Of Boolean, XUnit, XUnit, Integer)(bReturn, XBlockWidth, XBlockHeight, iStart)
    End Function
    Public Shared Function DrawFieldTabletsMCQ(ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XDisplacement As XPoint, ByVal oGridRect As Int32Rect, ByVal iGridHeight As Integer, ByVal iFontSizeMultiplier As Integer, ByVal oOverflowRows As List(Of Integer), ByRef oInputField As InputFieldDocumentStore.InputField, Optional ByVal oTabletImages As List(Of XImage) = Nothing, Optional ByVal oTabletDisplacements As List(Of Tuple(Of XRect, Integer, Integer, XRect, List(Of ElementStruc))) = Nothing) As Rect
        ' draws tablets for MCQ field
        Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

        Dim fSingleBlockWidth As Double = BlockHeight.Point * 2
        Dim XAdjustedWidth As New XUnit((XImageWidth.Point - (oInputField.TabletGroups * fSingleBlockWidth)) / oInputField.TabletGroups)
        Dim oSingleBoxSize As New XSize(XImageWidth.Point, BlockHeight.Point)
        Dim oReturnRect As New Rect

        ' draw letters
        Const fFontSize As Double = 10
        Const fTabletHeight As Double = 0.5
        Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
        Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
        Dim fScaledFontSize As Double = fFontSize * 0.8 * (fSingleBlockWidth * fTabletHeight) / oTestFont.GetHeight
        Dim oArielFont As New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)

        Dim oStringFormat As New XStringFormat()
        oStringFormat.Alignment = XStringAlignment.Center
        oStringFormat.LineAlignment = XLineAlignment.Center

        Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
        oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
        oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Subject, MigraDoc.DocumentObjectModel.Colors.DarkViolet)
        oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Field, MigraDoc.DocumentObjectModel.Colors.Green)

        Dim oLayoutList As New List(Of Tuple(Of Integer, Integer))
        For Each oTabletElements As Tuple(Of Rect, Integer, Integer, List(Of ElementStruc)) In oInputField.TabletDescriptionMCQ
            Dim oLayout As LayoutClass = GetLayout(XAdjustedWidth, XImageHeight, oGridRect.Height, iFontSizeMultiplier, oTabletElements.Item4, -1).Item1
            oLayoutList.Add(New Tuple(Of Integer, Integer)(Math.Max(oLayout.LineCount, 2), oLayout.LineCount))
        Next

        Dim iRowLimit As Integer = Math.Ceiling((Aggregate oLayoutHeight As Tuple(Of Integer, Integer) In oLayoutList Into Sum(oLayoutHeight.Item1)) / oInputField.TabletGroups)
        Dim iRowCount As Integer = 0
        For Each oLayoutHeight As Tuple(Of Integer, Integer) In oLayoutList
            iRowCount += oLayoutHeight.Item1
            If iRowCount >= iRowLimit Then
                Exit For
            End If
        Next

        ' draw tablets
        Dim iCurrentRow As Integer = 0
        Dim iCurrentColumn As Integer = 0
        Dim XCheckEmptyBitmapSize As XSize = Nothing
        For i = 0 To If(IsNothing(oTabletDisplacements), oInputField.TabletDescriptionMCQ.Count - 1, oTabletDisplacements.Count - 1)
            Dim sContentText As String = String.Empty
            Select Case oInputField.TabletContent
                Case Enumerations.TabletContentEnum.Number
                    sContentText = (i + If(oInputField.TabletStart = -2, 0, oInputField.TabletStart + 1)).ToString
                Case Enumerations.TabletContentEnum.Letter
                    sContentText = Converter.ConvertNumberToLetter(i + Math.Max(oInputField.TabletStart, 0), True)
            End Select

            Dim fDisplacementLeft As Double = 0
            Dim fDisplacementTop As Double = 0

            ' use supplied tablet displacements if given
            If (Not IsNothing(oTabletDisplacements)) Then
                fDisplacementLeft = oTabletDisplacements(i).Item1.X + oTabletDisplacements(i).Item1.Width / 2
                fDisplacementTop = oTabletDisplacements(i).Item1.Y + oTabletDisplacements(i).Item1.Height / 2
            Else
                fDisplacementLeft = XDisplacement.X + (fSingleBlockWidth / 2) + ((XAdjustedWidth.Point + fSingleBlockWidth) * iCurrentColumn)
                fDisplacementTop = XDisplacement.Y + (oLayoutList(i).Item1 * iFontSizeMultiplier * oSingleBoxSize.Height / 2) + (iCurrentRow * oSingleBoxSize.Height)
            End If

            XCheckEmptyBitmapSize = DrawTablet(oXGraphics, New XPoint(fDisplacementLeft, fDisplacementTop), sContentText, oArielFont, oStringFormat, If(IsNothing(oTabletImages) OrElse i >= oTabletImages.Count, Nothing, oTabletImages(i)))
            Dim XLeft As New XUnit(fDisplacementLeft - (XCheckEmptyBitmapSize.Width / 2))
            Dim XTop As XUnit = New XUnit(fDisplacementTop - (XCheckEmptyBitmapSize.Height / 2))
            Dim XWidth As New XUnit(XCheckEmptyBitmapSize.Width)
            Dim XHeight As New XUnit(XCheckEmptyBitmapSize.Height)
            Dim oImageRect As New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point)

            If IsNothing(oTabletImages) Then
                oInputField.AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oImageRect, Nothing, String.Empty, 0, i, False, 0, New Tuple(Of Single)(0)))
            End If

            If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                oReturnRect = oImageRect
            Else
                oReturnRect.Union(oImageRect)
            End If

            ' draw text
            Dim iLinesInBox As Integer = 0
            Dim iLinesPerLine As Integer = 0
            Dim oElements As List(Of ElementStruc) = Nothing
            Dim oAdjustedWidth As XUnit = Nothing
            Dim oAdjustedHeight As XUnit = Nothing
            If (Not IsNothing(oTabletDisplacements)) Then
                oAdjustedWidth = oTabletDisplacements(i).Item4.Width
                oAdjustedHeight = oTabletDisplacements(i).Item4.Height
                fDisplacementLeft = oTabletDisplacements(i).Item4.X
                fDisplacementTop = oTabletDisplacements(i).Item4.Y
                iLinesInBox = oTabletDisplacements(i).Item2
                iLinesPerLine = oTabletDisplacements(i).Item3
                oElements = oTabletDisplacements(i).Item5
            Else
                oAdjustedWidth = XAdjustedWidth
                oAdjustedHeight = XImageHeight
                fDisplacementLeft = XDisplacement.X + fSingleBlockWidth + ((XAdjustedWidth.Point + fSingleBlockWidth) * iCurrentColumn)
                fDisplacementTop = XDisplacement.Y + ((oLayoutList(i).Item1 - oLayoutList(i).Item2) * iFontSizeMultiplier * oSingleBoxSize.Height / 2) + (iCurrentRow * oSingleBoxSize.Height)
                iLinesInBox = oGridRect.Height
                iLinesPerLine = iFontSizeMultiplier
                oElements = oInputField.TabletDescriptionMCQ(i).Item4
            End If

            DrawFieldText(oXGraphics, oAdjustedWidth, oAdjustedHeight, New XPoint(fDisplacementLeft, fDisplacementTop), iLinesInBox, iLinesPerLine, oElements, -1, MigraDoc.DocumentObjectModel.ParagraphAlignment.Left, oFontColourDictionary)

            If IsNothing(oTabletDisplacements) Then
                If (Not IsNothing(oInputField.TabletDescriptionMCQ)) AndAlso oInputField.TabletDescriptionMCQ(i).Item1.IsEmpty Then
                    oInputField.TabletDescriptionMCQ(i) = New Tuple(Of Rect, Integer, Integer, List(Of ElementStruc))(New Rect(fDisplacementLeft, fDisplacementTop, XAdjustedWidth, XImageHeight), oGridRect.Height, iFontSizeMultiplier, oInputField.TabletDescriptionMCQ(i).Item4)
                End If

                Dim oTextRect As New Rect(fDisplacementLeft, fDisplacementTop, XAdjustedWidth.Point, oLayoutList(i).Item2 * iFontSizeMultiplier * oSingleBoxSize.Height)
                If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                    oReturnRect = oTextRect
                Else
                    oReturnRect.Union(oTextRect)
                End If

                iCurrentRow += oLayoutList(i).Item1
                If i < oLayoutList.Count - 1 AndAlso iCurrentRow + oLayoutList(i + 1).Item1 > iRowCount Then
                    iCurrentRow = 0
                    iCurrentColumn += 1
                End If
            End If
        Next

        oXGraphics.SmoothingMode = oSmoothingMode
        Return oReturnRect
    End Function
    Public Class Suspender
        ' suspends
        Implements IDisposable

        Private Shared m_SuspendDictionary As New Dictionary(Of BaseFormItem, Tuple(Of Integer, Boolean))
        Private Shared m_GlobalSuspendState As Boolean = False
        Private m_FormItem As BaseFormItem = Nothing
        Private m_PreviousGlobalSuspendState As Boolean = False

        Sub New()
            m_FormItem = Nothing
            m_PreviousGlobalSuspendState = GlobalSuspendUpdates
            GlobalSuspendUpdates = True
        End Sub
        Sub New(ByVal oFormItem As BaseFormItem, ByVal bContentChanged As Boolean)
            m_FormItem = oFormItem
            If Not m_SuspendDictionary.ContainsKey(m_FormItem) Then
                m_SuspendDictionary.Add(m_FormItem, New Tuple(Of Integer, Boolean)(0, bContentChanged))
            End If
            m_SuspendDictionary(m_FormItem) = New Tuple(Of Integer, Boolean)(m_SuspendDictionary(m_FormItem).Item1 + 1, m_SuspendDictionary(m_FormItem).Item2)
        End Sub
        Public Shared Property GlobalSuspendUpdates As Boolean
            Get
                Return m_GlobalSuspendState
            End Get
            Set(value As Boolean)
                m_GlobalSuspendState = value
                If Not m_GlobalSuspendState Then
                    ' update all pending content changes
                    For i = m_SuspendDictionary.Count - 1 To 0 Step -1
                        Dim oFormItem As BaseFormItem = m_SuspendDictionary.Keys(i)
                        If m_SuspendDictionary(oFormItem).Item1 <= 0 Then
                            RemoveItem(oFormItem)
                        End If
                    Next
                End If
            End Set
        End Property
        Private Shared Sub RemoveItem(ByVal oFormItem As BaseFormItem)
            If Not IsNothing(oFormItem) Then
                If m_SuspendDictionary(oFormItem).Item2 Then
                    ' trigger content changed 
                    oFormItem.ContentChanged()
                End If

                ' remove dictionary reference
                m_SuspendDictionary.Remove(oFormItem)
            End If
        End Sub
        Public Shared Sub DeleteEntry(ByVal oFormItem As BaseFormItem)
            ' removes entry from dictionary
            If m_SuspendDictionary.ContainsKey(oFormItem) Then
                m_SuspendDictionary.Remove(oFormItem)
            End If
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Shadows Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If IsNothing(m_FormItem) Then
                        ' global suspend
                        GlobalSuspendUpdates = m_PreviousGlobalSuspendState
                    Else
                        ' item suspend
                        m_SuspendDictionary(m_FormItem) = New Tuple(Of Integer, Boolean)(m_SuspendDictionary(m_FormItem).Item1 - 1, m_SuspendDictionary(m_FormItem).Item2)
                        If (Not m_GlobalSuspendState) AndAlso m_SuspendDictionary(m_FormItem).Item1 <= 0 Then
                            RemoveItem(m_FormItem)
                        End If
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
    Public Overrides Sub Display()
        Dim oScrollViewerBlockContent As Controls.ScrollViewer = ElementDictionary("ScrollViewerBlockContent")
        Dim oScrollBar As Controls.Primitives.ScrollBar = CType(oScrollViewerBlockContent.Template.FindName("PART_VerticalScrollBar", oScrollViewerBlockContent), Controls.Primitives.ScrollBar)

        If Not IsNothing(oScrollBar) Then
            Dim oBitmapSource As Imaging.BitmapSource = GetSectionTitleBitmapSource().Item1
            Dim fBlockSpacing As Double = PDFHelper.BlockSpacer.Inch * RenderResolution300
            Dim fScaleFactor As Double = Math.Min((oScrollViewerBlockContent.ActualWidth - oScrollBar.ActualWidth) / oBitmapSource.Width, oScrollViewerBlockContent.ActualHeight / oBitmapSource.Height)
            Dim fImageWidth As Double = oBitmapSource.Width * fScaleFactor
            Dim fImageHeight As Double = oBitmapSource.Height * fScaleFactor

            ' set image
            Dim oRectangleBlockBackground As Shapes.Rectangle = ElementDictionary("RectangleBlockBackground")
            With oRectangleBlockBackground
                .Width = fImageWidth
                .Height = fImageHeight
            End With

            Dim oImageBlockContent As Controls.Image = ElementDictionary("ImageBlockContent")
            With oImageBlockContent
                .Width = fImageWidth
                .Height = fImageHeight
                .Source = oBitmapSource
                .Tag = fScaleFactor
            End With

            ' set overlay canvas
            Dim oCanvasBlockContent As Controls.Canvas = ElementDictionary("CanvasBlockContent")
            With oCanvasBlockContent
                .Width = fImageWidth
                .Height = fImageHeight
                .Margin = New Thickness(0, 0, oScrollBar.ActualWidth, 0)
                .Children.Clear()
            End With
        End If
    End Sub
    Private Shared Function BitmapToBitmapSourceSmall(ByVal oBitmap As System.Drawing.Bitmap) As Media.Imaging.BitmapSource
        ' converts GDI bitmap to WPF bitmapsource (small sizes)
        If IsNothing(oBitmap) Then
            Return Nothing
        Else
            Dim oHbitmap As IntPtr = oBitmap.GetHbitmap()
            Dim oBitmapSource As Media.Imaging.BitmapSource = Nothing
            Try
                oBitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(oHbitmap, IntPtr.Zero, Int32Rect.Empty, Media.Imaging.BitmapSizeOptions.FromWidthAndHeight(oBitmap.Width, oBitmap.Height))
            Catch ex As OutOfMemoryException
                CommonFunctions.ClearMemory(True)
                oBitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(oHbitmap, IntPtr.Zero, Int32Rect.Empty, Media.Imaging.BitmapSizeOptions.FromWidthAndHeight(oBitmap.Width, oBitmap.Height))
            End Try
            NativeMethods.DeleteObject(oHbitmap)
            Return oBitmapSource
        End If
    End Function
    Public Overrides Sub RenderPDF(ByRef oPDFDocument As Pdf.PdfDocument, ByRef XCurrentHeight As XUnit, ByRef oParamList As Dictionary(Of String, Object))
        Dim XSectionSize As XSize = GetSectionTitleDimensions()

        FormPDF.AddPage(oPDFDocument, XCurrentHeight, XSectionSize.Height, oParamList)
        Dim XDisplacement As New XPoint(PDFHelper.PageLimitLeft.Point, PDFHelper.PageLimitTop.Point + XCurrentHeight.Point)

        ' add section title
        Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oParamList(KeyPDFPages)(oParamList(KeyPDFPages).Count - 1).Item1, XGraphicsPdfPageOptions.Append)
            DrawSectionTitleDirect(oXGraphics, New XUnit(XSectionSize.Width), New XUnit(XSectionSize.Height), XDisplacement)
        End Using
        XCurrentHeight = New XUnit(XCurrentHeight.Point + XSectionSize.Height + PDFHelper.BlockSpacer.Point)

        ' reset current numbering to zero
        If oParamList.ContainsKey(KeyNumberingCurrent) Then
            oParamList(KeyNumberingCurrent) = Start
        Else
            oParamList.Add(KeyNumberingCurrent, Start)
        End If

        ' set current section
        If oParamList.ContainsKey(KeyCurrentSection) Then
            oParamList(KeyCurrentSection) = Me
        Else
            oParamList.Add(KeyCurrentSection, Me)
        End If

        ' runs through children
        Dim oLineBlockList As New List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String()))
        Dim oChildDictionary As New Dictionary(Of Guid, Type)
        For Each oChild In Children
            If oChild.Value.Equals(GetType(FormatterGroup)) Then
                Dim oGroup As FormatterGroup = FormMain.FindChild(oChild.Key)
                For Each oGroupChild In oGroup.Children
                    oChildDictionary.Add(oGroupChild.Key, oGroupChild.Value)
                Next
            Else
                oChildDictionary.Add(oChild.Key, oChild.Value)
            End If
        Next

        For Each oChild In oChildDictionary
            If oChild.Value.Equals(GetType(FormBlock)) Then
                Dim oBlock As FormBlock = FormMain.FindChild(oChild.Key)
                Dim iBlockCount As Integer = Aggregate oLineBlock In oLineBlockList Into Sum(oLineBlock.Item1.BlockWidth)
                If iBlockCount + oBlock.BlockWidth > PDFHelper.PageBlockWidth Then
                    ' output current line and reset block list
                    RenderCurrentLine(oPDFDocument, XCurrentHeight, oParamList, oLineBlockList)
                End If

                ' add to block list
                oLineBlockList.Add(New Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String())(oBlock, oBlock.GetBlockDimensions(True), Nothing))
            Else
                ' output current line and reset block list
                RenderCurrentLine(oPDFDocument, XCurrentHeight, oParamList, oLineBlockList)

                Select Case oChild.Value
                    Case GetType(FormSubSection)
                        Dim oSubSection As FormSubSection = FormMain.FindChild(oChild.Key)
                        oParamList(KeyNumberingCurrent) = GetNumbering(oSubSection)

                        ' set a divider if the previous field is not another subsection or divider
                        Dim iCurrentIndex As Integer = oChildDictionary.ToList.IndexOf(oChild)
                        If oSubSection.Children.ContainsValue(GetType(FormatterDivider)) Then
                            If iCurrentIndex = 0 Then
                                SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                            ElseIf Not (oChildDictionary.Values(iCurrentIndex - 1).Equals(GetType(FormatterDivider)) Or oChildDictionary.Values(iCurrentIndex - 1).Equals(GetType(FormSubSection))) Then
                                If oChildDictionary.Values(iCurrentIndex - 1).Equals(GetType(FormSubSection)) Then
                                    Dim oPreviousSubSection As FormSubSection = FormMain.FindChild(oChildDictionary.Keys(iCurrentIndex - 1))
                                    If Not oPreviousSubSection.Children.ContainsValue(GetType(FormatterDivider)) Then
                                        SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                                    End If
                                Else
                                    SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                                End If
                            End If
                        End If

                        If ContinuousNumbering Then
                            If oParamList.ContainsKey(KeySubNumberingCurrent) Then
                                oParamList.Remove(KeySubNumberingCurrent)
                            End If
                            oSubSection.RenderPDF(oPDFDocument, XCurrentHeight, oParamList)
                        Else
                            If Not oParamList.ContainsKey(KeySubNumberingCurrent) Then
                                oParamList.Add(KeySubNumberingCurrent, 0)
                            End If
                            oSubSection.RenderPDF(oPDFDocument, XCurrentHeight, oParamList)
                            oParamList.Remove(KeySubNumberingCurrent)
                        End If

                        ' sets a divider after the oSubSection fields
                        If oSubSection.Children.ContainsValue(GetType(FormatterDivider)) Then
                            If (iCurrentIndex = oChildDictionary.Count - 1) OrElse Not oChildDictionary.Values(iCurrentIndex + 1).Equals(GetType(FormatterDivider)) Then
                                SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                            End If
                        End If
                    Case GetType(FormMCQ)
                        Dim oMCQ As FormMCQ = FormMain.FindChild(oChild.Key)
                        oParamList(KeyNumberingCurrent) = GetNumbering(oMCQ)

                        ' set a divider if the previous field is not another MCQ or divider
                        Dim iCurrentIndex As Integer = oChildDictionary.ToList.IndexOf(oChild)
                        If oMCQ.Children.ContainsValue(GetType(FormatterDivider)) Then
                            If iCurrentIndex = 0 Then
                                SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                            ElseIf Not (oChildDictionary.Values(iCurrentIndex - 1).Equals(GetType(FormatterDivider)) Or oChildDictionary.Values(iCurrentIndex - 1).Equals(GetType(FormMCQ))) Then
                                If oChildDictionary.Values(iCurrentIndex - 1).Equals(GetType(FormMCQ)) Then
                                    Dim oPreviousMCQ As FormMCQ = FormMain.FindChild(oChildDictionary.Keys(iCurrentIndex - 1))
                                    If Not oPreviousMCQ.Children.ContainsValue(GetType(FormatterDivider)) Then
                                        SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                                    End If
                                Else
                                    SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                                End If
                            End If
                        End If

                        If ContinuousNumbering Then
                            If oParamList.ContainsKey(KeySubNumberingCurrent) Then
                                oParamList.Remove(KeySubNumberingCurrent)
                            End If
                            oMCQ.RenderPDF(oPDFDocument, XCurrentHeight, oParamList)
                        Else
                            If Not oParamList.ContainsKey(KeySubNumberingCurrent) Then
                                oParamList.Add(KeySubNumberingCurrent, 0)
                            End If
                            oMCQ.RenderPDF(oPDFDocument, XCurrentHeight, oParamList)
                            oParamList.Remove(KeySubNumberingCurrent)
                        End If

                        ' sets a divider after the MCQ fields
                        If oMCQ.Children.ContainsValue(GetType(FormatterDivider)) Then
                            If (iCurrentIndex = oChildDictionary.Count - 1) OrElse Not oChildDictionary.Values(iCurrentIndex + 1).Equals(GetType(FormatterDivider)) Then
                                SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                            End If
                        End If
                    Case GetType(FormatterDivider)
                        SetDivider(oPDFDocument, XCurrentHeight, oParamList)
                    Case GetType(FormatterGroup)
                    Case GetType(FormatterPageBreak)
                        XCurrentHeight = New XUnit(PDFHelper.PageLimitBottomExBarcode.Point + 1)
                        FormPDF.AddPage(oPDFDocument, XCurrentHeight, 0, oParamList)
                End Select
            End If
        Next

        ' output current line and reset block list
        RenderCurrentLine(oPDFDocument, XCurrentHeight, oParamList, oLineBlockList)

        ' remove current section
        If oParamList.ContainsKey(KeyCurrentSection) Then
            oParamList.Remove(KeyCurrentSection)
        End If
    End Sub
    Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
        Dim oImage As DrawingImage = CType(m_Icons("CCMImageSample").Clone, DrawingImage)
        Dim fDesiredWidth As Double = fReferenceHeight * FieldTypeMultiplier * FieldTypeAspectRatio
        Dim fDesiredHeight As Double = fReferenceHeight * FieldTypeMultiplier
        CType(oImage.Drawing, DrawingGroup).Transform = New ScaleTransform(fDesiredWidth / oImage.Width, fDesiredHeight / oImage.Height)
        Return oImage
    End Function
    Public Overrides Sub TitleChanged()
        Dim sTitle As String = "Title: " + SectionTitle
        sTitle += vbCr + "Type: "
        Select Case m_DataObject.NumberingType
            Case Enumerations.Numbering.Number
                sTitle += "Number"
                sTitle += vbCr + "Start: " + (m_DataObject.Start + 1).ToString
            Case Enumerations.Numbering.LetterSmall
                sTitle += "Lower Case Letter"
                sTitle += vbCr + "Start: " + Converter.ConvertNumberToLetter(m_DataObject.Start, False)
            Case Enumerations.Numbering.LetterBig
                sTitle += "Upper Case Letter"
                sTitle += vbCr + "Start: " + Converter.ConvertNumberToLetter(m_DataObject.Start, True)
        End Select
        sTitle += vbCr + "Numbering: "
        Dim oNumberingList As New List(Of String)

        If NumberingBorder Then
            oNumberingList.Add("Border")
        End If
        If NumberingBackground Then
            oNumberingList.Add("Background")
        End If
        If ContinuousNumbering Then
            oNumberingList.Add("Continuous")
        End If
        sTitle += String.Join(", ", oNumberingList.ToArray)

        Title = sTitle
    End Sub
    Public Class DisplayObservableCollection
        Inherits ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)

        Public Sub Refresh(ByVal oDisplayList As List(Of Common.HighlightComboBox.HCBDisplay), ByRef oFrameworkElement As FrameworkElement, ByVal oDataContext As Object)
            CheckReentrancy()

            If Not IsNothing(oFrameworkElement) Then
                oFrameworkElement.DataContext = Nothing
            End If

            Clear()
            For Each oDisplay In oDisplayList
                Me.Add(oDisplay)
            Next

            If Not IsNothing(oFrameworkElement) Then
                oFrameworkElement.DataContext = oDataContext
            End If

            OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End Sub
    End Class
    Public Property SelectedOrientationDisplay As Common.HighlightComboBox.HCBDisplay
        Get
            Dim sOrientation As String = [Enum].GetName(GetType(PageOrientation), m_DataObject.SelectedOrientation)
            Return (From iIndex In Enumerable.Range(0, m_PageOrientation.Count) Where m_PageOrientation(iIndex).Name = sOrientation Select m_PageOrientation(iIndex)).First
        End Get
        Set(value As Common.HighlightComboBox.HCBDisplay)
        End Set
    End Property
    Public Property SelectedSizeDisplay As Common.HighlightComboBox.HCBDisplay
        Get
            Dim sSize As String = [Enum].GetName(GetType(PageSize), m_DataObject.SelectedSize)
            Return (From iIndex In Enumerable.Range(0, m_PageSize.Count) Where m_PageSize(iIndex).Name = sSize Select m_PageSize(iIndex)).First
        End Get
        Set(value As Common.HighlightComboBox.HCBDisplay)
        End Set
    End Property
    Private Sub ConfigureWIAScannerDefault(ByVal oDevice As WIA.Device, ByVal fPageWidth As Single, ByVal fPageHeight As Single)
        ' configures scanner
        ' scan greyscale, 300 dpi
        ' page width and height are in inches
        Dim oWIADeviceManager As New WIA.DeviceManager
        For Each oDeviceInfo As WIA.DeviceInfo In oWIADeviceManager.DeviceInfos
            If oDeviceInfo.DeviceID = SelectedScannerSource.Item2 Then
                ' set greyscale scan intent
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_CurrentIntent, 2)

                ' set scan extent
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_HorizontalStartScanPixel, 0)
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_VerticalStartScanPixel, 0)
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_HorizontalScanSizePixels, fPageWidth * 300)
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_VerticalScanSizePixels, fPageHeight * 300)

                ' set scan resolution to 300 dpi
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_HorizontalResolution, RenderResolution300)
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_VerticalResolution, RenderResolution300)

                ' set brightness and contrast
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_Brightness, 0)
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_Contrast, 0)

                ' set bit depth to 8 bits
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_BitsPerPixel, 8)

                ' set automatic feeder
                ' set to 1 for ADF, 2 for flatbed
                SetWIAItemProperty(oDevice, WIAProperty.WIA_PROPERTY_DocumentHandlingSelect, 1)
            End If
        Next
    End Sub
    Private Sub SetWIAItemProperty(ByVal oDevice As WIA.Device, ByVal oPropertyID As WIAProperty, ByVal iPropertyValue As Integer, Optional ByVal oList As List(Of Integer) = Nothing)
        ' set device property
        Dim oItem As WIA.Item = oDevice.Items(1)
        For Each oProperty As WIA.Property In oItem.Properties
            If oProperty.PropertyID = oPropertyID Then
                oProperty.Value = iPropertyValue
                Exit For
            End If
        Next
    End Sub
    Private Shared WIAFormatBMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
    Private Shared WiaFormatPNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Private Shared WiaFormatGIF As String = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
    Private Shared WiaFormatJPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    Private Shared WiaFormatTIFF As String = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    Private Enum WIAError As Integer
        WIA_ERROR_GENERAL_ERROR = 1
        WIA_ERROR_PAPER_JAM = 2
        WIA_ERROR_PAPER_EMPTY = 3
        WIA_ERROR_PAPER_PROBLEM = 4
        WIA_ERROR_OFFLINE = 5
        WIA_ERROR_BUSY = 6
        WIA_ERROR_WARMING_UP = 7
        WIA_ERROR_USER_INTERVENTION = 8
        WIA_ERROR_ITEM_DELETED = 9
        WIA_ERROR_DEVICE_COMMUNICATION = 10
        WIA_ERROR_INVALID_COMMAND = 11
        WIA_ERROR_INCORRECT_HARDWARE_SETTING = 12
        WIA_ERROR_DEVICE_LOCKED = 13
        WIA_ERROR_EXCEPTION_IN_DRIVER = 14
        WIA_ERROR_INVALID_DRIVER_RESPONSE = 15
        WIA_S_NO_DEVICE_AVAILABLE = 21
    End Enum
    Private Enum WIACOmpression As Integer
        WIA_COMPRESSION_JPEG = 5
    End Enum
    Private Enum WIAProperty As Integer
        WIA_PROPERTY_CurrentIntent = 6146
        WIA_PROPERTY_HorizontalResolution = 6147
        WIA_PROPERTY_VerticalResolution = 6148
        WIA_PROPERTY_HorizontalStartScanPixel = 6149
        WIA_PROPERTY_VerticalStartScanPixel = 6150
        WIA_PROPERTY_HorizontalScanSizePixels = 6151
        WIA_PROPERTY_VerticalScanSizePixels = 6152
        WIA_PROPERTY_Brightness = 6154
        WIA_PROPERTY_Contrast = 6155
        WIA_PROPERTY_DocumentHandlingSelect = 3088
        WIA_PROPERTY_DocumentHandlingStatus = 3087
        WIA_PROPERTY_BitsPerPixel = 4104
        WIA_PROPERTY_Format = 4106
        WIA_PROPERTY_Compression = 4107
        WIA_PROPERTY_VALUE_Color = 1
        WIA_PROPERTY_VALUE_Gray = 2
        WIA_PROPERTY_VALUE_BlackAndWhite = 4
    End Enum
    Public Shared Function GetCornerstoneImage(ByVal fMargin As Single, ByVal fHorizontalResolution As Single, ByVal fVerticalResolution As Single, ByVal bFillBackground As Boolean) As System.Drawing.Bitmap
        ' draw three concentric circles with a border in pixel width
        ' resolution is in DPI
        Dim fPixelWidth As Single = 5 * CornerstoneSeperation * fHorizontalResolution / 2.54
        Dim fPixelHeight As Single = 5 * CornerstoneSeperation * fVerticalResolution / 2.54
        Dim fPixelSeperationHorizontal As Single = CornerstoneSeperation * fHorizontalResolution / 2.54
        Dim fPixelSeperationVertical As Single = CornerstoneSeperation * fVerticalResolution / 2.54
        Dim oTabletImage As New System.Drawing.Bitmap(fPixelWidth + (fMargin * 2), fPixelHeight + (fMargin * 2), System.Drawing.Imaging.PixelFormat.Format32bppArgb)

        Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oTabletImage)
            oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
            oGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
            oGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic

            If bFillBackground Then
                oGraphics.FillRectangle(System.Drawing.Brushes.White, -1, -1, oTabletImage.Width + 2, oTabletImage.Height + 2)
            End If

            Dim oPen As New System.Drawing.Pen(System.Drawing.Color.Black, 3 * fVerticalResolution / 300)
            oGraphics.DrawEllipse(oPen, fMargin, fMargin, fPixelWidth, fPixelHeight)
            oGraphics.DrawEllipse(oPen, fMargin + fPixelSeperationHorizontal, fMargin + fPixelSeperationVertical, 3 * fPixelSeperationHorizontal, 3 * fPixelSeperationVertical)
            oGraphics.DrawEllipse(oPen, fMargin + (fPixelSeperationHorizontal * 2), fMargin + (fPixelSeperationVertical * 2), fPixelSeperationHorizontal, fPixelSeperationVertical)
        End Using
        oTabletImage.SetResolution(fHorizontalResolution, fVerticalResolution)

        Return oTabletImage
    End Function
    Public Shared Function DrawTablet(ByRef oXGraphics As XGraphics, ByVal XCenterPoint As XPoint, ByVal sTabletContent As String, ByVal oFont As XFont, ByVal oStringFormat As XStringFormat, Optional ByVal oTabletImage As XImage = Nothing, Optional ByVal fSingleBlockWidth As Double = 0) As XSize
        ' draws a single tablet with the center point given
        ' the content is limited to two characters only
        ' draws tablet
        Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
        If IsNothing(oTabletImage) Then
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
            Const fTabletHeight As Double = 0.5
            If fSingleBlockWidth = 0 Then
                fSingleBlockWidth = BlockHeight.Point * 2
            End If
            Dim oCheckEmpty As Media.DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty)
            Dim oCheckEmptyBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckEmpty, Double.MaxValue, fSingleBlockWidth * fTabletHeight * RenderResolution300 / 72, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim oXCheckEmptyImage As XImage = PdfSharp.Drawing.XImage.FromGdiPlusImage(oCheckEmptyBitmap)
            Dim XCheckEmptyBitmapSize As New XSize(oXCheckEmptyImage.PointWidth, oXCheckEmptyImage.PointHeight)

            oXGraphics.DrawImage(oXCheckEmptyImage, XCenterPoint.X - XCheckEmptyBitmapSize.Width / 2, XCenterPoint.Y - XCheckEmptyBitmapSize.Height / 2)

            ' draws tablet text
            oXGraphics.DrawString(Left(sTabletContent, 2), oFont, XBrushes.Black, XCenterPoint.X, XCenterPoint.Y, oStringFormat)

            ' clean up
            oCheckEmptyBitmap.Dispose()

            oXGraphics.SmoothingMode = oSmoothingMode
            Return XCheckEmptyBitmapSize
        Else
            oXGraphics.SmoothingMode = XSmoothingMode.None
            Dim XTabletImageSize As New XSize(oTabletImage.PointWidth, oTabletImage.PointHeight)

            oXGraphics.DrawImage(oTabletImage, XCenterPoint.X - oTabletImage.PointWidth / 2, XCenterPoint.Y - oTabletImage.PointHeight / 2)

            oXGraphics.SmoothingMode = oSmoothingMode
            Return XTabletImageSize
        End If
    End Function
    Private Shared Sub SaveScreenFile(ByVal sScreenFileName As String, ByVal oDetectedImages As List(Of Tuple(Of String, Emgu.CV.Matrix(Of Byte))))
        ' saves mark list 
        Dim oMarkList As List(Of String) = (From oImage In oDetectedImages Select oImage.Item1).Distinct.ToList
        Dim iMaxCount As Integer = Aggregate sMark In oMarkList Into Max(oDetectedImages.ToArray.Count(Function(x) x.Item1 = sMark))

        Using oScreenMatrix As New Emgu.CV.Matrix(Of Byte)(((BoxSize + 1) * iMaxCount) + 1, ((BoxSize + 1) * oMarkList.Count) + 1, 3)
            oScreenMatrix.SetValue(New Emgu.CV.Structure.MCvScalar(0, 0, Byte.MaxValue))

            For i = 0 To oMarkList.Count - 1
                Dim sMark As String = oMarkList(i)
                Dim oMarkMatrices As List(Of Emgu.CV.Matrix(Of Byte)) = (From oImage In oDetectedImages Where oImage.Item1 = sMark Select oImage.Item2).ToList
                For j = 0 To oMarkMatrices.Count - 1
                    Using oBoxSubRect As Emgu.CV.Matrix(Of Byte) = oScreenMatrix.GetSubRect(New System.Drawing.Rectangle((i * (BoxSize + 1)) + 1, (j * (BoxSize + 1)) + 1, BoxSize, BoxSize))
                        Using oBoxImage As Emgu.CV.Matrix(Of Byte) = DetectorFunctions.ProcessMat(oMarkMatrices(j), BoxSize, True)
                            Emgu.CV.CvInvoke.CvtColor(oBoxImage, oBoxSubRect, Emgu.CV.CvEnum.ColorConversion.Gray2Bgr)
                        End Using
                    End Using
                Next
            Next
            Converter.MatToBitmap(oScreenMatrix.Mat, ScreenResolution096).Save(sScreenFileName, System.Drawing.Imaging.ImageFormat.Png)
        End Using
    End Sub
    Public Class PointCollection
        Inherits List(Of SinglePoint)
        Implements ICloneable, IDisposable

#Region "Functions"
        Public Shared Function GetDistance(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Double
            ' returns the distance between the two points
            Return oPoint1.Point.DistanceTo(oPoint2.Point)
        End Function
        Public Shared Function InRange(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Boolean
            ' checks to see if the distance between the two points is less than the sum of their radii
            Return GetDistance(oPoint1, oPoint2) < oPoint1.Radius + oPoint2.Radius
        End Function
        Public Shared Function InCircle(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Boolean
            ' checks to see if the distance between the two points is less than the larger circle's radius
            Return GetDistance(oPoint1, oPoint2) < Math.Max(oPoint1.Radius, oPoint2.Radius)
        End Function
        Public Shared Function GetNeighbouringPoints(ByRef oPoint As SinglePoint, ByVal oPointCircleList As PointCollection) As PointCollection
            ' gets a list of neighbouring points within range from the reference point
            ' fLowerLimitMaxRadius gives a lower limit to the max radius and helps small circles to join up
            Dim oNeighbouringPointsCollection As New PointCollection
            For Each oTestPoint As SinglePoint In oPointCircleList
                Dim fMinRadius As Single = Math.Min(oPoint.Radius, oTestPoint.Radius)
                Dim fDistance As Single = GetDistance(oTestPoint, oPoint)
                If fDistance > fMinRadius * 0.1 And InRange(oPoint, oTestPoint) Then
                    oNeighbouringPointsCollection.Add(oTestPoint.Clone)
                End If
            Next
            Return oNeighbouringPointsCollection
        End Function
        Public Shared Function GetAngle(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Double
            ' returns the angle of the line described by the two points
            Return Math.Atan2(oPoint2.Point.Y - oPoint1.Point.Y, oPoint2.Point.X - oPoint1.Point.X)
        End Function
        Public Shared Function MatchAngle(ByVal fAngle As Double, ByVal fMatchAngle As Double) As Double
            ' changes the input angle to ± 180 degrees of the match angle
            Dim fLowerAngle As Double = fMatchAngle - Math.PI
            Dim fUpperAngle As Double = fMatchAngle + Math.PI
            Dim fReturnAngle As Double = fAngle

            Dim iCount As Integer = 0
            Do Until (fReturnAngle >= fLowerAngle And fReturnAngle < fUpperAngle) OrElse iCount >= 10
                iCount += 1
                If fReturnAngle >= fUpperAngle Then
                    fReturnAngle -= CircleAngle
                ElseIf fReturnAngle < fLowerAngle Then
                    fReturnAngle += CircleAngle
                End If
            Loop

            Return fReturnAngle
        End Function
#End Region
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oPointCollection As New PointCollection
            For Each oPoint In Me
                oPointCollection.Add(oPoint.Clone)
            Next
            Return oPointCollection
        End Function
        Public Overrides Function Equals(obj As [Object]) As Boolean
            ' Check for null values and compare run-time types.
            If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                Return False
            End If

            Dim oPointCollection As PointCollection = DirectCast(obj, PointCollection)
            Return GetHashCode.Equals(oPointCollection.GetHashCode)
        End Function
        Public Overrides Function GetHashCode() As Integer
            Dim iHashCode As Long = Aggregate oPoint As SinglePoint In Me Into Sum(CLng(oPoint.GetHashCode))
            Return iHashCode Mod Integer.MaxValue
        End Function
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    Clear()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            End If
            disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
        Structure Segment
            Implements ICloneable

            Public Segment As AForge.Math.Geometry.LineSegment

            Sub New(ByVal oSegment As AForge.Math.Geometry.LineSegment)
                Segment = oSegment
            End Sub
            Sub New(ByVal oStart As SinglePoint, ByVal oEnd As SinglePoint)
                Segment = New AForge.Math.Geometry.LineSegment(oStart.Point, oEnd.Point)
            End Sub
            Public ReadOnly Property Angle As Double
                ' returns the angle between the line segment and the horizontal line in radians
                Get
                    Dim oSegmentLine As AForge.Math.Geometry.Line = AForge.Math.Geometry.Line.FromPoints(Segment.Start, Segment.End)
                    If oSegmentLine.IsHorizontal Then
                        Return 0
                    ElseIf oSegmentLine.IsVertical Then
                        Return Math.PI / 2
                    Else
                        Dim oHorizontalLine As AForge.Math.Geometry.Line = AForge.Math.Geometry.Line.FromPoints(New AForge.Point(0, 0), New AForge.Point(1, 0))
                        Dim fMinimumAngle As Double = oSegmentLine.GetAngleBetweenLines(oHorizontalLine) * Math.PI / 180
                        If Math.Sign(Segment.End.X - Segment.Start.X) = Math.Sign(Segment.End.Y - Segment.Start.Y) Then
                            ' forward slope
                            Return fMinimumAngle
                        Else
                            ' backwards slope
                            Return Math.PI - fMinimumAngle
                        End If
                    End If
                End Get
            End Property
            Public ReadOnly Property Length As Double
                Get
                    Return Segment.Length
                End Get
            End Property
            Public Overrides Function Equals(obj As [Object]) As Boolean
                ' Check for null values and compare run-time types.
                If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                    Return False
                End If

                ' reverses the segment to check equality
                Dim oSegment As Segment = DirectCast(obj, Segment)
                Return GetHashCode.Equals(oSegment.GetHashCode) Or (New AForge.Math.Geometry.LineSegment(Segment.End, Segment.Start)).GetHashCode.Equals(oSegment.GetHashCode)
            End Function
            Public Overrides Function GetHashCode() As Integer
                Return Segment.GetHashCode
            End Function
            Public Function Clone() As Object Implements ICloneable.Clone
                Return New Segment(New AForge.Math.Geometry.LineSegment(Segment.Start, Segment.End))
            End Function
        End Structure
        Structure SinglePoint
            Implements ICloneable

            Public Radius As Single
            Public Point As AForge.Point
            Public Selected As Boolean

            Sub New(ByVal fRadius As Single, ByVal oSinglePoint As AForge.Point, ByVal bSelected As Boolean)
                Radius = fRadius
                Point = oSinglePoint
                Selected = bSelected
            End Sub
            Sub New(ByVal oSinglePoint As SinglePoint)
                Radius = oSinglePoint.Radius
                Point = oSinglePoint.Point
                Selected = oSinglePoint.Selected
            End Sub
            Sub New(ByVal fRadius As Single?, ByVal oSinglePoint As AForge.Point?, ByVal bSelected As Boolean?, ByVal oPoint As SinglePoint)
                If IsNothing(fRadius) Then
                    Radius = oPoint.Radius
                Else
                    Radius = fRadius
                End If
                If IsNothing(oSinglePoint) Then
                    Point = oPoint.Point
                Else
                    Point = oSinglePoint
                End If
                If IsNothing(bSelected) Then
                    Selected = oPoint.Selected
                Else
                    Selected = bSelected
                End If
            End Sub
            Public ReadOnly Property DrawingPoint As System.Drawing.Point
                Get
                    Return New System.Drawing.Point(Point.X, Point.Y)
                End Get
            End Property
            Public Overrides Function Equals(obj As [Object]) As Boolean
                ' Check for null values and compare run-time types.
                If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                    Return False
                End If

                Dim oPoint As SinglePoint = DirectCast(obj, SinglePoint)
                Return GetHashCode.Equals(oPoint.GetHashCode)
            End Function
            Public Overrides Function GetHashCode() As Integer
                Return Radius.GetHashCode Xor Point.GetHashCode Xor Selected.GetHashCode
            End Function
            Public Function Clone() As Object Implements ICloneable.Clone
                Return New SinglePoint(Radius, Point, Selected)
            End Function
        End Structure
    End Class
    Class PointList
        Implements ICloneable, IDisposable

        Private m_PointList As List(Of SinglePoint)
        Private m_Scale As Integer
        Private Shared CircleAngle As Double = 2 * Math.PI

        Sub New()
            m_PointList = New List(Of SinglePoint)
        End Sub
        Sub New(ByVal oPointList As List(Of SinglePoint))
            m_PointList = oPointList
        End Sub
        Sub New(ByVal oPointArray As SinglePoint())
            m_PointList = New List(Of SinglePoint)
            m_PointList.AddRange(oPointArray)
        End Sub
        Public Property PointCollection As List(Of SinglePoint)
            Get
                Return m_PointList
            End Get
            Set(value As List(Of SinglePoint))
                m_PointList = value
            End Set
        End Property
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oPointList As New Common.Common.PointCollection
            With oPointList
                .Clear()
                For Each oPoint As SinglePoint In Common.Common.PointCollection
                    .Add(New SinglePoint(oPoint))
                Next
            End With
            Return oPointList
        End Function
        Public Function InList(ByVal oPoint As SinglePoint) As Boolean
            ' checks to see if the point is in the list
            Dim bReturn As Boolean = False
            For Each oCurrentPoint As SinglePoint In m_PointList
                If oPoint.Equals(oCurrentPoint) Then
                    bReturn = True
                    Exit For
                End If
            Next

            Return bReturn
        End Function
        Public Function BoundingBox() As System.Drawing.Rectangle
            Dim oRectangle As Rect = Rect.Empty
            For Each oPoint In m_PointList
                Dim oPointRectangle = New Rect(oPoint.Point.X - oPoint.Radius, oPoint.Point.Y - oPoint.Radius, oPoint.Radius * 2, oPoint.Radius * 2)
                If oRectangle.IsEmpty Then
                    oRectangle = oPointRectangle
                Else
                    oRectangle.Union(oPointRectangle)
                End If
            Next
            Return New System.Drawing.Rectangle(oRectangle.X, oRectangle.Y, oRectangle.Width, oRectangle.Height)
        End Function
        Public Shared Function GetAngle(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Double
            ' returns the angle of the line described by the two points
            Return Math.Atan2(oPoint2.Point.Y - oPoint1.Point.Y, oPoint2.Point.X - oPoint1.Point.X)
        End Function
        Public Shared Function MatchAngle(ByVal fAngle As Double, ByVal fMatchAngle As Double) As Double
            ' changes the input angle to ± 180 degrees of the match angle
            Dim fLowerAngle As Double = fMatchAngle - Math.PI
            Dim fUpperAngle As Double = fMatchAngle + Math.PI
            Dim fReturnAngle As Double = fAngle

            Dim iCount As Integer = 0
            Do Until (fReturnAngle >= fLowerAngle And fReturnAngle < fUpperAngle) OrElse iCount >= 10
                iCount += 1
                If fReturnAngle >= fUpperAngle Then
                    fReturnAngle -= CircleAngle
                ElseIf fReturnAngle < fLowerAngle Then
                    fReturnAngle += CircleAngle
                End If
            Loop

            Return fReturnAngle
        End Function
        Public Shared Function CompareAngle(ByVal fAngle As Double, ByVal fMatchAngle As Double, ByVal fRange As Double) As Boolean
            ' compares the two angles and returns true if they are within fRange of each other
            Return CompareAngleValue(fAngle, fMatchAngle) <= fRange
        End Function
        Public Shared Function CompareAngleValue(ByVal fAngle As Double, ByVal fMatchAngle As Double) As Double
            ' compares the two angles and returns the difference
            Return Math.Abs(fMatchAngle - MatchAngle(fAngle, fMatchAngle))
        End Function
        Public Shared Function GetDistance(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Double
            ' returns the distance between the two points
            Return Math.Sqrt((oPoint2.Point.Y - oPoint1.Point.Y) * (oPoint2.Point.Y - oPoint1.Point.Y) + (oPoint2.Point.X - oPoint1.Point.X) * (oPoint2.Point.X - oPoint1.Point.X))
        End Function
        Public Shared Function InRange(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Boolean
            ' checks to see if the distance between the two points is less than the sum of their radii
            Return GetDistance(oPoint1, oPoint2) < oPoint1.Radius + oPoint2.Radius
        End Function
        Public Shared Function InCircle(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Boolean
            ' checks to see if the distance between the two points is less than the larger circle's radius
            Return GetDistance(oPoint1, oPoint2) < Math.Max(oPoint1.Radius, oPoint2.Radius)
        End Function
        Public Shared Function GetNeighbouringPoints(ByRef oPoint As SinglePoint, ByVal oPointCircleList As Common.Common.PointCollection, ByVal bAscending As Boolean, ByVal fLowerLimitMaxRadius As Single) As Common.Common.PointCollection
            ' gets a list of neighbouring points within range from the reference point
            ' fLowerLimitMaxRadius gives a lower limit to the max radius and helps small circles to join up
            Dim oNeighbouringPointsList As New Common.Common.PointCollection
            For Each oTestPoint As SinglePoint In oPointCircleList
                Dim fMinRadius As Single = Math.Min(oPoint.Radius, oTestPoint.Radius)
                Dim fDistance As Single = GetDistance(oTestPoint, oPoint)
                If fDistance > fMinRadius * 0.1 And (InRange(oPoint, oTestPoint) Or fDistance < fLowerLimitMaxRadius) Then
                    oNeighbouringPointsList.Add(New SinglePoint(oTestPoint.Radius, oTestPoint.Point, oTestPoint.Count, fDistance, oTestPoint.GUID))
                End If
            Next

            ' reorder by distance
            If bAscending Then
                oNeighbouringPointsList = (From oSinglePoint In oNeighbouringPointsList Order By oSinglePoint.Distance Ascending Select oSinglePoint).ToList
            Else
                oNeighbouringPointsList = (From oSinglePoint In oNeighbouringPointsList Order By oSinglePoint.Distance Descending Select oSinglePoint).ToList
            End If

            Return oNeighbouringPointsList
        End Function
        Public Shared Function GetPointIndex(ByRef oPoint As SinglePoint, ByVal oPointList As Common.Common.PointCollection) As Integer
            ' returns the index of the point based on GUID equivalence
            Dim iReturnIndex As Integer = -1
            For i = 0 To oPointList.Count - 1
                If oPoint.Equals(oPointList(i)) Then
                    iReturnIndex = i
                    Exit For
                End If
            Next
            Return iReturnIndex
        End Function
        Public Shared Function OutlierPoints(ByVal oStartPoint As SinglePoint, ByVal oTestPoints As List(Of SinglePoint), ByVal oEndPoint As SinglePoint, ByVal fEpsilon As Double) As List(Of SinglePoint)
            ' returns true if the test point is an outlier
            Dim oReturnPoints As New List(Of SinglePoint)

            Using oContour As New Emgu.CV.Util.VectorOfPoint()
                oContour.Push({oStartPoint.Point})
                For Each oPoint As SinglePoint In oTestPoints
                    oContour.Push({oPoint.Point})
                Next
                oContour.Push({oEndPoint.Point})

                Using oSimplifiedContour As New Emgu.CV.Util.VectorOfPoint()
                    Emgu.CV.CvInvoke.ApproxPolyDP(oContour, oSimplifiedContour, fEpsilon, True)
                    Dim oExclusionPoints As List(Of System.Drawing.Point) = oSimplifiedContour.ToArray.ToList

                    For Each oPoint As SinglePoint In oTestPoints
                        If oExclusionPoints.Contains(oPoint.Point) Then
                            oReturnPoints.Add(oPoint)
                        End If
                    Next
                End Using
            End Using

            Return oReturnPoints
        End Function
        Public Overrides Function Equals(obj As Object) As Boolean
            ' Check for null values and compare run-time types.
            If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                Return False
            End If

            Dim oPointList As Common.Common.PointCollection = obj
            Return GetHashCode.Equals(oPointList.GetHashCode)
        End Function
        Public Overrides Function GetHashCode() As Integer
            Dim iHashCode As Long = Aggregate oPoint As SinglePoint In m_PointList Into Sum(CLng(oPoint.GetHashCode))
            Return iHashCode Mod Integer.MaxValue
        End Function
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    m_PointList.Clear()
                    m_PointList = Nothing
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            End If
            disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
        Structure SinglePoint
            Private m_Radius As Single
            Private m_Point As System.Drawing.Point
            Private m_Count As Integer
            Private m_Distance As Single
            Private m_GUID As Guid

            Sub New(ByVal fRadius As Single, ByVal oPoint As System.Drawing.Point, ByVal iCount As Integer, ByVal fDistance As Single)
                m_Radius = fRadius
                m_Point = oPoint
                m_Count = iCount
                m_Distance = fDistance
                m_GUID = GUID.NewGuid
            End Sub
            Sub New(ByVal fRadius As Single, ByVal oPoint As System.Drawing.Point, ByVal iCount As Integer, ByVal fDistance As Single, ByVal oGUID As Guid)
                m_Radius = fRadius
                m_Point = oPoint
                m_Count = iCount
                m_Distance = fDistance
                m_GUID = oGUID
            End Sub
            Sub New(ByVal oSinglePoint As SinglePoint)
                m_Radius = oSinglePoint.Radius
                m_Point = oSinglePoint.Point
                m_Count = oSinglePoint.Count
                m_Distance = oSinglePoint.Distance
                m_GUID = oSinglePoint.m_GUID
            End Sub
            Sub New(ByVal fRadius As Single?, ByVal oPoint As System.Drawing.Point?, ByVal iCount As Integer?, ByVal fDistance As Single?, ByVal oGUID As Guid?, ByVal oSinglePoint As SinglePoint)
                If IsNothing(fRadius) Then
                    m_Radius = oSinglePoint.Radius
                Else
                    m_Radius = fRadius
                End If
                If IsNothing(oPoint) Then
                    m_Point = oSinglePoint.Point
                Else
                    m_Point = oPoint
                End If
                If IsNothing(iCount) Then
                    m_Count = oSinglePoint.Count
                Else
                    m_Count = iCount
                End If
                If IsNothing(fDistance) Then
                    m_Distance = oSinglePoint.Distance
                Else
                    m_Distance = fDistance
                End If
                If IsNothing(oGUID) Then
                    m_GUID = oSinglePoint.m_GUID
                Else
                    m_GUID = oGUID
                End If
            End Sub
            Public ReadOnly Property Radius As Single
                Get
                    Return m_Radius
                End Get
            End Property
            Public ReadOnly Property Point As System.Drawing.Point
                Get
                    Return m_Point
                End Get
            End Property
            Public ReadOnly Property Count As Integer
                Get
                    Return m_Count
                End Get
            End Property
            Public ReadOnly Property Distance As Single
                Get
                    Return m_Distance
                End Get
            End Property
            Public ReadOnly Property GUID As Guid
                Get
                    Return m_GUID
                End Get
            End Property
            Public ReadOnly Property AForgePoint As AForge.Point
                Get
                    Return New AForge.Point(m_Point.X, m_Point.Y)
                End Get
            End Property
            Public Overrides Function Equals(obj As Object) As Boolean
                ' Check for null values and compare run-time types.
                If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                    Return False
                End If

                Dim oSinglePoint As SinglePoint = obj
                Return m_GUID.Equals(oSinglePoint.m_GUID)
            End Function
            Public Overrides Function GetHashCode() As Integer
                Return m_GUID.GetHashCode
            End Function
        End Structure
    End Class
    Private Sub SetScanner()
        ' sets default scanner which is the first WIA scanner, then the first TWAIN scanner
        Dim oScannerSources As List(Of Tuple(Of Twain32Enumerations.ScannerSource, String, String)) = GetScannerSources()

        If oScannerSources.Count > 0 Then
            If oScannerSources.First.Item1 = Twain32Enumerations.ScannerSource.TWAIN Then
                ' select twain source
                m_CommonScanner.SelectScannerSource(oScannerSources.First.Item2)
            End If
            SelectedScannerSource = oScannerSources.First
        Else
            m_CommonScanner.SelectScannerSource(String.Empty)
            SelectedScannerSource = Nothing
        End If
    End Sub
    Private Shared Sub ExportData(ByVal sFileName As String)
        ' exports data to spreadsheet
        Const WorksheetName As String = "Data"
        Const CommentsName As String = "Comments"
        Const StatsName As String = "Statistics"
        Dim oScannerCollection As ScannerCollection = Root.GridMain.Resources("scannercollection")

        ' group by subject and then by ItemNumberSort
        'Dim oSubjectList As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore Select oFieldCollection.SubjectName).ToList
        'Dim oGroupedFieldList As List(Of Tuple(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, List(Of Field))(sSubjectName, (From oField As Field In oFieldCollection Where oField.SubjectName = sSubjectName Order By oField.ItemNumberSort Ascending Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        'Dim oItemList As List(Of Tuple(Of Enumerations.FieldTypeEnum, String, String, String, String)) = (From oTuple As Tuple(Of Enumerations.FieldTypeEnum, String, String, String, String) In (From oField As Field In oFieldCollection Order By oField.ItemNumberSort Ascending Select New Tuple(Of Enumerations.FieldTypeEnum, String, String, String, String)(oField.FieldType, oField.ItemNumberText, oField.SubjectName, oField.ItemNumber, oField.ItemSubNumber) Distinct)).ToList

        'Dim oItemNumberListChoice As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.Choice Select oField.Numbering Distinct).ToList
        'Dim oFieldDictionaryChoice As Dictionary(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field)))(sSubjectName, (From sItemNumberText As String In oItemNumberListChoice Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sItemNumberText, (From oField As Field In oGroupedFieldList(sSubjectName) Where oField.ItemNumberText = sItemNumberText AndAlso oField.FieldType = Enumerations.FieldTypeEnum.Choice Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value))).ToDictionary(Function(x) x.Key, Function(x) x.Value)

        'Dim oItemNumberListChoiceVertical As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.ChoiceVertical Select oField.Numbering Distinct).ToList
        'Dim oFieldDictionaryChoiceVertical As Dictionary(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field)))(sSubjectName, (From sItemNumberText As String In oItemNumberListChoiceVertical Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sItemNumberText, (From oField As Field In oGroupedFieldList(sSubjectName) Where oField.ItemNumberText = sItemNumberText AndAlso oField.FieldType = Enumerations.FieldTypeEnum.ChoiceVertical Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value))).ToDictionary(Function(x) x.Key, Function(x) x.Value)

        'Dim oItemNumberListChoiceVerticalMCQ As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ Select oField.Numbering Distinct).ToList
        'Dim oFieldDictionaryChoiceVerticalMCQ As Dictionary(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field)))(sSubjectName, (From sItemNumberText As String In oItemNumberListChoiceVerticalMCQ Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sItemNumberText, (From oField As Field In oGroupedFieldList(sSubjectName) Where oField.ItemNumberText = sItemNumberText AndAlso oField.FieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value))).ToDictionary(Function(x) x.Key, Function(x) x.Value)

        'Dim oItemNumberListHandwriting As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.Handwriting Select oField.Numbering Distinct).ToList
        'Dim oFieldDictionaryHandwriting As Dictionary(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field)))(sSubjectName, (From sItemNumberText As String In oItemNumberListHandwriting Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sItemNumberText, (From oField As Field In oGroupedFieldList(sSubjectName) Where oField.ItemNumberText = sItemNumberText AndAlso oField.FieldType = Enumerations.FieldTypeEnum.Handwriting Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value))).ToDictionary(Function(x) x.Key, Function(x) x.Value)

        'Dim oItemNumberListBoxChoice As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.BoxChoice Select oField.Numbering Distinct).ToList
        'Dim oFieldDictionaryBoxChoice As Dictionary(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field)))(sSubjectName, (From sItemNumberText As String In oItemNumberListBoxChoice Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sItemNumberText, (From oField As Field In oGroupedFieldList(sSubjectName) Where oField.ItemNumberText = sItemNumberText AndAlso oField.FieldType = Enumerations.FieldTypeEnum.BoxChoice Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value))).ToDictionary(Function(x) x.Key, Function(x) x.Value)

        'Dim oItemNumberListFree As List(Of String) = (From oFieldCollection As FieldDocumentStore.FieldCollection In oScannerCollection.FieldDocumentStore.FieldCollectionStore From oField As FieldDocumentStore.Field In oFieldCollection.Fields Where oField.FieldType = Enumerations.FieldTypeEnum.Free Select oField.Numbering Distinct).ToList
        'Dim oFieldDictionaryFree As Dictionary(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field))) = (From sSubjectName As String In oSubjectList Select New KeyValuePair(Of String, Dictionary(Of String, List(Of FieldDocumentStore.Field)))(sSubjectName, (From sItemNumberText As String In oItemNumberListFree Select New KeyValuePair(Of String, List(Of FieldDocumentStore.Field))(sItemNumberText, (From oField As Field In oGroupedFieldList(sSubjectName) Where oField.ItemNumberText = sItemNumberText AndAlso oField.FieldType = Enumerations.FieldTypeEnum.Free Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value))).ToDictionary(Function(x) x.Key, Function(x) x.Value)


        Dim oModifiedItemList As New List(Of Tuple(Of Enumerations.FieldTypeEnum, String, String, String))
        For Each oItem As Tuple(Of Enumerations.FieldTypeEnum, String, String, String, String) In oItemList
            Dim oChoiceFound As List(Of Tuple(Of String, String)) = (From sSubjectName As String In oFieldDictionaryChoice.Keys From sItemNumberText As String In oFieldDictionaryChoice(sSubjectName).Keys Where sSubjectName = oItem.Item3 And sItemNumberText = oItem.Item2 Select New Tuple(Of String, String)(sSubjectName, sItemNumberText) Distinct).ToList
            Dim oChoiceVerticalFound As List(Of Tuple(Of String, String)) = (From sSubjectName As String In oFieldDictionaryChoiceVertical.Keys From sItemNumberText As String In oFieldDictionaryChoiceVertical(sSubjectName).Keys Where sSubjectName = oItem.Item3 And sItemNumberText = oItem.Item2 Select New Tuple(Of String, String)(sSubjectName, sItemNumberText)).ToList
            Dim oChoiceVerticalMCQFound As List(Of Tuple(Of String, String)) = (From sSubjectName As String In oFieldDictionaryChoiceVerticalMCQ.Keys From sItemNumberText As String In oFieldDictionaryChoiceVerticalMCQ(sSubjectName).Keys Where sSubjectName = oItem.Item3 And sItemNumberText = oItem.Item2 Select New Tuple(Of String, String)(sSubjectName, sItemNumberText)).ToList
            Dim oHandwritingFound As List(Of Tuple(Of String, String)) = (From sSubjectName As String In oFieldDictionaryHandwriting.Keys From sItemNumberText As String In oFieldDictionaryHandwriting(sSubjectName).Keys Where sSubjectName = oItem.Item3 And sItemNumberText = oItem.Item2 Select New Tuple(Of String, String)(sSubjectName, sItemNumberText)).ToList
            Dim oBoxChoiceFound As List(Of Tuple(Of String, String)) = (From sSubjectName As String In oFieldDictionaryBoxChoice.Keys From sItemNumberText As String In oFieldDictionaryBoxChoice(sSubjectName).Keys Where sSubjectName = oItem.Item3 And sItemNumberText = oItem.Item2 Select New Tuple(Of String, String)(sSubjectName, sItemNumberText)).ToList
            Dim oFreeFound As List(Of Tuple(Of String, String)) = (From sSubjectName As String In oFieldDictionaryFree.Keys From sItemNumberText As String In oFieldDictionaryFree(sSubjectName).Keys Where sSubjectName = oItem.Item3 And sItemNumberText = oItem.Item2 Select New Tuple(Of String, String)(sSubjectName, sItemNumberText)).ToList

            Select Case oItem.Item1
                Case Enumerations.FieldTypeEnum.Choice
                    If oChoiceFound.Count > 0 Then
                        oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(Enumerations.FieldTypeEnum.Choice, oFieldDictionaryChoice(oChoiceFound(0).Item1)(oChoiceFound(0).Item2)(0).ItemNumberText + "C", oFieldDictionaryChoice(oChoiceFound(0).Item1)(oChoiceFound(0).Item2)(0).ItemNumberText, String.Empty))
                    End If
                Case Enumerations.FieldTypeEnum.ChoiceVertical
                    If oChoiceVerticalFound.Count > 0 Then
                        oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(Enumerations.FieldTypeEnum.ChoiceVertical, oFieldDictionaryChoiceVertical(oChoiceVerticalFound(0).Item1)(oChoiceVerticalFound(0).Item2)(0).ItemNumberText + "CV", oFieldDictionaryChoiceVertical(oChoiceVerticalFound(0).Item1)(oChoiceVerticalFound(0).Item2)(0).ItemNumberText, String.Empty))
                    End If
                Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    If oChoiceVerticalMCQFound.Count > 0 Then
                        oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(Enumerations.FieldTypeEnum.ChoiceVerticalMCQ, oFieldDictionaryChoiceVerticalMCQ(oChoiceVerticalMCQFound(0).Item1)(oChoiceVerticalMCQFound(0).Item2)(0).ItemNumberText + "MCQ", oFieldDictionaryChoiceVerticalMCQ(oChoiceVerticalMCQFound(0).Item1)(oChoiceVerticalMCQFound(0).Item2)(0).ItemNumberText, String.Empty))
                    End If
                Case Enumerations.FieldTypeEnum.Handwriting
                    If oHandwritingFound.Count > 0 Then
                        oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(Enumerations.FieldTypeEnum.Handwriting, oFieldDictionaryHandwriting(oHandwritingFound(0).Item1)(oHandwritingFound(0).Item2)(0).ItemNumberText + "H", oFieldDictionaryHandwriting(oHandwritingFound(0).Item1)(oHandwritingFound(0).Item2)(0).ItemNumberText, oFieldDictionaryHandwriting(oHandwritingFound(0).Item1)(oHandwritingFound(0).Item2)(0).ItemSubNumber))
                    End If
                Case Enumerations.FieldTypeEnum.BoxChoice
                    If oBoxChoiceFound.Count > 0 Then
                        oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(Enumerations.FieldTypeEnum.BoxChoice, oFieldDictionaryBoxChoice(oBoxChoiceFound(0).Item1)(oBoxChoiceFound(0).Item2)(0).ItemNumberText + "B", oFieldDictionaryBoxChoice(oBoxChoiceFound(0).Item1)(oBoxChoiceFound(0).Item2)(0).ItemNumberText, oFieldDictionaryBoxChoice(oBoxChoiceFound(0).Item1)(oBoxChoiceFound(0).Item2)(0).ItemSubNumber))
                    End If
                Case Enumerations.FieldTypeEnum.Free
                    If oFreeFound.Count > 0 Then
                        oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(Enumerations.FieldTypeEnum.Free, oFieldDictionaryFree(oFreeFound(0).Item1)(oFreeFound(0).Item2)(0).ItemNumberText + "F", oFieldDictionaryFree(oFreeFound(0).Item1)(oFreeFound(0).Item2)(0).ItemNumberText, oFieldDictionaryFree(oFreeFound(0).Item1)(oFreeFound(0).Item2)(0).ItemSubNumber))
                    End If
                Case Else
                    oModifiedItemList.Add(New Tuple(Of Enumerations.FieldTypeEnum, String, String, String)(oItem.Item1, oItem.Item2, oItem.Item4, oItem.Item5))
            End Select
        Next
        oModifiedItemList = oModifiedItemList.Distinct.ToList

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
        For j = 0 To oModifiedItemList.Count - 1
            Select Case oModifiedItemList(j).Item1
                Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical
                    oWorksheet.Cell(1, 3 + j).Value = "'" + oModifiedItemList(j).Item2
                Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    oWorksheet.Cell(1, 3 + j).Value = "'" + oModifiedItemList(j).Item2
                Case Enumerations.FieldTypeEnum.Handwriting
                    oWorksheet.Cell(1, 3 + j).Value = "'" + oModifiedItemList(j).Item2
                Case Enumerations.FieldTypeEnum.BoxChoice
                    oWorksheet.Cell(1, 3 + j).Value = "'" + oModifiedItemList(j).Item2
                Case Enumerations.FieldTypeEnum.Free
                    oWorksheet.Cell(1, 3 + j).Value = "'" + oModifiedItemList(j).Item2
            End Select
            oWorksheet.Cell(1, 3 + j).Style.Font.Bold = True

            If oModifiedItemList(j).Item1 = Enumerations.FieldTypeEnum.Choice Or oModifiedItemList(j).Item1 = Enumerations.FieldTypeEnum.ChoiceVertical Or (oModifiedItemList(j).Item1 = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ) Then
                Select Case oModifiedItemList(j).Item1
                    Case Enumerations.FieldTypeEnum.Choice
                        oComments.Add(New Tuple(Of String, Boolean)("Choice Field: " + oModifiedItemList(j).Item2, True))
                    Case Enumerations.FieldTypeEnum.ChoiceVertical
                        oComments.Add(New Tuple(Of String, Boolean)("Choice Vertical Field: " + oModifiedItemList(j).Item2, True))
                    Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        oComments.Add(New Tuple(Of String, Boolean)("Choice MCQ Field: " + oModifiedItemList(j).Item3, True))
                End Select

                If oSubjectList.Count > 0 Then
                    Dim iCurrentj As Integer = j
                    Dim iCurrentIndex As Integer = 0
                    Dim oSelectedFields As List(Of Field) = (From oField As Field In oGroupedFieldList(oSubjectList(0)) Where oField.ItemNumberText = oModifiedItemList(iCurrentj).Item3 AndAlso oField.FieldType = oModifiedItemList(iCurrentj).Item1).ToList
                    For Each oField As Field In oSelectedFields
                        Select Case oField.FieldType
                            Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical
                                Dim iMaxCount As Integer = Math.Min(Math.Max(oField.TabletDescriptionTop.Count, oField.TabletDescriptionBottom.Count), oField.ImageCount)
                                For k = 0 To iMaxCount - 1
                                    Dim sTabletDescriptionTop As String = If(IsNothing(oField.TabletDescriptionTop(k).Item2), String.Empty, oField.TabletDescriptionTop(k).Item2)
                                    Dim sTabletDescriptionBottom As String = If(IsNothing(oField.TabletDescriptionBottom(k).Item2), String.Empty, oField.TabletDescriptionBottom(k).Item2)

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
                                For k = 0 To oField.TabletDescriptionMCQ.Count - 1
                                    Dim sContentText As String = String.Empty
                                    Select Case oField.TabletContent
                                        Case Enumerations.TabletContentEnum.Number
                                            sContentText = (k + If(oField.TabletStart = -2, 0, oField.TabletStart + 1)).ToString
                                        Case Enumerations.TabletContentEnum.Letter
                                            sContentText = Converter.ConvertNumberToLetter(k + Math.Max(oField.TabletStart, 0), True)
                                    End Select

                                    Dim sDescription As String = String.Join(" ", (From oElement As ElementStruc In oField.TabletDescriptionMCQ(k).Item4 Select oElement.Text))
                                    Dim sTabletDescription As String = sContentText + ": " + sDescription
                                    oComments.Add(New Tuple(Of String, Boolean)(sTabletDescription, False))
                                Next
                        End Select
                    Next
                End If
                oComments.Add(New Tuple(Of String, Boolean)(String.Empty, False))
            End If
        Next
        For i = 0 To oSubjectList.Count - 1
            oWorksheet.Cell(2 + i, 1).Value = (1 + i).ToString
            oWorksheet.Cell(2 + i, 2).Value = oSubjectList(i)

            For j = 0 To oModifiedItemList.Count - 1
                Dim iCurrenti As Integer = i
                Dim iCurrentj As Integer = j
                Select Case oModifiedItemList(j).Item1
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        Dim oSelectedField As Field = (From oField As Field In oGroupedFieldList(oSubjectList(iCurrenti)) Where oField.ItemNumberText = oModifiedItemList(iCurrentj).Item3 AndAlso oField.FieldType = oModifiedItemList(iCurrentj).Item1).First
                        If oSelectedField.TabletSingleChoiceOnly Then
                            Dim oMarkTrue As List(Of Integer) = (From iIndex As Integer In Enumerable.Range(0, oSelectedField.MarkCount) Where oSelectedField.MarkChoice2(iIndex) Select iIndex).ToList
                            If oMarkTrue.Count > 0 Then
                                Dim sContentText As String = String.Empty
                                Select Case oSelectedField.TabletContent
                                    Case Enumerations.TabletContentEnum.Letter
                                        sContentText = Converter.ConvertNumberToLetter(oMarkTrue.First + Math.Max(oSelectedField.TabletStart, 0), True)
                                    Case Else
                                        sContentText = (oMarkTrue.First + If(oSelectedField.TabletStart = -2, 0, oSelectedField.TabletStart + 1)).ToString
                                End Select
                                oWorksheet.Cell(2 + i, 3 + j).Value = sContentText
                            End If
                        Else
                            oWorksheet.Cell(2 + i, 3 + j).Value = oSelectedField.MarkChoiceCombined2
                        End If
                    Case Enumerations.FieldTypeEnum.Handwriting
                        Dim oSelectedField As Field = (From oField As Field In oGroupedFieldList(oSubjectList(iCurrenti)) Where oField.ItemNumberText = oModifiedItemList(iCurrentj).Item3 And oField.ItemSubNumber = oModifiedItemList(iCurrentj).Item4).First
                        oWorksheet.Cell(2 + i, 3 + j).Value = oSelectedField.MarkHandwritingCombined2
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        Dim oSelectedField As Field = (From oField As Field In oGroupedFieldList(oSubjectList(iCurrenti)) Where oField.ItemNumberText = oModifiedItemList(iCurrentj).Item3 And oField.ItemSubNumber = oModifiedItemList(iCurrentj).Item4).First
                        oWorksheet.Cell(2 + i, 3 + j).Value = oSelectedField.MarkBoxChoiceCombined2
                    Case Enumerations.FieldTypeEnum.Free
                        Dim oSelectedField As Field = (From oField As Field In oGroupedFieldList(oSubjectList(iCurrenti)) Where oField.ItemNumberText = oModifiedItemList(iCurrentj).Item3 And oField.ItemSubNumber = oModifiedItemList(iCurrentj).Item4).First
                        oWorksheet.Cell(2 + i, 3 + j).Value = oSelectedField.MarkFree2
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
        Dim oCurrentFieldList As New List(Of Field)
        Dim oCriticalFieldList As New List(Of Field)
        Dim oNonCriticalFieldList As New List(Of Field)
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
            oCurrentFieldList = (From oField In oFieldCollection Where oField.FieldType = oCurrentFieldType Select oField).ToList
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
                    Dim oFieldList As List(Of Field) = (From oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oBoxList(oField.GUID).Count) Where (Not oField.Images(oBoxList(oField.GUID)(iIndex)).Item5) AndAlso ((Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3)) Or (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3))) Select oField).ToList
                    Dim oSubjectNameList As List(Of String) = (From oField In oFieldList Select oField.SubjectName Distinct).ToList
                    Dim oFieldDictionary As Dictionary(Of String, List(Of Field)) = (From sSubjectName As String In oSubjectNameList Select New KeyValuePair(Of String, List(Of Field))(sSubjectName, (From oField In oFieldList Where oField.SubjectName = sSubjectName Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    For Each sSubjectName As String In oFieldDictionary.Keys
                        iCurrentRow += 1
                        Dim sCurrentText As String = sSubjectName + ":"
                        For Each oField As Field In oFieldDictionary(sSubjectName)
                            sCurrentText += "[" + oField.ItemNumberText + "]"
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
                    Dim oFieldList As List(Of Field) = (From oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where (((Not oField.MarkChoice0(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice0(iIndex) <> oField.MarkChoice2(iIndex))) Or (((Not oField.MarkChoice1(iIndex)) Or (Not oField.MarkChoice2(iIndex))) AndAlso (oField.MarkChoice1(iIndex) <> oField.MarkChoice2(iIndex))) Select oField).ToList
                    Dim oSubjectNameList As List(Of String) = (From oField In oFieldList Select oField.SubjectName Distinct).ToList
                    Dim oFieldDictionary As Dictionary(Of String, List(Of Field)) = (From sSubjectName As String In oSubjectNameList Select New KeyValuePair(Of String, List(Of Field))(sSubjectName, (From oField In oFieldList Where oField.SubjectName = sSubjectName Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    For Each sSubjectName As String In oFieldDictionary.Keys
                        iCurrentRow += 1
                        Dim sCurrentText As String = sSubjectName + ":"
                        For Each oField As Field In oFieldDictionary(sSubjectName)
                            sCurrentText += "[" + oField.ItemNumberText + "]"
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
                    Dim oFieldList As List(Of Field) = (From oField In oCurrentFieldList From iIndex In Enumerable.Range(0, oField.Images.Count) Where ((Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item1) <> Trim(oField.MarkText(iIndex).Item3))) Or ((Not oField.Images(iIndex).Item5) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> String.Empty And Trim(oField.MarkText(iIndex).Item3) <> String.Empty) AndAlso (Trim(oField.MarkText(iIndex).Item2) <> Trim(oField.MarkText(iIndex).Item3))) Select oField).ToList
                    Dim oSubjectNameList As List(Of String) = (From oField In oFieldList Select oField.SubjectName Distinct).ToList
                    Dim oFieldDictionary As Dictionary(Of String, List(Of Field)) = (From sSubjectName As String In oSubjectNameList Select New KeyValuePair(Of String, List(Of Field))(sSubjectName, (From oField In oFieldList Where oField.SubjectName = sSubjectName Select oField).ToList)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    For Each sSubjectName As String In oFieldDictionary.Keys
                        iCurrentRow += 1
                        Dim sCurrentText As String = sSubjectName + ":"
                        For Each oField As Field In oFieldDictionary(sSubjectName)
                            sCurrentText += "[" + oField.ItemNumberText + "]"
                        Next
                        oStatsWorksheet.Cell(iCurrentRow, 1).Value = sCurrentText
                    Next
            End Select

            iCurrentRow += 1
        Next

        Dim iNoData As Integer = Aggregate oField In oFieldCollection Where oField.DataPresent = Field.DataPresentNone Into Count
        Dim iPartialData As Integer = Aggregate oField In oFieldCollection Where oField.DataPresent = Field.DataPresentPartial Into Count
        Dim iFullData As Integer = Aggregate oField In oFieldCollection Where oField.DataPresent = Field.DataPresentFull Into Count

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Summary"
        oStatsWorksheet.Cell(iCurrentRow, 1).Style.Font.SetUnderline(ClosedXML.Excel.XLFontUnderlineValues.Single)

        iCurrentRow += 1
        oStatsWorksheet.Cell(iCurrentRow, 1).Value = "Total Fields: "
        oStatsWorksheet.Cell(iCurrentRow, 2).Value = oFieldCollection.Count.ToString

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
        m_CommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Data saved to file " + oFileInfo.Name + "."))
    End Sub
    Private Sub RenderFieldItem()
        Dim iCurrent As Integer = 1
        For Each oItem In oGridFieldContents.Children
            If oItem.GetType.Equals(GetType(BorderFormItem)) Then
                Dim oBorderFormItem As BorderFormItem = oItem

                oBorderFormItem.InvalidateMeasure()
                oBorderFormItem.InvalidateArrange()
                oBorderFormItem.UpdateLayout()

                If oBorderFormItem.ActualWidth > 0 And oBorderFormItem.ActualHeight > 0 Then
                    Dim bounds As Rect = VisualTreeHelper.GetDescendantBounds(oBorderFormItem)

                    Dim width = bounds.Width + bounds.X
                    Dim height = bounds.Height + bounds.Y

                    Dim rtb As New Imaging.RenderTargetBitmap(CInt(Math.Round(width, MidpointRounding.AwayFromZero)), CInt(Math.Round(height, MidpointRounding.AwayFromZero)), ScreenResolution096, ScreenResolution096, PixelFormats.Pbgra32)

                    Dim dv As New DrawingVisual()
                    Using ctx As DrawingContext = dv.RenderOpen()
                        Dim vb As New VisualBrush(oBorderFormItem)
                        ctx.DrawRectangle(vb, Nothing, New Rect(New Point(bounds.X, bounds.Y), New Point(width, height)))
                    End Using

                    rtb.Render(dv)

                    Converter.BitmapSourceToBitmap(DirectCast(rtb.GetAsFrozen(), ImageSource), 300).Save("D:\Downloads\BorderFormItem" + iCurrent.ToString.PadLeft(2, "0") + ".tif", System.Drawing.Imaging.ImageFormat.Tiff)
                    iCurrent += 1
                End If
            End If
        Next
    End Sub
    Public Const SpireBarcodeKey As String = "QEYYS-2XFDC-VBSMT-2INYI-OL6VD"
    Public Shared Function GetBarcodeImage(ByVal sBarcodeData As String) As Tuple(Of System.Drawing.Bitmap, XUnit, XUnit)
        ' gets barcode image
        Spire.Barcode.BarcodeSettings.ApplyKey(SpireBarcodeKey)
        Dim oBarcodeSettings As New Spire.Barcode.BarcodeSettings()
        With oBarcodeSettings
            .Code128SetMode = Spire.Barcode.Code128SetMode.OnlyB
            .Data = Trim(sBarcodeData)
            .DpiX = RenderResolution300
            .DpiY = RenderResolution300
            .ResolutionType = Spire.Barcode.ResolutionType.UseDpi
            .ShowText = False
            .ShowTopText = False
            .Type = Spire.Barcode.BarCodeType.Code128
            .Unit = System.Drawing.GraphicsUnit.Point
        End With

        Dim oBarcodeGenerator As New Spire.Barcode.BarCodeGenerator(oBarcodeSettings)
        Dim oImage As System.Drawing.Bitmap = oBarcodeGenerator.GenerateImage
        Dim XWidth As New XUnit(oImage.Width * 72 / oImage.HorizontalResolution)
        Dim XHeight As New XUnit(oImage.Height * 72 / oImage.VerticalResolution)

        Return New Tuple(Of System.Drawing.Bitmap, XUnit, XUnit)(oImage, XWidth, XHeight)
    End Function
End Class
Public Class Twain32Shared
#Region "Variables"
    Public Const ServerCommPath As String = "\PRIVATE$\Survey2-Twain32ServerComm"
    Public Const ClientCommPath As String = "\PRIVATE$\Survey2-Twain32ClientComm"
    Public Const PathPrefix As String = "FormatName:DIRECT=OS:"
#End Region

    Public Shared Function Twain32Serialize(Of T)(ByVal data As T) As Byte()
        Dim content As Byte() = {}
        Dim formatter As IFormatter = New Formatters.Binary.BinaryFormatter()
        Using oMemoryStream = New MemoryStream()
            formatter.Serialize(oMemoryStream, data)
            content = oMemoryStream.ToArray
        End Using
        Return content
    End Function
    Public Shared Function Twain32Deserialize(Of T)(ByVal datastream As Byte()) As T
        Dim theObject As T = Nothing
        Dim formatter As IFormatter = New Formatters.Binary.BinaryFormatter()
        Using oMemoryStream = New MemoryStream(datastream)
            theObject = DirectCast(formatter.Deserialize(oMemoryStream), T)
        End Using
        Return theObject
    End Function
    <Serializable()> Public Structure Twain32MessageCache
        Implements IDisposable

        Public Type As MessageType
        Public RequestCode As Guid
        Public Text As String
        Public Colour As Windows.Media.Color
        Public DateTime As Date
        Public Content As Object

        Public Sub New(ByVal oMessageType As MessageType, ByVal oRequestCode As Guid, Optional ByVal oText As String = "", Optional ByVal oColour As Windows.Media.Color = Nothing, Optional ByVal oDateTime As Date = Nothing, Optional ByVal oContent As Object = Nothing)
            disposedValue = False

            Type = oMessageType
            RequestCode = oRequestCode
            Text = oText
            Colour = oColour
            DateTime = oDateTime
            Content = oContent
        End Sub
        Public Enum MessageType As Integer
            None = 0
            Close
            GetTWAINScannerSources
        End Enum
#Region " IDisposable Support "
        Private disposedValue As Boolean
        Public Sub Close()
            Dispose()
        End Sub
        Private Sub Dispose(ByVal disposing As Boolean)
            ' IDisposable
            If Not Me.disposedValue Then
                If disposing Then
                    ' free other state (managed objects)
                    ' unblock waiting threads
                End If

                ' free your own state (unmanaged objects)
                ' set large fields to null
            End If
            Me.disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            ' This code added by Visual Basic to correctly implement the disposable pattern.
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Structure
End Class
Public Class Twain32Shared
    Public Enum ScannerSource As Integer
        None = 0
        TWAIN
        WIA
    End Enum
    Public Class CommonScanner
        Public Const NewString As String = "New"
        Public Const GetTWAINScannerSourcesString As String = "GetTWAINScannerSources"
        Public Const OpenSessionString As String = "OpenSession"
        Public Const CloseSessionString As String = "CloseSession"
        Public Const ScannerIsActiveGetString As String = "ScannerIsActiveGet"
        Public Const ScannerIsActiveSetString As String = "ScannerIsActiveSet"
        Public Const SelectedScannerSourceGetString As String = "SelectedScannerSourceGet"
        Public Const SelectedScannerSourceSetString As String = "SelectedScannerSourceSet"
        Public Const ScannerConfigureString As String = "ScannerConfigure"
        Public Const ScannerScanString As String = "ScannerScan"
        Public Shared Scanner32GUID As New Guid("{d73b4483-1b6c-4031-ad36-75e1b490bbab}")

        Private m_ScannerIsActive As Boolean
        Private m_AppID As NTwain.Data.TWIdentity
        Private WithEvents m_Session As NTwain.TwainSession
        Private WithEvents m_Host32BitConnection As Host32BitConnection
        Private m_SelectedScannerSource As Tuple(Of ScannerSource, String)
        Public Event ReturnMessage(ByVal oColour As Windows.Media.Color, ByVal oDateTime As Date, ByVal sMessage As String)
        Public Event SetScannerSelectedChanged(ByVal sScannerName As String)
        Public Event ReturnScannedImage(ByVal oBitmap As System.Drawing.Bitmap)
        Public Event CloseEvent()

        Public Sub New(ByVal oKnownTypes As List(Of Type), Optional ByVal bBypass As Boolean = True)
            If bBypass Then
                ' runs the 32 bit scanner module
                m_Host32BitConnection = New Host32BitConnection
                If m_Host32BitConnection.StartHost32Bit("Twain32", "PlugIn\Scanner", oKnownTypes, Scanner32GUID) Then
                    m_Host32BitConnection.SendMessage(Of String, String)(NewString, String.Empty)
                Else
                    RaiseEvent ReturnMessage(Windows.Media.Colors.Red, Date.Now, "Unable to initialise scan module.")
                End If
            Else
                ' initialises TWAIN module
                ScannerIsActiveByPass = False

                If IsNothing(m_AppID) Then
                    m_AppID = NTwain.Data.TWIdentity.CreateFromAssembly(NTwain.Data.DataGroups.Image, Reflection.Assembly.GetExecutingAssembly)
                End If
                OpenSession(False)

                Dim oScannerSources As List(Of Tuple(Of ScannerSource, String)) = GetTWAINScannerSources(False)
                If oScannerSources.Count > 0 Then
                    m_SelectedScannerSource = oScannerSources(0)
                Else
                    m_SelectedScannerSource = Nothing
                End If

                CloseSession(False)
            End If
        End Sub
        Public Sub Close(Optional ByVal bBypass As Boolean = True)
            If bBypass Then
                If Not IsNothing(m_Host32BitConnection) Then
                    m_Host32BitConnection.StopHost32Bit()
                End If
            Else
                If (Not IsNothing(m_AppID)) AndAlso (Not IsNothing(m_Session)) AndAlso m_Session.IsDsmOpen Then
                    m_Session.Close()
                End If
            End If
        End Sub
        Public Sub CloseEventHandler() Handles m_Host32BitConnection.CloseEvent
            RaiseEvent CloseEvent()
        End Sub
        Public Function GetTWAINScannerSources(Optional ByVal bBypass As Boolean = True) As List(Of Tuple(Of ScannerSource, String))
            If bBypass Then
                Return m_Host32BitConnection.SendMessage(Of String, List(Of Tuple(Of ScannerSource, String)))(GetTWAINScannerSourcesString, String.Empty)
            Else
                ' gets scanner sources
                Dim oScannerSources As New List(Of Tuple(Of ScannerSource, String))
                Dim oSources As List(Of NTwain.DataSource) = m_Session.GetSources.ToList
                For Each oSource As NTwain.DataSource In oSources
                    oScannerSources.Add(New Tuple(Of ScannerSource, String)(ScannerSource.TWAIN, oSource.Name))
                Next
                Return oScannerSources
            End If
        End Function
        Private Function GetTWAINDataSource(ByVal sName As String) As NTwain.DataSource
            ' gets a scanner data source based on the supplied name
            Dim oSources As List(Of NTwain.DataSource) = m_Session.GetSources.ToList
            For Each oSource As NTwain.DataSource In oSources
                If oSource.Name = sName Then
                    Return oSource
                End If
            Next
            Return Nothing
        End Function
        Public Sub OpenSession(Optional ByVal bBypass As Boolean = True)
            If bBypass Then
                m_Host32BitConnection.SendMessage(Of String, String)(OpenSessionString, String.Empty)
            Else
                ' opens new TWAIN session
                m_Session = New NTwain.TwainSession(m_AppID)
                m_Session.SynchronizationContext = System.Threading.SynchronizationContext.Current
                m_Session.Open()
            End If
        End Sub
        Public Sub CloseSession(Optional ByVal bBypass As Boolean = True)
            If bBypass Then
                m_Host32BitConnection.SendMessage(Of String, String)(CloseSessionString, String.Empty)
            Else
                m_Session.Close()
            End If
        End Sub
        Public Sub ScannerConfigure(Optional ByVal bBypass As Boolean = True)
            If bBypass Then
                m_Host32BitConnection.SendMessage(Of String, String)(ScannerConfigureString, String.Empty)
            Else
                OpenSession(False)

                Dim oSource As NTwain.DataSource = GetTWAINDataSource(SelectedScannerSourceByPass.Item2)

                oSource.Open()

                ScannerIsActiveByPass = True
                If oSource.IsOpen Then
                    Dim oHandle As IntPtr = Process.GetCurrentProcess().MainWindowHandle
                    oSource.Enable(NTwain.SourceEnableMode.ShowUIOnly, False, oHandle)
                End If
            End If
        End Sub
        Public Sub ScannerScan(Optional ByVal bBypass As Boolean = True)
            If bBypass Then
                m_Host32BitConnection.SendMessage(Of String, String)(ScannerScanString, String.Empty)
            Else
                OpenSession(False)

                Dim oSource As NTwain.DataSource = GetTWAINDataSource(SelectedScannerSourceByPass.Item2)

                oSource.Open()

                ScannerIsActiveByPass = True
                If oSource.IsOpen Then
                    Dim oHandle As IntPtr = Process.GetCurrentProcess().MainWindowHandle
                    oSource.Enable(NTwain.SourceEnableMode.ShowUI, False, oHandle)
                End If
            End If
        End Sub
        Public Property ScannerIsActive() As Boolean
            Get
                Return m_Host32BitConnection.SendMessage(Of String, Boolean)(ScannerIsActiveGetString, String.Empty)
            End Get
            Set(value As Boolean)
                m_Host32BitConnection.SendMessage(Of Boolean, String)(ScannerIsActiveSetString, value)
            End Set
        End Property
        Public Property ScannerIsActiveByPass() As Boolean
            Get
                Return m_ScannerIsActive
            End Get
            Set(value As Boolean)
                m_ScannerIsActive = value
            End Set
        End Property
        Public Property SelectedScannerSource() As Tuple(Of ScannerSource, String)
            Get
                Return m_Host32BitConnection.SendMessage(Of String, Tuple(Of ScannerSource, String))(SelectedScannerSourceGetString, String.Empty)
            End Get
            Set(value As Tuple(Of ScannerSource, String))
                m_Host32BitConnection.SendMessage(Of Tuple(Of ScannerSource, String), String)(SelectedScannerSourceSetString, value)
            End Set
        End Property
        Public Property SelectedScannerSourceByPass() As Tuple(Of ScannerSource, String)
            Get
                Return m_SelectedScannerSource
            End Get
            Set(value As Tuple(Of ScannerSource, String))
                m_SelectedScannerSource = value
            End Set
        End Property
        Private Sub Twain_TransferReadyHandler(sender As Object, e As NTwain.TransferReadyEventArgs) Handles m_Session.TransferReady
            ' sets up the data transfer mechanism
            Dim oSource As NTwain.DataSource = GetTWAINDataSource(m_SelectedScannerSource.Item2)

            If oSource.Capabilities.ICapXferMech.GetValues.Contains(NTwain.Data.XferMech.Native) Then
                oSource.Capabilities.ICapXferMech.SetValue(NTwain.Data.XferMech.Native)

            ElseIf oSource.Capabilities.ICapXferMech.GetValues.Contains(NTwain.Data.XferMech.File) Then
                oSource.Capabilities.ICapXferMech.SetValue(NTwain.Data.XferMech.File)

                If oSource.Capabilities.ICapImageFileFormat.GetValues.Contains(NTwain.Data.FileFormat.Bmp) Then
                    Dim oFileXfer As New NTwain.Data.TWSetupFileXfer
                    With oFileXfer
                        .Format = NTwain.Data.FileFormat.Bmp
                        .FileName = IO.Path.GetTempPath + "TwainCapture.bmp"
                        If IO.File.Exists(.FileName) Then
                            IO.File.Delete(.FileName)
                        End If
                    End With

                    oSource.DGControl.SetupFileXfer.Set(oFileXfer)
                End If
            End If
        End Sub
        Private Sub Twain_DataTransferredHandler(sender As Object, e As NTwain.DataTransferredEventArgs) Handles m_Session.DataTransferred
            ' converts the scanned data to a bitmapsource
            Dim oImage As Windows.Media.Imaging.BitmapSource = Nothing

            If e.NativeData <> IntPtr.Zero Then
                Using oStream As IO.Stream = e.GetNativeImageStream()
                    If Not IsNothing(oStream) Then
                        Dim oBitmapImage As New Windows.Media.Imaging.BitmapImage
                        oBitmapImage.BeginInit()
                        oBitmapImage.CacheOption = Windows.Media.Imaging.BitmapCacheOption.OnLoad
                        oBitmapImage.DecodePixelHeight = 0
                        oBitmapImage.DecodePixelWidth = 0
                        oBitmapImage.StreamSource = oStream
                        oBitmapImage.EndInit()
                        If (oBitmapImage.CanFreeze) Then
                            oBitmapImage.Freeze()
                        End If

                        oImage = oBitmapImage
                    End If
                End Using
            ElseIf Not String.IsNullOrEmpty(e.FileDataPath) Then
                oImage = New Windows.Media.Imaging.BitmapImage(New Uri(e.FileDataPath))

                If IO.File.Exists(e.FileDataPath) Then
                    IO.File.Delete(e.FileDataPath)
                End If
            End If

            Using oBitmap As System.Drawing.Bitmap = BitmapSourceToBitmap(oImage, CInt(oImage.DpiX))
                Using oGrayscaleBitmap As System.Drawing.Bitmap = BitmapConvertGrayscale(oBitmap)
                    RaiseEvent ReturnScannedImage(oGrayscaleBitmap)
                End Using
            End Using
        End Sub
        Private Sub Twain_TransferErrorHandler(sender As Object, e As NTwain.TransferErrorEventArgs) Handles m_Session.TransferError
            ' scan error
            RaiseEvent ReturnMessage(Windows.Media.Colors.Red, Date.Now, "Scan error: " + e.Exception.Message)
        End Sub
        Private Sub Twain_StateChangedHandler(senser As Object, e As EventArgs) Handles m_Session.StateChanged
            ' change in state
            If ScannerIsActiveByPass And m_Session.State = 4 Then
                ScannerIsActiveByPass = False
                Dim oSource As NTwain.DataSource = GetTWAINDataSource(m_SelectedScannerSource.Item2)
                oSource.Close()
                m_Session.Close()
            End If
        End Sub
        Private Delegate Sub ReturnMessageDelegate(ByVal oColour As Windows.Media.Color, ByVal oDateTime As Date, ByVal sMessage As String, ByVal bInvoking As Boolean)
        Private Sub ReturnMessageHandler(ByVal oColour As Windows.Media.Color, ByVal oDateTime As Date, ByVal sMessage As String, Optional ByVal bInvoking As Boolean = False) Handles m_Host32BitConnection.ReturnMessage
            If IsNothing(Windows.Application.Current) Then
                RaiseEvent ReturnMessage(oColour, oDateTime, sMessage)
            Else
                If bInvoking Then
                    RaiseEvent ReturnMessage(oColour, oDateTime, sMessage)
                Else
                    Windows.Application.Current.Dispatcher.BeginInvoke(New ReturnMessageDelegate(AddressOf ReturnMessageHandler), {oColour, oDateTime, sMessage, True})
                End If
            End If
        End Sub
        Private Delegate Sub SetScannerSelectedChangedDelegate(ByVal sScannerName As String, ByVal bInvoking As Boolean)
        Public Sub SetScannerSelectedHandler(ByVal sScannerName As String, Optional ByVal bInvoking As Boolean = False) Handles m_Host32BitConnection.SetScannerSelected
            If IsNothing(Windows.Application.Current) Then
                RaiseEvent SetScannerSelectedChanged(sScannerName)
            Else
                If bInvoking Then
                    RaiseEvent SetScannerSelectedChanged(sScannerName)
                Else
                    Windows.Application.Current.Dispatcher.BeginInvoke(New SetScannerSelectedChangedDelegate(AddressOf SetScannerSelectedHandler), {sScannerName, True})
                End If
            End If
        End Sub
        Private Delegate Sub ReturnScannedImageDelegate(ByVal oBitmap As System.Drawing.Bitmap, ByVal bInvoking As Boolean)
        Private Sub ReturnScannedImageHandler(ByVal oBitmap As System.Drawing.Bitmap, Optional ByVal bInvoking As Boolean = False) Handles m_Host32BitConnection.ReturnScannedImage
            If IsNothing(Windows.Application.Current) Then
                RaiseEvent ReturnScannedImage(oBitmap)
            Else
                If bInvoking Then
                    RaiseEvent ReturnScannedImage(oBitmap)
                Else
                    Windows.Application.Current.Dispatcher.BeginInvoke(New ReturnScannedImageDelegate(AddressOf ReturnScannedImageHandler), {oBitmap, True})
                End If
            End If
        End Sub
        Private Shared Function BitmapSourceToBitmap(ByVal oBitmapSource As Windows.Media.Imaging.BitmapSource, ByVal oResolution As Single) As System.Drawing.Bitmap
            ' converts WPF bitmapsource to GDI bitmap
            Dim oFormatConvertedBitmap As New Windows.Media.Imaging.FormatConvertedBitmap
            oFormatConvertedBitmap.BeginInit()
            oFormatConvertedBitmap.Source = oBitmapSource
            oFormatConvertedBitmap.DestinationFormat = Windows.Media.PixelFormats.Bgra32
            oFormatConvertedBitmap.EndInit()

            Dim oBitmap As New System.Drawing.Bitmap(oFormatConvertedBitmap.PixelWidth, oFormatConvertedBitmap.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Dim oBitmapData = oBitmap.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, oBitmap.Size), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            oFormatConvertedBitmap.CopyPixels(Windows.Int32Rect.Empty, oBitmapData.Scan0, oBitmapData.Height * oBitmapData.Stride, oBitmapData.Stride)
            oBitmap.UnlockBits(oBitmapData)
            oBitmap.SetResolution(oResolution, oResolution)
            Return oBitmap
        End Function
        Private Shared Function BitmapConvertGrayscale(ByVal oImage As System.Drawing.Bitmap) As System.Drawing.Bitmap
            If IsNothing(oImage) Then
                Return Nothing
            ElseIf oImage.PixelFormat = System.Drawing.Imaging.PixelFormat.Format8bppIndexed Then
                Return oImage.Clone
            ElseIf oImage.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppArgb Then
                Dim oReturnBitmap As System.Drawing.Bitmap = AForge.Imaging.Filters.Grayscale.CommonAlgorithms.BT709.Apply(oImage)
                oReturnBitmap.SetResolution(oImage.HorizontalResolution, oImage.VerticalResolution)
                Return oReturnBitmap
            Else
                Return Nothing
            End If
        End Function
        Public Shared Function GetKnownTypes() As List(Of Type)
            Dim oKnownTypes As New List(Of Type)
            Return oKnownTypes
        End Function
    End Class
    Public Class Host32BitConnection
        Public Const CloseString As String = "Close"
        Public Const ValidateString As String = "Validate"
        Public Const MessageString As String = "Message"
        Public Const EmptyString As String = ""
        Public Const OKString As String = "OK"
        Public Const EventMessageString As String = "EventMessage"
        Public Const EventSetScannerString As String = "EventSetScanner"
        Public Const EventReturnImageString As String = "EventReturnImage"
        Private oSenderPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
        Private oReceiverPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
        Private oEventSenderPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
        Private oEventReceiverPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
        Private oStreamString As StreamString = Nothing
        Private oEventStreamString As StreamString = Nothing
        Private oKnownTypesLocal As List(Of Type)
        Public Event ReturnMessage(ByVal oColour As Windows.Media.Color, ByVal oDateTime As Date, ByVal sMessage As String)
        Public Event SetScannerSelected(ByVal sToolTip As String)
        Public Event ReturnScannedImage(ByVal oBitmap As System.Drawing.Bitmap)
        Public Event CloseEvent()

        Public Function StartHost32Bit(ByVal sProcessName As String, ByVal sSubDirectory As String, ByVal oKnownTypes As List(Of Type), ByVal oGUID As Guid) As Boolean
            ' starts a 32 bit process and establishes a named pipe connection
            Dim oProcessArrayList As New List(Of Process)
            oProcessArrayList.AddRange(Process.GetProcessesByName(sProcessName))

            ' terminates any running processes
            If oProcessArrayList.Count > 0 Then
                For Each oProcess As Process In oProcessArrayList
                    oProcess.Kill()
                Next
            End If

            oSenderPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.Out, IO.HandleInheritability.Inheritable)
            oReceiverPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.In, IO.HandleInheritability.Inheritable)
            oEventSenderPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.Out, IO.HandleInheritability.Inheritable)
            oEventReceiverPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.In, IO.HandleInheritability.Inheritable)

            Task.Run(AddressOf Listener)

            ' locate calling assembly location
            Dim oFileIO As New IO.FileInfo(Reflection.Assembly.GetCallingAssembly.Location)
            Dim sFullName As String = oFileIO.DirectoryName + "\"
            If sSubDirectory = String.Empty Then
                sFullName += sProcessName + ".exe"
            Else
                sFullName += sSubDirectory + "\" + sProcessName + ".exe"
            End If

            Dim senderID As String = oSenderPipe.GetClientHandleAsString()
            Dim receiverID As String = oReceiverPipe.GetClientHandleAsString()
            Dim eventsenderID As String = oEventSenderPipe.GetClientHandleAsString()
            Dim eventreceiverID As String = oEventReceiverPipe.GetClientHandleAsString()

            Dim oProcessStartInfo As New ProcessStartInfo(sFullName, senderID + " " + receiverID + " " + eventsenderID + " " + eventreceiverID)
            oProcessStartInfo.CreateNoWindow = True
            oProcessStartInfo.UseShellExecute = False

            Try
                Process.Start(oProcessStartInfo)
            Catch ex As Win32Exception
                Return False
            End Try
            If Process.GetProcessesByName(sProcessName).Length = 0 Then
                ' not successful in loading
                oSenderPipe.Dispose()
                oReceiverPipe.Dispose()
                oEventSenderPipe.Dispose()
                oEventReceiverPipe.Dispose()
                Return False
            Else
                oSenderPipe.DisposeLocalCopyOfClientHandle()
                oReceiverPipe.DisposeLocalCopyOfClientHandle()
                oEventSenderPipe.DisposeLocalCopyOfClientHandle()
                oEventReceiverPipe.DisposeLocalCopyOfClientHandle()

                oKnownTypesLocal = oKnownTypes

                Dim bReturn As Boolean = False
                oStreamString = New StreamString(oSenderPipe, oReceiverPipe)
                oStreamString.WriteString(ValidateString)
                If oStreamString.ReadString() = oGUID.ToString Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function
        Public Sub StopHost32Bit()
            ' stops the 32 bit process
            If Not IsNothing(oStreamString) Then
                oStreamString.WriteString(CloseString)
                System.Threading.Thread.Sleep(50)
                oSenderPipe.Dispose()
                oReceiverPipe.Dispose()
                oEventSenderPipe.Dispose()
                oEventReceiverPipe.Dispose()
            End If
        End Sub
        Public Function SendMessage(Of T, U)(ByVal sFunctionName As String, ByVal oInputData As T) As U
            ' sends a message with input data type T and returns with data type U
            ' order of events
            ' 1) send 'message string'
            ' 2) send data as base64 string
            ' 3) send function name
            ' 4) receive data as base64 string
            Dim oReturn As U = Nothing
            Dim oUnicodeEncoding As New Text.UnicodeEncoding

            If Not IsNothing(oStreamString) Then
                ' 1) send 'message string'
                oStreamString.WriteString(MessageString)
                Dim sMessage As String = oStreamString.ReadString
                If sMessage = OKString Then

                    ' 2) send data as base64 string
                    Dim sDataString As String = SerializeDataContractText(oInputData, oKnownTypesLocal)
                    oStreamString.WriteString(sDataString)
                    sMessage = oStreamString.ReadString
                    If sMessage = OKString Then

                        ' 3) send function name
                        oStreamString.WriteString(sFunctionName)

                        ' 4) receive data as base64 string
                        sMessage = oStreamString.ReadString
                        oReturn = DeserializeDataContractText(sMessage, oKnownTypesLocal)
                    End If
                End If
            End If
            Return oReturn
        End Function
        Private Sub Listener()
            ' listens to event messages
            Do While True
                Try
                    ' initialises streamstring if not done
                    If IsNothing(oEventStreamString) Then
                        oEventStreamString = New StreamString(oEventSenderPipe, oEventReceiverPipe)
                    End If

                    Dim sMessage As String = oEventStreamString.ReadString
                    Select Case sMessage
                        Case MessageString
                            ' order of events
                            ' 1) receive 'message string'
                            ' 3) receive data as base64 string
                            ' 4) receive function name
                            ' 5) send data as base64 string

                            ' 1) receive 'message string'
                            oEventStreamString.WriteString(OKString)

                            ' 2) receive data as base64 string
                            sMessage = oEventStreamString.ReadString

                            Dim oTObject = DeserializeDataContractText(sMessage, oKnownTypesLocal)
                            oEventStreamString.WriteString(OKString)

                            ' 3) receive function name
                            sMessage = oEventStreamString.ReadString
                            Dim oReturnObject As Object = String.Empty
                            Select Case sMessage
                                Case CloseString
                                    RaiseEvent CloseEvent()
                                    Exit Do
                                Case EventMessageString
                                    Dim oReturnMessage As Tuple(Of Windows.Media.Color, Date, String) = oTObject
                                    RaiseEvent ReturnMessage(oReturnMessage.Item1, oReturnMessage.Item2, oReturnMessage.Item3)
                                Case EventSetScannerString
                                    RaiseEvent SetScannerSelected(oTObject)
                                Case EventReturnImageString
                                    Dim oBitmap As System.Drawing.Bitmap = oTObject
                                    oBitmap.SetResolution(CInt(oBitmap.HorizontalResolution), CInt(oBitmap.VerticalResolution))
                                    RaiseEvent ReturnScannedImage(oBitmap)
                            End Select

                            ' 4) send data as base64 string
                            ' send only if there is a return variable
                            Dim sDataString As String = SerializeDataContractText(oReturnObject, oKnownTypesLocal)
                            oEventStreamString.WriteString(sDataString)
                    End Select
                Catch e As IO.IOException
                    ' if connection interrupted, then exit
                    Exit Do
                End Try
            Loop
        End Sub
        Public Class StreamString
            Private m_WriteStream As IO.Stream
            Private m_ReadStream As IO.Stream
            Private m_Encoding As Text.UnicodeEncoding

            Public Sub New(ByRef oWriteStream As IO.Stream, ByRef oReadStream As IO.Stream)
                m_WriteStream = oWriteStream
                m_ReadStream = oReadStream
                m_Encoding = New Text.UnicodeEncoding
            End Sub
            Public Function ReadString() As String
                Try
                    Dim byteArray(3) As Byte
                    For i = 0 To 3
                        Dim iByte As Integer = m_ReadStream.ReadByte()
                        If iByte = -1 Then
                            Return String.Empty
                        Else
                            byteArray(i) = iByte
                        End If
                    Next
                    Dim iBufferLength As Integer = BitConverter.ToInt32(byteArray, 0)
                    Dim bBuffer As Array = Array.CreateInstance(GetType(Byte), iBufferLength)
                    Dim iBufferReadCount As Integer = m_ReadStream.Read(bBuffer, 0, iBufferLength)

                    If iBufferReadCount = iBufferLength Then
                        Return m_Encoding.GetString(bBuffer)
                    Else
                        Return String.Empty
                    End If
                Catch ex As System.OverflowException
                    Return String.Empty
                End Try
            End Function
            Public Sub WriteString(outString As String)
                Dim bBuffer As Byte() = m_Encoding.GetBytes(outString)
                Dim iBufferLength As Integer = bBuffer.Length
                Dim byteArray As Byte() = BitConverter.GetBytes(iBufferLength)

                For i = 0 To 3
                    m_WriteStream.WriteByte(byteArray(i))
                Next
                m_WriteStream.Write(bBuffer, 0, iBufferLength)
                m_WriteStream.Flush()
            End Sub
        End Class
        Private Shared Function SerializeDataContractText(ByVal data As Object, ByVal oKnownTypes As List(Of Type)) As String
            ' serialise using data contract serialiser
            ' returns base64 text
            Using oMemoryStream As New IO.MemoryStream
                SerializeDataContractStream(oMemoryStream, data, oKnownTypes)
                Dim bByteArray As Byte() = oMemoryStream.ToArray
                Return Convert.ToBase64String(bByteArray)
            End Using
        End Function
        Private Shared Function DeserializeDataContractText(ByVal sBase64String As String, Optional ByVal oKnownTypes As List(Of Type) = Nothing) As Object
            ' deserialise from base64 text
            Using oMemoryStream As New IO.MemoryStream(Convert.FromBase64String(sBase64String))
                Return DeserializeDataContractStream(oMemoryStream, oKnownTypes)
            End Using
        End Function
        Private Shared Sub SerializeDataContractStream(ByRef oStream As IO.Stream, ByVal data As Object, ByVal oKnownTypes As List(Of Type))
            ' serialise to stream
            Dim oDataContractSerializer As DataContractSerializer = Nothing
            oDataContractSerializer = New DataContractSerializer(GetType(Object), oKnownTypes)
            oDataContractSerializer.WriteObject(oStream, data)
        End Sub
        Private Shared Function DeserializeDataContractStream(ByRef oStream As IO.Stream, ByVal oKnownTypes As List(Of Type)) As Object
            ' deserialise from stream
            Dim oXmlDictionaryReaderQuotas As New Xml.XmlDictionaryReaderQuotas()
            oXmlDictionaryReaderQuotas.MaxArrayLength = 100000000
            Dim oXmlDictionaryReader As Xml.XmlDictionaryReader = Xml.XmlDictionaryReader.CreateTextReader(oStream, oXmlDictionaryReaderQuotas)

            Dim theObject As Object = Nothing
            Try
                Dim oDataContractSerializer As DataContractSerializer = Nothing
                oDataContractSerializer = New DataContractSerializer(GetType(Object), oKnownTypes)
                theObject = oDataContractSerializer.ReadObject(oXmlDictionaryReader, True)
            Catch ex As SerializationException
            End Try

            oXmlDictionaryReader.Close()
            Return theObject
        End Function
    End Class
End Class
Module Twain32Server
    Private Const iMaxErrorCount As Integer = 10
    Private WithEvents oCommonScanner As CommonScanner
    Private oStreamString As Host32BitConnection.StreamString = Nothing
    Private oEventStreamString As Host32BitConnection.StreamString = Nothing
    Private oSenderPipe As AnonymousPipeClientStream = Nothing
    Private oReceiverPipe As AnonymousPipeClientStream = Nothing
    Private oEventSenderPipe As AnonymousPipeClientStream = Nothing
    Private oEventReceiverPipe As AnonymousPipeClientStream = Nothing
    Private iErrorCount As Integer = 0

    Sub Main(ByVal args As String())
        ' attach a debugger
        'If Not Debugger.IsAttached Then
        'Debugger.Launch()
        'End If

        Dim parentSenderID As String = args(0)
        Dim parentReceiverID As String = args(1)
        Dim parentEventSenderID As String = args(2)
        Dim parentEventReceiverID As String = args(3)

        oSenderPipe = New AnonymousPipeClientStream(PipeDirection.Out, parentReceiverID)
        oReceiverPipe = New AnonymousPipeClientStream(PipeDirection.In, parentSenderID)
        oEventSenderPipe = New AnonymousPipeClientStream(PipeDirection.Out, parentEventReceiverID)
        oEventReceiverPipe = New AnonymousPipeClientStream(PipeDirection.In, parentEventSenderID)

        ' main messaging loop
        Do While True
            Try
                ' initialises streamstring if not done
                If IsNothing(oStreamString) Then
                    oStreamString = New Host32BitConnection.StreamString(oSenderPipe, oReceiverPipe)
                End If

                Dim sMessage As String = oStreamString.ReadString
                Select Case sMessage
                    Case Host32BitConnection.ValidateString
                        iErrorCount = 0
                        oStreamString.WriteString(CommonScanner.Scanner32GUID.ToString)
                    Case Host32BitConnection.CloseString
                        iErrorCount = 0
                        SendMessage(Of String, String)(Host32BitConnection.CloseString, String.Empty)
                        Cleanup()
                        Exit Do
                    Case Host32BitConnection.MessageString
                        iErrorCount = 0

                        ' order of events
                        ' 1) receive 'message string'
                        ' 3) receive data as base64 string
                        ' 4) receive function name
                        ' 5) send data as base64 string

                        ' 1) receive 'message string'
                        oStreamString.WriteString(Host32BitConnection.OKString)

                        ' 2) receive data as base64 string
                        sMessage = oStreamString.ReadString

                        Dim oTObject = DeserializeDataContractText(sMessage, CommonScanner.GetKnownTypes)
                        oStreamString.WriteString(Host32BitConnection.OKString)

                        ' 3) receive function name
                        sMessage = oStreamString.ReadString
                        Dim oReturnObject As Object = String.Empty
                        Select Case sMessage
                            Case CommonScanner.NewString
                                oCommonScanner = New CommonScanner(CommonScanner.GetKnownTypes, False)

                                ' set scanner selected
                                If IsNothing(oCommonScanner.SelectedScannerSourceByPass) Then
                                    oCommonScanner.SetScannerSelectedHandler(String.Empty)
                                Else
                                    oCommonScanner.SetScannerSelectedHandler(oCommonScanner.SelectedScannerSourceByPass.Item2)
                                End If
                            Case CommonScanner.GetTWAINScannerSourcesString
                                oReturnObject = oCommonScanner.GetTWAINScannerSources(False)
                            Case CommonScanner.OpenSessionString
                                oCommonScanner.OpenSession(False)
                            Case CommonScanner.CloseSessionString
                                oCommonScanner.CloseSession(False)
                            Case CommonScanner.ScannerIsActiveGetString
                                oReturnObject = oCommonScanner.ScannerIsActiveByPass
                            Case CommonScanner.ScannerIsActiveSetString
                                oCommonScanner.ScannerIsActiveByPass = oTObject
                            Case CommonScanner.SelectedScannerSourceGetString
                                oReturnObject = oCommonScanner.SelectedScannerSourceByPass
                            Case CommonScanner.SelectedScannerSourceSetString
                                oCommonScanner.SelectedScannerSourceByPass = oTObject
                            Case CommonScanner.ScannerConfigureString
                                oCommonScanner.ScannerConfigure(False)
                            Case CommonScanner.ScannerScanString
                                oCommonScanner.ScannerScan(False)
                        End Select

                        ' 4) send data as base64 string
                        ' send only if there is a return variable
                        Dim sDataString As String = SerializeDataContractText(oReturnObject, CommonScanner.GetKnownTypes)
                        oStreamString.WriteString(sDataString)
                    Case Host32BitConnection.EmptyString
                        ' communication error
                        iErrorCount += 1
                End Select
            Catch e As IO.IOException
                ' if connection interrupted, then exit
                Cleanup()
                Exit Do
            End Try

            If iErrorCount > iMaxErrorCount Then
                ' communication interrupted
                Cleanup()
                Exit Do
            End If
        Loop

        oCommonScanner.Close()
    End Sub
    Private Sub Cleanup()
        oSenderPipe.Dispose()
        oReceiverPipe.Dispose()
        oEventSenderPipe.Dispose()
        oEventReceiverPipe.Dispose()
    End Sub
    Private Function SendMessage(Of T, U)(ByVal sFunctionName As String, ByVal oInputData As T) As U
        ' sends a message with input data type T and returns with data type U
        ' order of events
        ' 1) send 'message string'
        ' 2) send data as base64 string
        ' 3) send function name
        ' 4) receive data as base64 string
        Dim oReturn As U = Nothing
        Dim oUnicodeEncoding As New Text.UnicodeEncoding

        If IsNothing(oEventStreamString) Then
            oEventStreamString = New Host32BitConnection.StreamString(oEventSenderPipe, oEventReceiverPipe)
        End If

        If Not IsNothing(oEventStreamString) Then
            ' 1) send 'message string'
            oEventStreamString.WriteString(Host32BitConnection.MessageString)
            Dim sMessage As String = oEventStreamString.ReadString
            If sMessage = Host32BitConnection.OKString Then

                ' 2) send data as base64 string
                Dim sDataString As String = SerializeDataContractText(oInputData, CommonScanner.GetKnownTypes)
                oEventStreamString.WriteString(sDataString)
                sMessage = oEventStreamString.ReadString
                If sMessage = Host32BitConnection.OKString Then

                    ' 3) send function name
                    oEventStreamString.WriteString(sFunctionName)

                    ' 4) receive data as base64 string
                    sMessage = oEventStreamString.ReadString
                    oReturn = DeserializeDataContractText(sMessage, CommonScanner.GetKnownTypes)
                End If
            End If
        End If
        Return oReturn
    End Function
    Private Sub ReturnMessage(ByVal oColour As Color, ByVal oDateTime As Date, ByVal sMessage As String) Handles oCommonScanner.ReturnMessage
        SendMessage(Of Tuple(Of Color, Date, String), String)(Host32BitConnection.EventMessageString, New Tuple(Of Color, Date, String)(oColour, oDateTime, sMessage))
    End Sub
    Private Sub SetScannerChanged(ByVal sScannerName As String) Handles oCommonScanner.SetScannerSelectedChanged
        SendMessage(Of String, String)(Host32BitConnection.EventSetScannerString, sScannerName)
    End Sub
    Private Sub ReturnScannedImage(ByVal oBitmap As System.Drawing.Bitmap) Handles oCommonScanner.ReturnScannedImage
        SendMessage(Of System.Drawing.Bitmap, String)(Host32BitConnection.EventReturnImageString, oBitmap)
    End Sub
    Private Function SerializeDataContractText(ByVal data As Object, ByVal oKnownTypes As List(Of Type)) As String
        ' serialise using data contract serialiser
        ' returns base64 text
        Using oMemoryStream As New IO.MemoryStream
            SerializeDataContractStream(oMemoryStream, data, oKnownTypes)
            Dim bByteArray As Byte() = oMemoryStream.ToArray
            Return Convert.ToBase64String(bByteArray)
        End Using
    End Function
    Private Function DeserializeDataContractText(ByVal sBase64String As String, Optional ByVal oKnownTypes As List(Of Type) = Nothing) As Object
        ' deserialise from base64 text
        Using oMemoryStream As New IO.MemoryStream(Convert.FromBase64String(sBase64String))
            Return DeserializeDataContractStream(oMemoryStream, oKnownTypes)
        End Using
    End Function
    Private Sub SerializeDataContractStream(ByRef oStream As IO.Stream, ByVal data As Object, ByVal oKnownTypes As List(Of Type))
        ' serialise to stream
        Dim oDataContractSerializer As DataContractSerializer = Nothing
        oDataContractSerializer = New DataContractSerializer(GetType(Object), oKnownTypes)
        oDataContractSerializer.WriteObject(oStream, data)
    End Sub
    Private Function DeserializeDataContractStream(ByRef oStream As IO.Stream, ByVal oKnownTypes As List(Of Type)) As Object
        ' deserialise from stream
        Dim oXmlDictionaryReaderQuotas As New Xml.XmlDictionaryReaderQuotas()
        oXmlDictionaryReaderQuotas.MaxArrayLength = 100000000
        Dim oXmlDictionaryReader As Xml.XmlDictionaryReader = Xml.XmlDictionaryReader.CreateTextReader(oStream, oXmlDictionaryReaderQuotas)

        Dim theObject As Object = Nothing
        Try
            Dim oDataContractSerializer As DataContractSerializer = Nothing
            oDataContractSerializer = New DataContractSerializer(GetType(Object), oKnownTypes)
            theObject = oDataContractSerializer.ReadObject(oXmlDictionaryReader, True)
        Catch ex As SerializationException
        End Try

        oXmlDictionaryReader.Close()
        Return theObject
    End Function
End Module
Module Twain32Server
    Private oStreamString As StreamString = Nothing
    Private oKeepAliveStreamString As StreamString = Nothing
    Private oSenderPipe As AnonymousPipeClientStream = Nothing
    Private oReceiverPipe As AnonymousPipeClientStream = Nothing
    Private oKeepAliveSenderPipe As AnonymousPipeClientStream = Nothing
    Private oKeepAliveReceiverPipe As AnonymousPipeClientStream = Nothing
    Private Blocker As Threading.ManualResetEvent
    Private ErrorCount As Integer = 0
    Private DoLoop As Boolean = True

    Sub Main(ByVal args As String())
        ' attach a debugger
        If Not Debugger.IsAttached Then
            Debugger.Launch()
        End If

        Dim parentSenderID As String = args(0)
        Dim parentReceiverID As String = args(1)
        Dim parentKeepAliveSenderID As String = args(2)
        Dim parentKeepAliveReceiverID As String = args(3)

        oSenderPipe = New AnonymousPipeClientStream(PipeDirection.Out, parentReceiverID)
        oReceiverPipe = New AnonymousPipeClientStream(PipeDirection.In, parentSenderID)
        oKeepAliveSenderPipe = New AnonymousPipeClientStream(PipeDirection.Out, parentKeepAliveReceiverID)
        oKeepAliveReceiverPipe = New AnonymousPipeClientStream(PipeDirection.In, parentKeepAliveSenderID)

        ' initialises streamstring if not done
        Dim bSuccess As Boolean = True
        Try
            If IsNothing(oStreamString) Then
                oStreamString = New StreamString(oSenderPipe, oReceiverPipe)
            End If

            If IsNothing(oKeepAliveStreamString) Then
                oKeepAliveStreamString = New StreamString(oKeepAliveSenderPipe, oKeepAliveReceiverPipe)
            End If
        Catch e As IO.IOException
            ' if connection interrupted, then exit
            Cleanup()
            bSuccess = False
        End Try

        If bSuccess Then
            Blocker = New Threading.ManualResetEvent(False)
            Task.Run(AddressOf Listener)
            Task.Run(AddressOf KeepAlive)
            Blocker.WaitOne()
            Cleanup()

            If Not IsNothing(Blocker) Then
                Blocker.Dispose()
            End If
        End If
    End Sub
    Private Sub Cleanup()
        DoLoop = False

        oSenderPipe.Dispose()
        oReceiverPipe.Dispose()
        oKeepAliveSenderPipe.Dispose()
        oKeepAliveReceiverPipe.Dispose()

        oSenderPipe = Nothing
        oReceiverPipe = Nothing
        oKeepAliveSenderPipe = Nothing
        oKeepAliveReceiverPipe = Nothing
    End Sub
    Private Sub Listener()
        ' main messaging loop
        Do
            Try
                If IsNothing(oStreamString) Then
                    DoLoop = False
                Else
                    Dim sMessage As String = oStreamString.ReadString
                    Select Case sMessage
                        Case Constants.ValidateString
                            ErrorCount = 0
                            oStreamString.WriteString(Constants.Twain32GUID.ToString)
                        Case Constants.EmptyString
                            ' communication error
                            ErrorCount += 1
                    End Select
                End If
            Catch e As IO.IOException
                ' if connection interrupted, then exit
                DoLoop = False
            End Try

            If ErrorCount > Constants.MaxErrorCount Then
                ' communication interrupted
                DoLoop = False
            End If
        Loop While DoLoop

        If Not IsNothing(Blocker) Then
            Blocker.Set()
        End If
    End Sub
    Private Sub KeepAlive()
        ' keep alive loop
        Dim oSetDate As Date = Date.Now
        Do
            Try
                If IsNothing(oKeepAliveStreamString) Then
                    DoLoop = False
                Else
                    Dim sMessage As String = oKeepAliveStreamString.ReadString
                    Select Case sMessage
                        Case Constants.ValidateString
                            ErrorCount = 0
                            oKeepAliveStreamString.WriteString(Constants.Twain32GUID.ToString)
                        Case Constants.KeepAliveString
                            oSetDate = Date.Now
                        Case Constants.EmptyString
                            ' communication error
                            ErrorCount += 1
                    End Select
                End If
            Catch e As IO.IOException
                ' if connection interrupted, then exit
                DoLoop = False
            End Try

            If ErrorCount > Constants.MaxErrorCount Then
                ' communication interrupted
                DoLoop = False
            End If

            ' if the keep alive timer has not been signaled for more than the set amount of time, then unblock the main thread
            If (Date.Now - oSetDate).TotalMilliseconds > Constants.KeepAliveTimer * 5 Then
                Blocker.Set()
            End If
        Loop While DoLoop

        If Not IsNothing(Blocker) Then
            Blocker.Set()
        End If
    End Sub
    Private Function SerializeDataContractText(ByVal data As Object, ByVal oKnownTypes As List(Of Type)) As String
        ' serialise using data contract serialiser
        ' returns base64 text
        Using oMemoryStream As New IO.MemoryStream
            SerializeDataContractStream(oMemoryStream, data, oKnownTypes)
            Dim bByteArray As Byte() = oMemoryStream.ToArray
            Return Convert.ToBase64String(bByteArray)
        End Using
    End Function
    Private Function DeserializeDataContractText(ByVal sBase64String As String, Optional ByVal oKnownTypes As List(Of Type) = Nothing) As Object
        ' deserialise from base64 text
        Using oMemoryStream As New IO.MemoryStream(Convert.FromBase64String(sBase64String))
            Return DeserializeDataContractStream(oMemoryStream, oKnownTypes)
        End Using
    End Function
    Private Sub SerializeDataContractStream(ByRef oStream As IO.Stream, ByVal data As Object, ByVal oKnownTypes As List(Of Type))
        ' serialise to stream
        Dim oDataContractSerializer As DataContractSerializer = Nothing
        oDataContractSerializer = New DataContractSerializer(GetType(Object), oKnownTypes)
        oDataContractSerializer.WriteObject(oStream, data)
    End Sub
    Private Function DeserializeDataContractStream(ByRef oStream As IO.Stream, ByVal oKnownTypes As List(Of Type)) As Object
        ' deserialise from stream
        Dim oXmlDictionaryReaderQuotas As New Xml.XmlDictionaryReaderQuotas()
        oXmlDictionaryReaderQuotas.MaxArrayLength = 100000000
        Dim oXmlDictionaryReader As Xml.XmlDictionaryReader = Xml.XmlDictionaryReader.CreateTextReader(oStream, oXmlDictionaryReaderQuotas)

        Dim theObject As Object = Nothing
        Try
            Dim oDataContractSerializer As DataContractSerializer = Nothing
            oDataContractSerializer = New DataContractSerializer(GetType(Object), oKnownTypes)
            theObject = oDataContractSerializer.ReadObject(oXmlDictionaryReader, True)
        Catch ex As SerializationException
        End Try

        oXmlDictionaryReader.Close()
        Return theObject
    End Function
End Module
Public Class CommonScanner
    Implements IDisposable

    Private Const ProcessName As String = "Twain32"
    Private Const ProcessSubDirectory As String = ""

    Private oStreamString As StreamString = Nothing
    Private oKeepAliveStreamString As StreamString = Nothing
    Private oSenderPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
    Private oReceiverPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
    Private oKeepAliveSenderPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
    Private oKeepAliveReceiverPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
    Private DoLoop As Boolean = True

    Public Event SetScannerSelectedChanged(ByVal sScannerName As String)
    Sub New()
        ' starts a 32 bit process and establishes a named pipe connection
        Dim oProcessArrayList As New List(Of Process)
        oProcessArrayList.AddRange(Process.GetProcessesByName(ProcessName))

        ' terminates any running processes
        If oProcessArrayList.Count > 0 Then
            For Each oProcess As Process In oProcessArrayList
                oProcess.Kill()
            Next
        End If

        oSenderPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.Out, IO.HandleInheritability.Inheritable)
        oReceiverPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.In, IO.HandleInheritability.Inheritable)
        oKeepAliveSenderPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.Out, IO.HandleInheritability.Inheritable)
        oKeepAliveReceiverPipe = New IO.Pipes.AnonymousPipeServerStream(IO.Pipes.PipeDirection.In, IO.HandleInheritability.Inheritable)

        ' locate calling assembly location
        Dim oFileIO As New IO.FileInfo(Reflection.Assembly.GetCallingAssembly.Location)
        Dim sFullName As String = oFileIO.DirectoryName + "\"
        If ProcessSubDirectory = String.Empty Then
            sFullName += ProcessName + ".exe"
        Else
            sFullName += ProcessSubDirectory + "\" + ProcessName + ".exe"
        End If

        Dim senderID As String = oSenderPipe.GetClientHandleAsString()
        Dim receiverID As String = oReceiverPipe.GetClientHandleAsString()
        Dim KeepAlivesenderID As String = oKeepAliveSenderPipe.GetClientHandleAsString()
        Dim KeepAlivereceiverID As String = oKeepAliveReceiverPipe.GetClientHandleAsString()

        Dim oProcessStartInfo As New ProcessStartInfo(sFullName, senderID + " " + receiverID + " " + KeepAlivesenderID + " " + KeepAlivereceiverID)
        oProcessStartInfo.CreateNoWindow = True
        oProcessStartInfo.UseShellExecute = False

        Dim bSuccess As Boolean = True
        Try
            Process.Start(oProcessStartInfo)
        Catch ex As Win32Exception
            bSuccess = False
        End Try

        If bSuccess Then
            If Process.GetProcessesByName(ProcessName).Length = 0 Then
                ' not successful in loading
                oSenderPipe.Dispose()
                oReceiverPipe.Dispose()
                oKeepAliveSenderPipe.Dispose()
                oKeepAliveReceiverPipe.Dispose()
                Close()
            Else
                oSenderPipe.DisposeLocalCopyOfClientHandle()
                oReceiverPipe.DisposeLocalCopyOfClientHandle()
                oKeepAliveSenderPipe.DisposeLocalCopyOfClientHandle()
                oKeepAliveReceiverPipe.DisposeLocalCopyOfClientHandle()

                oStreamString = New StreamString(oSenderPipe, oReceiverPipe)
                oStreamString.WriteString(Constants.ValidateString)
                If oStreamString.ReadString() = Constants.Twain32GUID.ToString Then
                    oKeepAliveStreamString = New StreamString(oKeepAliveSenderPipe, oKeepAliveReceiverPipe)
                    oKeepAliveStreamString.WriteString(Constants.ValidateString)
                    If oKeepAliveStreamString.ReadString() = Constants.Twain32GUID.ToString Then
                        Task.Run(AddressOf Listener)
                        Task.Run(AddressOf KeepAlive)
                    Else
                        Close()
                    End If
                Else
                    Close()
                End If
            End If
        Else
            Close()
        End If
    End Sub
    Private Sub Cleanup()
        DoLoop = False

        oSenderPipe.Dispose()
        oReceiverPipe.Dispose()
        oKeepAliveSenderPipe.Dispose()
        oKeepAliveReceiverPipe.Dispose()

        oSenderPipe = Nothing
        oReceiverPipe = Nothing
        oKeepAliveSenderPipe = Nothing
        oKeepAliveReceiverPipe = Nothing
    End Sub
    Private Sub Listener()
        Do
            Try
                If IsNothing(oStreamString) Then
                    DoLoop = False
                Else
                    'oStreamString.WriteString(Constants.KeepAliveString)
                    System.Threading.Thread.Sleep(Constants.KeepAliveTimer * 2)
                End If
            Catch e As IO.IOException
                ' if connection interrupted, then exit
                DoLoop = False
            End Try
        Loop While DoLoop
    End Sub
    Private Sub KeepAlive()
        Do
            Try
                If IsNothing(oKeepAliveStreamString) Then
                    DoLoop = False
                Else
                    oKeepAliveStreamString.WriteString(Constants.KeepAliveString)
                    System.Threading.Thread.Sleep(Constants.KeepAliveTimer * 2)
                End If
            Catch e As IO.IOException
                ' if connection interrupted, then exit
                DoLoop = False
            End Try
        Loop While DoLoop
    End Sub
    Public Function GetTWAINScannerSources() As List(Of Tuple(Of Enumerations.ScannerSource, String))
        Return Nothing
    End Function
    Public Sub ScannerConfigure()

    End Sub
#Region " IDisposable Support "
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Public Sub Close()
        Dispose()
    End Sub
    Private Overloads Sub Dispose(ByVal disposing As Boolean)
        SyncLock (Me)
            If disposing Then
                'dispose of unmanaged resources
                Cleanup()
            End If
        End SyncLock
    End Sub
#End Region
End Class
<DataContract> Public Class FieldDocumentStore
    Implements ICloneable

    <DataMember> Public FieldCollectionStore As New List(Of FieldCollection)
    <DataMember> Public PDFTemplate As Byte()

    Public Function Clone() As Object Implements ICloneable.Clone
        Dim oFieldDocumentStore As New FieldDocumentStore
        With oFieldDocumentStore
            .FieldCollectionStore.Clear()
            For Each oFieldCollection As FieldCollection In FieldCollectionStore
                .FieldCollectionStore.Add(oFieldCollection.Clone)
            Next

            .PDFTemplate = PDFTemplate.Clone
        End With
        Return oFieldDocumentStore
    End Function
    <DataContract> Public Class FieldCollection
        Implements ICloneable

        Public Event UpdateEvent()

        ' collection of input fields from a single form
        <DataMember> Public Fields As New List(Of Field)

        ' collection of individualised bar codes for each form page
        ' can be used To determine the number Of pages In the form
        <DataMember> Public BarCodes As New List(Of String)
        <DataMember> Public RawBarCodes As New List(Of Tuple(Of String, String, String))

        ' same as from the properties page
        <DataMember> Public FormTitle As String = String.Empty
        <DataMember> Public FormAuthor As String = String.Empty
        <DataMember> Public FormTopic As String = String.Empty

        ' the subject name for this form
        <DataMember> Public SubjectName As String = String.Empty

        ' accessory text
        <DataMember> Public AppendText As String = String.Empty

        <DataMember> Public DateCreated As Date = Date.MinValue

        <DataMember> Public MatrixStore As New MatrixStore

        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oFields As New FieldCollection
            With oFields
                .Fields.Clear()
                For Each oField As Field In Fields
                    .Fields.Add(oField.Clone)
                Next

                .BarCodes.Clear()
                For Each sBarcodeData As String In BarCodes
                    .BarCodes.Add(sBarcodeData)
                Next

                .RawBarCodes.Clear()
                For Each oRawBarCode As Tuple(Of String, String, String) In RawBarCodes
                    .RawBarCodes.Add(oRawBarCode)
                Next

                .FormTitle = FormTitle
                .FormAuthor = FormAuthor
                .FormTopic = FormTopic
                .SubjectName = SubjectName
                .AppendText = AppendText
                .DateCreated = DateCreated
                .MatrixStore = MatrixStore.Clone
            End With
            Return oFields
        End Function
        Public Sub CleanMatrix()
            ' runs through matrix store and cleans out extra images
            Dim oGUIDStore As List(Of Guid) = (From oField As Field In Fields From oImage In oField.Images Select oImage.Item2 Distinct).ToList
            MatrixStore.CleanMatrix(oGUIDStore)
        End Sub
        Public Sub Update()
            RaiseEvent UpdateEvent()
        End Sub
    End Class
    <DataContract()> Public Class Field
        Implements INotifyPropertyChanged, IEditableObject, ICloneable

        ' Data for undoing canceled edits.
        Private temp_Field As Field = Nothing
        Private m_Editing As Boolean = False
        Private m_Parent As FieldCollection = Nothing

#Region "INotifyPropertyChanged"
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub OnPropertyChangedLocal(ByVal sName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
        End Sub
#End Region
#Region "IEditableObject"
        Public Sub BeginEdit() Implements IEditableObject.BeginEdit
            If Not Me.m_Editing Then
                Me.temp_Field = Me.MemberwiseClone()
                Me.m_Editing = True
            End If
        End Sub
        Public Sub CancelEdit() Implements IEditableObject.CancelEdit
            If m_Editing = True Then
                Me.m_FieldType = Me.temp_Field.m_FieldType
                Me.m_GUID = Me.temp_Field.m_GUID
                Me.m_CharacterASCII = Me.temp_Field.m_CharacterASCII
                Me.m_Numbering = Me.temp_Field.m_Numbering
                Me.m_Order = Me.temp_Field.m_Order
                Me.m_Location = Me.temp_Field.m_Location
                Me.m_PageNumber = Me.temp_Field.m_PageNumber
                Me.m_Critical = Me.temp_Field.m_Critical
                Me.m_Images = Me.temp_Field.m_Images
                Me.m_TabletContent = Me.temp_Field.m_TabletContent
                Me.m_TabletStart = Me.temp_Field.m_TabletStart
                Me.m_TabletGroups = Me.temp_Field.m_TabletGroups

                Me.m_TabletDescriptionTop.Clear()
                Me.m_TabletDescriptionTop.AddRange(Me.temp_Field.m_TabletDescriptionTop)

                Me.m_TabletDescriptionBottom.Clear()
                Me.m_TabletDescriptionBottom.AddRange(Me.temp_Field.m_TabletDescriptionBottom)

                Me.m_TabletDescriptionMCQ.Clear()
                Me.m_TabletDescriptionMCQ.AddRange(Me.temp_Field.m_TabletDescriptionMCQ)

                Me.m_TabletMCQParams = Me.temp_Field.m_TabletMCQParams
                Me.m_TabletAlignment = Me.temp_Field.m_TabletAlignment
                Me.m_TabletSingleChoiceOnly = Me.temp_Field.m_TabletSingleChoiceOnly
                Me.m_TabletLimit = Me.temp_Field.m_TabletLimit

                Me.m_FormTitle = Me.temp_Field.m_FormTitle
                Me.m_SubjectName = Me.temp_Field.m_SubjectName
                Me.m_AppendText = Me.temp_Field.m_AppendText
                Me.m_RawBarCodes.Clear()
                Me.m_RawBarCodes.AddRange(Me.temp_Field.m_RawBarCodes)
                Me.m_MarkText = Me.temp_Field.m_MarkText

                Me.m_Editing = False
            End If
        End Sub
        Public Sub EndEdit() Implements IEditableObject.EndEdit
            If m_Editing = True Then
                Me.temp_Field = Nothing
                Me.m_Editing = False
            End If
        End Sub
#End Region
#Region "Data Members"
        ' each input field represents a single input field within a block

        ' field type
        <DataMember> Private m_GUID As Guid = GUID.Empty

        ' field type
        <DataMember> Private m_FieldType As Enumerations.FieldTypeEnum = Enumerations.FieldTypeEnum.Undefined

        ' for handwriting fields, this is the allowed input type
        <DataMember> Private m_CharacterASCII As Enumerations.CharacterASCII = Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric

        ' display numbering and order
        <DataMember> Private m_Numbering As String
        <DataMember> Private m_Order As Integer

        ' this is the location of the entire field within the page, excluding the margin
        <DataMember> Private m_Location As Rect

        ' this is the page number, zero based
        <DataMember> Private m_PageNumber As Integer = Integer.MinValue

        ' critical field for verification
        <DataMember> Private m_Critical As Boolean = False

        ' the rectangles denote the displacement in points from the page limit upper left along with width and height for each individual image
        ' these rectangles are actual field locations without taking into account the set margins
        ' in contrast, the bitmaps are extracted with the margin present around the rectangles
        ' the third item represents the mark for the individual image
        ' the fourth and fifth item represent the zero-based row and column numbers
        ' the sixth item indicates whether this is a fixed mark (ie. cannot be changed)
        ' the seventh item is the bitmap resolution
        ' the eighth item is the diagonal proportion (fDiagonal / fInnerDiagonal) from the scanned tablet
        <DataMember> Private m_Images As New List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))

        ' tablet descriptors
        <DataMember> Private m_TabletContent As Enumerations.TabletContentEnum = Enumerations.TabletContentEnum.None
        <DataMember> Private m_TabletStart As Integer = 0
        <DataMember> Private m_TabletGroups As Integer = 1
        <DataMember> Private m_TabletDescriptionTop As New List(Of Tuple(Of Rect, String))
        <DataMember> Private m_TabletDescriptionBottom As New List(Of Tuple(Of Rect, String))
        <DataMember> Private m_TabletDescriptionMCQ As New List(Of Tuple(Of Rect, Integer, Integer, List(Of ElementStruc)))
        <DataMember> Private m_TabletMCQParams As Tuple(Of Double, Double, Point, Int32Rect, Integer, Integer, List(Of Integer))
        <DataMember> Private m_TabletAlignment As Enumerations.AlignmentEnum = Enumerations.AlignmentEnum.Center
        <DataMember> Private m_TabletSingleChoiceOnly As Boolean = False
        <DataMember> Private m_TabletLimit As Double = 0

        ' items needed for scanner display
        <DataMember> Private m_FormTitle As String = String.Empty
        <DataMember> Private m_SubjectName As String = String.Empty
        <DataMember> Private m_AppendText As String = String.Empty
        <DataMember> Private m_RawBarCodes As New List(Of Tuple(Of String, String, String))
        <DataMember> Private m_MarkText As New List(Of Tuple(Of String, String, String))
#End Region
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oField As New Field
            With oField
                .m_FieldType = m_FieldType
                .m_GUID = m_GUID
                .m_CharacterASCII = m_CharacterASCII
                .m_Numbering = m_Numbering
                .m_Order = m_Order
                .m_Location = m_Location
                .m_PageNumber = m_PageNumber
                .m_Critical = m_Critical

                For Each oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images
                    .m_Images.Add(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))(oImage.Item1, oImage.Item2, oImage.Item3, oImage.Item4, oImage.Item5, oImage.Item6, oImage.Item7, oImage.Rest))
                Next

                .m_TabletContent = m_TabletContent
                .m_TabletStart = m_TabletStart
                .m_TabletGroups = m_TabletGroups
                .m_TabletDescriptionTop.AddRange(m_TabletDescriptionTop)
                .m_TabletDescriptionBottom.AddRange(m_TabletDescriptionBottom)
                .m_TabletDescriptionMCQ.AddRange(m_TabletDescriptionMCQ)
                .m_TabletMCQParams = m_TabletMCQParams
                .m_TabletAlignment = m_TabletAlignment
                .m_TabletSingleChoiceOnly = m_TabletSingleChoiceOnly
                .m_TabletLimit = m_TabletLimit

                .m_FormTitle = m_FormTitle
                .m_SubjectName = m_SubjectName
                .m_AppendText = m_AppendText
                .m_RawBarCodes.AddRange(m_RawBarCodes)
                .m_MarkText.AddRange(m_MarkText)
            End With
            Return oField
        End Function
#Region "Properties"
        Public Property Parent As FieldCollection
            Get
                Return m_Parent
            End Get
            Set(value As FieldCollection)
                m_Parent = value
            End Set
        End Property
        Public Property FieldType As Enumerations.FieldTypeEnum
            Get
                Return m_FieldType
            End Get
            Set(value As Enumerations.FieldTypeEnum)
                m_FieldType = value
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
        Public Property CharacterASCII As Enumerations.CharacterASCII
            Get
                Return m_CharacterASCII
            End Get
            Set(value As Enumerations.CharacterASCII)
                m_CharacterASCII = value
            End Set
        End Property
        Public Property Numbering As String
            Get
                Return m_Numbering
            End Get
            Set(value As String)
                m_Numbering = value
            End Set
        End Property
        Public Property Order As Integer
            Get
                Return m_Order
            End Get
            Set(value As Integer)
                m_Order = value
            End Set
        End Property
        Public ReadOnly Property OrderSort As String
            Get
                Return m_Order.ToString
            End Get
        End Property
        Public Property Location As Rect
            Get
                Return m_Location
            End Get
            Set(value As Rect)
                m_Location = value
                OnPropertyChangedLocal("Location")
            End Set
        End Property
        Public Property PageNumberText As String
            Get
                Return (m_PageNumber + 1).ToString
            End Get
            Set(value As String)
                If value <> (m_PageNumber + 1).ToString Then
                    m_PageNumber = CInt(Val(value)) - 1
                    OnPropertyChangedLocal("PageNumberText")
                    OnPropertyChangedLocal("PageNumber")
                End If
            End Set
        End Property
        Public Property PageNumber As Integer
            Get
                Return m_PageNumber
            End Get
            Set(value As Integer)
                If value <> m_PageNumber Then
                    m_PageNumber = value
                    OnPropertyChangedLocal("PageNumberText")
                    OnPropertyChangedLocal("PageNumber")
                End If
            End Set
        End Property
        Public Property Critical As Boolean
            Get
                Return m_Critical
            End Get
            Set(value As Boolean)
                m_Critical = value
                OnPropertyChangedLocal("Critical")
            End Set
        End Property
        Public Property TabletContent As Enumerations.TabletContentEnum
            Get
                Return m_TabletContent
            End Get
            Set(value As Enumerations.TabletContentEnum)
                If value <> m_TabletContent Then
                    m_TabletContent = value
                    OnPropertyChangedLocal("TabletContent")
                End If
            End Set
        End Property
        Public Property TabletStart As Integer
            Get
                Return m_TabletStart
            End Get
            Set(value As Integer)
                m_TabletStart = value
                OnPropertyChangedLocal("TabletStart")
            End Set
        End Property
        Public Property TabletGroups As Integer
            Get
                Return m_TabletGroups
            End Get
            Set(value As Integer)
                m_TabletGroups = value
                OnPropertyChangedLocal("TabletGroups")
            End Set
        End Property
        Public Property TabletDescriptionTop As List(Of Tuple(Of Rect, String))
            Get
                Return m_TabletDescriptionTop
            End Get
            Set(value As List(Of Tuple(Of Rect, String)))
                m_TabletDescriptionTop.Clear()
                m_TabletDescriptionTop.AddRange(value)
                OnPropertyChangedLocal("TabletDescriptionTop")
            End Set
        End Property
        Public Property TabletDescriptionBottom As List(Of Tuple(Of Rect, String))
            Get
                Return m_TabletDescriptionBottom
            End Get
            Set(value As List(Of Tuple(Of Rect, String)))
                m_TabletDescriptionBottom.Clear()
                m_TabletDescriptionBottom.AddRange(value)
                OnPropertyChangedLocal("TabletDescriptionBottom")
            End Set
        End Property
        Public Property TabletDescriptionMCQ As List(Of Tuple(Of Rect, Integer, Integer, List(Of ElementStruc)))
            Get
                Return m_TabletDescriptionMCQ
            End Get
            Set(value As List(Of Tuple(Of Rect, Integer, Integer, List(Of ElementStruc))))
                m_TabletDescriptionMCQ.Clear()
                m_TabletDescriptionMCQ.AddRange(value)
                OnPropertyChangedLocal("TabletDescriptionMCQ")
            End Set
        End Property
        Public Property TabletMCQParams As Tuple(Of Double, Double, Point, Int32Rect, Integer, Integer, List(Of Integer))
            Get
                Return m_TabletMCQParams
            End Get
            Set(value As Tuple(Of Double, Double, Point, Int32Rect, Integer, Integer, List(Of Integer)))
                m_TabletMCQParams = value
                OnPropertyChangedLocal("TabletMCQParams")
            End Set
        End Property
        Public Property TabletAlignment As Enumerations.AlignmentEnum
            Get
                Return m_TabletAlignment
            End Get
            Set(value As Enumerations.AlignmentEnum)
                If value <> m_TabletAlignment Then
                    m_TabletAlignment = value
                    OnPropertyChangedLocal("TabletAlignment")
                End If
            End Set
        End Property
        Public Property TabletSingleChoiceOnly As Boolean
            Get
                Return m_TabletSingleChoiceOnly
            End Get
            Set(value As Boolean)
                m_TabletSingleChoiceOnly = value
                OnPropertyChangedLocal("TabletSingleChoiceOnly")
            End Set
        End Property
        Public Property TabletLimit As Double
            Get
                Return m_TabletLimit
            End Get
            Set(value As Double)
                If value <> m_TabletLimit Then
                    m_TabletLimit = value
                    OnPropertyChangedLocal("TabletLimit")
                End If
            End Set
        End Property
        Public Property FormTitle As String
            Get
                Return m_FormTitle
            End Get
            Set(value As String)
                If value <> m_FormTitle Then
                    m_FormTitle = value
                    OnPropertyChangedLocal("FormTitle")
                End If
            End Set
        End Property
        Public Property SubjectName As String
            Get
                Return If(m_AppendText = String.Empty, m_SubjectName, m_SubjectName + " (" + m_AppendText + ")")
            End Get
            Set(value As String)
                If value <> m_SubjectName Then
                    m_SubjectName = value
                    OnPropertyChangedLocal("SubjectName")
                End If
            End Set
        End Property
#End Region
#Region "Images"
        Public ReadOnly Property ImageCount As Integer
            Get
                Return m_Images.Count
            End Get
        End Property
        Public ReadOnly Property Images As List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))
            Get
                Dim oReturnImages As New List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)))
                For Each oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images
                    oReturnImages.Add(oImage)
                Next
                Return oReturnImages
            End Get
        End Property
        Public Function GetImage(ByVal iIndex As Integer) As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))
            If iIndex >= 0 And iIndex < m_Images.Count Then
                Return New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(m_Images(iIndex).Item1, m_Images(iIndex).Item2, m_Images(iIndex).Item3, m_Images(iIndex).Item4, m_Images(iIndex).Item5, m_Images(iIndex).Item6, m_Images(iIndex).Item7, New Tuple(Of Single)(m_Images(iIndex).Rest.Item1))
            Else
                Return Nothing
            End If
        End Function
        Public Sub SetImage(ByVal iIndex As Integer, ByVal oValue As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single)), Optional ByVal oMatrix As Emgu.CV.Matrix(Of Byte) = Nothing, Optional ByVal oMatrixStore As MatrixStore = Nothing)
            If iIndex >= 0 And iIndex < m_Images.Count Then
                If Not IsNothing(oMatrixStore) Then
                    oMatrixStore.SetMatrix(oValue.Item2, m_Images(iIndex).Item2, m_Images(iIndex).Rest.Item2, oMatrix)
                End If
                m_Images(iIndex) = New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))(oValue.Item1, oValue.Item2, oValue.Item3, oValue.Item4, oValue.Item5, oValue.Item6, oValue.Item7, New Tuple(Of Single, Guid)(oValue.Rest.Item1, m_Images(iIndex).Rest.Item2))
            End If
        End Sub
        Public Sub RemoveImage(ByVal iIndex As Integer, Optional ByVal oMatrixStore As MatrixStore = Nothing)
            If iIndex >= 0 And iIndex < m_Images.Count Then
                If Not IsNothing(oMatrixStore) Then
                    oMatrixStore.MatrixCleared(m_Images(iIndex).Item2, m_Images(iIndex).Rest.Item2)
                End If
                m_Images.RemoveAt(iIndex)
            End If
        End Sub
        Public Sub AddImage(ByVal oValue As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single)), Optional ByVal oMatrix As Emgu.CV.Matrix(Of Byte) = Nothing, Optional ByVal oMatrixStore As MatrixStore = Nothing)
            Dim oGUIDSource As Guid = GUID.NewGuid
            If Not IsNothing(oMatrixStore) Then
                oMatrixStore.SetMatrix(oValue.Item2, Nothing, oGUIDSource, oMatrix)
            End If
            m_Images.Add(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))(oValue.Item1, oValue.Item2, oValue.Item3, oValue.Item4, oValue.Item5, oValue.Item6, oValue.Item7, New Tuple(Of Single, Guid)(oValue.Rest.Item1, oGUIDSource)))
        End Sub
#End Region
#Region "Marks"
        Const DataPresentNone As String = "None"
        Const DataPresentFull As String = "Full"
        Const DataPresentPartial As String = "Partial"
        Public Enum DataPresentEnum As Integer
            DataNone = 0
            DataPartial
            DataFull
        End Enum
        Public Property DataPresent As DataPresentEnum
            Get
                ' no data is present if all marks are empty 
                ' partial data is present if either the detected or verified marks is present but the final mark is empty
                Dim iTotalDMarkCount As Integer = 0
                Select Case m_FieldType
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        Dim oBoxIndexList As List(Of Tuple(Of Integer, Boolean)) = (From iIndex As Integer In Enumerable.Range(0, m_Images.Count) Where m_Images(iIndex).Item5 = -1 Select New Tuple(Of Integer, Boolean)(iIndex, m_Images(iIndex).Item6)).ToList
                        iTotalDMarkCount = Aggregate iIndex In Enumerable.Range(0, oBoxIndexList.Count) Where Trim(MarkBoxChoice0(iIndex)) <> String.Empty Or Trim(MarkBoxChoice1(iIndex)) <> String.Empty Or Trim(MarkBoxChoice2(iIndex)) <> String.Empty And Not oBoxIndexList(iIndex).Item2 Into Count()
                        If iTotalDMarkCount = 0 Then
                            Return DataPresentNone
                        ElseIf (Aggregate iIndex In Enumerable.Range(0, oBoxIndexList.Count) Where ((Trim(MarkBoxChoice0(iIndex)) <> String.Empty Or Trim(MarkBoxChoice1(iIndex)) <> String.Empty) And Trim(MarkBoxChoice2(iIndex)) = String.Empty) And Not oBoxIndexList(iIndex).Item2 Into Count()) > 0 Then
                            Return DataPresentPartial
                        Else
                            Return DataPresentFull
                        End If
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        iTotalDMarkCount = Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where MarkChoice0(iIndex) Or MarkChoice1(iIndex) Or MarkChoice2(iIndex) Into Count()
                        If iTotalDMarkCount = 0 Then
                            Return DataPresentNone
                        ElseIf (Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where (MarkChoice0(iIndex) Or MarkChoice1(iIndex)) And (Not MarkChoice2(iIndex)) Into Count()) > 0 Then
                            Return DataPresentPartial
                        Else
                            Return DataPresentFull
                        End If
                    Case Enumerations.FieldTypeEnum.Handwriting
                        iTotalDMarkCount = Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where Trim(MarkHandwriting0(iIndex)) <> String.Empty Or Trim(MarkHandwriting1(iIndex)) <> String.Empty Or Trim(MarkHandwriting2(iIndex)) <> String.Empty And Not m_Images(iIndex).Item6 Into Count()
                        If iTotalDMarkCount = 0 Then
                            Return DataPresentNone
                        ElseIf (Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where (Trim(MarkHandwriting0(iIndex)) <> String.Empty Or Trim(MarkHandwriting1(iIndex)) <> String.Empty) And Trim(MarkHandwriting2(iIndex)) = String.Empty And Not m_Images(iIndex).Item6 Into Count()) > 0 Then
                            Return DataPresentPartial
                        Else
                            Return DataPresentFull
                        End If
                    Case Enumerations.FieldTypeEnum.Free
                        If Trim(MarkFree0) = String.Empty And Trim(MarkFree1) = String.Empty And Trim(MarkFree2) = String.Empty Then
                            Return DataPresentNone
                        ElseIf (Trim(MarkFree0) <> String.Empty Or Trim(MarkFree1) <> String.Empty) And Trim(MarkFree2) = String.Empty Then
                            Return DataPresentPartial
                        Else
                            Return DataPresentFull
                        End If
                    Case Else
                        Return DataPresentNone
                End Select
            End Get
            Set(value As DataPresentEnum)
            End Set
        End Property
        Public ReadOnly Property MarkCount As Integer
            Get
                Return m_MarkText.Count
            End Get
        End Property
        Public Property MarkBoxChoice0(ByVal iIndex As Integer) As String
            Get
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    Dim sChar As String = Left(m_MarkText(iIndex).Item1, 1)
                    If IsNumeric(Val(sChar)) Then
                        Return sChar
                    Else
                        Return String.Empty
                    End If
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                Dim sValue As String = Trim(Left(value, 1))
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    If (sValue = String.Empty Or IsNumeric(sValue)) And m_MarkText(iIndex).Item1 <> sValue Then
                        m_MarkText(iIndex) = New Tuple(Of String, String, String)(sValue, m_MarkText(iIndex).Item2, m_MarkText(iIndex).Item3)
                        SetFinal(iIndex)
                        OnPropertyChangedLocal("MarkBoxChoice0")
                        UpdateImages()
                    End If
                End If
            End Set
        End Property
        Public Property MarkBoxChoice1(ByVal iIndex As Integer) As String
            Get
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    Dim sChar As String = Left(m_MarkText(iIndex).Item2, 1)
                    If IsNumeric(Val(sChar)) Then
                        Return sChar
                    Else
                        Return String.Empty
                    End If
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                Dim sValue As String = Trim(Left(value, 1))
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    If (sValue = String.Empty Or IsNumeric(sValue)) And m_MarkText(iIndex).Item2 <> sValue Then
                        m_MarkText(iIndex) = New Tuple(Of String, String, String)(m_MarkText(iIndex).Item1, sValue, m_MarkText(iIndex).Item3)
                        SetFinal(iIndex)
                        OnPropertyChangedLocal("MarkBoxChoice1")
                        UpdateImages()
                    End If
                End If
            End Set
        End Property
        Public Property MarkBoxChoice2(ByVal iIndex As Integer) As String
            Get
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    Dim sChar As String = Left(m_MarkText(iIndex).Item3, 1)
                    If IsNumeric(Val(sChar)) Then
                        Return sChar
                    Else
                        Return String.Empty
                    End If
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                Dim sValue As String = Trim(Left(value, 1))
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    If (sValue = String.Empty Or IsNumeric(sValue)) And m_MarkText(iIndex).Item3 <> sValue Then
                        m_MarkText(iIndex) = New Tuple(Of String, String, String)(m_MarkText(iIndex).Item1, m_MarkText(iIndex).Item2, sValue)
                        OnPropertyChangedLocal("DataPresent")
                        OnPropertyChangedLocal("MarkBoxChoice2")
                        UpdateImages()
                    End If
                End If
            End Set
        End Property
        Public Property MarkChoice0(ByVal iIndex As Integer) As Boolean
            Get
                If m_MarkText.Count = 0 Then
                    Return False
                Else
                    Dim iActualIndex As Integer = iIndex
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > m_MarkText.Count - 1 Then
                        iActualIndex = m_MarkText.Count - 1
                    End If

                    Return If(Trim(m_MarkText(iActualIndex).Item1) = String.Empty, False, True)
                End If
            End Get
            Set(value As Boolean)
                If m_MarkText.Count > 0 Then
                    Dim iActualIndex As Integer = iIndex
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > m_MarkText.Count - 1 Then
                        iActualIndex = m_MarkText.Count - 1
                    End If

                    If value Then
                        ' clear off all marks first
                        If m_TabletSingleChoiceOnly And (Not IsNothing(Parent)) Then
                            For i = 0 To m_MarkText.Count - 1
                                m_MarkText(i) = New Tuple(Of String, String, String)(String.Empty, m_MarkText(i).Item2, m_MarkText(i).Item3)
                            Next
                        End If

                        m_MarkText(iActualIndex) = New Tuple(Of String, String, String)("X", m_MarkText(iActualIndex).Item2, m_MarkText(iActualIndex).Item3)
                    Else
                        m_MarkText(iActualIndex) = New Tuple(Of String, String, String)(String.Empty, m_MarkText(iActualIndex).Item2, m_MarkText(iActualIndex).Item3)
                    End If
                End If
                SetFinal(iIndex)
                OnPropertyChangedLocal(Data.Binding.IndexerName)
                UpdateImages()
            End Set
        End Property
        Public Property MarkChoice1(ByVal iIndex As Integer) As Boolean
            Get
                If m_MarkText.Count = 0 Then
                    Return False
                Else
                    Dim iActualIndex As Integer = iIndex
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > m_MarkText.Count - 1 Then
                        iActualIndex = m_MarkText.Count - 1
                    End If

                    Return If(Trim(m_MarkText(iActualIndex).Item2) = String.Empty, False, True)
                End If
            End Get
            Set(value As Boolean)
                If m_MarkText.Count > 0 Then
                    Dim iActualIndex As Integer = iIndex
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > m_MarkText.Count - 1 Then
                        iActualIndex = m_MarkText.Count - 1
                    End If

                    If value Then
                        ' clear off all marks first
                        If m_TabletSingleChoiceOnly And (Not IsNothing(Parent)) Then
                            For i = 0 To m_MarkText.Count - 1
                                m_MarkText(i) = New Tuple(Of String, String, String)(m_MarkText(i).Item1, String.Empty, m_MarkText(i).Item3)
                            Next
                        End If

                        m_MarkText(iActualIndex) = New Tuple(Of String, String, String)(m_MarkText(iActualIndex).Item1, "X", m_MarkText(iActualIndex).Item3)
                    Else
                        m_MarkText(iActualIndex) = New Tuple(Of String, String, String)(m_MarkText(iActualIndex).Item1, String.Empty, m_MarkText(iActualIndex).Item3)
                    End If
                End If
                SetFinal(iIndex)
                OnPropertyChangedLocal(Data.Binding.IndexerName)
                UpdateImages()
            End Set
        End Property
        Public Property MarkChoice2(ByVal iIndex As Integer) As Boolean
            Get
                If m_MarkText.Count = 0 Then
                    Return False
                Else
                    Dim iActualIndex As Integer = iIndex
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > m_MarkText.Count - 1 Then
                        iActualIndex = m_MarkText.Count - 1
                    End If

                    Return If(Trim(m_MarkText(iActualIndex).Item3) = String.Empty, False, True)
                End If
            End Get
            Set(value As Boolean)
                If m_MarkText.Count > 0 Then
                    Dim iActualIndex As Integer = iIndex
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > m_MarkText.Count - 1 Then
                        iActualIndex = m_MarkText.Count - 1
                    End If

                    If value Then
                        ' clear off all marks first
                        If m_TabletSingleChoiceOnly And (Not IsNothing(Parent)) Then
                            For i = 0 To m_MarkText.Count - 1
                                m_MarkText(i) = New Tuple(Of String, String, String)(m_MarkText(i).Item1, m_MarkText(i).Item2, String.Empty)
                            Next
                        End If

                        m_MarkText(iActualIndex) = New Tuple(Of String, String, String)(m_MarkText(iActualIndex).Item1, m_MarkText(iActualIndex).Item2, "X")
                    Else
                        m_MarkText(iActualIndex) = New Tuple(Of String, String, String)(m_MarkText(iActualIndex).Item1, m_MarkText(iActualIndex).Item2, String.Empty)
                    End If
                End If
                OnPropertyChangedLocal("DataPresent")
                OnPropertyChangedLocal(Data.Binding.IndexerName)
                UpdateImages()
            End Set
        End Property
        Public Property MarkHandwriting0(ByVal iIndex As Integer) As String
            Get
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    Return m_MarkText(iIndex).Item1
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    If m_MarkText(iIndex).Item1 <> value Then
                        m_MarkText(iIndex) = New Tuple(Of String, String, String)(Left(Trim(value), 1), m_MarkText(iIndex).Item2, m_MarkText(iIndex).Item3)
                        SetFinal(iIndex)
                        OnPropertyChangedLocal(Data.Binding.IndexerName)
                        UpdateImages()
                    End If
                End If
            End Set
        End Property
        Public Property MarkHandwriting1(ByVal iIndex As Integer) As String
            Get
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    Return m_MarkText(iIndex).Item2
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    If m_MarkText(iIndex).Item2 <> value Then
                        m_MarkText(iIndex) = New Tuple(Of String, String, String)(m_MarkText(iIndex).Item1, Left(Trim(value), 1), m_MarkText(iIndex).Item3)
                        SetFinal(iIndex)
                        OnPropertyChangedLocal(Data.Binding.IndexerName)
                        UpdateImages()
                    End If
                End If
            End Set
        End Property
        Public Property MarkHandwriting2(ByVal iIndex As Integer) As String
            Get
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    Return m_MarkText(iIndex).Item3
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_MarkText.Count > 0 AndAlso (iIndex >= 0 And iIndex <= m_MarkText.Count - 1) Then
                    If m_MarkText(iIndex).Item3 <> value Then
                        m_MarkText(iIndex) = New Tuple(Of String, String, String)(m_MarkText(iIndex).Item1, m_MarkText(iIndex).Item2, Left(Trim(value), 1))
                        OnPropertyChangedLocal("DataPresent")
                        OnPropertyChangedLocal(Data.Binding.IndexerName)
                        UpdateImages()
                    End If
                End If
            End Set
        End Property
        Public Property MarkFree0 As String
            Get
                If m_MarkText.Count > 0 Then
                    Return m_MarkText(0).Item1
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_MarkText.Count = 0 OrElse value <> m_MarkText(0).Item1 Then
                    Dim oMarkTuple As New Tuple(Of String, String, String)(value, If(m_Critical, If(m_MarkText.Count = 0, String.Empty, m_MarkText(0).Item2), value), If(m_Critical, If(m_MarkText.Count = 0, String.Empty, m_MarkText(0).Item3), value))
                    m_MarkText.Clear()
                    If value <> String.Empty Then
                        m_MarkText.Add(oMarkTuple)
                    End If
                    SetFinal(0)
                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                    UpdateImages()
                End If
            End Set
        End Property
        Public Property MarkFree1 As String
            Get
                If m_MarkText.Count > 0 Then
                    Return m_MarkText(0).Item2
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_MarkText.Count = 0 OrElse value <> m_MarkText(0).Item2 Then
                    Dim oMarkTuple As New Tuple(Of String, String, String)(If(m_Critical, If(m_MarkText.Count = 0, String.Empty, m_MarkText(0).Item1), value), value, If(m_Critical, If(m_MarkText.Count = 0, String.Empty, m_MarkText(0).Item3), value))
                    m_MarkText.Clear()
                    If value <> String.Empty Then
                        m_MarkText.Add(oMarkTuple)
                    End If
                    SetFinal(0)
                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                    UpdateImages()
                End If
            End Set
        End Property
        Public Property MarkFree2 As String
            Get
                If m_MarkText.Count > 0 Then
                    Return m_MarkText(0).Item3
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_MarkText.Count = 0 OrElse value <> m_MarkText(0).Item3 Then
                    Dim oMarkTuple As New Tuple(Of String, String, String)(If(m_MarkText.Count = 0, String.Empty, m_MarkText(0).Item1), If(m_MarkText.Count = 0, String.Empty, m_MarkText(0).Item2), value)
                    m_MarkText.Clear()
                    If value <> String.Empty Then
                        m_MarkText.Add(oMarkTuple)
                    End If
                    OnPropertyChangedLocal("DataPresent")
                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                    UpdateImages()
                End If
            End Set
        End Property
        Private Sub UpdateImages()
            If Not IsNothing(Parent) Then
                Parent.Update()
            End If
        End Sub
        Private Sub SetFinal(ByVal iIndex As Integer)
            ' if the detected and verified are in agreement, the final mark is set
            ' if not in agreement, set the final mark to blank
            Select Case m_FieldType
                Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    If MarkChoice0(iIndex) = MarkChoice1(iIndex) Then
                        MarkChoice2(iIndex) = MarkChoice0(iIndex)
                    Else
                        MarkChoice2(iIndex) = False
                    End If
                Case Enumerations.FieldTypeEnum.BoxChoice
                    If MarkBoxChoice0(iIndex) = MarkBoxChoice1(iIndex) Then
                        MarkBoxChoice2(iIndex) = MarkBoxChoice0(iIndex)
                    Else
                        MarkBoxChoice2(iIndex) = String.Empty
                    End If
                Case Enumerations.FieldTypeEnum.Handwriting
                    If MarkHandwriting0(iIndex) = MarkHandwriting1(iIndex) Then
                        MarkHandwriting2(iIndex) = MarkHandwriting0(iIndex)
                    Else
                        MarkHandwriting2(iIndex) = String.Empty
                    End If
                Case Enumerations.FieldTypeEnum.Free
                    If MarkFree0 = MarkFree1 Then
                        MarkFree2 = MarkFree0
                    Else
                        MarkFree2 = String.Empty
                    End If
            End Select
            OnPropertyChangedLocal("DataPresent")
        End Sub
#End Region
    End Class
    <DataContract> Public Class MatrixStore
        Implements ICloneable, IDisposable

        ' the store consists of a guid indexed dictionary of a compressed byte array, width, height, and the references to the object
        <DataMember> Private m_Store As New ConcurrentDictionary(Of Guid, Tuple(Of Byte(), Integer, Integer, List(Of Guid)))

        Public Function GetMatrix(ByVal oGUID As Guid) As Emgu.CV.Matrix(Of Byte)
            ' gets matrix from store
            If m_Store.ContainsKey(oGUID) Then
                Return Converter.ArrayToMatrix(CommonFunctions.Decompress(m_Store(oGUID).Item1), m_Store(oGUID).Item2, m_Store(oGUID).Item3)
            Else
                Return Nothing
            End If
        End Function
        Public Sub SetMatrix(ByVal oNewGUID As Guid, ByVal oOldGUID As Guid, ByVal oGUIDSource As Guid, Optional oMatrix As Emgu.CV.Matrix(Of Byte) = Nothing)
            ' stores matrix
            If (Not IsNothing(oOldGUID)) AndAlso m_Store.ContainsKey(oOldGUID) Then
                MatrixCleared(oOldGUID, oGUIDSource)
            End If

            ' proceed only if the guid is valid
            If Not oNewGUID.Equals(Guid.Empty) Then
                If Not m_Store.ContainsKey(oNewGUID) Then
                    m_Store.TryAdd(oNewGUID, New Tuple(Of Byte(), Integer, Integer, List(Of Guid))(Nothing, 0, 0, New List(Of Guid)))
                End If

                ' add guid source if not present
                Dim oCurrentGUIDList As List(Of Guid) = m_Store(oNewGUID).Item4
                If Not oCurrentGUIDList.Contains(oGUIDSource) Then
                    oCurrentGUIDList.Add(oGUIDSource)
                End If

                If IsNothing(oMatrix) Then
                    ' if no matrix supplied, then just update guid source list
                    m_Store(oNewGUID) = New Tuple(Of Byte(), Integer, Integer, List(Of Guid))(m_Store(oNewGUID).Item1, m_Store(oNewGUID).Item2, m_Store(oNewGUID).Item3, oCurrentGUIDList)
                Else
                    ' create a new byte store if not present
                    m_Store(oNewGUID) = New Tuple(Of Byte(), Integer, Integer, List(Of Guid))(If(IsNothing(m_Store(oNewGUID).Item1), CommonFunctions.Compress(Converter.MatrixToArray(oMatrix)), m_Store(oNewGUID).Item1), oMatrix.Width, oMatrix.Height, oCurrentGUIDList)
                End If
            End If
        End Sub
        Public Sub ReplaceMatrix(ByVal oGUID As Guid, oMatrix As Emgu.CV.Matrix(Of Byte))
            ' replaces the existing matrix but leaves all references unchanged
            If (Not IsNothing(oGUID)) AndAlso m_Store.ContainsKey(oGUID) Then
                m_Store(oGUID) = New Tuple(Of Byte(), Integer, Integer, List(Of Guid))(CommonFunctions.Compress(Converter.MatrixToArray(oMatrix)), oMatrix.Width, oMatrix.Height, m_Store(oGUID).Item4)
            End If
        End Sub
        Public Function GetMat(ByVal oGUID As Guid) As Emgu.CV.Mat
            ' gets matrix from store
            If m_Store.ContainsKey(oGUID) Then
                Return Converter.ArrayToMatrix(CommonFunctions.Decompress(m_Store(oGUID).Item1), m_Store(oGUID).Item2, m_Store(oGUID).Item3).Mat
            Else
                Return Nothing
            End If
        End Function
        Public Function GetMatrixWidth(ByVal oGUID As Guid) As Integer
            ' gets matrix width from store
            If m_Store.ContainsKey(oGUID) Then
                Return m_Store(oGUID).Item2
            Else
                Return 0
            End If
        End Function
        Public Function GetMatrixHeight(ByVal oGUID As Guid) As Integer
            ' gets matrix height from store
            If m_Store.ContainsKey(oGUID) Then
                Return m_Store(oGUID).Item3
            Else
                Return 0
            End If
        End Function
        Public Sub MatrixCleared(ByVal oGUID As Guid, ByVal oGUIDSource As Guid)
            ' clears matrix from store
            If m_Store.ContainsKey(oGUID) Then
                ' remove the source guid
                If m_Store(oGUID).Item4.Contains(oGUIDSource) Then
                    m_Store(oGUID).Item4.Remove(oGUIDSource)
                End If

                ' if no more guid reference present, then remove the matrix
                If m_Store(oGUID).Item4.Count = 0 Then
                    m_Store.TryRemove(oGUID, Nothing)
                End If
            End If
        End Sub
        Public Sub ClearAll()
            ' clears all store contents
            m_Store.Clear()
        End Sub
        Public Sub Transfer(ByRef oMatrixStore As MatrixStore)
            ' transfers the contents of this matrixstore to the new matrix store
            If IsNothing(oMatrixStore) Then
                oMatrixStore = New MatrixStore
            End If
            With oMatrixStore
                .ClearAll()

                ' add keys first
                For Each oGUID As Guid In m_Store.Keys
                    .m_Store.TryAdd(oGUID, Nothing)
                Next

                ' transfer values
                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn
                                                            Dim oBufferIn As MatrixStore = CType(oParam.Item4, MatrixStore)
                                                            Dim oBufferOut As MatrixStore = CType(oParam.Item5, MatrixStore)

                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                Dim oGUID As Guid = oBufferOut.m_Store.Keys(localY)
                                                                Dim oItem As Tuple(Of Byte(), Integer, Integer, List(Of Guid)) = Nothing
                                                                oBufferIn.m_Store.TryRemove(oGUID, oItem)
                                                                oBufferOut.m_Store.TryUpdate(oGUID, oItem, Nothing)
                                                            Next
                                                        End Sub
                CommonFunctions.ParallelRun(m_Store.Count, Me, oMatrixStore, Nothing, Nothing, TaskDelegate)
            End With
        End Sub
        Public Sub CleanMatrix(ByVal oGUIDStore As List(Of Guid))
            ' runs through matrix store and cleans out extra images
            Dim oMatrixGUIDList As List(Of Guid) = m_Store.Keys.ToList
            For Each oGUID As Guid In oMatrixGUIDList
                If Not oGUIDStore.Contains(oGUID) Then
                    m_Store.TryRemove(oGUID, Nothing)
                End If
            Next
        End Sub
        Public Function Clone() As Object Implements ICloneable.Clone
            CommonFunctions.ClearMemory()
            Dim oMatrixStore As New MatrixStore
            With oMatrixStore
                For Each oGUID As Guid In m_Store.Keys
                    .m_Store.TryAdd(oGUID, New Tuple(Of Byte(), Integer, Integer, List(Of Guid))(m_Store(oGUID).Item1.Clone, m_Store(oGUID).Item2, m_Store(oGUID).Item3, New List(Of Guid)(m_Store(oGUID).Item4)))
                Next
            End With
            Return oMatrixStore
        End Function
        Public Function Efficiency() As Double
            ' returns the compression efficiency of the whole store
            Dim iTotalBytesCompressed As Integer = 0
            Dim iTotalBytes As Integer = 0
            Dim iStoreCount As Integer = m_Store.Count

            Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                        Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                        For y = 0 To oParam.Item2 - 1
                                                            Dim localY As Integer = y + oParam.Item1

                                                            Dim oGUID As Guid = m_Store.Keys(localY)
                                                            Dim oItem As Tuple(Of Byte(), Integer, Integer, List(Of Guid)) = m_Store(oGUID)

                                                            System.Threading.Interlocked.Add(iTotalBytesCompressed, oItem.Item1.Count)
                                                            System.Threading.Interlocked.Add(iTotalBytes, oItem.Item2 * oItem.Item3)
                                                        Next
                                                    End Sub
            CommonFunctions.ParallelRun(iStoreCount, Nothing, Nothing, Nothing, Nothing, TaskDelegate)

            Return iTotalBytesCompressed / iTotalBytes
        End Function
        Public Function Efficiency(ByVal oGUID As Guid) As Double
            ' returns the compression efficiency as a fraction for the byte store
            Dim oItem As Tuple(Of Byte(), Integer, Integer, List(Of Guid)) = m_Store(oGUID)
            Return oItem.Item1.Count / (oItem.Item2 * oItem.Item3)
        End Function
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                End If
                m_Store = Nothing
            End If
            disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
End Class

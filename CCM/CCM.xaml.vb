Imports PdfSharp
Imports PdfSharp.Drawing
Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports Common.Common
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Media
Imports System.Runtime.Serialization

Public Class CCM
    Implements FunctionInterface

#Region "Variables"
    ' PluginName defines the friendly name of the plugin
    ' Priority determines the order in which the buttons are arranged on the main page. The lower the number, the earlier it is placed
    Private Const PluginName As String = "CCM"
    Private Const Priority As Integer = 1
    Private Const FieldTypeMultiplier As Double = 0.1
    Private Const FieldTypeAspectRatio As Double = 4
    Private Const WorksheetSubjects As String = "Subjects"
    Private Const StringSubject As String = "[Subject]"
    Private Const StringTag As String = "[Tag]"
    Private Const StringMCQ As String = "[MCQ]"
    Private Const DragBaseFormItem As String = "BaseFormItem"
    Private Shared m_Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer))
    Private Shared ButtonDictionary As New Dictionary(Of String, FrameworkElement)
    Private Shared DisplayDictionary As New Dictionary(Of String, FrameworkElement)
    Private Shared DisplayAdd As New List(Of String) From {"RectangleBlockBackground", "ImageBlockContent", "CanvasBlockContent", "ImageBlockViewer", "DataGridSubjectsTags", "ScrollViewerBlockContent"}
    Private Shared m_Icons As Dictionary(Of String, ImageSource)
    Private Shared oImageStore As ImageStore
    Private Shared SelectedItem As BaseFormItem
    Private WithEvents m_RectangleFormScrollUp As Shapes.Rectangle
    Private WithEvents m_RectangleFormScrollDown As Shapes.Rectangle

    ' insert your own GUIDs here
    Private Shared oMainFormGUID As New Guid("{d49665a8-bc40-4aff-8b24-854e623cb966}")
#End Region
#Region "Constants"
    Public Const BorderSpacerMultiplier As Double = 0.02
    Public Const CornerRadiusMultiplier As Double = 0.01
#End Region
#Region "FunctionInterface"
    Public Function GetDataTypes() As List(Of Tuple(Of Guid, Type)) Implements FunctionInterface.GetDataTypes
        ' returns a list of GUIDs and variable type representing the data types that the plug-in creates
        Dim oDataTypes As New List(Of Tuple(Of Guid, Type))
        oDataTypes.Add(New Tuple(Of Guid, Type)(oMainFormGUID, GetType(FormMain)))

        Return oDataTypes
    End Function
    Public Function CheckDataTypes(ByRef oCommonVariables As CommonVariables) As Boolean Implements FunctionInterface.CheckDataTypes
        ' checks the commonvariable data store to see if the variables objects required have been properly initialised
        Dim bCheck As Boolean = True

        ' initialise FormMain
        If oCommonVariables.DataStore(oMainFormGUID).GetType.Equals(GetType(FormMain)) Then
            PostInit()
        Else
            bCheck = False
        End If

        Return bCheck
    End Function
    Public Function GetIdentifier() As Tuple(Of Guid, ImageSource, String, Integer) Implements FunctionInterface.GetIdentifier
        ' returns the identifiers: GUID, icon, friendly name, and priority for the plugin
        Return New Tuple(Of Guid, ImageSource, String, Integer)(Guid.NewGuid, Converter.BitmapToBitmapSource(My.Resources.IconCCM), PluginName, Priority)
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

        BaseFormItem.Root = Me

        BaseFormItem.Suspender.GlobalSuspendUpdates = True

        ' Add any initialization after the InitializeComponent() call.
        CommonFunctions.IterateElementDictionary(ButtonDictionary, GridMain, "Button")

        Dim oGridContents As Controls.Grid = BaseFormItem.Root.GridContents
        IterateDisplayDictionary(oGridContents)
        SelectedItem = Nothing

        ' initially set all buttons to disabled and clear tooltips
        For i = 0 To 2
            Dim oButton As Controls.Button = ButtonDictionary("Button" + i.ToString)
            Dim oImage As Controls.Image = ButtonDictionary("ButtonImage" + i.ToString)
            oButton.ToolTip = String.Empty
            oButton.IsEnabled = False
            oImage.Source = Nothing
        Next

        SetIcons()
    End Sub
    Private Sub PostInit()
        ' runs post-initialisation routine
        BaseFormItem.FormMain = CType(oCommonVariables.DataStore(oMainFormGUID), FormMain)

        Dim oFormProperties As New FormProperties
        oFormProperties.Parent = BaseFormItem.FormMain

        Dim oFormHeader As New FormHeader
        oFormHeader.Parent = BaseFormItem.FormMain

        Dim oFormBody As New FormBody
        oFormBody.Parent = BaseFormItem.FormMain

        Dim oFormFooter As New FormFooter
        oFormFooter.Parent = BaseFormItem.FormMain

        Dim oFormPDF As New FormPDF
        oFormPDF.Parent = BaseFormItem.FormMain

        Dim oFormExport As New FormExport
        oFormExport.Parent = BaseFormItem.FormMain

        FormBlock.CanvasBlockContent = CanvasBlockContent

        If IsNothing(FormSection.ImageViewer) Then
            FormSection.ImageViewer = New FormSection.ImageViewerClass
        Else
            FormSection.ImageViewer.Clear()
        End If

        m_RectangleFormScrollUp = RectangleFormScrollUp
        m_RectangleFormScrollDown = BaseFormItem.Root.RectangleFormScrollDown

        SetBindings()

        BaseFormItem.Suspender.GlobalSuspendUpdates = False



    End Sub
    Private Sub SetBindings()
        BaseFormItem.FormProperties = BaseFormItem.FormMain.GetFormItems(Of FormProperties).First
        BaseFormItem.FormHeader = BaseFormItem.FormMain.GetFormItems(Of FormHeader).First
        BaseFormItem.FormPageHeader = BaseFormItem.FormMain.GetFormItems(Of FormPageHeader).First
        BaseFormItem.FormFormHeader = BaseFormItem.FormMain.GetFormItems(Of FormFormHeader).First
        BaseFormItem.FormBody = BaseFormItem.FormMain.GetFormItems(Of FormBody).First
        BaseFormItem.FormFooter = BaseFormItem.FormMain.GetFormItems(Of FormFooter).First
        BaseFormItem.FormPDF = BaseFormItem.FormMain.GetFormItems(Of FormPDF).First
        BaseFormItem.FormExport = BaseFormItem.FormMain.GetFormItems(Of FormExport).First

        BaseFormItem.FormProperties.SetBindings()
        BaseFormItem.FormHeader.SetBindings()
        BaseFormItem.FormPageHeader.SetBindings()
        BaseFormItem.FormFormHeader.SetBindings()
        BaseFormItem.FormBody.SetBindings()
        BaseFormItem.FormFooter.SetBindings()
        BaseFormItem.FormPDF.SetBindings()
        BaseFormItem.FormExport.SetBindings()
    End Sub
    Private Sub PageLoaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ' initialises on module selection
        If Not IsNothing(SelectedItem) Then
            SelectedItem.Click()
        End If
        If oSettings.DeveloperMode Then
            SubjectsExportHelp.Visibility = Visibility.Visible
        Else
            SubjectsExportHelp.Visibility = Visibility.Collapsed
        End If
    End Sub
    Private Sub SetIcons()
        ' initialises icon store
        m_Icons = New Dictionary(Of String, ImageSource)
        m_Icons.Add("CCMProperties", Converter.BitmapToBitmapSource(My.Resources.CCMProperties))
        m_Icons.Add("CCMBlock", Converter.BitmapToBitmapSource(My.Resources.CCMBlock))
        m_Icons.Add("CCMBody", Converter.BitmapToBitmapSource(My.Resources.CCMBody))
        m_Icons.Add("CCMField", Converter.BitmapToBitmapSource(My.Resources.CCMField))
        m_Icons.Add("CCMFooter", Converter.BitmapToBitmapSource(My.Resources.CCMFooter))
        m_Icons.Add("CCMHeader", Converter.BitmapToBitmapSource(My.Resources.CCMHeader))
        m_Icons.Add("CCMPageHeader", Converter.BitmapToBitmapSource(My.Resources.CCMPageHeader))
        m_Icons.Add("CCMFormHeader", Converter.BitmapToBitmapSource(My.Resources.CCMFormHeader))
        m_Icons.Add("CCMSection", Converter.BitmapToBitmapSource(My.Resources.CCMSection))
        m_Icons.Add("CCMSubSection", Converter.BitmapToBitmapSource(My.Resources.CCMSubSection))
        m_Icons.Add("CCMText", Converter.XamlToDrawingImage(My.Resources.CCMText))
        m_Icons.Add("CCMImage", Converter.XamlToDrawingImage(My.Resources.CCMImage))
        m_Icons.Add("CCMMCQ", Converter.BitmapToBitmapSource(My.Resources.CCMMCQ))
        m_Icons.Add("CCMPDF", Converter.BitmapToBitmapSource(My.Resources.CCMPDF))
        m_Icons.Add("CCMMCQData", Converter.BitmapToBitmapSource(My.Resources.CCMMCQData))
        m_Icons.Add("CCMSubjects", Converter.BitmapToBitmapSource(My.Resources.CCMSubjects))
        m_Icons.Add("CCMDivider", Converter.BitmapToBitmapSource(My.Resources.CCMDivider))
        m_Icons.Add("CCMGroup", Converter.BitmapToBitmapSource(My.Resources.CCMGroup))
        m_Icons.Add("CCMPageBreak", Converter.BitmapToBitmapSource(My.Resources.CCMPageBreak))
        m_Icons.Add("CCMPlus", Converter.XamlToDrawingImage(My.Resources.CCMPlus))
        m_Icons.Add("CCMMinus", Converter.XamlToDrawingImage(My.Resources.CCMMinus))
        m_Icons.Add("CCMUp", Converter.XamlToDrawingImage(My.Resources.CCMUp))
        m_Icons.Add("CCMDown", Converter.XamlToDrawingImage(My.Resources.CCMDown))
        m_Icons.Add("CCMBack", Converter.XamlToDrawingImage(My.Resources.CCMBack))
        m_Icons.Add("CCMForward", Converter.XamlToDrawingImage(My.Resources.CCMForward))
        m_Icons.Add("CCMBorder", Converter.XamlToDrawingImage(My.Resources.CCMBorder))
        m_Icons.Add("CCMBackground", Converter.XamlToDrawingImage(My.Resources.CCMBackground))
        m_Icons.Add("CCMNumbering", Converter.XamlToDrawingImage(My.Resources.CCMNumbering))
        m_Icons.Add("CCMSave", Converter.XamlToDrawingImage(My.Resources.CCMSave))
        m_Icons.Add("CCMLock", Converter.XamlToDrawingImage(My.Resources.CCMLock))
        m_Icons.Add("CCMVisible", Converter.XamlToDrawingImage(My.Resources.CCMVisible))
        m_Icons.Add("CCMNotVisible", Converter.XamlToDrawingImage(My.Resources.CCMNotVisible))
        m_Icons.Add("CCMSubject", Converter.XamlToDrawingImage(My.Resources.CCMSubject))
        m_Icons.Add("CCMTag", Converter.XamlToDrawingImage(My.Resources.CCMTag))
        m_Icons.Add("CCMClipboard", Converter.XamlToDrawingImage(My.Resources.CCMClipboard))
        m_Icons.Add("CCMSelection", Converter.XamlToDrawingImage(My.Resources.CCMSelection))
        m_Icons.Add("CCMJustificationLeft", Converter.XamlToDrawingImage(My.Resources.CCMJustificationLeft))
        m_Icons.Add("CCMJustificationCenter", Converter.XamlToDrawingImage(My.Resources.CCMJustificationCenter))
        m_Icons.Add("CCMJustificationRight", Converter.XamlToDrawingImage(My.Resources.CCMJustificationRight))
        m_Icons.Add("CCMJustificationJustify", Converter.XamlToDrawingImage(My.Resources.CCMJustificationJustify))
        m_Icons.Add("CCMFontBold", Converter.XamlToDrawingImage(My.Resources.CCMFontBold))
        m_Icons.Add("CCMFontItalic", Converter.XamlToDrawingImage(My.Resources.CCMFontItalic))
        m_Icons.Add("CCMFontUnderline", Converter.XamlToDrawingImage(My.Resources.CCMFontUnderline))
        m_Icons.Add("CCMAlignmentL", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentL))
        m_Icons.Add("CCMAlignmentT", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentT))
        m_Icons.Add("CCMAlignmentR", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentR))
        m_Icons.Add("CCMAlignmentB", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentB))
        m_Icons.Add("CCMAlignmentC", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentC))
        m_Icons.Add("CCMAlignmentUL", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentUL))
        m_Icons.Add("CCMAlignmentUR", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentUR))
        m_Icons.Add("CCMAlignmentLR", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentLR))
        m_Icons.Add("CCMAlignmentLL", Converter.XamlToDrawingImage(My.Resources.CCMAlignmentLL))
        m_Icons.Add("CCMStretchFill", Converter.XamlToDrawingImage(My.Resources.CCMStretchFill))
        m_Icons.Add("CCMStretchUniform", Converter.XamlToDrawingImage(My.Resources.CCMStretchUniform))
        m_Icons.Add("CCMImageBlue", Converter.XamlToDrawingImage(My.Resources.CCMImageBlue))
        m_Icons.Add("CCMCheck", Converter.XamlToDrawingImage(My.Resources.CCMCheck))
        m_Icons.Add("CCMHandwriting", Converter.XamlToDrawingImage(My.Resources.CCMHandwriting))
        m_Icons.Add("CCMPencil", Converter.XamlToDrawingImage(My.Resources.CCMPencil))
        m_Icons.Add("CCMBoxCheck", Converter.CombineDrawingImage(Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty), Converter.XamlToDrawingImage(My.Resources.CCMHandwriting), 0.6))
        m_Icons.Add("CCMCheckEmpty", Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty))
        m_Icons.Add("CCMCheckMark", Converter.XamlToDrawingImage(My.Resources.CCMCheckMark))
        m_Icons.Add("CCMCritical", Converter.XamlToDrawingImage(My.Resources.CCMCritical))
        m_Icons.Add("CCMOne", Converter.XamlToDrawingImage(My.Resources.CCMOne))
        m_Icons.Add("CCMCross", Converter.XamlToDrawingImage(My.Resources.CCMCross))
        m_Icons.Add("CCMMCQLoad", Converter.XamlToDrawingImage(My.Resources.CCMMCQLoad))
        m_Icons.Add("CCMPDF1", Converter.XamlToDrawingImage(My.Resources.CCMPDF1))
        m_Icons.Add("CCMReplacePDF", Converter.XamlToDrawingImage(My.Resources.CCMReplacePDF))
        m_Icons.Add("CCMLoremIpsum", Converter.XamlToDrawingImage(My.Resources.CCMLoremIpsum))
        m_Icons.Add("CCMTrashValid", Converter.XamlToDrawingImage(My.Resources.CCMTrashValid))
        m_Icons.Add("CCMImageSample", Converter.XamlToDrawingImage(My.Resources.CCMImageSample))

        ImageFormScrollUp.Source = m_Icons("CCMUp")
        ImageFormScrollDown.Source = m_Icons("CCMDown")
        PageHeaderAdd.HBSource = m_Icons("CCMPlus")
        PageHeaderRemove.HBSource = m_Icons("CCMMinus")
        PageHeaderPrevious.HBSource = m_Icons("CCMBack")
        PageHeaderNext.HBSource = m_Icons("CCMForward")
        PageHeaderFontSizeIncrease.HBSource = m_Icons("CCMPlus")
        PageHeaderFontSizeDecrease.HBSource = m_Icons("CCMMinus")
        BlockAddRow.HBSource = m_Icons("CCMPlus")
        BlockRemoveRow.HBSource = m_Icons("CCMMinus")
        BlockAddColumn.HBSource = m_Icons("CCMPlus")
        BlockRemoveColumn.HBSource = m_Icons("CCMMinus")
        BlockIncreaseWidth.HBSource = m_Icons("CCMUp")
        BlockDecreaseWidth.HBSource = m_Icons("CCMDown")
        StaticTextShowActive.HBSource = m_Icons("CCMVisible")
        StaticTextAddSubject.HBSource = m_Icons("CCMSubject")
        StaticTextAddTag.HBSource = m_Icons("CCMTag")
        StaticTextAdd.HBSource = m_Icons("CCMPlus")
        StaticTextRemove.HBSource = m_Icons("CCMMinus")
        StaticTextPrevious.HBSource = m_Icons("CCMBack")
        StaticTextNext.HBSource = m_Icons("CCMForward")
        BlockPlaceItem.HBSource = m_Icons("CCMSelection")
        JustificationLeft.HBSource = m_Icons("CCMJustificationLeft")
        JustificationCenter.HBSource = m_Icons("CCMJustificationCenter")
        JustificationRight.HBSource = m_Icons("CCMJustificationRight")
        JustificationJustify.HBSource = m_Icons("CCMJustificationJustify")
        FontBold.HBSource = m_Icons("CCMFontBold")
        FontItalic.HBSource = m_Icons("CCMFontItalic")
        FontUnderline.HBSource = m_Icons("CCMFontUnderline")
        FontSizeIncrease.HBSource = m_Icons("CCMPlus")
        FontSizeDecrease.HBSource = m_Icons("CCMMinus")
        FieldBorderIncreaseWidth.HBSource = m_Icons("CCMPlus")
        FieldBorderDecreaseWidth.HBSource = m_Icons("CCMMinus")
        FieldBackgroundDarken.HBSource = m_Icons("CCMPlus")
        FieldBackgroundLighten.HBSource = m_Icons("CCMMinus")
        FieldCritical.HBSource = m_Icons("CCMCritical")
        FieldSingle.HBSource = m_Icons("CCMOne")
        FieldExclude.HBSource = m_Icons("CCMCross")
        BlockIncreaseStart.HBSource = m_Icons("CCMPlus")
        BlockDecreaseStart.HBSource = m_Icons("CCMMinus")
        BlockPreviousType.HBSource = m_Icons("CCMBack")
        BlockNextType.HBSource = m_Icons("CCMForward")
        MCQLoadData.HBSource = m_Icons("CCMMCQLoad")
        TagsPrevious.HBSource = m_Icons("CCMBack")
        TagsNext.HBSource = m_Icons("CCMForward")
        TagsLoadData.HBSource = m_Icons("CCMTag")
        TagsAddRow.HBSource = m_Icons("CCMPlus")
        TagsRemoveRow.HBSource = m_Icons("CCMMinus")
        TagsAddColumn.HBSource = m_Icons("CCMPlus")
        TagsRemoveColumn.HBSource = m_Icons("CCMMinus")
        SubjectsLoadData.HBSource = m_Icons("CCMMCQLoad")
        SubjectsSavePDF.HBSource = m_Icons("CCMPDF1")
        SubjectsReplacePDF.HBSource = m_Icons("CCMReplacePDF")
        SubjectsExportHelp.HBSource = m_Icons("CCMSave")
        SectionTagGUID.HBSource = m_Icons("CCMLock")
        SectionTagCopy.HBSource = m_Icons("CCMClipboard")
        AlignmentL.HBSource = m_Icons("CCMAlignmentL")
        AlignmentT.HBSource = m_Icons("CCMAlignmentT")
        AlignmentR.HBSource = m_Icons("CCMAlignmentR")
        AlignmentB.HBSource = m_Icons("CCMAlignmentB")
        AlignmentC.HBSource = m_Icons("CCMAlignmentC")
        AlignmentUL.HBSource = m_Icons("CCMAlignmentUL")
        AlignmentUR.HBSource = m_Icons("CCMAlignmentUR")
        AlignmentLR.HBSource = m_Icons("CCMAlignmentLR")
        AlignmentLL.HBSource = m_Icons("CCMAlignmentLL")
        StretchFill.HBSource = m_Icons("CCMStretchFill")
        StretchUniform.HBSource = m_Icons("CCMStretchUniform")
        FieldBoxChoicePreviousLabel.HBSource = m_Icons("CCMBack")
        FieldBoxChoiceNextLabel.HBSource = m_Icons("CCMForward")
        FieldChoicePreviousTabletContent.HBSource = m_Icons("CCMBack")
        FieldChoiceNextTabletContent.HBSource = m_Icons("CCMForward")
        FieldChoiceAddTablet.HBSource = m_Icons("CCMPlus")
        FieldChoiceRemoveTablet.HBSource = m_Icons("CCMMinus")
        FieldChoiceIncreaseStart.HBSource = m_Icons("CCMPlus")
        FieldChoiceDecreaseStart.HBSource = m_Icons("CCMMinus")
        FieldChoiceIncreaseGroups.HBSource = m_Icons("CCMPlus")
        FieldChoiceDecreaseGroups.HBSource = m_Icons("CCMMinus")
        FieldChoicePreviousTopDescription.HBSource = m_Icons("CCMBack")
        FieldChoiceNextTopDescription.HBSource = m_Icons("CCMForward")
        FieldChoicePreviousBottomDescription.HBSource = m_Icons("CCMBack")
        FieldChoiceNextBottomDescription.HBSource = m_Icons("CCMForward")
        FieldHandwritingAddBlock.HBSource = m_Icons("CCMPlus")
        FieldHandwritingRemoveBlock.HBSource = m_Icons("CCMMinus")
        FieldHandwritingPreviousColumn.HBSource = m_Icons("CCMBack")
        FieldHandwritingNextColumn.HBSource = m_Icons("CCMForward")
        FieldHandwritingPreviousRow.HBSource = m_Icons("CCMBack")
        FieldHandwritingNextRow.HBSource = m_Icons("CCMForward")
        FieldHandwritingPreviousType.HBSource = m_Icons("CCMUp")
        FieldHandwritingNextType.HBSource = m_Icons("CCMDown")
        ImageLoad.HBSource = m_Icons("CCMImageBlue")
        ImageRecycleBin.Source = m_Icons("CCMTrashValid")
        ImageExport.Source = m_Icons("CCMSave")
    End Sub
    Private Sub SetImages(ByVal oGridMain As Controls.Grid)
        ' initiates image store
        If IsNothing(oImageStore) Then
            oImageStore = New ImageStore
        Else
            oImageStore.Clear()
        End If

        ' sets images by type
        Dim oImageFields As New List(Of Type) From {GetType(FieldNumbering), GetType(FieldBorder), GetType(FieldBackground), GetType(FieldText), GetType(FieldImage), GetType(FieldChoice), GetType(FieldChoiceVertical), GetType(FieldBoxChoice), GetType(FieldHandwriting), GetType(FieldFree)}
        For Each oType In oImageFields
            If oType.GetMember("FieldTypeImage").Count > 0 Then
                ' get method info for the continuing function
                Dim oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo = Nothing
                If oType.GetMember("GetFieldTypeImageProcess").Count > 0 Then
                    oMethodInfoGetFieldTypeImageProcess = oType.GetMethod("GetFieldTypeImageProcess", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                End If

                Dim oMethodInfoFieldTypeImage As Reflection.MethodInfo = oType.GetMethod("FieldTypeImage", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                Dim oFormItemFieldTypeImage As ImageSource = oMethodInfoFieldTypeImage.Invoke(Nothing, {oGridMain.ActualHeight, oMethodInfoGetFieldTypeImageProcess})
                If Not IsNothing(oFormItemFieldTypeImage) Then
                    oImageStore.AddImage(oType, oFormItemFieldTypeImage)
                End If
            End If
        Next
    End Sub
    Private Sub IterateDisplayDictionary(oFrameworkElement As FrameworkElement)
        ' iterates through the child elements and add to the display dictionary those who are of specific types
        Dim oChildren As IEnumerable = LogicalTreeHelper.GetChildren(oFrameworkElement)
        For Each oChild In oChildren
            Dim oChildFrameworkElement As FrameworkElement = TryCast(oChild, FrameworkElement)
            If Not IsNothing(oChildFrameworkElement) Then
                If (oChildFrameworkElement.Name <> String.Empty) AndAlso (Not DisplayDictionary.ContainsKey(oChildFrameworkElement.Name)) Then
                    If (oChildFrameworkElement.GetType.Equals(GetType(Controls.DockPanel)) OrElse oChildFrameworkElement.GetType.Equals(GetType(Controls.StackPanel)) OrElse oChildFrameworkElement.GetType.Equals(GetType(Controls.Grid)) OrElse oChildFrameworkElement.GetType.Equals(GetType(PDFViewer.PDFViewer)) OrElse oChildFrameworkElement.GetType.Equals(GetType(Common.HighlightTextBox))) Then
                        DisplayDictionary.Add(oChildFrameworkElement.Name, oChildFrameworkElement)
                    ElseIf DisplayAdd.Contains(oChildFrameworkElement.Name) Then
                        DisplayDictionary.Add(oChildFrameworkElement.Name, oChildFrameworkElement)
                    End If
                End If

                ' iterate through all subelements except for the pdf viewer
                If Not oChildFrameworkElement.GetType.Equals(GetType(PDFViewer.PDFViewer)) Then
                    IterateDisplayDictionary(oChildFrameworkElement)
                End If
            End If
        Next
    End Sub
    Private Shared Sub LeftArrangeFormItems()
        ' arrange the form items on the left panel
        Using oSuspender As New BaseFormItem.Suspender()
            Dim oGridFormContents As Controls.Grid = BaseFormItem.Root.GridFormContents
            oGridFormContents.Children.Clear()
            oGridFormContents.RowDefinitions.Clear()
            Dim oGridMain As Controls.Grid = BaseFormItem.Root.GridMain

            If Not IsNothing(BaseFormItem.FormMain) Then
                LeftIterateFormItem(BaseFormItem.FormMain, oGridFormContents, oGridMain, True)
            End If
        End Using
    End Sub
    Private Shared Sub LeftIterateFormItem(ByRef oFormItem As BaseFormItem, ByRef oGridFormContents As Controls.Grid, ByRef oGridMain As Controls.Grid, ByVal bExpanded As Boolean)
        ' runs through the children of the form item and display them
        For Each oChild In oFormItem.Children
            If Not oFormItem.IgnoreFields.Contains(oChild.Value) Then
                Dim oChildFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(oChild.Key)
                If bExpanded Then
                    oChildFormItem.TitleChanged()

                    Dim oRowDefinition As New Controls.RowDefinition
                    oRowDefinition.Height = GridLength.Auto
                    oGridFormContents.RowDefinitions.Add(oRowDefinition)

                    Dim oPropInfoName As Reflection.PropertyInfo = oChild.Value.GetProperty("Name", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoIconName As Reflection.PropertyInfo = oChild.Value.GetProperty("IconName", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoTagGuid As Reflection.PropertyInfo = oChild.Value.GetProperty("TagGuid", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim sFormItemName As String = oPropInfoName.GetValue(Nothing, Nothing)
                    Dim sFormItemIconName As String = oPropInfoIconName.GetValue(Nothing, Nothing)
                    Dim sFormItemTagGuid As String = oPropInfoTagGuid.GetValue(Nothing, Nothing)
                    Dim oFormItemIcon As ImageSource = If(m_Icons.ContainsKey(sFormItemIconName), m_Icons(sFormItemIconName), Nothing)
                    Dim sHeader As String = sFormItemName + If(oChildFormItem.Numbering = String.Empty, String.Empty, " " + oChildFormItem.Numbering)

                    Dim oFieldItemParam As New Tuple(Of String, Double, ImageSource, ImageSource)(sHeader, oChildFormItem.Multiplier, oFormItemIcon, Nothing)
                    Dim oFieldItem As New Common.FieldItem
                    oFieldItem.FIContent = oFieldItemParam

                    ' set binding for title
                    oFieldItem.DataContext = oChildFormItem
                    Dim oBinding As New Data.Binding
                    oBinding.Path = New PropertyPath("Title")
                    oBinding.Mode = Data.BindingMode.OneWay
                    oBinding.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                    oFieldItem.SetBinding(Common.FieldItem.FITitleProperty, oBinding)

                    oFieldItem.FIReference = New Tuple(Of Double, Double)(oGridMain.ActualWidth, oGridMain.ActualHeight)
                    Dim oBorderFormItem As New BorderFormItem(oChildFormItem.GUID, Not BaseFormItem.StaticFields.Contains(oChild.Value), oChild.Value, sFormItemTagGuid)
                    oChildFormItem.BorderFormItem = oBorderFormItem
                    With oBorderFormItem
                        .Child = oFieldItem
                        .BorderBrush = Brushes.Black
                        .BorderThickness = New Thickness(1)
                        .HorizontalAlignment = HorizontalAlignment.Stretch
                        .VerticalAlignment = VerticalAlignment.Stretch
                        .OpacityMask = Brushes.Black
                        .Margin = New Thickness(5, 0, 5, 0)
                        .CornerRadius = New CornerRadius(oGridMain.ActualHeight * CornerRadiusMultiplier)
                    End With

                    Controls.Grid.SetRow(oBorderFormItem, oGridFormContents.RowDefinitions.Count - 1)
                    Controls.Grid.SetColumn(oBorderFormItem, oChildFormItem.Level + 1)
                    Controls.Grid.SetColumnSpan(oBorderFormItem, oGridFormContents.ColumnDefinitions.Count - oChildFormItem.Level - 2)
                    Controls.Grid.SetZIndex(oBorderFormItem, 1)

                    oBorderFormItem.Visibility = Visibility.Visible
                    oGridFormContents.Children.Add(oBorderFormItem)
                    oChildFormItem.ChangeTransparent()

                    ' set expander
                    Dim iVisibleChildren As Integer = 0
                    For Each oType In oChildFormItem.Children.Values
                        If Not oChildFormItem.IgnoreFields.Contains(oType) Then
                            iVisibleChildren += 1
                            Exit For
                        End If
                    Next
                    If iVisibleChildren > 0 Then
                        oChildFormItem.ImageExpander = New ImageExpander(oGridFormContents, oChildFormItem)
                        Controls.Grid.SetRow(oChildFormItem.ImageExpander, oGridFormContents.RowDefinitions.Count - 1)
                        Controls.Grid.SetColumn(oChildFormItem.ImageExpander, oChildFormItem.Level)
                        Controls.Grid.SetZIndex(oChildFormItem.ImageExpander, 1)

                        oChildFormItem.ImageExpander.Visibility = Visibility.Visible
                        oGridFormContents.Children.Add(oChildFormItem.ImageExpander)
                    End If

                    ' set spacer
                    Dim oRowDefinitionSpacer As New Controls.RowDefinition
                    oRowDefinitionSpacer.Height = GridLength.Auto
                    oGridFormContents.RowDefinitions.Add(oRowDefinitionSpacer)

                    ' set border spacer
                    oChildFormItem.BorderSpacer = New BorderSpacer(oGridMain.ActualHeight, True, oBorderFormItem)
                    Controls.Grid.SetRow(oChildFormItem.BorderSpacer, oGridFormContents.RowDefinitions.Count - 1)
                    Controls.Grid.SetColumn(oChildFormItem.BorderSpacer, 0)
                    Controls.Grid.SetColumnSpan(oChildFormItem.BorderSpacer, oGridFormContents.ColumnDefinitions.Count)
                    Controls.Grid.SetZIndex(oChildFormItem.BorderSpacer, 1)

                    oChildFormItem.BorderSpacer.Visibility = Visibility.Visible
                    oGridFormContents.Children.Add(oChildFormItem.BorderSpacer)
                End If

                ' iterate through children
                LeftIterateFormItem(oChildFormItem, oGridFormContents, oGridMain, bExpanded And oChildFormItem.Expanded)
            End If
        Next
    End Sub
    Private Sub GridMainSizeChanged(sender As Object, e As SizeChangedEventArgs) Handles GridMain.SizeChanged
        If e.NewSize.Width > 0 And e.NewSize.Height > 0 Then
            SetImages(GridMain)
        End If

        LeftArrangeFormItems()
        If IsNothing(SelectedItem) Then
            BaseFormItem.FormMain.RightSetAllowedFields()
        Else
            SelectedItem.RightSetAllowedFields()
            SelectedItem.Display()
        End If
    End Sub
    Private Sub Page_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles Me.Loaded
        ' checks to perform when navigating to this page
        ' if a section is selected, then refresh the display to account for a display resolution change in the configuration
        If (Not IsNothing(SelectedItem)) AndAlso SelectedItem.GetType.Equals(GetType(FormSection)) Then
            Dim oSection As FormSection = SelectedItem
            oSection.Display()
        End If
    End Sub
#End Region
#Region "Classes"
    <DataContract(IsReference:=True)> Public Class FormMain
        ' top level form item
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public ChildDictionary As Dictionary(Of Guid, BaseFormItem)
            <DataMember> Public Tags As List(Of TagClass)

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    For Each oKeyGUID As Guid In ChildDictionary.Keys
                        .ChildDictionary.Add(oKeyGUID, ChildDictionary(oKeyGUID))
                        .Tags = New List(Of TagClass)
                        For Each oTag In Tags
                            .Tags.Add(oTag.Clone)
                        Next
                    Next
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Property Tags As List(Of TagClass)
            Get
                Return m_DataObject.Tags
            End Get
            Set(value As List(Of TagClass))
                m_DataObject.Tags = value
            End Set
        End Property
        Public Overrides Sub TitleChanged()
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .ChildDictionary = New Dictionary(Of Guid, BaseFormItem)
                .Tags = New List(Of TagClass)
            End With
            If GUID.Equals(Guid.Empty) Then
                GUID = Guid.NewGuid
            End If
            m_DataObject.ChildDictionary.Add(GUID, Me)
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "FormMain"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 0
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return String.Empty
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Return New List(Of String)
            End Get
        End Property
        Public Overrides Sub SetBindings()
        End Sub
        Public Overrides Sub Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return ""
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
        Public Function FindChild(ByVal oGUID As Guid) As BaseFormItem
            ' gets a form item from its guid
            If m_DataObject.ChildDictionary.ContainsKey(oGUID) Then
                Return m_DataObject.ChildDictionary(oGUID)
            Else
                Return Nothing
            End If
        End Function
        Public Function ContainsChild(ByVal oGUID As Guid) As Boolean
            ' checks if the dictionary contains this guid
            Return m_DataObject.ChildDictionary.ContainsKey(oGUID)
        End Function
        Public Sub AddChild(ByRef oFormItem As BaseFormItem)
            ' adds a child to the dictionary
            If oFormItem.GUID.Equals(Guid.Empty) Then
                oFormItem.GUID = Guid.NewGuid
            End If
            If Not m_DataObject.ChildDictionary.ContainsKey(oFormItem.GUID) Then
                m_DataObject.ChildDictionary.Add(oFormItem.GUID, oFormItem)
            End If
        End Sub
        Public Sub RemoveChild(ByVal oFormItem As BaseFormItem)
            ' removes a child from the dictionary
            If m_DataObject.ChildDictionary.ContainsKey(oFormItem.GUID) Then
                m_DataObject.ChildDictionary.Remove(oFormItem.GUID)
            End If
        End Sub
        Public Sub SetLevel()
            ' resets the level on each child item
            For Each oChildGUID In Children.Keys
                IterateSetLevel(oChildGUID)
            Next
        End Sub
        Private Sub IterateSetLevel(ByVal oGUID As Guid)
            ' resets the level on each child item
            Dim oFormItem As BaseFormItem = FindChild(oGUID)
            oFormItem.Level = oFormItem.Parent.Level + 1
            For Each oChildGUID In oFormItem.Children.Keys
                IterateSetLevel(oChildGUID)
            Next
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FormProperties
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass
        Private m_PageOrientation As ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)
        Private m_PageSize As ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)
        Private m_DataContext As FrameworkElement

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public FormTitle As String
            <DataMember> Public FormAuthor As String
            <DataMember> Public FormTopic As String
            <DataMember> Public SelectedOrientation As PageOrientation
            <DataMember> Public SelectedSize As PageSize

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .FormTitle = FormTitle
                    .FormAuthor = FormAuthor
                    .FormTopic = FormTopic
                    .SelectedOrientation = SelectedOrientation
                    .SelectedSize = SelectedSize
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Property FormTitle As String
            Get
                Return m_DataObject.FormTitle
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.FormTitle = value
                    OnPropertyChangedLocal("FormTitle")
                End Using
            End Set
        End Property
        Public Property FormAuthor As String
            Get
                Return m_DataObject.FormAuthor
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.FormAuthor = value
                    OnPropertyChangedLocal("FormAuthor")
                End Using
            End Set
        End Property
        Public Property FormTopic As String
            Get
                Return m_DataObject.FormTopic
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.FormTopic = value
                    OnPropertyChangedLocal("FormTopic")
                End Using
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "Title:  " + m_DataObject.FormTitle
            Title += vbCr + "Author: " + m_DataObject.FormAuthor
            Title += vbCr + "Page: " + [Enum].GetName(GetType(PageSize), SelectedSize) + " " + [Enum].GetName(GetType(PageOrientation), SelectedOrientation)
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .FormTitle = String.Empty
                .FormAuthor = String.Empty
                .FormTopic = String.Empty
                .SelectedOrientation = PDFHelper.DefaultPageOrientation
                .SelectedSize = PDFHelper.DefaultPageSize
            End With

            PDFHelper.ResetPage(SelectedOrientation, SelectedSize)

            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Properties"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.5
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMProperties"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelProperties")
                oDisplayFilter.Add("PropertiesTitle")
                oDisplayFilter.Add("PropertiesAuthor")
                oDisplayFilter.Add("PropertiesTopic")
                Return oDisplayFilter
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelProperties.DataContext = Me
            m_DataContext = Root.DockPanelProperties

            Root.PropertiesOrientation.DataContext = Me
            Root.PropertiesSize.DataContext = Me

            m_PageOrientation = New ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)
            For Each sOrientation In [Enum].GetNames(GetType(PageOrientation))
                m_PageOrientation.Add(New Common.HighlightComboBox.HCBDisplay(sOrientation, Guid.Empty, False))
            Next

            m_PageSize = New ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)
            For Each oPageSize In PDFHelper.PageDictionary(m_DataObject.SelectedOrientation)
                m_PageSize.Add(New Common.HighlightComboBox.HCBDisplay([Enum].GetName(GetType(PageSize), oPageSize), Guid.Empty, False))
            Next
            Update()

            PDFHelper.ResetPage(SelectedOrientation, SelectedSize)
            FormPageHeader.SetPageHeaderBackground()

            Dim oBindingProperties1 As New Data.Binding
            oBindingProperties1.Path = New PropertyPath("FormTitle")
            oBindingProperties1.Mode = Data.BindingMode.TwoWay
            oBindingProperties1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.PropertiesTitle.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingProperties1)

            Dim oBindingProperties2 As New Data.Binding
            oBindingProperties2.Path = New PropertyPath("FormAuthor")
            oBindingProperties2.Mode = Data.BindingMode.TwoWay
            oBindingProperties2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.PropertiesAuthor.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingProperties2)

            Dim oBindingProperties3 As New Data.Binding
            oBindingProperties3.Path = New PropertyPath("FormTopic")
            oBindingProperties3.Mode = Data.BindingMode.TwoWay
            oBindingProperties3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.PropertiesTopic.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingProperties3)
        End Sub
        Public Overrides Sub Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "2e14a0d2-12ec-4a4a-9733-2ec9f1db7341"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
        Public Property SelectedOrientation As PageOrientation
            Get
                Return m_DataObject.SelectedOrientation
            End Get
            Set(value As PageOrientation)
                If m_DataObject.SelectedOrientation <> value Then
                    m_DataObject.SelectedOrientation = value

                    ' reset page sizes
                    m_PageSize.Clear()
                    For Each oPageSize In PDFHelper.PageDictionary(m_DataObject.SelectedOrientation)
                        m_PageSize.Add(New Common.HighlightComboBox.HCBDisplay([Enum].GetName(GetType(PageSize), oPageSize), Guid.Empty, False))
                    Next

                    SelectedSize = PDFHelper.DefaultPageSize
                End If
            End Set
        End Property
        Public Property SelectedSize As PageSize
            Get
                Return m_DataObject.SelectedSize
            End Get
            Set(value As PageSize)
                Using oSuspender As New Suspender(Me, True)
                    If m_DataObject.SelectedSize <> value Then
                        m_DataObject.SelectedSize = value
                    End If

                    PDFHelper.ResetPage(SelectedOrientation, SelectedSize)
                    FormPageHeader.SetPageHeaderBackground()
                    FormFormHeader.BlockWidth = PDFHelper.PageBlockWidth
                    Dim oBlockList As List(Of FormBlock) = FormMain.GetFormItems(Of FormBlock)
                    For Each oBlock In oBlockList
                        oBlock.BlockWidth = oBlock.BlockWidth
                    Next
                    Dim oMCQList As List(Of FormMCQ) = FormMain.GetFormItems(Of FormMCQ)
                    For Each oMCQ In oMCQList
                        oMCQ.BlockWidth = oMCQ.BlockWidth
                    Next
                    Dim oSubSectionList As List(Of FormSubSection) = FormMain.GetFormItems(Of FormSubSection)
                    For Each oSubSection In oSubSectionList
                        oSubSection.BlockWidth = oSubSection.BlockWidth
                    Next

                    Dim oSectionList As List(Of FormSection) = FormMain.GetFormItems(Of FormSection)
                    For Each oSection In oSectionList
                        oSection.ImageTracker.Clear()
                    Next

                    Update()
                End Using
            End Set
        End Property
        Public Property SelectedOrientationDisplay As Common.HighlightComboBox.HCBDisplay
            Get
                Dim sOrientation As String = [Enum].GetName(GetType(PageOrientation), m_DataObject.SelectedOrientation)
                Return New Common.HighlightComboBox.HCBDisplay(sOrientation, Guid.NewGuid, False)
            End Get
            Set(value As Common.HighlightComboBox.HCBDisplay)
            End Set
        End Property
        Public Property SelectedSizeDisplay As Common.HighlightComboBox.HCBDisplay
            Get
                Dim sSize As String = [Enum].GetName(GetType(PageSize), m_DataObject.SelectedSize)
                Return New Common.HighlightComboBox.HCBDisplay(sSize, Guid.NewGuid, False)
            End Get
            Set(value As Common.HighlightComboBox.HCBDisplay)
            End Set
        End Property
        Public Property PageOrientation As ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)
            Get
                Return m_PageOrientation
            End Get
            Set(value As ObservableCollection(Of Common.HighlightComboBox.HCBDisplay))
            End Set
        End Property
        Public Property PageSize As ObservableCollection(Of Common.HighlightComboBox.HCBDisplay)
            Get
                Return m_PageSize
            End Get
            Set(value As ObservableCollection(Of Common.HighlightComboBox.HCBDisplay))
            End Set
        End Property
        Private Sub Update()
            OnPropertyChangedLocal("PageOrientation")
            OnPropertyChangedLocal("SelectedOrientationDisplay")
            OnPropertyChangedLocal("PageSize")
            OnPropertyChangedLocal("SelectedSizeDisplay")
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FormHeader
        Inherits BaseFormItem

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
        End Sub
        Public Overrides Sub Initialise()
            If ExtractFields(Of FormPageHeader)(GetType(FormPageHeader)).Count = 0 Then
                Dim oFormPageHeader As New FormPageHeader
                oFormPageHeader.Parent = Me
            End If
            If ExtractFields(Of FormFormHeader)(GetType(FormFormHeader)).Count = 0 Then
                Dim oFormFormHeader As New FormFormHeader
                oFormFormHeader.Parent = Me
            End If
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Header"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.5
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMHeader"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Return New List(Of String)
            End Get
        End Property
        Public Overrides Sub SetBindings()
        End Sub
        Public Overrides Sub Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "a3b9f44c-8ee0-4614-9e16-e7419b7b3f6d"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FormPageHeader
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public LineNumber As Integer
            <DataMember> Public Lines As New List(Of String)
            <DataMember> Public FontSizeMultiplier As Integer

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .LineNumber = LineNumber
                    .Lines.AddRange(Lines)
                    .FontSizeMultiplier = FontSizeMultiplier
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Title = "Line Count: " + Lines.Count.ToString
            Title += vbCr + "Font Size:   " + m_DataObject.FontSizeMultiplier.ToString
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .LineNumber = -1
                .FontSizeMultiplier = 1
            End With
            TitleChanged()
        End Sub
#End Region
#Region "Lines"
        Public Property LineNumber As Integer
            Get
                Return m_DataObject.LineNumber + 1
            End Get
            Set(value As Integer)
                If m_DataObject.Lines.Count = 0 Then
                    m_DataObject.LineNumber = -1
                ElseIf value - 1 >= 0 And value - 1 < m_DataObject.Lines.Count Then
                    m_DataObject.LineNumber = value - 1
                End If
                OnPropertyChangedLocal("LineNumber")
                OnPropertyChangedLocal("LineNumberText")
                OnPropertyChangedLocal("Line")
            End Set
        End Property
        Public Property LineNumberText As String
            Get
                Return (m_DataObject.LineNumber + 1).ToString
            End Get
            Set(value As String)
                LineNumber = CInt(Val(value))
            End Set
        End Property
        Public Property Line As String
            Get
                If m_DataObject.Lines.Count = 0 Then
                    Return String.Empty
                ElseIf m_DataObject.LineNumber >= 0 And m_DataObject.LineNumber < m_DataObject.Lines.Count Then
                    Return m_DataObject.Lines(m_DataObject.LineNumber)
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If m_DataObject.LineNumber >= 0 And m_DataObject.LineNumber < m_DataObject.Lines.Count Then
                    Using oSuspender As New Suspender(Me, True)
                        m_DataObject.Lines(m_DataObject.LineNumber) = value
                        OnPropertyChangedLocal("Line")
                    End Using
                End If
            End Set
        End Property
        Public ReadOnly Property Lines As List(Of String)
            Get
                Return m_DataObject.Lines
            End Get
        End Property
        Public Sub AddLine(ByVal sLine As String)
            If m_DataObject.Lines.Count = 0 Then
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Lines.Add(sLine)
                    m_DataObject.LineNumber = 0
                    OnPropertyChangedLocal("LineNumber")
                    OnPropertyChangedLocal("LineNumberText")
                    OnPropertyChangedLocal("Line")
                End Using
            ElseIf m_DataObject.LineNumber >= 0 And m_DataObject.LineNumber < m_DataObject.Lines.Count Then
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Lines.Insert(m_DataObject.LineNumber, sLine)
                    OnPropertyChangedLocal("Line")
                End Using
            End If
        End Sub
        Public Sub AddLine()
            If m_DataObject.Lines.Count = 0 Then
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Lines.Add(String.Empty)
                    m_DataObject.LineNumber = 0
                    OnPropertyChangedLocal("LineNumber")
                    OnPropertyChangedLocal("LineNumberText")
                    OnPropertyChangedLocal("Line")
                End Using
            ElseIf m_DataObject.LineNumber >= 0 And m_DataObject.LineNumber < m_DataObject.Lines.Count Then
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Lines.Insert(m_DataObject.LineNumber, String.Empty)
                    OnPropertyChangedLocal("Line")
                End Using
            End If
        End Sub
        Public Sub RemoveLine(ByVal iLineNumber As Integer)
            If iLineNumber >= 0 And iLineNumber < m_DataObject.Lines.Count Then
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Lines.RemoveAt(iLineNumber)
                    If m_DataObject.Lines.Count = 0 Then
                        m_DataObject.LineNumber = -1
                    Else
                        m_DataObject.LineNumber = Math.Min(m_DataObject.LineNumber, m_DataObject.Lines.Count - 1)
                    End If
                    OnPropertyChangedLocal("LineNumber")
                    OnPropertyChangedLocal("LineNumberText")
                    OnPropertyChangedLocal("Line")
                End Using
            End If
        End Sub
        Public Sub RemoveLine()
            If m_DataObject.LineNumber >= 0 And m_DataObject.LineNumber < m_DataObject.Lines.Count Then
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Lines.RemoveAt(m_DataObject.LineNumber)
                    If m_DataObject.Lines.Count = 0 Then
                        m_DataObject.LineNumber = -1
                    Else
                        m_DataObject.LineNumber = Math.Min(m_DataObject.LineNumber, m_DataObject.Lines.Count - 1)
                    End If
                    OnPropertyChangedLocal("LineNumber")
                    OnPropertyChangedLocal("LineNumberText")
                    OnPropertyChangedLocal("Line")
                End Using
            End If
        End Sub
        Public Sub PreviousLine()
            If m_DataObject.Lines.Count > 0 Then
                If m_DataObject.LineNumber < 1 Then
                    m_DataObject.LineNumber = 0
                Else
                    m_DataObject.LineNumber -= 1
                End If
                m_DataObject.LineNumber = Math.Min(m_DataObject.LineNumber, m_DataObject.Lines.Count - 1)
                OnPropertyChangedLocal("LineNumber")
                OnPropertyChangedLocal("LineNumberText")
                OnPropertyChangedLocal("Line")
            End If
        End Sub
        Public Sub NextLine()
            If m_DataObject.Lines.Count > 0 Then
                If m_DataObject.LineNumber > m_DataObject.Lines.Count - 2 Then
                    m_DataObject.LineNumber = m_DataObject.Lines.Count - 1
                Else
                    m_DataObject.LineNumber += 1
                End If
                m_DataObject.LineNumber = Math.Max(m_DataObject.LineNumber, 0)
                OnPropertyChangedLocal("LineNumber")
                OnPropertyChangedLocal("LineNumberText")
                OnPropertyChangedLocal("Line")
            End If
        End Sub
#End Region
#Region "Font"
        Private Const MaxFontSizeMultiplier As Integer = 3

        Public Property FontSizeMultiplier As Integer
            Get
                Return m_DataObject.FontSizeMultiplier
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 1 Then
                        m_DataObject.FontSizeMultiplier = 1
                    ElseIf value > MaxFontSizeMultiplier Then
                        m_DataObject.FontSizeMultiplier = MaxFontSizeMultiplier
                    Else
                        m_DataObject.FontSizeMultiplier = value
                    End If
                    OnPropertyChangedLocal("FontSizeMultiplier")
                    OnPropertyChangedLocal("FontSizeMultiplierText")
                End Using
            End Set
        End Property
        Public Property FontSizeMultiplierText As String
            Get
                Return m_DataObject.FontSizeMultiplier.ToString
            End Get
            Set(value As String)
                FontSizeMultiplier = CInt(Val(value))
            End Set
        End Property
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Page Header"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.25
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMPageHeader"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelPageHeader")
                oDisplayFilter.Add("PageHeaderCurrent")
                oDisplayFilter.Add("PageHeaderFontSizeMultiplier")
                oDisplayFilter.Add("PageHeaderContent")
                Return oDisplayFilter
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelPageHeader.DataContext = Me

            Dim oBindingPageHeader1 As New Data.Binding
            oBindingPageHeader1.Path = New PropertyPath("LineNumberText")
            oBindingPageHeader1.Mode = Data.BindingMode.TwoWay
            oBindingPageHeader1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.PageHeaderCurrent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingPageHeader1)

            Dim oBindingPageHeader2 As New Data.Binding
            oBindingPageHeader2.Path = New PropertyPath("Line")
            oBindingPageHeader2.Mode = Data.BindingMode.TwoWay
            oBindingPageHeader2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.PageHeaderContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingPageHeader2)

            Dim oBindingPageHeader3 As New Data.Binding
            oBindingPageHeader3.Path = New PropertyPath("FontSizeMultiplierText")
            oBindingPageHeader3.Mode = Data.BindingMode.TwoWay
            oBindingPageHeader3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.PageHeaderFontSizeMultiplier.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingPageHeader3)
        End Sub
        Public Overrides Sub Display()
            Dim oHeaderText As ImageSource = Converter.BitmapToBitmapSource(GetPageHeaderText(Lines, FontSizeMultiplier, XStringAlignment.Center))
            Dim oImagePageHeader As Controls.Image = Root.ImagePageHeader
            Dim oImagePageHeaderText As Controls.Image = Root.ImagePageHeaderText
            If Not IsNothing(oHeaderText) Then
                oImagePageHeaderText.Width = oImagePageHeader.ActualWidth * oHeaderText.Width / oImagePageHeader.Source.Width
                oImagePageHeaderText.Height = oImagePageHeader.ActualHeight * oHeaderText.Height / oImagePageHeader.Source.Height
            End If
            oImagePageHeaderText.Source = oHeaderText
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "709edbd7-b0d9-47a1-94a9-7af43d20601f"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
        Public Sub SetPageHeaderBackground()
            ' sets page header background
            Dim oImagePageHeader As Controls.Image = Root.ImagePageHeader
            oImagePageHeader.Source = GetPageHeaderBackground()
            FormPDF.Changed = True
        End Sub
        Private Shared Function GetPageHeaderBackground() As Imaging.BitmapSource
            ' creates a 4 cm height bitmap to contain the drawn page header background, with margins of 5 cm to either side
            Dim oReturnBitmapSource As Imaging.BitmapSource = Nothing
            Const BitmapWidthMargins As Double = 5
            Const BitmapHeight As Double = 4
            Dim BitmapWidth As Double = PDFHelper.PageFullWidth.Centimeter - (BitmapWidthMargins * 2)
            Dim oBackgroundSize As New XSize(XUnit.FromCentimeter(BitmapWidth).Point, XUnit.FromCentimeter(BitmapHeight).Point)

            Dim iBitmapWidth As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Width).Inch * RenderResolution300)
            Dim iBitmapHeight As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Height).Inch * RenderResolution300)
            Using oBitmap As New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                oBitmap.SetResolution(RenderResolution300, RenderResolution300)

                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, oBackgroundSize, XGraphicsUnit.Point)
                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                        oXGraphics.DrawRectangle(XBrushes.White, New XRect(oBackgroundSize))
                        Dim oCenterPoint As New XPoint(oXGraphics.PageSize.Width / 2, oXGraphics.PageSize.Height / 2)

                        Dim oCornerstone As XImage = XImage.FromGdiPlusImage(PDFHelper.GetCornerstoneImage(SpacingSmall, RenderResolution300, RenderResolution300, False))
                        Dim width As Double = XUnit.FromInch(oCornerstone.PixelWidth / RenderResolution300).Point
                        Dim height As Double = XUnit.FromInch(oCornerstone.PixelHeight / RenderResolution300).Point
                        Dim fSpacing As Double = PDFHelper.Cornerstone(1).Item3.X - PDFHelper.Cornerstone(0).Item3.X
                        Dim oLeftCornerstoneCenterPoint As New XPoint(oCenterPoint.X - fSpacing / 2, oCenterPoint.Y)
                        Dim oRightCornerstoneCenterPoint As New XPoint(oCenterPoint.X + fSpacing / 2, oCenterPoint.Y)
                        oXGraphics.DrawImage(oCornerstone, oLeftCornerstoneCenterPoint.X - width / 2, oLeftCornerstoneCenterPoint.Y - height / 2, width, height)
                        oXGraphics.DrawImage(oCornerstone, oRightCornerstoneCenterPoint.X - width / 2, oLeftCornerstoneCenterPoint.Y - height / 2, width, height)
                    End Using
                End Using

                oReturnBitmapSource = Converter.BitmapToBitmapSource(oBitmap)
            End Using

            Return oReturnBitmapSource
        End Function
        Public Shared Sub DrawPageHeader(ByRef oXGraphics As XGraphics, ByVal oCenterPoint As XPoint, ByVal oHeaderTextList As List(Of String), ByVal iFontSizeMultiplier As Integer, ByVal oAlignment As XStringAlignment, ByVal oBrushColour As XBrush)
            ' draws the page header text on the supplied XGraphics at the specified centerpoint
            ' this is scaled to fit into 1.5 cm height box, with margins of 6 cm on either side
            Const fFontSize As Double = 5
            Const BoxWidthMargins As Double = 6
            Const BoxHeight As Double = 1.5
            Dim BoxWidth As Double = PDFHelper.PageFullWidth.Centimeter - (BoxWidthMargins * 2)
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oStringFormat As New XStringFormat
            oStringFormat.Alignment = oAlignment
            oStringFormat.LineAlignment = XLineAlignment.Center
            Dim oBoxSize As New XSize(XUnit.FromCentimeter(BoxWidth).Point, XUnit.FromCentimeter(BoxHeight).Point)
            Dim fActualMultiplier As Double = (((iFontSizeMultiplier - 1) / 2) + 1)
            Dim oMeasuredSize As Tuple(Of Double, Double, Double) = GetTextSize(oXGraphics, oHeaderTextList, fFontSize * fActualMultiplier, oBoxSize, Enumerations.ScaleDirection.Both, Enumerations.ScaleType.Down)
            Dim oArielFont As New XFont(FontArial, fFontSize * fActualMultiplier * oMeasuredSize.Item1, XFontStyle.Regular, oFontOptions)

            Dim fCurrentHeightPosition As Double = 0
            For Each sHeaderText As String In oHeaderTextList
                Dim oSize As XSize = oXGraphics.MeasureString(Trim(sHeaderText), oArielFont)
                Dim oXRect As New XRect(oCenterPoint.X - oMeasuredSize.Item2 / 2, oCenterPoint.Y - oMeasuredSize.Item3 / 2 + fCurrentHeightPosition, oMeasuredSize.Item2, oSize.Height)
                fCurrentHeightPosition += oArielFont.GetHeight
                oXGraphics.DrawString(Trim(sHeaderText), oArielFont, oBrushColour, oXRect, oStringFormat)
            Next
        End Sub
        Public Shared Function GetPageHeaderText(ByVal oHeaderTextList As List(Of String), ByVal iFontSizeMultiplier As Integer, ByVal oAlignment As XStringAlignment) As System.Drawing.Bitmap
            ' draws just the text of the bitmap
            Dim oTextBlockSize As XSize = MeasurePageHeader(oHeaderTextList)

            If oTextBlockSize.Width = 0 Or oTextBlockSize.Height = 0 Then
                Return Nothing
            Else
                Dim fScaleFactor As Double = RenderResolution300 / 72
                Dim oBitmap As New System.Drawing.Bitmap(CInt(Math.Ceiling(oTextBlockSize.Width * fScaleFactor)), CInt(Math.Ceiling(oTextBlockSize.Height * fScaleFactor)))
                oBitmap.SetResolution(RenderResolution300, RenderResolution300)

                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, oTextBlockSize, XGraphicsUnit.Point)
                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                        Dim oCenterPoint As New XPoint(oXGraphics.PageSize.Width / 2, oXGraphics.PageSize.Height / 2)
                        DrawPageHeader(oXGraphics, oCenterPoint, oHeaderTextList, iFontSizeMultiplier, oAlignment, XBrushes.Black)
                    End Using
                End Using

                Return oBitmap
            End If
        End Function
        Private Shared Function MeasurePageHeader(ByVal oHeaderTextList As List(Of String)) As XSize
            ' measures the page header text size
            Const fFontSize As Double = 10
            Const BoxWidth As Double = 9
            Const BoxHeight As Double = 1.5
            Dim oBoxSize As New XSize(XUnit.FromCentimeter(BoxWidth).Point, XUnit.FromCentimeter(BoxHeight).Point)

            Dim oMeasuredSize As Tuple(Of Double, Double, Double) = Nothing
            Using oXGraphics As XGraphics = XGraphics.CreateMeasureContext(oBoxSize, XGraphicsUnit.Point, XPageDirection.Downwards)
                oMeasuredSize = GetTextSize(oXGraphics, oHeaderTextList, fFontSize, oBoxSize, Enumerations.ScaleDirection.Both, Enumerations.ScaleType.Down)
            End Using

            Return New XSize(oMeasuredSize.Item2, oMeasuredSize.Item3)
        End Function
        Public Shared Function GetTextSize(ByRef oXGraphics As XGraphics, ByVal oTextList As List(Of String), ByVal fFontSize As Double, ByVal oBoxSize As XSize, ByVal oScaleDirection As Enumerations.ScaleDirection, ByVal oScaleType As Enumerations.ScaleType) As Tuple(Of Double, Double, Double)
            ' this function gets the measured size for the drawing function
            ' scaledirection scales the text list to fit within the supplied box size, either in both, vertical, or horizontal directions
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oArielFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)

            ' gets text size
            Dim fMeasuredWidth As Double = 0
            Dim fMeasuredHeight As Double = 0
            Dim fScaleFactor As Double = 0
            Dim oArielFontScaled As XFont = Nothing
            For Each sText As String In oTextList
                Dim oSize As XSize = oXGraphics.MeasureString(Trim(sText), oArielFont)
                fMeasuredWidth = Math.Max(fMeasuredWidth, oSize.Width)
                fMeasuredHeight += oArielFont.GetHeight
            Next

            ' remeasure after scaling
            Select Case oScaleType
                Case Enumerations.ScaleType.Both
                    Select Case oScaleDirection
                        Case Enumerations.ScaleDirection.Both
                            fScaleFactor = Math.Min(oBoxSize.Width / fMeasuredWidth, oBoxSize.Height / fMeasuredHeight)
                        Case Enumerations.ScaleDirection.Vertical
                            fScaleFactor = oBoxSize.Height / fMeasuredHeight
                        Case Enumerations.ScaleDirection.Horizontal
                            fScaleFactor = oBoxSize.Width / fMeasuredWidth
                    End Select
                Case Enumerations.ScaleType.Down
                    Select Case oScaleDirection
                        Case Enumerations.ScaleDirection.Both
                            fScaleFactor = Math.Min(Math.Min(oBoxSize.Width / fMeasuredWidth, 1), Math.Min(oBoxSize.Height / fMeasuredHeight, 1))
                        Case Enumerations.ScaleDirection.Vertical
                            fScaleFactor = Math.Min(oBoxSize.Height / fMeasuredHeight, 1)
                        Case Enumerations.ScaleDirection.Horizontal
                            fScaleFactor = Math.Min(oBoxSize.Width / fMeasuredWidth, 1)
                    End Select
                Case Enumerations.ScaleType.Up
                    Select Case oScaleDirection
                        Case Enumerations.ScaleDirection.Both
                            fScaleFactor = Math.Min(Math.Max(oBoxSize.Width / fMeasuredWidth, 1), Math.Max(oBoxSize.Height / fMeasuredHeight, 1))
                        Case Enumerations.ScaleDirection.Vertical
                            fScaleFactor = Math.Max(oBoxSize.Height / fMeasuredHeight, 1)
                        Case Enumerations.ScaleDirection.Horizontal
                            fScaleFactor = Math.Max(oBoxSize.Width / fMeasuredWidth, 1)
                    End Select
            End Select

            oArielFontScaled = New XFont(FontArial, fFontSize * fScaleFactor, XFontStyle.Regular, oFontOptions)

            fMeasuredWidth = 0
            fMeasuredHeight = 0
            For Each sText As String In oTextList
                Dim oSize As XSize = oXGraphics.MeasureString(Trim(sText), oArielFontScaled)
                fMeasuredWidth = Math.Max(fMeasuredWidth, oSize.Width)
                fMeasuredHeight += oArielFontScaled.GetHeight
            Next

            Return New Tuple(Of Double, Double, Double)(fScaleFactor, fMeasuredWidth, fMeasuredHeight)
        End Function
    End Class
    <DataContract(IsReference:=True)> Public Class FormFormHeader
        Inherits FormBlock

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Title = vbCr + "Rows:   " + GridHeight.ToString
            Title += vbCr + "Columns: " + GridWidth.ToString
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Form Header"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.25
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMFormHeader"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Return MyBase.DisplayFilter
            End Get
        End Property
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "4a6a89db-04bb-482e-8a46-d964c16be57f"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FieldBorder), GetType(FieldBackground), GetType(FieldText), GetType(FieldImage)}
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FormBlock
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass
        Protected m_ShowGrid As Boolean
        Protected m_ShowActive As Boolean
        Public Shared FormHeaderRectangleOverlay As Shapes.Rectangle(,)
        Public Shared WithEvents CanvasBlockContent As Controls.Canvas
        Private Shared SelectionActive As New Tuple(Of Boolean, Enumerations.InputDevice)(False, Enumerations.InputDevice.NoDevice)
        Private Shared SelectionStart As Point = Nothing
        Private Shared RectangleFormHeaderSelection As Shapes.Rectangle

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public GridWidth As Integer
            <DataMember> Public GridHeight As Integer
            <DataMember> Public BlockWidth As Integer
            <DataMember> Public TabletSingleChoiceOnly As Boolean

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .GridWidth = GridWidth
                    .GridHeight = GridHeight
                    .BlockWidth = BlockWidth
                    .TabletSingleChoiceOnly = TabletSingleChoiceOnly
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            CType(oFormItem, FormBlock).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Private Const MaxGridHeight As Integer = 48
        Public Property GridWidth As Integer
            Get
                Return m_DataObject.GridWidth
            End Get
            Set(value As Integer)
                Dim iOldGridWidth As Integer = m_DataObject.GridWidth

                If value < 1 Then
                    m_DataObject.GridWidth = 1
                ElseIf value > MaxGridWidth() Then
                    m_DataObject.GridWidth = MaxGridWidth()
                Else
                    m_DataObject.GridWidth = value
                End If

                If iOldGridWidth <> m_DataObject.GridWidth Then
                    Using oSuspender As New Suspender(Me, True)
                        ResetFields(New Int32Rect(0, 0, m_DataObject.GridWidth, m_DataObject.GridHeight))
                        OnPropertyChangedLocal("GridWidth")
                        OnPropertyChangedLocal("GridWidthText")
                    End Using
                End If
            End Set
        End Property
        Public Property GridWidthText As String
            Get
                Return m_DataObject.GridWidth.ToString
            End Get
            Set(value As String)
                GridWidth = CInt(Val(value))
            End Set
        End Property
        Private Function MaxGridWidth() As Integer
            Return 8 * EffectiveBlockWidth
        End Function
        Public Property GridHeight As Integer
            Get
                Return m_DataObject.GridHeight
            End Get
            Set(value As Integer)
                Dim iOldGridHeight As Integer = m_DataObject.GridHeight

                If value < 1 Then
                    m_DataObject.GridHeight = 1
                ElseIf value > MaxGridHeight Then
                    m_DataObject.GridHeight = MaxGridHeight
                Else
                    m_DataObject.GridHeight = value
                End If

                If iOldGridHeight <> m_DataObject.GridHeight Then
                    Using oSuspender As New Suspender(Me, True)
                        ResetFields(New Int32Rect(0, 0, m_DataObject.GridWidth, m_DataObject.GridHeight))
                        OnPropertyChangedLocal("GridHeight")
                        OnPropertyChangedLocal("GridHeightText")
                    End Using
                End If
            End Set
        End Property
        Public Property GridHeightText As String
            Get
                Return m_DataObject.GridHeight.ToString
            End Get
            Set(value As String)
                GridHeight = CInt(Val(value))
            End Set
        End Property
        Public Overridable Property BlockWidth As Integer
            Get
                Dim oSubSection As FormSubSection = TryCast(Parent, FormSubSection)
                If IsNothing(oSubSection) Then
                    Return m_DataObject.BlockWidth
                Else
                    Return oSubSection.BlockWidth
                End If
            End Get
            Set(value As Integer)
                Dim oSubSection As FormSubSection = TryCast(Parent, FormSubSection)
                If IsNothing(oSubSection) Then
                    Dim iOldBlockWidth As Integer = m_DataObject.BlockWidth

                    If Me.GetType.Equals(GetType(FormFormHeader)) Then
                        m_DataObject.BlockWidth = PDFHelper.PageBlockWidth
                    Else
                        If value < 0 Then
                            m_DataObject.BlockWidth = 0
                        ElseIf value > PDFHelper.PageBlockWidth Then
                            m_DataObject.BlockWidth = PDFHelper.PageBlockWidth
                        Else
                            m_DataObject.BlockWidth = value
                        End If
                    End If

                    If iOldBlockWidth <> m_DataObject.BlockWidth Then
                        ' reset grid width
                        Using oSuspender As New Suspender(Me, True)
                            Dim iOldGridWidth As Integer = m_DataObject.GridWidth
                            If m_DataObject.GridWidth > MaxGridWidth() Then
                                m_DataObject.GridWidth = MaxGridWidth()
                            End If

                            ResetFields(Nothing)
                            OnPropertyChangedLocal("GridWidth")
                            OnPropertyChangedLocal("GridWidthText")
                            OnPropertyChangedLocal("BlockWidth")
                            OnPropertyChangedLocal("BlockWidthText")
                        End Using
                    End If
                End If
            End Set
        End Property
        Public ReadOnly Property EffectiveBlockWidth As Integer
            Get
                Return If(BlockWidth = 0, PDFHelper.PageBlockWidth, BlockWidth)
            End Get
        End Property
        Public Overridable Property BlockWidthText As String
            Get
                If BlockWidth = 0 Then
                    Return "Max"
                Else
                    Return BlockWidth.ToString
                End If
            End Get
            Set(value As String)
                BlockWidth = CInt(Val(value))
            End Set
        End Property
        Public Property ShowGrid As Boolean
            Get
                Return m_ShowGrid
            End Get
            Set(value As Boolean)
                Using oSuspender As New Suspender(Me, True)
                    m_ShowGrid = value
                    OnPropertyChangedLocal("ShowGrid")
                End Using
            End Set
        End Property
        Public Property ShowActive As Boolean
            Get
                Return m_ShowActive
            End Get
            Set(value As Boolean)
                Using oSuspender As New Suspender(Me, True)
                    m_ShowActive = value
                    OnPropertyChangedLocal("ShowActive")
                End Using
            End Set
        End Property
        Public Function LeftIndent() As XUnit
            ' returns left indent width depending on the numbering
            ' numbering, remove 3/20 of width of a single block
            Dim iNumbering As Integer = 0
            Dim oSubSection As FormSubSection = TryCast(Parent, FormSubSection)
            If IsNothing(oSubSection) Then
                iNumbering = Math.Min((Aggregate oType As Type In Children.Values Into Count(oType.Equals(GetType(FieldNumbering)))), 1)
            Else
                iNumbering = Math.Min((Aggregate oType As Type In oSubSection.Children.Values Into Count(oType.Equals(GetType(FieldNumbering)))), 1)
            End If

            Dim fTwentiethWidth As Double = PDFHelper.BlockWidth.Point / 20
            Return New XUnit(3 * fTwentiethWidth * iNumbering)
        End Function
        Public Property TabletSingleChoiceOnly As Boolean
            Get
                Return m_DataObject.TabletSingleChoiceOnly
            End Get
            Set(value As Boolean)
                m_DataObject.TabletSingleChoiceOnly = value
                OnPropertyChangedLocal("TabletSingleChoiceOnly")
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = vbCr + "Blocks:   " + BlockWidthText
            Title += vbCr + "Rows:   " + m_DataObject.GridHeight.ToString
            Title += vbCr + "Columns: " + m_DataObject.GridWidth.ToString
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .GridWidth = 1
                .GridHeight = 1
                .BlockWidth = PDFHelper.PageBlockWidth
                .TabletSingleChoiceOnly = True
            End With
            m_ShowGrid = True
            m_ShowActive = True
            FormHeaderRectangleOverlay = Nothing
            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Block"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.0
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMBlock"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelBlock")
                oDisplayFilter.Add("StackPanelBlockRow")
                oDisplayFilter.Add("StackPanelBlockColumn")
                oDisplayFilter.Add("StackPanelShowGrid")
                oDisplayFilter.Add("BlockRow")
                oDisplayFilter.Add("BlockColumn")
                oDisplayFilter.Add("RectangleBlockBackground")
                oDisplayFilter.Add("ImageBlockContent")
                oDisplayFilter.Add("CanvasBlockContent")
                oDisplayFilter.Add("ScrollViewerBlockContent")
                If (Me.GetType.Equals(GetType(FormBlock))) AndAlso (Not IsNothing(Parent)) AndAlso (Not Parent.GetType.Equals(GetType(FormSubSection))) Then
                    oDisplayFilter.Add("StackPanelBlockWidth")
                    oDisplayFilter.Add("BlockWidth")
                End If
                Return oDisplayFilter
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelBlock.DataContext = Me
            Root.StackPanelBlockColumn.DataContext = Me
            Root.StackPanelBlockRow.DataContext = Me
            Root.StackPanelBlockWidth.DataContext = Me

            Dim oBindingBlock1 As New Data.Binding
            oBindingBlock1.Path = New PropertyPath("GridWidthText")
            oBindingBlock1.Mode = Data.BindingMode.TwoWay
            oBindingBlock1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockColumn.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBlock1)

            Dim oBindingBlock2 As New Data.Binding
            oBindingBlock2.Path = New PropertyPath("GridHeightText")
            oBindingBlock2.Mode = Data.BindingMode.TwoWay
            oBindingBlock2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockRow.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBlock2)

            Dim oBindingBlock3 As New Data.Binding
            oBindingBlock3.Path = New PropertyPath("BlockWidthText")
            oBindingBlock3.Mode = Data.BindingMode.TwoWay
            oBindingBlock3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockWidth.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBlock3)

            Dim oBindingBlock4 As New Data.Binding
            oBindingBlock4.Path = New PropertyPath("ShowGrid")
            oBindingBlock4.Mode = Data.BindingMode.TwoWay
            oBindingBlock4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockShowGrid.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingBlock4)
        End Sub
        Public Overrides Sub Display()
            Dim oScrollViewerBlockContent As Controls.ScrollViewer = Root.ScrollViewerBlockContent
            Dim oScrollBar As Controls.Primitives.ScrollBar = CType(oScrollViewerBlockContent.Template.FindName("PART_VerticalScrollBar", oScrollViewerBlockContent), Controls.Primitives.ScrollBar)

            If Not IsNothing(oScrollBar) Then
                oScrollViewerBlockContent.UpdateLayout()

                ' block for display only scales horizontally
                Dim oScaleDirection As Enumerations.ScaleDirection = GetScaleDirection()
                Dim oBitmapSource As Imaging.BitmapSource = GetBlockBitmapSource.Item1
                Dim fLeftIndentPixel As Double = LeftIndent.Point * RenderResolution300 / 72
                Dim fBlockSpacing As Double = PDFHelper.BlockSpacer.Inch * RenderResolution300
                Dim fScaleFactor As Double = Math.Min((oScrollViewerBlockContent.ActualWidth - oScrollBar.ActualWidth) / oBitmapSource.PixelWidth, oScrollViewerBlockContent.ActualHeight / oBitmapSource.PixelHeight)
                Dim fImageWidth As Double = oBitmapSource.PixelWidth * fScaleFactor
                Dim fImageHeight As Double = oBitmapSource.PixelHeight * fScaleFactor

                ' set image
                With Root.RectangleBlockBackground
                    .Width = fImageWidth
                    .Height = fImageHeight
                End With

                With Root.ImageBlockContent
                    .Width = fImageWidth
                    .Height = fImageHeight
                    .Source = oBitmapSource
                    .Tag = fScaleFactor
                End With

                ' set overlay canvas
                With Root.CanvasBlockContent
                    .Width = fImageWidth
                    .Height = fImageHeight
                    .Margin = New Thickness(0, 0, oScrollBar.ActualWidth, 0)
                    .Children.Clear()
                End With

                If Root.CanvasBlockContent.Visibility = Visibility.Visible Then
                    ' set overlay rectangles
                    Dim oVisibility As Visibility = If(m_ShowGrid, Visibility.Visible, Visibility.Hidden)
                    ReDim FormHeaderRectangleOverlay(GridWidth - 1, GridHeight - 1)
                    For x = 0 To GridWidth - 1
                        For y = 0 To GridHeight - 1
                            Dim oRectangle As New Shapes.Rectangle
                            With oRectangle
                                If (oScaleDirection And Enumerations.ScaleDirection.Horizontal) = Enumerations.ScaleDirection.Horizontal Then
                                    .Width = (oBitmapSource.PixelWidth - (2 * fBlockSpacing) - fLeftIndentPixel) * fScaleFactor / GridWidth
                                    Controls.Canvas.SetLeft(oRectangle, ((fBlockSpacing + fLeftIndentPixel) * fScaleFactor) + x * .Width)
                                Else
                                    .Width = (oBitmapSource.PixelWidth - fLeftIndentPixel) * fScaleFactor / GridWidth
                                    Controls.Canvas.SetLeft(oRectangle, (fLeftIndentPixel * fScaleFactor) + x * .Width)
                                End If
                                If (oScaleDirection And Enumerations.ScaleDirection.Vertical) = Enumerations.ScaleDirection.Vertical Then
                                    .Height = (oBitmapSource.PixelHeight - (2 * fBlockSpacing)) * fScaleFactor / GridHeight
                                    Controls.Canvas.SetTop(oRectangle, (fBlockSpacing * fScaleFactor) + y * .Height)
                                Else
                                    .Height = oBitmapSource.PixelHeight * fScaleFactor / GridHeight
                                    Controls.Canvas.SetTop(oRectangle, y * .Height)
                                End If
                                .Stroke = Brushes.CornflowerBlue
                                .StrokeThickness = 1
                                .StrokeDashArray = New DoubleCollection From {4}
                                .Fill = Brushes.Transparent
                                .Visibility = oVisibility
                            End With
                            Controls.Canvas.SetZIndex(oRectangle, 1)
                            Root.CanvasBlockContent.Children.Add(oRectangle)
                            FormHeaderRectangleOverlay(x, y) = oRectangle
                        Next
                    Next
                End If
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "25ea913b-4509-440f-aa5f-66ca4f4040e4"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Dim oAllowedFields As New List(Of Type)

                ' leave out numbering when in a subsection
                If IsNothing(Parent) OrElse (Not Parent.GetType.Equals(GetType(FormSubSection))) Then
                    oAllowedFields.Add(GetType(FieldNumbering))
                End If

                oAllowedFields.AddRange({GetType(FieldBorder), GetType(FieldBackground), GetType(FieldText), GetType(FieldImage), GetType(FieldChoice), GetType(FieldChoiceVertical), GetType(FieldBoxChoice), GetType(FieldHandwriting), GetType(FieldFree)})
                oAllowedFields.AddRange(FormattingFields)
                Return oAllowedFields
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedSingleOnly As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FieldNumbering)}
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Dim XBlockSize As Tuple(Of XUnit, XUnit, List(Of Integer)) = GetBlockDimensions(True)
            Dim XBlockWidth As XUnit = XBlockSize.Item1
            Dim XBlockHeight As XUnit = XBlockSize.Item2
            Dim XExpandedBlockWidth As New XUnit(XBlockSize.Item1.Point + (2 * PDFHelper.BlockSpacer.Point))
            Dim XExpandedBlockHeight As New XUnit(XBlockSize.Item2.Point + (2 * PDFHelper.BlockSpacer.Point))
            Dim fFormItemHeight As Double = XExpandedBlockHeight.Point
            Dim XContentDisplacement As XPoint = GetContentDisplacement()

            FormPDF.AddPage(fFormItemHeight)
            Dim XDisplacement As New XPoint(PDFHelper.PageLimitLeft.Point, PDFHelper.PageLimitTop.Point + ParamList.Value(ParamList.KeyXCurrentHeight).Point)
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(ParamList.Value(ParamList.KeyPDFPages)(ParamList.Value(ParamList.KeyPDFPages).Count - 1).Item1, XGraphicsPdfPageOptions.Append)
                DrawBlockDirect(True, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, XDisplacement, XBlockWidth, XBlockHeight, XContentDisplacement, XBlockSize.Item3)
            End Using
            ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(ParamList.Value(ParamList.KeyXCurrentHeight).Point + XExpandedBlockHeight.Point + PDFHelper.BlockSpacer.Point)
        End Sub
        Public Sub ResetFields(ByVal oGridRect As Int32Rect)
            If IsNothing(oGridRect) Then
                oGridRect = Int32Rect.Empty
            End If
            For Each oGUID As Guid In Children.Keys
                Dim oFormItem As BaseFormItem = FormMain.FindChild(oGUID)

                Dim oFormField As FormField = TryCast(oFormItem, FormField)
                If Not IsNothing(oFormField) Then
                    Dim oGridRectTest As New Rect(oGridRect.X, oGridRect.Y, oGridRect.Width, oGridRect.Height)
                    Dim oGridRectItem As New Rect(oFormField.GridRect.X, oFormField.GridRect.Y, oFormField.GridRect.Width, oFormField.GridRect.Height)
                    If Not (oGridRectTest.Contains(oGridRectItem)) Then
                        oFormField.ResetGridRect()
                    End If
                End If
            Next
        End Sub
#Region "Block Drawing"
        Private Function GetBlockBitmapSource() As Tuple(Of Imaging.BitmapSource, XUnit, XUnit)
            ' gets the block bitmap with margins as indicated by scale direction
            Dim oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = GetBlockDimensions(False)
            Dim XBlockWidth As XUnit = oBlockDimensions.Item1
            Dim XBlockHeight As XUnit = oBlockDimensions.Item2

            Dim oExpandedBlockDimensions As Tuple(Of XUnit, XUnit, XSize) = GetExpandedBlockDimensions(oBlockDimensions)
            Dim XExpandedBlockWidth As XUnit = oExpandedBlockDimensions.Item1
            Dim XExpandedBlockHeight As XUnit = oExpandedBlockDimensions.Item2
            Dim XExpandedBlockSize As XSize = oExpandedBlockDimensions.Item3
            Dim XContentDisplacement As XPoint = GetContentDisplacement()

            Dim width As Integer = Math.Ceiling(XExpandedBlockWidth.Inch * RenderResolution300)
            Dim height As Integer = Math.Ceiling(XExpandedBlockHeight.Inch * RenderResolution300)

            Dim oBitmapSource As Imaging.BitmapSource = Nothing
            Using oBitmap As New System.Drawing.Bitmap(width, height)
                oBitmap.SetResolution(RenderResolution300, RenderResolution300)

                Dim oSection As FormSection = FindParent(GetType(FormSection))
                Dim oSubSection As FormSubSection = TryCast(Parent, FormSubSection)
                Using oParamListNumberingCurrent As New ParamList(ParamList.KeyNumberingCurrent, If(IsNothing(oSection), Nothing, If(IsNothing(oSubSection), oSection.GetNumbering(Me), oSection.GetNumbering(oSubSection))))
                    Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                        oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                        oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                        Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XExpandedBlockSize, XGraphicsUnit.Point)
                            Dim oBlockReturn = DrawBlockDirect(False, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, New XPoint(0, 0), XBlockWidth, XBlockHeight, XContentDisplacement, oBlockDimensions.Item3)
                        End Using
                    End Using

                    oBitmapSource = Converter.BitmapToBitmapSource(oBitmap)
                End Using
            End Using

            Return New Tuple(Of Imaging.BitmapSource, XUnit, XUnit)(oBitmapSource, oBlockDimensions.Item1, oBlockDimensions.Item2)
        End Function
        Public Function DrawBlockDirect(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XBlockWidth As XUnit, XBlockHeight As XUnit, ByVal XDisplacement As XPoint, ByVal XContentWidth As XUnit, XContentHeight As XUnit, ByVal XContentDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing) As Tuple(Of XUnit, XUnit)
            ' draw directly on the supplied XGraphics with the specified displacement
            Dim oContainer As XGraphicsContainer = Nothing

            ' extract fields
            Dim oFormFields As List(Of FormField) = ExtractFields(Of FormField)(GetType(FormField))
            Dim oBackgroundFields As List(Of FieldBackground) = ExtractFields(Of FieldBackground)(GetType(FormField))
            Dim oBorderFields As List(Of FieldBorder) = ExtractFields(Of FieldBorder)(GetType(FormField))
            Dim oNumberingFields As List(Of FieldNumbering) = ExtractFields(Of FieldNumbering)(GetType(FormField))
            oFormFields.RemoveAll(Function(x) oBackgroundFields.Cast(Of FormField).Contains(x))
            oFormFields.RemoveAll(Function(x) oBorderFields.Cast(Of FormField).Contains(x))
            oFormFields.RemoveAll(Function(x) oNumberingFields.Cast(Of FormField).Contains(x))
            oFormFields.RemoveAll(Function(x) Not x.Placed)

            Dim XCombinedDisplacement As New XPoint(XDisplacement.X + XContentDisplacement.X, XDisplacement.Y + XContentDisplacement.Y)

            ' draw background
            oContainer = oXGraphics.BeginContainer()
            Dim oFieldBackgroundWholeBlockList As List(Of FieldBackground) = (From oBackground In oBackgroundFields Where oBackground.WholeBlock Select oBackground).ToList
            Dim oFieldBackgroundList As List(Of FieldBackground) = (From oBackground In oBackgroundFields Where (Not oBackground.WholeBlock) Select oBackground).ToList

            If oFieldBackgroundWholeBlockList.Count > 0 Then
                Select Case oFieldBackgroundWholeBlockList.First.Lightness
                    Case 1
                        PDFHelper.DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.WhiteSmoke)
                    Case 2
                        PDFHelper.DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.LightGray)
                    Case 3
                        PDFHelper.DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.Silver)
                    Case 4
                        PDFHelper.DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.DarkGray)
                End Select
            Else
                PDFHelper.DrawFieldBackground(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, XBrushes.White)
            End If
            For Each oBackground In oFieldBackgroundList
                Select Case oBackground.Lightness
                    Case 1
                        PDFHelper.DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XCombinedDisplacement, XBrushes.WhiteSmoke, GridWidth, GridHeight, oBackground.GridRect)
                    Case 2
                        PDFHelper.DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XCombinedDisplacement, XBrushes.LightGray, GridWidth, GridHeight, oBackground.GridRect)
                    Case 3
                        PDFHelper.DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XCombinedDisplacement, XBrushes.Silver, GridWidth, GridHeight, oBackground.GridRect)
                    Case 4
                        PDFHelper.DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XCombinedDisplacement, XBrushes.DarkGray, GridWidth, GridHeight, oBackground.GridRect)
                End Select
            Next
            oXGraphics.EndContainer(oContainer)

            ' draw numbering
            oContainer = oXGraphics.BeginContainer()
            Dim oSubSection As FormSubSection = TryCast(Parent, FormSubSection)
            If IsNothing(oSubSection) Then
                ' standard processing
                DrawNumbering(oXGraphics, XDisplacement, XContentDisplacement, oNumberingFields, oFieldBackgroundWholeBlockList)
            Else
                ' block is part of a subsection
                Dim oBlockGuidList As List(Of Guid) = (From oChild In oSubSection.Children Where oChild.Value.Equals(GetType(FormBlock)) Select oChild.Key).ToList
                If oBlockGuidList.Count > 0 AndAlso oBlockGuidList.First.Equals(GUID) Then
                    ' only number the first block
                    Dim oSubSectionNumberingFields As List(Of FieldNumbering) = oSubSection.ExtractFields(Of FieldNumbering)(GetType(FormField))
                    DrawNumbering(oXGraphics, XDisplacement, XContentDisplacement, oSubSectionNumberingFields, oFieldBackgroundWholeBlockList)
                End If
            End If
            oXGraphics.EndContainer(oContainer)

            ' runs through each form field that has been placed
            For Each oFormField In oFormFields
                oContainer = oXGraphics.BeginContainer()
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

                ' adds field count for text fields in subsections
                If oFormField.GetType.Equals(GetType(FieldText)) AndAlso ParamList.ContainsKey(ParamList.KeyTextArray) Then
                    Dim oText As FieldText = oFormField
                    Using oParamListKeyFieldCount As New ParamList(ParamList.KeyFieldCount, oText.GetTagNumber.Item2)
                        oFormField.DrawFieldContents(bRender, oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, oOverflowRows)
                    End Using
                Else
                    oFormField.DrawFieldContents(bRender, oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, oOverflowRows)
                End If
                oXGraphics.EndContainer(oContainer)
            Next

            ' draw border
            oContainer = oXGraphics.BeginContainer()
            Dim oFieldBorderWholeBlockList As List(Of FieldBorder) = (From oBorder In oBorderFields Where oBorder.WholeBlock Select oBorder).ToList
            Dim oFieldBorderList As List(Of FieldBorder) = (From oBorder In oBorderFields Where (Not oBorder.WholeBlock) Select oBorder).ToList

            If oFieldBorderWholeBlockList.Count > 0 Then
                PDFHelper.DrawFieldBorder(oXGraphics, XBlockWidth, XBlockHeight, XDisplacement, oFieldBorderWholeBlockList.First.BorderWidth, 0, 0, Nothing)
            End If
            For Each oBorder In oFieldBorderList
                PDFHelper.DrawFieldBorder(oXGraphics, XContentWidth, XContentHeight, XCombinedDisplacement, oBorder.BorderWidth, GridWidth, GridHeight, oBorder.GridRect)
            Next
            oXGraphics.EndContainer(oContainer)

            Return New Tuple(Of XUnit, XUnit)(XBlockWidth, XBlockHeight)
        End Function
        Private Sub DrawNumbering(ByRef oXGraphics As XGraphics, ByVal XDisplacement As XPoint, ByVal XContentDisplacement As XPoint, ByVal oNumberingFields As List(Of FieldNumbering), ByVal oFieldBackgroundWholeBlockList As List(Of FieldBackground))
            ' draws numbering on block
            If oNumberingFields.Count > 0 Then
                Dim oFieldNumbering As FieldNumbering = oNumberingFields.First
                If Not oFieldNumbering.Spacer Then
                    ' set background tint
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
                    If (Not IsNothing(oSection)) AndAlso ParamList.ContainsKey(ParamList.KeyNumberingCurrent) Then
                        With oSection
                            PDFHelper.DrawFieldNumbering(oXGraphics, New XPoint(XDisplacement.X + XContentDisplacement.X, XDisplacement.Y + XContentDisplacement.Y), ParamList.Value(ParamList.KeyNumberingCurrent), If(ParamList.ContainsKey(ParamList.KeySubNumberingCurrent), ParamList.Value(ParamList.KeySubNumberingCurrent), -1), .NumberingType, .NumberingBackground, .NumberingBorder, oBorderBackground)
                        End With
                    End If
                End If
            End If
        End Sub
        Public Function GetBlockDimensions(ByVal bUseOverflows As Boolean, Optional ByVal oTextArray As String() = Nothing) As Tuple(Of XUnit, XUnit, List(Of Integer))
            ' get dimensions of item
            Dim XBlockWidth As XUnit = XUnit.FromPoint((EffectiveBlockWidth * PDFHelper.BlockWidth.Point) + ((EffectiveBlockWidth - 1) * 3 * PDFHelper.BlockSpacer.Point))
            Dim XBlockHeight As XUnit = XUnit.FromPoint(GridHeight * PDFHelper.BlockHeight.Point)

            ' run through each row and extract the maximum overflow for that row and add it to the running total
            Dim fLeftIndent As Double = LeftIndent()
            Dim oOverflowRows As List(Of Integer) = Nothing
            If bUseOverflows Then
                GetBlockDimensionsNext(XBlockWidth, XBlockHeight, fLeftIndent, oTextArray, oOverflowRows)
            End If

            XBlockHeight = XUnit.FromPoint((GridHeight + If(IsNothing(oOverflowRows), 0, oOverflowRows.Sum)) * PDFHelper.BlockHeight.Point)

            Return New Tuple(Of XUnit, XUnit, List(Of Integer))(XBlockWidth, XBlockHeight, oOverflowRows)
        End Function
        Private Sub GetBlockDimensionsNext(ByVal XBlockWidth As XUnit, XBlockHeight As XUnit, ByVal fLeftIndent As Double, ByVal oTextArray As String(), ByRef oOverflowRows As List(Of Integer))
            Dim iFieldCount As Integer = 0
            Dim oFormFields As List(Of FormField) = ((From oChild In Children Where oChild.Value.IsSubclassOf(GetType(FormField)) AndAlso (Not CType(FormMain.FindChild(oChild.Key), FormField).GridRect.IsEmpty) Select CType(FormMain.FindChild(oChild.Key), FormField))).ToList
            For Each oFormField In oFormFields
                Select Case oFormField.GetType
                    Case GetType(FieldText)
                        Dim oNewElements As New List(Of ElementStruc)
                        Dim oFieldText As FieldText = oFormField
                        For Each oElement As ElementStruc In oFieldText.Elements
                            Select Case oElement.ElementType
                                Case ElementStruc.ElementTypeEnum.Field
                                    If IsNothing(oTextArray) Then
                                        oNewElements.Add(oElement)
                                    Else
                                        If iFieldCount >= oTextArray.Length Then
                                            oNewElements.Add(oElement)
                                        Else
                                            oNewElements.Add(New ElementStruc(oTextArray(iFieldCount), ElementStruc.ElementTypeEnum.Text, oElement.FontBold, oElement.FontItalic, oElement.FontUnderline))
                                        End If
                                    End If
                                Case Else
                                    oNewElements.Add(oElement)
                            End Select
                        Next

                        Dim oLayout As Tuple(Of LayoutClass, Double) = PDFHelper.GetLayout(New XUnit((XBlockWidth.Point - fLeftIndent) * (oFormField.GridRect.Width / GridWidth)), New XUnit(XBlockHeight.Point * (oFormField.GridRect.Height / GridHeight)), oFormField.GridRect.Height, oFieldText.FontSizeMultiplier, oNewElements, -1)
                        Dim oLayoutClass As LayoutClass = oLayout.Item1

                        oFormField.GridAdjustedHeight = oLayoutClass.LineCount * oFieldText.FontSizeMultiplier

                        iFieldCount += 1
                End Select
            Next

            oOverflowRows = GetOverflowRows()
        End Sub
        Public Function GetOverflowRows() As List(Of Integer)
            ' get list of overflows by row
            Dim oOverflowRows As New List(Of Integer)
            Dim oFieldTextList As List(Of FieldText) = ((From oChild In Children Where oChild.Value.Equals(GetType(FieldText)) Select CType(FormMain.FindChild(oChild.Key), FieldText))).ToList
            For i = oFieldTextList.Count - 1 To 0 Step -1
                If oFieldTextList(i).GridRect.IsEmpty OrElse oFieldTextList(i).GridAdjustedHeight <= oFieldTextList(i).GridRect.Height Then
                    oFieldTextList.RemoveAt(i)
                End If
            Next

            Dim oFieldList As List(Of FormField) = ((From oChild In Children Where oChild.Value.IsSubclassOf(GetType(FormField)) Select CType(FormMain.FindChild(oChild.Key), FormField))).ToList
            For i = 0 To GridHeight - 1
                If oFieldTextList.Count > 0 Then
                    Dim iCurrentRow As Integer = i
                    Dim oOverflowList As List(Of System.Drawing.Rectangle) = (From oFieldText In oFieldTextList Where oFieldText.GridRect.Y + oFieldText.GridRect.Height - 1 = iCurrentRow Select New System.Drawing.Rectangle(oFieldText.GridRect.X, oFieldText.GridRect.Y + oFieldText.GridRect.Height, oFieldText.GridRect.Width, oFieldText.GridAdjustedHeight - oFieldText.GridRect.Height)).ToList

                    If oOverflowList.Count > 0 Then
                        ' check for overlap with other fields in the overflow area
                        Dim iOverflow As Integer = 0
                        For j = 0 To oOverflowList.Count - 1
                            Dim iNoOverlapCount As Integer = 0
                            For k = oOverflowList(j).Y To oOverflowList(j).Y + oOverflowList(j).Height - 1
                                ' look for fields which are not whole block fields
                                Dim iCurrentJ As Integer = j
                                Dim iCurrentK As Integer = k
                                Dim iFoundOverlap As Integer = Aggregate oFormItem In oFieldList Where (Not oFormItem.SingleOnly) AndAlso Converter.Int32RectToRectangle(oFormItem.GridRect).IntersectsWith(New System.Drawing.Rectangle(oOverflowList(iCurrentJ).X, iCurrentK, oOverflowList(iCurrentJ).Width, 1)) Into Count
                                Dim iFoundSameRow As Integer = Aggregate oFormItem In oFieldList Where (Not oFormItem.SingleOnly) AndAlso Converter.Int32RectToRectangle(oFormItem.GridRect).IntersectsWith(New System.Drawing.Rectangle(0, iCurrentK, GridWidth, 1)) Into Count

                                ' no overlap is detected when there is no field occupying the overlap block and another field is on the same row
                                If iFoundOverlap = 0 And iFoundSameRow > 0 Then
                                    iNoOverlapCount += 1
                                Else
                                    Exit For
                                End If
                            Next
                            iOverflow = Math.Max(iOverflow, oOverflowList(j).Height - iNoOverlapCount)
                        Next
                        oOverflowRows.Add(iOverflow)
                    Else
                        oOverflowRows.Add(0)
                    End If
                Else
                    oOverflowRows.Add(0)
                End If
            Next
            Return oOverflowRows
        End Function
        Private Function GetScaleDirection() As Enumerations.ScaleDirection
            ' gets the scale direction based on the type of block rendered
            Dim oScaleDirection As Enumerations.ScaleDirection = Enumerations.ScaleDirection.None
            Select Case Me.GetType
                Case GetType(FormBlock)
                    oScaleDirection = Enumerations.ScaleDirection.Both
                Case Else
                    oScaleDirection = Enumerations.ScaleDirection.Both
            End Select
            Return oScaleDirection
        End Function
        Public Function GetContentDisplacement() As XPoint
            ' gets content displacement based on the type of block rendered
            Dim oScaleDirection As Enumerations.ScaleDirection = GetScaleDirection()
            Dim fDisplacementX As Double = 0
            Dim fDisplacementY As Double = 0
            If (oScaleDirection And Enumerations.ScaleDirection.Horizontal) = Enumerations.ScaleDirection.Horizontal Then
                fDisplacementX = PDFHelper.BlockSpacer.Point
            End If
            If (oScaleDirection And Enumerations.ScaleDirection.Vertical) = Enumerations.ScaleDirection.Vertical Then
                fDisplacementY = PDFHelper.BlockSpacer.Point
            End If
            Return New XPoint(fDisplacementX, fDisplacementY)
        End Function
        Public Function GetExpandedBlockDimensions(ByVal oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer))) As Tuple(Of XUnit, XUnit, XSize)
            ' gets expanded block size
            Dim oScaleDirection As Enumerations.ScaleDirection = GetScaleDirection()
            Dim XExpandedBlockWidth As XUnit = XUnit.Zero
            Dim XExpandedBlockHeight As XUnit = XUnit.Zero

            ' get expanded block dimensions
            If (oScaleDirection And Enumerations.ScaleDirection.Horizontal) = Enumerations.ScaleDirection.Horizontal Then
                XExpandedBlockWidth = New XUnit(oBlockDimensions.Item1.Point + (2 * PDFHelper.BlockSpacer.Point))
            Else
                XExpandedBlockWidth = oBlockDimensions.Item1
            End If
            If (oScaleDirection And Enumerations.ScaleDirection.Vertical) = Enumerations.ScaleDirection.Vertical Then
                XExpandedBlockHeight = New XUnit(oBlockDimensions.Item2.Point + (2 * PDFHelper.BlockSpacer.Point))
            Else
                XExpandedBlockHeight = oBlockDimensions.Item2
            End If
            Dim XExpandedBlockSize As New XSize(XExpandedBlockWidth.Point, XExpandedBlockHeight.Point)

            Return New Tuple(Of XUnit, XUnit, XSize)(XExpandedBlockWidth, XExpandedBlockHeight, XExpandedBlockSize)
        End Function
#End Region
#Region "Canvas Overlay"
        Private Shared Sub CanvasBlockContentMouseHandler(sender As Object, e As EventArgs) Handles CanvasBlockContent.MouseDown, CanvasBlockContent.MouseUp, CanvasBlockContent.MouseMove, CanvasBlockContent.MouseEnter, CanvasBlockContent.MouseLeave
            ' handles mouse event 
            If CType(Root.BlockPlaceItem, Common.HighlightButton).HBSelected AndAlso ((Not SelectionActive.Item1) Or SelectionActive.Item2 = Enumerations.InputDevice.Mouse) Then
                Dim oLeftButton As Input.MouseButtonState = Nothing
                Dim oPosition As Point = Nothing

                Select Case e.GetType
                    Case GetType(Input.MouseButtonEventArgs)
                        Dim oMouseButtonEventArgs As Input.MouseButtonEventArgs = e
                        oLeftButton = oMouseButtonEventArgs.LeftButton
                        oPosition = oMouseButtonEventArgs.GetPosition(CanvasBlockContent)
                    Case GetType(Input.MouseEventArgs)
                        Dim oMouseEventArgs As Input.MouseEventArgs = e
                        oLeftButton = oMouseEventArgs.LeftButton
                        oPosition = oMouseEventArgs.GetPosition(CanvasBlockContent)
                End Select

                ' check if position is directly over the canvas
                Dim oImageBlockContent As Controls.Image = Root.ImageBlockContent
                Dim fSpacing As Double = PDFHelper.BlockSpacer.Point * CDbl(oImageBlockContent.Tag) * RenderResolution300 / 72

                Dim bOverCanvas As Boolean = False
                If oPosition.X >= fSpacing And oPosition.Y >= fSpacing And oPosition.X < CanvasBlockContent.ActualWidth - fSpacing And oPosition.Y < CanvasBlockContent.ActualHeight - fSpacing Then
                    bOverCanvas = True
                End If

                ' checks to see if an active selection is ongoing
                If SelectionActive.Item1 Then
                    If bOverCanvas Then
                        If oLeftButton = Input.MouseButtonState.Pressed Then
                            ' movement within selection
                            CanvasBlockContentSelection(Enumerations.InputDeviceState.StateMove, oPosition)
                        Else
                            ' end selection
                            CanvasBlockContent.ReleaseMouseCapture()
                            SelectionActive = New Tuple(Of Boolean, Enumerations.InputDevice)(False, Enumerations.InputDevice.NoDevice)
                            CanvasBlockContentSelection(Enumerations.InputDeviceState.StateEnd, oPosition)
                        End If
                    Else
                        ' mouse has left the canvas
                        ' end selection
                        CanvasBlockContent.ReleaseMouseCapture()
                        SelectionActive = New Tuple(Of Boolean, Enumerations.InputDevice)(False, Enumerations.InputDevice.NoDevice)
                        CanvasBlockContentSelection(Enumerations.InputDeviceState.StateEnd, oPosition)
                    End If
                Else
                    If oLeftButton = Input.MouseButtonState.Pressed Then
                        ' start a new selection
                        CanvasBlockContent.CaptureMouse()
                        SelectionActive = New Tuple(Of Boolean, Enumerations.InputDevice)(True, Enumerations.InputDevice.Mouse)
                        CanvasBlockContentSelection(Enumerations.InputDeviceState.StateStart, oPosition)
                    End If
                End If
            End If
        End Sub
        Private Shared Sub CanvasBlockContentSelection(ByVal oState As Enumerations.InputDeviceState, ByVal oLocation As Point)
            ' controls mouse selection for item positioning
            Select Case oState
                Case Enumerations.InputDeviceState.StateStart
                    SelectionStart = oLocation

                    RectangleFormHeaderSelection = New Shapes.Rectangle
                    With RectangleFormHeaderSelection
                        .Width = 0
                        .Height = 0
                        .Stroke = New SolidColorBrush(Color.FromArgb(&HFF, &HFF, &HA5, &H0))
                        .StrokeThickness = 2
                        .StrokeDashArray = New DoubleCollection From {4}
                        .Fill = Brushes.Transparent
                        .Visibility = Visibility.Visible
                    End With
                    Controls.Canvas.SetLeft(RectangleFormHeaderSelection, oLocation.X)
                    Controls.Canvas.SetTop(RectangleFormHeaderSelection, oLocation.Y)
                    Controls.Canvas.SetZIndex(RectangleFormHeaderSelection, 2)
                    CanvasBlockContent.Children.Add(RectangleFormHeaderSelection)
                    CanvasBlockContentFindGrid(SelectionStart, oLocation, False)
                Case Enumerations.InputDeviceState.StateMove
                    With RectangleFormHeaderSelection
                        .Width = Math.Abs(SelectionStart.X - oLocation.X)
                        .Height = Math.Abs(SelectionStart.Y - oLocation.Y)
                    End With
                    Controls.Canvas.SetLeft(RectangleFormHeaderSelection, Math.Min(SelectionStart.X, oLocation.X))
                    Controls.Canvas.SetTop(RectangleFormHeaderSelection, Math.Min(SelectionStart.Y, oLocation.Y))
                    CanvasBlockContentFindGrid(SelectionStart, oLocation, False)
                Case Enumerations.InputDeviceState.StateEnd
                    ' add logic to convert to grid
                    If Not IsNothing(RectangleFormHeaderSelection) Then
                        Dim oRect As Tuple(Of Boolean, Int32Rect) = CanvasBlockContentFindGrid(SelectionStart, oLocation, True)

                        CType(Root.BlockPlaceItem, Common.HighlightButton).HBSelected = False
                        CanvasBlockContent.Children.Remove(RectangleFormHeaderSelection)
                        RectangleFormHeaderSelection = Nothing
                        SelectionStart = Nothing

                        Dim oFormField As FormField = TryCast(SelectedItem, FormField)
                        If Not IsNothing(oFormField) Then
                            Using oSuspender As New Suspender(oFormField, True)
                                If oRect.Item1 Then
                                    oFormField.ResetGridRect()
                                Else
                                    oFormField.GridRect = oRect.Item2
                                End If
                            End Using
                        End If
                    End If
            End Select
        End Sub
        Public Shared Sub BlockPlaceItemClicked()
            ' deselection of the place item button terminates the selection
            If SelectionActive.Item1 AndAlso (SelectionActive.Item2 = Enumerations.InputDevice.Mouse And (Not CType(Root.BlockPlaceItem, Common.HighlightButton).HBSelected)) Then
                SelectionActive = New Tuple(Of Boolean, Enumerations.InputDevice)(False, Enumerations.InputDevice.NoDevice)
                CanvasBlockContentSelection(Enumerations.InputDeviceState.StateEnd, Input.Mouse.GetPosition(CanvasBlockContent))
            End If
        End Sub
        Private Shared Function CanvasBlockContentFindGrid(ByVal oStartLocation As Point, ByVal oCurrentLocation As Point, ByVal bClear As Boolean) As Tuple(Of Boolean, Int32Rect)
            ' converts the start and end point of the selection into grid coordinates and highlights the underlying grid rectangles
            ' returns the grids selected
            Dim oFormField As FormField = TryCast(SelectedItem, FormField)
            Dim oBlock As FormBlock = If(IsNothing(oFormField), Nothing, TryCast(oFormField.Parent, FormBlock))
            If IsNothing(oFormField) Or IsNothing(oBlock) Then
                Return Nothing
            Else
                Dim fLeft As Double = Math.Min(oStartLocation.X, oCurrentLocation.X)
                Dim fTop As Double = Math.Min(oStartLocation.Y, oCurrentLocation.Y)
                Dim fRight As Double = Math.Max(oStartLocation.X, oCurrentLocation.X)
                Dim fBottom As Double = Math.Max(oStartLocation.Y, oCurrentLocation.Y)

                Dim xMin As Integer = FormHeaderRectangleOverlay.GetLength(0) - 1
                Dim yMin As Integer = FormHeaderRectangleOverlay.GetLength(1) - 1
                Dim xMax As Integer = 0
                Dim yMax As Integer = 0
                Dim bRestricted As Boolean = False
                Dim iRectangleWidthCount As Integer = FormHeaderRectangleOverlay.GetLength(0)
                Dim iRectangleHeightCount As Integer = FormHeaderRectangleOverlay.GetLength(1)

                Dim fRectangleLeft As Double = Controls.Canvas.GetLeft(FormHeaderRectangleOverlay(0, 0))
                Dim fRectangleTop As Double = Controls.Canvas.GetTop(FormHeaderRectangleOverlay(0, 0))
                Dim fRectangleWidth As Double = FormHeaderRectangleOverlay(0, 0).ActualWidth
                Dim fRectangleHeight As Double = FormHeaderRectangleOverlay(0, 0).ActualHeight

                Dim iLeft As Integer = Math.Max(Math.Floor((fLeft - fRectangleLeft) / fRectangleWidth), 0)
                Dim iTop As Integer = Math.Max(Math.Floor((fTop - fRectangleTop) / fRectangleHeight), 0)
                Dim iRight As Integer = Math.Min(Math.Floor((fRight - fRectangleLeft) / fRectangleWidth), iRectangleWidthCount - 1)
                Dim iBottom As Integer = Math.Min(Math.Floor((fBottom - fRectangleTop) / fRectangleHeight), iRectangleHeightCount - 1)
                Dim iWidth As Integer = iRight - iLeft + 1
                Dim iHeight As Integer = iBottom - iTop + 1

                ' special processing for fields
                Dim oPlacedFields As List(Of FormField) = (From oChild In oBlock.Children Where (Not oChild.Key.Equals(oFormField.GUID)) AndAlso oChild.Value.IsSubclassOf(GetType(FormField)) AndAlso CType(FormMain.FindChild(oChild.Key), FormField).Placed AndAlso (Not CType(FormMain.FindChild(oChild.Key), FormField).WholeBlock) Select CType(FormMain.FindChild(oChild.Key), FormField)).ToList
                Dim oRectFormField As New Rect(iLeft, iTop, iWidth, iHeight)
                Dim bRectangleColourArray(iRectangleWidthCount - 1, iRectangleHeightCount - 1) As Boolean
                For x = 0 To iRectangleWidthCount - 1
                    For y = 0 To iRectangleHeightCount - 1
                        bRectangleColourArray(x, y) = False
                    Next
                Next

                ' IsRestricted refers to fields which cannot have another field overlapping, and is typical of input fields
                Dim oRestrictedFields As List(Of FormField) = (From oChild In oPlacedFields Where oChild.IsRestricted Select oChild).ToList
                For Each oRestrictedField In oRestrictedFields
                    For x = oRestrictedField.GridRect.X To oRestrictedField.GridRect.X + oRestrictedField.GridRect.Width - 1
                        For y = oRestrictedField.GridRect.Y To oRestrictedField.GridRect.Y + oRestrictedField.GridRect.Height - 1
                            bRectangleColourArray(x, y) = True
                        Next
                    Next
                Next

                ' NoOverlap refers to fields which can surround another field completely like a border or background, but cannot cut through another field
                If oFormField.NoOverlap Then
                    For Each oPlacedField In oPlacedFields
                        Dim oRectPlacedField As New Rect(oPlacedField.GridRect.X, oPlacedField.GridRect.Y, oPlacedField.GridRect.Width, oPlacedField.GridRect.Height)

                        ' fields which have no intersection with the NoOverlap field are completely separate
                        ' fields in union with the NoOverlap field and have the same size as the NoOverlap field are completely surrounded
                        Dim oIntersectField As Rect = Rect.Intersect(oRectFormField, oRectPlacedField)
                        If Not oIntersectField.IsEmpty Then
                            ' intersection is present, check for union
                            Dim oUnionField As Rect = Rect.Union(oRectFormField, oRectPlacedField)
                            If Not oUnionField.Equals(oRectFormField) Then
                                ' the placed field is not completely enclosed by the NoOverlap field
                                For x As Integer = oIntersectField.X To oIntersectField.X + oIntersectField.Width - 1
                                    For y As Integer = oIntersectField.Y To oIntersectField.Y + oIntersectField.Height - 1
                                        bRectangleColourArray(x, y) = True
                                    Next
                                Next
                            End If
                        End If
                    Next
                End If

                ' colour restricted fields red
                For x = 0 To iRectangleWidthCount - 1
                    For y = 0 To iRectangleHeightCount - 1
                        Dim oRectangle As Shapes.Rectangle = FormHeaderRectangleOverlay(x, y)
                        If x >= oRectFormField.X AndAlso y >= oRectFormField.Y AndAlso x < oRectFormField.X + oRectFormField.Width AndAlso y < oRectFormField.Y + oRectFormField.Height Then
                            If bRectangleColourArray(x, y) Then
                                ' restricted rectangle, colour red
                                oRectangle.Fill = New SolidColorBrush(Color.FromArgb(&H66, &HFF, &H0, &H0))
                                bRestricted = True
                            Else
                                ' non-restricted rectangle, colour orange
                                oRectangle.Fill = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                            End If
                        Else
                            ' not selected
                            oRectangle.Fill = Brushes.Transparent
                        End If
                    Next
                Next

                bRestricted = bRestricted Or FieldSpecificProcessing(oRectFormField)

                If bRestricted Then
                    RectangleFormHeaderSelection.Stroke = New SolidColorBrush(Color.FromArgb(&HFF, &HFF, &H0, &H0))
                Else
                    RectangleFormHeaderSelection.Stroke = New SolidColorBrush(Color.FromArgb(&HFF, &HFF, &HA5, &H0))
                End If

                Return New Tuple(Of Boolean, Int32Rect)(bRestricted, New Int32Rect(oRectFormField.X, oRectFormField.Y, oRectFormField.Width, oRectFormField.Height))
            End If
        End Function
        Private Shared Function FieldSpecificProcessing(ByVal oRectFormField As Rect) As Boolean
            ' runs processing specific to each field type to determine additional restrictions
            Dim oFormField As FormField = TryCast(SelectedItem, FormField)
            If IsNothing(oFormField) Then
                Return False
            Else
                Return oFormField.FieldSpecificProcessing(oRectFormField)
            End If
        End Function
#End Region
    End Class
    <DataContract(IsReference:=True)> Public MustInherit Class FormField
        Inherits BaseFormItem

        ' SingleOnly - only one field per block
        ' IsRestricted - fields which cannot have another field overlapping, and is typical of input fields
        ' NoOverlap - fields which can surround another field completely like a border or background, but cannot cut through another field

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public GridRect As Int32Rect
            <DataMember> Public GridAdjustedHeight As Integer
            <DataMember> Public WholeBlock As Boolean
            <DataMember> Public Critical As Boolean
            <DataMember> Public Exclude As Boolean
            <DataMember> Public Alignment As Enumerations.AlignmentEnum

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .GridRect = GridRect
                    .GridAdjustedHeight = GridAdjustedHeight
                    .WholeBlock = WholeBlock
                    .Critical = Critical
                    .Exclude = Exclude
                    .Alignment = Alignment
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            CType(oFormItem, FormField).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Property GridRect As Int32Rect
            Get
                Return m_DataObject.GridRect
            End Get
            Set(value As Int32Rect)
                Dim oBlock As FormBlock = TryCast(Parent, FormBlock)
                m_DataObject.GridRect = value
                m_DataObject.GridAdjustedHeight = value.Height
                RectHeightChanged()
                RectWidthChanged()
                If Not IsNothing(oBlock) Then
                    Using oSuspender As New Suspender(oBlock, True)
                    End Using
                End If
            End Set
        End Property
        Public Property GridAdjustedHeight As Integer
            Get
                Return m_DataObject.GridAdjustedHeight
            End Get
            Set(value As Integer)
                m_DataObject.GridAdjustedHeight = value
            End Set
        End Property
        Public Property WholeBlock As Boolean
            Get
                Return m_DataObject.WholeBlock
            End Get
            Set(value As Boolean)
                Dim oBlock As FormBlock = TryCast(Parent, FormBlock)
                If value AndAlso (Not IsNothing(oBlock)) Then
                    Dim iWholeBlockCount As Integer = Aggregate oChild In Parent.Children Where oChild.Value.Equals(Me.GetType) And CType(FormMain.FindChild(oChild.Key), FormField).WholeBlock Into Count
                    If iWholeBlockCount = 0 Then
                        m_DataObject.WholeBlock = value
                        OnPropertyChangedLocal("WholeBlock")
                        ResetGridRect()
                    End If
                Else
                    m_DataObject.WholeBlock = value
                    OnPropertyChangedLocal("WholeBlock")
                    ResetGridRect()
                End If

                If Not IsNothing(oBlock) Then
                    Using oSuspender As New Suspender(oBlock, True)
                    End Using
                End If
            End Set
        End Property
        Public Property Critical As Boolean
            Get
                Return m_DataObject.Critical
            End Get
            Set(value As Boolean)
                m_DataObject.Critical = value
                OnPropertyChangedLocal("Critical")
            End Set
        End Property
        Public Property Exclude As Boolean
            Get
                Return m_DataObject.Exclude
            End Get
            Set(value As Boolean)
                m_DataObject.Exclude = value
                OnPropertyChangedLocal("Exclude")
            End Set
        End Property
        Protected Overridable Sub RectHeightChanged()
        End Sub
        Protected Overridable Sub RectWidthChanged()
        End Sub
        Public Overrides Sub TitleChanged()
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .GridRect = Int32Rect.Empty
                .GridAdjustedHeight = 0
                .WholeBlock = False
                .Critical = False
                .Exclude = False
                .Alignment = Enumerations.AlignmentEnum.Center
            End With
        End Sub
#End Region
#Region "Alignment"
        Public Property Alignment As Enumerations.AlignmentEnum
            Get
                Return m_DataObject.Alignment
            End Get
            Set(value As Enumerations.AlignmentEnum)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Alignment = value
                    OnPropertyChangedLocal("Alignment")
                    OnPropertyChangedLocal("SelectedL")
                    OnPropertyChangedLocal("SelectedUL")
                    OnPropertyChangedLocal("SelectedT")
                    OnPropertyChangedLocal("SelectedUR")
                    OnPropertyChangedLocal("SelectedR")
                    OnPropertyChangedLocal("SelectedLR")
                    OnPropertyChangedLocal("SelectedB")
                    OnPropertyChangedLocal("SelectedLL")
                    OnPropertyChangedLocal("SelectedC")
                End Using
            End Set
        End Property
        Public ReadOnly Property SelectedL As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.Left Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedUL As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.UpperLeft Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedT As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.Top Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedUR As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.UpperRight Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedR As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.Right Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedLR As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.LowerRight Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedB As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.Bottom Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedLL As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.LowerLeft Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedC As Boolean
            Get
                If m_DataObject.Alignment = Enumerations.AlignmentEnum.Center Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
#End Region
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.0
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                If Not IsNothing(Parent) Then
                    oDisplayFilter.AddRange(Parent.DisplayFilter)
                End If
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            If Not IsNothing(Parent) Then
                Parent.SetBindings()
            End If

            Root.StackPanelWholeBlock.DataContext = Me
            Root.StackPanelAlignmentStart.DataContext = Me
            Root.StackPanelAlignmentExtra1.DataContext = Me
            Root.StackPanelAlignmentExtra2.DataContext = Me
            Root.StackPanelAlignmentBase.DataContext = Me

            Dim oBindingFormField1 As New Data.Binding
            oBindingFormField1.Path = New PropertyPath("WholeBlock")
            oBindingFormField1.Mode = Data.BindingMode.TwoWay
            oBindingFormField1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldWholeBlock.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingFormField1)

            Dim oBindingAlignment1 As New Data.Binding
            oBindingAlignment1.Path = New PropertyPath("SelectedL")
            oBindingAlignment1.Mode = Data.BindingMode.OneWay
            oBindingAlignment1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentL.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment1)

            Dim oBindingAlignment2 As New Data.Binding
            oBindingAlignment2.Path = New PropertyPath("SelectedUL")
            oBindingAlignment2.Mode = Data.BindingMode.OneWay
            oBindingAlignment2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentUL.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment2)

            Dim oBindingAlignment3 As New Data.Binding
            oBindingAlignment3.Path = New PropertyPath("SelectedT")
            oBindingAlignment3.Mode = Data.BindingMode.OneWay
            oBindingAlignment3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentT.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment3)

            Dim oBindingAlignment4 As New Data.Binding
            oBindingAlignment4.Path = New PropertyPath("SelectedUR")
            oBindingAlignment4.Mode = Data.BindingMode.OneWay
            oBindingAlignment4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentUR.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment4)

            Dim oBindingAlignment5 As New Data.Binding
            oBindingAlignment5.Path = New PropertyPath("SelectedR")
            oBindingAlignment5.Mode = Data.BindingMode.OneWay
            oBindingAlignment5.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentR.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment5)

            Dim oBindingAlignment6 As New Data.Binding
            oBindingAlignment6.Path = New PropertyPath("SelectedLR")
            oBindingAlignment6.Mode = Data.BindingMode.OneWay
            oBindingAlignment6.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentLR.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment6)

            Dim oBindingAlignment7 As New Data.Binding
            oBindingAlignment7.Path = New PropertyPath("SelectedB")
            oBindingAlignment7.Mode = Data.BindingMode.OneWay
            oBindingAlignment7.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentB.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment7)

            Dim oBindingAlignment8 As New Data.Binding
            oBindingAlignment8.Path = New PropertyPath("SelectedLL")
            oBindingAlignment8.Mode = Data.BindingMode.OneWay
            oBindingAlignment8.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentLL.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment8)

            Dim oBindingAlignment9 As New Data.Binding
            oBindingAlignment9.Path = New PropertyPath("SelectedC")
            oBindingAlignment9.Mode = Data.BindingMode.OneWay
            oBindingAlignment9.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.AlignmentC.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingAlignment9)
        End Sub
        Public Overrides Sub Display()
            If Not IsNothing(Parent) Then
                Parent.Display()
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return ""
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
        Public MustOverride ReadOnly Property IsRestricted As Boolean
        Public MustOverride ReadOnly Property NoOverlap As Boolean
        Public Shared Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return Nothing
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
        End Sub
        Public MustOverride Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
        Public Overridable Function FieldSpecificProcessing(ByVal oRectFormField As Rect) As Boolean
            ' field specific processing to check whether placement is allowed
            ' default value is false (no problem)
            Return False
        End Function
        Public ReadOnly Property Placed As Boolean
            Get
                If m_DataObject.GridRect.Width = 0 Or m_DataObject.GridRect.Height = 0 Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property
        Protected Shared Function GetFieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            ' creates an ImageSource bitmap with a 4:1 aspect ratio to represent the contents of the field type
            Dim oBackgroundSize As New XSize(XUnit.FromInch(fReferenceHeight * FieldTypeMultiplier * FieldTypeAspectRatio / RenderResolution300).Point, XUnit.FromInch(fReferenceHeight * FieldTypeMultiplier / RenderResolution300).Point)
            Dim iBitmapWidth As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Width).Inch * RenderResolution300)
            Dim iBitmapHeight As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Height).Inch * RenderResolution300)
            Using oBitmap As New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight)
                oBitmap.SetResolution(RenderResolution300, RenderResolution300)
                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, oBackgroundSize, XGraphicsUnit.Point)
                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                        If IsNothing(oMethodInfoGetFieldTypeImageProcess) Then
                            GetFieldTypeImageProcess(fReferenceHeight, oBackgroundSize, oXGraphics, iBitmapWidth, iBitmapHeight)
                        Else
                            oMethodInfoGetFieldTypeImageProcess.Invoke(Nothing, {fReferenceHeight, oBackgroundSize, oXGraphics, iBitmapWidth, iBitmapHeight})
                        End If
                    End Using
                End Using
                Return Converter.BitmapToBitmapSource(oBitmap)
            End Using
        End Function
        Public Sub ResetGridRect()
            m_DataObject.GridRect = Int32Rect.Empty
            m_DataObject.GridAdjustedHeight = 0
            ChangeTransparent()
        End Sub
        Public Overrides Sub ChangeTransparent()
            If Not IsNothing(BorderFormItem) Then
                If SingleOnly() Then
                    BorderFormItem.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                Else
                    If Placed Or WholeBlock Or IsNothing(Parent) Then
                        BorderFormItem.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                    Else
                        BorderFormItem.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
                    End If
                End If
            End If
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FieldBorder
        Inherits FormField

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public BorderWidth As Integer

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .BorderWidth = BorderWidth
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldBorder).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Private Const MaxBorderWidth As Integer = 4
        Public Property BorderWidth As Integer
            Get
                Return m_DataObject.BorderWidth
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 1 Then
                        m_DataObject.BorderWidth = 1
                    ElseIf value > MaxBorderWidth Then
                        m_DataObject.BorderWidth = MaxBorderWidth
                    Else
                        m_DataObject.BorderWidth = value
                    End If
                    OnPropertyChangedLocal("BorderWidth")
                    OnPropertyChangedLocal("BorderWidthText")
                End Using
            End Set
        End Property
        Public Property BorderWidthText As String
            Get
                Return m_DataObject.BorderWidth.ToString
            End Get
            Set(value As String)
                BorderWidth = CInt(Val(value))
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "Border Width: " + BorderWidthText
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            WholeBlock = True

            m_DataObject = New DataObjectClass
            With m_DataObject
                .BorderWidth = 1
            End With
            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Border"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMBorder"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelFieldBorder")
                oDisplayFilter.Add("FieldBorderWidth")
                If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                    oDisplayFilter.Add("StackPanelSelection")
                    oDisplayFilter.Add("StackPanelWholeBlock")
                End If
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return True
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            PDFHelper.DrawFieldBorder(oXGraphics, New XUnit(oBackgroundSize.Width), New XUnit(oBackgroundSize.Height), New XPoint(0, 0), 2)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Throw New NotImplementedException
        End Sub
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelFieldBorder.DataContext = Me

            Dim oBindingBorder1 As New Data.Binding
            oBindingBorder1.Path = New PropertyPath("BorderWidthText")
            oBindingBorder1.Mode = Data.BindingMode.TwoWay
            oBindingBorder1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldBorderWidth.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBorder1)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "9bd4f50e-a5fd-44ff-a2e2-aa607aeea730"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FieldBackground
        Inherits FormField

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public Lightness As Integer

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .Lightness = Lightness
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldBackground).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Private Const MaxLightness As Integer = 4
        Public Property Lightness As Integer
            Get
                Return m_DataObject.Lightness
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 1 Then
                        m_DataObject.Lightness = 1
                    ElseIf value > MaxLightness Then
                        m_DataObject.Lightness = MaxLightness
                    Else
                        m_DataObject.Lightness = value
                    End If
                    OnPropertyChangedLocal("Lightness")
                    OnPropertyChangedLocal("LightnessText")
                End Using
            End Set
        End Property
        Public Property LightnessText As String
            Get
                Return m_DataObject.Lightness.ToString
            End Get
            Set(value As String)
                Lightness = CInt(Val(value))
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "Background: "
            Select Case Lightness
                Case 1
                    Title += "White Smoke"
                Case 2
                    Title += "Light Gray"
                Case 3
                    Title += "Silver"
                Case 4
                    Title += "Dark Gray"
            End Select
            If WholeBlock Then
                Title += vbCr + "Whole Block"
            End If
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            WholeBlock = True

            m_DataObject = New DataObjectClass
            With m_DataObject
                .Lightness = 1
            End With
            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Background"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMBackground"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelFieldBackground")
                oDisplayFilter.Add("FieldBackgroundLightness")
                If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                    oDisplayFilter.Add("StackPanelSelection")
                    oDisplayFilter.Add("StackPanelWholeBlock")
                End If
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return True
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            PDFHelper.DrawFieldBackground(oXGraphics, New XUnit(oBackgroundSize.Width), New XUnit(oBackgroundSize.Height), New XPoint(0, 0), XBrushes.LightGray)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Throw New NotImplementedException
        End Sub
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelFieldBackground.DataContext = Me

            Dim oBindingBackground1 As New Data.Binding
            oBindingBackground1.Path = New PropertyPath("LightnessText")
            oBindingBackground1.Mode = Data.BindingMode.TwoWay
            oBindingBackground1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldBackgroundLightness.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBackground1)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "9fb433da-ed6a-4fc9-8351-5af48042760f"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FieldNumbering
        Inherits FormField

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public Spacer As Boolean

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .Spacer = Spacer
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldNumbering).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Property Spacer As Boolean
            Get
                Return m_DataObject.Spacer
            End Get
            Set(value As Boolean)
                If m_DataObject.Spacer <> value Then
                    Using oSuspender As New Suspender(Me, True, True)
                        m_DataObject.Spacer = value
                        OnPropertyChangedLocal("Spacer")
                    End Using
                End If
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            If Spacer Then
                Title = "Spacer"
            Else
                Title = String.Empty
            End If
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            WholeBlock = True

            m_DataObject = New DataObjectClass
            With m_DataObject
                .Spacer = False
            End With
            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Numbering"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMNumbering"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelNumbering")
                oDisplayFilter.Add("BlockSpacer")

                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return False
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            Const fSizeFraction As Double = 0.8
            Dim oNumberSize As New XUnit(Math.Min(oBackgroundSize.Width, oBackgroundSize.Height) * fSizeFraction)
            Dim oElements As New List(Of ElementStruc)
            oElements.Add(New ElementStruc("1.1.1", ElementStruc.ElementTypeEnum.Text, Enumerations.FontEnum.Bold))
            Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
            PDFHelper.DrawFieldBackground(oXGraphics, New XUnit(oNumberSize.Point * 3), oNumberSize, New XPoint((oXGraphics.PageSize.Width - (oNumberSize.Point * 3)) / 2, (oXGraphics.PageSize.Height - oNumberSize.Point) / 2), XBrushes.LightGray)
            PDFHelper.DrawFieldText(oXGraphics, New XUnit(oNumberSize.Point * 3), oNumberSize, New XPoint((oXGraphics.PageSize.Width - (oNumberSize.Point * 3)) / 2, (oXGraphics.PageSize.Height - oNumberSize.Point) / 2), 1, 1, oElements, -1, MigraDoc.DocumentObjectModel.ParagraphAlignment.Center, oFontColourDictionary)
            PDFHelper.DrawFieldBorder(oXGraphics, New XUnit(oNumberSize.Point * 3), oNumberSize, New XPoint((oXGraphics.PageSize.Width - (oNumberSize.Point * 3)) / 2, (oXGraphics.PageSize.Height - oNumberSize.Point) / 2), 1)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Throw New NotImplementedException
        End Sub
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelNumbering.DataContext = Me

            Dim oBindingNumbering1 As New Data.Binding
            oBindingNumbering1.Path = New PropertyPath("Spacer")
            oBindingNumbering1.Mode = Data.BindingMode.TwoWay
            oBindingNumbering1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockSpacer.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingNumbering1)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "4d0ab196-ed3b-4be2-90d8-0576d2fedfe0"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FieldText
        Inherits FormField

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public Justification As Enumerations.JustificationEnum
            <DataMember> Public FontSizeMultiplier As Integer
            <DataMember> Public Elements As New List(Of ElementStruc)
            <DataMember> Public ElementNumber As Integer

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .Justification = Justification
                    .FontSizeMultiplier = FontSizeMultiplier
                    .Elements.AddRange(Elements)
                    .ElementNumber = Math.Min(ElementNumber, .Elements.Count - 1)
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldText).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Dim WordLimit As Integer = 10
            Const PixelsPerWord As Integer = 350
            If Not IsNothing(Root.GridMain) AndAlso Root.GridMain.ActualWidth > 0 Then
                WordLimit = Math.Max(Root.GridMain.ActualWidth / PixelsPerWord, 1)
            End If

            Dim sText As String = GetAllText()
            Dim oTextArray As String() = sText.Split(" ")
            ReDim Preserve oTextArray(Math.Min(oTextArray.Count - 1, WordLimit - 1))
            Title = String.Join(" ", oTextArray)
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .Justification = Enumerations.JustificationEnum.Center
                .FontSizeMultiplier = 1
                .ElementNumber = -1
            End With
            AddElement()
            TitleChanged()
        End Sub
#End Region
#Region "Text Elements"
        Public Property ElementNumber As Integer
            Get
                Return m_DataObject.ElementNumber + 1
            End Get
            Set(value As Integer)
                If m_DataObject.Elements.Count = 0 Then
                    m_DataObject.ElementNumber = -1
                ElseIf value - 1 >= 0 And value - 1 < m_DataObject.Elements.Count Then
                    m_DataObject.ElementNumber = value - 1
                End If
                ElementPropertyChange()
            End Set
        End Property
        Public Property ElementNumberText As String
            Get
                Return (m_DataObject.ElementNumber + 1).ToString
            End Get
            Set(value As String)
                ElementNumber = CInt(Val(value))
            End Set
        End Property
        Public Property ElementText As String
            Get
                If m_DataObject.Elements.Count = 0 Then
                    Return String.Empty
                ElseIf m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                    Return m_DataObject.Elements(m_DataObject.ElementNumber).Text
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                If (m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count) AndAlso m_DataObject.Elements(m_DataObject.ElementNumber).ElementType = ElementStruc.ElementTypeEnum.Text Then
                    m_DataObject.Elements(m_DataObject.ElementNumber) = New ElementStruc(value, m_DataObject.Elements(m_DataObject.ElementNumber).ElementType, m_DataObject.Elements(m_DataObject.ElementNumber).Font)
                    ElementPropertyChange()
                End If
            End Set
        End Property
        Public Property Element As ElementStruc
            Get
                If m_DataObject.Elements.Count = 0 Then
                    Return New ElementStruc(String.Empty)
                ElseIf m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                    Return m_DataObject.Elements(m_DataObject.ElementNumber)
                Else
                    Return New ElementStruc(String.Empty)
                End If
            End Get
            Set(value As ElementStruc)
                If m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                    m_DataObject.Elements(m_DataObject.ElementNumber) = value
                    ElementPropertyChange()
                End If
            End Set
        End Property
        Public ReadOnly Property Elements As List(Of ElementStruc)
            Get
                Return m_DataObject.Elements
            End Get
        End Property
        Public ReadOnly Property ProcessedElements As List(Of ElementStruc)
            Get
                Return (From iIndex As Integer In Enumerable.Range(0, m_DataObject.Elements.Count) Select GetProcessedText(m_DataObject.Elements(iIndex), iIndex)).ToList
            End Get
        End Property
        Public ReadOnly Property IsActive As Boolean
            Get
                If m_DataObject.Elements.Count = 0 Then
                    Return False
                ElseIf m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                    If m_DataObject.Elements(m_DataObject.ElementNumber).ElementType = ElementStruc.ElementTypeEnum.Text Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            End Get
        End Property
        Public Sub AddElement(ByVal sElementText As String)
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(sElementText))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub AddElement(ByVal sElementText As String, ByVal oElementTypeEnum As ElementStruc.ElementTypeEnum)
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(sElementText, oElementTypeEnum))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub AddElement(ByVal sElementText As String, ByVal oElementTypeEnum As ElementStruc.ElementTypeEnum, ByVal oFont As Enumerations.FontEnum)
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(sElementText, oElementTypeEnum, oFont))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub AddElement(ByVal sElementText As String, ByVal oElementTypeEnum As ElementStruc.ElementTypeEnum, ByVal bBold As Boolean, ByVal bItalic As Boolean, ByVal bUnderline As Boolean)
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(sElementText, oElementTypeEnum, bBold, bItalic, bUnderline))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub AddElement()
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(String.Empty))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub AddElements(ByVal oElements As List(Of ElementStruc))
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                For Each oElement As ElementStruc In oElements
                    m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, oElement)
                    m_DataObject.ElementNumber += 1
                Next
                ElementPropertyChange()
            End If
        End Sub
        Public Sub ReplaceElements(ByVal oElements As List(Of ElementStruc))
            m_DataObject.Elements.Clear()
            m_DataObject.ElementNumber = -1
            AddElements(oElements)
        End Sub
        Public Sub AddTag()
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(String.Empty, ElementStruc.ElementTypeEnum.Field))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub AddSubject()
            If m_DataObject.ElementNumber >= -1 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.Insert(m_DataObject.ElementNumber + 1, New ElementStruc(String.Empty, ElementStruc.ElementTypeEnum.Subject))
                m_DataObject.ElementNumber += 1
                ElementPropertyChange()
            End If
        End Sub
        Public Sub RemoveElement(ByVal iElementNumber As Integer)
            If m_DataObject.Elements.Count > 1 And iElementNumber >= 0 And iElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.RemoveAt(iElementNumber)
                If m_DataObject.Elements.Count = 0 Then
                    m_DataObject.ElementNumber = -1
                Else
                    m_DataObject.ElementNumber = Math.Min(m_DataObject.ElementNumber, m_DataObject.Elements.Count - 1)
                End If
                ElementPropertyChange()
            End If
        End Sub
        Public Sub RemoveElement()
            If m_DataObject.Elements.Count > 1 And m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                m_DataObject.Elements.RemoveAt(m_DataObject.ElementNumber)
                If m_DataObject.Elements.Count = 0 Then
                    m_DataObject.ElementNumber = -1
                Else
                    m_DataObject.ElementNumber = Math.Min(m_DataObject.ElementNumber, m_DataObject.Elements.Count - 1)
                End If
                ElementPropertyChange()
            End If
        End Sub
        Public Sub ClearElements()
            m_DataObject.Elements.Clear()
            m_DataObject.ElementNumber = -1
            ElementPropertyChange()
        End Sub
        Public Sub PreviousElement()
            If m_DataObject.Elements.Count > 0 Then
                If m_DataObject.ElementNumber < 1 Then
                    m_DataObject.ElementNumber = 0
                Else
                    m_DataObject.ElementNumber -= 1
                End If
                m_DataObject.ElementNumber = Math.Min(m_DataObject.ElementNumber, m_DataObject.Elements.Count - 1)
                ElementPropertyChange()
            End If
        End Sub
        Public Sub NextElement()
            If m_DataObject.Elements.Count > 0 Then
                If m_DataObject.ElementNumber > m_DataObject.Elements.Count - 2 Then
                    m_DataObject.ElementNumber = m_DataObject.Elements.Count - 1
                Else
                    m_DataObject.ElementNumber += 1
                End If
                m_DataObject.ElementNumber = Math.Max(m_DataObject.ElementNumber, 0)
                ElementPropertyChange()
            End If
        End Sub
        Public Function GetProcessedText(ByVal oElementStruc As ElementStruc, ByVal iIndex As Integer) As ElementStruc
            Select Case oElementStruc.ElementType
                Case ElementStruc.ElementTypeEnum.Subject
                    Return New ElementStruc("[Subject]")
                Case ElementStruc.ElementTypeEnum.Field
                    Dim oSubSection As FormSubSection = FindParent(GetType(FormSubSection))
                    If IsNothing(oSubSection) Then
                        Return New ElementStruc("[Tag " + Numbering + "-" + (iIndex + 1).ToString + "]")
                    Else
                        Return New ElementStruc(GetTagNumber(iIndex).Item1)
                    End If
                Case ElementStruc.ElementTypeEnum.Template
                    Return New ElementStruc("[Template " + Numbering + "-" + (iIndex + 1).ToString + "]")
                Case Else
                    Return oElementStruc
            End Select
        End Function
        Private Function GetAllText() As String
            ' returns the text of all elements joined together
            Dim sText As String = String.Empty
            For Each oElement In m_DataObject.Elements
                sText += oElement.Text
            Next
            Return sText
        End Function
        Private Sub ElementPropertyChange()
            Using oSuspender As New Suspender(Me, True)
                OnPropertyChangedLocal("ElementNumber")
                OnPropertyChangedLocal("ElementNumberText")
                OnPropertyChangedLocal("Element")
                OnPropertyChangedLocal("ElementText")
                OnPropertyChangedLocal("IsActive")
                OnPropertyChangedLocal("Elements")
                OnPropertyChangedLocal("ElementTexts")
                OnPropertyChangedLocal("Font")
                OnPropertyChangedLocal("FontBold")
                OnPropertyChangedLocal("FontItalic")
                OnPropertyChangedLocal("FontUnderline")
            End Using
        End Sub
        Public Function GetTagNumber(Optional ByVal iIndex As Integer = -1) As Tuple(Of String, Integer)
            ' gets the tag number and index
            Dim oSubSection As FormSubSection = FindParent(GetType(FormSubSection))
            If IsNothing(oSubSection) Then
                Return New Tuple(Of String, Integer)(If(iIndex = -1, String.Empty, "[Tag " + Numbering + "-" + (iIndex + 1).ToString + "]"), -1)
            Else
                Dim oTextList As List(Of FieldText) = oSubSection.GetFormItems(Of FieldText)()
                Dim iTextIndex As Integer = oTextList.IndexOf(Me)
                Return New Tuple(Of String, Integer)(If(iIndex = -1, String.Empty, "[Tag " + oSubSection.Numbering + "-" + (iTextIndex + 1).ToString + "]"), iTextIndex)
            End If
        End Function
#End Region
#Region "Justification"
        Public Property Justification As Enumerations.JustificationEnum
            Get
                Return m_DataObject.Justification
            End Get
            Set(value As Enumerations.JustificationEnum)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Justification = value
                    OnPropertyChangedLocal("Justification")
                    OnPropertyChangedLocal("JustificationLeft")
                    OnPropertyChangedLocal("JustificationCenter")
                    OnPropertyChangedLocal("JustificationRight")
                    OnPropertyChangedLocal("JustificationJustify")
                End Using
            End Set
        End Property
        Public ReadOnly Property JustificationLeft As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Left Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property JustificationCenter As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Center Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property JustificationRight As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Right Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property JustificationJustify As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Justify Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
#End Region
#Region "Font"
        Private Const MaxFontSizeMultiplier As Integer = 8

        Public Property Font As Enumerations.FontEnum
            Get
                If m_DataObject.ElementNumber >= 0 And m_DataObject.ElementNumber < m_DataObject.Elements.Count Then
                    Return m_DataObject.Elements(m_DataObject.ElementNumber).Font
                Else
                    Return Enumerations.FontEnum.None
                End If
            End Get
            Set(value As Enumerations.FontEnum)
                Using oSuspender As New Suspender(Me, True)
                    Dim oFont As Enumerations.FontEnum = m_DataObject.Elements(m_DataObject.ElementNumber).Font Xor value
                    If Elements.Count > 0 Then
                        Element = New ElementStruc(Element.Text, Element.ElementType, oFont)
                    End If
                    OnPropertyChangedLocal("Font")
                    OnPropertyChangedLocal("FontBold")
                    OnPropertyChangedLocal("FontItalic")
                    OnPropertyChangedLocal("FontUnderline")
                End Using
            End Set
        End Property
        Public Property FontSizeMultiplier As Integer
            Get
                Return m_DataObject.FontSizeMultiplier
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 1 Then
                        m_DataObject.FontSizeMultiplier = 1
                    ElseIf value > MaxFontSizeMultiplier Then
                        m_DataObject.FontSizeMultiplier = MaxFontSizeMultiplier
                    Else
                        m_DataObject.FontSizeMultiplier = value
                    End If
                    OnPropertyChangedLocal("FontSizeMultiplier")
                    OnPropertyChangedLocal("FontSizeMultiplierText")
                End Using
            End Set
        End Property
        Public Property FontSizeMultiplierText As String
            Get
                Return m_DataObject.FontSizeMultiplier.ToString
            End Get
            Set(value As String)
                FontSizeMultiplier = CInt(Val(value))
            End Set
        End Property
        Public ReadOnly Property FontBold As Boolean
            Get
                If (m_DataObject.Elements(m_DataObject.ElementNumber).Font And Enumerations.FontEnum.Bold).Equals(Enumerations.FontEnum.Bold) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property FontItalic As Boolean
            Get
                If (m_DataObject.Elements(m_DataObject.ElementNumber).Font And Enumerations.FontEnum.Italic).Equals(Enumerations.FontEnum.Italic) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property FontUnderline As Boolean
            Get
                If (m_DataObject.Elements(m_DataObject.ElementNumber).Font And Enumerations.FontEnum.Underline).Equals(Enumerations.FontEnum.Underline) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Text"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMText"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("BlockContent")
                oDisplayFilter.Add("StackPanelFieldStaticText1")

                ' display add tags only if in subsection
                If Not IsNothing(FindParent(GetType(FormSubSection))) Then
                    oDisplayFilter.Add("StackPanelFieldStaticText2")
                End If

                oDisplayFilter.Add("StackPanelFieldStaticText3")
                oDisplayFilter.Add("StackPanelSelection")
                oDisplayFilter.Add("StackPanelJustification")
                oDisplayFilter.Add("StackPanelFont")
                oDisplayFilter.Add("FontSizeMultiplier")
                oDisplayFilter.Add("StaticTextCurrent")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return False
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            Dim oElements As New List(Of ElementStruc)
            For i = 0 To PDFHelper.LoremIpsum.Count - 2
                oElements.Add(New ElementStruc(PDFHelper.LoremIpsum(i), ElementStruc.ElementTypeEnum.Text, Enumerations.FontEnum.None))
                oElements.Add(New ElementStruc(" [Field " + i.ToString + "] ", ElementStruc.ElementTypeEnum.Field, Enumerations.FontEnum.Bold))
            Next
            oElements.Add(New ElementStruc(PDFHelper.LoremIpsum(PDFHelper.LoremIpsum.Count - 1), ElementStruc.ElementTypeEnum.Text, Enumerations.FontEnum.None))

            Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Field, MigraDoc.DocumentObjectModel.Colors.Red)

            PDFHelper.DrawFieldText(oXGraphics, oBackgroundSize.Width, oBackgroundSize.Height, New XPoint((oXGraphics.PageSize.Width - oBackgroundSize.Width) / 2, (oXGraphics.PageSize.Height - oBackgroundSize.Height) / 2), 5, 1, oElements, -1, MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify, oFontColourDictionary)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim oCurrentSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            ' optional parameters
            Dim oTextArray As String() = If(ParamList.ContainsKey(ParamList.KeyTextArray), ParamList.Value(ParamList.KeyTextArray), Nothing)
            Dim iFieldCount As Integer = If(ParamList.ContainsKey(ParamList.KeyFieldCount), ParamList.Value(ParamList.KeyFieldCount), -1)
            Dim sSubjectName As String = If(ParamList.ContainsKey(ParamList.KeyCurrentSubject), ParamList.Value(ParamList.KeyCurrentSubject), String.Empty)

            Dim oXRect As New XRect(XFieldDisplacement.X, XFieldDisplacement.Y, XImageWidth.Point, XImageHeight.Point)
            oXGraphics.IntersectClip(oXRect)

            Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Subject, MigraDoc.DocumentObjectModel.Colors.DarkViolet)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Field, MigraDoc.DocumentObjectModel.Colors.Green)

            Dim oJustification As MigraDoc.DocumentObjectModel.ParagraphAlignment = Nothing
            Select Case Justification
                Case Enumerations.JustificationEnum.Left
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left
                Case Enumerations.JustificationEnum.Center
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Center
                Case Enumerations.JustificationEnum.Right
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Right
                Case Enumerations.JustificationEnum.Justify
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify
            End Select

            ' the text array if supplied is used for rendering while if not present, then a display format is assumed
            If bRender Then
                ' render format
                Dim oNewElements As New List(Of ElementStruc)
                For Each oElement As ElementStruc In Elements
                    Select Case oElement.ElementType
                        Case ElementStruc.ElementTypeEnum.Field
                            If IsNothing(oTextArray) OrElse iFieldCount >= oTextArray.Length Then
                                oNewElements.Add(oElement)
                            Else
                                oNewElements.Add(New ElementStruc(oTextArray(iFieldCount), ElementStruc.ElementTypeEnum.Text, oElement.FontBold, oElement.FontItalic, oElement.FontUnderline))
                            End If
                        Case ElementStruc.ElementTypeEnum.Subject
                            oNewElements.Add(New ElementStruc(sSubjectName, ElementStruc.ElementTypeEnum.Text, oElement.FontBold, oElement.FontItalic, oElement.FontUnderline))
                        Case Else
                            oNewElements.Add(oElement)
                    End Select
                Next
                PDFHelper.DrawFieldText(oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, GridAdjustedHeight, FontSizeMultiplier, oNewElements, -1, oJustification, oFontColourDictionary)
            Else
                ' display format
                If CType(Parent, FormBlock).ShowActive Then
                    PDFHelper.DrawFieldText(oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, If(IsNothing(oOverflowRows), GridRect.Height, GridAdjustedHeight), FontSizeMultiplier, ProcessedElements, ElementNumber - 1, oJustification, oFontColourDictionary)
                Else
                    PDFHelper.DrawFieldText(oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, If(IsNothing(oOverflowRows), GridRect.Height, GridAdjustedHeight), FontSizeMultiplier, ProcessedElements, -1, oJustification, oFontColourDictionary)
                End If
            End If

            oXGraphics.SmoothingMode = oCurrentSmoothingMode
        End Sub
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelFieldStaticText1.DataContext = Me
            Root.StackPanelFieldStaticText2.DataContext = Me
            Root.StackPanelFieldStaticText3.DataContext = Me
            Root.BlockContent.DataContext = Me

            ' bind to parent block
            Dim oBlock As FormBlock = TryCast(Parent, FormBlock)
            If Not IsNothing(oBlock) Then
                Root.StaticTextShowActive.DataContext = oBlock

                Dim oBindingBlock1 As New Data.Binding
                oBindingBlock1.Path = New PropertyPath("ShowActive")
                oBindingBlock1.Mode = Data.BindingMode.OneWay
                oBindingBlock1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                Root.StaticTextShowActive.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingBlock1)
            End If

            Dim oBindingText1 As New Data.Binding
            oBindingText1.Path = New PropertyPath("ElementNumberText")
            oBindingText1.Mode = Data.BindingMode.TwoWay
            oBindingText1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.StaticTextCurrent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingText1)

            Dim oBindingText2 As New Data.Binding
            oBindingText2.Path = New PropertyPath("ElementText")
            oBindingText2.Mode = Data.BindingMode.TwoWay
            oBindingText2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingText2)

            Dim oBindingText3 As New Data.Binding
            oBindingText3.Path = New PropertyPath("IsActive")
            oBindingText3.Mode = Data.BindingMode.OneWay
            oBindingText3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockContent.SetBinding(Common.HighlightTextBox.IsEnabledProperty, oBindingText3)

            BindingJustification(Me)
            BindingFont(Me)
        End Sub
        Public Shared Sub BindingJustification(ByVal oFormItem As BaseFormItem)
            Root.StackPanelJustification.DataContext = oFormItem

            Dim oBindingJustification1 As New Data.Binding
            oBindingJustification1.Path = New PropertyPath("JustificationLeft")
            oBindingJustification1.Mode = Data.BindingMode.OneWay
            oBindingJustification1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.JustificationLeft.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingJustification1)

            Dim oBindingJustification2 As New Data.Binding
            oBindingJustification2.Path = New PropertyPath("JustificationCenter")
            oBindingJustification2.Mode = Data.BindingMode.OneWay
            oBindingJustification2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.JustificationCenter.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingJustification2)

            Dim oBindingJustification3 As New Data.Binding
            oBindingJustification3.Path = New PropertyPath("JustificationRight")
            oBindingJustification3.Mode = Data.BindingMode.OneWay
            oBindingJustification3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.JustificationRight.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingJustification3)

            Dim oBindingJustification4 As New Data.Binding
            oBindingJustification4.Path = New PropertyPath("JustificationJustify")
            oBindingJustification4.Mode = Data.BindingMode.OneWay
            oBindingJustification4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.JustificationJustify.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingJustification4)
        End Sub
        Public Shared Sub BindingFont(ByVal oFormItem As BaseFormItem)
            Root.StackPanelFont.DataContext = oFormItem

            Dim oBindingFont1 As New Data.Binding
            oBindingFont1.Path = New PropertyPath("FontBold")
            oBindingFont1.Mode = Data.BindingMode.OneWay
            oBindingFont1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FontBold.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingFont1)

            Dim oBindingFont2 As New Data.Binding
            oBindingFont2.Path = New PropertyPath("FontItalic")
            oBindingFont2.Mode = Data.BindingMode.OneWay
            oBindingFont2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FontItalic.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingFont2)

            Dim oBindingFont3 As New Data.Binding
            oBindingFont3.Path = New PropertyPath("FontUnderline")
            oBindingFont3.Mode = Data.BindingMode.OneWay
            oBindingFont3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FontUnderline.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingFont3)

            Dim oBindingFont4 As New Data.Binding
            oBindingFont4.Path = New PropertyPath("FontSizeMultiplierText")
            oBindingFont4.Mode = Data.BindingMode.TwoWay
            oBindingFont4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FontSizeMultiplier.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingFont4)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "aec241fb-f012-4809-9f62-5bfb50907351"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FieldImage
        Inherits FormField

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public Stretch As Enumerations.StretchEnum
            <DataMember> Public Image As System.Drawing.Image
            <DataMember> Public ImageName As String

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .Stretch = Stretch
                    .Image = If(IsNothing(Image), Nothing, Image.Clone)
                    .ImageName = ImageName
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldImage).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Dim sTitle As String = "Alignment: "
            Select Case Alignment
                Case Enumerations.AlignmentEnum.Left
                    sTitle += "Left"
                Case Enumerations.AlignmentEnum.Top
                    sTitle += "Top"
                Case Enumerations.AlignmentEnum.Right
                    sTitle += "Right"
                Case Enumerations.AlignmentEnum.Bottom
                    sTitle += "Bottom"
                Case Enumerations.AlignmentEnum.Center
                    sTitle += "Center"
            End Select
            Select Case Stretch
                Case Enumerations.StretchEnum.Fill
                    sTitle += vbCr + "Stretch: Fill"
                Case Enumerations.StretchEnum.Uniform
                    sTitle += vbCr + "Stretch: Uniform"
            End Select
            sTitle += vbCr + "Image: " + m_DataObject.ImageName
            Title = sTitle
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .Stretch = Enumerations.StretchEnum.Uniform
            End With
            TitleChanged()
        End Sub
#End Region
#Region "Stretch"
        Public Property Stretch As Enumerations.StretchEnum
            Get
                Return m_DataObject.Stretch
            End Get
            Set(value As Enumerations.StretchEnum)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Stretch = value
                    OnPropertyChangedLocal("Stretch")
                    OnPropertyChangedLocal("StretchFill")
                    OnPropertyChangedLocal("StretchUniform")
                End Using
            End Set
        End Property
        Public ReadOnly Property StretchFill As Boolean
            Get
                If m_DataObject.Stretch = Enumerations.StretchEnum.Fill Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property StretchUniform As Boolean
            Get
                If m_DataObject.Stretch = Enumerations.StretchEnum.Uniform Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
#End Region
#Region "Image"
        Private m_XImage As XImage
        Private m_ImageChanged As Boolean = True
        Private m_SetWidth As Integer = 0
        Private m_SetHeight As Integer = 0
        Private m_SetStretch As Enumerations.StretchEnum = Enumerations.StretchEnum.Fill

        Public ReadOnly Property XImage(ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal oStretch As Enumerations.StretchEnum, Optional bReset As Boolean = False) As XImage
            Get
                ' resamples if necessary
                If (Not IsNothing(m_DataObject.Image)) AndAlso (bReset Or m_SetWidth <> iWidth Or m_SetHeight <> iHeight Or m_SetStretch <> oStretch Or m_ImageChanged) Then
                    m_ImageChanged = False

                    Dim iStretchWidth As Integer = 0
                    Dim iStretchHeight As Integer = 0

                    Select Case oStretch
                        Case Enumerations.StretchEnum.Fill
                            iStretchWidth = iWidth
                            iStretchHeight = iHeight
                        Case Enumerations.StretchEnum.Uniform
                            Dim fScale As Double = Math.Min(iWidth / m_DataObject.Image.Width, iHeight / m_DataObject.Image.Height)
                            iStretchWidth = m_DataObject.Image.Width * fScale
                            iStretchHeight = m_DataObject.Image.Height * fScale
                    End Select

                    Dim oResampledBitmap As New System.Drawing.Bitmap(m_DataObject.Image, iStretchWidth, iStretchHeight)
                    oResampledBitmap.SetResolution(RenderResolution300, RenderResolution300)
                    m_XImage = XImage.FromGdiPlusImage(oResampledBitmap)

                    m_SetWidth = iWidth
                    m_SetHeight = iHeight
                    m_SetStretch = oStretch
                End If
                Return m_XImage
            End Get
        End Property
        Public ReadOnly Property ImageName As String
            Get
                Return m_DataObject.ImageName
            End Get
        End Property
        Public ReadOnly Property ImagePresent As Boolean
            Get
                If IsNothing(m_DataObject.Image) Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property
        Public Sub SetImage(ByRef oImage As System.Drawing.Image, ByVal sImageName As String)
            Using oSuspender As New Suspender(Me, True)
                m_DataObject.Image = oImage
                m_DataObject.ImageName = sImageName
                m_ImageChanged = True
            End Using
        End Sub
        Public Function GetImage() As System.Drawing.Bitmap
            If IsNothing(m_DataObject.Image) Then
                Return Nothing
            Else
                Return Converter.BitmapConvertGrayscale(m_DataObject.Image)
            End If
        End Function
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Image"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMImage"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelSelection")
                oDisplayFilter.Add("StackPanelAlignmentStart")
                oDisplayFilter.Add("StackPanelAlignmentExtra1")
                oDisplayFilter.Add("StackPanelAlignmentBase")
                oDisplayFilter.Add("StackPanelStretch")
                oDisplayFilter.Add("StackPanelFieldImage")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Dim oImage As DrawingImage = CType(m_Icons("CCMImageSample").Clone, DrawingImage)
            Dim oBackgroundSize As New XSize(XUnit.FromInch(fReferenceHeight * FieldTypeMultiplier * FieldTypeAspectRatio / RenderResolution300).Point, XUnit.FromInch(fReferenceHeight * FieldTypeMultiplier / RenderResolution300).Point)
            Dim iBitmapWidth As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Width).Inch * RenderResolution300)
            Dim iBitmapHeight As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Height).Inch * RenderResolution300)
            Return Converter.DrawingImageToBitmapSource(oImage, iBitmapWidth, iBitmapHeight, Enumerations.StretchEnum.Uniform)
        End Function
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim oXImage As XImage = XImage(XImageWidth.Inch * RenderResolution300, XImageHeight.Inch * RenderResolution300, Stretch)

            If ImagePresent Then
                Dim fAlignmentDisplacementX As Double = 0
                Dim fAlignmentDisplacementY As Double = 0
                Select Case Alignment
                    Case Enumerations.AlignmentEnum.Left
                        fAlignmentDisplacementY = (XImageHeight.Point - oXImage.PointHeight) / 2
                    Case Enumerations.AlignmentEnum.UpperLeft
                    Case Enumerations.AlignmentEnum.Top
                        fAlignmentDisplacementX = (XImageWidth.Point - oXImage.PointWidth) / 2
                    Case Enumerations.AlignmentEnum.UpperRight
                        fAlignmentDisplacementX = XImageWidth.Point - oXImage.PointWidth
                    Case Enumerations.AlignmentEnum.Right
                        fAlignmentDisplacementX = XImageWidth.Point - oXImage.PointWidth
                        fAlignmentDisplacementY = (XImageHeight.Point - oXImage.PointHeight) / 2
                    Case Enumerations.AlignmentEnum.LowerRight
                        fAlignmentDisplacementX = XImageWidth.Point - oXImage.PointWidth
                        fAlignmentDisplacementY = XImageHeight.Point - oXImage.PointHeight
                    Case Enumerations.AlignmentEnum.Bottom
                        fAlignmentDisplacementX = (XImageWidth.Point - oXImage.PointWidth) / 2
                        fAlignmentDisplacementY = XImageHeight.Point - oXImage.PointHeight
                    Case Enumerations.AlignmentEnum.LowerLeft
                        fAlignmentDisplacementY = XImageHeight.Point - oXImage.PointHeight
                    Case Enumerations.AlignmentEnum.Center
                        fAlignmentDisplacementX = (XImageWidth.Point - oXImage.PointWidth) / 2
                        fAlignmentDisplacementY = (XImageHeight.Point - oXImage.PointHeight) / 2
                End Select
                oXGraphics.DrawImage(oXImage, XFieldDisplacement.X + fAlignmentDisplacementX, XFieldDisplacement.Y + fAlignmentDisplacementY, oXImage.PointWidth, oXImage.PointHeight)
            End If
        End Sub
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelStretch.DataContext = Me

            Dim oBindingStretch1 As New Data.Binding
            oBindingStretch1.Path = New PropertyPath("StretchFill")
            oBindingStretch1.Mode = Data.BindingMode.OneWay
            oBindingStretch1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.StretchFill.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingStretch1)

            Dim oBindingStretch2 As New Data.Binding
            oBindingStretch2.Path = New PropertyPath("StretchUniform")
            oBindingStretch2.Mode = Data.BindingMode.OneWay
            oBindingStretch2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.StretchUniform.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingStretch2)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "2a91d005-4e55-48f3-87ee-120f754fab8d"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FieldChoice
        Inherits FormField

        <DataMember> Protected m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public TabletCount As Integer
            <DataMember> Public TabletContent As Enumerations.TabletContentEnum
            <DataMember> Public TabletStart As Integer
            <DataMember> Public TabletGroups As Integer
            <DataMember> Public TabletDescriptionTop As New List(Of String)
            <DataMember> Public TabletDescriptionBottom As New List(Of String)
            <DataMember> Public TabletDescriptionMCQ As New List(Of List(Of ElementStruc))
            <DataMember> Public CurrentDescriptionTop As Integer
            <DataMember> Public CurrentDescriptionBottom As Integer
            <DataMember> Public BlockCount As Integer
            <DataMember> Public HandwritingContent As String()

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .TabletCount = TabletCount
                    .TabletContent = TabletContent
                    .TabletStart = TabletStart
                    .TabletGroups = TabletGroups
                    .TabletDescriptionTop.AddRange(TabletDescriptionTop)
                    .TabletDescriptionBottom.AddRange(TabletDescriptionBottom)
                    If Not IsNothing(TabletDescriptionMCQ) Then
                        .TabletDescriptionMCQ.AddRange(TabletDescriptionMCQ)
                    End If
                    .CurrentDescriptionTop = CurrentDescriptionTop
                    .CurrentDescriptionBottom = CurrentDescriptionBottom
                    .BlockCount = BlockCount
                    .HandwritingContent = If(IsNothing(HandwritingContent), Nothing, HandwritingContent.Clone)
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldChoice).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Protected Overrides Sub RectHeightChanged()
            MyBase.RectHeightChanged()

            OnPropertyChangedLocal("TabletGroups")
            OnPropertyChangedLocal("TabletGroupsText")
        End Sub
        Protected Overrides Sub RectWidthChanged()
            MyBase.RectWidthChanged()

            Dim iMaxTabletCount As Integer = GetMaxTabletCount()
            If m_DataObject.TabletCount > iMaxTabletCount Then
                m_DataObject.TabletCount = iMaxTabletCount
                OnPropertyChangedLocal("TabletCount")
                OnPropertyChangedLocal("TabletCountText")
            End If
        End Sub
        Public Overrides Sub TitleChanged()
            Title = "Horizontal" + vbCr + "Tablet Count: " + TabletCountText + If(Critical, vbCrLf + "Critical Field", String.Empty)
            Title += If(Exclude, vbCr + "Excluded", String.Empty)
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .TabletCount = 0
                .TabletContent = Enumerations.TabletContentEnum.None
                .TabletStart = 0
                .TabletGroups = 1
                .CurrentDescriptionTop = -1
                .CurrentDescriptionBottom = -1
            End With
            TitleChanged()
        End Sub
#End Region
#Region "Choice"
        Public Property TabletCount As Integer
            Get
                Return m_DataObject.TabletCount
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                        Dim iMaxTabletCount As Integer = GetMaxTabletCount()
                        If value < 0 Then
                            m_DataObject.TabletCount = 0
                        ElseIf value > iMaxTabletCount Then
                            m_DataObject.TabletCount = iMaxTabletCount
                        Else
                            m_DataObject.TabletCount = value
                        End If
                    Else
                        m_DataObject.TabletCount = 0
                    End If

                    ' add or remove tablets to match tablet count
                    ' if a tablet descriptor is not empty, then do not remove
                    If m_DataObject.TabletDescriptionTop.Count > m_DataObject.TabletCount Then
                        For i = m_DataObject.TabletDescriptionTop.Count - 1 To m_DataObject.TabletCount Step -1
                            If Trim(m_DataObject.TabletDescriptionTop(i)) = String.Empty Then
                                m_DataObject.TabletDescriptionTop.RemoveAt(i)
                            Else
                                Exit For
                            End If
                        Next
                    ElseIf m_DataObject.TabletDescriptionTop.Count < m_DataObject.TabletCount Then
                        For i = 0 To m_DataObject.TabletCount - m_DataObject.TabletDescriptionTop.Count - 1
                            m_DataObject.TabletDescriptionTop.Add(String.Empty)
                        Next
                    End If
                    If m_DataObject.TabletDescriptionBottom.Count > m_DataObject.TabletCount Then
                        For i = m_DataObject.TabletDescriptionBottom.Count - 1 To m_DataObject.TabletCount Step -1
                            If Trim(m_DataObject.TabletDescriptionBottom(i)) = String.Empty Then
                                m_DataObject.TabletDescriptionBottom.RemoveAt(i)
                            Else
                                Exit For
                            End If
                        Next
                    ElseIf m_DataObject.TabletDescriptionBottom.Count < m_DataObject.TabletCount Then
                        For i = 0 To m_DataObject.TabletCount - m_DataObject.TabletDescriptionBottom.Count - 1
                            m_DataObject.TabletDescriptionBottom.Add(String.Empty)
                        Next
                    End If
                    If m_DataObject.TabletDescriptionTop.Count = 0 Then
                        m_DataObject.CurrentDescriptionTop = -1
                    Else
                        m_DataObject.CurrentDescriptionTop = Math.Min(Math.Max(0, m_DataObject.CurrentDescriptionTop), m_DataObject.TabletCount - 1)
                    End If
                    If m_DataObject.TabletDescriptionBottom.Count = 0 Then
                        m_DataObject.CurrentDescriptionBottom = -1
                    Else
                        m_DataObject.CurrentDescriptionBottom = Math.Min(Math.Max(0, m_DataObject.CurrentDescriptionBottom), m_DataObject.TabletCount - 1)
                    End If

                    OnPropertyChangedLocal("TabletCount")
                    OnPropertyChangedLocal("TabletCountText")
                    OnPropertyChangedLocal("TabletGroups")
                    OnPropertyChangedLocal("TabletGroupsText")
                    OnPropertyChangedLocal("CurrentDescriptionTop")
                    OnPropertyChangedLocal("CurrentDescriptionTopText")
                    OnPropertyChangedLocal("CurrentDescriptionBottom")
                    OnPropertyChangedLocal("CurrentDescriptionBottomText")
                    OnPropertyChangedLocal("TabletDescriptionTopCurrent")
                    OnPropertyChangedLocal("TabletDescriptionBottomCurrent")
                End Using
            End Set
        End Property
        Public Property TabletCountText As String
            Get
                Return m_DataObject.TabletCount.ToString
            End Get
            Set(value As String)
                TabletCount = CInt(Val(value))
            End Set
        End Property
        Public Property TabletContent As Enumerations.TabletContentEnum
            Get
                Return m_DataObject.TabletContent
            End Get
            Set(value As Enumerations.TabletContentEnum)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.TabletContent = value

                    ' reset tablet start if necessary
                    If m_DataObject.TabletContent = Enumerations.TabletContentEnum.Letter And m_DataObject.TabletStart = -1 Then
                        m_DataObject.TabletStart = 0
                    End If

                    OnPropertyChangedLocal("TabletContent")
                    OnPropertyChangedLocal("TabletContentText")
                    OnPropertyChangedLocal("TabletStart")
                    OnPropertyChangedLocal("TabletStartText")
                End Using
            End Set
        End Property
        Public Property TabletContentText As String
            Get
                Return [Enum].GetNames(GetType(Enumerations.TabletContentEnum))(CType(m_DataObject.TabletContent, Integer))
            End Get
            Set(value As String)
            End Set
        End Property
        Public Property TabletStart As Integer
            Get
                Return m_DataObject.TabletStart
            End Get
            Set(value As Integer)
                If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                    Using oSuspender As New Suspender(Me, True)
                        Const MaxTabletStart As Integer = 25
                        Dim iLowerLimit As Integer = If(m_DataObject.TabletContent = Enumerations.TabletContentEnum.Letter, 0, -1)
                        If value < iLowerLimit Then
                            m_DataObject.TabletStart = iLowerLimit
                        ElseIf value > MaxTabletStart Then
                            m_DataObject.TabletStart = MaxTabletStart
                        Else
                            m_DataObject.TabletStart = value
                        End If

                        OnPropertyChangedLocal("TabletStart")
                        OnPropertyChangedLocal("TabletStartText")
                    End Using
                End If
            End Set
        End Property
        Public Property TabletStartText As String
            Get
                Return (TabletStart + 1).ToString
            End Get
            Set(value As String)
                TabletStart = CInt(Val(value)) - 1
            End Set
        End Property
        Public Property TabletGroups As Integer
            Get
                Return Math.Max(m_DataObject.TabletGroups, 1)
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    Dim iLowerLimit As Integer = 1
                    If (Not IsNothing(Parent)) AndAlso (Parent.GetType.Equals(GetType(FormBlock)) Or Parent.GetType.Equals(GetType(FormMCQ))) Then
                        Dim MaxTabletGroups As Integer = 0
                        Select Case Me.GetType
                            Case GetType(FieldChoice)
                                MaxTabletGroups = Math.Min(Math.Floor(GridRect.Height / 3), TabletCount)
                            Case GetType(FieldChoiceVerticalMCQ)
                                Dim oBlock As FormBlock = Parent
                                MaxTabletGroups = oBlock.BlockWidth
                            Case Else
                                MaxTabletGroups = GridRect.Width
                        End Select
                        If value < iLowerLimit Then
                            m_DataObject.TabletGroups = iLowerLimit
                        ElseIf value > MaxTabletGroups Then
                            m_DataObject.TabletGroups = MaxTabletGroups
                        Else
                            m_DataObject.TabletGroups = value
                        End If
                    Else
                        m_DataObject.TabletGroups = iLowerLimit
                    End If

                    OnPropertyChangedLocal("TabletGroups")
                    OnPropertyChangedLocal("TabletGroupsText")
                End Using
            End Set
        End Property
        Public Property TabletGroupsText As String
            Get
                Return TabletGroups.ToString
            End Get
            Set(value As String)
                TabletGroups = CInt(Val(value))
            End Set
        End Property
        Protected Overridable Function GetMaxTabletCount() As Integer
            Dim oBlock As FormBlock = Parent
            If IsNothing(oBlock) Then
                Return 0
            Else
                ' each tablet width is equals to two row heights, but need to account for tablet groups
                Dim fImageWidth As Double = ((oBlock.GetBlockDimensions(False).Item1.Point - oBlock.LeftIndent.Point) * GridRect.Width / oBlock.GridWidth) - (2 * SpacingLarge * 72 / RenderResolution300)
                Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
                Dim iMaxTabletCount As Integer = 0
                If fImageWidth >= fSingleBlockWidth Then
                    iMaxTabletCount = Math.Floor(fImageWidth / fSingleBlockWidth)
                End If
                Return iMaxTabletCount * TabletGroups
            End If
        End Function
        Public Property CurrentDescriptionTop As Integer
            Get
                Return m_DataObject.CurrentDescriptionTop
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If m_DataObject.TabletDescriptionTop.Count = 0 Then
                        m_DataObject.CurrentDescriptionTop = -1
                    Else
                        Dim iMaxCount As Integer = TabletCount
                        If value < 0 Then
                            m_DataObject.CurrentDescriptionTop = 0
                        ElseIf value > iMaxCount - 1 Then
                            m_DataObject.CurrentDescriptionTop = iMaxCount - 1
                        Else
                            m_DataObject.CurrentDescriptionTop = value
                        End If
                    End If

                    OnPropertyChangedLocal("CurrentDescriptionTop")
                    OnPropertyChangedLocal("CurrentDescriptionTopText")
                    OnPropertyChangedLocal("TabletDescriptionTopCurrent")
                End Using
            End Set
        End Property
        Public Property CurrentDescriptionTopText As String
            Get
                Return (m_DataObject.CurrentDescriptionTop + 1).ToString
            End Get
            Set(value As String)
                CurrentDescriptionTop = CInt(Val(value)) - 1
            End Set
        End Property
        Public Property CurrentDescriptionBottom As Integer
            Get
                Return m_DataObject.CurrentDescriptionBottom
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If m_DataObject.TabletDescriptionBottom.Count = 0 Then
                        m_DataObject.CurrentDescriptionBottom = -1
                    Else
                        If value < 0 Then
                            m_DataObject.CurrentDescriptionBottom = 0
                        ElseIf value > TabletCount - 1 Then
                            m_DataObject.CurrentDescriptionBottom = TabletCount - 1
                        Else
                            m_DataObject.CurrentDescriptionBottom = value
                        End If
                    End If

                    OnPropertyChangedLocal("CurrentDescriptionBottom")
                    OnPropertyChangedLocal("CurrentDescriptionBottomText")
                    OnPropertyChangedLocal("TabletDescriptionBottomCurrent")
                End Using
            End Set
        End Property
        Public Property CurrentDescriptionBottomText As String
            Get
                Return (m_DataObject.CurrentDescriptionBottom + 1).ToString
            End Get
            Set(value As String)
                CurrentDescriptionBottom = CInt(Val(value)) - 1
            End Set
        End Property
        Public Property TabletDescriptionTopCurrent As String
            Get
                If m_DataObject.TabletDescriptionTop.Count > 0 Then
                    Return m_DataObject.TabletDescriptionTop(m_DataObject.CurrentDescriptionTop)
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    If m_DataObject.TabletDescriptionTop.Count > 0 Then
                        m_DataObject.TabletDescriptionTop(m_DataObject.CurrentDescriptionTop) = value
                    End If
                    OnPropertyChangedLocal("TabletDescriptionTopCurrent")
                End Using
            End Set
        End Property
        Public Property TabletDescriptionBottomCurrent As String
            Get
                If m_DataObject.TabletDescriptionBottom.Count > 0 Then
                    Return m_DataObject.TabletDescriptionBottom(m_DataObject.CurrentDescriptionBottom)
                Else
                    Return String.Empty
                End If
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    If m_DataObject.TabletDescriptionBottom.Count > 0 Then
                        m_DataObject.TabletDescriptionBottom(m_DataObject.CurrentDescriptionBottom) = value
                    End If
                    OnPropertyChangedLocal("TabletDescriptionBottomCurrent")
                End Using
            End Set
        End Property
        Public ReadOnly Property TabletDescriptionTop As List(Of String)
            Get
                Return m_DataObject.TabletDescriptionTop
            End Get
        End Property
        Public ReadOnly Property TabletDescriptionTopTuple As List(Of Tuple(Of Rect, String))
            Get
                Return (From sDescription In m_DataObject.TabletDescriptionTop Select New Tuple(Of Rect, String)(Rect.Empty, sDescription)).ToList
            End Get
        End Property
        Public ReadOnly Property TabletDescriptionBottom As List(Of String)
            Get
                Return m_DataObject.TabletDescriptionBottom
            End Get
        End Property
        Public ReadOnly Property TabletDescriptionBottomTuple As List(Of Tuple(Of Rect, String))
            Get
                Return (From sDescription In m_DataObject.TabletDescriptionBottom Select New Tuple(Of Rect, String)(Rect.Empty, sDescription)).ToList
            End Get
        End Property
        Public Overridable ReadOnly Property TabletLabelTop As String
            Get
                Return "Top"
            End Get
        End Property
        Public Overridable Property TopDescriptionContentToolTip As String
            Get
                Return "Top Description"
            End Get
            Set(value As String)
            End Set
        End Property
        Public Overridable ReadOnly Property TabletLabelBottom As String
            Get
                Return "Bottom"
            End Get
        End Property
        Public Overridable Property BottomDescriptionContentToolTip As String
            Get
                Return "Bottom Description"
            End Get
            Set(value As String)
            End Set
        End Property
        Public Property TabletSingleChoiceOnly As Boolean
            Get
                If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                    Return CType(Parent, FormBlock).TabletSingleChoiceOnly
                Else
                    Return False
                End If
            End Get
            Set(value As Boolean)
                If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                    CType(Parent, FormBlock).TabletSingleChoiceOnly = value
                    OnPropertyChangedLocal("TabletSingleChoiceOnly")
                End If
            End Set
        End Property
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "User Choice"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMCheck"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelSelection")
                oDisplayFilter.Add("StackPanelSingle")
                oDisplayFilter.Add("StackPanelInput")
                oDisplayFilter.Add("StackPanelFieldChoice")
                oDisplayFilter.Add("StackPanelFieldChoiceExtraTop")
                oDisplayFilter.Add("StackPanelFieldChoiceExtraBottom")
                oDisplayFilter.Add("StackPanelFieldChoiceMore")
                oDisplayFilter.Add("StackPanelAlignmentStart")
                oDisplayFilter.Add("StackPanelAlignmentBase")
                oDisplayFilter.Add("FieldChoiceTabletContent")
                oDisplayFilter.Add("FieldChoiceTabletCount")
                oDisplayFilter.Add("FieldChoiceTabletStart")
                oDisplayFilter.Add("FieldChoiceTabletGroups")
                oDisplayFilter.Add("FieldChoiceTopDescription")
                oDisplayFilter.Add("FieldChoiceTopDescriptionContent")
                oDisplayFilter.Add("FieldChoiceBottomDescription")
                oDisplayFilter.Add("FieldChoiceBottomDescriptionContent")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return False
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim oCheckEmpty As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty)
            Dim oCheckEmptyBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckEmpty, iBitmapWidth / 8, iBitmapHeight, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim XCheckEmptyBitmapSize As New XSize(CDbl(oCheckEmptyBitmap.Width) * 72 / RenderResolution300, CDbl(oCheckEmptyBitmap.Height) * 72 / RenderResolution300)
            Dim oXCheckEmptyImage As XImage = PdfSharp.Drawing.XImage.FromGdiPlusImage(oCheckEmptyBitmap)

            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
            Dim oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, oCheckEmptyBitmap.Width * 1.5, oCheckEmptyBitmap.Height * 1.5, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim XCheckMarkBitmapSize As New XSize(CDbl(oCheckMarkBitmap.Width) * 72 / RenderResolution300, CDbl(oCheckMarkBitmap.Height) * 72 / RenderResolution300)
            Dim oXCheckMarkImage As XImage = PdfSharp.Drawing.XImage.FromGdiPlusImage(oCheckMarkBitmap)

            ' determine displacements
            Dim fFirstDisplacement As Double = XCheckEmptyBitmapSize.Width / 2
            Dim fLastDisplacement As Double = oBackgroundSize.Width - fFirstDisplacement
            Dim fSpacing As Double = (fLastDisplacement - fFirstDisplacement) / 4
            Dim XDisplacementList As New List(Of Tuple(Of Char, XUnit))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("A", New XUnit(fFirstDisplacement)))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("B", New XUnit(fFirstDisplacement + fSpacing)))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("C", New XUnit(fFirstDisplacement + (fSpacing * 2))))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("D", New XUnit(fFirstDisplacement + (fSpacing * 3))))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("E", New XUnit(fLastDisplacement)))
            Dim YDisplacement As New XUnit((iBitmapHeight * 72 / RenderResolution300) / 2)

            ' draw ellipses
            For Each XDisplacement As Tuple(Of Char, XUnit) In XDisplacementList
                oXGraphics.DrawImage(oXCheckEmptyImage, New XUnit(XDisplacement.Item2.Point - fFirstDisplacement), New XUnit(YDisplacement.Point - (XCheckEmptyBitmapSize.Height / 2)), oXCheckEmptyImage.PointWidth, oXCheckEmptyImage.PointHeight)
            Next

            ' draw letters
            Const fFontSize As Double = 10
            Const fDescriptionFontScale As Double = 0.75
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * 0.8 * (oCheckEmptyBitmap.Height * 72 / RenderResolution300) / oTestFont.GetHeight
            Dim oArielFont As New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)
            Dim oArielFontDescription As New XFont(FontArial, fScaledFontSize * fDescriptionFontScale, XFontStyle.Regular, oFontOptions)

            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            For Each XDisplacement As Tuple(Of Char, XUnit) In XDisplacementList
                Dim oXSize As XSize = oXGraphics.MeasureString(XDisplacement.Item1, oArielFont)
                oXGraphics.DrawString(XDisplacement.Item1, oArielFont, XBrushes.Black, XDisplacement.Item2, YDisplacement, oStringFormat)
            Next

            ' draw descriptions
            Dim oDescriptionList As List(Of String) = PDFHelper.LoremIpsumWordList(0, XDisplacementList.Count * 2, True, 5)
            Dim fFontHeight As Double = oArielFontDescription.GetHeight * 72 / RenderResolution300
            For i = 0 To XDisplacementList.Count - 1
                Dim XDisplacement As Tuple(Of Char, XUnit) = XDisplacementList(i)
                oXGraphics.DrawString(oDescriptionList(i), oArielFontDescription, XBrushes.Black, XDisplacement.Item2, New XUnit(YDisplacement.Point - XCheckEmptyBitmapSize.Height), oStringFormat)
                oXGraphics.DrawString(oDescriptionList(XDisplacementList.Count + i), oArielFontDescription, XBrushes.Black, XDisplacement.Item2, New XUnit(YDisplacement.Point + XCheckEmptyBitmapSize.Height), oStringFormat)
            Next

            ' draw check mark
            oXGraphics.DrawImage(oXCheckMarkImage, New XUnit(XDisplacementList(3).Item2.Point - (oXCheckMarkImage.PointWidth / 2)), New XUnit(YDisplacement.Point - (oXCheckMarkImage.PointHeight / 2)), oXCheckMarkImage.PointWidth, oXCheckMarkImage.PointHeight)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
            Dim fAlignmentDisplacementX As Double = 0
            Dim fAlignmentDisplacementY As Double = 0
            Dim iTabletColumns As Integer = Math.Ceiling(TabletCount / TabletGroups)
            Select Case Alignment
                Case Enumerations.AlignmentEnum.Left
                    fAlignmentDisplacementY = (XImageHeight.Point - ((TabletGroups - 1) * fSingleBlockWidth * 3 / 2)) / 2
                Case Enumerations.AlignmentEnum.Right
                    fAlignmentDisplacementX = XImageWidth.Point - (iTabletColumns * fSingleBlockWidth)
                    fAlignmentDisplacementY = (XImageHeight.Point - ((TabletGroups - 1) * fSingleBlockWidth * 3 / 2)) / 2
                Case Enumerations.AlignmentEnum.Center
                    fAlignmentDisplacementX = (XImageWidth.Point - (iTabletColumns * fSingleBlockWidth)) / 2
                    fAlignmentDisplacementY = (XImageHeight.Point - ((TabletGroups - 1) * fSingleBlockWidth * 3 / 2)) / 2
            End Select

            ' process input fields if present
            Dim oField As FieldDocumentStore.Field = Nothing
            If ParamList.ContainsKey(ParamList.KeyFieldCollection) AndAlso Not Exclude Then
                oField = New FieldDocumentStore.Field
                With oField
                    .FieldType = Enumerations.FieldTypeEnum.Choice
                    .GUID = Guid.NewGuid
                    .Numbering = Numbering
                    .PageNumber = ParamList.Value(ParamList.KeyPDFPages).Count
                    .Critical = Critical
                    .TabletStart = TabletStart
                    .TabletGroups = TabletGroups
                    .TabletContent = TabletContent
                    .TabletDescriptionTop.AddRange(TabletDescriptionTopTuple)
                    .TabletDescriptionBottom.AddRange(TabletDescriptionBottomTuple)
                    .TabletAlignment = Alignment
                    .TabletSingleChoiceOnly = TabletSingleChoiceOnly
                End With
            End If

            Dim oTabletReturn As Rect = PDFHelper.DrawFieldTablets(oXGraphics, New XPoint(XFieldDisplacement.X + fAlignmentDisplacementX, XFieldDisplacement.Y + fAlignmentDisplacementY), TabletCount, TabletStart, TabletGroups, TabletContent, oField, Nothing, TabletDescriptionTopTuple, TabletDescriptionBottomTuple)

            If Not IsNothing(oField) Then
                oField.Location = oTabletReturn
                ParamList.Value(ParamList.KeyFieldCollection).Fields.Add(oField)
            End If
        End Sub
        Public Overrides Function FieldSpecificProcessing(ByVal oRectFormField As Rect) As Boolean
            ' field specific processing to check whether placement is allowed
            ' default value is false (no problem)
            ' minimum of 3 rows height
            If oRectFormField.Height < 3 Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelSingle.DataContext = Me
            Root.StackPanelInput.DataContext = Me
            Root.StackPanelFieldChoice.DataContext = Me
            Root.StackPanelFieldChoiceExtraTop.DataContext = Me
            Root.StackPanelFieldChoiceExtraBottom.DataContext = Me
            Root.StackPanelFieldChoiceMore.DataContext = Me
            Root.StackPanelAlignmentStart.DataContext = Me
            Root.StackPanelAlignmentBase.DataContext = Me

            Dim oBindingField1 As New Data.Binding
            oBindingField1.Path = New PropertyPath("Critical")
            oBindingField1.Mode = Data.BindingMode.TwoWay
            oBindingField1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldCritical.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField1)

            Dim oBindingField2 As New Data.Binding
            oBindingField2.Path = New PropertyPath("Exclude")
            oBindingField2.Mode = Data.BindingMode.TwoWay
            oBindingField2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldExclude.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField2)

            Dim oBindingField3 As New Data.Binding
            oBindingField3.Path = New PropertyPath("TabletSingleChoiceOnly")
            oBindingField3.Mode = Data.BindingMode.TwoWay
            oBindingField3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldSingle.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField3)

            Dim oBindingChoice1 As New Data.Binding
            oBindingChoice1.Path = New PropertyPath("TabletCountText")
            oBindingChoice1.Mode = Data.BindingMode.TwoWay
            oBindingChoice1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTabletCount.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice1)

            Dim oBindingChoice2 As New Data.Binding
            oBindingChoice2.Path = New PropertyPath("TabletStartText")
            oBindingChoice2.Mode = Data.BindingMode.TwoWay
            oBindingChoice2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTabletStart.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice2)

            Dim oBindingChoice3 As New Data.Binding
            oBindingChoice3.Path = New PropertyPath("TabletGroupsText")
            oBindingChoice3.Mode = Data.BindingMode.TwoWay
            oBindingChoice3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTabletGroups.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice3)

            Dim oBindingChoice4 As New Data.Binding
            oBindingChoice4.Path = New PropertyPath("CurrentDescriptionTopText")
            oBindingChoice4.Mode = Data.BindingMode.TwoWay
            oBindingChoice4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTopDescription.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice4)

            Dim oBindingChoice5 As New Data.Binding
            oBindingChoice5.Path = New PropertyPath("TopDescriptionContentToolTip")
            oBindingChoice5.Mode = Data.BindingMode.TwoWay
            oBindingChoice5.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTopDescriptionContent.SetBinding(Common.HighlightTextBox.HTBInnerToolTipProperty, oBindingChoice5)

            Dim oBindingChoice6 As New Data.Binding
            oBindingChoice6.Path = New PropertyPath("CurrentDescriptionBottomText")
            oBindingChoice6.Mode = Data.BindingMode.TwoWay
            oBindingChoice6.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceBottomDescription.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice6)

            Dim oBindingChoice7 As New Data.Binding
            oBindingChoice7.Path = New PropertyPath("BottomDescriptionContentToolTip")
            oBindingChoice7.Mode = Data.BindingMode.TwoWay
            oBindingChoice7.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceBottomDescriptionContent.SetBinding(Common.HighlightTextBox.HTBInnerToolTipProperty, oBindingChoice7)

            Dim oBindingChoice8 As New Data.Binding
            oBindingChoice8.Path = New PropertyPath("TabletDescriptionTopCurrent")
            oBindingChoice8.Mode = Data.BindingMode.TwoWay
            oBindingChoice8.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTopDescriptionContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice8)

            Dim oBindingChoice9 As New Data.Binding
            oBindingChoice9.Path = New PropertyPath("TabletDescriptionBottomCurrent")
            oBindingChoice9.Mode = Data.BindingMode.TwoWay
            oBindingChoice9.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceBottomDescriptionContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice9)

            Dim oBindingChoice10 As New Data.Binding
            oBindingChoice10.Path = New PropertyPath("TabletContentText")
            oBindingChoice10.Mode = Data.BindingMode.TwoWay
            oBindingChoice10.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTabletContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingChoice10)

            Dim oBindingChoice11 As New Data.Binding
            oBindingChoice11.Path = New PropertyPath("TabletLabelTop")
            oBindingChoice11.Mode = Data.BindingMode.OneWay
            oBindingChoice11.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTextTopDescription.SetBinding(Controls.TextBlock.TextProperty, oBindingChoice11)

            Dim oBindingChoice12 As New Data.Binding
            oBindingChoice12.Path = New PropertyPath("TabletLabelBottom")
            oBindingChoice12.Mode = Data.BindingMode.OneWay
            oBindingChoice12.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldChoiceTextBottomDescription.SetBinding(Controls.TextBlock.TextProperty, oBindingChoice12)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "30a1b20b-eb5e-4523-a40e-b1084049f4d0"
            End Get
        End Property
        Public Sub SetTabletDescriptionTop(ByVal oDescriptions As List(Of String))
            m_DataObject.TabletDescriptionTop.Clear()
            m_DataObject.TabletDescriptionTop.AddRange(oDescriptions)
            OnPropertyChangedLocal("CurrentDescriptionTop")
            OnPropertyChangedLocal("CurrentDescriptionTopText")
            OnPropertyChangedLocal("TabletDescriptionTopCurrent")
            OnPropertyChangedLocal("TabletDescriptionTop")
        End Sub
        Public Sub SetTabletDescriptionBottom(ByVal oDescriptions As List(Of String))
            m_DataObject.TabletDescriptionBottom.Clear()
            m_DataObject.TabletDescriptionBottom.AddRange(oDescriptions)
            OnPropertyChangedLocal("CurrentDescriptionBottom")
            OnPropertyChangedLocal("CurrentDescriptionBottomText")
            OnPropertyChangedLocal("TabletDescriptionBottomCurrent")
            OnPropertyChangedLocal("TabletDescriptionBottom")
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FieldChoiceVertical
        Inherits FieldChoice

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldChoiceVertical).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Title = "Vertical" + vbCr + "Tablet Count: " + TabletCountText + If(Critical, vbCrLf + "Critical Field", String.Empty)
            Title += If(Exclude, vbCr + "Excluded", String.Empty)
        End Sub
#End Region
#Region "Choice"
        Protected Overrides Function GetMaxTabletCount() As Integer
            Dim oBlock As FormBlock = Parent
            If IsNothing(oBlock) Then
                Return 0
            Else
                ' each tablet height is equal to two row heights
                Return Math.Floor(GridRect.Height / 2) * TabletGroups
            End If
        End Function
        Public Overrides ReadOnly Property TabletLabelTop As String
            Get
                Return "Left"
            End Get
        End Property
        Public Overrides Property TopDescriptionContentToolTip As String
            Get
                Return "Left Description"
            End Get
            Set(value As String)
            End Set
        End Property
        Public Overrides ReadOnly Property TabletLabelBottom As String
            Get
                Return "Right"
            End Get
        End Property
        Public Overrides Property BottomDescriptionContentToolTip As String
            Get
                Return "Right Description"
            End Get
            Set(value As String)
            End Set
        End Property
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "User Choice Vertical"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "90f02a2a-fc31-4e06-ad4b-54f651784124"
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim oCheckEmpty As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty)
            Dim oCheckEmptyBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckEmpty, iBitmapWidth, iBitmapHeight / 4, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim XCheckEmptyBitmapSize As New XSize(CDbl(oCheckEmptyBitmap.Width) * 72 / RenderResolution300, CDbl(oCheckEmptyBitmap.Height) * 72 / RenderResolution300)
            Dim oXCheckEmptyImage As XImage = PdfSharp.Drawing.XImage.FromGdiPlusImage(oCheckEmptyBitmap)

            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
            Dim oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, oCheckEmptyBitmap.Width * 1.5, oCheckEmptyBitmap.Height * 1.5, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim XCheckMarkBitmapSize As New XSize(CDbl(oCheckMarkBitmap.Width) * 72 / RenderResolution300, CDbl(oCheckMarkBitmap.Height) * 72 / RenderResolution300)
            Dim oXCheckMarkImage As XImage = PdfSharp.Drawing.XImage.FromGdiPlusImage(oCheckMarkBitmap)

            ' determine displacements
            Dim XDisplacementList As New List(Of Tuple(Of Char, XUnit))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("A", New XUnit((oBackgroundSize.Height / 2) - (XCheckEmptyBitmapSize.Height * 1.25))))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("B", New XUnit(oBackgroundSize.Height / 2)))
            XDisplacementList.Add(New Tuple(Of Char, XUnit)("C", New XUnit((oBackgroundSize.Height / 2) + (XCheckEmptyBitmapSize.Height * 1.25))))
            Dim XDisplacement As New XUnit((iBitmapWidth * 72 / RenderResolution300) / 2)

            ' draw ellipses
            For Each YDisplacement As Tuple(Of Char, XUnit) In XDisplacementList
                oXGraphics.DrawImage(oXCheckEmptyImage, New XUnit(XDisplacement.Point - (XCheckEmptyBitmapSize.Width / 2)), New XUnit(YDisplacement.Item2.Point - (XCheckEmptyBitmapSize.Height / 2)), oXCheckEmptyImage.PointWidth, oXCheckEmptyImage.PointHeight)
            Next

            ' draw letters
            Const fFontSize As Double = 10
            Const fDescriptionFontScale As Double = 0.75
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * 0.8 * (oCheckEmptyBitmap.Height * 72 / RenderResolution300) / oTestFont.GetHeight
            Dim oArielFont As New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)
            Dim oArielFontDescription As New XFont(FontArial, fScaledFontSize * fDescriptionFontScale, XFontStyle.Regular, oFontOptions)

            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            For Each YDisplacement As Tuple(Of Char, XUnit) In XDisplacementList
                oXGraphics.DrawString(YDisplacement.Item1, oArielFont, XBrushes.Black, XDisplacement, YDisplacement.Item2, oStringFormat)
            Next

            Dim oStringFormatLeft As New XStringFormat()
            oStringFormatLeft.Alignment = XStringAlignment.Near
            oStringFormatLeft.LineAlignment = XLineAlignment.Center

            Dim oStringFormatRight As New XStringFormat()
            oStringFormatRight.Alignment = XStringAlignment.Far
            oStringFormatRight.LineAlignment = XLineAlignment.Center

            ' draw descriptions
            Dim oDescriptionList As New List(Of String)
            For i = 0 To (XDisplacementList.Count * 2) - 1
                oDescriptionList.Add(PDFHelper.LoremIpsumPhrase(i * 2, 2, True))
            Next

            Dim fFontHeight As Double = oArielFontDescription.GetHeight * 72 / RenderResolution300
            For i = 0 To XDisplacementList.Count - 1
                Dim YDisplacement As Tuple(Of Char, XUnit) = XDisplacementList(i)
                oXGraphics.DrawString(oDescriptionList(i), oArielFontDescription, XBrushes.Black, New XRect(0, YDisplacement.Item2.Point - fFontHeight / 2, (oBackgroundSize.Width / 2) - XCheckEmptyBitmapSize.Width, fFontHeight), oStringFormatRight)
                oXGraphics.DrawString(oDescriptionList(XDisplacementList.Count + i), oArielFontDescription, XBrushes.Black, New XRect((oBackgroundSize.Width / 2) + XCheckEmptyBitmapSize.Width, YDisplacement.Item2.Point - fFontHeight / 2, (oBackgroundSize.Width / 2) - XCheckEmptyBitmapSize.Width, fFontHeight), oStringFormatLeft)
            Next

            ' draw check mark
            oXGraphics.DrawImage(oXCheckMarkImage, New XUnit(XDisplacement.Point - (oXCheckMarkImage.PointWidth / 2)), New XUnit(XDisplacementList(2).Item2.Point - (oXCheckMarkImage.PointHeight / 2)), oXCheckMarkImage.PointWidth, oXCheckMarkImage.PointHeight)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
            Dim oXRect As New XRect(XFieldDisplacement.X, XFieldDisplacement.Y, XImageWidth.Point, XImageHeight.Point)
            oXGraphics.IntersectClip(oXRect)

            ' process input fields if present
            Dim oField As FieldDocumentStore.Field = Nothing
            If ParamList.ContainsKey(ParamList.KeyFieldCollection) AndAlso Not Exclude Then
                oField = New FieldDocumentStore.Field
                With oField
                    .FieldType = Enumerations.FieldTypeEnum.ChoiceVertical
                    .GUID = Guid.NewGuid
                    .Numbering = Numbering
                    .PageNumber = ParamList.Value(ParamList.KeyPDFPages).Count
                    .Critical = Critical
                    .TabletStart = TabletStart
                    .TabletGroups = TabletGroups
                    .TabletContent = TabletContent
                    .TabletDescriptionTop.AddRange(TabletDescriptionTopTuple)
                    .TabletDescriptionBottom.AddRange(TabletDescriptionBottomTuple)
                    .TabletAlignment = Alignment
                    .TabletSingleChoiceOnly = TabletSingleChoiceOnly
                End With
            End If

            Dim iTabletRows As Integer = Math.Ceiling(TabletCount / TabletGroups)
            Dim fAlignmentDisplacementY As Double = (XImageHeight.Point - (iTabletRows * fSingleBlockWidth)) / 2
            Dim oTabletReturn As Tuple(Of Rect, Double) = Nothing
            Select Case Alignment
                Case Enumerations.AlignmentEnum.Left
                    oTabletReturn = PDFHelper.DrawFieldTabletsVertical(oXGraphics, XImageWidth, XUnit.Zero, New XPoint(XFieldDisplacement.X, XFieldDisplacement.Y + fAlignmentDisplacementY), TabletCount, TabletStart, TabletGroups, TabletContent, oField, Alignment, Nothing, Nothing, TabletDescriptionBottomTuple)
                Case Enumerations.AlignmentEnum.Center
                    oTabletReturn = PDFHelper.DrawFieldTabletsVertical(oXGraphics, XImageWidth, XUnit.Zero, New XPoint(XFieldDisplacement.X, XFieldDisplacement.Y + fAlignmentDisplacementY), TabletCount, TabletStart, TabletGroups, TabletContent, oField, Alignment, Nothing, TabletDescriptionTopTuple, TabletDescriptionBottomTuple)
                Case Enumerations.AlignmentEnum.Right
                    oTabletReturn = PDFHelper.DrawFieldTabletsVertical(oXGraphics, XImageWidth, XUnit.Zero, New XPoint(XFieldDisplacement.X, XFieldDisplacement.Y + fAlignmentDisplacementY), TabletCount, TabletStart, TabletGroups, TabletContent, oField, Alignment, Nothing, TabletDescriptionTopTuple, Nothing)
            End Select

            If (Not IsNothing(oField)) And (Not IsNothing(oTabletReturn)) Then
                oField.Location = oTabletReturn.Item1
                oField.TabletLimit = oTabletReturn.Item2
                ParamList.Value(ParamList.KeyFieldCollection).Fields.Add(oField)
            End If
        End Sub
        Public Overrides Function FieldSpecificProcessing(ByVal oRectFormField As Rect) As Boolean
            ' field specific processing to check whether placement is allowed
            ' default value is false (no problem)
            ' minimum of 2 rows height per tablet
            If (oRectFormField.Width < TabletGroups) OrElse oRectFormField.Height < Math.Floor(Math.Max(TabletCount, 1) / TabletGroups) * 2 Then
                Return True
            Else
                Return False
            End If
        End Function
    End Class
    <DataContract(IsReference:=True)> Public Class FieldChoiceVerticalMCQ
        Inherits FieldChoice

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldChoiceVerticalMCQ).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Choice"
        Public ReadOnly Property TabletDescriptionMCQ As List(Of List(Of ElementStruc))
            Get
                Return m_DataObject.TabletDescriptionMCQ
            End Get
        End Property
        Public ReadOnly Property TabletDescriptionMCQTuple As List(Of Tuple(Of Rect, Integer, Integer, List(Of ElementStruc)))
            Get
                Return (From oDescription In m_DataObject.TabletDescriptionMCQ Select New Tuple(Of Rect, Integer, Integer, List(Of ElementStruc))(Rect.Empty, -1, -1, oDescription)).ToList
            End Get
        End Property
        Public Sub SetTabletDescriptionMCQ(ByVal oDescriptions As List(Of List(Of ElementStruc)))
            m_DataObject.TabletDescriptionMCQ.Clear()
            m_DataObject.TabletDescriptionMCQ.AddRange(oDescriptions)
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "User Choice Vertical MCQ"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return ""
            End Get
        End Property
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Const FontSizeMultiplier As Integer = 1
            If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormMCQ)) Then
                Dim oMCQ As FormMCQ = Parent
                Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
                Dim oXRect As New XRect(XFieldDisplacement.X, XFieldDisplacement.Y, XImageWidth.Point, XImageHeight.Point)
                oXGraphics.IntersectClip(oXRect)

                ' process input fields if present
                Dim oField As New FieldDocumentStore.Field
                With oField
                    .FieldType = Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                    .GUID = Guid.NewGuid
                    .Numbering = Numbering
                    .PageNumber = If(ParamList.ContainsKey(ParamList.KeyPDFPages), ParamList.Value(ParamList.KeyPDFPages).Count, -1)
                    .TabletStart = TabletStart
                    .Critical = Critical
                    .TabletGroups = oMCQ.BlockWidth
                    .TabletContent = TabletContent
                    .TabletDescriptionMCQ.AddRange(TabletDescriptionMCQTuple)
                    .TabletAlignment = Alignment
                    .TabletSingleChoiceOnly = TabletSingleChoiceOnly
                    .TabletMCQParams = New Tuple(Of Double, Double, Point, Int32Rect, Integer, Integer, List(Of Integer))(XImageWidth.Point, XImageHeight.Point, New Point(XFieldDisplacement.X, XFieldDisplacement.Y), GridRect, oMCQ.GridHeight, FontSizeMultiplier, oOverflowRows)
                End With

                Dim oTabletReturn As Rect = PDFHelper.DrawFieldTabletsMCQ(oXGraphics, XImageWidth, XImageHeight, XFieldDisplacement, GridRect, oMCQ.GridHeight, FontSizeMultiplier, oOverflowRows, oField)

                If ParamList.ContainsKey(ParamList.KeyFieldCollection) AndAlso (Not Exclude) AndAlso (Not IsNothing(oTabletReturn)) Then
                    oField.Location = oTabletReturn
                    ParamList.Value(ParamList.KeyFieldCollection).Fields.Add(oField)
                End If
            End If
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FieldBoxChoice
        Inherits FieldChoice

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldBoxChoice).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Private m_CurrentRow As Integer
        Public Property BlockCount As Integer
            Get
                Return m_DataObject.BlockCount
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                        Dim iMaxBlockCount As Integer = GridRect.Height / 2
                        If value < 0 Then
                            m_DataObject.BlockCount = 0
                        ElseIf value > iMaxBlockCount Then
                            m_DataObject.BlockCount = iMaxBlockCount
                        Else
                            m_DataObject.BlockCount = value
                        End If
                    Else
                        m_DataObject.BlockCount = 0
                    End If

                    ResetHandwritingContent()
                    OnPropertyChangedLocal("BlockCount")
                    OnPropertyChangedLocal("BlockCountText")
                End Using
            End Set
        End Property
        Public Property BlockCountText As String
            Get
                Return m_DataObject.BlockCount.ToString
            End Get
            Set(value As String)
                BlockCount = CInt(Val(value))
            End Set
        End Property
        Public Property CurrentRow As Integer
            Get
                Return m_CurrentRow
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 0 Then
                        m_CurrentRow = 0
                    ElseIf value > m_DataObject.BlockCount - 1 Then
                        m_CurrentRow = m_DataObject.BlockCount - 1
                    Else
                        m_CurrentRow = value
                    End If

                    OnPropertyChangedLocal("CurrentRow")
                    OnPropertyChangedLocal("CurrentRowText")
                    OnPropertyChangedLocal("CurrentContent")
                End Using
            End Set
        End Property
        Public Property CurrentRowText As String
            Get
                Return (m_CurrentRow + 1).ToString
            End Get
            Set(value As String)
            End Set
        End Property
        Public Property CurrentContent As String
            Get
                Return m_DataObject.HandwritingContent(m_CurrentRow)
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.HandwritingContent(m_CurrentRow) = Left(Trim(value), 1).Replace("-", "–")
                    CurrentDescriptionTop = CurrentDescriptionTop
                    OnPropertyChangedLocal("CurrentContent")
                End Using
            End Set
        End Property
        Public Shadows Property CurrentDescriptionTop As Integer
            Get
                Return m_DataObject.CurrentDescriptionTop
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If m_DataObject.TabletDescriptionTop.Count = 0 Then
                        m_DataObject.CurrentDescriptionTop = -1
                    Else
                        Dim iMaxCount As Integer = BoxChoiceGroupCount()
                        If value < 0 Then
                            m_DataObject.CurrentDescriptionTop = 0
                        ElseIf value > iMaxCount - 1 Then
                            m_DataObject.CurrentDescriptionTop = iMaxCount - 1
                        Else
                            m_DataObject.CurrentDescriptionTop = value
                        End If
                    End If

                    OnPropertyChangedLocal("CurrentDescriptionTop")
                    OnPropertyChangedLocal("CurrentDescriptionTopText")
                    OnPropertyChangedLocal("TabletDescriptionTopCurrent")
                End Using
            End Set
        End Property
        Public ReadOnly Property HandwritingContent(ByVal iRow As Integer) As String
            Get
                If iRow >= 0 And iRow <= m_DataObject.BlockCount - 1 Then
                    Return m_DataObject.HandwritingContent(iRow)
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Private Sub ResetHandwritingContent()
            ' reset handwritingcontent
            Dim oNewHandwritingContent(Math.Max(m_DataObject.BlockCount - 1, 0)) As String
            For x = 0 To oNewHandwritingContent.Length - 1
                oNewHandwritingContent(x) = String.Empty
            Next
            For x = 0 To Math.Min(m_DataObject.HandwritingContent.Length, oNewHandwritingContent.Length) - 1
                oNewHandwritingContent(x) = If(IsNothing(m_DataObject.HandwritingContent(x)), String.Empty, m_DataObject.HandwritingContent(x))
            Next
            m_DataObject.HandwritingContent = oNewHandwritingContent

            ' add or remove tablets to match tablet count
            ' if a tablet descriptor is not empty, then do not remove
            If m_DataObject.TabletDescriptionTop.Count > m_DataObject.BlockCount Then
                For i = m_DataObject.TabletDescriptionTop.Count - 1 To m_DataObject.BlockCount Step -1
                    If Trim(m_DataObject.TabletDescriptionTop(i)) = String.Empty Then
                        m_DataObject.TabletDescriptionTop.RemoveAt(i)
                    Else
                        Exit For
                    End If
                Next
            ElseIf m_DataObject.TabletDescriptionTop.Count < m_DataObject.BlockCount Then
                For i = 0 To m_DataObject.BlockCount - m_DataObject.TabletDescriptionTop.Count - 1
                    m_DataObject.TabletDescriptionTop.Add(String.Empty)
                Next
            End If

            If m_DataObject.TabletDescriptionTop.Count = 0 Then
                m_DataObject.CurrentDescriptionTop = -1
            Else
                m_DataObject.CurrentDescriptionTop = Math.Min(Math.Max(0, m_DataObject.CurrentDescriptionTop), m_DataObject.BlockCount - 1)
            End If

            CurrentDescriptionTop = CurrentDescriptionTop
        End Sub
        Private Function BoxChoiceGroupCount() As Integer
            If BlockCount > 0 Then
                Dim iGroupCount As Integer = 1
                Dim bTextFound As Boolean = False
                For i = 0 To BlockCount - 1
                    If HandwritingContent(i) <> String.Empty Then
                        bTextFound = True
                    End If
                    If HandwritingContent(i) = String.Empty And bTextFound Then
                        bTextFound = False
                        iGroupCount += 1
                    End If
                Next
                Return iGroupCount
            Else
                Return 0
            End If
        End Function
        Public Overrides Sub TitleChanged()
            Title = "Box Count: " + BlockCountText + If(Critical, vbCrLf + "Critical Field", String.Empty)
            Title += If(Exclude, vbCr + "Excluded", String.Empty)
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            With m_DataObject
                .BlockCount = 0
                ReDim .HandwritingContent(0)
                .HandwritingContent(0) = String.Empty
                .TabletStart = -2
                .TabletCount = 10
                .TabletContent = Enumerations.TabletContentEnum.Number
            End With
            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Box Choice"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMBoxCheck"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                If Not IsNothing(Parent) Then
                    oDisplayFilter.AddRange(Parent.DisplayFilter)
                End If
                oDisplayFilter.Add("StackPanelSelection")
                oDisplayFilter.Add("StackPanelInput")
                oDisplayFilter.Add("StackPanelAlignmentStart")
                oDisplayFilter.Add("StackPanelAlignmentExtra1")
                oDisplayFilter.Add("StackPanelAlignmentExtra2")
                oDisplayFilter.Add("StackPanelAlignmentBase")
                oDisplayFilter.Add("StackPanelFieldBoxChoice")
                oDisplayFilter.Add("StackPanelFieldHandwriting")
                oDisplayFilter.Add("StackPanelFieldHandwritingExtra2")
                oDisplayFilter.Add("FieldBoxChoiceLabel")
                oDisplayFilter.Add("FieldBoxChoiceLabelContent")
                oDisplayFilter.Add("FieldHandwritingBlockCount")
                oDisplayFilter.Add("FieldHandwritingCurrentRow")
                oDisplayFilter.Add("FieldHandwritingCurrentContent")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            ' draw boxes
            Const fTabletHeight As Double = 0.5
            Dim sWord As String = "376"
            Dim fDimension As Double = oBackgroundSize.Height * 0.96 / sWord.Count
            Dim oCheckEmpty As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty)
            Dim oCheckEmptyBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckEmpty, Double.MaxValue, fDimension * fTabletHeight * RenderResolution300 / 72, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim XCheckEmptyBitmapSize As New XSize(CDbl(oCheckEmptyBitmap.Width) * 72 / RenderResolution300, CDbl(oCheckEmptyBitmap.Height) * 72 / RenderResolution300)
            Dim oXCheckEmptyImage As XImage = XImage.FromGdiPlusImage(oCheckEmptyBitmap)

            ' determine displacements
            Dim fFirstDisplacementBox As Double = oBackgroundSize.Height * 0.02 + fDimension / 2
            Dim fLastDisplacementBox As Double = oBackgroundSize.Height - fFirstDisplacementBox
            Dim fBackgroundDisplacementX As Double = (oBackgroundSize.Width - ((fFirstDisplacementBox * 1.75) + (fDimension * 5 / 8) + (9 * fDimension * 0.65) + (XCheckEmptyBitmapSize.Width / 2))) / 2
            Dim fSpacingBox As Double = (fLastDisplacementBox - fFirstDisplacementBox) / (sWord.Count - 1)
            Dim YDisplacementListBox As New List(Of Tuple(Of Char, XUnit))
            For i = 0 To sWord.Count - 1
                YDisplacementListBox.Add(New Tuple(Of Char, XUnit)(Mid(sWord, i + 1, 1), New XUnit(fFirstDisplacementBox + (fSpacingBox * i))))
            Next
            Dim XDisplacementBox As New XUnit(fDimension / 2)

            ' draw rectangles
            For Each YDisplacement As Tuple(Of Char, XUnit) In YDisplacementListBox
                PDFHelper.DrawFieldBorder(oXGraphics, New XUnit(fDimension), New XUnit(fDimension), New XPoint(fBackgroundDisplacementX + XDisplacementBox.Point - fDimension / 2, YDisplacement.Item2.Point - fDimension / 2), 0.5)
            Next

            ' draw letters
            Const fFontSize As Double = 10
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontComicSansMS, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * 0.8 * fDimension / oTestFont.GetHeight
            Dim oComicSansFont As New XFont(FontComicSansMS, fScaledFontSize, XFontStyle.Regular, oFontOptions)

            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            For Each YDisplacement As Tuple(Of Char, XUnit) In YDisplacementListBox
                Dim oXSize As XSize = oXGraphics.MeasureString(YDisplacement.Item1, oComicSansFont)
                oXGraphics.DrawString(YDisplacement.Item1, oComicSansFont, XBrushes.Black, New XUnit(fBackgroundDisplacementX + XDisplacementBox.Point), YDisplacement.Item2, oStringFormat)
            Next

            ' draw tablets
            Dim oCheckMark As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckMark)
            Dim oCheckMarkBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckMark, oCheckEmptyBitmap.Width * 1.5, oCheckEmptyBitmap.Height * 1.5, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim XCheckMarkBitmapSize As New XSize(CDbl(oCheckMarkBitmap.Width) * 72 / RenderResolution300, CDbl(oCheckMarkBitmap.Height) * 72 / RenderResolution300)
            Dim oXCheckMarkImage As XImage = XImage.FromGdiPlusImage(oCheckMarkBitmap)

            Dim oTestFontTablet As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSizeTablet As Double = fFontSize * 0.8 * (oCheckEmptyBitmap.Height * 72 / RenderResolution300) / oTestFont.GetHeight
            Dim oArielFontTablet As New XFont(FontArial, fScaledFontSizeTablet, XFontStyle.Regular, oFontOptions)
            Dim fDisplacementTabletX As Double = fFirstDisplacementBox * 1.75
            Dim fDisplacementTabletY = fDimension * 1 / 16
            Dim oCenterPointList As New List(Of List(Of XPoint))
            For y = 0 To 2
                oCenterPointList.Add(New List(Of XPoint))
                For x = 0 To 9
                    Dim XCenterPoint As New XPoint(fBackgroundDisplacementX + fDisplacementTabletX + (fDimension * 5 / 8) + (x * fDimension * 0.65), (y * fDimension) + fDisplacementTabletY + (((((x Mod 2) * 2) + 1) * fDimension / 4) - ((((x Mod 2) * 2) - 1) * fDimension / 16)))
                    DrawTabletNoBackground(oXGraphics, XCenterPoint, x.ToString, oArielFontTablet, oStringFormat, fDimension, RenderResolution300)
                    oCenterPointList(y).Add(XCenterPoint)
                Next
            Next

            ' draw check mark
            oXGraphics.DrawImage(oXCheckMarkImage, New XUnit(oCenterPointList(0)(3).X - (oXCheckMarkImage.PointWidth / 2)), New XUnit(oCenterPointList(0)(3).Y - (oXCheckMarkImage.PointHeight / 2)), oXCheckMarkImage.PointWidth, oXCheckMarkImage.PointHeight)
            oXGraphics.DrawImage(oXCheckMarkImage, New XUnit(oCenterPointList(1)(7).X - (oXCheckMarkImage.PointWidth / 2)), New XUnit(oCenterPointList(1)(7).Y - (oXCheckMarkImage.PointHeight / 2)), oXCheckMarkImage.PointWidth, oXCheckMarkImage.PointHeight)
            oXGraphics.DrawImage(oXCheckMarkImage, New XUnit(oCenterPointList(2)(6).X - (oXCheckMarkImage.PointWidth / 2)), New XUnit(oCenterPointList(2)(6).Y - (oXCheckMarkImage.PointHeight / 2)), oXCheckMarkImage.PointWidth, oXCheckMarkImage.PointHeight)
        End Sub
        Private Shared Function DrawTabletNoBackground(ByRef oXGraphics As XGraphics, ByVal XCenterPoint As XPoint, ByVal sTabletContent As String, ByVal oFont As XFont, ByVal oStringFormat As XStringFormat, ByVal fSingleBlockWidth As Double, ByVal iResolution As Integer) As XSize
            ' draws a single tablet with the center point given
            ' the content is limited to two characters only
            ' draws tablet
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
            Const fTabletHeight As Double = 0.5
            Dim oCheckEmpty As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMCheckEmpty)
            Dim oCheckEmptyBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oCheckEmpty, Double.MaxValue, fSingleBlockWidth * fTabletHeight * iResolution / 72, Enumerations.StretchEnum.Uniform), iResolution)
            Dim oXCheckEmptyImage As XImage = XImage.FromGdiPlusImage(oCheckEmptyBitmap)
            Dim XCheckEmptyBitmapSize As New XSize(oXCheckEmptyImage.PointWidth, oXCheckEmptyImage.PointHeight)

            oXGraphics.DrawImage(oXCheckEmptyImage, XCenterPoint.X - XCheckEmptyBitmapSize.Width / 2, XCenterPoint.Y - XCheckEmptyBitmapSize.Height / 2)

            ' draws tablet text
            oXGraphics.DrawString(Left(sTabletContent, 2), oFont, XBrushes.Black, XCenterPoint.X, XCenterPoint.Y, oStringFormat)

            ' clean up
            oCheckEmptyBitmap.Dispose()

            oXGraphics.SmoothingMode = oSmoothingMode
            Return XCheckEmptyBitmapSize
        End Function
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
            Dim fLabelHeight As Double = fSingleBlockWidth / 2
            Dim fBoxChoiceWidth As Double = fLabelHeight + (fSingleBlockWidth * 3 / 2) + (10 * fSingleBlockWidth)

            Dim fAlignmentDisplacementX As Double = 0
            Dim fAlignmentDisplacementY As Double = 0
            Select Case Alignment
                Case Enumerations.AlignmentEnum.Left
                    fAlignmentDisplacementX = fLabelHeight
                    fAlignmentDisplacementY = (XImageHeight.Point - (BlockCount * fSingleBlockWidth)) / 2
                Case Enumerations.AlignmentEnum.UpperLeft
                    fAlignmentDisplacementX = fLabelHeight
                Case Enumerations.AlignmentEnum.Top
                    fAlignmentDisplacementX = fLabelHeight + (XImageWidth.Point - fBoxChoiceWidth) / 2
                Case Enumerations.AlignmentEnum.UpperRight
                    fAlignmentDisplacementX = fLabelHeight + XImageWidth.Point - fBoxChoiceWidth
                Case Enumerations.AlignmentEnum.Right
                    fAlignmentDisplacementX = fLabelHeight + XImageWidth.Point - fBoxChoiceWidth
                    fAlignmentDisplacementY = (XImageHeight.Point - (BlockCount * fSingleBlockWidth)) / 2
                Case Enumerations.AlignmentEnum.LowerRight
                    fAlignmentDisplacementX = fLabelHeight + XImageWidth.Point - fBoxChoiceWidth
                    fAlignmentDisplacementY = XImageHeight.Point - (BlockCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.Bottom
                    fAlignmentDisplacementX = fLabelHeight + (XImageWidth.Point - fBoxChoiceWidth) / 2
                    fAlignmentDisplacementY = XImageHeight.Point - (BlockCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.LowerLeft
                    fAlignmentDisplacementX = fLabelHeight
                    fAlignmentDisplacementY = XImageHeight.Point - (BlockCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.Center
                    fAlignmentDisplacementX = fLabelHeight + (XImageWidth.Point - fBoxChoiceWidth) / 2
                    fAlignmentDisplacementY = (XImageHeight.Point - (BlockCount * fSingleBlockWidth)) / 2
            End Select

            Dim oField As FieldDocumentStore.Field = Nothing
            If ParamList.ContainsKey(ParamList.KeyFieldCollection) AndAlso Not Exclude Then
                oField = New FieldDocumentStore.Field
                With oField
                    .FieldType = Enumerations.FieldTypeEnum.BoxChoice
                    .GUID = Guid.NewGuid
                    .Numbering = Numbering
                    .PageNumber = ParamList.Value(ParamList.KeyPDFPages).Count
                    .Critical = Critical
                    .TabletStart = TabletStart
                    .TabletContent = TabletContent
                    .TabletAlignment = Alignment
                    .TabletSingleChoiceOnly = True
                    .Location = Rect.Empty
                End With
            End If

            Dim XWidth As New XUnit(fSingleBlockWidth)
            Dim XHeight As New XUnit(fSingleBlockWidth)
            Dim XLeft As XUnit = Nothing
            Dim XTop As XUnit = Nothing

            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            Dim fFontSize As Double = PDFHelper.GetScaledFontSize(XWidth, XHeight, XHeight)
            Dim oArielFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fLabelFontSize As Double = PDFHelper.GetScaledFontSize(XImageHeight, fLabelHeight, fLabelHeight)
            Dim oArielFontLabel As New XFont(FontArial, fLabelFontSize, XFontStyle.Italic, oFontOptions)

            ' add block and label location
            If Not IsNothing(oField) Then
                oField.Location = New Rect(XFieldDisplacement.X + fAlignmentDisplacementX, XFieldDisplacement.Y + fAlignmentDisplacementY, fSingleBlockWidth, fSingleBlockWidth * BlockCount)
                oField.Location = Rect.Union(oField.Location, New Rect(XFieldDisplacement.X + fAlignmentDisplacementX - fLabelHeight, XFieldDisplacement.Y + fAlignmentDisplacementY, fLabelHeight, fSingleBlockWidth * BlockCount))
            End If

            For y = 0 To BlockCount - 1
                XLeft = New XUnit(XFieldDisplacement.X + fAlignmentDisplacementX)
                XTop = New XUnit(XFieldDisplacement.Y + fAlignmentDisplacementY + (y * fSingleBlockWidth))

                ' draw text if present
                If HandwritingContent(y) = String.Empty Then
                    If bRender Then
                        PDFHelper.DrawFieldBorder(oXGraphics, XWidth, XHeight, New XPoint(XLeft, XTop), 0.5)
                    Else
                        If CurrentRow <> y Then
                            PDFHelper.DrawFieldBorder(oXGraphics, XWidth, XHeight, New XPoint(XLeft, XTop), 0.5)
                        End If
                    End If

                    ' render tablets
                    Dim oLocation As Rect = PDFHelper.DrawFieldTablets(oXGraphics, New XPoint(XFieldDisplacement.X + fAlignmentDisplacementX + (fSingleBlockWidth * 3 / 2), XFieldDisplacement.Y + fAlignmentDisplacementY + (y * fSingleBlockWidth)), TabletCount, TabletStart, 1, TabletContent, oField, Nothing,,, y)
                    If Not IsNothing(oField) Then
                        oField.Location = Rect.Union(oField.Location, oLocation)
                    End If
                Else
                    PDFHelper.DrawStringRotated(oXGraphics, New XPoint(XLeft.Point + XWidth.Point / 2, XTop.Point + XHeight.Point / 2), oArielFont, oStringFormat, HandwritingContent(y), -90)
                End If

                If Not IsNothing(oField) Then
                    oField.AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point), Nothing, HandwritingContent(y), y, -1, HandwritingContent(y) <> String.Empty, 0, New Tuple(Of Single)(0)))
                End If
            Next

            ' draw current box
            If (Not bRender) And BlockCount > 0 Then
                XLeft = New XUnit(XFieldDisplacement.X + fAlignmentDisplacementX)
                XTop = New XUnit(XFieldDisplacement.Y + fAlignmentDisplacementY + (CurrentRow * fSingleBlockWidth))
                Dim oXRedPen As XPen = New XPen(XColors.Red, 0.5)
                PDFHelper.DrawFieldBorderPen(oXGraphics, oXRedPen, XWidth, XHeight, New XPoint(XLeft, XTop))
            End If

            ' draw labels
            If BlockCount > 0 Then
                Dim sContent As String = String.Empty
                For i = BlockCount - 1 To 0 Step -1
                    If HandwritingContent(i) = String.Empty Then
                        sContent += "X"
                    Else
                        sContent += " "
                    End If
                Next
                Dim oContentSplit As String() = sContent.Split(" ")

                Dim fPosition As Double = BlockCount
                Dim iCurrentLabel As Integer = -1
                For i = 0 To oContentSplit.Count - 1
                    Dim sContentBlock = oContentSplit(i)
                    fPosition -= sContentBlock.Length / 2
                    If sContentBlock.Count > 0 Then
                        iCurrentLabel += 1
                        If TabletDescriptionTop(iCurrentLabel) <> String.Empty Then
                            PDFHelper.DrawStringRotated(oXGraphics, New XPoint(XFieldDisplacement.X + fAlignmentDisplacementX - fLabelHeight / 2, XFieldDisplacement.Y + fAlignmentDisplacementY + (fPosition * fSingleBlockWidth)), oArielFontLabel, oStringFormat, TabletDescriptionTop(iCurrentLabel), -90)
                        End If
                    End If
                    fPosition -= (sContentBlock.Length / 2) + 1
                Next
            End If

            If Not IsNothing(oField) Then
                ParamList.Value(ParamList.KeyFieldCollection).Fields.Add(oField)
            End If
        End Sub
        Public Overrides Function FieldSpecificProcessing(ByVal oRectFormField As Rect) As Boolean
            ' field specific processing to check whether placement is allowed
            ' default value is false (no problem)
            ' minimum of 2 rows height per block
            'Dim iRowCount As Integer = Math.Max(CInt(Math.Floor(GridRect.Height / 2)), 0)
            Dim iRowCount As Integer = BlockCount

            If oRectFormField.Height < Math.Max(iRowCount, 1) * 2 Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelInput.DataContext = Me
            Root.StackPanelFieldHandwriting.DataContext = Me
            Root.StackPanelFieldHandwritingExtra2.DataContext = Me
            Root.StackPanelFieldBoxChoice.DataContext = Me

            Dim oBindingField1 As New Data.Binding
            oBindingField1.Path = New PropertyPath("Critical")
            oBindingField1.Mode = Data.BindingMode.TwoWay
            oBindingField1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldCritical.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField1)

            Dim oBindingField2 As New Data.Binding
            oBindingField2.Path = New PropertyPath("Exclude")
            oBindingField2.Mode = Data.BindingMode.TwoWay
            oBindingField2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldExclude.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField2)

            Dim oBindingBoxChoice1 As New Data.Binding
            oBindingBoxChoice1.Path = New PropertyPath("CurrentDescriptionTopText")
            oBindingBoxChoice1.Mode = Data.BindingMode.TwoWay
            oBindingBoxChoice1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldBoxChoiceLabel.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBoxChoice1)

            Dim oBindingBoxChoice2 As New Data.Binding
            oBindingBoxChoice2.Path = New PropertyPath("TabletDescriptionTopCurrent")
            oBindingBoxChoice2.Mode = Data.BindingMode.TwoWay
            oBindingBoxChoice2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldBoxChoiceLabelContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBoxChoice2)

            Dim oBindingBoxChoice3 As New Data.Binding
            oBindingBoxChoice3.Path = New PropertyPath("BlockCountText")
            oBindingBoxChoice3.Mode = Data.BindingMode.TwoWay
            oBindingBoxChoice3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingBlockCount.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBoxChoice3)

            Dim oBindingBoxChoice4 As New Data.Binding
            oBindingBoxChoice4.Path = New PropertyPath("CurrentRowText")
            oBindingBoxChoice4.Mode = Data.BindingMode.TwoWay
            oBindingBoxChoice4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingCurrentRow.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBoxChoice4)

            Dim oBindingBoxChoice5 As New Data.Binding
            oBindingBoxChoice5.Path = New PropertyPath("CurrentContent")
            oBindingBoxChoice5.Mode = Data.BindingMode.TwoWay
            oBindingBoxChoice5.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingCurrentContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingBoxChoice5)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "a38625ff-0302-42e8-ab9f-b90548452c2d"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FieldHandwriting
        Inherits FormField

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public BlockCount As Integer
            <DataMember> Public RowCount As Integer
            <DataMember> Public HandwritingContent As String()()
            <DataMember> Public CharacterASCII As Enumerations.CharacterASCII

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .BlockCount = BlockCount
                    .RowCount = RowCount
                    .HandwritingContent = HandwritingContent.Clone
                    .CharacterASCII = CharacterASCII
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldHandwriting).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Protected Overrides Sub RectHeightChanged()
            MyBase.RectHeightChanged()

            m_DataObject.RowCount = Math.Max(CInt(Math.Floor(GridRect.Height / 2)), 0)
            m_CurrentRow = Math.Max(Math.Min(m_CurrentRow, m_DataObject.RowCount - 1), 0)
            ResetHandwritingContent()

            OnPropertyChangedLocal("RowCount")
            OnPropertyChangedLocal("RowCountText")
            OnPropertyChangedLocal("CurrentRow")
            OnPropertyChangedLocal("CurrentRowText")
            OnPropertyChangedLocal("CurrentContent")
        End Sub
        Protected Overrides Sub RectWidthChanged()
            MyBase.RectWidthChanged()

            Dim iMaxBlockCount As Integer = GetMaxBlockCount()
            If m_DataObject.BlockCount > iMaxBlockCount Then
                m_DataObject.BlockCount = iMaxBlockCount
                OnPropertyChangedLocal("BlockCount")
                OnPropertyChangedLocal("BlockCountText")
                OnPropertyChangedLocal("CurrentColumn")
                OnPropertyChangedLocal("CurrentColumnText")
                OnPropertyChangedLocal("CurrentContent")
            End If
        End Sub
        Public Overrides Sub TitleChanged()
            Title = "Block Count: " + BlockCountText + If(Critical, vbCrLf + "Critical Field", String.Empty)
            Title += If(Exclude, vbCr + "Excluded", String.Empty)
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                m_DataObject.BlockCount = 0
                m_DataObject.RowCount = 0
                ReDim m_DataObject.HandwritingContent(0)
                ReDim m_DataObject.HandwritingContent(0)(0)
                m_DataObject.HandwritingContent(0)(0) = String.Empty
                m_DataObject.CharacterASCII = Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric
            End With
            TitleChanged()
        End Sub
#End Region
#Region "Handwriting"
        Private m_CurrentRow As Integer
        Private m_CurrentColumn As Integer

        Public Property BlockCount As Integer
            Get
                Return m_DataObject.BlockCount
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If (Not IsNothing(Parent)) AndAlso Parent.GetType.Equals(GetType(FormBlock)) Then
                        Dim iMaxBlockCount As Integer = GetMaxBlockCount()
                        If value < 0 Then
                            m_DataObject.BlockCount = 0
                        ElseIf value > iMaxBlockCount Then
                            m_DataObject.BlockCount = iMaxBlockCount
                        Else
                            m_DataObject.BlockCount = value
                        End If
                    Else
                        m_DataObject.BlockCount = 0
                    End If
                    m_CurrentColumn = Math.Max(Math.Min(m_CurrentColumn, m_DataObject.BlockCount - 1), 0)
                    ResetHandwritingContent()

                    OnPropertyChangedLocal("BlockCount")
                    OnPropertyChangedLocal("BlockCountText")
                    OnPropertyChangedLocal("CurrentColumnText")
                    OnPropertyChangedLocal("CurrentContent")
                    OnPropertyChangedLocal("CurrentDescriptionTop")
                    OnPropertyChangedLocal("CurrentDescriptionTopText")
                    OnPropertyChangedLocal("TabletDescriptionTopCurrent")
                End Using
            End Set
        End Property
        Public Property BlockCountText As String
            Get
                Return m_DataObject.BlockCount.ToString
            End Get
            Set(value As String)
                BlockCount = CInt(Val(value))
            End Set
        End Property
        Private Function GetMaxBlockCount() As Integer
            ' each block width is equals to two row heights
            Dim oBlock As FormBlock = Parent
            If IsNothing(oBlock) Then
                Return 0
            Else
                Dim fImageWidth As Double = ((oBlock.GetBlockDimensions(False).Item1.Point - oBlock.LeftIndent.Point) * GridRect.Width / oBlock.GridWidth) - (2 * SpacingLarge * 72 / RenderResolution300)
                Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
                Dim iMaxBlockCount As Integer = 0
                If fImageWidth >= fSingleBlockWidth Then
                    iMaxBlockCount = Math.Floor(fImageWidth / fSingleBlockWidth)
                End If
                Return iMaxBlockCount
            End If
        End Function
        Public Property RowCount As Integer
            Get
                Return m_DataObject.RowCount
            End Get
            Set(value As Integer)
            End Set
        End Property
        Public Property RowCountText As String
            Get
                Return m_DataObject.RowCount.ToString
            End Get
            Set(value As String)
            End Set
        End Property
        Public Property CurrentRow As Integer
            Get
                Return m_CurrentRow
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 0 Then
                        m_CurrentRow = 0
                    ElseIf value > m_DataObject.RowCount - 1 Then
                        m_CurrentRow = m_DataObject.RowCount - 1
                    Else
                        m_CurrentRow = value
                    End If

                    OnPropertyChangedLocal("CurrentRow")
                    OnPropertyChangedLocal("CurrentRowText")
                    OnPropertyChangedLocal("CurrentContent")
                End Using
            End Set
        End Property
        Public Property CurrentRowText As String
            Get
                Return (m_CurrentRow + 1).ToString
            End Get
            Set(value As String)
            End Set
        End Property
        Public Property CurrentColumn As Integer
            Get
                Return m_CurrentColumn
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 0 Then
                        m_CurrentColumn = 0
                    ElseIf value > m_DataObject.BlockCount - 1 Then
                        m_CurrentColumn = m_DataObject.BlockCount - 1
                    Else
                        m_CurrentColumn = value
                    End If

                    OnPropertyChangedLocal("CurrentColumn")
                    OnPropertyChangedLocal("CurrentColumnText")
                    OnPropertyChangedLocal("CurrentContent")
                End Using
            End Set
        End Property
        Public Property CurrentColumnText As String
            Get
                Return (m_CurrentColumn + 1).ToString
            End Get
            Set(value As String)
            End Set
        End Property
        Public Property CurrentContent As String
            Get
                Return m_DataObject.HandwritingContent(m_CurrentRow)(m_CurrentColumn)
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.HandwritingContent(m_CurrentRow)(m_CurrentColumn) = Left(Trim(value), 1).Replace("-", "–")
                    OnPropertyChangedLocal("CurrentContent")
                End Using
            End Set
        End Property
        Public ReadOnly Property HandwritingContent(ByVal iRow As Integer, ByVal iColumn As Integer) As String
            Get
                If iRow >= 0 And iRow <= m_DataObject.RowCount - 1 And iColumn >= 0 And iColumn <= m_DataObject.BlockCount - 1 Then
                    Return m_DataObject.HandwritingContent(iRow)(iColumn)
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public Property CharacterASCII As Enumerations.CharacterASCII
            Get
                Return m_DataObject.CharacterASCII
            End Get
            Set(value As Enumerations.CharacterASCII)
                m_DataObject.CharacterASCII = value
                OnPropertyChangedLocal("CharacterASCII")
                OnPropertyChangedLocal("CharacterASCIIText")
            End Set
        End Property
        Public Property CharacterASCIIText As String
            Get
                Select Case m_DataObject.CharacterASCII
                    Case Enumerations.CharacterASCII.None, Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric
                        Return "All"
                    Case Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers
                        Return "Alphanumeric"
                    Case Enumerations.CharacterASCII.Numbers
                        Return "Numeric"
                    Case Else
                        Return "None"
                End Select
            End Get
            Set(value As String)
            End Set
        End Property
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Handwriting"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMHandwriting"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelSelection")
                oDisplayFilter.Add("StackPanelInput")
                oDisplayFilter.Add("StackPanelAlignmentStart")
                oDisplayFilter.Add("StackPanelAlignmentExtra1")
                oDisplayFilter.Add("StackPanelAlignmentExtra2")
                oDisplayFilter.Add("StackPanelAlignmentBase")
                oDisplayFilter.Add("StackPanelFieldHandwriting")
                oDisplayFilter.Add("StackPanelFieldHandwritingMore")
                oDisplayFilter.Add("StackPanelFieldHandwritingExtra1")
                oDisplayFilter.Add("StackPanelFieldHandwritingExtra2")
                oDisplayFilter.Add("FieldHandwritingBlockCount")
                oDisplayFilter.Add("FieldHandwritingRowCount")
                oDisplayFilter.Add("FieldHandwritingType")
                oDisplayFilter.Add("FieldHandwritingCurrentColumn")
                oDisplayFilter.Add("FieldHandwritingCurrentRow")
                oDisplayFilter.Add("FieldHandwritingCurrentContent")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return False
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim sWord As String = "LOREM"
            Dim fDimension As Double = oBackgroundSize.Width * 0.8 / sWord.Count

            ' determine displacements
            Dim fFirstDisplacement As Double = oBackgroundSize.Width * 0.1 + fDimension / 2
            Dim fLastDisplacement As Double = oBackgroundSize.Width - fFirstDisplacement
            Dim fSpacing As Double = (fLastDisplacement - fFirstDisplacement) / (sWord.Count - 1)
            Dim XDisplacementList As New List(Of Tuple(Of Char, XUnit))
            For i = 0 To sWord.Count - 1
                XDisplacementList.Add(New Tuple(Of Char, XUnit)(Mid(sWord, i + 1, 1), New XUnit(fFirstDisplacement + (fSpacing * i))))
            Next
            Dim YDisplacement As New XUnit((iBitmapHeight * 72 / RenderResolution300) / 2)

            ' draw rectangles
            For Each XDisplacement As Tuple(Of Char, XUnit) In XDisplacementList
                PDFHelper.DrawFieldBorder(oXGraphics, New XUnit(fDimension), New XUnit(fDimension), New XPoint(XDisplacement.Item2.Point - fDimension / 2, YDisplacement.Point - fDimension / 2), 0.5)
            Next

            ' draw letters
            Const fFontSize As Double = 10
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontComicSansMS, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * 0.8 * fDimension / oTestFont.GetHeight
            Dim oComicSansFont As New XFont(FontComicSansMS, fScaledFontSize, XFontStyle.Regular, oFontOptions)

            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            For Each XDisplacement As Tuple(Of Char, XUnit) In XDisplacementList
                Dim oXSize As XSize = oXGraphics.MeasureString(XDisplacement.Item1, oComicSansFont)
                oXGraphics.DrawString(XDisplacement.Item1, oComicSansFont, XBrushes.Black, XDisplacement.Item2, YDisplacement, oStringFormat)
            Next
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
            Dim fAlignmentDisplacementX As Double = 0
            Dim fAlignmentDisplacementY As Double = 0
            Select Case Alignment
                Case Enumerations.AlignmentEnum.Left
                    fAlignmentDisplacementY = (XImageHeight.Point - (RowCount * fSingleBlockWidth)) / 2
                Case Enumerations.AlignmentEnum.UpperLeft
                Case Enumerations.AlignmentEnum.Top
                    fAlignmentDisplacementX = (XImageWidth.Point - (BlockCount * fSingleBlockWidth)) / 2
                Case Enumerations.AlignmentEnum.UpperRight
                    fAlignmentDisplacementX = XImageWidth.Point - (BlockCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.Right
                    fAlignmentDisplacementX = XImageWidth.Point - (BlockCount * fSingleBlockWidth)
                    fAlignmentDisplacementY = (XImageHeight.Point - (RowCount * fSingleBlockWidth)) / 2
                Case Enumerations.AlignmentEnum.LowerRight
                    fAlignmentDisplacementX = XImageWidth.Point - (BlockCount * fSingleBlockWidth)
                    fAlignmentDisplacementY = XImageHeight.Point - (RowCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.Bottom
                    fAlignmentDisplacementX = (XImageWidth.Point - (BlockCount * fSingleBlockWidth)) / 2
                    fAlignmentDisplacementY = XImageHeight.Point - (RowCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.LowerLeft
                    fAlignmentDisplacementY = XImageHeight.Point - (RowCount * fSingleBlockWidth)
                Case Enumerations.AlignmentEnum.Center
                    fAlignmentDisplacementX = (XImageWidth.Point - (BlockCount * fSingleBlockWidth)) / 2
                    fAlignmentDisplacementY = (XImageHeight.Point - (RowCount * fSingleBlockWidth)) / 2
            End Select

            ' process input fields if present
            Dim oField As FieldDocumentStore.Field = Nothing
            If ParamList.ContainsKey(ParamList.KeyFieldCollection) AndAlso Not Exclude Then
                oField = New FieldDocumentStore.Field
                With oField
                    .FieldType = Enumerations.FieldTypeEnum.Handwriting
                    .GUID = Guid.NewGuid
                    .CharacterASCII = CharacterASCII
                    .Numbering = Numbering
                    .Location = New Rect(XFieldDisplacement.X + fAlignmentDisplacementX, XFieldDisplacement.Y + fAlignmentDisplacementY, fSingleBlockWidth * BlockCount, fSingleBlockWidth * RowCount)
                    .PageNumber = ParamList.Value(ParamList.KeyPDFPages).Count
                    .Critical = Critical
                End With
            End If

            Dim XWidth As New XUnit(fSingleBlockWidth)
            Dim XHeight As New XUnit(fSingleBlockWidth)
            Dim XLeft As XUnit = Nothing
            Dim XTop As XUnit = Nothing

            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            Dim fFontSize As Double = PDFHelper.GetScaledFontSize(XWidth, XHeight, XHeight)
            Dim oArielFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            For x = 0 To BlockCount - 1
                For y = 0 To RowCount - 1
                    XLeft = New XUnit(XFieldDisplacement.X + fAlignmentDisplacementX + (x * fSingleBlockWidth))
                    XTop = New XUnit(XFieldDisplacement.Y + fAlignmentDisplacementY + (y * fSingleBlockWidth))

                    ' draw text if present
                    If HandwritingContent(y, x) = String.Empty Then
                        If bRender Then
                            PDFHelper.DrawFieldBorder(oXGraphics, XWidth, XHeight, New XPoint(XLeft, XTop), 0.5)
                        Else
                            If CurrentColumn <> x Or CurrentRow <> y Then
                                PDFHelper.DrawFieldBorder(oXGraphics, XWidth, XHeight, New XPoint(XLeft, XTop), 0.5)
                            End If
                        End If
                    Else
                        oXGraphics.DrawString(HandwritingContent(y, x), oArielFont, XBrushes.Black, XLeft.Point + XWidth.Point / 2, XTop.Point + XHeight.Point / 2, oStringFormat)
                    End If

                    If Not IsNothing(oField) Then
                        oField.AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point), Nothing, HandwritingContent(y, x), y, x, HandwritingContent(y, x) <> String.Empty, 0, New Tuple(Of Single)(0)))
                    End If
                Next
            Next

            ' draw current box
            If (Not bRender) And BlockCount > 0 Then
                XLeft = New XUnit(XFieldDisplacement.X + fAlignmentDisplacementX + (CurrentColumn * fSingleBlockWidth))
                XTop = New XUnit(XFieldDisplacement.Y + fAlignmentDisplacementY + (CurrentRow * fSingleBlockWidth))
                Dim oXRedPen As XPen = New XPen(XColors.Red, 0.5)
                PDFHelper.DrawFieldBorderPen(oXGraphics, oXRedPen, XWidth, XHeight, New XPoint(XLeft, XTop))
            End If

            If Not IsNothing(oField) Then
                ParamList.Value(ParamList.KeyFieldCollection).Fields.Add(oField)
            End If
        End Sub
        Public Overrides Function FieldSpecificProcessing(ByVal oRectFormField As Rect) As Boolean
            ' field specific processing to check whether placement is allowed
            ' default value is false (no problem)
            ' minimum of 2 rows height per block
            Dim iRowCount As Integer = Math.Max(CInt(Math.Floor(GridRect.Height / 2)), 0)
            If oRectFormField.Height < Math.Max(iRowCount, 1) * 2 Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelInput.DataContext = Me
            Root.StackPanelFieldHandwriting.DataContext = Me
            Root.StackPanelFieldHandwritingMore.DataContext = Me
            Root.StackPanelFieldHandwritingExtra1.DataContext = Me
            Root.StackPanelFieldHandwritingExtra2.DataContext = Me

            Dim oBindingField1 As New Data.Binding
            oBindingField1.Path = New PropertyPath("Critical")
            oBindingField1.Mode = Data.BindingMode.TwoWay
            oBindingField1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldCritical.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField1)

            Dim oBindingField2 As New Data.Binding
            oBindingField2.Path = New PropertyPath("Exclude")
            oBindingField2.Mode = Data.BindingMode.TwoWay
            oBindingField2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldExclude.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField2)

            Dim oBindingHandwriting1 As New Data.Binding
            oBindingHandwriting1.Path = New PropertyPath("BlockCountText")
            oBindingHandwriting1.Mode = Data.BindingMode.TwoWay
            oBindingHandwriting1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingBlockCount.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingHandwriting1)

            Dim oBindingHandwriting2 As New Data.Binding
            oBindingHandwriting2.Path = New PropertyPath("RowCountText")
            oBindingHandwriting2.Mode = Data.BindingMode.TwoWay
            oBindingHandwriting2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingRowCount.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingHandwriting2)

            Dim oBindingHandwriting3 As New Data.Binding
            oBindingHandwriting3.Path = New PropertyPath("CurrentColumnText")
            oBindingHandwriting3.Mode = Data.BindingMode.TwoWay
            oBindingHandwriting3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingCurrentColumn.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingHandwriting3)

            Dim oBindingHandwriting4 As New Data.Binding
            oBindingHandwriting4.Path = New PropertyPath("CurrentRowText")
            oBindingHandwriting4.Mode = Data.BindingMode.TwoWay
            oBindingHandwriting4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingCurrentRow.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingHandwriting4)

            Dim oBindingHandwriting5 As New Data.Binding
            oBindingHandwriting5.Path = New PropertyPath("CurrentContent")
            oBindingHandwriting5.Mode = Data.BindingMode.TwoWay
            oBindingHandwriting5.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingCurrentContent.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingHandwriting5)

            Dim oBindingHandwriting6 As New Data.Binding
            oBindingHandwriting6.Path = New PropertyPath("CharacterASCIIText")
            oBindingHandwriting6.Mode = Data.BindingMode.TwoWay
            oBindingHandwriting6.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldHandwritingType.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingHandwriting6)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "138f280b-8b08-44c4-966d-54ea4b6eb462"
            End Get
        End Property
        Private Sub ResetHandwritingContent()
            Dim oNewHandwritingContent(Math.Max(m_DataObject.RowCount - 1, 0))() As String
            For y = 0 To oNewHandwritingContent.Length - 1
                ReDim oNewHandwritingContent(y)(Math.Max(m_DataObject.BlockCount - 1, 0))
                For x = 0 To oNewHandwritingContent(y).Length - 1
                    oNewHandwritingContent(y)(x) = String.Empty
                Next
            Next
            For y = 0 To Math.Min(m_DataObject.HandwritingContent.Length, oNewHandwritingContent.Length) - 1
                For x = 0 To Math.Min(m_DataObject.HandwritingContent(y).Length, oNewHandwritingContent(y).Length) - 1
                    oNewHandwritingContent(y)(x) = If(IsNothing(m_DataObject.HandwritingContent(y)(x)), String.Empty, m_DataObject.HandwritingContent(y)(x))
                Next
            Next
            m_DataObject.HandwritingContent = oNewHandwritingContent
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FieldFree
        Inherits FormField

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FieldFree).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Title = If(Critical, "Critical Field", String.Empty)
            Title += If(Exclude, vbCr + "Excluded", String.Empty)
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Free Entry"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMPencil"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.AddRange(MyBase.DisplayFilter)
                oDisplayFilter.Add("StackPanelSelection")
                oDisplayFilter.Add("StackPanelInput")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides ReadOnly Property IsRestricted As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property NoOverlap As Boolean
            Get
                Return False
            End Get
        End Property
        Public Shared Shadows Function FieldTypeImage(ByVal fReferenceHeight As Double, ByVal oMethodInfoGetFieldTypeImageProcess As Reflection.MethodInfo) As ImageSource
            Return GetFieldTypeImage(fReferenceHeight, oMethodInfoGetFieldTypeImageProcess)
        End Function
        Public Shared Shadows Sub GetFieldTypeImageProcess(ByVal fReferenceHeight As Double, ByVal oBackgroundSize As XSize, ByVal oXGraphics As XGraphics, ByVal iBitmapWidth As Integer, ByVal iBitmapHeight As Integer)
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            ' draw border
            PDFHelper.DrawFieldBorder(oXGraphics, New XUnit(oBackgroundSize.Width * 0.95), New XUnit(oBackgroundSize.Height * 0.95), New XPoint(oBackgroundSize.Width * 0.025, oBackgroundSize.Height * 0.025), 0.5)

            ' draw lorem ipsum
            Dim oLoremIpsum As DrawingImage = Converter.XamlToDrawingImage(My.Resources.CCMLoremIpsum)
            Dim oLoremIpsumBitmap As System.Drawing.Bitmap = Converter.BitmapSourceToBitmap(Converter.DrawingImageToBitmapSource(oLoremIpsum, iBitmapWidth * 0.8, iBitmapHeight * 0.8, Enumerations.StretchEnum.Uniform), RenderResolution300)
            Dim oXLoremIpsumImage As XImage = PdfSharp.Drawing.XImage.FromGdiPlusImage(oLoremIpsumBitmap)
            Dim XDisplacement As New XUnit(((iBitmapWidth - oLoremIpsumBitmap.Width) * 72 / RenderResolution300) / 2)
            Dim YDisplacement As New XUnit(((iBitmapHeight - oLoremIpsumBitmap.Height) * 72 / RenderResolution300) / 2)
            oXGraphics.DrawImage(oXLoremIpsumImage, XDisplacement, YDisplacement, oXLoremIpsumImage.PointWidth, oXLoremIpsumImage.PointHeight)
        End Sub
        Public Overrides Sub DrawFieldContents(ByVal bRender As Boolean, ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XFieldDisplacement As XPoint, Optional ByVal oOverflowRows As List(Of Integer) = Nothing)
            Dim XLeft As New XUnit(XFieldDisplacement.X + PDFHelper.XMargin.Point)
            Dim XTop As New XUnit(XFieldDisplacement.Y + PDFHelper.XMargin.Point)
            Dim XWidth As New XUnit(XImageWidth.Point - (PDFHelper.XMargin.Point * 2))
            Dim XHeight As New XUnit(XImageHeight.Point - (PDFHelper.XMargin.Point * 2))
            PDFHelper.DrawFieldBorder(oXGraphics, XWidth, XHeight, New XPoint(XLeft.Point, XTop.Point), 0.5)

            ' process input fields if present
            If ParamList.ContainsKey(ParamList.KeyFieldCollection) AndAlso Not Exclude Then
                Dim oField As New FieldDocumentStore.Field
                With oField
                    .FieldType = Enumerations.FieldTypeEnum.Free
                    .GUID = Guid.NewGuid
                    .Numbering = Numbering
                    .Location = New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point)
                    .PageNumber = ParamList.Value(ParamList.KeyPDFPages).Count
                    .AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point), Nothing, String.Empty, 0, 0, False, 0, New Tuple(Of Single)(0)))
                End With

                ParamList.Value(ParamList.KeyFieldCollection).Fields.Add(oField)
            End If
        End Sub
        Public Overrides Sub SetBindings()
            MyBase.SetBindings()

            Root.StackPanelInput.DataContext = Me

            Dim oBindingField1 As New Data.Binding
            oBindingField1.Path = New PropertyPath("Critical")
            oBindingField1.Mode = Data.BindingMode.TwoWay
            oBindingField1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldCritical.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField1)

            Dim oBindingField2 As New Data.Binding
            oBindingField2.Path = New PropertyPath("Exclude")
            oBindingField2.Mode = Data.BindingMode.TwoWay
            oBindingField2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FieldExclude.SetBinding(Common.HighlightButton.HBSelectedProperty, oBindingField2)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "83acaeb5-c8e9-439a-8812-fe0e5f709ba2"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FormMCQ
        Inherits FormBlock

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public BlockWidth As Integer
            <DataMember> Public MCQ As New MCQClass(StringMCQ)

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .BlockWidth = BlockWidth
                    .MCQ = MCQ.Clone
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
            CType(oFormItem, FormMCQ).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Overrides Property BlockWidth As Integer
            Get
                Return m_DataObject.BlockWidth
            End Get
            Set(value As Integer)
                Dim iOldBlockWidth As Integer = m_DataObject.BlockWidth

                If Me.GetType.Equals(GetType(FormFormHeader)) Then
                    m_DataObject.BlockWidth = PDFHelper.PageBlockWidth
                Else
                    If value < 1 Then
                        m_DataObject.BlockWidth = 1
                    ElseIf value > PDFHelper.PageBlockWidth Then
                        m_DataObject.BlockWidth = PDFHelper.PageBlockWidth
                    Else
                        m_DataObject.BlockWidth = value
                    End If
                End If

                If iOldBlockWidth <> m_DataObject.BlockWidth Then
                    Using oSuspender As New Suspender(Me, True)
                        OnPropertyChangedLocal("BlockWidth")
                        OnPropertyChangedLocal("BlockWidthText")
                        UpdateAnswers()
                    End Using
                End If
            End Set
        End Property
        Public Overrides Property BlockWidthText As String
            Get
                Return m_DataObject.BlockWidth.ToString
            End Get
            Set(value As String)
                BlockWidth = CInt(Val(value))
            End Set
        End Property
        Public Property MCQText As String
            Get
                If IsNothing(m_DataObject) OrElse IsNothing(m_DataObject.MCQ) Then
                    Return String.Empty
                Else
                    Return m_DataObject.MCQ.Name
                End If
            End Get
            Set(value As String)
                m_DataObject.MCQ.Name = value
                OnPropertyChangedLocal("MCQText")
                TitleChanged()
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "MCQ:   " + MCQText
        End Sub
        Public Overrides Sub Initialise()
            MyBase.Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .BlockWidth = PDFHelper.PageBlockWidth
            End With
            TitleChanged()

            ResetQuestionAnswers()

            ShowGrid = False
            ShowActive = False
        End Sub
        Private Sub ResetQuestionAnswers()
            ' set question and answer fields to default
            Dim oFieldTextList As List(Of FieldText) = ExtractFields(Of FieldText)(GetType(FieldText))
            Dim oFieldChoiceVerticalMCQ As List(Of FieldChoiceVerticalMCQ) = ExtractFields(Of FieldChoiceVerticalMCQ)(GetType(FieldChoiceVerticalMCQ))
            Dim oFieldFreeList As List(Of FieldFree) = ExtractFields(Of FieldFree)(GetType(FieldFree))

            Dim oQuestion As FieldText = Nothing
            Dim oAnswers As FieldChoiceVerticalMCQ = Nothing
            Dim oFree As FieldFree = Nothing
            If oFieldTextList.Count = 0 Then
                oQuestion = New FieldText
                oQuestion.Parent = Me
            Else
                oQuestion = oFieldTextList.First
            End If
            If oFieldChoiceVerticalMCQ.Count = 0 Then
                oAnswers = New FieldChoiceVerticalMCQ
                oAnswers.Parent = Me
            Else
                oAnswers = oFieldChoiceVerticalMCQ.First
            End If
            If oFieldFreeList.Count = 0 Then
                oFree = New FieldFree
                oFree.Parent = Me
            Else
                oFree = oFieldFreeList.First
            End If

            oQuestion.GridRect = New Int32Rect(0, 0, 1, 1)
            oQuestion.Alignment = Enumerations.AlignmentEnum.Left
            oQuestion.Justification = Enumerations.JustificationEnum.Justify
            oQuestion.ClearElements()
            oQuestion.AddElement(PDFHelper.LoremIpsum(0))

            oAnswers.SetTabletDescriptionMCQ(New List(Of List(Of ElementStruc)) From {New List(Of ElementStruc) From {New ElementStruc(PDFHelper.LoremIpsum(1))}, New List(Of ElementStruc) From {New ElementStruc(PDFHelper.LoremIpsum(2))}, New List(Of ElementStruc) From {New ElementStruc(PDFHelper.LoremIpsum(3))}, New List(Of ElementStruc) From {New ElementStruc(PDFHelper.LoremIpsum(4))}, New List(Of ElementStruc) From {New ElementStruc(PDFHelper.LoremIpsum(5))}, New List(Of ElementStruc) From {New ElementStruc(PDFHelper.LoremIpsum(6))}})
            oAnswers.Alignment = Enumerations.AlignmentEnum.Left
            oAnswers.TabletCount = oAnswers.TabletDescriptionMCQ.Count
            oAnswers.TabletStart = 0
            oAnswers.TabletGroups = BlockWidth
            oAnswers.TabletContent = Enumerations.TabletContentEnum.Letter
            oAnswers.GridRect = New Int32Rect(0, 1, 1, GetFieldHeight(oAnswers).Item2)

            oFree.ResetGridRect()

            GridWidth = 1
            GridHeight = oQuestion.GridRect.Height + oAnswers.GridRect.Height

            GetBlockDimensions(True)
            UpdateGridHeight()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "MCQ Header"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.0
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMMCQ"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelBlock")
                oDisplayFilter.Add("StackPanelBlockWidth")
                oDisplayFilter.Add("StackPanelMCQ")
                oDisplayFilter.Add("BlockWidth")
                oDisplayFilter.Add("BlockMCQ")
                oDisplayFilter.Add("RectangleBlockBackground")
                oDisplayFilter.Add("ImageBlockContent")
                oDisplayFilter.Add("ScrollViewerBlockContent")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelBlock.DataContext = Me
            Root.StackPanelBlockWidth.DataContext = Me
            Root.StackPanelMCQ.DataContext = Me

            Dim oBindingMCQ1 As New Data.Binding
            oBindingMCQ1.Path = New PropertyPath("BlockWidthText")
            oBindingMCQ1.Mode = Data.BindingMode.TwoWay
            oBindingMCQ1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockWidth.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingMCQ1)

            Dim oBindingMCQ2 As New Data.Binding
            oBindingMCQ2.Path = New PropertyPath("MCQText")
            oBindingMCQ2.Mode = Data.BindingMode.TwoWay
            oBindingMCQ2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockMCQ.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingMCQ2)
        End Sub
        Public Overrides Sub Display()
            MyBase.Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "f8c3cf2c-6451-4262-991f-a014025e25ca"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterDivider), GetType(FieldNumbering)}
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedSingleOnly As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterDivider)}
            End Get
        End Property
        Public Overrides ReadOnly Property IgnoreFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FieldText), GetType(FieldChoiceVerticalMCQ), GetType(FieldFree)}
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            ' run through all MCQ items
            ' add to MCQ list
            Using oSuspender As New Suspender()
                Dim oLineMCQList As New List(Of Tuple(Of Integer, Double, List(Of Integer)))
                For i = 0 To m_DataObject.MCQ.MCQStore.Count - 1
                    Dim oMCQItem As MCQClass.MCQItem = m_DataObject.MCQ.MCQStore(i)
                    Dim oMCQDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = GetMCQDimensions(oMCQItem)
                    oLineMCQList.Add(New Tuple(Of Integer, Double, List(Of Integer))(i, oMCQDimensions.Item2.Point, oMCQDimensions.Item3))
                Next

                ' use items from the list to fill up columns of approximately equal height
                Dim iPageRepeatCount As Integer = 0
                Dim oColumnList As New Dictionary(Of Integer, List(Of Tuple(Of Integer, Double, List(Of Integer))))
                Dim iColumnCount As Integer = Math.Floor(PDFHelper.PageBlockWidth / BlockWidth)
                For i = 0 To iColumnCount - 1
                    oColumnList.Add(i, New List(Of Tuple(Of Integer, Double, List(Of Integer))))
                Next
                Do Until oLineMCQList.Count = 0 Or iPageRepeatCount > 1
                    Dim fRemainingPageAllowance As Double = PDFHelper.PageLimitBottomExBarcode.Point - PDFHelper.PageLimitTop.Point - ParamList.Value(ParamList.KeyXCurrentHeight).Point

                    Dim iCurrentColumn As Integer = 0
                    Dim bLoopExited As Boolean = False
                    For i = 0 To oLineMCQList.Count - 1
                        Dim fCurrentColumnHeight As Double = (Aggregate iRow In Enumerable.Range(0, oColumnList(iCurrentColumn).Count) Into Sum(oColumnList(iCurrentColumn)(iRow).Item2)) + ((Math.Max(oColumnList(iCurrentColumn).Count - 1, 0)) * 2 * PDFHelper.BlockSpacer.Point)
                        If fCurrentColumnHeight + oLineMCQList(i).Item2 > fRemainingPageAllowance Then
                            If i = 0 Then
                                ' no space to render on this page, increment and start again
                                ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(PDFHelper.PageLimitBottom.Point * 2)
                                FormPDF.AddPage(0)

                                ' exit loop
                                iPageRepeatCount += 1
                                bLoopExited = True
                                Exit For
                            ElseIf iCurrentColumn = iColumnCount - 1 Then
                                ' already at last column, increment page and remove items from the column list
                                RenderColumnList(oColumnList)

                                ' increment page
                                ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(PDFHelper.PageLimitBottom.Point * 2)
                                FormPDF.AddPage(0)

                                ' remove rendered items from the line MCQ list
                                For j = i - 1 To 0 Step -1
                                    oLineMCQList.RemoveAt(j)
                                Next

                                ' clear column list
                                For j = 0 To iColumnCount - 1
                                    oColumnList(j).Clear()
                                Next

                                ' exit loop
                                bLoopExited = True
                                Exit For
                            Else
                                ' go to next column and add
                                iCurrentColumn += 1
                                oColumnList(iCurrentColumn).Add(oLineMCQList(i))
                            End If
                        Else
                            ' add to current column
                            oColumnList(iCurrentColumn).Add(oLineMCQList(i))
                        End If
                    Next

                    ' final render
                    If (Not bLoopExited) Or iPageRepeatCount > 1 Then
                        ' clear column list
                        For j = 0 To iColumnCount - 1
                            oColumnList(j).Clear()
                        Next

                        Dim fColumnHeight As Double = ((Aggregate iIndex In Enumerable.Range(0, oLineMCQList.Count) Into Sum(oLineMCQList(iIndex).Item2)) + (oLineMCQList.Count * 2 * PDFHelper.BlockSpacer.Point)) / iColumnCount
                        iCurrentColumn = 0
                        For i = 0 To oLineMCQList.Count - 1
                            ' add to current column
                            oColumnList(iCurrentColumn).Add(oLineMCQList(i))
                            Dim fCurrentColumnHeight As Double = (Aggregate iRow In Enumerable.Range(0, oColumnList(iCurrentColumn).Count) Into Sum(oColumnList(iCurrentColumn)(iRow).Item2)) + ((Math.Max(oColumnList(iCurrentColumn).Count - 1, 0)) * 2 * PDFHelper.BlockSpacer.Point)
                            If fCurrentColumnHeight > fColumnHeight OrElse ((i < oLineMCQList.Count - 1) AndAlso oLineMCQList(i + 1).Item2 > fColumnHeight) Then
                                iCurrentColumn += 1
                            End If
                        Next

                        ' renders final column list
                        RenderColumnList(oColumnList)
                        oLineMCQList.Clear()
                        oColumnList.Clear()
                    End If
                Loop

                If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                    ParamList.Value(ParamList.KeyNumberingCurrent) += 1
                End If

                ' restore initial state
                ResetQuestionAnswers()
            End Using
        End Sub
        Public Sub DisplayPDF()
            ' renders the MCQ to a bitmap
            Using oSuspender As New Suspender()
                If m_DataObject.MCQ.MCQStore.Count > 0 Then
                    Dim oMCQImage As Imaging.BitmapSource = Nothing
                    Dim oSection As FormSection = ParamList.Value(ParamList.KeyCurrentSection)
                    If Not IsNothing(oSection) Then
                        oMCQImage = oSection.ImageTracker.GetImage(FormSection.ImageTrackerClass.GetList({Me}))
                        If IsNothing(oMCQImage) Then
                            Dim oLineMCQList As New List(Of Tuple(Of Integer, Double, List(Of Integer)))
                            For i = 0 To m_DataObject.MCQ.MCQStore.Count - 1
                                Dim oMCQItem As MCQClass.MCQItem = m_DataObject.MCQ.MCQStore(i)
                                Dim oMCQDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = GetMCQDimensions(oMCQItem)
                                oLineMCQList.Add(New Tuple(Of Integer, Double, List(Of Integer))(i, oMCQDimensions.Item2.Point, oMCQDimensions.Item3))
                            Next

                            ' use items from the list to fill up columns of approximately equal height
                            Dim iPageRepeatCount As Integer = 0
                            Dim oColumnList As New Dictionary(Of Integer, List(Of Tuple(Of Integer, Double, List(Of Integer))))
                            Dim iColumnCount As Integer = Math.Floor(PDFHelper.PageBlockWidth / BlockWidth)
                            For i = 0 To iColumnCount - 1
                                oColumnList.Add(i, New List(Of Tuple(Of Integer, Double, List(Of Integer))))
                            Next

                            Dim fColumnHeight As Double = ((Aggregate iIndex In Enumerable.Range(0, oLineMCQList.Count) Into Sum(oLineMCQList(iIndex).Item2)) + (oLineMCQList.Count * 2 * PDFHelper.BlockSpacer.Point)) / iColumnCount
                            Dim iCurrentColumn As Integer = 0
                            For i = 0 To oLineMCQList.Count - 1
                                ' add to current column
                                oColumnList(iCurrentColumn).Add(oLineMCQList(i))
                                Dim fCurrentColumnHeight As Double = (Aggregate iRow In Enumerable.Range(0, oColumnList(iCurrentColumn).Count) Into Sum(oColumnList(iCurrentColumn)(iRow).Item2)) + (oColumnList(iCurrentColumn).Count * 2 * PDFHelper.BlockSpacer.Point)
                                If fCurrentColumnHeight > fColumnHeight Then
                                    iCurrentColumn += 1
                                End If
                            Next

                            Dim fActualHeight As Double = 0
                            For i = 0 To iColumnCount - 1
                                Dim iCurrent As Integer = i
                                fActualHeight = Math.Max(fActualHeight, (Aggregate iRow In Enumerable.Range(0, oColumnList(iCurrent).Count) Into Sum(oColumnList(iCurrent)(iRow).Item2)) + (oColumnList(iCurrent).Count * 2 * PDFHelper.BlockSpacer.Point))
                            Next

                            ' render to bitmap
                            Dim XMCQWidth As New XUnit(PDFHelper.PageLimitWidth.Point)
                            Dim XMCQHeight As New XUnit(fActualHeight)
                            Dim XMCQSize As New XSize(XMCQWidth.Point, XMCQHeight.Point)

                            Using oBitmap As New System.Drawing.Bitmap(Math.Ceiling(XMCQWidth.Inch * oSection.SectionResolution), Math.Ceiling(XMCQHeight.Inch * oSection.SectionResolution), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                oBitmap.SetResolution(oSection.SectionResolution, oSection.SectionResolution)

                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                    oGraphics.FillRectangle(System.Drawing.Brushes.White, 0, 0, oBitmap.Width, oBitmap.Height)
                                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XMCQSize, XGraphicsUnit.Point)
                                        Dim XBlockWidth As XUnit = XUnit.FromPoint((BlockWidth * PDFHelper.BlockWidth.Point) + ((BlockWidth - 1) * 3 * PDFHelper.BlockSpacer.Point))
                                        Dim fCurrentDisplacementX As Double = 0

                                        For i = 0 To oColumnList.Keys.Count - 1
                                            Dim fCurrentDisplacementY As Double = 0
                                            For j = 0 To oColumnList(i).Count - 1
                                                ' set question and answer parameters
                                                Dim oMCQItem As MCQClass.MCQItem = m_DataObject.MCQ.MCQStore(oColumnList(i)(j).Item1)
                                                GetMCQDimensions(oMCQItem)

                                                Dim XBlockHeight As XUnit = oColumnList(i)(j).Item2
                                                Dim XExpandedBlockWidth As New XUnit(XBlockWidth.Point + (2 * PDFHelper.BlockSpacer.Point))
                                                Dim XExpandedBlockHeight As New XUnit(XBlockHeight.Point + (2 * PDFHelper.BlockSpacer.Point))

                                                Dim XDisplacement As New XPoint(fCurrentDisplacementX, fCurrentDisplacementY)
                                                DrawBlockDirect(True, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, XDisplacement, XBlockWidth, XBlockHeight, New XPoint(PDFHelper.BlockSpacer.Point, PDFHelper.BlockSpacer.Point), oColumnList(i)(j).Item3)

                                                fCurrentDisplacementY += XExpandedBlockHeight.Point

                                                ' increments subnumbering if present, otherwise increase numbering
                                                If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                                                    ParamList.Value(ParamList.KeySubNumberingCurrent) += 1
                                                Else
                                                    ParamList.Value(ParamList.KeyNumberingCurrent) += 1
                                                End If
                                            Next

                                            ' increment displacement
                                            fCurrentDisplacementX += XBlockWidth.Point + (PDFHelper.BlockSpacer.Point * 3)
                                        Next
                                    End Using
                                End Using

                                oMCQImage = Converter.BitmapToBitmapSource(oBitmap)
                            End Using

                            ' restore initial state
                            ResetQuestionAnswers()

                            oSection.ImageTracker.Add(oMCQImage, FormSection.ImageTrackerClass.GetList({Me}))
                        End If

                        If Not IsNothing(oMCQImage) Then
                            FormSection.ImageViewer.Add(oMCQImage, "MCQ Image")
                        End If
                    End If

                    If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                        ParamList.Value(ParamList.KeyNumberingCurrent) += 1
                    End If
                End If
            End Using
        End Sub
        Private Function GetMCQDimensions(ByVal oMCQItem As MCQClass.MCQItem) As Tuple(Of XUnit, XUnit, List(Of Integer))
            Dim oFieldTextList As List(Of FieldText) = ExtractFields(Of FieldText)(GetType(FieldText))
            Dim oFieldChoiceVerticalMCQ As List(Of FieldChoiceVerticalMCQ) = ExtractFields(Of FieldChoiceVerticalMCQ)(GetType(FieldChoiceVerticalMCQ))
            Dim oFieldFreeList As List(Of FieldFree) = ExtractFields(Of FieldFree)(GetType(FieldFree))

            Dim oQuestion As FieldText = Nothing
            Dim oAnswers As FieldChoiceVerticalMCQ = Nothing
            Dim oFree As FieldFree = Nothing
            If oFieldTextList.Count = 0 Then
                oQuestion = New FieldText
                oQuestion.Parent = Me
            Else
                oQuestion = oFieldTextList.First
            End If
            If oFieldChoiceVerticalMCQ.Count = 0 Then
                oAnswers = New FieldChoiceVerticalMCQ
                oAnswers.Parent = Me
            Else
                oAnswers = oFieldChoiceVerticalMCQ.First
            End If
            If oFieldFreeList.Count = 0 Then
                oFree = New FieldFree
                oFree.Parent = Me
            Else
                oFree = oFieldFreeList.First
            End If

            oQuestion.GridRect = New Int32Rect(0, 0, 1, 1)
            oQuestion.Alignment = Enumerations.AlignmentEnum.Left
            oQuestion.Justification = Enumerations.JustificationEnum.Justify
            oQuestion.ReplaceElements(oMCQItem.QuestionElements)

            oAnswers.SetTabletDescriptionMCQ(oMCQItem.Answers)
            oAnswers.Alignment = Enumerations.AlignmentEnum.Left
            oAnswers.TabletCount = oAnswers.TabletDescriptionMCQ.Count
            oAnswers.TabletStart = 0
            oAnswers.TabletGroups = BlockWidth
            oAnswers.TabletContent = oMCQItem.TabletContent
            oAnswers.GridRect = New Int32Rect(0, 1, 1, GetFieldHeight(oAnswers).Item2)

            If oMCQItem.Free = 0 Then
                oFree.ResetGridRect()
            Else
                oFree.GridRect = New Int32Rect(0, 2 + oAnswers.GridRect.Height, 1, oMCQItem.Free)
            End If

            GridWidth = 1
            GridHeight = oQuestion.GridRect.Height + oAnswers.GridRect.Height + oFree.GridRect.Height + 1

            GetBlockDimensions(True)
            UpdateGridHeight()

            Return GetBlockDimensions(True)
        End Function
        Private Sub RenderColumnList(ByRef oColumnList As Dictionary(Of Integer, List(Of Tuple(Of Integer, Double, List(Of Integer)))))
            ' renders current column list
            Dim fCurrentDisplacementX As Double = PDFHelper.PageLimitLeft.Point
            Dim XBlockWidth As XUnit = XUnit.FromPoint((BlockWidth * PDFHelper.BlockWidth.Point) + ((BlockWidth - 1) * 3 * PDFHelper.BlockSpacer.Point))
            Dim oDisplacementYList As New List(Of Double)

            For i = 0 To oColumnList.Keys.Count - 1
                Dim fCurrentDisplacementY As Double = 0
                For j = 0 To oColumnList(i).Count - 1
                    ' set question and answer parameters
                    Dim oMCQItem As MCQClass.MCQItem = m_DataObject.MCQ.MCQStore(oColumnList(i)(j).Item1)
                    GetMCQDimensions(oMCQItem)

                    Dim XBlockHeight As XUnit = oColumnList(i)(j).Item2
                    Dim XExpandedBlockWidth As New XUnit(XBlockWidth.Point + (2 * PDFHelper.BlockSpacer.Point))
                    Dim XExpandedBlockHeight As New XUnit(XBlockHeight.Point + (2 * PDFHelper.BlockSpacer.Point))

                    Dim XDisplacement As New XPoint(fCurrentDisplacementX, PDFHelper.PageLimitTop.Point + ParamList.Value(ParamList.KeyXCurrentHeight).Point + fCurrentDisplacementY)
                    Using oXGraphics As XGraphics = XGraphics.FromPdfPage(ParamList.Value(ParamList.KeyPDFPages)(ParamList.Value(ParamList.KeyPDFPages).Count - 1).Item1, XGraphicsPdfPageOptions.Append)
                        DrawBlockDirect(True, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, XDisplacement, XBlockWidth, XBlockHeight, New XPoint(PDFHelper.BlockSpacer.Point, PDFHelper.BlockSpacer.Point), oColumnList(i)(j).Item3)
                    End Using

                    fCurrentDisplacementY += XExpandedBlockHeight.Point

                    ' increments subnumbering if present, otherwise increase numbering
                    If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                        ParamList.Value(ParamList.KeySubNumberingCurrent) += 1
                    Else
                        ParamList.Value(ParamList.KeyNumberingCurrent) += 1
                    End If
                Next

                oDisplacementYList.Add(fCurrentDisplacementY)

                ' increment displacement
                fCurrentDisplacementX += XBlockWidth.Point + (PDFHelper.BlockSpacer.Point * 3)
            Next

            ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(ParamList.Value(ParamList.KeyXCurrentHeight).Point + oDisplacementYList.Max)
        End Sub
        Public Sub ImportQuestions()
            ' imports questions from Excel file
            ' each worksheet represents one question
            ' allows formatting with italic, underline and bold type
            ' first column lists fields, second column lists content, all subsequent columns are ignored
            ' Question: question
            ' Format: 'letter' or 'number', default is 'letter'
            ' Answer: correct answers as a list of letters eg. "AD" for 'A' and 'D' correct
            ' if no answer supplied, then this question does not have a fixed answer
            ' Answers: each subsequent non-empty line is treated as an answer, and rich text is accepted
            Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
            oOpenFileDialog.FileName = String.Empty
            oOpenFileDialog.DefaultExt = "*.xlsx"
            oOpenFileDialog.Multiselect = False
            oOpenFileDialog.Filter = "Excel Spreadsheet|*.xlsx"
            oOpenFileDialog.Title = "Load MCQ Questions From File"
            Dim result? As Boolean = oOpenFileDialog.ShowDialog()
            If result = True Then
                Dim oFileInfo As New IO.FileInfo(oOpenFileDialog.FileName)
                Dim oExcelDocument As ClosedXML.Excel.XLWorkbook = Nothing
                Try
                    oExcelDocument = New ClosedXML.Excel.XLWorkbook(oOpenFileDialog.FileName)
                Catch ex1 As System.IO.FileFormatException
                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oFileInfo.Name + ". File format error."))
                    Exit Sub
                Catch ex2 As System.IO.IOException
                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oFileInfo.Name + ". File access error."))
                    Exit Sub
                End Try

                Dim oMCQ As New MCQClass(Left(oFileInfo.Name, Len(oFileInfo.Name) - Len(oFileInfo.Extension)))
                For Each oWorksheet As ClosedXML.Excel.IXLWorksheet In oExcelDocument.Worksheets
                    ' check for correct format
                    If Trim(oWorksheet.Cell(1, 1).GetString.ToUpper) = "QUESTION" And Trim(oWorksheet.Cell(2, 1).GetString.ToUpper) = "FORMAT" And Trim(oWorksheet.Cell(3, 1).GetString.ToUpper) = "SINGLE" And Trim(oWorksheet.Cell(4, 1).GetString.ToUpper) = "CRITICAL" And Trim(oWorksheet.Cell(5, 1).GetString.ToUpper) = "FREE" And Trim(oWorksheet.Cell(6, 1).GetString.ToUpper) = "ANSWER" And Trim(oWorksheet.Cell(7, 1).GetString.ToUpper) = "ANSWERS" Then
                        Dim oMCQItem As New MCQClass.MCQItem
                        With oMCQItem
                            Dim oQuestion As ClosedXML.Excel.IXLCell = oWorksheet.Cell(1, 2)
                            .QuestionElements.AddRange(ConvertRichText(oQuestion))
                            If Trim(oWorksheet.Cell(2, 2).GetString.ToUpper) = "NUMBER" Then
                                .TabletContent = Enumerations.TabletContentEnum.Number
                            ElseIf Trim(oWorksheet.Cell(2, 2).GetString.ToUpper) = "LETTER" Then
                                .TabletContent = Enumerations.TabletContentEnum.Letter
                            Else
                                .TabletContent = Enumerations.TabletContentEnum.None
                            End If
                            If Trim(oWorksheet.Cell(3, 2).GetString.ToUpper) = "Y" Then
                                .TabletSingleChoiceOnly = True
                            Else
                                .TabletSingleChoiceOnly = False
                            End If

                            If Trim(oWorksheet.Cell(4, 2).GetString.ToUpper) = "Y" Then
                                .Critical = True
                            Else
                                .Critical = False
                            End If

                            .Free = oWorksheet.Cell(5, 2).GetValue(Of Integer)()

                            Dim oRichText As New List(Of ElementStruc)
                            Dim iCurrentRow As Integer = 7
                            Do Until IsNothing(oRichText)
                                Dim oReturn As ClosedXML.Excel.IXLCell = oWorksheet.Cell(iCurrentRow, 2)
                                oRichText = ConvertRichText(oReturn)
                                If Not IsNothing(oRichText) Then
                                    .Answers.Add(oRichText)
                                End If
                                iCurrentRow += 1
                            Loop

                            ' only add answer list after the answers themselves have been added
                            Dim oAnswerList As Char() = Trim(oWorksheet.Cell(6, 2).GetString.ToUpper).ToCharArray
                            For Each sChar As Char In oAnswerList
                                Dim iNumber As Integer = Converter.ConvertLetterToNumber(sChar)
                                If iNumber >= 0 And iNumber < .Answers.Count And (Not .CorrectAnswers.Contains(iNumber)) Then
                                    .CorrectAnswers.Add(iNumber)
                                End If
                            Next
                        End With

                        If oMCQItem.Answers.Count > 0 Then
                            oMCQ.MCQStore.Add(oMCQItem)
                        End If
                    Else
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error processing worksheet """ + oWorksheet.Name + """."))
                    End If
                Next
                If oMCQ.MCQStore.Count > 0 Then
                    Using oSuspender As New Suspender(Me, True)
                        m_DataObject.MCQ = oMCQ
                        OnPropertyChangedLocal("MCQText")
                    End Using
                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "MCQ set """ + m_DataObject.MCQ.Name + """ added."))
                Else
                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "No valid MCQs to add."))
                End If
            End If
        End Sub
        Private Function ConvertRichText(ByVal oCell As ClosedXML.Excel.IXLCell) As List(Of ElementStruc)
            ' converts the supplied rich text into a list of elements
            Dim oElements As New List(Of ElementStruc)
            Dim oRichText As ClosedXML.Excel.IXLRichText = oCell.RichText
            If oRichText.Length > 0 Then
                Dim bBold As Boolean = oRichText(0).Bold
                Dim bItalic As Boolean = oRichText(0).Italic
                Dim oUnderline As ClosedXML.Excel.XLFontUnderlineValues = oRichText(0).Underline

                Dim sText As String = oRichText(0).Text
                For i = 1 To oRichText.Count - 1
                    If oRichText(i).Bold = bBold And oRichText(i).Italic = bItalic And oRichText(i).Underline = oUnderline Then
                        sText += oRichText(i).Text
                    Else
                        oElements.Add(New ElementStruc(sText, ElementStruc.ElementTypeEnum.Text, bBold, bItalic, If(oUnderline = ClosedXML.Excel.XLFontUnderlineValues.None, False, True)))
                        sText = oRichText(i).Text
                        bBold = oRichText(i).Bold
                        bItalic = oRichText(i).Italic
                        oUnderline = oRichText(i).Underline
                    End If
                Next
                If sText.Length > 0 Then
                    oElements.Add(New ElementStruc(sText, ElementStruc.ElementTypeEnum.Text, bBold, bItalic, If(oUnderline = ClosedXML.Excel.XLFontUnderlineValues.None, False, True)))
                End If
            Else
                Return Nothing
            End If
            Return oElements
        End Function
        Private Sub UpdateGridHeight()
            Dim oFormFieldList As List(Of FormField) = ExtractFields(Of FormField)(GetType(FormField))
            GridHeight = Aggregate oFormField In oFormFieldList Into Max(oFormField.GridRect.Y + Math.Min(oFormField.GridRect.Height, oFormField.GridAdjustedHeight))
        End Sub
        Private Sub UpdateAnswers()
            ' updates height for answers when block width changes
            Dim oFieldChoiceVerticalMCQList As List(Of FieldChoiceVerticalMCQ) = ExtractFields(Of FieldChoiceVerticalMCQ)(GetType(FieldChoiceVerticalMCQ))
            If oFieldChoiceVerticalMCQList.Count > 0 Then
                Dim oAnswers As FieldChoiceVerticalMCQ = oFieldChoiceVerticalMCQList.First
                oAnswers.TabletGroups = BlockWidth
                oAnswers.GridRect = New Int32Rect(0, 1, 1, GetFieldHeight(oAnswers).Item2)
                GetBlockDimensions(True)
                UpdateGridHeight()
            End If
        End Sub
        Public Function GetFieldHeight(ByVal oFieldMCQ As FieldChoiceVerticalMCQ) As Tuple(Of Double, Integer)
            ' gets the field height
            ' return value is a tuple of the row count and the actual height in points
            Dim XBlockWidth As XUnit = XUnit.FromPoint((BlockWidth * PDFHelper.BlockWidth.Point) + ((BlockWidth - 1) * 3 * PDFHelper.BlockSpacer.Point))
            Dim fLeftIndent As Double = LeftIndent()
            Dim XImageWidth As XUnit = XUnit.FromPoint(XBlockWidth.Point - fLeftIndent)
            Dim oMCQHeight As Tuple(Of Double, Integer, List(Of Tuple(Of Integer, Integer))) = PDFHelper.GetFieldTabletsMCQHeight(XImageWidth, 1, oFieldMCQ.TabletGroups, oFieldMCQ.TabletDescriptionMCQTuple)
            Return New Tuple(Of Double, Integer)(oMCQHeight.Item1, oMCQHeight.Item2)
        End Function
        <DataContract(IsReference:=True)> Public Class MCQClass
            Implements ICloneable

            <DataMember> Public MCQStore As New ObservableCollection(Of MCQItem)
            <DataMember> Private m_Name As String = String.Empty

            Sub New(ByVal sName As String)
                m_Name = sName
            End Sub
            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oMCQClass As New MCQClass(m_Name)
                With oMCQClass
                    For Each oMCQItem As MCQItem In MCQStore
                        .MCQStore.Add(oMCQItem.Clone)
                    Next
                End With
                Return oMCQClass
            End Function
            Public Property Name As String
                Get
                    Return m_Name
                End Get
                Set(value As String)
                    m_Name = value
                End Set
            End Property
            <DataContract(IsReference:=True)> Public Class MCQItem
                Implements ICloneable
                Implements INotifyPropertyChanged

                <DataMember> Private m_QuestionElements As New List(Of ElementStruc)
                <DataMember> Private m_TabletContent As Enumerations.TabletContentEnum
                <DataMember> Private m_TabletSingleChoiceOnly As Boolean
                <DataMember> Private m_Critical As Boolean
                <DataMember> Private m_CorrectAnswers As New List(Of Integer)
                <DataMember> Private m_Answers As New List(Of List(Of ElementStruc))
                <DataMember> Private m_Free As Integer

                Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
                Protected Sub OnPropertyChangedLocal(ByVal sName As String)
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
                End Sub
                Public Function Clone() As Object Implements ICloneable.Clone
                    Return New MCQItem(m_QuestionElements, m_TabletContent, m_TabletSingleChoiceOnly, m_Critical, m_CorrectAnswers, m_Answers, m_Free)
                End Function
                Sub New()
                    m_TabletContent = Enumerations.TabletContentEnum.None
                    m_TabletSingleChoiceOnly = False
                    m_Critical = False
                    m_Free = -1
                End Sub
                Sub New(ByVal oQuestionElements As List(Of ElementStruc), ByVal oTabletContent As Enumerations.TabletContentEnum, ByVal bTabletSingleChoiceOnly As Boolean, ByVal bCritical As Boolean, ByVal oCorrectAnswers As List(Of Integer), ByVal oAnswers As List(Of List(Of ElementStruc)), ByVal iFree As Integer)
                    m_QuestionElements.AddRange(oQuestionElements)
                    m_TabletContent = oTabletContent
                    m_TabletSingleChoiceOnly = bTabletSingleChoiceOnly
                    m_Critical = bCritical
                    m_CorrectAnswers.AddRange(oCorrectAnswers)
                    m_Answers.AddRange(oAnswers)
                    m_Free = iFree
                End Sub
                Public Property QuestionElements As List(Of ElementStruc)
                    Get
                        Return m_QuestionElements
                    End Get
                    Set(value As List(Of ElementStruc))
                        m_QuestionElements.Clear()
                        m_QuestionElements.AddRange(value)
                        OnPropertyChangedLocal("QuestionElements")
                    End Set
                End Property
                Public Property TabletContent As Enumerations.TabletContentEnum
                    Get
                        Return m_TabletContent
                    End Get
                    Set(value As Enumerations.TabletContentEnum)
                        m_TabletContent = value
                        OnPropertyChangedLocal("TabletContent")
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
                Public Property Critical As Boolean
                    Get
                        Return m_Critical
                    End Get
                    Set(value As Boolean)
                        m_Critical = value
                        OnPropertyChangedLocal("Critical")
                    End Set
                End Property
                Public Property CorrectAnswers As List(Of Integer)
                    Get
                        Return m_CorrectAnswers
                    End Get
                    Set(value As List(Of Integer))
                        m_CorrectAnswers.Clear()
                        m_CorrectAnswers.AddRange(value)
                        OnPropertyChangedLocal("CorrectAnswers")
                    End Set
                End Property
                Public Property Answers As List(Of List(Of ElementStruc))
                    Get
                        Return m_Answers
                    End Get
                    Set(value As List(Of List(Of ElementStruc)))
                        m_Answers.Clear()
                        m_Answers.AddRange(value)
                        OnPropertyChangedLocal("Answers")
                    End Set
                End Property
                Public Property Free As Integer
                    Get
                        Return m_Free
                    End Get
                    Set(value As Integer)
                        m_Free = value
                        OnPropertyChangedLocal("Free")
                    End Set
                End Property
            End Class
        End Class
    End Class
    <DataContract(IsReference:=True)> Public Class FormBody
        Inherits BaseFormItem

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
        End Sub
        Public Overrides Sub Initialise()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Body"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.5
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMBody"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Return New List(Of String)
            End Get
        End Property
        Public Overrides Sub SetBindings()
        End Sub
        Public Overrides Sub Display()
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "1da230c5-b539-4070-9c07-49f5673bea83"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormSection), GetType(FormatterPageBreak), GetType(FormatterDivider)}
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
        Public Overrides ReadOnly Property FormattingFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterPageBreak), GetType(FormatterDivider)}
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FormFooter
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public ShowTitle As Boolean
            <DataMember> Public ShowPage As Boolean
            <DataMember> Public ShowSubject As Boolean

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .ShowTitle = ShowTitle
                    .ShowPage = ShowPage
                    .ShowSubject = ShowSubject
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Property ShowTitle As Boolean
            Get
                Return m_DataObject.ShowTitle
            End Get
            Set(value As Boolean)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.ShowTitle = value
                    OnPropertyChangedLocal("ShowTitle")
                End Using
            End Set
        End Property
        Public Property ShowPage As Boolean
            Get
                Return m_DataObject.ShowPage
            End Get
            Set(value As Boolean)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.ShowPage = value
                    OnPropertyChangedLocal("ShowPage")
                End Using
            End Set
        End Property
        Public Property ShowSubject As Boolean
            Get
                Return m_DataObject.ShowSubject
            End Get
            Set(value As Boolean)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.ShowSubject = value
                    OnPropertyChangedLocal("ShowSubject")
                End Using
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Dim sTitle As String = String.Empty
            If m_DataObject.ShowTitle Then
                sTitle += "Title: Visible"
            Else
                sTitle += "Title: Hidden"
            End If
            If m_DataObject.ShowPage Then
                sTitle += vbCr + "Page: Visible"
            Else
                sTitle += vbCr + "Page: Hidden"
            End If
            If m_DataObject.ShowSubject Then
                sTitle += vbCr + "Subject: Visible"
            Else
                sTitle += vbCr + "Subject: Hidden"
            End If
            Title = sTitle
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                m_DataObject.ShowTitle = True
                m_DataObject.ShowPage = True
                m_DataObject.ShowSubject = True
            End With

            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Footer"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.5
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMFooter"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelFooter")
                Return oDisplayFilter
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelFooter.DataContext = Me

            Dim oBindingFooter1 As New Data.Binding
            oBindingFooter1.Path = New PropertyPath("ShowTitle")
            oBindingFooter1.Mode = Data.BindingMode.TwoWay
            oBindingFooter1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FooterShowTitle.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingFooter1)

            Dim oBindingFooter2 As New Data.Binding
            oBindingFooter2.Path = New PropertyPath("ShowPage")
            oBindingFooter2.Mode = Data.BindingMode.TwoWay
            oBindingFooter2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FooterShowPage.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingFooter2)

            Dim oBindingFooter3 As New Data.Binding
            oBindingFooter3.Path = New PropertyPath("ShowSubject")
            oBindingFooter3.Mode = Data.BindingMode.TwoWay
            oBindingFooter3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.FooterShowSubject.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingFooter3)
        End Sub
        Public Overrides Sub Display()
            Dim iCurrentPage As Integer = 0
            Dim iPageCount As Integer = -1
            Dim oPDFViewerControl As PDFViewer.PDFViewer = Root.PDFViewerControl
            If Not IsNothing(oPDFViewerControl.PDFDocument) Then
                iCurrentPage = oPDFViewerControl.CurrentPage
                iPageCount = oPDFViewerControl.PDFDocument.PageCount
            End If

            Dim oBarcodeText As Tuple(Of String, String, String) = FormPDF.GetBarcodeText(iCurrentPage + 1, iPageCount, "[Subject]", String.Empty)
            CType(Root.ImageFooterBarcode, Controls.Image).Source = Converter.BitmapToBitmapSource(PDFHelper.GetBarcode(oBarcodeText.Item1, oBarcodeText.Item2, oBarcodeText.Item3))
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "7564f869-1a6c-4421-a095-a6e94c3287be"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FormPDF
        Inherits BaseFormItem

        Public Changed As Boolean
        Private PDFViewerControl As PDFViewer.PDFViewer

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            Throw New NotImplementedException
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Title = "Page Count: " + If(IsNothing(PDFViewerControl), "0", PDFViewerControl.PDFDocument.PageCount.ToString)
        End Sub
        Public Overrides Sub Initialise()
            PDFViewerControl = Root.PDFViewerControl
            Changed = True
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "PDF Preview"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.5
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMPDF"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("PDFViewerControl")
                Return oDisplayFilter
            End Get
        End Property
        Public Overrides Sub SetBindings()
        End Sub
        Public Overrides Sub Display()
            If Changed Then
                Dim XCurrentHeight As XUnit = New XUnit(PDFHelper.PageLimitBottom.Point * 2)

                Using oParamListFieldCollection As New ParamList(ParamList.KeyFieldCollection, New FieldDocumentStore.FieldCollection)
                    Using oParamListRemoveBarcode As New ParamList(ParamList.KeyRemoveBarcode, False)
                        Using oParamListPDFDocument As New ParamList(ParamList.KeyPDFDocument, PDFViewerControl.PDFDocument)
                            Using oParamXCurrentHeight As New ParamList(ParamList.KeyXCurrentHeight, XCurrentHeight)
                                Using oParamListSubjectName As New ParamList(ParamList.KeyCurrentSubject, StringSubject)
                                    RenderPDF()
                                End Using
                            End Using
                            oParamListPDFDocument.DisposeReturn(PDFViewerControl.PDFDocument)
                        End Using
                    End Using
                End Using

                PDFViewerControl.Update()
                TitleChanged()
                Changed = False
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "d8881070-ec61-4015-b54f-0b8b2585bb80"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            ' gets the PDF document
            If IsNothing(ParamList.Value(ParamList.KeyPDFDocument)) Then
                ParamList.Value(ParamList.KeyPDFDocument) = New Pdf.PdfDocument
            End If

            If ParamList.Value(ParamList.KeyPDFDocument).PageCount > 0 Then
                For i = ParamList.Value(ParamList.KeyPDFDocument).PageCount - 1 To 0 Step -1
                    ParamList.Value(ParamList.KeyPDFDocument).Pages.RemoveAt(i)
                Next
            End If

            Using oParamListPageCount As New ParamList(ParamList.KeyPageCount, -1)
                Using oParamListTopic As New ParamList(ParamList.KeyTopic, FormProperties.FormTopic)
                    Using oParamListAppendText As New ParamList(ParamList.KeyAppendText, String.Empty)
                        Using oParamListPDFPages As New ParamList(ParamList.KeyPDFPages, New List(Of Tuple(Of Pdf.PdfPage, String, String, String)))
                            Using oParamListBarCodes As New ParamList(ParamList.KeyBarCodes, New List(Of String))
                                Using oParamDisplacement As New ParamList(ParamList.KeyDisplacement, New XPoint(PDFHelper.PageLimitLeft.Point, PDFHelper.PageLimitTop.Point))
                                    ' insert logic for document info
                                    With ParamList.Value(ParamList.KeyPDFDocument)
                                        .Info.Title = FormProperties.FormTitle
                                        .Info.Author = FormProperties.FormAuthor + " (" + ModuleName + " - " + PluginName + ")"
                                        .Info.Subject = ParamList.Value(ParamList.KeyTopic)
                                    End With

                                    ' add the first page
                                    AddPage(0)

                                    ' add the form header
                                    FormFormHeader.RenderPDF()

                                    ' add sections
                                    For i = 0 To FormBody.Children.Count - 1
                                        Select Case FormBody.Children.Values(i)
                                            Case GetType(FormSection)
                                                Dim oSection As FormSection = FormMain.FindChild(FormBody.Children.Keys(i))
                                                oSection.RenderPDF()

                                                ' add section spacing if not the last item and no page break after
                                                If i < FormBody.Children.Count - 1 AndAlso (Not FormBody.Children.Values(i + 1).Equals(GetType(FormatterPageBreak))) Then
                                                    Dim fSectionSpacing As Double = PDFHelper.BlockSpacer.Point
                                                    ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(ParamList.Value(ParamList.KeyXCurrentHeight).Point + fSectionSpacing)
                                                    FormPDF.AddPage(fSectionSpacing)
                                                End If
                                            Case GetType(FormatterPageBreak)
                                                ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(PDFHelper.PageLimitBottomExBarcode.Point + 1)
                                                FormPDF.AddPage(0)
                                        End Select
                                    Next

                                    ' add bar codes
                                    AddPageBarcodes()
                                    ParamList.Value(ParamList.KeyFieldCollection).BarCodes = ParamList.Value(ParamList.KeyBarCodes)
                                    ParamList.Value(ParamList.KeyFieldCollection).RawBarCodes = (From oPDFPage As Tuple(Of Pdf.PdfPage, String, String, String) In CType(ParamList.Value(ParamList.KeyPDFPages), List(Of Tuple(Of Pdf.PdfPage, String, String, String))) Select New Tuple(Of String, String, String)(oPDFPage.Item2, oPDFPage.Item3, oPDFPage.Item4)).ToList

                                    ' set order on fields
                                    Dim iOrder As Integer = 0
                                    For Each oField As FieldDocumentStore.Field In ParamList.Value(ParamList.KeyFieldCollection).Fields
                                        oField.Order = iOrder
                                        iOrder += 1
                                    Next
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        End Sub
#Region "PDF"
        Public Sub AddPage(ByVal fFormItemHeight As Double)
            ' adds a new page and reset current height
            Dim bPaginate As Boolean = True
            If ParamList.ContainsKey(ParamList.KeySuspendPagination) AndAlso ParamList.Value(ParamList.KeySuspendPagination) Then
                bPaginate = False
            End If

            ' paginate only if not suspended
            If bPaginate Then
                If ParamList.Value(ParamList.KeyXCurrentHeight).Point + fFormItemHeight > PDFHelper.PageLimitBottomExBarcode.Point Then
                    ParamList.Value(ParamList.KeyXCurrentHeight) = XUnit.Zero
                    ParamList.Value(ParamList.KeyPDFPages).Add(New Tuple(Of Pdf.PdfPage, String, String, String)(CreatePDFPage(ParamList.Value(ParamList.KeyPDFDocument), FormPageHeader.Lines, FormPageHeader.FontSizeMultiplier), String.Empty, String.Empty, String.Empty))
                End If
            End If
        End Sub
        Private Shared Function CreatePDFPage(ByRef oPDFDocument As Pdf.PdfDocument, ByVal oHeaderTextList As List(Of String), ByVal iFontSizeMultiplier As Integer) As Pdf.PdfPage
            ' creates basic PDF page
            Dim oPage As Pdf.PdfPage = oPDFDocument.AddPage()
            oPage.Orientation = FormProperties.SelectedOrientation
            oPage.Size = FormProperties.SelectedSize

            ' draws the dotted line on the left side at 1 cm
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                If oXGraphics.PageUnit = XGraphicsUnit.Point Then
                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                    Dim oBlackPen As New XPen(XColors.Black, 1)
                    oBlackPen.DashStyle = XDashStyle.Dot
                    oXGraphics.DrawLine(oBlackPen, New XPoint(XUnit.FromCentimeter(1.5).Point, XUnit.FromCentimeter(2.5).Point), New XPoint(XUnit.FromCentimeter(1.5).Point, oPage.Height))
                    oXGraphics.DrawLine(oBlackPen, New XPoint(XUnit.FromCentimeter(1.5).Point, XUnit.FromCentimeter(2.5).Point), New XPoint(XUnit.FromCentimeter(4).Point, 0))
                End If
            End Using

            ' adds "Detachable Margin" to the left margin
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                If oXGraphics.PageUnit = XGraphicsUnit.Point Then
                    Const sDetachableMargin As String = "Detachable Margin"
                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                    Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
                    Dim oArielFontItalic As New XFont(FontArial, 12, XFontStyle.Italic, oFontOptions)
                    Dim oCenterPoint As New XPoint(XUnit.FromCentimeter(1).Point, oXGraphics.PageSize.Height / 2)
                    Dim oStringFormat As New XStringFormat
                    oStringFormat.Alignment = XStringAlignment.Center
                    oStringFormat.LineAlignment = XLineAlignment.Center
                    oXGraphics.RotateAtTransform(270, oCenterPoint)
                    oXGraphics.DrawString(sDetachableMargin, oArielFontItalic, XBrushes.Black, PDFHelper.GetTextXRect(sDetachableMargin, oArielFontItalic, oXGraphics, oCenterPoint, Enumerations.CenterAlignment.Center), oStringFormat)
                End If
            End Using

            ' draw cornerstones
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                If oXGraphics.PageUnit = XGraphicsUnit.Point Then
                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                    Dim oXImage As XImage = XImage.FromGdiPlusImage(PDFHelper.GetCornerstoneImage(SpacingSmall, RenderResolution300 * 2, RenderResolution300 * 2, False))
                    Dim width As Double = oXImage.PixelWidth * 72 / oXImage.HorizontalResolution
                    Dim height As Double = oXImage.PixelHeight * 72 / oXImage.VerticalResolution

                    For Each oCornerstone As Tuple(Of Integer, Integer, XPoint) In PDFHelper.Cornerstone
                        oXGraphics.DrawImage(oXImage, oCornerstone.Item3.X - width / 2, oCornerstone.Item3.Y - height / 2, width, height)
                    Next
                End If
            End Using

            ' sets page header
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                If oXGraphics.PageUnit = XGraphicsUnit.Point Then
                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                    Dim oCenterPointX As XUnit = XUnit.FromPoint(XUnit.FromCentimeter(6).Point + (oXGraphics.PageSize.Width - XUnit.FromCentimeter(4.5).Point - XUnit.FromCentimeter(6).Point) / 2)
                    Dim oCenterPointY As XUnit = XUnit.FromCentimeter(1.5)
                    FormPageHeader.DrawPageHeader(oXGraphics, New XPoint(oCenterPointX, oCenterPointY), oHeaderTextList, iFontSizeMultiplier, XStringAlignment.Center, XBrushes.Black)
                End If
            End Using

            Return oPage
        End Function
        Public Sub AddPageBarcodes()
            ' adds the page barcodes
            For iPage = 0 To ParamList.Value(ParamList.KeyPDFPages).Count - 1
                Dim oBarcodeText As Tuple(Of String, String, String) = GetBarcodeText(iPage + 1, ParamList.Value(ParamList.KeyPDFPages).Count, ParamList.Value(ParamList.KeyCurrentSubject), ParamList.Value(ParamList.KeyAppendText))
                AddPDFBarcode(ParamList.Value(ParamList.KeyPDFPages)(iPage).Item1, oBarcodeText.Item1, oBarcodeText.Item2, oBarcodeText.Item3)
                ParamList.Value(ParamList.KeyPDFPages)(iPage) = New Tuple(Of Pdf.PdfPage, String, String, String)(ParamList.Value(ParamList.KeyPDFPages)(iPage).Item1, oBarcodeText.Item1, oBarcodeText.Item2, oBarcodeText.Item3)

                If Left(ParamList.Value(ParamList.KeyCurrentSubject), 1) = "[" And Right(ParamList.Value(ParamList.KeyCurrentSubject), 1) = "]" Then
                    Dim oAltBarcodeText As Tuple(Of String, String, String) = GetBarcodeText(ParamList.Value(ParamList.KeyPDFPages).Count, ParamList.Value(ParamList.KeyPageCount), "[Template]", String.Empty)
                    ParamList.Value(ParamList.KeyBarCodes).Add(oAltBarcodeText.Item1)
                Else
                    ParamList.Value(ParamList.KeyBarCodes).Add(oBarcodeText.Item1)
                End If
            Next
        End Sub
        Public Shared Function GetBarcodeText(ByVal iCurrentPage As Integer, ByVal iPageCount As Integer, ByVal sSubjectName As String, ByVal sAppendText As String) As Tuple(Of String, String, String)
            ' gets barcode text fields
            ' if subject name is enclosed with square brackets, then omit form title from bar code
            Dim sBarcodeData As String = String.Empty
            If Left(sSubjectName, 1) = "[" And Right(sSubjectName, 1) = "]" Then
                sBarcodeData = SeparatorChar + iCurrentPage.ToString + SeparatorChar + sSubjectName
            Else
                sBarcodeData = FormProperties.FormTitle.Replace(SeparatorChar, String.Empty).GetHashCode.ToString + SeparatorChar + iCurrentPage.ToString + SeparatorChar + sSubjectName.GetHashCode.ToString + If(sAppendText = String.Empty, String.Empty, SeparatorChar + sAppendText)
            End If

            Dim sBarcodeTopText As String = String.Empty
            Dim sBarcodeBottomText As String = String.Empty
            If (FormFooter.ShowTitle And FormProperties.FormTitle <> String.Empty) And FormFooter.ShowPage Then
                If iPageCount < 1 Then
                    sBarcodeTopText = FormProperties.FormTitle + " - Page " + iCurrentPage.ToString
                Else
                    sBarcodeTopText = FormProperties.FormTitle + " - Page " + iCurrentPage.ToString + "/" + iPageCount.ToString
                End If
            ElseIf (FormFooter.ShowTitle And FormProperties.FormTitle <> String.Empty) Then
                sBarcodeTopText = FormProperties.FormTitle
            ElseIf FormFooter.ShowPage Then
                If iPageCount < 1 Then
                    sBarcodeTopText = "Page " + iCurrentPage.ToString
                Else
                    sBarcodeTopText = "Page " + iCurrentPage.ToString + "/" + iPageCount.ToString
                End If
            End If
            If FormFooter.ShowSubject Then
                sBarcodeBottomText = sSubjectName
            End If

            Return New Tuple(Of String, String, String)(sBarcodeData, sBarcodeTopText, sBarcodeBottomText)
        End Function
        Private Shared Sub AddPDFBarcode(ByRef oPDFPage As Pdf.PdfPage, ByVal sBarcodeData As String, ByVal sBarcodeTopText As String, ByVal sBarcodeBottomText As String)
            ' adds the barcode to the page
            If sBarcodeData.Length <= PDFHelper.BarcodeCharLimit Then
                ' sets identification bar code
                If sBarcodeData <> String.Empty Then
                    ' if the barcode data is empty, then do not draw the barcode
                    PDFHelper.DrawBarcode(oPDFPage, sBarcodeData, sBarcodeTopText, sBarcodeBottomText)
                End If
            End If
        End Sub
#End Region
    End Class
    <DataContract(IsReference:=True)> Public Class FormExport
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public Subjects As TagClass

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .Subjects = Subjects.Clone
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            CType(oFormItem, FormExport).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Overrides Sub TitleChanged()
            Title = "Subjects: " + m_DataObject.Subjects.Count.ToString
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .Subjects = New TagClass(WorksheetSubjects, WorksheetSubjects, 1)
            End With
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Export"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.5
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMSubjects"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelBlock")
                oDisplayFilter.Add("StackPanelSubjects")
                oDisplayFilter.Add("StackPanelTagsDataRow")
                oDisplayFilter.Add("DataGridSubjectsTags")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelBlock.DataContext = Me
            Root.StackPanelSubjects.DataContext = Me
        End Sub
        Public Overrides Sub Display()
            ' validate the selected tag number and rebuild the datagrid columns
            Dim oDataGridSubjectsTags As Controls.DataGrid = Root.DataGridSubjectsTags

            oDataGridSubjectsTags.Columns.Clear()
            If (Not IsNothing(m_DataObject.Subjects)) Then
                For i = 0 To m_DataObject.Subjects.Width
                    Dim oColumn As New Controls.DataGridTextColumn
                    With oColumn
                        If i = 0 Then
                            .Width = New Controls.DataGridLength(oDataGridSubjectsTags.ActualWidth * 0.05, Controls.DataGridLengthUnitType.Pixel)
                            .Header = "No."
                            .IsReadOnly = True

                            Dim oBinding As New Data.Binding
                            oBinding.Path = New PropertyPath("Number")
                            oBinding.Mode = Data.BindingMode.TwoWay
                            oBinding.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                            .Binding = oBinding
                        Else
                            .Width = New Controls.DataGridLength(1, Controls.DataGridLengthUnitType.Star)
                            .Header = m_DataObject.Subjects.Header

                            Dim oBinding As New Data.Binding
                            oBinding.Path = New PropertyPath("Tags[0]")
                            oBinding.Mode = Data.BindingMode.TwoWay
                            oBinding.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                            .Binding = oBinding
                        End If
                    End With
                    oDataGridSubjectsTags.Columns.Add(oColumn)
                Next
                oDataGridSubjectsTags.ItemsSource = m_DataObject.Subjects.TagStore
            Else
                oDataGridSubjectsTags.ItemsSource = Nothing
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "f3ee6d66-99fe-41b1-8c28-4ff78949c27a"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Throw New NotImplementedException
        End Sub
        Public Sub AddRow()
            Using oSuspender As New Suspender(Me, True)
                Dim oDataGridSubjectsTags As Controls.DataGrid = Root.DataGridSubjectsTags
                Dim oEmptyList As New List(Of String)
                m_DataObject.Subjects.Add({String.Empty}, oDataGridSubjectsTags.SelectedIndex)
            End Using
        End Sub
        Public Sub RemoveRow()
            Using oSuspender As New Suspender(Me, True)
                Dim oDataGridSubjectsTags As Controls.DataGrid = Root.DataGridSubjectsTags
                If m_DataObject.Subjects.Count > 0 Then
                    m_DataObject.Subjects.Remove(oDataGridSubjectsTags.SelectedIndex)
                End If
            End Using
        End Sub
        Public Sub ImportSubjects()
            ' add a new subject sheet
            ' scans for the largest contiguous rectangular area starting from the top left
            Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
            oOpenFileDialog.FileName = String.Empty
            oOpenFileDialog.DefaultExt = "*.xlsx"
            oOpenFileDialog.Multiselect = False
            oOpenFileDialog.Filter = "Excel Spreadsheet|*.xlsx"
            oOpenFileDialog.Title = "Load Subjects From File"
            Dim result? As Boolean = oOpenFileDialog.ShowDialog()
            If result = True Then
                Using oSuspender As New Suspender(Me, True)
                    Dim oFileInfo As New IO.FileInfo(oOpenFileDialog.FileName)
                    Dim oExcelDocument As ClosedXML.Excel.XLWorkbook = Nothing
                    Try
                        oExcelDocument = New ClosedXML.Excel.XLWorkbook(oOpenFileDialog.FileName)
                    Catch ex1 As System.IO.FileFormatException
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oFileInfo.Name + ". File format error."))
                        Exit Sub
                    Catch ex2 As System.IO.IOException
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oFileInfo.Name + ". File access error."))
                        Exit Sub
                    End Try

                    m_DataObject.Subjects.Clear()

                    ' load only subject sheet
                    For Each oWorksheet As ClosedXML.Excel.IXLWorksheet In oExcelDocument.Worksheets
                        If Trim(oWorksheet.Name) = WorksheetSubjects Then
                            ' check contiguous area
                            Dim yLimit As Integer = Integer.MaxValue
                            Dim sReturn As String = "Tags"
                            Dim iCurrentRow As Integer = 1
                            Do Until sReturn = String.Empty
                                sReturn = Trim(oWorksheet.Cell(iCurrentRow, 1).GetString)
                                iCurrentRow += 1
                            Loop

                            yLimit = Math.Min(yLimit, iCurrentRow - 2)

                            For i = 1 To yLimit
                                Dim oRow As New List(Of String)
                                oRow.Add(Trim(oWorksheet.Cell(i, 1).GetString))
                                m_DataObject.Subjects.Add(oRow)
                            Next

                            Exit For
                        End If
                    Next

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Subjects loaded from file " + oFileInfo.Name + "."))
                End Using
            End If
        End Sub
        Public Sub ExportSubjects()
            ' exports in opendocument spreadsheet format
            ' add sheet to spreadsheet if saving over existing file
            Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
            oSaveFileDialog.FileName = String.Empty
            oSaveFileDialog.DefaultExt = "*.xlsx"
            oSaveFileDialog.Filter = "Excel Spreadsheet|*.xlsx"
            oSaveFileDialog.Title = "Save Subjects To File"
            oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
            Dim result? As Boolean = oSaveFileDialog.ShowDialog()
            If result = True Then
                Dim oExcelDocument As ClosedXML.Excel.XLWorkbook = Nothing
                If IO.File.Exists(oSaveFileDialog.FileName) Then
                    ' load existing file
                    Dim oCurrentFileInfo As New IO.FileInfo(oSaveFileDialog.FileName)
                    Try
                        oExcelDocument = New ClosedXML.Excel.XLWorkbook(oSaveFileDialog.FileName)
                    Catch ex1 As System.IO.FileFormatException
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oCurrentFileInfo.Name + ". File format error."))
                        Exit Sub
                    Catch ex2 As System.IO.IOException
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oCurrentFileInfo.Name + ". File access error."))
                        Exit Sub
                    End Try

                    ' delete subjects worksheet if found
                    For i = oExcelDocument.Worksheets.Count - 1 To 0 Step -1
                        If Trim(oExcelDocument.Worksheets(i).Name) = WorksheetSubjects Then
                            oExcelDocument.Worksheets.Delete(i + 1)
                        End If
                    Next
                Else
                    ' create new Excel document
                    oExcelDocument = New ClosedXML.Excel.XLWorkbook

                    For Each oCurrentWorksheet As ClosedXML.Excel.IXLWorksheet In oExcelDocument.Worksheets
                        oExcelDocument.Worksheets.Delete(oCurrentWorksheet.Name)
                    Next
                End If

                Dim oWorkSheet As ClosedXML.Excel.IXLWorksheet = oExcelDocument.Worksheets.Add(WorksheetSubjects, 0)

                ' set headers
                For i = 0 To m_DataObject.Subjects.Count - 1
                    oWorkSheet.Cell(1 + i, 1).Value = m_DataObject.Subjects.Tag(i, 0)
                Next
                oWorkSheet.Column(1).AdjustToContents()

                oExcelDocument.SaveAs(oSaveFileDialog.FileName)

                Dim oFileInfo As New IO.FileInfo(oSaveFileDialog.FileName)
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Tags saved to file " + oFileInfo.Name + "."))
            End If
        End Sub
        Public Sub SavePDF()
            ' cycles through all subject names and creates individualised PDFs with placement information for input fields
            Dim oFolderBrowserDialog As New Forms.FolderBrowserDialog
            oFolderBrowserDialog.Description = "Save to PDF by Subject Name"
            oFolderBrowserDialog.ShowNewFolderButton = True
            oFolderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop
            If oFolderBrowserDialog.ShowDialog = Forms.DialogResult.OK Then
                Dim oSubjectsList As List(Of String) = (From iSubjects In Enumerable.Range(0, m_DataObject.Subjects.Count) Select m_DataObject.Subjects.Tag(iSubjects, 0)).ToList
                SavePDFCommon(oFolderBrowserDialog.SelectedPath, oSubjectsList, "PDF files saved.", String.Empty)
            End If
        End Sub
        Public Sub ReplacePDF()
            ' updates the PDF and accompanying field info file while keeping the same subjects and bar codes
            Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
            oOpenFileDialog.FileName = String.Empty
            oOpenFileDialog.DefaultExt = "*.gz"
            oOpenFileDialog.Multiselect = False
            oOpenFileDialog.Filter = "GZip Files|*.gz"
            oOpenFileDialog.Title = "Update PDF File Content"
            Dim result? As Boolean = oOpenFileDialog.ShowDialog()
            If result = True Then
                Dim oNewFieldDocumentStore As FieldDocumentStore = CommonFunctions.DeserializeDataContractFile(Of FieldDocumentStore)(oOpenFileDialog.FileName)
                If IsNothing(oNewFieldDocumentStore) Then
                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Unable to load field definitions from file."))
                Else
                    Dim oSubjectsList As List(Of String) = (From oFieldCollection In oNewFieldDocumentStore.FieldCollectionStore Select oFieldCollection.SubjectName).ToList
                    Dim oFileInfo As New IO.FileInfo(oOpenFileDialog.FileName)
                    SavePDFCommon(oFileInfo.DirectoryName, oSubjectsList, "PDF files replaced.", String.Empty)
                End If
            End If
        End Sub
        Public Sub ExportHelp()
            ' exports help file
            ' this is in a format of a dictionary of GUID and byte arrays representing sections with tags that have valid GUIDs
            ' the byte arrays are PDF files generated from the relevant sections
            Dim oHelpDictionary As New HelpFile
            Dim oSections As List(Of FormSection) = FormMain.GetFormItems(Of FormSection)
            For Each oSection In oSections
                Dim oParseGUID As Guid = Guid.Empty
                Guid.TryParse(oSection.Tag, oParseGUID)

                ' valid guid found
                If Not oParseGUID.Equals(Guid.Empty) Then
                    ' renders the PDF of this section only
                    ParamList.Clear()

                    Dim XCurrentHeight As New XUnit(0)
                    Dim oFieldCollection As New FieldDocumentStore.FieldCollection
                    With oFieldCollection
                        .FormTitle = "Help File"
                        .FormAuthor = ModuleName
                        .FormTopic = FormProperties.FormTopic
                        .SubjectName = oParseGUID.ToString
                        .AppendText = String.Empty
                        .DateCreated = Date.Now
                    End With

                    Using oParamListFieldCollection As New ParamList(ParamList.KeyFieldCollection, oFieldCollection)
                        Using oParamListRemoveBarcode As New ParamList(ParamList.KeyRemoveBarcode, False)
                            Using oParamListPDFDocument As New ParamList(ParamList.KeyPDFDocument, New Pdf.PdfDocument)
                                Using oParamListPDFPages As New ParamList(ParamList.KeyPDFPages, New List(Of Tuple(Of Pdf.PdfPage, String, String, String)))
                                    Using oParamXCurrentHeight As New ParamList(ParamList.KeyXCurrentHeight, XCurrentHeight)
                                        Using oParamListSubjectName As New ParamList(ParamList.KeyCurrentSubject, oFieldCollection.SubjectName)
                                            ' insert logic for document info
                                            With ParamList.Value(ParamList.KeyPDFDocument)
                                                .Info.Title = oFieldCollection.FormTitle
                                                .Info.Author = oFieldCollection.FormAuthor
                                                .Info.Subject = oFieldCollection.SubjectName
                                            End With

                                            ' add sections
                                            Dim oMeasureHelp As Tuple(Of XUnit, XUnit) = Nothing
                                            Using oParamDisplacement As New ParamList(ParamList.KeyDisplacement, New XPoint(PDFHelper.PageLimitLeft.Point, PDFHelper.PageLimitTop.Point))
                                                ' add the first page
                                                Dim oPage As Pdf.PdfPage = ParamList.Value(ParamList.KeyPDFDocument).AddPage()
                                                oPage.Orientation = FormProperties.SelectedOrientation
                                                oPage.Size = FormProperties.SelectedSize
                                                ParamList.Value(ParamList.KeyPDFPages).Add(New Tuple(Of Pdf.PdfPage, String, String, String)(oPage, String.Empty, String.Empty, String.Empty))

                                                oMeasureHelp = oSection.MeasureHelpPDF()
                                            End Using

                                            Using oParamDisplacement As New ParamList(ParamList.KeyDisplacement, New XPoint(PDFHelper.BlockSpacer.Point, PDFHelper.BlockSpacer.Point))
                                                Dim oPage As Pdf.PdfPage = ParamList.Value(ParamList.KeyPDFDocument).AddPage()
                                                oPage.Orientation = FormProperties.SelectedOrientation
                                                oPage.Width = New XUnit(oMeasureHelp.Item1.Point + (PDFHelper.BlockSpacer.Point * 2))
                                                oPage.Height = New XUnit(oMeasureHelp.Item2.Point + (PDFHelper.BlockSpacer.Point * 2))
                                                ParamList.Value(ParamList.KeyPDFPages).Add(New Tuple(Of Pdf.PdfPage, String, String, String)(oPage, String.Empty, String.Empty, String.Empty))

                                                Using oParamListSuspendPagination As New ParamList(ParamList.KeySuspendPagination, True)
                                                    oSection.RenderPDF()
                                                End Using
                                            End Using

                                            ' save PDF
                                            Dim oPDFDocument As Pdf.PdfDocument = Nothing
                                            oParamListPDFDocument.DisposeReturn(oPDFDocument)
                                            Using oMemoryStream As New IO.MemoryStream
                                                oPDFDocument.Save(oMemoryStream)
                                                oPDFDocument.Close()
                                                oPDFDocument.Dispose()

                                                Dim bBytes As Byte() = oMemoryStream.ToArray
                                                If Not oHelpDictionary.Dictionary.ContainsKey(oParseGUID) Then
                                                    oHelpDictionary.Dictionary.Add(oParseGUID, bBytes)
                                                End If
                                            End Using
                                        End Using
                                    End Using
                                End Using
                            End Using
                        End Using
                    End Using
                End If
            Next

            ' save help file
            If oHelpDictionary.Dictionary.Count > 0 Then
                Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
                oSaveFileDialog.FileName = String.Empty
                oSaveFileDialog.DefaultExt = "*.shp"
                oSaveFileDialog.Filter = "Survey Help File|*.shp"
                oSaveFileDialog.Title = "Save Sections To Help File"
                oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
                Dim result? As Boolean = oSaveFileDialog.ShowDialog()
                If result = True Then
                    Dim sHelpFile As String = CommonFunctions.ReplaceExtension(oSaveFileDialog.FileName, "shp")
                    If IO.File.Exists(sHelpFile) Then
                        IO.File.Delete(sHelpFile)
                    End If

                    CommonFunctions.SerializeDataContractFile(sHelpFile, oHelpDictionary, New List(Of Type) From {GetType(HelpFile)}, False, "shp")
                    Dim oFileInfo As New IO.FileInfo(sHelpFile)

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Help file " + oFileInfo.Name + " saved."))
                End If
            Else
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "No valid sections to save."))
            End If
        End Sub
        Private Shared Sub SavePDFCommon(ByVal sSelectedPath As String, ByVal oSubjectsList As List(Of String), ByVal sMessage As String, ByVal sAppendText As String)
            ' common save routine
            Dim XCurrentHeight As XUnit = XUnit.Zero
            If oSubjectsList.Count > 0 Then
                Using oSuspender As New Suspender
                    Dim oFieldDocumentStore As New FieldDocumentStore
                    For iSubjects = 0 To oSubjectsList.Count - 1
                        XCurrentHeight = New XUnit(PDFHelper.PageLimitBottom.Point * 2)
                        Dim oFieldCollection As New FieldDocumentStore.FieldCollection
                        With oFieldCollection
                            .FormTitle = FormProperties.FormTitle
                            .FormAuthor = FormProperties.FormAuthor
                            .FormTopic = FormProperties.FormTopic
                            .SubjectName = oSubjectsList(iSubjects)
                            .AppendText = sAppendText
                            .DateCreated = Date.Now
                        End With

                        Using oParamListFieldCollection As New ParamList(ParamList.KeyFieldCollection, oFieldCollection)
                            Using oParamListRemoveBarcode As New ParamList(ParamList.KeyRemoveBarcode, False)
                                Using oParamListPDFDocument As New ParamList(ParamList.KeyPDFDocument, New Pdf.PdfDocument)
                                    Using oParamXCurrentHeight As New ParamList(ParamList.KeyXCurrentHeight, XCurrentHeight)
                                        Using oParamListSubjectName As New ParamList(ParamList.KeyCurrentSubject, oSubjectsList(iSubjects))
                                            FormPDF.RenderPDF()

                                            Dim sFileName As String = String.Empty
                                            If Trim(Converter.AlphaNumericOnly(FormProperties.FormTitle, False)) <> String.Empty Then
                                                sFileName += Trim(Converter.AlphaNumericOnly(FormProperties.FormTitle, False)) + "_"
                                            End If
                                            sFileName += iSubjects.ToString.PadLeft(Len(oSubjectsList.Count.ToString), "0") + "_" + ParamList.Value(ParamList.KeyCurrentSubject) + ".pdf"
                                            sFileName = sSelectedPath + "\" + CommonFunctions.SafeFileName(sFileName)

                                            Dim oPDFDocument As Pdf.PdfDocument = Nothing
                                            oParamListPDFDocument.DisposeReturn(oPDFDocument)
                                            oPDFDocument.Save(sFileName)
                                            oPDFDocument.Close()

                                            oFieldDocumentStore.FieldCollectionStore.Add(ParamList.Value(ParamList.KeyFieldCollection).Clone)
                                        End Using
                                    End Using
                                End Using
                            End Using
                        End Using
                    Next

                    ' get PDF document without bar codes
                    XCurrentHeight = New XUnit(PDFHelper.PageLimitBottom.Point * 2)
                    Using oParamListFieldCollection As New ParamList(ParamList.KeyFieldCollection, New FieldDocumentStore.FieldCollection)
                        Using oParamListRemoveBarcode As New ParamList(ParamList.KeyRemoveBarcode, True)
                            Using oParamListPDFDocument As New ParamList(ParamList.KeyPDFDocument, New Pdf.PdfDocument)
                                Using oParamXCurrentHeight As New ParamList(ParamList.KeyXCurrentHeight, XCurrentHeight)
                                    Using oParamListSubjectName As New ParamList(ParamList.KeyCurrentSubject, String.Empty)
                                        FormPDF.RenderPDF()

                                        Dim oPDFDocument As Pdf.PdfDocument = Nothing
                                        oParamListPDFDocument.DisposeReturn(oPDFDocument)
                                        Using oMemoryStream As New IO.MemoryStream
                                            oPDFDocument.Save(oMemoryStream)
                                            oPDFDocument.Close()
                                            oPDFDocument.Dispose()
                                            oFieldDocumentStore.PDFTemplate = oMemoryStream.ToArray
                                        End Using
                                    End Using
                                End Using
                            End Using
                        End Using
                    End Using

                    oFieldDocumentStore.SetMarks()
                    CommonFunctions.SerializeDataContractFile(sSelectedPath + "\FieldInfo.gz", oFieldDocumentStore, Root.LocalKnownTypes)

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, sMessage))
                End Using
            Else
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "No subjects. PDF files not saved."))
            End If
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FormSection
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass
        Public Shared ImageViewer As ImageViewerClass
        Public ImageTracker As ImageTrackerClass
        Public SectionResolution As Single = 0

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public SectionTitle As String
            <DataMember> Public Justification As Enumerations.JustificationEnum
            <DataMember> Public Font As Enumerations.FontEnum
            <DataMember> Public FontSizeMultiplier As Integer
            <DataMember> Public ContinuousNumbering As Boolean
            <DataMember> Public NumberingBorder As Boolean
            <DataMember> Public NumberingBackground As Boolean
            <DataMember> Public NumberingType As Enumerations.Numbering
            <DataMember> Public Start As Integer
            <DataMember> Public GridHeight As Integer
            <DataMember> Public Tag As String

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .SectionTitle = SectionTitle
                    .Justification = Justification
                    .Font = Font
                    .FontSizeMultiplier = FontSizeMultiplier
                    .ContinuousNumbering = ContinuousNumbering
                    .NumberingBorder = NumberingBorder
                    .NumberingBackground = NumberingBackground
                    .NumberingType = NumberingType
                    .Start = Start
                    .GridHeight = GridHeight
                    .Tag = Tag
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            CType(oFormItem, FormSection).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Private Const MaxGridHeight As Integer = 12
        Public Property SectionTitle As String
            Get
                Return m_DataObject.SectionTitle
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.SectionTitle = value
                    OnPropertyChangedLocal("SectionTitle")
                End Using
            End Set
        End Property
        Public Property Start As Integer
            Get
                Return m_DataObject.Start
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True, True)
                    If value < 0 Then
                        m_DataObject.Start = 0
                    Else
                        m_DataObject.Start = value
                    End If
                    OnPropertyChangedLocal("Start")
                    OnPropertyChangedLocal("StartText")
                End Using
            End Set
        End Property
        Public Property StartText As String
            Get
                Return (m_DataObject.Start + 1).ToString
            End Get
            Set(value As String)
                Start = CInt(Val(value)) - 1
            End Set
        End Property
        Public Property NumberingType As Enumerations.Numbering
            Get
                Return m_DataObject.NumberingType
            End Get
            Set(value As Enumerations.Numbering)
                Using oSuspender As New Suspender(Me, True, True)
                    m_DataObject.NumberingType = value
                    OnPropertyChangedLocal("NumberingType")
                    OnPropertyChangedLocal("NumberingTypeText")
                End Using
            End Set
        End Property
        Public ReadOnly Property NumberingTypeText As String
            Get
                Select Case m_DataObject.NumberingType
                    Case Enumerations.Numbering.Number
                        Return "Number"
                    Case Enumerations.Numbering.LetterSmall
                        Return "Lower Case Letter"
                    Case Enumerations.Numbering.LetterBig
                        Return "Upper Case Letter"
                    Case Else
                        Return String.Empty
                End Select
            End Get
        End Property
        Public Property GridHeight As Integer
            Get
                Return m_DataObject.GridHeight
            End Get
            Set(value As Integer)
                Dim iOldGridHeight As Integer = m_DataObject.GridHeight

                If value < 1 Then
                    m_DataObject.GridHeight = 1
                ElseIf value > MaxGridHeight Then
                    m_DataObject.GridHeight = MaxGridHeight
                Else
                    m_DataObject.GridHeight = value
                End If

                If iOldGridHeight <> m_DataObject.GridHeight Then
                    Using oSuspender As New Suspender(Me, True)
                        OnPropertyChangedLocal("GridHeight")
                        OnPropertyChangedLocal("GridHeightText")
                    End Using
                End If
            End Set
        End Property
        Public Property GridHeightText As String
            Get
                Return m_DataObject.GridHeight.ToString
            End Get
            Set(value As String)
                GridHeight = CInt(Val(value))
            End Set
        End Property
        Public Property Tag As String
            Get
                Return m_DataObject.Tag
            End Get
            Set(value As String)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Tag = Trim(value)
                    OnPropertyChangedLocal("Tag")
                End Using
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "Title: " + SectionTitle
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .SectionTitle = String.Empty
                .Justification = Enumerations.JustificationEnum.Left
                .Font = Enumerations.FontEnum.None
                .FontSizeMultiplier = 1
                .ContinuousNumbering = True
                .NumberingBorder = False
                .NumberingBackground = False
                .NumberingType = Enumerations.Numbering.Number
                .Start = 0
                .GridHeight = 1
            End With
            ImageTracker = New ImageTrackerClass
            TitleChanged()
        End Sub
#End Region
#Region "Justification"
        Public Property Justification As Enumerations.JustificationEnum
            Get
                Return m_DataObject.Justification
            End Get
            Set(value As Enumerations.JustificationEnum)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Justification = value
                    OnPropertyChangedLocal("Justification")
                    OnPropertyChangedLocal("JustificationLeft")
                    OnPropertyChangedLocal("JustificationCenter")
                    OnPropertyChangedLocal("JustificationRight")
                    OnPropertyChangedLocal("JustificationJustify")
                End Using
            End Set
        End Property
        Public ReadOnly Property JustificationLeft As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Left Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property JustificationCenter As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Center Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property JustificationRight As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Right Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property JustificationJustify As Boolean
            Get
                If m_DataObject.Justification = Enumerations.JustificationEnum.Justify Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
#End Region
#Region "Font"
        Private Const MaxFontSizeMultiplier As Integer = 4

        Public Property Font As Enumerations.FontEnum
            Get
                Return m_DataObject.Font
            End Get
            Set(value As Enumerations.FontEnum)
                Using oSuspender As New Suspender(Me, True)
                    m_DataObject.Font = m_DataObject.Font Xor value
                    OnPropertyChangedLocal("Font")
                    OnPropertyChangedLocal("FontBold")
                    OnPropertyChangedLocal("FontItalic")
                    OnPropertyChangedLocal("FontUnderline")
                End Using
            End Set
        End Property
        Public Property FontSizeMultiplier As Integer
            Get
                Return m_DataObject.FontSizeMultiplier
            End Get
            Set(value As Integer)
                Using oSuspender As New Suspender(Me, True)
                    If value < 1 Then
                        m_DataObject.FontSizeMultiplier = 1
                    ElseIf value > MaxFontSizeMultiplier Then
                        m_DataObject.FontSizeMultiplier = MaxFontSizeMultiplier
                    Else
                        m_DataObject.FontSizeMultiplier = value
                    End If
                    OnPropertyChangedLocal("FontSizeMultiplier")
                    OnPropertyChangedLocal("FontSizeMultiplierText")
                End Using
            End Set
        End Property
        Public Property FontSizeMultiplierText As String
            Get
                Return m_DataObject.FontSizeMultiplier.ToString
            End Get
            Set(value As String)
                FontSizeMultiplier = CInt(Val(value))
            End Set
        End Property
        Public ReadOnly Property FontBold As Boolean
            Get
                If (m_DataObject.Font And Enumerations.FontEnum.Bold).Equals(Enumerations.FontEnum.Bold) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property FontItalic As Boolean
            Get
                If (m_DataObject.Font And Enumerations.FontEnum.Italic).Equals(Enumerations.FontEnum.Italic) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property FontUnderline As Boolean
            Get
                If (m_DataObject.Font And Enumerations.FontEnum.Underline).Equals(Enumerations.FontEnum.Underline) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
#End Region
#Region "Numbering"
        Public Property ContinuousNumbering As Boolean
            Get
                Return m_DataObject.ContinuousNumbering
            End Get
            Set(value As Boolean)
                Using oSuspender As New Suspender(Me, True, True)
                    m_DataObject.ContinuousNumbering = value
                    OnPropertyChangedLocal("ContinuousNumbering")
                End Using
            End Set
        End Property
        Public Property NumberingBorder As Boolean
            Get
                Return m_DataObject.NumberingBorder
            End Get
            Set(value As Boolean)
                If m_DataObject.NumberingBorder <> value Then
                    Using oSuspender As New Suspender(Me, True, True)
                        m_DataObject.NumberingBorder = value
                        OnPropertyChangedLocal("NumberingBorder")
                    End Using
                End If
            End Set
        End Property
        Public Property NumberingBackground As Boolean
            Get
                Return m_DataObject.NumberingBackground
            End Get
            Set(value As Boolean)
                If m_DataObject.NumberingBackground <> value Then
                    Using oSuspender As New Suspender(Me, True, True)
                        m_DataObject.NumberingBackground = value
                        OnPropertyChangedLocal("NumberingBackground")
                    End Using
                End If
            End Set
        End Property
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Section"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.0
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMSection"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelBlock")
                oDisplayFilter.Add("StackPanelSection")
                oDisplayFilter.Add("StackPanelSectionExtra")
                oDisplayFilter.Add("StackPanelBlockRow")
                oDisplayFilter.Add("StackPanelSectionTitle")
                oDisplayFilter.Add("StackPanelJustification")
                oDisplayFilter.Add("StackPanelFont")
                oDisplayFilter.Add("BlockStart")
                oDisplayFilter.Add("BlockType")
                oDisplayFilter.Add("BlockRow")
                oDisplayFilter.Add("SectionTitle")
                oDisplayFilter.Add("FontSizeMultiplier")
                oDisplayFilter.Add("ImageBlockViewer")
                oDisplayFilter.Add("ScrollViewerBlockContent")
                If oSettings.DeveloperMode Then
                    oDisplayFilter.Add("StackPanelSectionTag")
                    oDisplayFilter.Add("SectionTag")
                End If
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.StackPanelSection.DataContext = Me
            Root.StackPanelSectionExtra.DataContext = Me
            Root.StackPanelBlockRow.DataContext = Me
            Root.StackPanelSectionTitle.DataContext = Me
            Root.StackPanelSectionTag.DataContext = Me
            Root.ImageBlockViewer.DataContext = ImageViewer

            Dim oBindingSection1 As New Data.Binding
            oBindingSection1.Path = New PropertyPath("NumberingBorder")
            oBindingSection1.Mode = Data.BindingMode.TwoWay
            oBindingSection1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockNumberingBorder.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingSection1)

            Dim oBindingSection2 As New Data.Binding
            oBindingSection2.Path = New PropertyPath("NumberingBackground")
            oBindingSection2.Mode = Data.BindingMode.TwoWay
            oBindingSection2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockNumberingBackground.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingSection2)

            Dim oBindingSection3 As New Data.Binding
            oBindingSection3.Path = New PropertyPath("StartText")
            oBindingSection3.Mode = Data.BindingMode.TwoWay
            oBindingSection3.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockStart.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingSection3)

            Dim oBindingSection4 As New Data.Binding
            oBindingSection4.Path = New PropertyPath("NumberingTypeText")
            oBindingSection4.Mode = Data.BindingMode.OneWay
            oBindingSection4.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockType.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingSection4)

            Dim oBindingSection5 As New Data.Binding
            oBindingSection5.Path = New PropertyPath("ContinuousNumbering")
            oBindingSection5.Mode = Data.BindingMode.TwoWay
            oBindingSection5.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockContinuousNumbering.SetBinding(Common.HighlightCheckBox.HCBCheckedProperty, oBindingSection5)

            Dim oBindingSection6 As New Data.Binding
            oBindingSection6.Path = New PropertyPath("GridHeightText")
            oBindingSection6.Mode = Data.BindingMode.TwoWay
            oBindingSection6.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockRow.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingSection6)

            Dim oBindingSection7 As New Data.Binding
            oBindingSection7.Path = New PropertyPath("SectionTitle")
            oBindingSection7.Mode = Data.BindingMode.TwoWay
            oBindingSection7.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.SectionTitle.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingSection7)

            Dim oBindingSection8 As New Data.Binding
            oBindingSection8.Path = New PropertyPath("Tag")
            oBindingSection8.Mode = Data.BindingMode.TwoWay
            oBindingSection8.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.SectionTag.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingSection8)

            FieldText.BindingJustification(Me)
            FieldText.BindingFont(Me)
        End Sub
        Public Overrides Sub Display()
            Dim oScrollViewerBlockContent As Controls.ScrollViewer = Root.ScrollViewerBlockContent
            Dim oScrollBar As Controls.Primitives.ScrollBar = CType(oScrollViewerBlockContent.Template.FindName("PART_VerticalScrollBar", oScrollViewerBlockContent), Controls.Primitives.ScrollBar)

            If Not IsNothing(oScrollBar) Then
                CheckInvalidate()
                ImageViewer.Clear()

                ' add section title
                Dim oSectionTitleBitmap As ImageSource = ImageTracker.GetImage(ImageTrackerClass.GetList({Me}))
                If IsNothing(oSectionTitleBitmap) Then
                    oSectionTitleBitmap = GetSectionTitleBitmapSource(SectionResolution).Item1
                    ImageTracker.Add(oSectionTitleBitmap, ImageTrackerClass.GetList({Me}))
                End If
                ImageViewer.Add(oSectionTitleBitmap, "Section Title")

                ' reset current numbering to zero
                Using oParamListNumberingCurrent As New ParamList(ParamList.KeyNumberingCurrent, Start)
                    RenderCommon(False)
                End Using
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "d7c5c68a-efaa-4931-932b-6293af5f73a1"
            End Get
        End Property
        Private Sub RenderCommon(ByVal bRender As Boolean)
            ' common rendering
            ' set current section
            Using oParamListCurrentSection As New ParamList(ParamList.KeyCurrentSection, Me)
                ' runs through children
                Dim oLineBlockList As New List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize)))
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
                    If bRender OrElse (Not (oChild.Value.IsSubclassOf(GetType(FormFormatter)) OrElse oChild.Value.IsSubclassOf(GetType(FormField)))) Then
                        If oChild.Value.Equals(GetType(FormBlock)) Then
                            ' if the block above is a divider, then output current line and reset block list
                            Dim oBlock As FormBlock = FormMain.FindChild(oChild.Key)
                            Dim iItemIndex As Integer = oChildDictionary.Keys.ToList.IndexOf(oChild.Key)
                            If oLineBlockList.Count > 0 AndAlso oChildDictionary.Values(iItemIndex - 1).Equals(GetType(FormatterDivider)) Then
                                CheckDivider(oLineBlockList, oChildDictionary, bRender)
                                OutputCurrentLine(bRender, oLineBlockList)

                                ' add to block list
                                Dim oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = oBlock.GetBlockDimensions(True)
                                oLineBlockList.Add(New Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))(oBlock, oBlockDimensions, Nothing, oBlock.GetExpandedBlockDimensions(oBlockDimensions)))
                            Else
                                Dim iBlockCount As Integer = Aggregate oLineBlock In oLineBlockList Into Sum(oLineBlock.Item1.EffectiveBlockWidth)
                                If iBlockCount + oBlock.EffectiveBlockWidth > PDFHelper.PageBlockWidth Then
                                    ' output current line and reset block list
                                    CheckDivider(oLineBlockList, oChildDictionary, bRender)
                                    OutputCurrentLine(bRender, oLineBlockList)
                                End If

                                ' add to block list
                                Dim oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = oBlock.GetBlockDimensions(True)
                                oLineBlockList.Add(New Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))(oBlock, oBlockDimensions, Nothing, oBlock.GetExpandedBlockDimensions(oBlockDimensions)))
                            End If
                        Else
                            ' output current line and reset block list
                            CheckDivider(oLineBlockList, oChildDictionary, bRender)
                            OutputCurrentLine(bRender, oLineBlockList)

                            Select Case oChild.Value
                                Case GetType(FormSubSection)
                                    Dim oSubSection As FormSubSection = FormMain.FindChild(oChild.Key)
                                    ParamList.Value(ParamList.KeyNumberingCurrent) = GetNumbering(oSubSection)

                                    If ContinuousNumbering Then
                                        CheckDivider(oChild, oChildDictionary, bRender)
                                        If bRender Then
                                            oSubSection.RenderPDF()
                                        Else
                                            oSubSection.DisplayPDF()
                                        End If
                                    Else
                                        Using oParamListSubNumberingCurrent As New ParamList(ParamList.KeySubNumberingCurrent, 0)
                                            CheckDivider(oChild, oChildDictionary, bRender)
                                            If bRender Then
                                                oSubSection.RenderPDF()
                                            Else
                                                oSubSection.DisplayPDF()
                                            End If
                                        End Using
                                    End If
                                Case GetType(FormMCQ)
                                    Dim oMCQ As FormMCQ = FormMain.FindChild(oChild.Key)
                                    ParamList.Value(ParamList.KeyNumberingCurrent) = GetNumbering(oMCQ)

                                    If ContinuousNumbering Then
                                        CheckDivider(oChild, oChildDictionary, bRender)
                                        If bRender Then
                                            oMCQ.RenderPDF()
                                        Else
                                            oMCQ.DisplayPDF()
                                        End If
                                    Else
                                        Using oParamListSubNumberingCurrent As New ParamList(ParamList.KeySubNumberingCurrent, 0)
                                            CheckDivider(oChild, oChildDictionary, bRender)
                                            If bRender Then
                                                oMCQ.RenderPDF()
                                            Else
                                                oMCQ.DisplayPDF()
                                            End If
                                        End Using
                                    End If
                                Case GetType(FormatterPageBreak)
                                    Dim bPaginate As Boolean = True
                                    If ParamList.ContainsKey(ParamList.KeySuspendPagination) AndAlso ParamList.Value(ParamList.KeySuspendPagination) Then
                                        bPaginate = False
                                    End If

                                    ' paginate only if not suspended
                                    If bPaginate Then
                                        ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(PDFHelper.PageLimitBottomExBarcode.Point + 1)
                                        FormPDF.AddPage(0)
                                    End If
                            End Select
                        End If
                    End If
                Next

                ' output current line and reset block list
                CheckDivider(oLineBlockList, oChildDictionary, bRender)
                OutputCurrentLine(bRender, oLineBlockList)
            End Using
        End Sub
        Private Sub OutputCurrentLine(ByVal bRender As Boolean, ByRef oLineBlockList As List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))))
            ' outputs the current line
            If bRender Then
                RenderCurrentLine(oLineBlockList)
            Else
                DisplayCurrentLine(oLineBlockList)
            End If
        End Sub
        Private Shared Sub CheckDivider(oChild As KeyValuePair(Of Guid, Type), ByVal oChildDictionary As Dictionary(Of Guid, Type), ByVal bRender As Boolean)
            ' check for divider setting
            Dim iItemIndex As Integer = oChildDictionary.Keys.ToList.IndexOf(oChild.Key)
            ' set divider with no line if anything else
            Dim bDividerLine As Boolean = False
            Dim oFormItem As BaseFormItem = FormMain.FindChild(oChild.Key)
            ' set line divider if:
            ' 1) the child item is an MCQ or subsection with a divider
            If (oChild.Value.Equals(GetType(FormMCQ)) Or oChild.Value.Equals(GetType(FormSubSection))) AndAlso oFormItem.Children.ContainsValue(GetType(FormatterDivider)) Then
                bDividerLine = True
            End If

            ' not the first item in the list
            If iItemIndex > 0 Then
                Dim oFormItemAbove As BaseFormItem = FormMain.FindChild(oChildDictionary.Keys(iItemIndex - 1))
                ' and the item above is not a border or background
                If Not (oFormItemAbove.GetType.Equals(GetType(FieldBorder)) Or oFormItemAbove.GetType.Equals(GetType(FieldBackground))) Then
                    ' 2) the child item above is a divider
                    If oFormItemAbove.GetType.Equals(GetType(FormatterDivider)) Then
                        bDividerLine = True
                    End If

                    ' 3) the child item above is an MCQ or subsection with a divider
                    If (oFormItemAbove.GetType.Equals(GetType(FormMCQ)) Or oFormItemAbove.GetType.Equals(GetType(FormSubSection))) AndAlso oFormItemAbove.Children.ContainsValue(GetType(FormatterDivider)) Then
                        bDividerLine = True
                    End If
                End If
            End If

            ' sets the divider type
            If bRender Then
                SetDivider(bDividerLine)
            Else
                If bDividerLine Then
                    ImageViewer.Add(PDFHelper.SpacerBitmapSourceLine.Item1, "Spacer With Line")
                Else
                    ImageViewer.Add(PDFHelper.SpacerBitmapSourceNoLine.Item1, "Spacer Without Line")
                End If
            End If
        End Sub
        Private Shared Sub CheckDivider(ByVal oLineBlockList As List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))), ByVal oChildDictionary As Dictionary(Of Guid, Type), ByVal bRender As Boolean)
            If oLineBlockList.Count > 0 Then
                Dim oChild As KeyValuePair(Of Guid, Type) = oChildDictionary.ElementAt(oChildDictionary.Keys.ToList.IndexOf(oLineBlockList.First.Item1.GUID))
                CheckDivider(oChild, oChildDictionary, bRender)
            End If
        End Sub
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormSubSection), GetType(FormMCQ), GetType(FormBlock), GetType(FieldBorder), GetType(FieldBackground), GetType(FormatterPageBreak), GetType(FormatterDivider), GetType(FormatterGroup)}
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            Dim XSectionSize As XSize = GetSectionTitleDimensions()

            FormPDF.AddPage(XSectionSize.Height)
            Dim XDisplacement As New XPoint(ParamList.Value(ParamList.KeyDisplacement).X, ParamList.Value(ParamList.KeyDisplacement).Y + ParamList.Value(ParamList.KeyXCurrentHeight).Point)

            ' add section title
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(ParamList.Value(ParamList.KeyPDFPages)(ParamList.Value(ParamList.KeyPDFPages).Count - 1).Item1, XGraphicsPdfPageOptions.Append)
                DrawSectionTitleDirect(oXGraphics, New XUnit(XSectionSize.Width), New XUnit(XSectionSize.Height), XDisplacement)
            End Using
            ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(ParamList.Value(ParamList.KeyXCurrentHeight).Point + XSectionSize.Height)

            Using oParamListNumberingCurrent As New ParamList(ParamList.KeyNumberingCurrent, Start)
                RenderCommon(True)
            End Using
        End Sub
        Public Function MeasureHelpPDF() As Tuple(Of XUnit, XUnit)
            ' measures the rendered section for the help file
            ' returned values are without borders
            ' store orientation of first page
            Dim fStartHeight As Double = ParamList.Value(ParamList.KeyXCurrentHeight).Point

            Using oParamListSuspendPagination As New ParamList(ParamList.KeySuspendPagination, True)
                RenderPDF()
            End Using

            Dim fEndHeight As Double = ParamList.Value(ParamList.KeyXCurrentHeight).Point

            ' reset pdf document
            If ParamList.Value(ParamList.KeyPDFDocument).PageCount > 0 Then
                For i = ParamList.Value(ParamList.KeyPDFDocument).PageCount - 1 To 0 Step -1
                    ParamList.Value(ParamList.KeyPDFDocument).Pages.RemoveAt(i)
                Next
            End If
            ParamList.Value(ParamList.KeyPDFPages).Clear

            ' reset current height
            ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(fStartHeight)

            Return New Tuple(Of XUnit, XUnit)(PDFHelper.PageLimitWidth, New XUnit(fEndHeight - fStartHeight))
        End Function
        Public Overrides ReadOnly Property FormattingFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterPageBreak), GetType(FormatterDivider), GetType(FormatterGroup)}
            End Get
        End Property
        Public Function GetSectionTitleBitmapSource(ByVal fResolution As Single) As Tuple(Of Imaging.BitmapSource, XUnit, XUnit)
            ' gets bitmap for the section title
            Dim XSectionSize As XSize = GetSectionTitleDimensions()
            Dim XSectionWidth As XUnit = XSectionSize.Width
            Dim XSectionHeight As XUnit = XSectionSize.Height


            Dim oBitmapSource As Imaging.BitmapSource = Nothing
            Using oBitmap As New System.Drawing.Bitmap(CInt(Math.Ceiling(XSectionWidth.Inch * fResolution)), CInt(Math.Ceiling(XSectionHeight.Inch * fResolution)), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                oBitmap.SetResolution(fResolution, fResolution)

                Dim oParamList As New Dictionary(Of String, Object)
                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XSectionSize, XGraphicsUnit.Point)
                        Dim oBlockReturn = DrawSectionTitleDirect(oXGraphics, XSectionWidth, XSectionHeight, New XPoint(0, 0)).Item1
                    End Using
                End Using

                oBitmapSource = Converter.BitmapToBitmapSource(oBitmap)
            End Using

            Return New Tuple(Of Imaging.BitmapSource, XUnit, XUnit)(oBitmapSource, XSectionWidth, XSectionHeight)
        End Function
        Public Function GetSectionTitleDimensions() As XSize
            ' get dimensions of item
            Dim XSectionWidth As XUnit = PDFHelper.PageLimitWidth.Point
            Dim XSectionHeight As XUnit = XUnit.FromPoint((GridHeight * PDFHelper.BlockHeight.Point) + (PDFHelper.BlockSpacer.Point * 2))
            Return New XSize(XSectionWidth.Point, XSectionHeight.Point)
        End Function
        Public Function DrawSectionTitleDirect(ByRef oXGraphics As XGraphics, ByVal XSectionWidth As XUnit, XSectionHeight As XUnit, ByVal XDisplacement As XPoint) As Tuple(Of XUnit, XUnit)
            ' draw directly on the supplied XGraphics with the specified displacement
            Dim oContainer As XGraphicsContainer = Nothing

            ' extract fields
            Dim oBackgroundFields As List(Of FieldBackground) = ExtractFields(Of FieldBackground)(GetType(FormField))
            Dim oBorderFields As List(Of FieldBorder) = ExtractFields(Of FieldBorder)(GetType(FormField))

            ' draw background
            oContainer = oXGraphics.BeginContainer()
            Dim oFieldBackgroundWholeBlockList As List(Of FieldBackground) = (From oBackground In oBackgroundFields Where oBackground.WholeBlock Select oBackground).ToList
            If oFieldBackgroundWholeBlockList.Count > 0 Then
                Select Case oFieldBackgroundWholeBlockList.First.Lightness
                    Case 1
                        PDFHelper.DrawFieldBackground(oXGraphics, XSectionWidth, XSectionHeight, XDisplacement, XBrushes.WhiteSmoke)
                    Case 2
                        PDFHelper.DrawFieldBackground(oXGraphics, XSectionWidth, XSectionHeight, XDisplacement, XBrushes.LightGray)
                    Case 3
                        PDFHelper.DrawFieldBackground(oXGraphics, XSectionWidth, XSectionHeight, XDisplacement, XBrushes.Silver)
                    Case 4
                        PDFHelper.DrawFieldBackground(oXGraphics, XSectionWidth, XSectionHeight, XDisplacement, XBrushes.DarkGray)
                End Select
            Else
                PDFHelper.DrawFieldBackground(oXGraphics, XSectionWidth, XSectionHeight, XDisplacement, XBrushes.White)
            End If
            oXGraphics.EndContainer(oContainer)

            ' set section title
            oContainer = oXGraphics.BeginContainer()

            Dim sSectionFullTitle As String = String.Empty

            Dim oParseGUID As Guid = Guid.Empty
            Guid.TryParse(Tag, oParseGUID)
            If oParseGUID.Equals(Guid.Empty) Then
                sSectionFullTitle = "Section " + Numbering.ToString
                If Trim(SectionTitle) <> String.Empty Then
                    sSectionFullTitle += " - " + Trim(SectionTitle)
                End If
            Else
                sSectionFullTitle = Trim(SectionTitle)
            End If

            Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
            Dim oJustification As MigraDoc.DocumentObjectModel.ParagraphAlignment = Nothing
            Select Case Justification
                Case Enumerations.JustificationEnum.Left
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left
                Case Enumerations.JustificationEnum.Center
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Center
                Case Enumerations.JustificationEnum.Right
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Right
                Case Enumerations.JustificationEnum.Justify
                    oJustification = MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify
            End Select

            Dim oElements As New List(Of ElementStruc)
            oElements.Add(New ElementStruc(sSectionFullTitle, ElementStruc.ElementTypeEnum.Text, FontBold, FontItalic, FontUnderline))
            Dim xImageWidth As XUnit = XUnit.FromPoint(XSectionWidth.Point - (PDFHelper.BlockSpacer.Point * 2))
            Dim xImageHeight As XUnit = XUnit.FromPoint(GridHeight * PDFHelper.BlockHeight.Point)
            Dim oXRect As New XRect(PDFHelper.BlockSpacer.Point + XDisplacement.X, PDFHelper.BlockSpacer.Point + XDisplacement.Y, xImageWidth.Point, xImageHeight.Point)
            oXGraphics.IntersectClip(oXRect)

            PDFHelper.DrawFieldText(oXGraphics, xImageWidth, xImageHeight, New XPoint(((XSectionWidth.Point - xImageWidth.Point) / 2) + +XDisplacement.X, ((XSectionHeight.Point - xImageHeight.Point) / 2) + XDisplacement.Y), GridHeight, FontSizeMultiplier, oElements, -1, oJustification, oFontColourDictionary)
            oXGraphics.EndContainer(oContainer)

            ' draw border
            oContainer = oXGraphics.BeginContainer()
            Dim oFieldBorderWholeBlockList As List(Of FieldBorder) = (From oBorder In oBorderFields Where oBorder.WholeBlock Select oBorder).ToList
            If oFieldBorderWholeBlockList.Count > 0 Then
                PDFHelper.DrawFieldBorder(oXGraphics, XSectionWidth, XSectionHeight, XDisplacement, oFieldBorderWholeBlockList.First.BorderWidth, 0, 0, Nothing)
            End If
            oXGraphics.EndContainer(oContainer)

            Return New Tuple(Of XUnit, XUnit)(XSectionWidth, XSectionHeight)
        End Function
        Public Shared Sub RenderCurrentLine(ByRef oLineBlockList As List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))))
            ' renders current line
            Dim oContainer As XGraphicsContainer = Nothing
            If oLineBlockList.Count > 0 Then
                Dim oSection As FormSection = ParamList.Value(ParamList.KeyCurrentSection)
                Dim fFormItemHeight As Double = Aggregate oLineBlock In oLineBlockList Into Max(oLineBlock.Item4.Item2.Point)
                FormPDF.AddPage(fFormItemHeight)

                Dim fCurrentDisplacementX As Double = ParamList.Value(ParamList.KeyDisplacement).X
                For Each oCurrentBlock In oLineBlockList
                    ' do not change numbering if parent is a subsection
                    If Not oCurrentBlock.Item1.Parent.GetType.Equals(GetType(FormSubSection)) Then
                        ParamList.Value(ParamList.KeyNumberingCurrent) = oSection.GetNumbering(oCurrentBlock.Item1)
                    End If

                    'add text array
                    Using oParamListTextArray As New ParamList(ParamList.KeyTextArray, oCurrentBlock.Item3)
                        Dim XBlockWidth As XUnit = oCurrentBlock.Item2.Item1
                        Dim XBlockHeight As XUnit = oCurrentBlock.Item2.Item2

                        Dim XExpandedBlockWidth As XUnit = oCurrentBlock.Item4.Item1
                        Dim XExpandedBlockHeight As XUnit = oCurrentBlock.Item4.Item2
                        Dim XContentDisplacement As XPoint = oCurrentBlock.Item1.GetContentDisplacement()

                        Using oXGraphics As XGraphics = XGraphics.FromPdfPage(ParamList.Value(ParamList.KeyPDFPages)(ParamList.Value(ParamList.KeyPDFPages).Count - 1).Item1, XGraphicsPdfPageOptions.Append)
                            oContainer = oXGraphics.BeginContainer()
                            oCurrentBlock.Item1.DrawBlockDirect(True, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, New XPoint(fCurrentDisplacementX, ParamList.Value(ParamList.KeyDisplacement).Y + ParamList.Value(ParamList.KeyXCurrentHeight).Point), XBlockWidth, XBlockHeight, XContentDisplacement, oCurrentBlock.Item2.Item3)
                            oXGraphics.EndContainer(oContainer)
                        End Using
                    End Using

                    If oCurrentBlock.Item1.Parent.GetType.Equals(GetType(FormSubSection)) Then
                        If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                            ' not continuous numbering
                            ParamList.Value(ParamList.KeySubNumberingCurrent) += 1
                        Else
                            ' continuous numbering
                            ParamList.Value(ParamList.KeyNumberingCurrent) += 1
                        End If
                    End If

                    ' increment displacement
                    fCurrentDisplacementX += oCurrentBlock.Item2.Item1.Point + (PDFHelper.BlockSpacer.Point * 3)
                Next

                ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(ParamList.Value(ParamList.KeyXCurrentHeight).Point + fFormItemHeight)
                oLineBlockList.Clear()
            End If
        End Sub
        Public Sub DisplayCurrentLine(ByRef oLineBlockList As List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))), Optional ByRef oBitmapStore As List(Of Imaging.BitmapSource) = Nothing)
            ' renders current line to a bitmapsource
            If oLineBlockList.Count > 0 Then
                Dim oLineImage As Imaging.BitmapSource = ImageTracker.GetImage(ImageTrackerClass.GetList((From oLineBlock In oLineBlockList Select oLineBlock.Item1).Cast(Of BaseFormItem).ToArray))
                If IsNothing(oLineImage) Then
                    Dim XLineWidth As New XUnit(PDFHelper.PageLimitWidth.Point)
                    Dim XLineHeight As New XUnit(Aggregate oLineBlock In oLineBlockList Into Max(oLineBlock.Item4.Item2.Point))
                    Dim XLineSize As New XSize(XLineWidth.Point, XLineHeight.Point)

                    Dim fCurrentDisplacementX As Double = 0
                    Using oBitmap As New System.Drawing.Bitmap(Math.Ceiling(XLineWidth.Inch * SectionResolution), Math.Ceiling(XLineHeight.Inch * SectionResolution), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                        oBitmap.SetResolution(SectionResolution, SectionResolution)

                        Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                            oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                            oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                            oGraphics.FillRectangle(System.Drawing.Brushes.White, 0, 0, oBitmap.Width, oBitmap.Height)
                            Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XLineSize, XGraphicsUnit.Point)
                                For Each oCurrentBlock In oLineBlockList
                                    ' do not change numbering if parent is a subsection
                                    If Not oCurrentBlock.Item1.Parent.GetType.Equals(GetType(FormSubSection)) Then
                                        ParamList.Value(ParamList.KeyNumberingCurrent) = GetNumbering(oCurrentBlock.Item1)
                                    End If

                                    'add text array
                                    Using oParamListTextArray As New ParamList(ParamList.KeyTextArray, oCurrentBlock.Item3)
                                        Dim XBlockWidth As XUnit = oCurrentBlock.Item2.Item1
                                        Dim XBlockHeight As XUnit = oCurrentBlock.Item2.Item2

                                        Dim XExpandedBlockWidth As XUnit = oCurrentBlock.Item4.Item1
                                        Dim XExpandedBlockHeight As XUnit = oCurrentBlock.Item4.Item2
                                        Dim XContentDisplacement As XPoint = oCurrentBlock.Item1.GetContentDisplacement()

                                        oCurrentBlock.Item1.DrawBlockDirect(True, oXGraphics, XExpandedBlockWidth, XExpandedBlockHeight, New XPoint(fCurrentDisplacementX, 0), XBlockWidth, XBlockHeight, XContentDisplacement, oCurrentBlock.Item2.Item3)
                                    End Using

                                    If oCurrentBlock.Item1.Parent.GetType.Equals(GetType(FormSubSection)) Then
                                        If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                                            ' not continuous numbering
                                            ParamList.Value(ParamList.KeySubNumberingCurrent) += 1
                                        Else
                                            ' continuous numbering
                                            ParamList.Value(ParamList.KeyNumberingCurrent) += 1
                                        End If
                                    End If

                                    ' increment displacement
                                    fCurrentDisplacementX += oCurrentBlock.Item2.Item1.Point + (PDFHelper.BlockSpacer.Point * 3)
                                Next
                            End Using
                        End Using

                        oLineImage = Converter.BitmapToBitmapSource(oBitmap)

                        If IsNothing(oBitmapStore) Then
                            ImageTracker.Add(oLineImage, ImageTrackerClass.GetList((From oLineBlock In oLineBlockList Select oLineBlock.Item1).Cast(Of BaseFormItem).ToArray))
                        Else
                            oBitmapStore.Add(oLineImage)
                        End If
                    End Using
                End If

                oLineBlockList.Clear()
                If IsNothing(oBitmapStore) Then
                    ImageViewer.Add(oLineImage, "Block Line")
                End If
            End If
        End Sub
        Private Shared Sub IncrementCurrentLine(ByVal oPage As Pdf.PdfPage, ByVal XCurrentHeight As XUnit)
            Dim XBlockWidth As XUnit = XUnit.FromPoint(PDFHelper.PageLimitWidth.Point)
            Dim XDisplacement As New XPoint(PDFHelper.PageLimitLeft.Point, PDFHelper.PageLimitTop.Point + XCurrentHeight.Point)

            Dim oXPen As XPen = New XPen(XColors.Black, 0.25)
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage, XGraphicsPdfPageOptions.Append)
                oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                oXGraphics.DrawLine(oXPen, XDisplacement, New XPoint(XDisplacement.X + XBlockWidth.Point, XDisplacement.Y))
            End Using
        End Sub
        Public Function GetNumbering(ByVal oFormItem As BaseFormItem) As Integer
            ' gets the numbering of the supplied form item
            Dim iStart As Integer = Start
            For Each oChild In Children
                If oChild.Value.Equals(GetType(FormatterGroup)) Then
                    Dim oGroup As FormatterGroup = FormMain.FindChild(oChild.Key)
                    For Each oGroupChild In oGroup.Children
                        If oGroupChild.Key.Equals(oFormItem.GUID) Then
                            Return iStart
                        End If
                        GetNumberingMain(oGroupChild, iStart)
                    Next
                Else
                    If oChild.Key.Equals(oFormItem.GUID) Then
                        Return iStart
                    End If
                    GetNumberingMain(oChild, iStart)
                End If
            Next
            Return iStart
        End Function
        Private Sub GetNumberingMain(ByVal oChild As KeyValuePair(Of Guid, Type), ByRef iStart As Integer)
            ' main numbering routine
            If oChild.Value.Equals(GetType(FormSubSection)) Then
                Dim oSubSection As FormSubSection = FormMain.FindChild(oChild.Key)
                Dim oNumberingList = oSubSection.ExtractFields(Of FieldNumbering)(GetType(FormField))
                If oNumberingList.Count > 0 AndAlso (Not oNumberingList.First.Spacer) Then
                    If ContinuousNumbering Then
                        iStart += Math.Max(If(IsNothing(oSubSection.SelectedTagItem), 0, oSubSection.SelectedTagItem.Count), 1)
                    Else
                        iStart += 1
                    End If
                End If
            ElseIf oChild.Value.Equals(GetType(FormBlock)) Then
                Dim oBlock As FormBlock = FormMain.FindChild(oChild.Key)
                Dim oNumberingList = oBlock.ExtractFields(Of FieldNumbering)(GetType(FormField))
                If oNumberingList.Count > 0 AndAlso (Not oNumberingList.First.Spacer) Then
                    iStart += 1
                End If
            ElseIf oChild.Value.Equals(GetType(FormMCQ)) Then
                Dim oMCQ As FormMCQ = FormMain.FindChild(oChild.Key)
                Dim oNumberingList = oMCQ.ExtractFields(Of FieldNumbering)(GetType(FormField))
                If oNumberingList.Count > 0 AndAlso (Not oNumberingList.First.Spacer) Then
                    If ContinuousNumbering Then
                        iStart += Math.Max(oMCQ.DataObject.MCQ.MCQStore.Count, 1)
                    Else
                        iStart += 1
                    End If
                End If
            End If
        End Sub
        Public Shared Sub SetDivider(ByVal bDividerLine As Boolean)
            ' sets a horizontal divider
            Dim fFormItemHeight As Double = PDFHelper.SpacerBitmapSourceLine.Item3.Point
            FormPDF.AddPage(fFormItemHeight)
            If bDividerLine Then
                IncrementCurrentLine(ParamList.Value(ParamList.KeyPDFPages)(ParamList.Value(ParamList.KeyPDFPages).Count - 1).Item1, ParamList.Value(ParamList.KeyXCurrentHeight))
            End If
            ParamList.Value(ParamList.KeyXCurrentHeight) = New XUnit(ParamList.Value(ParamList.KeyXCurrentHeight).Point + fFormItemHeight)
        End Sub
        Public Sub GenerateGUIDKey()
            ' generates a new GUID key for the tag
            Tag = Guid.NewGuid.ToString
        End Sub
        Public Sub CopyKeyToClipboard()
            ' copies the GUID key to the clipboard
            If Tag <> String.Empty Then
                My.Computer.Clipboard.SetText(Tag)
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, Tag + " copied to clipboard."))
            End If
        End Sub
        Private Sub CheckInvalidate()
            ' checks to see if the render resolution has changed, and invalidate all child items if this is true
            If SectionResolution <> oSettings.RenderResolutionValue Then
                SectionResolution = oSettings.RenderResolutionValue

                PDFHelper.SetSpacerBitmaps()

                ImageTracker.Invalidate(Me, True)
            End If
        End Sub
        Public Class ImageViewerClass
            Implements INotifyPropertyChanged, IDisposable

            Private m_Images As ObservableCollection(Of ImageSource)
            Private m_Descriptions As List(Of String)
            Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
            Protected Sub OnPropertyChanged(ByVal sName As String)
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
            End Sub
            Sub New()
                m_Images = New ObservableCollection(Of ImageSource)
                m_Descriptions = New List(Of String)
            End Sub
            Public ReadOnly Property Images As ObservableCollection(Of ImageSource)
                Get
                    Return m_Images
                End Get
            End Property
            Public Sub Clear()
                m_Images.Clear()
                m_Descriptions.Clear()
                CommonFunctions.ClearMemory()
                OnPropertyChanged("Images")
            End Sub
            Public Sub Add(ByVal oImage As ImageSource, ByVal sDescription As String)
                ' adds new image
                m_Images.Add(oImage)
                m_Descriptions.Add(sDescription)
                OnPropertyChanged("Images")
            End Sub
#Region "IDisposable Support"
            Private disposedValue As Boolean
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not disposedValue Then
                    If disposing Then
                        m_Images.Clear()
                        m_Descriptions.Clear()
                    End If
                End If
                disposedValue = True
            End Sub
            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
            End Sub
#End Region
        End Class
        Public Class ImageTrackerClass
            Implements IDisposable

            ' the tuple consists of the compressed byte array, horizontal and vertical resolution, and image stride
            Private m_Tracker As Dictionary(Of GUIDArray, Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, PixelFormat))

            Sub New()
                m_Tracker = New Dictionary(Of GUIDArray, Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, PixelFormat))
            End Sub
            Public Shadows Sub Clear()
                m_Tracker.Clear()
                For Each oKey In m_Tracker.Keys
                    m_Tracker(oKey) = Nothing
                Next
                CommonFunctions.ClearMemory()
            End Sub
            Public Shadows Sub Add(ByVal oImage As Imaging.BitmapSource, ByVal oItemList As List(Of BaseFormItem))
                ' adds new image with a list of source form item guids
                Dim oOrderedGUIDArray As GUIDArray = OrderedGUIDArray(oItemList)
                If Not oOrderedGUIDArray.Contains(Guid.Empty) Then
                    m_Tracker.Add(OrderedGUIDArray(oItemList), Converter.BitmapSourceToByteArray(oImage, True))
                End If
            End Sub
            Public Function GetImage(ByVal oItemList As List(Of BaseFormItem)) As Imaging.BitmapSource
                ' gets image based on source form item guids
                Dim oOrderedGUIDArray As GUIDArray = OrderedGUIDArray(oItemList)
                If m_Tracker.ContainsKey(oOrderedGUIDArray) Then
                    Return Converter.ByteArrayToBitmapSource(m_Tracker(oOrderedGUIDArray))
                Else
                    Return Nothing
                End If
            End Function
            Public Sub Invalidate(ByVal oFormItem As BaseFormItem, ByVal bInvalidateChildren As Boolean)
                ' invalidates images based on the supplied form item guid
                InvalidateItem(oFormItem)

                ' invalidate children
                If bInvalidateChildren Then
                    Dim oFormItemList As List(Of BaseFormItem) = oFormItem.GetFormItems(Nothing)
                    For Each oChildFormItem In oFormItemList
                        InvalidateItem(oChildFormItem)
                    Next
                End If
            End Sub
            Private Sub InvalidateItem(ByVal oFormItem As BaseFormItem)
                ' invalidates images based on the supplied form item guid
                Dim oInvalidatedImageList As List(Of GUIDArray) = (From oGUIDArray In m_Tracker.Keys Where oGUIDArray.Contains(oFormItem.GUID) Select oGUIDArray).ToList
                For Each oGUIDArray In oInvalidatedImageList
                    m_Tracker.Remove(oGUIDArray)
                Next
            End Sub
            Public Function ItemPresent(ByVal oItemList As List(Of BaseFormItem)) As Boolean
                ' checks if the source item is present
                Dim oOrderedGUIDArray As GUIDArray = OrderedGUIDArray(oItemList)
                Return m_Tracker.ContainsKey(oOrderedGUIDArray)
            End Function
            Private Function OrderedGUIDArray(ByVal oItemList As List(Of BaseFormItem)) As GUIDArray
                ' returns an ordered list of guid extracted from the form items
                Return New GUIDArray((From oFormItem In oItemList Select oFormItem.GUID).OrderBy(Function(x) x).ToList)
            End Function
            Public Shared Function GetList(ByVal oFormItemArray As BaseFormItem()) As List(Of BaseFormItem)
                Return oFormItemArray.Cast(Of BaseFormItem).ToList
            End Function
            Public Class GUIDArray
                Public GUIDArray As List(Of Guid)
                Sub New(ByVal oGUIDArray As List(Of Guid))
                    GUIDArray = oGUIDArray
                End Sub
                Function Contains(value As Guid) As Boolean
                    Return GUIDArray.Contains(value)
                End Function
                Public Overloads Overrides Function Equals(obj As Object) As Boolean
                    If obj Is Nothing OrElse Not Me.GetType() Is obj.GetType() Then
                        Return False
                    End If

                    Dim oGUIDArray As GUIDArray = CType(obj, GUIDArray)
                    Return GUIDArray.SequenceEqual(oGUIDArray.GUIDArray)
                End Function
                Public Overrides Function GetHashCode() As Integer
                    Dim iHashCode As Integer = 0
                    If GUIDArray.Count <> 0 Then
                        iHashCode = GUIDArray(0).GetHashCode
                        For i = 1 To GUIDArray.Count - 1
                            iHashCode = iHashCode Xor GUIDArray(i).GetHashCode
                        Next
                    End If
                    Return iHashCode
                End Function
            End Class
#Region "IDisposable Support"
            Private disposedValue As Boolean
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not disposedValue Then
                    If disposing Then
                        Clear()
                    End If
                End If
                disposedValue = True
            End Sub
            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
            End Sub
#End Region
        End Class
    End Class
    <DataContract(IsReference:=True)> Public Class FormSubSection
        Inherits BaseFormItem

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public BlockWidth As Integer
            <DataMember> Public SelectedTag As Integer

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .BlockWidth = BlockWidth
                    .SelectedTag = SelectedTag
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            CType(oFormItem, FormSubSection).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Property BlockWidth As Integer
            Get
                Return m_DataObject.BlockWidth
            End Get
            Set(value As Integer)
                Dim iOldBlockWidth As Integer = m_DataObject.BlockWidth

                If value < 1 Then
                    m_DataObject.BlockWidth = 1
                ElseIf value > PDFHelper.PageBlockWidth Then
                    m_DataObject.BlockWidth = PDFHelper.PageBlockWidth
                Else
                    m_DataObject.BlockWidth = value
                End If

                If iOldBlockWidth <> m_DataObject.BlockWidth Then
                    Dim oBlockList As List(Of FormBlock) = ExtractFields(Of FormBlock)(GetType(FormBlock))
                    Using oSuspender As New Suspender(Me, True)
                        ' reset grid width for all child blocks
                        For Each oBlock In oBlockList
                            oBlock.GridWidth = oBlock.GridWidth
                        Next

                        OnPropertyChangedLocal("BlockWidth")
                        OnPropertyChangedLocal("BlockWidthText")
                    End Using
                End If
            End Set
        End Property
        Public Property BlockWidthText As String
            Get
                Return m_DataObject.BlockWidth.ToString
            End Get
            Set(value As String)
                BlockWidth = CInt(Val(value))
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "Tag:   " + SelectedTagText
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .BlockWidth = PDFHelper.PageBlockWidth
                .SelectedTag = -1
            End With
            TitleChanged()
        End Sub
#End Region
#Region "Tags"
        Public Property SelectedTag As Integer
            Get
                Return m_DataObject.SelectedTag
            End Get
            Set(value As Integer)
                If m_DataObject.SelectedTag <> value Then
                    Using oSuspender As New Suspender(Me, True)
                        m_DataObject.SelectedTag = value
                        OnPropertyChangedLocal("SelectedTag")
                        OnPropertyChangedLocal("SelectedTagText")
                        OnPropertyChangedLocal("SelectedTagItem")
                    End Using
                End If
            End Set
        End Property
        Public Property SelectedTagText As String
            Get
                If m_DataObject.SelectedTag = -1 Then
                    Return String.Empty
                Else
                    If (Not IsNothing(FormMain.Tags)) AndAlso m_DataObject.SelectedTag < FormMain.Tags.Count AndAlso m_DataObject.SelectedTag >= 0 Then
                        Return FormMain.Tags(m_DataObject.SelectedTag).Name
                    Else
                        Return String.Empty
                    End If
                End If
            End Get
            Set(value As String)
                Dim oNames As List(Of String) = (From oTag In FormMain.Tags Select oTag.Name).ToList
                If oNames.Contains(Trim(value)) Then
                    SelectedTag = oNames.IndexOf(Trim(value))
                End If
            End Set
        End Property
        Public ReadOnly Property SelectedTagItem As TagClass
            Get
                If FormMain.Tags.Count = 0 Then
                    Return Nothing
                Else
                    If m_DataObject.SelectedTag >= 0 Then
                        Return FormMain.Tags(m_DataObject.SelectedTag)
                    Else
                        Return Nothing
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property TagCount As Integer
            Get
                Return FormMain.Tags.Count
            End Get
        End Property
        Public Sub AddRow()
            If m_DataObject.SelectedTag >= 0 Then
                Using oSuspender As New Suspender(Me, True)
                    Dim oDataGridSubjectsTags As Controls.DataGrid = Root.DataGridSubjectsTags
                    Dim oEmptyList As New List(Of String)
                    For i = 0 To FormMain.Tags(m_DataObject.SelectedTag).Width - 1
                        oEmptyList.Add(String.Empty)
                    Next
                    FormMain.Tags(m_DataObject.SelectedTag).Add(oEmptyList, oDataGridSubjectsTags.SelectedIndex)
                End Using
            End If
        End Sub
        Public Sub RemoveRow()
            If m_DataObject.SelectedTag >= 0 Then
                Using oSuspender As New Suspender(Me, True)
                    Dim oDataGridSubjectsTags As Controls.DataGrid = Root.DataGridSubjectsTags
                    If FormMain.Tags(m_DataObject.SelectedTag).Count > 0 Then
                        FormMain.Tags(m_DataObject.SelectedTag).Remove(oDataGridSubjectsTags.SelectedIndex)
                    End If
                End Using
            End If
        End Sub
        Public Sub AddColumn()
            If m_DataObject.SelectedTag >= 0 Then
                Using oSuspender As New Suspender(Me, True)
                    FormMain.Tags(m_DataObject.SelectedTag).Width += 1
                End Using
            End If
        End Sub
        Public Sub RemoveColumn()
            If m_DataObject.SelectedTag >= 0 Then
                Using oSuspender As New Suspender(Me, True)
                    FormMain.Tags(m_DataObject.SelectedTag).Width -= 1
                End Using
            End If
        End Sub
        Public Sub ImportTags()
            ' add a new tag sheet from the first workbook sheet
            ' scans for the largest contiguous rectangular area starting from the top left
            Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
            oOpenFileDialog.FileName = String.Empty
            oOpenFileDialog.DefaultExt = "*.xlsx"
            oOpenFileDialog.Multiselect = False
            oOpenFileDialog.Filter = "Excel Spreadsheet|*.xlsx"
            oOpenFileDialog.Title = "Load Tags From File"
            Dim result? As Boolean = oOpenFileDialog.ShowDialog()
            If result = True Then
                Using oSuspender As New Suspender(Me, True)
                    Dim oFileInfo As New IO.FileInfo(oOpenFileDialog.FileName)
                    Dim oExcelDocument As ClosedXML.Excel.XLWorkbook = Nothing
                    Try
                        oExcelDocument = New ClosedXML.Excel.XLWorkbook(oOpenFileDialog.FileName)
                    Catch ex1 As System.IO.FileFormatException
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oFileInfo.Name + ". File format error."))
                        Exit Sub
                    Catch ex2 As System.IO.IOException
                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading " + oFileInfo.Name + ". File access error."))
                        Exit Sub
                    End Try

                    If IsNothing(FormMain.Tags) Then
                        FormMain.Tags = New List(Of TagClass)
                    End If
                    FormMain.Tags.Clear()

                    ' load all sheets except for the subject sheet
                    For Each oWorksheet As ClosedXML.Excel.IXLWorksheet In oExcelDocument.Worksheets
                        If Trim(oWorksheet.Name) <> WorksheetSubjects Then
                            ' check contiguous area
                            Dim xLimit As Integer = Integer.MaxValue
                            Dim sReturn As String = "Tags"
                            Dim iCurrentColumn As Integer = 1
                            Do Until sReturn = String.Empty
                                sReturn = Trim(oWorksheet.Cell(1, iCurrentColumn).GetString)
                                iCurrentColumn += 1
                            Loop
                            xLimit = iCurrentColumn - 2

                            Dim yLimit As Integer = Integer.MaxValue
                            For iColumn = 1 To xLimit
                                sReturn = "Tags"
                                Dim iCurrentRow As Integer = 1
                                Do Until sReturn = String.Empty
                                    sReturn = Trim(oWorksheet.Cell(iCurrentRow, iColumn).GetString)
                                    iCurrentRow += 1
                                Loop

                                yLimit = Math.Min(yLimit, iCurrentRow - 2)
                            Next

                            Dim oTags As New TagClass(Trim(oWorksheet.Name), "Item", xLimit)
                            For i = 1 To yLimit
                                Dim oRow As New List(Of String)
                                For j = 1 To xLimit
                                    oRow.Add(Trim(oWorksheet.Cell(i, j).GetString))
                                Next
                                oTags.Add(oRow)
                            Next
                            FormMain.Tags.Add(oTags)
                        End If
                    Next

                    ' run through all subsections in the form and reset the selecteditem to -1
                    Dim oSubSectionList As List(Of FormSubSection) = FormMain.GetFormItems(Of FormSubSection)
                    For Each oSubSection In oSubSectionList
                        oSubSection.SelectedTag = -1
                    Next

                    OnPropertyChangedLocal("SelectedTag")
                    OnPropertyChangedLocal("SelectedTagText")

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Tags loaded from file " + oFileInfo.Name + "."))
                End Using
            End If
        End Sub
        Public Sub ExportTags()
            ' exports in opendocument spreadsheet format
            Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
            oSaveFileDialog.FileName = String.Empty
            oSaveFileDialog.DefaultExt = "*.xlsx"
            oSaveFileDialog.Filter = "Excel Spreadsheet|*.xlsx"
            oSaveFileDialog.Title = "Save Tags To File"
            oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
            Dim result? As Boolean = oSaveFileDialog.ShowDialog()
            If result = True Then
                ' create new Excel document
                Dim oExcelDocument As New ClosedXML.Excel.XLWorkbook

                For i = oExcelDocument.Worksheets.Count - 1 To 0 Step -1
                    oExcelDocument.Worksheets.Delete(i + 1)
                Next

                ' set headers
                For iTagIndex As Integer = 0 To FormMain.Tags.Count - 1
                    Dim oWorksheet As ClosedXML.Excel.IXLWorksheet = oExcelDocument.AddWorksheet(Converter.AlphaNumericOnly(FormMain.Tags(iTagIndex).Name, False))

                    For i = 0 To FormMain.Tags(iTagIndex).Count - 1
                        For iColumn As Integer = 0 To FormMain.Tags(iTagIndex).Width - 1
                            oWorksheet.Cell(1 + i, 1 + iColumn).Value = FormMain.Tags(iTagIndex).Tag(i, iColumn)
                        Next
                    Next

                    For iColumn As Integer = 0 To FormMain.Tags(iTagIndex).Width - 1
                        oWorksheet.Column(1 + iColumn).AdjustToContents()
                    Next
                Next

                oExcelDocument.SaveAs(oSaveFileDialog.FileName)

                Dim oFileInfo As New IO.FileInfo(oSaveFileDialog.FileName)
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Tags saved to file " + oFileInfo.Name + "."))
            End If
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Sub-Section"
            End Get
        End Property
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.0
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMSubSection"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelBlock")
                oDisplayFilter.Add("StackPanelBlockWidth")
                oDisplayFilter.Add("StackPanelTags")
                oDisplayFilter.Add("BlockWidth")
                oDisplayFilter.Add("BlockTags")
                oDisplayFilter.Add("StackPanelTagsDataRow")
                oDisplayFilter.Add("StackPanelTagsDataColumn")
                oDisplayFilter.Add("DataGridSubjectsTags")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelBlock.DataContext = Me
            Root.StackPanelBlockWidth.DataContext = Me
            Root.StackPanelTags.DataContext = Me

            Dim oBindingTags1 As New Data.Binding
            oBindingTags1.Path = New PropertyPath("BlockWidthText")
            oBindingTags1.Mode = Data.BindingMode.TwoWay
            oBindingTags1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockWidth.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingTags1)

            Dim oBindingTags2 As New Data.Binding
            oBindingTags2.Path = New PropertyPath("SelectedTagText")
            oBindingTags2.Mode = Data.BindingMode.TwoWay
            oBindingTags2.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.BlockTags.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingTags2)
        End Sub
        Public Overrides Sub Display()
            ' validate the selected tag number and rebuild the datagrid columns
            Dim oDataGridSubjectsTags As Controls.DataGrid = Root.DataGridSubjectsTags

            oDataGridSubjectsTags.Columns.Clear()
            If (Not IsNothing(FormMain.Tags)) AndAlso FormMain.Tags.Count > 0 AndAlso m_DataObject.SelectedTag >= 0 Then
                For i = 0 To FormMain.Tags(m_DataObject.SelectedTag).Width
                    Dim oColumn As New Controls.DataGridTextColumn
                    With oColumn
                        If i = 0 Then
                            .Width = New Controls.DataGridLength(oDataGridSubjectsTags.ActualWidth * 0.05, Controls.DataGridLengthUnitType.Pixel)
                            .Header = "No."
                            .IsReadOnly = True

                            Dim oBinding As New Data.Binding
                            oBinding.Path = New PropertyPath("Number")
                            oBinding.Mode = Data.BindingMode.TwoWay
                            oBinding.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                            .Binding = oBinding
                        Else
                            .Width = New Controls.DataGridLength(1, Controls.DataGridLengthUnitType.Star)

                            If FormMain.Tags(m_DataObject.SelectedTag).Width = 1 Then
                                .Header = FormMain.Tags(m_DataObject.SelectedTag).Header
                            Else
                                .Header = FormMain.Tags(m_DataObject.SelectedTag).Header + " " + i.ToString
                            End If

                            Dim oBinding As New Data.Binding
                            oBinding.Path = New PropertyPath("Tags[" + (i - 1).ToString + "]")
                            oBinding.Mode = Data.BindingMode.TwoWay
                            oBinding.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                            .Binding = oBinding
                        End If
                    End With
                    oDataGridSubjectsTags.Columns.Add(oColumn)
                Next
                oDataGridSubjectsTags.ItemsSource = FormMain.Tags(m_DataObject.SelectedTag).TagStore
            Else
                oDataGridSubjectsTags.ItemsSource = Nothing
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "ad377093-a252-4ae9-b481-64f6634af835"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormBlock), GetType(FieldNumbering), GetType(FormatterDivider)}
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedSingleOnly As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterDivider)}
            End Get
        End Property
        Public Overrides Sub RenderPDF()
            If Not IsNothing(SelectedTagItem) Then
                Dim iSubSectionsPerLine As Integer = GetSubSectionsPerLine()
                Dim oLineBlockList As New List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize)))

                ' loop until all the tags have been processed
                Dim iCurrentTag As Integer = 0
                Do Until iCurrentTag >= SelectedTagItem.Count
                    For iColumn = 0 To iSubSectionsPerLine - 1
                        For Each oChild In Children
                            If oChild.Value.Equals(GetType(FormBlock)) Then
                                Dim oBlock As FormBlock = FormMain.FindChild(oChild.Key)
                                Dim iBlockCount As Integer = Aggregate oLineBlock In oLineBlockList Into Sum(oLineBlock.Item1.EffectiveBlockWidth)
                                If iBlockCount + oBlock.EffectiveBlockWidth > PDFHelper.PageBlockWidth Then
                                    ' output current line and reset block list
                                    FormSection.RenderCurrentLine(oLineBlockList)
                                End If

                                ' add to block list
                                Dim oTagList As String() = SelectedTagItem.TagStore(iCurrentTag).Tags.ToArray
                                Dim oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = oBlock.GetBlockDimensions(True, oTagList)
                                oLineBlockList.Add(New Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))(oBlock, oBlockDimensions, oTagList, oBlock.GetExpandedBlockDimensions(oBlockDimensions)))
                            End If
                        Next

                        ' exit loop if tags finished
                        iCurrentTag += 1
                        If iCurrentTag >= SelectedTagItem.Count Then
                            Exit For
                        End If
                    Next

                    ' output current line and reset block list
                    FormSection.RenderCurrentLine(oLineBlockList)
                Loop

                ' output current line and reset block list
                FormSection.RenderCurrentLine(oLineBlockList)
            End If
        End Sub
        Public Sub DisplayPDF()
            ' renders the MCQ to a bitmap
            Using oSuspender As New Suspender()
                If Not IsNothing(SelectedTagItem) Then
                    Dim oSubSectionImage As Imaging.BitmapSource = Nothing
                    Dim oSection As FormSection = ParamList.Value(ParamList.KeyCurrentSection)
                    If Not IsNothing(oSection) Then
                        oSubSectionImage = oSection.ImageTracker.GetImage(FormSection.ImageTrackerClass.GetList({Me}))
                        If IsNothing(oSubSectionImage) Then
                            Dim iSubSectionsPerLine As Integer = GetSubSectionsPerLine()
                            Dim oLineBlockList As New List(Of Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize)))
                            Dim oBitmapStore As New List(Of Imaging.BitmapSource)

                            ' loop until all the tags have been processed
                            Dim iCurrentTag As Integer = 0
                            Do Until iCurrentTag >= SelectedTagItem.Count
                                For iColumn = 0 To iSubSectionsPerLine - 1
                                    For Each oChild In Children
                                        If oChild.Value.Equals(GetType(FormBlock)) Then
                                            Dim oBlock As FormBlock = FormMain.FindChild(oChild.Key)
                                            Dim iBlockCount As Integer = Aggregate oLineBlock In oLineBlockList Into Sum(oLineBlock.Item1.EffectiveBlockWidth)
                                            If iBlockCount + oBlock.EffectiveBlockWidth > PDFHelper.PageBlockWidth Then
                                                ' output current line and reset block list
                                                oSection.DisplayCurrentLine(oLineBlockList, oBitmapStore)
                                            End If

                                            ' add to block list
                                            Dim oTagList As String() = SelectedTagItem.TagStore(iCurrentTag).Tags.ToArray
                                            Dim oBlockDimensions As Tuple(Of XUnit, XUnit, List(Of Integer)) = oBlock.GetBlockDimensions(True, oTagList)
                                            oLineBlockList.Add(New Tuple(Of FormBlock, Tuple(Of XUnit, XUnit, List(Of Integer)), String(), Tuple(Of XUnit, XUnit, XSize))(oBlock, oBlockDimensions, oTagList, oBlock.GetExpandedBlockDimensions(oBlockDimensions)))
                                        End If
                                    Next

                                    ' exit loop if tags finished
                                    iCurrentTag += 1
                                    If iCurrentTag >= SelectedTagItem.Count Then
                                        Exit For
                                    End If
                                Next

                                ' output current line and reset block list
                                oSection.DisplayCurrentLine(oLineBlockList, oBitmapStore)
                            Loop

                            ' output current line and reset block list
                            oSection.DisplayCurrentLine(oLineBlockList, oBitmapStore)

                            If oBitmapStore.Count > 0 Then
                                Dim fCurrentHeight As Double = 0
                                Dim fMaxWidth As Double = (Aggregate oBitmap In oBitmapStore Into Max(oBitmap.Width))
                                Dim fMaxHeight As Double = (Aggregate oBitmap In oBitmapStore Into Sum(oBitmap.Height)) + (Math.Max(oBitmapStore.Count - 1, 0) * PDFHelper.SpacerBitmapSourceLine.Item1.Height)
                                Dim fBitmapWidth As Double = (Aggregate oBitmap In oBitmapStore Into Max(oBitmap.PixelWidth))
                                Dim fBitmapHeight As Double = (Aggregate oBitmap In oBitmapStore Into Sum(oBitmap.PixelHeight)) + (Math.Max(oBitmapStore.Count - 1, 0) * PDFHelper.SpacerBitmapSourceLine.Item1.PixelHeight)
                                Dim oDrawingVisual As New DrawingVisual()
                                Dim oDrawingContext As DrawingContext = oDrawingVisual.RenderOpen()

                                ' draw bitmaps
                                oDrawingContext.DrawRectangle(Brushes.White, Nothing, New Rect(0, 0, fMaxWidth, fMaxHeight))
                                For Each oBitmap In oBitmapStore
                                    Dim oRect As New Rect(0, fCurrentHeight, oBitmap.Width, oBitmap.Height)
                                    oDrawingContext.DrawImage(oBitmap, oRect)
                                    fCurrentHeight += oBitmap.Height + PDFHelper.SpacerBitmapSourceLine.Item1.Height
                                Next

                                oDrawingContext.Close()

                                Dim oRenderTargetBitmap As New Imaging.RenderTargetBitmap(fBitmapWidth, fBitmapHeight, oSection.SectionResolution, oSection.SectionResolution, PixelFormats.Pbgra32)
                                oRenderTargetBitmap.Render(oDrawingVisual)
                                oSubSectionImage = oRenderTargetBitmap
                            End If

                            oSection.ImageTracker.Add(oSubSectionImage, FormSection.ImageTrackerClass.GetList({Me}))
                        End If

                        If Not IsNothing(oSubSectionImage) Then
                            FormSection.ImageViewer.Add(oSubSectionImage, "Sub-Section Image")
                        End If
                    End If

                    ' remove subnumbering key
                    If ParamList.ContainsKey(ParamList.KeySubNumberingCurrent) Then
                        ParamList.Value(ParamList.KeySubNumberingCurrent) += 1
                    End If
                End If
            End Using
        End Sub
        Public Overrides ReadOnly Property FormattingFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterDivider)}
            End Get
        End Property
        Private Function GetSubSectionsPerLine() As Integer
            ' returns the number of subsections that can fit in a line
            Dim iBlockCount As Integer = ExtractFields(Of FormBlock)(GetType(FormBlock)).Count
            Return Math.Max(Math.Floor(PDFHelper.PageBlockWidth / (iBlockCount * BlockWidth)), 1)
        End Function
    End Class
    <DataContract(IsReference:=True)> Public MustInherit Class FormFormatter
        Inherits BaseFormItem

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
        End Sub
#Region "Items"
        Public Overrides Sub TitleChanged()
        End Sub
        Public Overrides Sub Initialise()
        End Sub
#End Region
        Public Overrides ReadOnly Property Multiplier As Single
            Get
                Return 1.0
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                If Not IsNothing(Parent) Then
                    oDisplayFilter.AddRange(Parent.DisplayFilter)
                End If
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            If Not IsNothing(Parent) Then
                Parent.SetBindings()
            End If
        End Sub
        Public Overrides Sub Display()
            If Not IsNothing(Parent) Then
                Parent.Display()
            End If
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return ""
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overrides Sub RenderPDF()
        End Sub
    End Class
    <DataContract(IsReference:=True)> Public Class FormatterPageBreak
        Inherits FormFormatter

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Page Break"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMPageBreak"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "0192e74e-00de-4fcc-9344-5a49c99d9e41"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FormatterDivider
        Inherits FormFormatter

#Region "Serialisation"
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            MyBase.CloneProcessing(oFormItem)
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Divider"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMDivider"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "19907103-4ddc-42d9-845d-255a27b5fdb9"
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public Class FormatterGroup
        Inherits FormFormatter

        <DataMember> Private m_DataObject As DataObjectClass

#Region "Serialisation"
        <DataContract> Public Shadows Class DataObjectClass
            Implements ICloneable

            <DataMember> Public GroupTitle As String

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .GroupTitle = GroupTitle
                End With
                Return oDataObject
            End Function
        End Class
        Public Shadows Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Overrides Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
            CType(oFormItem, FormatterGroup).DataObject = Me.DataObject.Clone
        End Sub
#End Region
#Region "Items"
        Public Property GroupTitle As String
            Get
                Return m_DataObject.GroupTitle
            End Get
            Set(value As String)
                m_DataObject.GroupTitle = value
                OnPropertyChangedLocal("GroupTitle")
                TitleChanged()
            End Set
        End Property
        Public Overrides Sub TitleChanged()
            Title = "Title: " + GroupTitle
        End Sub
        Public Overrides Sub Initialise()
            m_DataObject = New DataObjectClass
            With m_DataObject
                .GroupTitle = String.Empty
            End With
            TitleChanged()
        End Sub
#End Region
        Public Shared Shadows ReadOnly Property Name As String
            Get
                Return "Group"
            End Get
        End Property
        Public Shared Shadows ReadOnly Property IconName As String
            Get
                Return "CCMGroup"
            End Get
        End Property
        Public Overrides ReadOnly Property DisplayFilter As List(Of String)
            Get
                Dim oDisplayFilter As New List(Of String)
                oDisplayFilter.Add("DockPanelBlock")
                oDisplayFilter.Add("StackPanelGroupTitle")
                oDisplayFilter.Add("GroupTitle")
                Return oDisplayFilter.Distinct.ToList
            End Get
        End Property
        Public Overrides Sub SetBindings()
            Root.DockPanelBlock.DataContext = Me
            Root.GroupTitle.DataContext = Me

            Dim oBindingGroup1 As New Data.Binding
            oBindingGroup1.Path = New PropertyPath("GroupTitle")
            oBindingGroup1.Mode = Data.BindingMode.TwoWay
            oBindingGroup1.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
            Root.GroupTitle.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBindingGroup1)
        End Sub
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return "2f45fd61-2826-4416-9ec4-83b85303d237"
            End Get
        End Property
        Public Overrides ReadOnly Property AllowedFields As List(Of Type)
            Get
                If IsNothing(Parent) Then
                    Return New List(Of Type)
                Else
                    ' get parent allowed fields, then remove all formatting fields, and add the local formatting fields
                    Dim oAllowedFields As New List(Of Type)
                    oAllowedFields.AddRange(Parent.AllowedFields)
                    For i = oAllowedFields.Count - 1 To 0 Step -1
                        If oAllowedFields(i).IsSubclassOf(GetType(FormFormatter)) Then
                            oAllowedFields.RemoveAt(i)
                        End If
                    Next
                    oAllowedFields.AddRange(FormattingFields)
                    Return oAllowedFields
                End If
            End Get
        End Property
        Public Overrides ReadOnly Property FormattingFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormatterDivider)}
            End Get
        End Property
    End Class
    <DataContract(IsReference:=True)> Public MustInherit Class BaseFormItem
        Implements INotifyPropertyChanged, IDisposable, ICloneable

        ' guid is unique for items in a tree, but is empty for those not in a tree
        ' isnumbered is true for items that update the numbering tree
        ' FormMain refers to the top most form item
        ' parent refers to the immediate parent
        ' children refers to the immediate child form items
        Public Shared Root As CCM
        Public Shared FormMain As FormMain
        Public Shared FormProperties As FormProperties
        Public Shared FormHeader As FormHeader
        Public Shared FormPageHeader As FormPageHeader
        Public Shared FormFormHeader As FormFormHeader
        Public Shared FormBody As FormBody
        Public Shared FormFooter As FormFooter
        Public Shared FormPDF As FormPDF
        Public Shared FormExport As FormExport
        Public Shared SectionList As List(Of List(Of BaseFormItem))
        Public Level As Integer
        Public Expanded As Boolean
        Public BorderFormItem As BorderFormItem
        Public ImageExpander As ImageExpander
        Public BorderSpacer As BorderSpacer
        <DataMember> Private m_DataObject As DataObjectClass
        Private m_Title As String

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub OnPropertyChangedLocal(ByVal sName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
        End Sub
        Sub New()
            OnDeserializing()
        End Sub
#Region "Serialisation"
        <OnDeserializing()> Private Sub OnDeserializing(Optional ByVal c As StreamingContext = Nothing)
            ' for version tolerant deserialisation
            Level = -1
            Expanded = False
            ImageExpander = Nothing
            BorderSpacer = Nothing
            SectionList = New List(Of List(Of BaseFormItem))

            If IsNothing(m_DataObject) Then
                m_DataObject = New DataObjectClass
            End If
            m_DataObject.m_GUID = Guid.Empty
            m_DataObject.m_IsNumbered = True
            m_DataObject.m_Parent = Nothing
            m_DataObject.m_Children = New Dictionary(Of Guid, Type)

            OnDeserializingOverride()

            Using oSuspender As New Suspender(Me, True)
                Initialise()
            End Using
        End Sub
        Public Overridable Sub OnDeserializingOverride()
            ' to be overriden in derived classes
            ' for version tolerant deserialisation
        End Sub
        <OnDeserialized()> Private Sub OnDeserialized(Optional ByVal c As StreamingContext = Nothing)
            ' for version tolerant deserialisation
            TitleChanged()
            OnDeserializedOverride()
        End Sub
        Public Overridable Sub OnDeserializedOverride()
            ' to be overriden in derived classes
            ' for version tolerant deserialisation
        End Sub
        <DataContract> Public Class DataObjectClass
            Implements ICloneable

            <DataMember> Public m_GUID As Guid
            <DataMember> Public m_IsNumbered As Boolean
            <DataMember> Public m_Parent As BaseFormItem
            Public m_Children As Dictionary(Of Guid, Type)

            <DataMember> Private Property ChildrenSave As Dictionary(Of Guid, String)
                Get
                    Return (From oChild In m_Children Select New KeyValuePair(Of Guid, String)(oChild.Key, oChild.Value.AssemblyQualifiedName)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                End Get
                Set(value As Dictionary(Of Guid, String))
                    Dim oKeyValuePairList As List(Of KeyValuePair(Of Guid, Type)) = (From oChild In value Select New KeyValuePair(Of Guid, Type)(oChild.Key, Type.GetType(oChild.Value))).ToList
                    If IsNothing(m_Children) Then
                        m_Children = New Dictionary(Of Guid, Type)
                    Else
                        m_Children.Clear()
                    End If
                    For Each oChild In oKeyValuePairList
                        m_Children.Add(oChild.Key, oChild.Value)
                    Next
                End Set
            End Property
            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oDataObject As New DataObjectClass
                With oDataObject
                    .m_GUID = m_GUID
                    .m_IsNumbered = m_IsNumbered
                    .m_Parent = m_Parent
                    .m_Children = New Dictionary(Of Guid, Type)
                End With
                Return oDataObject
            End Function
        End Class
        Public Property DataObject As DataObjectClass
            Get
                Return m_DataObject
            End Get
            Set(value As DataObjectClass)
                m_DataObject = value
            End Set
        End Property
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oType As Type = Me.GetType

            ' creates new form item with new guid but same data object
            Dim oFormItem As BaseFormItem = Activator.CreateInstance(oType)
            oFormItem.DataObject = Me.DataObject.Clone
            oFormItem.GUID = Guid.NewGuid

            ' specific processing for inherited items
            CloneProcessing(oFormItem)

            ' add children
            For Each oChild In Children
                Dim oChildItem As BaseFormItem = FormMain.FindChild(oChild.Key)
                Dim oNewChildItem As BaseFormItem = oChildItem.Clone
                oNewChildItem.Parent = oFormItem
            Next

            Return oFormItem
        End Function
        Public MustOverride Sub CloneProcessing(ByRef oFormItem As BaseFormItem)
#End Region
#Region "Items"
        Public MustOverride Sub TitleChanged()
        Public MustOverride Sub Initialise()
#End Region
        Public Property GUID As Guid
            Get
                Return m_DataObject.m_GUID
            End Get
            Set(value As Guid)
                m_DataObject.m_GUID = value
            End Set
        End Property
        Public Property Parent(Optional ByVal iInsertAfter As Integer = -2, Optional ByVal bCopy As Boolean = True) As BaseFormItem
            Get
                Return m_DataObject.m_Parent
            End Get
            Set(value As BaseFormItem)
                ' remove from previous parent
                Dim oOldParent As BaseFormItem = m_DataObject.m_Parent
                Dim bSameParent As Boolean = ((Not (IsNothing(value) Or IsNothing(Parent))) AndAlso value.GUID.Equals(m_DataObject.m_Parent.GUID))

                If Not IsNothing(m_DataObject.m_Parent) Then
                    m_DataObject.m_Parent.Children.Remove(GUID)
                End If

                ' add to new parent
                m_DataObject.m_Parent = value
                If IsNothing(m_DataObject.m_Parent) Then
                    If FormMain.ContainsChild(m_DataObject.m_GUID) Then
                        IterateRemoveChildren(m_DataObject.m_GUID)
                        FormMain.RemoveChild(Me)
                    End If
                    Level = -1
                Else
                    FormMain.AddChild(Me)
                    Level = m_DataObject.m_Parent.Level + 1

                    ' convert to list to handle insertions and additions
                    Dim oDictionaryList As List(Of KeyValuePair(Of Guid, Type)) = m_DataObject.m_Parent.Children.ToList
                    If iInsertAfter = -2 Then
                        ' add to the end
                        oDictionaryList.Add(New KeyValuePair(Of Guid, Type)(GUID, Me.GetType))
                    Else
                        ' insert
                        oDictionaryList.Insert(Math.Max(iInsertAfter + If(bSameParent And (Not bCopy), 0, 1), 0), New KeyValuePair(Of Guid, Type)(GUID, Me.GetType))
                    End If

                    ' extract numbering, border, and background fields to place at the front of the list
                    Dim oNumberingList As New List(Of KeyValuePair(Of Guid, Type))
                    Dim oBorderList As New List(Of KeyValuePair(Of Guid, Type))
                    Dim oBackgroundList As New List(Of KeyValuePair(Of Guid, Type))
                    For i = oDictionaryList.Count - 1 To 0 Step -1
                        Select Case oDictionaryList(i).Value
                            Case GetType(FieldNumbering)
                                oNumberingList.Add(oDictionaryList(i))
                                oDictionaryList.RemoveAt(i)
                            Case GetType(FieldBorder)
                                oBorderList.Add(oDictionaryList(i))
                                oDictionaryList.RemoveAt(i)
                            Case GetType(FieldBackground)
                                oBackgroundList.Add(oDictionaryList(i))
                                oDictionaryList.RemoveAt(i)
                        End Select
                    Next

                    ' convert back to dictionary
                    m_DataObject.m_Parent.Children.Clear()
                    For Each oKeyValuePair In oNumberingList
                        m_DataObject.m_Parent.Children.Add(oKeyValuePair.Key, oKeyValuePair.Value)
                    Next
                    For Each oKeyValuePair In oBorderList
                        m_DataObject.m_Parent.Children.Add(oKeyValuePair.Key, oKeyValuePair.Value)
                    Next
                    For Each oKeyValuePair In oBackgroundList
                        m_DataObject.m_Parent.Children.Add(oKeyValuePair.Key, oKeyValuePair.Value)
                    Next
                    For Each oKeyValuePair In oDictionaryList
                        m_DataObject.m_Parent.Children.Add(oKeyValuePair.Key, oKeyValuePair.Value)
                    Next

                    ' set levels
                    IterateChildrenLevel(m_DataObject.m_GUID)
                End If

                ' invalidate new parent if copied, and also invalidate old parent if moved
                If Not IsNothing(m_DataObject.m_Parent) Then
                    Dim oNewParentSection As FormSection = Nothing
                    If m_DataObject.m_Parent.GetType.Equals(GetType(FormSection)) Then
                        oNewParentSection = m_DataObject.m_Parent
                    Else
                        oNewParentSection = m_DataObject.m_Parent.FindParent(GetType(FormSection))
                    End If
                    If Not IsNothing(oNewParentSection) Then
                        oNewParentSection.ImageTracker.Invalidate(m_DataObject.m_Parent, False)
                    End If
                End If

                If Not bCopy AndAlso Not IsNothing(oOldParent) Then
                    Dim oOldParentSection As FormSection = Nothing
                    If oOldParent.GetType.Equals(GetType(FormSection)) Then
                        oOldParentSection = oOldParent
                    Else
                        oOldParentSection = oOldParent.FindParent(GetType(FormSection))
                    End If
                    If Not IsNothing(oOldParentSection) Then
                        oOldParentSection.ImageTracker.Invalidate(oOldParent, False)
                    End If
                End If
            End Set
        End Property
        Public ReadOnly Property Children As Dictionary(Of Guid, Type)
            Get
                Return m_DataObject.m_Children
            End Get
        End Property
        Public ReadOnly Property Numbering As String
            Get
                If IsNothing(m_DataObject.m_Parent) Then
                    Return String.Empty
                Else
                    Dim oParentMCQ As FormMCQ = FindParent(GetType(FormMCQ))
                    Dim oParentSubSection As FormSubSection = FindParent(GetType(FormSubSection))
                    If Not IsNothing(oParentMCQ) Then
                        ' special handling for MCQ
                        If oParentMCQ.IgnoreFields.Contains(Me.GetType) Then
                            ' render process active
                            Return oParentMCQ.GetNumbering(True)
                        Else
                            ' display numbering only
                            Return GetNumbering(True)
                        End If
                    ElseIf Not IsNothing(oParentSubSection) Then
                        ' special handling for subsections
                        Return GetNumbering(True)
                    ElseIf m_DataObject.m_Parent.IgnoreFields.Contains(Me.GetType) Then
                        Return String.Empty
                    Else
                        Return GetNumbering(False)
                    End If
                End If
            End Get
        End Property
        Public Shared ReadOnly Property Icon As ImageSource
            Get
                If m_Icons.ContainsKey(IconName) Then
                    Return m_Icons(IconName)
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public Property Title As String
            Get
                Return m_Title
            End Get
            Set(value As String)
                m_Title = value
                OnPropertyChangedLocal("Title")
            End Set
        End Property
        Public Shared ReadOnly Property Name As String
            Get
                Return String.Empty
            End Get
        End Property
        Public MustOverride ReadOnly Property Multiplier As Single
        Public Shared ReadOnly Property IconName As String
            Get
                Return String.Empty
            End Get
        End Property
        Public MustOverride ReadOnly Property DisplayFilter As List(Of String)
        Public MustOverride Sub SetBindings()
        Public MustOverride Sub Display()
        Public Shared Shadows ReadOnly Property TagGuid As String
            Get
                Return ""
            End Get
        End Property
        Public MustOverride ReadOnly Property AllowedFields As List(Of Type)
        Public Overridable ReadOnly Property AllowedSingleOnly As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overridable ReadOnly Property IgnoreFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Function SingleOnly(Optional ByVal oParent As BaseFormItem = Nothing) As Boolean
            If Not IsNothing(oParent) Then
                Return oParent.AllowedSingleOnly.Contains(Me.GetType)
            ElseIf Not IsNothing(Parent) Then
                Return Parent.AllowedSingleOnly.Contains(Me.GetType)
            Else
                Return False
            End If
        End Function
        Public MustOverride Sub RenderPDF()
        Public Shared ReadOnly Property MandatoryFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormSection), GetType(FormSubSection), GetType(FormMCQ), GetType(FormBlock)}
            End Get
        End Property
        Public Shared ReadOnly Property StaticFields As List(Of Type)
            Get
                Return New List(Of Type) From {GetType(FormProperties), GetType(FormHeader), GetType(FormPageHeader), GetType(FormFooter), GetType(FormPDF), GetType(FormExport)}
            End Get
        End Property
        Public Overridable ReadOnly Property FormattingFields As List(Of Type)
            Get
                Return New List(Of Type)
            End Get
        End Property
        Public Overridable Sub ChangeTransparent()
            If Not IsNothing(BorderFormItem) Then
                BorderFormItem.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
            End If
        End Sub
        Public Sub Click()
            Dim oFormHeaderTitle As Controls.TextBlock = Root.FormHeaderTitle
            oFormHeaderTitle.Text = Name
            SelectedItem = Me

            FilterDisplay()
            RightSetAllowedFields()
            SetBindings()
            Display()
        End Sub
        Public Sub ContentChanged()
            If Not IsNothing(FormPDF) Then
                ChangeTransparent()
                Display()
                TitleChanged()
                FormPDF.Changed = True
            End If
        End Sub
        Public Sub RightSetAllowedFields()
            ' sets the list of allowable fields
            ' all field templates have an empty guid
            Using oSuspender As New Suspender()
                Dim oGridMain As Controls.Grid = Root.GridMain
                Dim oGridFieldContents As Controls.Grid = Root.GridFieldContents
                oGridFieldContents.Children.Clear()
                oGridFieldContents.RowDefinitions.Clear()

                ' combine formatting fields with that of parent section
                Dim oFormattingFields As New List(Of Type)
                oFormattingFields.AddRange(FormattingFields)
                Dim oSection As FormSection = FindParent(GetType(FormSection))
                If (Not IsNothing(oSection)) Then
                    oFormattingFields.AddRange(oSection.FormattingFields)
                    oFormattingFields = oFormattingFields.Distinct.ToList
                End If

                ' allowed fields are those on the right column which are dependent on the selected item type
                Dim oCheckedAllowedFields As New List(Of Type)
                If (Me.GetType.IsSubclassOf(GetType(FormField)) Or Me.GetType.IsSubclassOf(GetType(FormFormatter))) AndAlso (Not IsNothing(Parent)) Then
                    ' use allowed field from parent if this is a form field
                    oCheckedAllowedFields.AddRange(Parent.AllowedFields)
                Else
                    oCheckedAllowedFields.AddRange(AllowedFields)
                End If
                For i = oCheckedAllowedFields.Count - 1 To 0 Step -1
                    If MandatoryFields.Contains(oCheckedAllowedFields(i)) Then
                        oCheckedAllowedFields.RemoveAt(i)
                    ElseIf (Not IsNothing(Parent)) AndAlso MandatoryFields.Contains(oCheckedAllowedFields(i)) Then
                        oCheckedAllowedFields.RemoveAt(i)
                    ElseIf oFormattingFields.Contains(oCheckedAllowedFields(i)) Then
                        oCheckedAllowedFields.RemoveAt(i)
                    End If
                Next

                ' set row definitions
                For i = 0 To Math.Max(oCheckedAllowedFields.Count, MandatoryFields.Count + oFormattingFields.Count) - 1
                    Dim oRowDefinition As New Controls.RowDefinition
                    oRowDefinition.Height = GridLength.Auto
                    oGridFieldContents.RowDefinitions.Add(oRowDefinition)
                Next

                For i = 0 To MandatoryFields.Count - 1
                    Dim oType As Type = MandatoryFields(i)
                    Dim oPropInfoName As Reflection.PropertyInfo = oType.GetProperty("Name", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoIconName As Reflection.PropertyInfo = oType.GetProperty("IconName", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoTagGuid As Reflection.PropertyInfo = oType.GetProperty("TagGuid", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim sFormItemName As String = oPropInfoName.GetValue(Nothing, Nothing)
                    Dim sFormItemIconName As String = oPropInfoIconName.GetValue(Nothing, Nothing)
                    Dim sFormItemTagGuid As String = oPropInfoTagGuid.GetValue(Nothing, Nothing)
                    Dim oFormItemIcon As ImageSource = If(m_Icons.ContainsKey(sFormItemIconName), m_Icons(sFormItemIconName), Nothing)

                    Dim oFieldItemParam As New Tuple(Of String, Double, ImageSource, ImageSource)(sFormItemName, 1.0, oFormItemIcon, Nothing)
                    Dim oFieldItem As New Common.FieldItem
                    oFieldItem.FIContent = oFieldItemParam

                    oFieldItem.FIReference = New Tuple(Of Double, Double)(oGridMain.ActualWidth, oGridMain.ActualHeight)
                    Dim oBorderFormItem As New BorderFormItem(Guid.Empty, False, oType, sFormItemTagGuid)
                    With oBorderFormItem
                        .Child = oFieldItem
                        .BorderBrush = Brushes.Black
                        .BorderThickness = New Thickness(1)
                        .HorizontalAlignment = HorizontalAlignment.Stretch
                        .VerticalAlignment = VerticalAlignment.Stretch
                        .OpacityMask = Brushes.Black
                        .Margin = New Thickness(5, 0, 5, 0)
                        .Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                        .CornerRadius = New CornerRadius(oGridMain.ActualHeight * CornerRadiusMultiplier)
                    End With

                    Controls.Grid.SetRow(oBorderFormItem, i * 2)
                    Controls.Grid.SetColumn(oBorderFormItem, 1)
                    Controls.Grid.SetColumnSpan(oBorderFormItem, 1)
                    Controls.Grid.SetZIndex(oBorderFormItem, 1)

                    oBorderFormItem.Visibility = Visibility.Visible
                    oGridFieldContents.Children.Add(oBorderFormItem)

                    Dim oRowDefinitionSpacer As New Controls.RowDefinition
                    oRowDefinitionSpacer.Height = GridLength.Auto
                    oGridFieldContents.RowDefinitions.Add(oRowDefinitionSpacer)

                    Dim oBorderSpacer As New BorderSpacer(oGridMain.ActualHeight, False)
                    Controls.Grid.SetRow(oBorderSpacer, (i * 2) + 1)
                    Controls.Grid.SetColumn(oBorderSpacer, 1)
                    Controls.Grid.SetColumnSpan(oBorderSpacer, 1)
                    Controls.Grid.SetZIndex(oBorderSpacer, 1)

                    oGridFieldContents.Children.Add(oBorderSpacer)
                Next

                For i = 0 To oFormattingFields.Count - 1
                    Dim iFormattingRow As Integer = i + MandatoryFields.Count
                    Dim oType As Type = oFormattingFields(i)
                    Dim oPropInfoName As Reflection.PropertyInfo = oType.GetProperty("Name", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoIconName As Reflection.PropertyInfo = oType.GetProperty("IconName", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoTagGuid As Reflection.PropertyInfo = oType.GetProperty("TagGuid", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim sFormItemName As String = oPropInfoName.GetValue(Nothing, Nothing)
                    Dim sFormItemIconName As String = oPropInfoIconName.GetValue(Nothing, Nothing)
                    Dim sFormItemTagGuid As String = oPropInfoTagGuid.GetValue(Nothing, Nothing)
                    Dim oFormItemIcon As ImageSource = If(m_Icons.ContainsKey(sFormItemIconName), m_Icons(sFormItemIconName), Nothing)

                    Dim oFieldItemParam As New Tuple(Of String, Double, ImageSource, ImageSource)(sFormItemName, 1.0, oFormItemIcon, Nothing)
                    Dim oFieldItem As New Common.FieldItem
                    oFieldItem.FIContent = oFieldItemParam

                    oFieldItem.FIReference = New Tuple(Of Double, Double)(oGridMain.ActualWidth, oGridMain.ActualHeight)
                    Dim oBorderFormItem As New BorderFormItem(Guid.Empty, False, oType, sFormItemTagGuid)
                    With oBorderFormItem
                        .Child = oFieldItem
                        .BorderBrush = Brushes.Black
                        .BorderThickness = New Thickness(1)
                        .HorizontalAlignment = HorizontalAlignment.Stretch
                        .VerticalAlignment = VerticalAlignment.Stretch
                        .OpacityMask = Brushes.Black
                        .Margin = New Thickness(5, 0, 5, 0)
                        .Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                        .CornerRadius = New CornerRadius(oGridMain.ActualHeight * CornerRadiusMultiplier)
                    End With

                    Controls.Grid.SetRow(oBorderFormItem, iFormattingRow * 2)
                    Controls.Grid.SetColumn(oBorderFormItem, 1)
                    Controls.Grid.SetColumnSpan(oBorderFormItem, 1)
                    Controls.Grid.SetZIndex(oBorderFormItem, 1)

                    oBorderFormItem.Visibility = Visibility.Visible
                    oGridFieldContents.Children.Add(oBorderFormItem)

                    Dim oRowDefinitionSpacer As New Controls.RowDefinition
                    oRowDefinitionSpacer.Height = GridLength.Auto
                    oGridFieldContents.RowDefinitions.Add(oRowDefinitionSpacer)

                    Dim oBorderSpacer As New BorderSpacer(oGridMain.ActualHeight, False)
                    Controls.Grid.SetRow(oBorderSpacer, (iFormattingRow * 2) + 1)
                    Controls.Grid.SetColumn(oBorderSpacer, 1)
                    Controls.Grid.SetColumnSpan(oBorderSpacer, 1)
                    Controls.Grid.SetZIndex(oBorderSpacer, 1)

                    oGridFieldContents.Children.Add(oBorderSpacer)
                Next

                For i = 0 To oCheckedAllowedFields.Count - 1
                    Dim oType As Type = oCheckedAllowedFields(i)
                    Dim oPropInfoName As Reflection.PropertyInfo = oType.GetProperty("Name", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoIconName As Reflection.PropertyInfo = oType.GetProperty("IconName", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim oPropInfoTagGuid As Reflection.PropertyInfo = oType.GetProperty("TagGuid", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy)
                    Dim sFormItemName As String = oPropInfoName.GetValue(Nothing, Nothing)
                    Dim sFormItemIconName As String = oPropInfoIconName.GetValue(Nothing, Nothing)
                    Dim sFormItemTagGuid As String = oPropInfoTagGuid.GetValue(Nothing, Nothing)
                    Dim oFormItemIcon As ImageSource = If(m_Icons.ContainsKey(sFormItemIconName), m_Icons(sFormItemIconName), Nothing)
                    Dim oFieldItemParam As New Tuple(Of String, Double, ImageSource, ImageSource)(sFormItemName, 1.0, oFormItemIcon, oImageStore.GetImage(oType))
                    Dim oFieldItem As New Common.FieldItem
                    oFieldItem.FIContent = oFieldItemParam

                    oFieldItem.FIReference = New Tuple(Of Double, Double)(oGridMain.ActualWidth, oGridMain.ActualHeight)
                    Dim oBorderFormItem As New BorderFormItem(Guid.Empty, False, oType, sFormItemTagGuid)
                    With oBorderFormItem
                        .Child = oFieldItem
                        .BorderBrush = Brushes.Black
                        .BorderThickness = New Thickness(1)
                        .HorizontalAlignment = HorizontalAlignment.Stretch
                        .VerticalAlignment = VerticalAlignment.Stretch
                        .OpacityMask = Brushes.Black
                        .Margin = New Thickness(5, 0, 5, 0)
                        .Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                        .CornerRadius = New CornerRadius(oGridMain.ActualHeight * CornerRadiusMultiplier)
                    End With

                    Controls.Grid.SetRow(oBorderFormItem, i * 2)
                    Controls.Grid.SetColumn(oBorderFormItem, 3)
                    Controls.Grid.SetColumnSpan(oBorderFormItem, 1)
                    Controls.Grid.SetZIndex(oBorderFormItem, 1)

                    oBorderFormItem.Visibility = Visibility.Visible
                    oGridFieldContents.Children.Add(oBorderFormItem)

                    Dim oRowDefinitionSpacer As New Controls.RowDefinition
                    oRowDefinitionSpacer.Height = GridLength.Auto
                    oGridFieldContents.RowDefinitions.Add(oRowDefinitionSpacer)

                    Dim oBorderSpacer As New BorderSpacer(oGridMain.ActualHeight, False)
                    Controls.Grid.SetRow(oBorderSpacer, (i * 2) + 1)
                    Controls.Grid.SetColumn(oBorderSpacer, 3)
                    Controls.Grid.SetColumnSpan(oBorderSpacer, 1)
                    Controls.Grid.SetZIndex(oBorderSpacer, 1)

                    oGridFieldContents.Children.Add(oBorderSpacer)
                Next

                For i = 0 To SectionList.Count - 1
                    Dim oCurrentSection As FormSection = (From oFormItem In SectionList(i) Where oFormItem.GetType.Equals(GetType(FormSection)) Select oFormItem).First
                    Dim iSectionRow As Integer = i + oCheckedAllowedFields.Count
                    Dim oType As Type = GetType(FormSection)
                    Dim oFieldItemParam As New Tuple(Of String, Double, ImageSource, ImageSource)(FormSection.Name, 1.0, m_Icons(FormSection.IconName), oImageStore.GetImage(oType))
                    Dim oFieldItem As New Common.FieldItem
                    oFieldItem.FIContent = oFieldItemParam

                    ' set binding for title
                    oFieldItem.DataContext = oCurrentSection
                    Dim oBinding As New Data.Binding
                    oBinding.Path = New PropertyPath("Title")
                    oBinding.Mode = Data.BindingMode.OneWay
                    oBinding.UpdateSourceTrigger = Data.UpdateSourceTrigger.PropertyChanged
                    oFieldItem.SetBinding(Common.FieldItem.FITitleProperty, oBinding)

                    oFieldItem.FIReference = New Tuple(Of Double, Double)(oGridMain.ActualWidth, oGridMain.ActualHeight)
                    Dim oBorderFormItem As New BorderFormItem(oCurrentSection.GUID, False, oType, FormSection.TagGuid)
                    With oBorderFormItem
                        .Child = oFieldItem
                        .BorderBrush = Brushes.Red
                        .BorderThickness = New Thickness(1)
                        .HorizontalAlignment = HorizontalAlignment.Stretch
                        .VerticalAlignment = VerticalAlignment.Stretch
                        .OpacityMask = Brushes.Black
                        .Margin = New Thickness(5, 0, 5, 0)
                        .Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
                        .CornerRadius = New CornerRadius(oGridMain.ActualHeight * CornerRadiusMultiplier)
                    End With

                    Controls.Grid.SetRow(oBorderFormItem, iSectionRow * 2)
                    Controls.Grid.SetColumn(oBorderFormItem, 3)
                    Controls.Grid.SetColumnSpan(oBorderFormItem, 1)
                    Controls.Grid.SetZIndex(oBorderFormItem, 1)

                    oBorderFormItem.Visibility = Visibility.Visible
                    oGridFieldContents.Children.Add(oBorderFormItem)

                    Dim oRowDefinitionSpacer As New Controls.RowDefinition
                    oRowDefinitionSpacer.Height = GridLength.Auto
                    oGridFieldContents.RowDefinitions.Add(oRowDefinitionSpacer)

                    Dim oBorderSpacer As New BorderSpacer(oGridMain.ActualHeight, False)
                    Controls.Grid.SetRow(oBorderSpacer, (iSectionRow * 2) + 1)
                    Controls.Grid.SetColumn(oBorderSpacer, 3)
                    Controls.Grid.SetColumnSpan(oBorderSpacer, 1)
                    Controls.Grid.SetZIndex(oBorderSpacer, 1)

                    oGridFieldContents.Children.Add(oBorderSpacer)
                Next
            End Using
        End Sub
        Public Function FindParent(ByVal oType As Type) As BaseFormItem
            ' finds a parent of the stated type, or returns nothing if not found
            Dim oCurrentParent As BaseFormItem = Parent
            Do Until IsNothing(oCurrentParent) OrElse oCurrentParent.GetType.Equals(oType)
                oCurrentParent = oCurrentParent.Parent
            Loop
            Return oCurrentParent
        End Function
        Public Function ExtractFields(Of T As BaseFormItem)(ByVal oBaseType As Type) As List(Of T)
            ' extracts form fields based on a supplied type
            If GetType(T).Equals(oBaseType) Then
                Return (From oChild In Children Where oChild.Value.Equals(GetType(T)) OrElse oChild.Value.IsSubclassOf(GetType(T)) Select CType(FormMain.FindChild(oChild.Key), T)).ToList
            Else
                Return (From oChild In Children Where oChild.Value.Equals(GetType(T)) Select CType(FormMain.FindChild(oChild.Key), T)).ToList
            End If
        End Function
        Public Function GetFormItems(Of T As BaseFormItem)() As List(Of T)
            ' iteratively gets child form items of the specified type
            Dim oFormItemList As New List(Of T)
            For Each oChildGUID In Children.Keys
                IterateGetFormItems(FormMain.FindChild(oChildGUID), GetType(T), oFormItemList)
            Next
            Return oFormItemList
        End Function
        Public Function GetFormItems(ByVal oType As Type) As List(Of BaseFormItem)
            ' iteratively gets child form items
            ' if the supplied type is nothing, then all child items are added
            Dim oFormItemList As New List(Of BaseFormItem)
            For Each oChildGUID In Children.Keys
                IterateGetFormItems(FormMain.FindChild(oChildGUID), oType, oFormItemList)
            Next
            Return oFormItemList
        End Function
        Private Function GetNumbering(ByVal bSubNumbering As Boolean) As String
            ' gets numbering for this item
            ' gets the numbering string which is a sequence of numbers separated by dots indicating the child position (1-based)
            If IsNothing(m_DataObject.m_Parent) Then
                Return String.Empty
            Else
                Dim oFilteredChildDictionary As Dictionary(Of Guid, Type) = Nothing
                Select Case m_DataObject.m_Parent.GetType
                    Case GetType(FormBody)
                        ' only number sections
                        If Not Me.GetType.Equals(GetType(FormSection)) Then
                            Return String.Empty
                        End If
                        oFilteredChildDictionary = (From oChild In m_DataObject.m_Parent.Children Where (Not m_DataObject.m_Parent.IgnoreFields.Contains(oChild.Value)) AndAlso oChild.Value.Equals(GetType(FormSection)) Select oChild).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    Case Else
                        oFilteredChildDictionary = (From oChild In m_DataObject.m_Parent.Children Where Not m_DataObject.m_Parent.IgnoreFields.Contains(oChild.Value) Select oChild).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                End Select

                If IsNothing(m_DataObject.m_Parent) OrElse m_DataObject.m_Parent.GetType.Equals(GetType(FormMain)) Then
                    Return String.Empty
                Else
                    Dim sNumberingString As String = (oFilteredChildDictionary.Keys.ToList.IndexOf(m_DataObject.m_GUID) + 1).ToString
                    Dim sParentNumbering As String = m_DataObject.m_Parent.GetNumbering(False)
                    If sParentNumbering <> String.Empty Then
                        sNumberingString = sParentNumbering + "." + sNumberingString
                    End If

                    ' add subnumbering for MCQ and subsections
                    If bSubNumbering Then
                        Dim oSection As FormSection = FindParent(GetType(FormSection))
                        If (Not IsNothing(oSection)) AndAlso ParamList.ContainsKey(ParamList.KeyNumberingCurrent) Then
                            With oSection
                                Dim sSubNumbering As String = PDFHelper.GetNumbering(ParamList.Value(ParamList.KeyNumberingCurrent), If(ParamList.ContainsKey(ParamList.KeySubNumberingCurrent), ParamList.Value(ParamList.KeySubNumberingCurrent), -1), .NumberingType)
                                sNumberingString += "(" + sSubNumbering + ")"
                            End With
                        End If
                    End If

                    Return sNumberingString
                End If
            End If
        End Function
        Private Sub IterateGetFormItems(Of T As BaseFormItem)(ByVal oFormItem As BaseFormItem, ByVal oType As Type, ByRef oFormItemList As List(Of T))
            If IsNothing(oType) Then
                oFormItemList.Add(oFormItem)
            Else
                If oFormItem.GetType.Equals(oType) Then
                    oFormItemList.Add(oFormItem)
                End If
            End If
            For Each oChildGUID In oFormItem.Children.Keys
                IterateGetFormItems(FormMain.FindChild(oChildGUID), oType, oFormItemList)
            Next
        End Sub
        Private Sub FilterDisplay()
            ' runs through all the display items in the grid and set to visible those specified by the DisplayFilter property
            For Each sElementName As String In DisplayDictionary.Keys
                If DisplayFilter.Contains(sElementName) Then
                    DisplayDictionary(sElementName).Visibility = Visibility.Visible
                Else
                    Select Case DisplayDictionary(sElementName).GetType
                        Case GetType(Controls.DockPanel), GetType(PDFViewer.PDFViewer), GetType(Shapes.Rectangle), GetType(Controls.Canvas), GetType(Controls.Image), GetType(Controls.ScrollViewer)
                            DisplayDictionary(sElementName).Visibility = Visibility.Hidden
                        Case GetType(Controls.StackPanel), GetType(Common.HighlightTextBox), GetType(Controls.ItemsControl), GetType(Controls.DataGrid)
                            DisplayDictionary(sElementName).Visibility = Visibility.Collapsed
                        Case GetType(Controls.Grid)
                            DisplayDictionary(sElementName).Visibility = Visibility.Visible
                        Case Else
                            DisplayDictionary(sElementName).Visibility = Visibility.Hidden
                    End Select
                End If
            Next

            ' screen grids with collapsed contents and set to collapse
            For Each sElementName As String In DisplayDictionary.Keys
                If DisplayDictionary(sElementName).GetType.Equals(GetType(Controls.Grid)) Then
                    Dim oGrid As Controls.Grid = DisplayDictionary(sElementName)
                    If oGrid.Name <> String.Empty Then
                        Dim oVisibility As Visibility = Visibility.Collapsed
                        For Each oChild In oGrid.Children
                            If oChild.GetType.Equals(GetType(Controls.DockPanel)) Then
                                Dim oDockPanel As Controls.DockPanel = oChild
                                For Each oDockChild As FrameworkElement In oDockPanel.Children
                                    If oDockChild.Visibility = Visibility.Visible Then
                                        oVisibility = Visibility.Visible
                                    End If
                                Next
                            End If
                        Next
                        oGrid.Visibility = oVisibility
                    End If
                End If
            Next
        End Sub
        Private Shared Sub IterateChildrenLevel(ByVal oGUID As Guid)
            ' when parent is changed, adjusts the level of all children
            Dim oFormItem As BaseFormItem = FormMain.FindChild(oGUID)

            For Each oChildGUID As Guid In oFormItem.Children.Keys
                Dim oChildFormItem As BaseFormItem = FormMain.FindChild(oChildGUID)
                If oFormItem.Level = -1 Then
                    oChildFormItem.Level = -1
                Else
                    oChildFormItem.Level = oFormItem.Level + 1
                End If
                IterateChildrenLevel(oChildGUID)
            Next
        End Sub
        Private Shared Sub IterateRemoveChildren(ByVal oGUID As Guid)
            ' when parent is removed, remove all children
            Dim oFormItem As BaseFormItem = FormMain.FindChild(oGUID)
            For Each oChildGUID As Guid In oFormItem.Children.Keys
                Dim oChildFormItem As BaseFormItem = FormMain.FindChild(oChildGUID)
                IterateRemoveChildren(oChildGUID)
                FormMain.RemoveChild(oChildFormItem)
            Next
        End Sub
        Public Class Suspender
            ' suspends
            Implements IDisposable

            Private Shared m_SuspendDictionary As New Dictionary(Of BaseFormItem, Tuple(Of Integer, Boolean, Boolean))
            Private Shared m_GlobalSuspendState As Boolean = False
            Private Shared m_PreventReentrancy As Boolean = False
            Private m_FormItem As BaseFormItem = Nothing
            Private m_PreviousGlobalSuspendState As Boolean = False

            Sub New()
                If Not m_PreventReentrancy Then
                    m_FormItem = Nothing
                    m_PreviousGlobalSuspendState = GlobalSuspendUpdates
                    GlobalSuspendUpdates = True
                End If
            End Sub
            Sub New(ByVal oFormItem As BaseFormItem, ByVal bContentChanged As Boolean, Optional ByVal bInvalidateSection As Boolean = False)
                If Not m_PreventReentrancy Then
                    m_FormItem = oFormItem
                    If Not m_SuspendDictionary.ContainsKey(m_FormItem) Then
                        m_SuspendDictionary.Add(m_FormItem, New Tuple(Of Integer, Boolean, Boolean)(0, bContentChanged, bInvalidateSection))
                    End If
                    m_SuspendDictionary(m_FormItem) = New Tuple(Of Integer, Boolean, Boolean)(m_SuspendDictionary(m_FormItem).Item1 + 1, m_SuspendDictionary(m_FormItem).Item2, bInvalidateSection)
                End If
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
                        m_PreventReentrancy = True
                        oFormItem.ContentChanged()
                        m_PreventReentrancy = False
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
                        If Not m_PreventReentrancy Then
                            If IsNothing(m_FormItem) Then
                                ' global suspend
                                GlobalSuspendUpdates = m_PreviousGlobalSuspendState
                            Else
                                If m_SuspendDictionary(m_FormItem).Item2 Then
                                    ' invalidates section display store
                                    Dim oSection As FormSection = TryCast(m_FormItem, FormSection)
                                    If IsNothing(oSection) Then
                                        oSection = m_FormItem.FindParent(GetType(FormSection))
                                    End If
                                    If Not IsNothing(oSection) Then
                                        oSection.ImageTracker.Invalidate(oSection, True)
                                    End If
                                End If

                                ' item suspend
                                m_SuspendDictionary(m_FormItem) = New Tuple(Of Integer, Boolean, Boolean)(m_SuspendDictionary(m_FormItem).Item1 - 1, m_SuspendDictionary(m_FormItem).Item2, m_SuspendDictionary(m_FormItem).Item3)
                                If (Not m_GlobalSuspendState) AndAlso m_SuspendDictionary(m_FormItem).Item1 <= 0 Then
                                    RemoveItem(m_FormItem)
                                End If
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
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Shadows Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    For i = Children.Keys.Count - 1 To 0 Step -1
                        Dim oFormItem As BaseFormItem = FormMain.FindChild(Children.Keys(i))
                        oFormItem.Dispose()
                        FormMain.RemoveChild(oFormItem)
                        Children.Remove(Children.Keys(i))
                    Next

                    Suspender.DeleteEntry(Me)
                    Parent(, False) = Nothing
                End If
            End If
            disposedValue = True
        End Sub
        Public Shadows Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
    Public Class BorderFormItem
        Inherits Controls.Border

        Private m_GUID As Guid
        Private m_Type As Type

        Sub New(ByVal oGUID As Guid, ByVal bAllowDrop As Boolean, ByVal oType As Type, ByVal sTag As String)
            m_GUID = oGUID
            m_Type = oType
            Me.AllowDrop = bAllowDrop
            Me.Tag = sTag
        End Sub
        Public ReadOnly Property GUID As Guid
            Get
                Return m_GUID
            End Get
        End Property
#Region "UI"
        Private Sub MouseMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles Me.MouseMove
            If e.LeftButton = Input.MouseButtonState.Pressed Then
                ' drag and drop operation move
                DragDrop.DoDragDrop(Me, New DataObject(DragBaseFormItem, New Tuple(Of Type, Guid, Boolean, String)(m_Type, m_GUID, False, Tag)), DragDropEffects.Move)
            ElseIf e.RightButton = Input.MouseButtonState.Pressed Then
                ' drag and drop operation copy
                DragDrop.DoDragDrop(Me, New DataObject(DragBaseFormItem, New Tuple(Of Type, Guid, Boolean, String)(m_Type, m_GUID, True, Tag)), DragDropEffects.Move)
            Else
                ' highlight semi-transparent
                Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            End If
        End Sub
        Private Sub MouseLeaveHandler(sender As Object, e As Input.MouseEventArgs) Handles Me.MouseLeave
            ' changes background colour to transparent
            ChangeTransparent()
        End Sub
        Private Sub BorderClick(sender As Object, e As EventArgs) Handles Me.MouseLeftButtonDown, Me.TouchDown
            ' triggers on clicking a formitem
            Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(m_GUID)
            If Not IsNothing(oFormItem) Then
                oFormItem.Click()
            End If

            If e.GetType.Equals(GetType(Input.TouchEventArgs)) Then
                CType(e, Input.TouchEventArgs).Handled = True
            End If
        End Sub
        Public Sub BorderFormItem_Drop(sender As Object, e As DragEventArgs, Optional ByVal iInsertAfter As Integer = -2) Handles Me.Drop
            ' handles drop event
            If e.Data.GetDataPresent(DragBaseFormItem) Then
                Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)
                If DropAllowed(oDragData) Then
                    Dim oDropFormItem As BaseFormItem = Nothing
                    Using oSuspender As New BaseFormItem.Suspender()
                        Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(GUID)
                        Dim oFormField As FormField = Nothing
                        Using oSuspenderFormItem As New BaseFormItem.Suspender(oFormItem, True)
                            If oDragData.Item2.Equals(Guid.Empty) Then
                                ' drag from right panel
                                oDropFormItem = Activator.CreateInstance(oDragData.Item1)
                            Else
                                ' check if item is from the section list
                                Dim oSectionList As List(Of List(Of BaseFormItem)) = (From oCurrentFormItemList As List(Of BaseFormItem) In BaseFormItem.SectionList From oCurrentFormItem As BaseFormItem In oCurrentFormItemList Where oCurrentFormItem.GUID.Equals(oDragData.Item2) Select oCurrentFormItemList).ToList
                                If oSectionList.Count > 0 Then
                                    Dim oSection As FormSection = (From oCurrentFormItem As BaseFormItem In oSectionList.First Where oCurrentFormItem.GUID.Equals(oDragData.Item2) Select oCurrentFormItem).First
                                    oDropFormItem = oSection
                                    BaseFormItem.SectionList.Remove(oSectionList.First)

                                    ' add all child items to form main dictionary
                                    For Each oChildFormItem As BaseFormItem In oSectionList.First
                                        Dim oNewGUID As Guid = Guid.NewGuid
                                        ' replace guid in parent list
                                        Dim oParentList As List(Of BaseFormItem) = (From oCurrentFormItem As BaseFormItem In oSectionList.First From oChild In oCurrentFormItem.Children Where oChild.Key.Equals(oChildFormItem.GUID) Select oCurrentFormItem).ToList
                                        For Each oParentItem In oParentList
                                            ' remove guid from parent list
                                            oParentItem.Children.Remove(oChildFormItem.GUID)

                                            ' add new guid to parent list
                                            oParentItem.Children.Add(oNewGUID, oChildFormItem.GetType)
                                        Next

                                        ' change the guid of the form item
                                        oChildFormItem.GUID = oNewGUID

                                        ' add to the form main list
                                        BaseFormItem.FormMain.AddChild(oChildFormItem)
                                    Next
                                Else
                                    ' drag from right panel
                                    oDropFormItem = BaseFormItem.FormMain.FindChild(oDragData.Item2)
                                    Dim oSection As FormSection = oDropFormItem.FindParent(GetType(FormSection))
                                    If Not IsNothing(oSection) Then
                                        oSection.ImageTracker.Clear()
                                    End If
                                End If
                            End If

                            If (Not oDragData.Item2.Equals(Guid.Empty)) AndAlso oDragData.Item3 Then
                                ' drag copy
                                Dim oNewFormItem As BaseFormItem = oDropFormItem.Clone
                                oNewFormItem.Parent(iInsertAfter, oDragData.Item3) = oFormItem
                            Else
                                oFormField = TryCast(oDropFormItem, FormField)
                                If Not IsNothing(oFormField) Then
                                    Dim iCount As Integer = Aggregate oChild In oFormItem.Children Where oChild.Value.Equals(oDragData.Item1) AndAlso CType(BaseFormItem.FormMain.FindChild(oChild.Key), FormField).WholeBlock Into Count
                                    If iCount > 0 Then
                                        oFormField.WholeBlock = False
                                    End If
                                End If

                                ' drag move
                                oDropFormItem.Parent(iInsertAfter, oDragData.Item3) = oFormItem
                            End If
                            oFormItem.Expanded = True
                        End Using
                        LeftArrangeFormItems()
                        oDropFormItem.BorderFormItem.BringIntoView()
                        oDropFormItem.RightSetAllowedFields()

                        ' processing for form fields
                        If Not IsNothing(oFormField) Then
                            oFormField.ChangeTransparent()
                        End If
                    End Using
                    oDropFormItem.Click()
                End If
            End If
        End Sub
        Private Sub BorderFormItem_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
            ' changes background colour to semi-transparent
            If e.Data.GetDataPresent(DragBaseFormItem) Then
                Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)

                If DropAllowed(oDragData) Then
                    Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
                Else
                    Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
                End If
            Else
                Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
            End If
        End Sub
        Private Sub BorderFormItem_DragLeave(sender As Object, e As DragEventArgs) Handles Me.DragLeave
            ' changes background colour to transparent
            ChangeTransparent()
        End Sub
        Private Sub ChangeTransparent()
            Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(m_GUID)
            If IsNothing(oFormItem) Then
                Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &HA5, &H0))
            Else
                oFormItem.ChangeTransparent()
            End If
        End Sub
        Public Function DropAllowed(ByVal oDragData As Tuple(Of Type, Guid, Boolean, String)) As Boolean
            ' check to see if a drop is allowed
            Dim bAllowed As Boolean = False

            ' only allow drops to non field templates
            Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(GUID)

            ' if the form item is being moved to the same parent but a different position, the allow
            If (Not oDragData.Item3) AndAlso oFormItem.Children.Keys.Contains(oDragData.Item2) Then
                bAllowed = True
            Else
                If oFormItem.AllowedFields.Contains(oDragData.Item1) Then
                    If oDragData.Item1.IsSubclassOf(GetType(FormField)) Then
                        If (oFormItem.GetType.Equals(GetType(FormFormHeader)) Or oFormItem.GetType.Equals(GetType(FormSection))) Then
                            ' for these two types, only allow a single type of each field
                            If Not oFormItem.Children.Values.Contains(oDragData.Item1) Then
                                bAllowed = True
                            End If
                        Else
                            ' the drop item is a field
                            Dim oDropFieldItem As FormField = Activator.CreateInstance(oDragData.Item1)
                            If (Not oDropFieldItem.SingleOnly(oFormItem)) OrElse (Not oFormItem.Children.Values.Contains(oDragData.Item1)) Then
                                ' field is not single only, or the form item does not contain one of these fields
                                bAllowed = True
                            End If
                        End If
                    ElseIf oDragData.Item1.IsSubclassOf(GetType(FormFormatter)) Then
                        ' the drop item is a format field
                        Dim oDropFormatterItem As FormFormatter = Activator.CreateInstance(oDragData.Item1)
                        If (Not oDropFormatterItem.SingleOnly(oFormItem)) OrElse (Not oFormItem.Children.Values.Contains(oDragData.Item1)) Then
                            ' field is not single only, or the form item does not contain one of these fields
                            bAllowed = True
                        End If
                    Else
                        ' the drop item is not a field
                        bAllowed = True
                    End If
                End If
            End If

            Return bAllowed
        End Function
#End Region
    End Class
    Public Class ImageExpander
        Inherits Controls.Image

        Private m_Grid As Controls.Grid
        Private m_FormItem As BaseFormItem
        Private m_Expanded As Boolean

        Public Sub New(ByRef oGrid As Controls.Grid, ByVal oFormItem As BaseFormItem)
            MyBase.New()
            m_Grid = oGrid

            With Me
                .HorizontalAlignment = HorizontalAlignment.Center
                .VerticalAlignment = VerticalAlignment.Center
                .Margin = New Thickness(0)
                .FormItem = oFormItem
            End With
        End Sub
        Public Property FormItem As BaseFormItem
            Get
                Return m_FormItem
            End Get
            Set(value As BaseFormItem)
                m_FormItem = value
                m_Expanded = m_FormItem.Expanded
                If m_Expanded Then
                    Source = m_Icons("CCMMinus")
                Else
                    Source = m_Icons("CCMPlus")
                End If
            End Set
        End Property
        Public Property Expanded As Boolean
            Get
                Return m_Expanded
            End Get
            Set(value As Boolean)
                m_Expanded = value
                m_FormItem.Expanded = value
                If m_Expanded Then
                    Source = m_Icons("CCMMinus")
                Else
                    Source = m_Icons("CCMPlus")
                End If
            End Set
        End Property
        Private Sub ExpanderClick(sender As Object, e As EventArgs) Handles Me.MouseLeftButtonDown, Me.TouchDown
            ' expands or collapses formitem
            Dim oAction As Action = Sub()
                                        Expanded = Not Expanded
                                        LeftArrangeFormItems()
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    End Class
    Public Class BorderSpacer
        Inherits Controls.Border

        Private m_ReferenceHeight As Double
        Private m_BorderFormItem As BorderFormItem

        Public Sub New(ByVal fReferenceHeight As Double, ByVal bAllowDrop As Boolean, Optional ByVal oBorderFormItem As BorderFormItem = Nothing)
            MyBase.New()
            m_ReferenceHeight = fReferenceHeight
            m_BorderFormItem = oBorderFormItem

            With Me
                .BorderBrush = Brushes.Black
                .BorderThickness = New Thickness(0)
                .HorizontalAlignment = HorizontalAlignment.Stretch
                .VerticalAlignment = VerticalAlignment.Stretch
                .Margin = New Thickness(0)
                .Background = Brushes.Transparent
                .OpacityMask = Brushes.Black
                .Height = m_ReferenceHeight * BorderSpacerMultiplier
                .AllowDrop = bAllowDrop
            End With
        End Sub
        Private Sub BorderSpacer_Drop(sender As Object, e As DragEventArgs) Handles Me.Drop
            ' changes background colour to transparent
            Me.Background = New SolidColorBrush(Colors.Transparent)

            ' handles drop event
            If (Not IsNothing(m_BorderFormItem)) AndAlso (Not m_BorderFormItem.GUID.Equals(Guid.Empty)) Then
                Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(m_BorderFormItem.GUID)

                ' check for first spacer
                ' this is determined by a drop referencing a form item which has children
                If oFormItem.Expanded AndAlso oFormItem.Children.Count > 0 Then
                    oFormItem.BorderFormItem.BorderFormItem_Drop(sender, e, -1)
                Else
                    Dim oFormItemParent As BaseFormItem = oFormItem.Parent
                    If (Not IsNothing(oFormItemParent)) AndAlso (Not IsNothing(oFormItemParent.BorderFormItem)) Then
                        Dim iInsertAfter As Integer = oFormItemParent.Children.Keys.ToList.IndexOf(m_BorderFormItem.GUID)
                        oFormItemParent.BorderFormItem.BorderFormItem_Drop(sender, e, iInsertAfter)
                    End If
                End If
            End If
        End Sub
        Private Sub BorderSpacer_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
            ' changes background colour to semi-transparent
            If e.Data.GetDataPresent(DragBaseFormItem) Then
                Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)

                If (Not IsNothing(m_BorderFormItem)) AndAlso (Not m_BorderFormItem.GUID.Equals(Guid.Empty)) Then
                    Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(m_BorderFormItem.GUID)

                    ' check for first spacer
                    ' this is determined by a drop referencing a form item which has children
                    If oFormItem.Expanded AndAlso oFormItem.Children.Count > 0 Then
                        If oFormItem.BorderFormItem.DropAllowed(oDragData) Then
                            Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
                        Else
                            Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
                        End If
                    Else
                        Dim oFormItemParent As BaseFormItem = oFormItem.Parent
                        If (Not IsNothing(oFormItemParent)) AndAlso (Not IsNothing(oFormItemParent.BorderFormItem)) Then
                            Dim iInsertAfter As Integer = oFormItemParent.Children.Keys.ToList.IndexOf(m_BorderFormItem.GUID)
                            If oFormItemParent.BorderFormItem.DropAllowed(oDragData) Then
                                Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
                            Else
                                Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
                            End If
                        Else
                            Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
                        End If
                    End If
                End If
            Else
                Me.Background = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
            End If
        End Sub
        Private Sub BorderSpacer_DragLeave(sender As Object, e As DragEventArgs) Handles Me.DragLeave
            ' changes background colour to transparent
            Me.Background = New SolidColorBrush(Colors.Transparent)
        End Sub
    End Class
    Public Class ImageStore
        Implements IDisposable

        Private m_Images As Dictionary(Of String, Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, Media.PixelFormat))

        Public Sub New()
            m_Images = New Dictionary(Of String, Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, Media.PixelFormat))
        End Sub
        Public Sub Clear()
            ' clears store and reclaims resources
            For i = m_Images.Count - 1 To 0 Step -1
                m_Images(m_Images.Keys(i)) = Nothing
                m_Images.Remove(m_Images.Keys(i))
            Next
            m_Images.Clear()
            CommonFunctions.ClearMemory()
        End Sub
        Public Sub AddImage(ByVal oType As Type, ByVal oImageSource As Imaging.BitmapSource)
            ' adds image to store
            Dim sKey As String = oType.ToString
            If m_Images.ContainsKey(sKey) Then
                m_Images(sKey) = Converter.BitmapSourceToByteArray(oImageSource, False)
            Else
                m_Images.Add(sKey, Converter.BitmapSourceToByteArray(oImageSource, False))
            End If
        End Sub
        Public Function GetImage(ByVal oType As Type) As Imaging.BitmapSource
            ' gets image from store
            Dim sKey As String = oType.ToString
            If m_Images.ContainsKey(sKey) Then
                Return Converter.ByteArrayToBitmapSource(m_Images(sKey))
            Else
                Return Nothing
            End If
        End Function
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Shadows Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    Clear()
                End If
            End If
            disposedValue = True
        End Sub
        Public Shadows Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
    Public Class ParamList
        Implements IDisposable
        Private Shared m_ParamList As New Dictionary(Of String, Tuple(Of Integer, Object))
        Private m_KeyString As String = String.Empty

#Region "ParamKeys"
        ' param list keys
        Public Const KeyFieldCollection As String = "FieldCollection"
        Public Const KeyPageCount As String = "PageCount"
        Public Const KeyCurrentPage As String = "CurrentPage"
        Public Const KeyRemoveBarcode As String = "RemoveBarcode"
        Public Const KeyPDFPages As String = "PDFPages"
        Public Const KeyBarCodes As String = "BarCodes"
        Public Const KeyTopic As String = "Topic"
        Public Const KeyCurrentSubject As String = "CurrentSubject"
        Public Const KeyAppendText As String = "AppendText"
        Public Const KeyNumberingCurrent As String = "NumberingCurrent"
        Public Const KeySubNumberingCurrent As String = "SubNumberingCurrent"
        Public Const KeyTextArray As String = "TextArray"
        Public Const KeyFieldCount As String = "FieldCount"
        Public Const KeyCurrentSection As String = "CurrentSection"
        Public Const KeyPDFDocument As String = "PDFDocument"
        Public Const KeyXCurrentHeight As String = "XCurrentHeight"
        Public Const KeySuspendPagination As String = "SuspendPagination"
        Public Const KeyDisplacement As String = "Displacement"
#End Region

        Sub New(ByVal sKeyString As String, ByVal oValueObject As Object)
            If Trim(sKeyString) <> String.Empty Then
                If m_ParamList.ContainsKey(sKeyString) Then
                    m_ParamList(sKeyString) = New Tuple(Of Integer, Object)(m_ParamList(sKeyString).Item1 + 1, oValueObject)
                Else
                    m_ParamList.Add(sKeyString, New Tuple(Of Integer, Object)(1, oValueObject))
                End If
                m_KeyString = sKeyString
            End If
        End Sub
        Public Shared Property Value(ByVal sKeyString As String) As Object
            Get
                If m_ParamList.ContainsKey(sKeyString) Then
                    Return m_ParamList(sKeyString).Item2
                Else
                    Throw New KeyNotFoundException
                End If
            End Get
            Set(value As Object)
                If m_ParamList.ContainsKey(sKeyString) Then
                    m_ParamList(sKeyString) = New Tuple(Of Integer, Object)(m_ParamList(sKeyString).Item1, value)
                Else
                    Throw New KeyNotFoundException
                End If
            End Set
        End Property
        Public Shared Function ContainsKey(ByVal sKeyString As String) As Boolean
            Return m_ParamList.ContainsKey(sKeyString)
        End Function
        Public Shared Sub Remove(ByVal sKeyString As String)
            If m_ParamList.ContainsKey(sKeyString) Then
                m_ParamList.Remove(sKeyString)
            Else
                Throw New KeyNotFoundException
            End If
        End Sub
        Public Sub DisposeReturn(ByRef oValueObject As Object)
            ' call this to return a value before disposing
            oValueObject = m_ParamList(m_KeyString).Item2
        End Sub
        Public Shared Sub Clear()
            For i = m_ParamList.Count - 1 To 0 Step -1
                m_ParamList(m_ParamList.Keys(i)) = Nothing
                m_ParamList.Remove(m_ParamList.Keys(i))
            Next
            m_ParamList.Clear()
            CommonFunctions.ClearMemory()
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Shadows Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    If m_ParamList.ContainsKey(m_KeyString) Then
                        If m_ParamList(m_KeyString).Item1 > 1 Then
                            ' decrement count
                            m_ParamList(m_KeyString) = New Tuple(Of Integer, Object)(m_ParamList(m_KeyString).Item1 - 1, m_ParamList(m_KeyString).Item2)
                        Else
                            ' remove
                            m_ParamList.Remove(m_KeyString)
                        End If
                    Else
                        Throw New KeyNotFoundException
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
    <DataContract> Public Class TagClass
        Implements ICloneable

        ' class for sub-section tag strings in CCM
        <DataMember> Public TagStore As New ObservableCollection(Of TagItem)
        <DataMember> Private m_Name As String = String.Empty
        <DataMember> Private m_Header As String = String.Empty
        <DataMember> Private m_Width As Integer = 0

        Private Const StringUndefined As String = "Undefined"

        Sub New(ByVal sName As String, ByVal sHeader As String, ByVal iSetWidth As Integer)
            m_Name = sName
            m_Header = sHeader
            m_Width = Math.Max(1, iSetWidth)
        End Sub
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oTagClass As New TagClass(m_Name, m_Header, m_Width)
            With oTagClass
                For Each oRow As TagItem In TagStore
                    .TagStore.Add(oRow.Clone)
                Next
            End With
            Return oTagClass
        End Function
        Public Sub Clear()
            TagStore.Clear()
        End Sub
        Public Sub Remove(ByVal iIndex As Integer)
            If iIndex >= 0 And iIndex < TagStore.Count Then
                TagStore.RemoveAt(iIndex)
                Renumber()
            End If
        End Sub
        Public Sub Add(ByVal oTags As List(Of String), Optional ByVal iIndex As Integer = -1)
            Dim oNewRowArray(m_Width - 1) As String
            For i = 0 To oNewRowArray.Count - 1
                oNewRowArray(i) = String.Empty
            Next
            For i = 0 To Math.Min(oNewRowArray.Count - 1, oTags.Count - 1)
                oNewRowArray(i) = oTags(i)
            Next

            Dim oNewRow As New TagItem(0, oNewRowArray.ToList)
            If iIndex >= 0 And iIndex < TagStore.Count Then
                ' insert tags
                TagStore.Insert(iIndex, oNewRow)
                Renumber()
            Else
                ' append tags
                oNewRow.Number = (TagStore.Count + 1).ToString
                TagStore.Add(oNewRow)
            End If
        End Sub
        Public Sub Add(ByVal oTags As String(), Optional ByVal iIndex As Integer = -1)
            Add(oTags.ToList, iIndex)
        End Sub
        Public Property Name As String
            Get
                Return m_Name
            End Get
            Set(value As String)
                m_Name = value
            End Set
        End Property
        Public ReadOnly Property Header As String
            Get
                Return m_Header
            End Get
        End Property
        Public Property Width As Integer
            Get
                Return m_Width
            End Get
            Set(value As Integer)
                If value > 0 And value <> m_Width Then
                    m_Width = value
                    For i = 0 To TagStore.Count - 1
                        Dim oNewRowArray(m_Width - 1) As String
                        For j = 0 To oNewRowArray.Count - 1
                            oNewRowArray(j) = String.Empty
                        Next
                        For j = 0 To Math.Min(oNewRowArray.Count - 1, TagStore(i).Tags.Count - 1)
                            oNewRowArray(j) = TagStore(i).Tags(j)
                        Next

                        TagStore(i).Tags = oNewRowArray.ToList
                    Next
                End If
            End Set
        End Property
        Public ReadOnly Property Count As Integer
            Get
                Return TagStore.Count
            End Get
        End Property
        Public Property Tag(ByVal iIndex As Integer, ByVal iCount As Integer) As String
            Get
                If m_Width > 0 And TagStore.Count > 0 Then
                    Dim iActualIndex As Integer = 0
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > TagStore.Count - 1 Then
                        iActualIndex = TagStore.Count - 1
                    Else
                        iActualIndex = iIndex
                    End If

                    Dim iActualCount As Integer = 0
                    If iCount < 0 Then
                        iActualCount = 0
                    ElseIf iCount > m_Width - 1 Then
                        iActualCount = m_Width - 1
                    Else
                        iActualCount = iCount
                    End If

                    Return TagStore(iActualIndex).Tags(iActualCount)
                Else
                    Return StringUndefined
                End If
            End Get
            Set(value As String)
                If m_Width > 0 And TagStore.Count > 0 Then
                    Dim iActualIndex As Integer = 0
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > TagStore.Count - 1 Then
                        iActualIndex = TagStore.Count - 1
                    Else
                        iActualIndex = iIndex
                    End If

                    Dim iActualCount As Integer = 0
                    If iCount < 0 Then
                        iActualCount = 0
                    ElseIf iCount > m_Width - 1 Then
                        iActualCount = m_Width - 1
                    Else
                        iActualCount = iCount
                    End If

                    TagStore(iActualIndex).Tags(iActualCount) = value
                End If
            End Set
        End Property
        Public ReadOnly Property TagArray(ByVal iIndex As Integer) As String()
            Get
                If m_Width > 0 And TagStore.Count > 0 Then
                    Dim iActualIndex As Integer = 0
                    If iIndex < 0 Then
                        iActualIndex = 0
                    ElseIf iIndex > TagStore.Count - 1 Then
                        iActualIndex = TagStore.Count - 1
                    Else
                        iActualIndex = iIndex
                    End If

                    Return TagStore(iActualIndex).Tags.ToArray
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Private Sub Renumber()
            For i = 0 To TagStore.Count - 1
                TagStore(i).Number = (i + 1).ToString
            Next
        End Sub
        <DataContract> Public Class TagItem
            Implements ICloneable
            Implements INotifyPropertyChanged

            <DataMember> Private m_Number As String = String.Empty
            <DataMember> Private m_Tags As New List(Of String)

            Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
            Protected Sub OnPropertyChangedLocal(ByVal sName As String)
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
            End Sub
            Public Function Clone() As Object Implements ICloneable.Clone
                Return New TagItem(m_Number, m_Tags)
            End Function
            Sub New(ByVal iNumber As Integer, ByVal oTags As List(Of String))
                m_Number = iNumber
                m_Tags.AddRange(oTags)
            End Sub
            Public Property Number As String
                Get
                    Return m_Number
                End Get
                Set(value As String)
                    m_Number = value
                    OnPropertyChangedLocal("Number")
                End Set
            End Property
            Public Property Tags As List(Of String)
                Get
                    Return m_Tags
                End Get
                Set(value As List(Of String))
                    m_Tags.Clear()
                    m_Tags.AddRange(value)
                    OnPropertyChangedLocal("Tags")
                End Set
            End Property
        End Class
    End Class
    Private Function LocalKnownTypes() As List(Of Type)
        ' gets a list of known local class types
        Dim oKnownTypes As New List(Of Type)
        oKnownTypes.Add(GetType(FormMain))
        oKnownTypes.Add(GetType(FormProperties))
        oKnownTypes.Add(GetType(FormHeader))
        oKnownTypes.Add(GetType(FormPageHeader))
        oKnownTypes.Add(GetType(FormFormHeader))
        oKnownTypes.Add(GetType(FormBlock))
        oKnownTypes.Add(GetType(FormField))
        oKnownTypes.Add(GetType(FieldBorder))
        oKnownTypes.Add(GetType(FieldBackground))
        oKnownTypes.Add(GetType(FieldNumbering))
        oKnownTypes.Add(GetType(FieldText))
        oKnownTypes.Add(GetType(FieldImage))
        oKnownTypes.Add(GetType(FieldChoice))
        oKnownTypes.Add(GetType(FieldChoiceVertical))
        oKnownTypes.Add(GetType(FieldChoiceVerticalMCQ))
        oKnownTypes.Add(GetType(FieldBoxChoice))
        oKnownTypes.Add(GetType(FieldHandwriting))
        oKnownTypes.Add(GetType(FieldFree))
        oKnownTypes.Add(GetType(FormBody))
        oKnownTypes.Add(GetType(FormFooter))
        oKnownTypes.Add(GetType(FormPDF))
        oKnownTypes.Add(GetType(FormExport))
        oKnownTypes.Add(GetType(FormSection))
        oKnownTypes.Add(GetType(FormSubSection))
        oKnownTypes.Add(GetType(FormMCQ))
        oKnownTypes.Add(GetType(FormFormatter))
        oKnownTypes.Add(GetType(FormatterPageBreak))
        oKnownTypes.Add(GetType(FormatterDivider))
        oKnownTypes.Add(GetType(FormatterGroup))
        oKnownTypes.Add(GetType(BaseFormItem))
        oKnownTypes.Add(GetType(FormMain.DataObjectClass))
        oKnownTypes.Add(GetType(FormProperties.DataObjectClass))
        oKnownTypes.Add(GetType(FormHeader.DataObjectClass))
        oKnownTypes.Add(GetType(FormPageHeader.DataObjectClass))
        oKnownTypes.Add(GetType(FormFormHeader.DataObjectClass))
        oKnownTypes.Add(GetType(FormBlock.DataObjectClass))
        oKnownTypes.Add(GetType(FormField.DataObjectClass))
        oKnownTypes.Add(GetType(FieldBorder.DataObjectClass))
        oKnownTypes.Add(GetType(FieldBackground.DataObjectClass))
        oKnownTypes.Add(GetType(FieldNumbering.DataObjectClass))
        oKnownTypes.Add(GetType(FieldText.DataObjectClass))
        oKnownTypes.Add(GetType(FieldImage.DataObjectClass))
        oKnownTypes.Add(GetType(FieldChoice.DataObjectClass))
        oKnownTypes.Add(GetType(FieldChoiceVertical.DataObjectClass))
        oKnownTypes.Add(GetType(FieldChoiceVerticalMCQ.DataObjectClass))
        oKnownTypes.Add(GetType(FieldBoxChoice.DataObjectClass))
        oKnownTypes.Add(GetType(FieldHandwriting.DataObjectClass))
        oKnownTypes.Add(GetType(FieldFree.DataObjectClass))
        oKnownTypes.Add(GetType(FormBody.DataObjectClass))
        oKnownTypes.Add(GetType(FormFooter.DataObjectClass))
        oKnownTypes.Add(GetType(FormPDF.DataObjectClass))
        oKnownTypes.Add(GetType(FormExport.DataObjectClass))
        oKnownTypes.Add(GetType(FormSection.DataObjectClass))
        oKnownTypes.Add(GetType(FormSubSection.DataObjectClass))
        oKnownTypes.Add(GetType(FormMCQ.DataObjectClass))
        oKnownTypes.Add(GetType(FormFormatter.DataObjectClass))
        oKnownTypes.Add(GetType(FormatterPageBreak.DataObjectClass))
        oKnownTypes.Add(GetType(FormatterDivider.DataObjectClass))
        oKnownTypes.Add(GetType(FormatterGroup.DataObjectClass))
        oKnownTypes.Add(GetType(BaseFormItem.DataObjectClass))
        oKnownTypes.Add(GetType(FormMCQ.MCQClass))
        oKnownTypes.Add(GetType(FormMCQ.MCQClass.MCQItem))
        oKnownTypes.Add(GetType(List(Of BaseFormItem)))
        oKnownTypes.Add(GetType(System.Drawing.Bitmap))

        Return oKnownTypes
    End Function
#End Region
#End Region
#Region "Buttons"
    ' add processing for the plugin buttons here
    Private Sub ClearForm_Button_Click(sender As Object, e As RoutedEventArgs)
        If MessageBox.Show("Clear form contents?", ModuleName, MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
            Using oSuspender As New BaseFormItem.Suspender()
                Dim oOldFormMain As FormMain = oCommonVariables.DataStore(oMainFormGUID)
                oOldFormMain.Dispose()
                oOldFormMain = Nothing

                oCommonVariables.DataStore.TryRemove(oMainFormGUID, Nothing)
                Dim oFormMain As New FormMain()
                oCommonVariables.DataStore.TryAdd(oMainFormGUID, oFormMain)

                PostInit()

                BaseFormItem.FormMain.SetLevel()
                LeftArrangeFormItems()
                BaseFormItem.FormBody.Click()
            End Using

            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Form contents cleared."))
        End If
    End Sub
    Private Sub LoadForm_Button_Click(sender As Object, e As RoutedEventArgs)
        If MessageBox.Show("Load form contents from file?", ModuleName, MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
            Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
            oOpenFileDialog.FileName = String.Empty
            oOpenFileDialog.DefaultExt = "*.xml"
            oOpenFileDialog.Multiselect = False
            oOpenFileDialog.Filter = "XML Files|*.xml"
            oOpenFileDialog.Title = "Load Form Contents From File"
            Dim result? As Boolean = oOpenFileDialog.ShowDialog()
            If result = True Then
                Dim oXmlDictionaryReaderQuotas As New Xml.XmlDictionaryReaderQuotas()
                oXmlDictionaryReaderQuotas.MaxArrayLength = 100000000

                Dim oXmlDictionaryReader As Xml.XmlDictionaryReader = Xml.XmlDictionaryReader.CreateTextReader(IO.File.OpenRead(oOpenFileDialog.FileName), oXmlDictionaryReaderQuotas)

                Dim oFormMain As FormMain = Nothing
                Try
                    Dim oDataContractSerializer As New DataContractSerializer(GetType(FormMain), CommonFunctions.GetKnownTypes(LocalKnownTypes))
                    oFormMain = oDataContractSerializer.ReadObject(oXmlDictionaryReader, True)
                Catch ex As SerializationException
                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error deserialising file."))
                    Exit Sub
                End Try

                oXmlDictionaryReader.Close()

                If Not IsNothing(oFormMain) Then
                    Dim oOldFormMain As FormMain = oCommonVariables.DataStore(oMainFormGUID)
                    oOldFormMain.Dispose()
                    oOldFormMain = Nothing

                    oCommonVariables.DataStore.TryRemove(oMainFormGUID, Nothing)
                    oCommonVariables.DataStore.TryAdd(oMainFormGUID, oFormMain)

                    BaseFormItem.FormMain = CType(oCommonVariables.DataStore(oMainFormGUID), FormMain)
                    BaseFormItem.FormMain.SetLevel()
                    SetBindings()
                    LeftArrangeFormItems()
                    BaseFormItem.FormBody.Click()

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Form contents loaded from file."))
                End If
            End If
        End If
    End Sub
    Private Sub SaveForm_Button_Click(sender As Object, e As RoutedEventArgs)
        ' save form contents
        Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
        oSaveFileDialog.FileName = String.Empty
        oSaveFileDialog.DefaultExt = "*.xml"
        oSaveFileDialog.Filter = "XML Files|*.xml"
        oSaveFileDialog.Title = "Save Form Contents To File"
        Dim result? As Boolean = oSaveFileDialog.ShowDialog()
        If result = True Then
            ' serialise main form
            Dim oFormMain As FormMain = CType(oCommonVariables.DataStore(oMainFormGUID), FormMain)
            Using oFileStream As IO.FileStream = IO.File.Create(oSaveFileDialog.FileName)
                Dim oDataContractSerializer As New DataContractSerializer(GetType(FormMain), CommonFunctions.GetKnownTypes(LocalKnownTypes))
                oDataContractSerializer.WriteObject(oFileStream, oFormMain)
            End Using

            oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Form contents saved to file."))
        End If
    End Sub
    Private Sub LoadSection_Button_Click(sender As Object, e As RoutedEventArgs)
        Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
        oOpenFileDialog.FileName = String.Empty
        oOpenFileDialog.DefaultExt = "*.gz"
        oOpenFileDialog.Multiselect = False
        oOpenFileDialog.Filter = "GZip Files|*.gz"
        oOpenFileDialog.Title = "Load Section Contents From File"
        Dim result? As Boolean = oOpenFileDialog.ShowDialog()
        If result = True Then
            Dim oFormItemList As List(Of BaseFormItem) = CommonFunctions.DeserializeDataContractFile(Of List(Of BaseFormItem))(oOpenFileDialog.FileName, LocalKnownTypes, False)
            If IsNothing(oFormItemList) Then
                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error loading section contents."))
            Else
                BaseFormItem.SectionList.Add(oFormItemList)

                ' update field list
                If IsNothing(SelectedItem) Then
                    BaseFormItem.FormMain.RightSetAllowedFields()
                Else
                    SelectedItem.RightSetAllowedFields()
                    SelectedItem.Display()
                End If

                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Section contents loaded from file."))
            End If
        End If
    End Sub
    Private Sub Exit_Button_Click(sender As Object, e As RoutedEventArgs)
        'FormMain.ImageViewer.Dispose()
        RaiseEvent ExitButtonClick(Guid.Empty)
    End Sub
    Private Sub LoadLink_Button_Click(sender As Object, e As RoutedEventArgs)
        ' loads linked plugin
        Dim oButton As Controls.Button = sender
        Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
        Dim iButton As Integer = Val(Right(oButton.Name, Len(oButton.Name) - Len("Button")))

        FormSection.ImageViewer.Dispose()
        RaiseEvent ExitButtonClick(m_Identifiers(oFilteredNames(iButton)).Item1)
    End Sub
#End Region
#Region "ButtonHelp"
    Public Const DragHelp As String = "Help"
    Public Shared Sub HelpMouseMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles Me.MouseMove
        ' allows drag and drop
        If e.LeftButton = Input.MouseButtonState.Pressed Then
            Dim oInputElement As IInputElement = Input.Mouse.DirectlyOver
            If (Not IsNothing(oInputElement)) AndAlso oInputElement.GetType.IsSubclassOf(GetType(FrameworkElement)) Then
                Dim oFrameworkElement As FrameworkElement = oInputElement
                Dim oGUID As Guid = Guid.Empty
                IterateElements(oFrameworkElement, oGUID)

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
        Dim oPoint As Point = Input.Mouse.GetPosition(oFrameworkElement)
        If oPoint.X >= 0 AndAlso oPoint.Y >= 0 AndAlso oPoint.X < oFrameworkElement.ActualWidth AndAlso oPoint.Y < oFrameworkElement.ActualHeight Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub ButtonHelpMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles ButtonHelp.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            ButtonHelp.Background = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
        End If
    End Sub
    Private Sub ButtonHelpLeaveHandler(sender As Object, e As EventArgs) Handles ButtonHelp.MouseLeave, ButtonHelp.TouchLeave
        ButtonHelp.Background = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub ButtonHelp_DragEnter(sender As Object, e As DragEventArgs) Handles ButtonHelp.DragEnter
        If e.Data.GetDataPresent(DragHelp) Or e.Data.GetDataPresent(DragBaseFormItem) Then
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
        ElseIf e.Data.GetDataPresent(DragBaseFormItem) Then
            Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)
            Guid.TryParse(oDragData.Item4, oDragGUID)
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
    Private Sub PropertiesOrientation_SelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs) Handles PropertiesOrientation.SelectionChanged
            ' updates the page orientation
            Dim oAction As Action = Sub()
                                        Dim oComboBoxMain As Controls.ComboBox = sender
                                        If BaseFormItem.FormProperties.SelectedOrientation <> [Enum].GetValues(GetType(PageOrientation))(oComboBoxMain.SelectedIndex) Then
                                            BaseFormItem.FormProperties.SelectedOrientation = [Enum].GetValues(GetType(PageOrientation))(oComboBoxMain.SelectedIndex)
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub PropertiesSize_SelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs) Handles PropertiesSize.SelectionChanged
        ' updates the page orientation
        Dim oAction As Action = Sub()
                                    Dim oComboBoxMain As Controls.ComboBox = sender
                                    If BaseFormItem.FormProperties.SelectedSize <> PDFHelper.PageDictionary(BaseFormItem.FormProperties.SelectedOrientation)(oComboBoxMain.SelectedIndex) Then
                                        BaseFormItem.FormProperties.SelectedSize = PDFHelper.PageDictionary(BaseFormItem.FormProperties.SelectedOrientation)(oComboBoxMain.SelectedIndex)
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PageHeaderAddHandler(sender As Object, e As EventArgs) Handles PageHeaderAdd.Click
        Dim oAction As Action = Sub()
                                    BaseFormItem.FormPageHeader.AddLine()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PageHeaderRemoveHandler(sender As Object, e As EventArgs) Handles PageHeaderRemove.Click
            Dim oAction As Action = Sub()
                                        BaseFormItem.FormPageHeader.RemoveLine()
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub PageHeaderPreviousHandler(sender As Object, e As EventArgs) Handles PageHeaderPrevious.Click
        Dim oAction As Action = Sub()
                                    BaseFormItem.FormPageHeader.PreviousLine()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PageHeaderNextHandler(sender As Object, e As EventArgs) Handles PageHeaderNext.Click
        Dim oAction As Action = Sub()
                                    BaseFormItem.FormPageHeader.NextLine()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PageHeaderFontSizeIncreaseHandler(sender As Object, e As EventArgs) Handles PageHeaderFontSizeIncrease.Click
        Dim oAction As Action = Sub()
                                    If Not IsNothing(PageHeaderFontSizeMultiplier.DataContext) Then
                                        BaseFormItem.FormPageHeader.FontSizeMultiplier += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub PageHeaderFontSizeDecreaseHandler(sender As Object, e As EventArgs) Handles PageHeaderFontSizeDecrease.Click
        Dim oAction As Action = Sub()
                                    If Not IsNothing(PageHeaderFontSizeMultiplier.DataContext) Then
                                        BaseFormItem.FormPageHeader.FontSizeMultiplier -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockPlaceItemHandler(sender As Object, e As EventArgs) Handles BlockPlaceItem.Click
        Dim oAction As Action = Sub()
                                    BlockPlaceItem.HBSelected = Not BlockPlaceItem.HBSelected
                                    FormBlock.BlockPlaceItemClicked()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockAddRowHandler(sender As Object, e As EventArgs) Handles BlockAddRow.Click
        Dim oAction As Action = Sub()
                                    Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                    Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oFormBlock) Then
                                        oFormBlock.GridHeight += 1
                                    ElseIf Not IsNothing(oFormBlockParent) Then
                                        oFormBlockParent.GridHeight += 1
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.GridHeight += 1
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.GridHeight += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockRemoveRowHandler(sender As Object, e As EventArgs) Handles BlockRemoveRow.Click
            Dim oAction As Action = Sub()
                                        Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                        Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                        Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                        Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                        If Not IsNothing(oFormBlock) Then
                                            oFormBlock.GridHeight -= 1
                                        ElseIf Not IsNothing(oFormBlockParent) Then
                                            oFormBlockParent.GridHeight -= 1
                                        ElseIf Not IsNothing(oSection) Then
                                            oSection.GridHeight -= 1
                                        ElseIf Not IsNothing(oSectionParent) Then
                                            oSectionParent.GridHeight -= 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub BlockAddColumnHandler(sender As Object, e As EventArgs) Handles BlockAddColumn.Click
        Dim oAction As Action = Sub()
                                    Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                    Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                    If Not IsNothing(oFormBlock) Then
                                        oFormBlock.GridWidth += 1
                                    ElseIf Not IsNothing(oFormBlockParent) Then
                                        oFormBlockParent.GridWidth += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockRemoveColumnHandler(sender As Object, e As EventArgs) Handles BlockRemoveColumn.Click
        Dim oAction As Action = Sub()
                                    Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                    Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                    If Not IsNothing(oFormBlock) Then
                                        oFormBlock.GridWidth -= 1
                                    ElseIf Not IsNothing(oFormBlockParent) Then
                                        oFormBlockParent.GridWidth -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockIncreaseWidthHandler(sender As Object, e As EventArgs) Handles BlockIncreaseWidth.Click
        Dim oAction As Action = Sub()
                                    Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                    Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                    Dim oFormMCQ As FormMCQ = TryCast(SelectedItem, FormMCQ)
                                    Dim oFormSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    If Not IsNothing(oFormBlock) Then
                                        oFormBlock.BlockWidth += 1
                                    ElseIf Not IsNothing(oFormBlockParent) Then
                                        oFormBlockParent.BlockWidth += 1
                                    ElseIf Not IsNothing(oFormMCQ) Then
                                        oFormMCQ.BlockWidth += 1
                                    ElseIf Not IsNothing(oFormSubSection) Then
                                        oFormSubSection.BlockWidth += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockDecreaseWidthHandler(sender As Object, e As EventArgs) Handles BlockDecreaseWidth.Click
            Dim oAction As Action = Sub()
                                        Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                        Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                        Dim oFormMCQ As FormMCQ = TryCast(SelectedItem, FormMCQ)
                                        Dim oFormSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                        If Not IsNothing(oFormBlock) Then
                                            oFormBlock.BlockWidth -= 1
                                        ElseIf Not IsNothing(oFormBlockParent) Then
                                            oFormBlockParent.BlockWidth -= 1
                                        ElseIf Not IsNothing(oFormMCQ) Then
                                            oFormMCQ.BlockWidth -= 1
                                        ElseIf Not IsNothing(oFormSubSection) Then
                                            oFormSubSection.BlockWidth -= 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub BlockShowGridHandler(sender As Object, e As EventArgs) Handles BlockShowGrid.Click
        Dim oAction As Action = Sub()
                                    Dim oFormBlock As FormBlock = TryCast(SelectedItem, FormBlock)
                                    Dim oFormBlockParent As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                    If Not IsNothing(oFormBlock) Then
                                        oFormBlock.ShowGrid = Not oFormBlock.ShowGrid
                                    ElseIf Not IsNothing(oFormBlockParent) Then
                                        oFormBlockParent.ShowGrid = Not oFormBlockParent.ShowGrid
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldSingleHandler(sender As Object, e As EventArgs) Handles FieldSingle.Click
        Dim oAction As Action = Sub()
                                    Dim oFieldChoice As FieldChoice = TryCast(SelectedItem, FieldChoice)
                                    If Not IsNothing(oFieldChoice) Then
                                        oFieldChoice.TabletSingleChoiceOnly = Not oFieldChoice.TabletSingleChoiceOnly
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldCriticalHandler(sender As Object, e As EventArgs) Handles FieldCritical.Click
            Dim oAction As Action = Sub()
                                        Dim oFormField As FormField = TryCast(SelectedItem, FormField)
                                        If Not IsNothing(oFormField) Then
                                            oFormField.Critical = Not oFormField.Critical
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldExcludeHandler(sender As Object, e As EventArgs) Handles FieldExclude.Click
        Dim oAction As Action = Sub()
                                    Dim oFormField As FormField = TryCast(SelectedItem, FormField)
                                    If Not IsNothing(oFormField) Then
                                        oFormField.Exclude = Not oFormField.Exclude
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub StaticTextShowActiveHandler(sender As Object, e As EventArgs) Handles StaticTextShowActive.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldText)) AndAlso (Not IsNothing(SelectedItem.Parent)) Then
                                        Dim oFormBlock As FormBlock = TryCast(SelectedItem.Parent, FormBlock)
                                        If Not IsNothing(oFormBlock) Then
                                            oFormBlock.ShowActive = Not oFormBlock.ShowActive
                                            If oFormBlock.ShowActive Then
                                                StaticTextShowActive.HBSource = m_Icons("CCMVisible")
                                                StaticTextShowActive.InnerToolTip = "Show Active Selection"
                                            Else
                                                StaticTextShowActive.HBSource = m_Icons("CCMNotVisible")
                                                StaticTextShowActive.InnerToolTip = "Hide Active Selection"
                                            End If
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub StaticTextAddSubjectHandler(sender As Object, e As EventArgs) Handles StaticTextAddSubject.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldText)) Then
                                            Dim oText As FieldText = SelectedItem
                                            oText.AddSubject()
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub StaticTextAddTagHandler(sender As Object, e As EventArgs) Handles StaticTextAddTag.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldText)) Then
                                        Dim oText As FieldText = SelectedItem
                                        oText.AddTag()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub StaticTextAddHandler(sender As Object, e As EventArgs) Handles StaticTextAdd.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldText)) Then
                                        Dim oText As FieldText = SelectedItem
                                        oText.AddElement(String.Empty, ElementStruc.ElementTypeEnum.Text, oText.Font)
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub StaticTextRemoveHandler(sender As Object, e As EventArgs) Handles StaticTextRemove.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldText)) Then
                                            Dim oText As FieldText = SelectedItem
                                            oText.RemoveElement()
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub StaticTextPreviousHandler(sender As Object, e As EventArgs) Handles StaticTextPrevious.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldText)) Then
                                        Dim oText As FieldText = SelectedItem
                                        oText.PreviousElement()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub StaticTextNextHandler(sender As Object, e As EventArgs) Handles StaticTextNext.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldText)) Then
                                        Dim oText As FieldText = SelectedItem
                                        oText.NextElement()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldWholeBlockHandler(sender As Object, e As EventArgs) Handles FieldWholeBlock.Click
            Dim oAction As Action = Sub()
                                        Dim oFormField As FormField = TryCast(SelectedItem, FormField)
                                        If Not IsNothing(oFormField) Then
                                            oFormField.WholeBlock = Not oFormField.WholeBlock
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldBorderIncreaseWidthHandler(sender As Object, e As EventArgs) Handles FieldBorderIncreaseWidth.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldBorder)) Then
                                        Dim oBorder As FieldBorder = SelectedItem
                                        oBorder.BorderWidth += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldBorderDecreaseWidthHandler(sender As Object, e As EventArgs) Handles FieldBorderDecreaseWidth.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldBorder)) Then
                                        Dim oBorder As FieldBorder = SelectedItem
                                        oBorder.BorderWidth -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldBackgroundDarkenHandler(sender As Object, e As EventArgs) Handles FieldBackgroundDarken.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldBackground)) Then
                                            Dim oFieldBackground As FieldBackground = SelectedItem
                                            oFieldBackground.Lightness += 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldBackgroundLightenHandler(sender As Object, e As EventArgs) Handles FieldBackgroundLighten.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldBackground)) Then
                                        Dim oFieldBackground As FieldBackground = SelectedItem
                                        oFieldBackground.Lightness -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoicePreviousTabletContentHandler(sender As Object, e As EventArgs) Handles FieldChoicePreviousTabletContent.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        Dim iLength As Integer = [Enum].GetValues(GetType(Enumerations.TabletContentEnum)).Length
                                        oFieldChoice.TabletContent = CType((oFieldChoice.TabletContent + iLength - 1) Mod iLength, Enumerations.TabletContentEnum)
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoiceNextTabletContentHandler(sender As Object, e As EventArgs) Handles FieldChoiceNextTabletContent.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                            Dim oFieldChoice As FieldChoice = SelectedItem
                                            Dim iLength As Integer = [Enum].GetValues(GetType(Enumerations.TabletContentEnum)).Length
                                            oFieldChoice.TabletContent = CType((oFieldChoice.TabletContent + 1) Mod iLength, Enumerations.TabletContentEnum)
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldChoiceAddTabletHandler(sender As Object, e As EventArgs) Handles FieldChoiceAddTablet.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.TabletCount += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoiceRemoveTabletHandler(sender As Object, e As EventArgs) Handles FieldChoiceRemoveTablet.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.TabletCount -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoiceIncreaseStartHandler(sender As Object, e As EventArgs) Handles FieldChoiceIncreaseStart.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                            Dim oFieldChoice As FieldChoice = SelectedItem
                                            oFieldChoice.TabletStart += 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldChoiceDecreaseStartHandler(sender As Object, e As EventArgs) Handles FieldChoiceDecreaseStart.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.TabletStart -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoiceIncreaseGroupsHandler(sender As Object, e As EventArgs) Handles FieldChoiceIncreaseGroups.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.TabletGroups += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoiceDecreaseGroupsHandler(sender As Object, e As EventArgs) Handles FieldChoiceDecreaseGroups.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                            Dim oFieldChoice As FieldChoice = SelectedItem
                                            oFieldChoice.TabletGroups -= 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldChoicePreviousTopDescriptionHandler(sender As Object, e As EventArgs) Handles FieldChoicePreviousTopDescription.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.CurrentDescriptionTop -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoiceNextTopDescriptionHandler(sender As Object, e As EventArgs) Handles FieldChoiceNextTopDescription.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.CurrentDescriptionTop += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldChoicePreviousBottomDescriptionHandler(sender As Object, e As EventArgs) Handles FieldChoicePreviousBottomDescription.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                            Dim oFieldChoice As FieldChoice = SelectedItem
                                            oFieldChoice.CurrentDescriptionBottom -= 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldChoiceNextBottomDescriptionHandler(sender As Object, e As EventArgs) Handles FieldChoiceNextBottomDescription.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldChoice)) OrElse SelectedItem.GetType.IsSubclassOf(GetType(FieldChoice)) Then
                                        Dim oFieldChoice As FieldChoice = SelectedItem
                                        oFieldChoice.CurrentDescriptionBottom += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldBoxChoicePreviousLabelDescriptionHandler(sender As Object, e As EventArgs) Handles FieldBoxChoicePreviousLabel.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldBoxChoice)) Then
                                        Dim oBoxChoice As FieldBoxChoice = SelectedItem
                                        oBoxChoice.CurrentDescriptionTop -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldBoxChoiceNextLabelHandler(sender As Object, e As EventArgs) Handles FieldBoxChoiceNextLabel.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldBoxChoice)) Then
                                        Dim oBoxChoice As FieldBoxChoice = SelectedItem
                                        oBoxChoice.CurrentDescriptionTop += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldHandwritingAddBlockHandler(sender As Object, e As EventArgs) Handles FieldHandwritingAddBlock.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                        Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                        oFieldHandwriting.BlockCount += 1
                                    ElseIf SelectedItem.GetType.Equals(GetType(FieldBoxChoice)) Then
                                        Dim oFieldBoxChoice As FieldBoxChoice = SelectedItem
                                        oFieldBoxChoice.BlockCount += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldHandwritingRemoveBlockHandler(sender As Object, e As EventArgs) Handles FieldHandwritingRemoveBlock.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                            Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                            oFieldHandwriting.BlockCount -= 1
                                        ElseIf SelectedItem.GetType.Equals(GetType(FieldBoxChoice)) Then
                                            Dim oFieldBoxChoice As FieldBoxChoice = SelectedItem
                                            oFieldBoxChoice.BlockCount -= 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldHandwritingPreviousColumnHandler(sender As Object, e As EventArgs) Handles FieldHandwritingPreviousColumn.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                        Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                        oFieldHandwriting.CurrentColumn -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldHandwritingNextColumnHandler(sender As Object, e As EventArgs) Handles FieldHandwritingNextColumn.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                        Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                        oFieldHandwriting.CurrentColumn += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldHandwritingPreviousRowHandler(sender As Object, e As EventArgs) Handles FieldHandwritingPreviousRow.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                            Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                            oFieldHandwriting.CurrentRow -= 1
                                        ElseIf SelectedItem.GetType.Equals(GetType(FieldBoxChoice)) Then
                                            Dim oFieldBoxChoice As FieldBoxChoice = SelectedItem
                                            oFieldBoxChoice.CurrentRow -= 1
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldHandwritingNextRowHandler(sender As Object, e As EventArgs) Handles FieldHandwritingNextRow.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                        Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                        oFieldHandwriting.CurrentRow += 1
                                    ElseIf SelectedItem.GetType.Equals(GetType(FieldBoxChoice)) Then
                                        Dim oFieldBoxChoice As FieldBoxChoice = SelectedItem
                                        oFieldBoxChoice.CurrentRow += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldHandwritingPreviousTypeHandler(sender As Object, e As EventArgs) Handles FieldHandwritingPreviousType.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                        Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                        Select Case oFieldHandwriting.CharacterASCII
                                            Case Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers
                                                oFieldHandwriting.CharacterASCII = Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric
                                            Case Enumerations.CharacterASCII.Numbers
                                                oFieldHandwriting.CharacterASCII = Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers
                                            Case Enumerations.CharacterASCII.None, Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric
                                                oFieldHandwriting.CharacterASCII = Enumerations.CharacterASCII.Numbers
                                            Case Else
                                                ' no change
                                        End Select
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FieldHandwritingNextTypeHandler(sender As Object, e As EventArgs) Handles FieldHandwritingNextType.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                            Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                            Select Case oFieldHandwriting.CharacterASCII
                                                Case Enumerations.CharacterASCII.None, Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric
                                                    oFieldHandwriting.CharacterASCII = Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers
                                                Case Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers
                                                    oFieldHandwriting.CharacterASCII = Enumerations.CharacterASCII.Numbers
                                                Case Enumerations.CharacterASCII.Numbers
                                                    oFieldHandwriting.CharacterASCII = Enumerations.CharacterASCII.Uppercase Or Enumerations.CharacterASCII.Numbers Or Enumerations.CharacterASCII.NonAlphaNumeric
                                                Case Else
                                                    ' no change
                                            End Select
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FieldHandwritingCurrentContentHandler(sender As Object, e As Controls.TextChangedEventArgs) Handles FieldHandwritingCurrentContent.TextChanged
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldHandwriting)) Then
                                        Dim oFieldHandwriting As FieldHandwriting = SelectedItem
                                        oFieldHandwriting.BlockCount = oFieldHandwriting.BlockCount
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockSpacerHandler(sender As Object, e As EventArgs) Handles BlockSpacer.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldNumbering)) Then
                                        Dim oFieldNumbering As FieldNumbering = SelectedItem
                                        oFieldNumbering.Spacer = Not oFieldNumbering.Spacer
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockNumberingBorderHandler(sender As Object, e As EventArgs) Handles BlockNumberingBorder.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.NumberingBorder = Not oSection.NumberingBorder
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.NumberingBorder = Not oSectionParent.NumberingBorder
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub BlockNumberingBackgroundHandler(sender As Object, e As EventArgs) Handles BlockNumberingBackground.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.NumberingBackground = Not oSection.NumberingBackground
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.NumberingBackground = Not oSectionParent.NumberingBackground
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockContinuousNumberingHandler(sender As Object, e As EventArgs) Handles BlockContinuousNumbering.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.ContinuousNumbering = Not oSection.ContinuousNumbering
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.ContinuousNumbering = Not oSectionParent.ContinuousNumbering
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockIncreaseStartHandler(sender As Object, e As EventArgs) Handles BlockIncreaseStart.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.Start += 1
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Start += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockDecreaseStartHandler(sender As Object, e As EventArgs) Handles BlockDecreaseStart.Click
            Dim oAction As Action = Sub()
                                        Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                        Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                        If Not IsNothing(oSection) Then
                                            oSection.Start -= 1
                                        ElseIf Not IsNothing(oSectionParent) Then
                                            oSectionParent.Start -= 1
                                        End If
                                    End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub BlockPreviousTypeHandler(sender As Object, e As EventArgs) Handles BlockPreviousType.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    If IsNothing(oSection) Then
                                        oSection = TryCast(SelectedItem.Parent, FormSection)
                                    End If
                                    If Not IsNothing(oSection) Then
                                        Select Case oSection.NumberingType
                                            Case Enumerations.Numbering.Number
                                                oSection.NumberingType = Enumerations.Numbering.LetterBig
                                            Case Enumerations.Numbering.LetterSmall
                                                oSection.NumberingType = Enumerations.Numbering.Number
                                            Case Enumerations.Numbering.LetterBig
                                                oSection.NumberingType = Enumerations.Numbering.LetterSmall
                                            Case Else
                                        End Select
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub BlockNextTypeHandler(sender As Object, e As EventArgs) Handles BlockNextType.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    If IsNothing(oSection) Then
                                        oSection = TryCast(SelectedItem.Parent, FormSection)
                                    End If
                                    If Not IsNothing(oSection) Then
                                        Select Case oSection.NumberingType
                                            Case Enumerations.Numbering.Number
                                                oSection.NumberingType = Enumerations.Numbering.LetterSmall
                                            Case Enumerations.Numbering.LetterSmall
                                                oSection.NumberingType = Enumerations.Numbering.LetterBig
                                            Case Enumerations.Numbering.LetterBig
                                                oSection.NumberingType = Enumerations.Numbering.Number
                                            Case Else
                                        End Select
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SectionTagGUIDHandler(sender As Object, e As EventArgs) Handles SectionTagGUID.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.GenerateGUIDKey()
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.GenerateGUIDKey()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SectionTagGUIDClearHandler(sender As Object, e As EventArgs) Handles SectionTagGUID.RightClick
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.Tag = String.Empty
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Tag = String.Empty
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SectionTagCopyHandler(sender As Object, e As EventArgs) Handles SectionTagCopy.Click
        Dim oAction As Action = Sub()
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oSection) Then
                                        oSection.CopyKeyToClipboard()
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.CopyKeyToClipboard()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub MCQLoadDataHandler(sender As Object, e As EventArgs) Handles MCQLoadData.Click
        Dim oAction As Action = Sub()
                                    Dim oMCQ As FormMCQ = TryCast(SelectedItem, FormMCQ)
                                    Dim oMCQParent As FormMCQ = TryCast(SelectedItem.Parent, FormMCQ)
                                    If Not IsNothing(oMCQ) Then
                                        oMCQ.ImportQuestions()
                                    ElseIf Not IsNothing(oMCQParent) Then
                                        oMCQParent.ImportQuestions()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsPreviousHandler(sender As Object, e As EventArgs) Handles TagsPrevious.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    If Not IsNothing(oSubSection) Then
                                        If oSubSection.TagCount > 0 AndAlso oSubSection.SelectedTag >= 0 Then
                                            oSubSection.SelectedTag -= 1
                                        End If
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        If oSubSectionParent.TagCount > 0 AndAlso oSubSectionParent.SelectedTag >= 0 Then
                                            oSubSectionParent.SelectedTag -= 1
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsNextHandler(sender As Object, e As EventArgs) Handles TagsNext.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    If Not IsNothing(oSubSection) Then
                                        If oSubSection.TagCount > 0 AndAlso oSubSection.SelectedTag < oSubSection.TagCount - 1 Then
                                            oSubSection.SelectedTag += 1
                                        End If
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        If oSubSectionParent.TagCount > 0 AndAlso oSubSectionParent.SelectedTag < oSubSectionParent.TagCount - 1 Then
                                            oSubSectionParent.SelectedTag += 1
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsLoadDataHandler(sender As Object, e As EventArgs) Handles TagsLoadData.Click, SubjectsLoadData.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oSubSection) Then
                                        oSubSection.ImportTags()
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        oSubSectionParent.ImportTags()
                                    ElseIf Not IsNothing(oExport) Then
                                        oExport.ImportSubjects()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsSaveDataHandler(sender As Object, e As EventArgs) Handles TagsLoadData.RightClick, SubjectsLoadData.RightClick
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oSubSection) Then
                                        oSubSection.ExportTags()
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        oSubSectionParent.ExportTags()
                                    ElseIf Not IsNothing(oExport) Then
                                        oExport.ExportSubjects()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsAddRowHandler(sender As Object, e As EventArgs) Handles TagsAddRow.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oSubSection) Then
                                        oSubSection.AddRow()
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        oSubSectionParent.AddRow()
                                    ElseIf Not IsNothing(oExport) Then
                                        oExport.AddRow()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsRemoveRowHandler(sender As Object, e As EventArgs) Handles TagsRemoveRow.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oSubSection) Then
                                        oSubSection.RemoveRow()
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        oSubSectionParent.RemoveRow()
                                    ElseIf Not IsNothing(oExport) Then
                                        oExport.RemoveRow()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsAddColumnHandler(sender As Object, e As EventArgs) Handles TagsAddColumn.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    If Not IsNothing(oSubSection) Then
                                        oSubSection.AddColumn()
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        oSubSectionParent.AddColumn()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub TagsRemoveColumnHandler(sender As Object, e As EventArgs) Handles TagsRemoveColumn.Click
        Dim oAction As Action = Sub()
                                    Dim oSubSection As FormSubSection = TryCast(SelectedItem, FormSubSection)
                                    Dim oSubSectionParent As FormSubSection = TryCast(SelectedItem.Parent, FormSubSection)
                                    If Not IsNothing(oSubSection) Then
                                        oSubSection.RemoveColumn()
                                    ElseIf Not IsNothing(oSubSectionParent) Then
                                        oSubSectionParent.RemoveColumn()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SubjectsSavePDFHandler(sender As Object, e As EventArgs) Handles SubjectsSavePDF.Click
        Dim oAction As Action = Sub()
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oExport) Then
                                        oExport.SavePDF()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SubjectsReplacePDFHandler(sender As Object, e As EventArgs) Handles SubjectsReplacePDF.Click
        Dim oAction As Action = Sub()
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oExport) Then
                                        oExport.ReplacePDF()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SubjectsExportHelpHandler(sender As Object, e As EventArgs) Handles SubjectsExportHelp.Click
        Dim oAction As Action = Sub()
                                    Dim oExport As FormExport = TryCast(SelectedItem, FormExport)
                                    If Not IsNothing(oExport) Then
                                        oExport.ExportHelp()
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub JustificationLeftHandler(sender As Object, e As EventArgs) Handles JustificationLeft.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Justification = Enumerations.JustificationEnum.Left
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Justification = Enumerations.JustificationEnum.Left
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Justification = Enumerations.JustificationEnum.Left
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub JustificationCenterHandler(sender As Object, e As EventArgs) Handles JustificationCenter.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Justification = Enumerations.JustificationEnum.Center
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Justification = Enumerations.JustificationEnum.Center
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Justification = Enumerations.JustificationEnum.Center
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub JustificationRightHandler(sender As Object, e As EventArgs) Handles JustificationRight.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Justification = Enumerations.JustificationEnum.Right
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Justification = Enumerations.JustificationEnum.Right
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Justification = Enumerations.JustificationEnum.Right
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub JustificationJustifyHandler(sender As Object, e As EventArgs) Handles JustificationJustify.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Justification = Enumerations.JustificationEnum.Justify
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Justification = Enumerations.JustificationEnum.Justify
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Justification = Enumerations.JustificationEnum.Justify
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub FontBoldHandler(sender As Object, e As EventArgs) Handles FontBold.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Font = Enumerations.FontEnum.Bold
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Font = Enumerations.FontEnum.Bold
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Font = Enumerations.FontEnum.Bold
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FontItalicHandler(sender As Object, e As EventArgs) Handles FontItalic.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Font = Enumerations.FontEnum.Italic
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Font = Enumerations.FontEnum.Italic
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Font = Enumerations.FontEnum.Italic
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FontUnderlineHandler(sender As Object, e As EventArgs) Handles FontUnderline.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.Font = Enumerations.FontEnum.Underline
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.Font = Enumerations.FontEnum.Underline
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.Font = Enumerations.FontEnum.Underline
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FontSizeIncreaseHandler(sender As Object, e As EventArgs) Handles FontSizeIncrease.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.FontSizeMultiplier += 1
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.FontSizeMultiplier += 1
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.FontSizeMultiplier += 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FontSizeDecreaseHandler(sender As Object, e As EventArgs) Handles FontSizeDecrease.Click
        Dim oAction As Action = Sub()
                                    Dim oText As FieldText = TryCast(SelectedItem, FieldText)
                                    Dim oSection As FormSection = TryCast(SelectedItem, FormSection)
                                    Dim oSectionParent As FormSection = TryCast(SelectedItem.Parent, FormSection)
                                    If Not IsNothing(oText) Then
                                        oText.FontSizeMultiplier -= 1
                                    ElseIf Not IsNothing(oSection) Then
                                        oSection.FontSizeMultiplier -= 1
                                    ElseIf Not IsNothing(oSectionParent) Then
                                        oSectionParent.FontSizeMultiplier -= 1
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentLHandler(sender As Object, e As EventArgs) Handles AlignmentL.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                        Dim oFormField As FormField = SelectedItem
                                        oFormField.Alignment = Enumerations.AlignmentEnum.Left
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentULHandler(sender As Object, e As EventArgs) Handles AlignmentUL.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                            Dim oFormField As FormField = SelectedItem
                                            oFormField.Alignment = Enumerations.AlignmentEnum.UpperLeft
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub AlignmentTHandler(sender As Object, e As EventArgs) Handles AlignmentT.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                        Dim oFormField As FormField = SelectedItem
                                        oFormField.Alignment = Enumerations.AlignmentEnum.Top
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentURHandler(sender As Object, e As EventArgs) Handles AlignmentUR.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                        Dim oFormField As FormField = SelectedItem
                                        oFormField.Alignment = Enumerations.AlignmentEnum.UpperRight
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentRHandler(sender As Object, e As EventArgs) Handles AlignmentR.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                            Dim oFormField As FormField = SelectedItem
                                            oFormField.Alignment = Enumerations.AlignmentEnum.Right
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub AlignmentLRHandler(sender As Object, e As EventArgs) Handles AlignmentLR.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                        Dim oFormField As FormField = SelectedItem
                                        oFormField.Alignment = Enumerations.AlignmentEnum.LowerRight
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentBHandler(sender As Object, e As EventArgs) Handles AlignmentB.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                        Dim oFormField As FormField = SelectedItem
                                        oFormField.Alignment = Enumerations.AlignmentEnum.Bottom
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentLLHandler(sender As Object, e As EventArgs) Handles AlignmentLL.Click
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                        Dim oFormField As FormField = SelectedItem
                                        oFormField.Alignment = Enumerations.AlignmentEnum.LowerLeft
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub AlignmentCHandler(sender As Object, e As EventArgs) Handles AlignmentC.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.IsSubclassOf(GetType(FormField)) Then
                                            Dim oFormField As FormField = SelectedItem
                                            oFormField.Alignment = Enumerations.AlignmentEnum.Center
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub StretchFillHandler(sender As Object, e As EventArgs) Handles StretchFill.Click
        Dim oAction As Action = Sub()
                                    Select Case SelectedItem.GetType
                                        Case GetType(FieldImage)
                                            Dim oImage As FieldImage = SelectedItem
                                            oImage.Stretch = Enumerations.StretchEnum.Fill
                                    End Select
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub StretchUniformHandler(sender As Object, e As EventArgs) Handles StretchUniform.Click
        Dim oAction As Action = Sub()
                                    Select Case SelectedItem.GetType
                                        Case GetType(FieldImage)
                                            Dim oImage As FieldImage = SelectedItem
                                            oImage.Stretch = Enumerations.StretchEnum.Uniform
                                    End Select
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ImageLoadHandler(sender As Object, e As EventArgs) Handles ImageLoad.Click
            Dim oAction As Action = Sub()
                                        If SelectedItem.GetType.Equals(GetType(FieldImage)) Then
                                            Dim oFieldImage As FieldImage = SelectedItem

                                            Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
                                            oOpenFileDialog.FileName = String.Empty
                                            oOpenFileDialog.DefaultExt = "*"
                                            oOpenFileDialog.Multiselect = False
                                            oOpenFileDialog.Filter = "All Files|*.*|JPEG Images|*.jpg;*.jpeg|TIFF Images|*.tif;*.tiff|PNG Images|*.png"
                                            oOpenFileDialog.Title = "Loading Image"
                                            Dim result? As Boolean = oOpenFileDialog.ShowDialog()
                                            If result = True Then
                                                Try
                                                    Dim oImage As System.Drawing.Bitmap = CommonFunctions.LoadBitmap(oOpenFileDialog.FileName)
                                                    Dim oFileInfo As New IO.FileInfo(oOpenFileDialog.FileName)
                                                    Dim sFileName As String = Left(oFileInfo.Name, Len(oFileInfo.Name) - Len(oFileInfo.Extension))
                                                    oImage.SetResolution(RenderResolution300, RenderResolution300)
                                                    oFieldImage.SetImage(oImage, sFileName)
                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Image """ + sFileName + """ added."))
                                                Catch ex As Exception
                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error adding image."))
                                                End Try
                                            End If
                                        End If
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub ImageSaveHandler(sender As Object, e As EventArgs) Handles ImageLoad.RightClick
        Dim oAction As Action = Sub()
                                    If SelectedItem.GetType.Equals(GetType(FieldImage)) Then
                                        Dim oFieldImage As FieldImage = SelectedItem

                                        Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
                                        oSaveFileDialog.FileName = String.Empty
                                        oSaveFileDialog.DefaultExt = "*.tif"
                                        oSaveFileDialog.Filter = "TIFF Images|*.tif"
                                        oSaveFileDialog.Title = "Saving Image"
                                        oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
                                        Dim result? As Boolean = oSaveFileDialog.ShowDialog()
                                        If result = True Then
                                            Try
                                                Dim oImage As System.Drawing.Bitmap = oFieldImage.GetImage
                                                If IsNothing(oImage) Then
                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "No image."))
                                                Else
                                                    Dim sFileName As String = CommonFunctions.ReplaceExtension(oSaveFileDialog.FileName, "tif")
                                                    CommonFunctions.SaveBitmap(sFileName, oImage, True)

                                                    Dim oFileInfo As New IO.FileInfo(sFileName)
                                                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Image """ + oFileInfo.Name + """ saved."))
                                                End If
                                            Catch ex As Exception
                                                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error saving image."))
                                            End Try
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FooterShowTitleHandler(sender As Object, e As EventArgs) Handles FooterShowTitle.Click
        Dim oAction As Action = Sub()
                                    BaseFormItem.FormFooter.ShowTitle = Not BaseFormItem.FormFooter.ShowTitle
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FooterShowPageHandler(sender As Object, e As EventArgs) Handles FooterShowPage.Click
        Dim oAction As Action = Sub()
                                    BaseFormItem.FormFooter.ShowPage = Not BaseFormItem.FormFooter.ShowPage
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub FooterShowSubjectHandler(sender As Object, e As EventArgs) Handles FooterShowSubject.Click
            Dim oAction As Action = Sub()
                                        BaseFormItem.FormFooter.ShowSubject = Not BaseFormItem.FormFooter.ShowSubject
                                    End Sub
            CommonFunctions.RepeatCheck(sender, e, oAction)
        End Sub
    Private Sub RectangleRecycleBinMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles RectangleRecycleBin.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            RectangleRecycleBin.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            RectangleRecycleBinBackground.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub RectangleRecycleBinLeaveHandler(sender As Object, e As EventArgs) Handles RectangleRecycleBin.MouseLeave, RectangleRecycleBin.TouchLeave
        RectangleRecycleBin.Fill = New SolidColorBrush(Colors.Transparent)
        RectangleRecycleBinBackground.Visibility = Visibility.Hidden
    End Sub
    Private Sub RectangleRecycleBin_DragEnter(sender As Object, e As DragEventArgs) Handles RectangleRecycleBin.DragEnter
        If e.Data.GetDataPresent(DragBaseFormItem) Then
            Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)
            If oDragData.Item2.Equals(Guid.Empty) Then
                RectangleRecycleBin.Fill = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
            Else
                RectangleRecycleBin.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            End If
        Else
            RectangleRecycleBin.Fill = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
        End If
        RectangleRecycleBinBackground.Visibility = Visibility.Visible
    End Sub
    Private Sub RectangleRecycleBin_DragLeave(sender As Object, e As DragEventArgs) Handles RectangleRecycleBin.DragLeave
        RectangleRecycleBin.Fill = New SolidColorBrush(Colors.Transparent)
        RectangleRecycleBinBackground.Visibility = Visibility.Hidden
    End Sub
    Private Sub RectangleRecycleBin_Drop(sender As Object, e As DragEventArgs) Handles RectangleRecycleBin.Drop
        ' handles drop event
        If e.Data.GetDataPresent(DragBaseFormItem) Then
            Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)

            If Not oDragData.Item2.Equals(Guid.Empty) Then
                Using oSuspender As New BaseFormItem.Suspender()
                    BaseFormItem.FormPDF.Changed = True
                    Dim oFormItem As BaseFormItem = BaseFormItem.FormMain.FindChild(oDragData.Item2)
                    Dim oParent As BaseFormItem = oFormItem.Parent
                    oFormItem.Dispose()
                    oParent.Click()
                    LeftArrangeFormItems()
                End Using
            End If

            RectangleRecycleBin.Fill = New SolidColorBrush(Colors.Transparent)
            RectangleRecycleBinBackground.Visibility = Visibility.Hidden
        End If
    End Sub
    Private Sub RectangleExportMoveHandler(sender As Object, e As Input.MouseEventArgs) Handles RectangleExport.MouseMove
        If Not e.LeftButton = Input.MouseButtonState.Pressed Then
            RectangleExport.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            RectangleExportBackground.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub RectangleExportLeaveHandler(sender As Object, e As EventArgs) Handles RectangleExport.MouseLeave, RectangleExport.TouchLeave
        RectangleExport.Fill = New SolidColorBrush(Colors.Transparent)
        RectangleExportBackground.Visibility = Visibility.Hidden
    End Sub
    Private Sub RectangleExport_DragEnter(sender As Object, e As DragEventArgs) Handles RectangleExport.DragEnter
        If e.Data.GetDataPresent(DragBaseFormItem) Then
            Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)
            If (Not oDragData.Item2.Equals(Guid.Empty)) AndAlso BaseFormItem.FormMain.FindChild(oDragData.Item2).GetType.Equals(GetType(FormSection)) Then
                RectangleExport.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
            Else
                RectangleExport.Fill = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
            End If
        Else
            RectangleExport.Fill = New SolidColorBrush(Color.FromArgb(&H33, &HFF, &H0, &H0))
        End If
        RectangleExportBackground.Visibility = Visibility.Visible
    End Sub
    Private Sub RectangleExport_DragLeave(sender As Object, e As DragEventArgs) Handles RectangleExport.DragLeave
        RectangleExport.Fill = New SolidColorBrush(Colors.Transparent)
        RectangleExportBackground.Visibility = Visibility.Hidden
    End Sub
    Private Sub RectangleExport_Drop(sender As Object, e As DragEventArgs) Handles RectangleExport.Drop
        ' handles drop event
        If e.Data.GetDataPresent(DragBaseFormItem) Then
            Dim oDragData As Tuple(Of Type, Guid, Boolean, String) = e.Data.GetData(DragBaseFormItem)
            If (Not oDragData.Item2.Equals(Guid.Empty)) AndAlso BaseFormItem.FormMain.FindChild(oDragData.Item2).GetType.Equals(GetType(FormSection)) Then
                ' saves section contents to a file
                Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
                oSaveFileDialog.FileName = String.Empty
                oSaveFileDialog.DefaultExt = "*.gz"
                oSaveFileDialog.Filter = "GZip Files|*.gz"
                oSaveFileDialog.Title = "Save Section Contents To File"
                oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
                Dim result? As Boolean = oSaveFileDialog.ShowDialog()
                If result = True Then
                    ' serialise main form
                    Dim oSection As FormSection = BaseFormItem.FormMain.FindChild(oDragData.Item2).Clone

                    ' remove all page breaks
                    Dim oPageBreakList As List(Of FormatterPageBreak) = oSection.GetFormItems(Of FormatterPageBreak)
                    For Each oPageBreak In oPageBreakList
                        oPageBreak.Parent = Nothing
                    Next

                    ' remove all child items from the form main dictionary
                    Dim oFormItemList As List(Of BaseFormItem) = oSection.GetFormItems(Nothing)
                    For Each oFormItem As BaseFormItem In oFormItemList
                        BaseFormItem.FormMain.RemoveChild(oFormItem)
                    Next

                    ' remove the cloned section
                    oSection.Parent = Nothing

                    ' add the cloned section to the form item list
                    oFormItemList.Add(oSection)

                    CommonFunctions.SerializeDataContractFile(oSaveFileDialog.FileName, oFormItemList, LocalKnownTypes, False)

                    oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Section contents saved to file."))
                End If

                RectangleExport.Fill = New SolidColorBrush(Colors.Transparent)
                RectangleExportBackground.Visibility = Visibility.Hidden
            End If
        End If
    End Sub
    Private Sub RectangleFormScrollUp_Drop(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollUp.Drop
        m_RectangleFormScrollUp.Fill = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub RectangleFormScrollUp_DragEnter(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollUp.DragEnter
        m_RectangleFormScrollUp.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
    End Sub
    Private Sub RectangleFormScrollUp_DragLeave(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollUp.DragLeave
        m_RectangleFormScrollUp.Fill = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub RectangleFormScrollUp_DragOver(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollUp.DragOver
        Dim oAction As Action = Sub()
                                    CType(BaseFormItem.Root.GridFormScrollViewer, Controls.ScrollViewer).LineUp()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub RectangleFormScrollDown_Drop(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollDown.Drop
        m_RectangleFormScrollDown.Fill = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub RectangleFormScrollDown_DragEnter(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollDown.DragEnter
        m_RectangleFormScrollDown.Fill = New SolidColorBrush(Color.FromArgb(&H33, &H0, &HFF, &H0))
    End Sub
    Private Sub RectangleFormScrollDown_DragLeave(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollDown.DragLeave
        m_RectangleFormScrollDown.Fill = New SolidColorBrush(Colors.Transparent)
    End Sub
    Private Sub RectangleFormScrollDown_DragOver(sender As Object, e As DragEventArgs) Handles m_RectangleFormScrollDown.DragOver
        Dim oAction As Action = Sub()
                                    CType(BaseFormItem.Root.GridFormScrollViewer, Controls.ScrollViewer).LineDown()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
#End Region
End Class

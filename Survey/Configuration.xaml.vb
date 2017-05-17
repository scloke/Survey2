Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports Common.Common

Public Class Configuration
#Region "Variables"
    ' PluginName defines the friendly name of the plugin
    ' Priority determines the order in which the buttons are arranged on the main page. The lower the number, the earlier it is placed
    Private Const PluginName As String = "Configuration"
    Private Shared m_Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer))
    Private Shared ButtonDictionary As New Dictionary(Of String, FrameworkElement)
#End Region
#Region "Main"
    Sub New(value As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)))
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        MessageHub1.Messages = oCommonVariables.Messages

        CommonFunctions.IterateElementDictionary(ButtonDictionary, GridMain, "Button")

        ' set buttons to link to the other plugins
        m_Identifiers = value
        Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
        For i = 0 To 3
            Dim oButton As Controls.Button = ButtonDictionary("Button" + i.ToString)
            Dim oImage As Controls.Image = ButtonDictionary("ButtonImage" + i.ToString)

            If i < oFilteredNames.Count Then
                oButton.ToolTip = "Go To " + m_Identifiers(oFilteredNames(i)).Item3
                oButton.IsEnabled = True
                oImage.Source = m_Identifiers(oFilteredNames(i)).Item2
            Else
                oButton.ToolTip = String.Empty
                oButton.IsEnabled = False
                oImage.Source = Nothing
            End If
        Next

        SetIcons()
        SetBindings()
        SetDeveloper()
    End Sub
    Private Sub SetBindings()
        Dim oBinding1 As New Binding
        oBinding1.Path = New PropertyPath("DefaultSave")
        oBinding1.Mode = BindingMode.TwoWay
        oBinding1.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        ConfigSaveLocation.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBinding1)

        Dim oBinding2 As New Binding
        oBinding2.Path = New PropertyPath("RenderResolutionText")
        oBinding2.Mode = BindingMode.TwoWay
        oBinding2.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        RenderResolution.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBinding2)

        Dim oBinding3 As New Binding
        oBinding3.Path = New PropertyPath("DeveloperModeText")
        oBinding3.Mode = BindingMode.TwoWay
        oBinding3.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        DeveloperMode.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBinding3)

        Dim oBinding4 As New Binding
        oBinding4.Path = New PropertyPath("HelpFileText")
        oBinding4.Mode = BindingMode.TwoWay
        oBinding4.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        HelpFile.SetBinding(Common.HighlightTextBox.HTBTextProperty, oBinding4)

        ConfigSaveLocation.DataContext = oSettings
        RenderResolution.DataContext = oSettings
        DeveloperMode.DataContext = oSettings
        HelpFile.DataContext = oSettings
    End Sub
    Private Sub SetIcons()
        ConfigSelectSave.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMSave)
        RenderLower.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMDown)
        RenderHigher.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMUp)
        DeveloperModeOn.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMTick)
        DeveloperModeOff.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMCross)
        LoadHelpFile.HBSource = Converter.XamlToDrawingImage(My.Resources.CCMQuestionMark)
    End Sub
    Private Sub SetDeveloper()
        ' sets visibility of developer items
        If oSettings.DeveloperMode Then
            RectangleHelpFile.Visibility = Visibility.Visible
            DockPanelHelpFile.Visibility = Visibility.Visible
        Else
            RectangleHelpFile.Visibility = Visibility.Hidden
            DockPanelHelpFile.Visibility = Visibility.Hidden
        End If
    End Sub
    Public Sub Exit_Button_Click(oActivatePluginGUID As Guid)
        Me.WindowTitle = ModuleName
        Me.OnReturn(New ReturnEventArgs(Of Object)(oActivatePluginGUID))
    End Sub
#End Region
#Region "FunctionInterface"
    WriteOnly Property Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer))
        Set(value As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)))
        End Set
    End Property
    ' the status message event handler to update the main program status window
    Public Event StatusMessage(oMessage As Messages.Message)
#End Region
#Region "Buttons"
    ' add processing for the plugin buttons here
    Private Sub Exit_Button_Click(sender As Object, e As RoutedEventArgs)
        Exit_Button_Click(Guid.Empty)
    End Sub
    Private Sub LoadLink_Button_Click(sender As Object, e As RoutedEventArgs)
        ' loads linked plugin
        Dim oButton As Controls.Button = sender
        Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
        Dim iButton As Integer = Val(Right(oButton.Name, Len(oButton.Name) - Len("Button")))

        Exit_Button_Click(m_Identifiers(oFilteredNames(iButton)).Item1)
    End Sub
    Private Sub ShowInfo_Button_Click(sender As Object, e As RoutedEventArgs)
        ' shows the information screen
        Dim oHelpDialog As New HelpDialog(GridMain.ActualWidth * 3 / 5, GridMain.ActualHeight * 22 / 23, GridButtons.ColumnDefinitions(0).ActualWidth / 2, GridMain.RowDefinitions(1).ActualHeight)
        oHelpDialog.GUID = New Guid("498200b8-be87-441d-9ef3-e5963c31ff02")
        oHelpDialog.IconLabel = HelpDialog.IconType.Info
        oHelpDialog.ShowDialog()
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
    Private Sub ConfigSelectSaveHandler(sender As Object, e As EventArgs) Handles ConfigSelectSave.Click
        Dim oAction As Action = Sub()
                                    Dim oFolderBrowserDialog As New Forms.FolderBrowserDialog
                                    oFolderBrowserDialog.Description = "Select Default Save Location"
                                    oFolderBrowserDialog.ShowNewFolderButton = True
                                    oFolderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop
                                    If oFolderBrowserDialog.ShowDialog = Forms.DialogResult.OK Then
                                        oSettings.DefaultSave = oFolderBrowserDialog.SelectedPath
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub ConfigSelectClearHandler(sender As Object, e As EventArgs) Handles ConfigSelectSave.RightClick
        Dim oAction As Action = Sub()
                                    If MessageBox.Show("Clear default save?", ModuleName, MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                                        oSettings.DefaultSave = String.Empty
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub RenderLowerHandler(sender As Object, e As EventArgs) Handles RenderLower.Click
        Dim oAction As Action = Sub()
                                    Select Case oSettings.RenderResolution
                                        Case BaseFunctions.BaseFunctions.RenderResolution.High
                                            oSettings.RenderResolution = BaseFunctions.BaseFunctions.RenderResolution.Medium
                                        Case BaseFunctions.BaseFunctions.RenderResolution.Medium
                                            oSettings.RenderResolution = BaseFunctions.BaseFunctions.RenderResolution.Low
                                        Case Else
                                    End Select
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub RenderHigherHandler(sender As Object, e As EventArgs) Handles RenderHigher.Click
        Dim oAction As Action = Sub()
                                    Select Case oSettings.RenderResolution
                                        Case BaseFunctions.BaseFunctions.RenderResolution.Low
                                            oSettings.RenderResolution = BaseFunctions.BaseFunctions.RenderResolution.Medium
                                        Case BaseFunctions.BaseFunctions.RenderResolution.Medium
                                            oSettings.RenderResolution = BaseFunctions.BaseFunctions.RenderResolution.High
                                        Case Else
                                    End Select
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub DeveloperModeOnHandler(sender As Object, e As EventArgs) Handles DeveloperModeOn.Click
        Dim oAction As Action = Sub()
                                    oSettings.DeveloperMode = True
                                    SetDeveloper()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub DeveloperModeOffHandler(sender As Object, e As EventArgs) Handles DeveloperModeOff.Click
        Dim oAction As Action = Sub()
                                    oSettings.DeveloperMode = False
                                    SetDeveloper()
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub LoadHelpFileHandler(sender As Object, e As EventArgs) Handles LoadHelpFile.Click
        Dim oAction As Action = Sub()
                                    Dim oOpenFileDialog As New Microsoft.Win32.OpenFileDialog
                                    oOpenFileDialog.FileName = String.Empty
                                    oOpenFileDialog.DefaultExt = "*.shp"
                                    oOpenFileDialog.Multiselect = False
                                    oOpenFileDialog.Filter = "Survey Help File|*.shp"
                                    oOpenFileDialog.Title = "Load Help File"
                                    Dim result? As Boolean = oOpenFileDialog.ShowDialog()
                                    If result = True Then
                                        If IO.File.Exists(oOpenFileDialog.FileName) Then
                                            Dim oFileInfo As New IO.FileInfo(oOpenFileDialog.FileName)
                                            Dim oHelpDictionary As HelpFile = CommonFunctions.DeserializeDataContractFile(Of HelpFile)(oOpenFileDialog.FileName, New List(Of Type) From {GetType(HelpFile)}, False, , String.Empty)
                                            If IsNothing(oHelpDictionary) Then
                                                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Red, Date.Now, "Error deserialising " + oFileInfo.Name + "."))
                                            Else
                                                oHelpDictionary.Name = oFileInfo.Name
                                                oSettings.HelpFile = oHelpDictionary

                                                oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Help file " + oFileInfo.Name + " loaded."))
                                            End If
                                        End If
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
    Private Sub SaveHelpTemplateHandler(sender As Object, e As EventArgs) Handles LoadHelpFile.RightClick
        Dim oAction As Action = Sub()
                                    Dim oSaveFileDialog As New Microsoft.Win32.SaveFileDialog
                                    oSaveFileDialog.FileName = String.Empty
                                    oSaveFileDialog.DefaultExt = "*.xml"
                                    oSaveFileDialog.Filter = "XML Files|*.xml"
                                    oSaveFileDialog.Title = "Save Help File Template"
                                    oSaveFileDialog.InitialDirectory = oSettings.DefaultSave
                                    Dim result? As Boolean = oSaveFileDialog.ShowDialog()
                                    If result = True Then
                                        IO.File.WriteAllText(oSaveFileDialog.FileName, My.Resources.Survey_v2_Help_File)
                                        Dim oFileInfo As New IO.FileInfo(oSaveFileDialog.FileName)

                                        oCommonVariables.Messages.Add(New Messages.Message(PluginName, Colors.Green, Date.Now, "Help file template " + oFileInfo.Name + " saved."))
                                    End If
                                End Sub
        CommonFunctions.RepeatCheck(sender, e, oAction)
    End Sub
#End Region
End Class

Imports Common.Common

Class MainPage
#Region "ProgramConstantsAndVariables"
    Private WithEvents oCommonPage As CommonPage
    Private WithEvents oConfiguration As Configuration
    Private oPlugInList As New List(Of Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer))
    Private Shared PlugInDirectory As String = AppDomain.CurrentDomain.BaseDirectory + "PlugIn\"
    Private Shared oIdentifierList As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)) = Nothing
#End Region
#Region "Main"
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.KeepAlive = True

        If PDFHelper.Initialise() Then
            oPlugInList.AddRange(LoadPlugIns)
            oPlugInList = (From oPlugIn As Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer) In oPlugInList Order By oPlugIn.Item5 Ascending Select oPlugIn).ToList
            oIdentifierList = (From oPlugIn As Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer) In oPlugInList Select New KeyValuePair(Of String, Tuple(Of Guid, ImageSource, String, Integer))(oPlugIn.Item4, New Tuple(Of Guid, ImageSource, String, Integer)(oPlugIn.Item2, oPlugIn.Item3, oPlugIn.Item4, oPlugIn.Item5))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
            For i = 0 To Math.Min(oPlugInList.Count - 1, 5)
                Dim oPlugIn As Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer) = oPlugInList(i)
                Dim oButtonCommon As Button = Grid1.FindName("ButtonCommon" + i.ToString)
                Dim oImageCommon As Image = Grid1.FindName("ImageCommon" + i.ToString)
                Dim oLabelCommon As Label = Grid1.FindName("LabelCommon" + i.ToString)

                oButtonCommon.Visibility = Visibility.Visible
                oButtonCommon.Tag = oPlugIn.Item2
                oImageCommon.Source = oPlugIn.Item3
                oLabelCommon.Content = oPlugIn.Item4

                ' connect status message handler
                AddHandler oPlugIn.Item1.StatusMessage, AddressOf StatusMessageHandler

                ' initialise data store variables
                Dim oDataTypes As List(Of Tuple(Of Guid, Type)) = oPlugIn.Item1.GetDataTypes
                For Each oDataType In oDataTypes
                    oCommonVariables.DataStore.TryAdd(oDataType.Item1, Activator.CreateInstance(oDataType.Item2))
                Next
            Next

            ' runs a final check to see if all variables initialised properly
            ' this is to screen for plugins with conflicting GUIDs
            For i = 0 To Math.Min(oPlugInList.Count - 1, 5)
                Dim oPlugIn As Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer) = oPlugInList(i)
                If Not oPlugIn.Item1.CheckDataTypes(oCommonVariables) Then
                    MessageBox.Show("Plug-in " + oPlugIn.Item4 + " not initialised properly. Exiting " + ModuleName, ModuleName, MessageBoxButton.OK, MessageBoxImage.Error)
                    ExitRoutine()
                End If
            Next

            For i = 0 To Math.Min(oPlugInList.Count - 1, 5)
                Dim oPlugIn As Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer) = oPlugInList(i)

                ' set identifier list
                oPlugIn.Item1.Identifiers = oIdentifierList
            Next
        Else
            MessageBox.Show("Error: PDFSharp page unit error.", ModuleName, MessageBoxButton.OK, MessageBoxImage.Error)
            ExitRoutine()
        End If
    End Sub
    Private Sub ReturnPageHandler(sender As Object, e As ReturnEventArgs(Of Object)) Handles oCommonPage.Return, oConfiguration.Return
        If Not IsNothing(e) Then
            Dim oReturnGUID As Guid = CType(e.Result, Guid)
            If Not oReturnGUID.Equals(Guid.Empty) Then
                ' activate associated plugin
                Dim oButtonList As List(Of Button) = (From iIndex As Integer In Enumerable.Range(0, Math.Min(oPlugInList.Count, 5)) Where oPlugInList(iIndex).Item2.Equals(oReturnGUID) Select CType(Grid1.FindName("ButtonCommon" + iIndex.ToString), Button)).ToList
                If oButtonList.Count > 0 Then
                    Common_Button_Click(oButtonList.First, New RoutedEventArgs)
                End If
            End If
        End If
    End Sub
    Private Sub ExitRoutine()
        ' cancels tasks and wait until they have completed
        oCommonVariables.Messages.Add(New Messages.Message("Main", Colors.Red, Date.Now, "Closing " + ModuleName))
        For i = 0 To Math.Min(oPlugInList.Count - 1, 5)
            Dim oPlugIn As Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer) = oPlugInList(i)
            RemoveHandler oPlugIn.Item1.StatusMessage, AddressOf StatusMessageHandler
        Next

        Application.Current.Shutdown()
    End Sub
    Private Sub StatusMessageHandler(oMessage As Messages.Message)
        If Not IsNothing(oCommonVariables) AndAlso (Not IsNothing(oMessage)) Then
            oCommonVariables.Messages.Add(oMessage)
        End If
    End Sub
    Private Shared Function LoadPlugIns() As List(Of Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer))
        ' returns a list containing the plug-in interface, the GUID, a bitmap with the icon, the friendly name, and the priority
        Dim oPlugInList As New List(Of Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer))
        Dim oDirectoryInfo As New IO.DirectoryInfo(PlugInDirectory)
        Dim oDirectoryList As List(Of IO.DirectoryInfo) = oDirectoryInfo.EnumerateDirectories.ToList
        For Each oDirectory In oDirectoryList
            Dim sDLLFile As String = oDirectory.FullName + "\" + oDirectory.Name + ".dll"
            If IO.File.Exists(sDLLFile) Then
                Dim oHash As Byte() = GenerateMD5(sDLLFile)
                Try
                    Dim oAssembly As Reflection.Assembly = Reflection.Assembly.LoadFrom(sDLLFile, oHash, System.Configuration.Assemblies.AssemblyHashAlgorithm.MD5)
                    For Each oType As Type In oAssembly.GetTypes
                        If Not IsNothing(oType.GetInterface("FunctionInterface")) Then
                            Dim oFunctionInterface As FunctionInterface = Activator.CreateInstance(oType)
                            Dim oIdentifier As Tuple(Of Guid, ImageSource, String, Integer) = oFunctionInterface.GetIdentifier
                            oPlugInList.Add(New Tuple(Of FunctionInterface, Guid, ImageSource, String, Integer)(oFunctionInterface, oIdentifier.Item1, oIdentifier.Item2, oIdentifier.Item3, oIdentifier.Item4))
                        End If
                    Next
                Catch ex As Exception
                End Try
            End If
        Next
        Return oPlugInList
    End Function
    Private Shared Function GenerateMD5(ByVal sFilename As String) As Byte()
        ' returns hash from file
        Dim oReturnHash As Byte() = Nothing
        Using oFileStream As IO.FileStream = IO.File.OpenRead(sFilename)
            oReturnHash = Security.Cryptography.MD5.Create().ComputeHash(oFileStream)
        End Using
        Return oReturnHash
    End Function
#End Region
#Region "Buttons"
    Private Sub Common_Button_Click(sender As Object, e As RoutedEventArgs)
        ' activates the respective buttons
        Dim oButton As Button = CType(sender, Button)
        Dim iButtonNumber As Integer = Val(Right(oButton.Name, oButton.Name.Length - "ButtonCommon".Length))

        WindowTitle = ModuleName + " - " + oPlugInList(iButtonNumber).Item4

        ' remove template control and dispose of old page if needed
        If Not IsNothing(oCommonPage) Then
            For i = oCommonPage.Grid1.Children.Count - 1 To 0 Step -1
                Dim oControl = oCommonPage.Grid1.Children(i)
                Dim iCount As Integer = Aggregate oPlugin In oPlugInList Where oPlugin.Item1.GetType.Equals(oControl.GetType) Into Count()
                If iCount > 0 Then
                    RemoveHandler CType(oControl, FunctionInterface).ExitButtonClick, AddressOf oCommonPage.Exit_Button_Click
                    oCommonPage.Grid1.Children.Remove(oControl)
                End If
            Next
            oCommonPage = Nothing
        End If

        oCommonPage = New CommonPage()
        Dim oTemplateControl As UserControl = CType(oPlugInList(iButtonNumber).Item1, UserControl)
        Dim oFunctionInterface As FunctionInterface = CType(oTemplateControl, FunctionInterface)
        oCommonPage.Grid1.Children.Add(oTemplateControl)
        Grid.SetColumn(oTemplateControl, 0)
        Grid.SetRow(oTemplateControl, 0)
        Grid.SetColumnSpan(oTemplateControl, 1)
        Grid.SetRowSpan(oTemplateControl, 2)
        AddHandler oFunctionInterface.ExitButtonClick, AddressOf oCommonPage.Exit_Button_Click

        NavigationService.Navigate(oCommonPage)
    End Sub
    Private Sub Configuration_Button_Click(sender As Object, e As RoutedEventArgs)
        ' configuration page
        Me.WindowTitle = ModuleName + " - Configuration"

        ' remove template control and dispose of old page if needed
        If Not IsNothing(oConfiguration) Then
            RemoveHandler oConfiguration.StatusMessage, AddressOf StatusMessageHandler
            oConfiguration = Nothing
        End If
        oConfiguration = New Configuration(oIdentifierList)
        AddHandler oConfiguration.StatusMessage, AddressOf StatusMessageHandler

        NavigationService.Navigate(oConfiguration)
    End Sub
    Private Sub Exit_Button_Click(sender As Object, e As RoutedEventArgs)
        If MessageBox.Show("Exit " + ModuleName + "?", ModuleName, MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
            ExitRoutine()
        End If
    End Sub
#End Region
End Class
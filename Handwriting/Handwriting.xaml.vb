Imports Common.Common
Imports System.Windows
Imports System.Windows.Media

Public Class Handwriting
    Implements FunctionInterface

#Region "Variables"
    ' PluginName defines the friendly name of the plugin
    ' Priority determines the order in which the buttons are arranged on the main page. The lower the number, the earlier it is placed
    Private Const PluginName As String = "Handwriting"
    Private Const Priority As Integer = 3

    Private Shared m_Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer))
    Private Shared ElementDictionary As New Dictionary(Of String, FrameworkElement)

    ' insert your own GUIDs here
#End Region

#Region "FunctionInterface"
    Public Function GetDataTypes() As List(Of Tuple(Of Guid, Type)) Implements FunctionInterface.GetDataTypes
        ' returns a list of GUIDs and variable type representing the data types that the plug-in creates
        Dim oDataTypes As New List(Of Tuple(Of Guid, Type))
        Return oDataTypes
    End Function
    Public Function CheckDataTypes(ByRef oCommonVariables As CommonVariables) As Boolean Implements FunctionInterface.CheckDataTypes
        ' checks the commonvariable data store to see if the variables objects required have been properly initialised
        Dim bCheck As Boolean = True

        Return bCheck
    End Function
    Public Function GetIdentifier() As Tuple(Of Guid, ImageSource, String, Integer) Implements FunctionInterface.GetIdentifier
        ' returns the identifiers: GUID, icon, friendly name, and priority for the plugin
        Return New Tuple(Of Guid, ImageSource, String, Integer)(Guid.NewGuid, Converter.BitmapToBitmapSource(My.Resources.IconHandwriting1), PluginName, Priority)
    End Function
    WriteOnly Property Identifiers As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)) Implements FunctionInterface.Identifiers
        Set(value As Dictionary(Of String, Tuple(Of Guid, ImageSource, String, Integer)))
            ' set buttons to link to the other plugins
            m_Identifiers = value
            Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
            For i = 0 To oFilteredNames.Count - 1
                Dim oButton As Controls.Button = ElementDictionary("Button" + i.ToString)
                Dim oImage As Controls.Image = ElementDictionary("ButtonImage" + i.ToString)
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
        CommonFunctions.IterateElementDictionary(ElementDictionary, GridMain)

        ' initially set all buttons to disabled and clear tooltips
        For i = 0 To 2
            Dim oButton As Controls.Button = ElementDictionary("Button" + i.ToString)
            Dim oImage As Controls.Image = ElementDictionary("ButtonImage" + i.ToString)
            oButton.ToolTip = String.Empty
            oButton.IsEnabled = False
            oImage.Source = Nothing
        Next
    End Sub
#End Region
#Region "Classes"
#End Region
#Region "Buttons"
    ' add processing for the plugin buttons here
    Private Sub Exit_Button_Click(sender As Object, e As RoutedEventArgs)
        'MainForm.ImageViewer.Dispose()
        RaiseEvent ExitButtonClick(Guid.Empty)
    End Sub
    Private Sub LoadLink_Button_Click(sender As Object, e As RoutedEventArgs)
        ' loads linked plugin
        Dim oButton As Controls.Button = sender
        Dim oFilteredNames As List(Of String) = (From sName As String In m_Identifiers.Keys Where sName <> PluginName Select sName).ToList
        Dim iButton As Integer = Val(Right(oButton.Name, Len(oButton.Name) - Len("Button")))

        'MainForm.ImageViewer.Dispose()
        RaiseEvent ExitButtonClick(m_Identifiers(oFilteredNames(iButton)).Item1)
    End Sub
    Private Sub Help_Button_Click(sender As Object, e As RoutedEventArgs)
        ' help button

    End Sub
#End Region
End Class
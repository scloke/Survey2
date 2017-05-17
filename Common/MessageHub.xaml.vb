Imports Common.Common
Imports System.Windows
Imports System.Windows.Controls

Public Class MessageHub
    Private WithEvents m_Messages As Messages
    Private Shared m_Icons As Dictionary(Of String, Media.ImageSource)
    Private oUIDispatcher As Threading.Dispatcher

    Public Property Messages As Messages
        Get
            Return m_Messages
        End Get
        Set(value As Messages)
            m_Messages = value
            m_Messages.SetDispatcher(oUIDispatcher)
            DataGridView1.DataContext = m_Messages.Messages
            Grid1.DataContext = m_Messages
            UpdateMessages()
        End Set
    End Property
    Public Sub New()
        InitializeComponent()

        oUIDispatcher = Threading.Dispatcher.CurrentDispatcher

        SetIcons()

        ButtonFilterImage.Source = GetIcon("IconGray")
    End Sub
    Private Sub SetIcons()
        m_Icons = New Dictionary(Of String, Media.ImageSource)
        m_Icons.Add("IconGray", Converter.BitmapToBitmapSource(My.Resources.IconGray))
        m_Icons.Add("IconBlue", Converter.BitmapToBitmapSource(My.Resources.IconBlue))
        m_Icons.Add("IconGreen", Converter.BitmapToBitmapSource(My.Resources.IconGreen))
        m_Icons.Add("IconRed", Converter.BitmapToBitmapSource(My.Resources.IconRed))
        m_Icons.Add("IconYellow", Converter.BitmapToBitmapSource(My.Resources.IconYellow))
    End Sub
    Public Shared Function GetIcon(ByVal sIconName As String) As Media.ImageSource
        If m_Icons.ContainsKey(sIconName) Then
            Return m_Icons(sIconName)
        Else
            Return Nothing
        End If
    End Function
    Private Sub UpdateMessages() Handles m_Messages.UpdateMessage
        If Not IsNothing(m_Messages) Then
            LabelCount.Content = m_Messages.SelectedCount
            LabelSource.Content = m_Messages.SelectedSourceTrimmed
            LabelDateTime.Content = m_Messages.SelectedDateTime
            LabelMessage.Content = m_Messages.SelectedMessageText
            CommonFunctions.RefreshControl(LabelMessage)
        End If
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DataGridView1.SelectionChanged
        If Not IsNothing(DataGridView1.SelectedItem) Then
            Messages.SelectedMessageIndex = DataGridView1.SelectedIndex
        Else
            If DataGridView1.Items.Count > 0 Then
                DataGridView1.SelectedIndex = 0
            End If
        End If
    End Sub
#Region "Buttons"
    Private Sub Message_Button_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf sender Is Controls.Button Then
            Dim oButton As Controls.Button = CType(sender, Controls.Button)
            Select Case oButton.Name
                Case "ButtonFirst"
                    Messages.SelectedMessageIndex = 0
                Case "ButtonBack"
                    If Messages.SelectedMessageIndex > 0 Then
                        Messages.SelectedMessageIndex -= 1
                    End If
                Case "ButtonForward"
                    If Messages.SelectedMessageIndex < Messages.Count - 1 Then
                        Messages.SelectedMessageIndex += 1
                    End If
                Case "ButtonLast"
                    Messages.SelectedMessageIndex = Messages.Count - 1
            End Select
            DataGridView1.SelectedIndex = Messages.SelectedMessageIndex
        End If
    End Sub
    Private Sub Expand_Button_Click(sender As Object, e As RoutedEventArgs)
        Dim oBitmap As System.Drawing.Bitmap = Nothing
        If Expander.Visibility = Windows.Visibility.Collapsed Then
            Expander.Visibility = Windows.Visibility.Visible
            oBitmap = My.Resources.IconEjectInverted
        Else
            Expander.Visibility = Windows.Visibility.Collapsed
            oBitmap = My.Resources.IconEject
        End If
        ImageEject.Source = Converter.BitmapToBitmapSource(oBitmap)
        oBitmap.Dispose()
    End Sub
    Private Sub Filter_Button_Click(sender As Object, e As RoutedEventArgs)
        ' changes filter colour
        Dim oCurrentColour As Media.Color = m_Messages.IncrementDisplayColour
        Select Case oCurrentColour
            Case Media.Colors.Black
                ButtonFilterImage.Source = GetIcon("IconGray")
            Case Media.Colors.Blue
                ButtonFilterImage.Source = GetIcon("IconBlue")
            Case Media.Colors.Green
                ButtonFilterImage.Source = GetIcon("IconGreen")
            Case Media.Colors.Red
                ButtonFilterImage.Source = GetIcon("IconRed")
            Case Media.Colors.Yellow
                ButtonFilterImage.Source = GetIcon("IconYellow")
            Case Else
                ButtonFilterImage.Source = GetIcon("IconGray")
        End Select
    End Sub
#End Region
End Class
Imports System.Collections.ObjectModel
Imports System.Windows

Public Class ScannerDialog
    Inherits Window

#Region "Variables"
    Public ScannerSources As New ObservableCollection(Of String)
    Public SelectedIndex As Integer
#End Region

    Public Sub New()
        Me.InitializeComponent()

        ListBoxScanner.ItemsSource = ScannerSources
    End Sub
#Region "Buttons"
    ' add processing for the buttons here
    Private Sub ScannerOK_Button_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
        SelectedIndex = ListBoxScanner.SelectedIndex
        Me.Close()
    End Sub
    Private Sub ScannerCancel_Button_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        SelectedIndex = -1
        Me.Close()
    End Sub
#End Region
End Class
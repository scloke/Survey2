Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports Common.Common
Imports System.IO

Class MainWindow
#Region "ProgramConstantsAndVariables"
    Private oMainPage As New MainPage
#End Region
#Region "MainProgram"
    Public Sub New()
        InitializeComponent()
        AppReload()
        AppInit()
    End Sub
    Protected Overrides Sub OnClosed(e As EventArgs)
        MyBase.OnClosed(e)
        AppClose()
        AppPersist()
    End Sub
    Private Sub AppReload()
        ' Load configuration data
        If My.Settings.SettingStore = String.Empty Then
            oSettings = New Settings
        Else
            Try
                Using oMemoryStream As New MemoryStream(Convert.FromBase64String(My.Settings.SettingStore))
                    oSettings = CommonFunctions.DeserializeDataContractStream(Of Settings)(oMemoryStream)
                    If IsNothing(oSettings) Then
                        oSettings = New Settings
                        MessageBox.Show("Settings file corrupted. New Settings file initialised.", ModuleName, MessageBoxButton.OK, MessageBoxImage.Warning)
                    End If
                End Using
            Catch ex As Exception
                oSettings = New Settings
                MessageBox.Show("Settings file corrupted. New Settings file initialised.", ModuleName, MessageBoxButton.OK, MessageBoxImage.Warning)
            End Try
        End If
    End Sub
    Private Sub AppPersist()
        ' Save configuration data
        Using oMemoryStream As New MemoryStream
            CommonFunctions.SerializeDataContractStream(Of Settings)(oMemoryStream, oSettings)
            My.Settings.SettingStore = Convert.ToBase64String(oMemoryStream.ToArray)
            My.Settings.Save()
        End Using
    End Sub
    Private Sub AppInit()
        NavigationService.Navigate(oMainPage)
    End Sub
    Private Sub AppClose()
    End Sub
#End Region
End Class

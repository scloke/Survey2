Imports Common.Common

Class CommonPage
#Region "ProgramConstantsAndVariables"
#End Region
#Region "MainProgram"
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        Me.KeepAlive = True

        ' Add any initialization after the InitializeComponent() call.
        MessageHub1.Messages = oCommonVariables.Messages
    End Sub
    Public Sub Exit_Button_Click(oActivatePluginGUID As Guid)
        Me.WindowTitle = ModuleName
        Me.OnReturn(New ReturnEventArgs(Of Object)(oActivatePluginGUID))
    End Sub
#End Region
End Class
Imports PdfSharp
Imports PdfSharp.Drawing
Imports BaseFunctions.BaseFunctions
Imports System.ComponentModel
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data
Imports System.Runtime.Serialization

Public Class BaseFunctions
    Public Const ScreenResolution096 As Single = 96
    Public Const ViewResolution150 As Single = 150
    Public Const RenderResolution300 As Single = 300
    Public Const MarkResolution600 As Single = 600
    Public Shared oSettings As Settings

    Public Enum RenderResolution As Integer
        Low = 0
        Medium
        High
    End Enum
End Class
Public Class Settings
    ' Persistent application variables
    Implements IDisposable, INotifyPropertyChanged

    Private m_DefaultSave As String
    Private m_DeveloperMode As Boolean
    Private m_RenderResolution As RenderResolution
    Private m_HelpFile As HelpFile

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChangedLocal(ByVal sName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
    End Sub
    Public Sub New()
        m_DefaultSave = String.Empty
        m_DeveloperMode = False
        m_RenderResolution = RenderResolution.High
        m_HelpFile = Nothing
    End Sub
    Public Property DefaultSave As String
        Get
            Return m_DefaultSave
        End Get
        Set(value As String)
            m_DefaultSave = value
            OnPropertyChangedLocal("DefaultSave")
        End Set
    End Property
    Public Property DeveloperMode As Boolean
        Get
            Return m_DeveloperMode
        End Get
        Set(value As Boolean)
            m_DeveloperMode = value
            OnPropertyChangedLocal("DeveloperMode")
            OnPropertyChangedLocal("DeveloperModeText")
        End Set
    End Property
    Public Property RenderResolution As RenderResolution
        Get
            Return m_RenderResolution
        End Get
        Set(value As RenderResolution)
            m_RenderResolution = value
            OnPropertyChangedLocal("RenderResolution")
            OnPropertyChangedLocal("RenderResolutionValue")
            OnPropertyChangedLocal("RenderResolutionText")
        End Set
    End Property
    Public Property HelpFile As HelpFile
        Get
            Return m_HelpFile
        End Get
        Set(value As HelpFile)
            m_HelpFile = value
            OnPropertyChangedLocal("HelpFile")
            OnPropertyChangedLocal("HelpFileText")
        End Set
    End Property
    Public ReadOnly Property RenderResolutionValue As Single
        Get
            Select Case m_RenderResolution
                Case RenderResolution.Low
                    Return ScreenResolution096
                Case RenderResolution.Medium
                    Return ViewResolution150
                Case RenderResolution.High
                    Return RenderResolution300
                Case Else
                    Return 0
            End Select
        End Get
    End Property
    Public ReadOnly Property RenderResolutionValueMax As Single
        Get
            Select Case m_RenderResolution
                Case RenderResolution.Low
                    Return ViewResolution150
                Case RenderResolution.Medium
                    Return RenderResolution300
                Case RenderResolution.High
                    Return MarkResolution600
                Case Else
                    Return 0
            End Select
        End Get
    End Property
    Public Property DeveloperModeText As String
        Get
            Return If(m_DeveloperMode, "On", "Off")
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property RenderResolutionText As String
        Get
            Return [Enum].GetName(GetType(RenderResolution), m_RenderResolution)
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property HelpFileText As String
        Get
            If IsNothing(m_HelpFile) Then
                Return String.Empty
            Else
                Return m_HelpFile.Name
            End If
        End Get
        Set(value As String)
        End Set
    End Property
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
            End If
        End SyncLock
    End Sub
#End Region
End Class
<DataContract> Public Class HelpFile
    Implements INotifyPropertyChanged

    Private m_Dictionary As New HelpDictionary
    Private m_Name As String = String.Empty

#Region "INotifyPropertyChanged"
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private Sub OnPropertyChangedLocal(ByVal sName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
    End Sub
#End Region
    Public Sub New()
        MyBase.New()
    End Sub
    <DataMember> Public Property Dictionary As HelpDictionary
        Get
            Return m_Dictionary
        End Get
        Set(value As HelpDictionary)
            m_Dictionary = value
        End Set
    End Property
    <DataMember> Public Property Name As String
        Get
            Return m_Name
        End Get
        Set(value As String)
            m_Name = value
            OnPropertyChangedLocal("Name")
        End Set
    End Property
    Public Function GetPDFDocument(ByVal oGUID As Guid) As Pdf.PdfDocument
        ' returns PDF document from store
        Dim oPDFDocument As Pdf.PdfDocument = Nothing
        If m_Dictionary.ContainsKey(oGUID) Then
            Using oMemoryStream As New IO.MemoryStream(m_Dictionary(oGUID))
                oPDFDocument = Pdf.IO.PdfReader.Open(oMemoryStream)
            End Using
        End If
        Return oPDFDocument
    End Function
    <CollectionDataContract> Public Class HelpDictionary
        Inherits Dictionary(Of Guid, Byte())
    End Class
End Class
#Region "WPFConverters"
<ValueConversion(GetType(Double), GetType(Double))> Public Class MultiplierConverter
    Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim fReturn As Double = CDbl(value) * CDbl(parameter)
        Return If(fReturn = 0, 1, fReturn)
    End Function
    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Return If(CDbl(parameter) = 0, CDbl(1), CDbl(value) / CDbl(parameter))
    End Function
End Class
Public Class MultiplierConverter2
    Implements IMultiValueConverter

    Public Function Convert(ByVal values() As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
        Dim fReturn As Double = 0
        If (Not (IsNothing(values(0)) Or IsNothing(values(1)))) AndAlso (Not (values(0).Equals(DependencyProperty.UnsetValue) Or values(1).Equals(DependencyProperty.UnsetValue))) Then
            fReturn = values(0) * values(1)
        End If
        Return If(fReturn = 0, 1, fReturn)
    End Function
    Public Function ConvertBack(ByVal value As Object, ByVal targetTypes() As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
        Throw New NotSupportedException
    End Function
End Class
<ValueConversion(GetType(Controls.DataGridRow), GetType(String))> Public Class RowToIndexConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim row As Controls.DataGridRow = TryCast(value, Controls.DataGridRow)
        If Not IsNothing(row) Then
            Return (row.GetIndex + 1).ToString
        Else
            Return String.Empty
        End If
    End Function
    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
#End Region

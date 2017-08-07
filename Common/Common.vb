Imports PdfSharp
Imports PdfSharp.Drawing
Imports BaseFunctions
Imports BaseFunctions.BaseFunctions
Imports Twain32Shared.Twain32Shared
Imports System.Collections.Concurrent
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Runtime.Serialization
Imports System.Collections.Specialized
Imports System.Windows
Imports System.Windows.Input

Public Class Common
#Region "ProgramConstantsAndVariables"
    Public Const ModuleName As String = "Survey"
    Public Const FontArial As String = "Arial"
    Public Const FontComicSansMS As String = "Comic Sans MS"
    Public Const MessageLimit As Integer = 100
    Public Const TaskCount As Integer = 8
    Public Const SeparatorChar As Char = "|"
    Public Const SpacingSmall As Single = 5
    Public Const SpacingLarge As Single = 15
    Public Const TemplateMatchCutoff As Double = 0.5
    Public Const BoxSize As Integer = 24
    Public Const ScaleFactor As Integer = 5
    Public Shared ParallelDictionary As New Dictionary(Of Integer, Tuple(Of List(Of Integer), List(Of Integer)))
    Public Shared WithEvents oCommonVariables As New CommonVariables
#End Region
#Region "Classes"
    Public Class CommonVariables
        ' Non-persistent application variables
        Implements IDisposable

        Public WithEvents Messages As Messages
        Public DataStore As New ConcurrentDictionary(Of Guid, Object)

        Public Sub New()
            Messages = New Messages
        End Sub
        Public Sub New(ByVal oCommonVariables As CommonVariables)
            Messages = oCommonVariables.Messages
            DataStore = oCommonVariables.DataStore
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
                End If
            End SyncLock
        End Sub
#End Region
    End Class
    Public Class Messages
        ' represents groups of messages with a selected message
        Implements INotifyPropertyChanged

        Private m_Messages As New ObservableCollection(Of Message)
        Private m_FilteredMessages As New ObservableCollection(Of Message)
        Private m_SelectedMessage As Integer = 0
        Private m_FilterColour As Integer = 0
        Private oMessageDispatcher As Threading.Dispatcher
        Public Shared Event UpdateMessage()

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub OnPropertyChanged(ByVal sName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
        End Sub
        Public Sub New()
        End Sub
        Public Sub SetDispatcher(ByVal oDispatcher As Threading.Dispatcher)
            ' sets the message dispatcher
            oMessageDispatcher = oDispatcher
        End Sub
        Public Sub Add(ByVal oMessage As Message)
            Dim emptyDelegate As Action = Sub()
                                              m_Messages.Add(oMessage)
                                              TruncateLimit()
                                              OnPropertyChanged(String.Empty)
                                              RaiseEvent UpdateMessage()
                                          End Sub
            InvokeDelegate(emptyDelegate)
        End Sub
        Public Sub AddRange(ByVal oMessages As ObservableCollection(Of Message))
            Dim emptyDelegate As Action = Sub()
                                              For Each oMessage In oMessages
                                                  m_Messages.Add(oMessage)
                                              Next
                                              TruncateLimit()
                                              OnPropertyChanged(String.Empty)
                                              RaiseEvent UpdateMessage()
                                          End Sub
            InvokeDelegate(emptyDelegate)
        End Sub
        Public Sub Clear()
            Dim emptyDelegate As Action = Sub()
                                              m_Messages.Clear()
                                              OnPropertyChanged(String.Empty)
                                              RaiseEvent UpdateMessage()
                                          End Sub
            InvokeDelegate(emptyDelegate)
        End Sub
        Public Sub Remove(ByVal oMessage As Message)
            Dim emptyDelegate As Action = Sub()
                                              m_Messages.Remove(oMessage)
                                              OnPropertyChanged(String.Empty)
                                              RaiseEvent UpdateMessage()
                                          End Sub
            InvokeDelegate(emptyDelegate)
        End Sub
        Public Sub RemoveAt(ByVal iIndex As Integer)
            Dim emptyDelegate As Action = Sub()
                                              m_Messages.RemoveAt(iIndex)
                                              OnPropertyChanged(String.Empty)
                                              RaiseEvent UpdateMessage()
                                          End Sub
            InvokeDelegate(emptyDelegate)
        End Sub
        Public ReadOnly Property Count() As Integer
            Get
                Return m_Messages.Count
            End Get
        End Property
        Public Function IncrementDisplayColour() As Media.Color
            Dim emptyDelegate As Func(Of Media.Color) = Function()
                                                            Dim oReturnColour As Media.Color
                                                            m_FilterColour = (m_FilterColour + 1) Mod 5
                                                            Select Case m_FilterColour
                                                                Case 0
                                                                    oReturnColour = Media.Colors.Black
                                                                Case 1
                                                                    oReturnColour = Media.Colors.Blue
                                                                Case 2
                                                                    oReturnColour = Media.Colors.Green
                                                                Case 3
                                                                    oReturnColour = Media.Colors.Red
                                                                Case 4
                                                                    oReturnColour = Media.Colors.Yellow
                                                                Case Else
                                                                    oReturnColour = Media.Colors.Black
                                                            End Select
                                                            UpdateFilteredMessages()
                                                            OnPropertyChanged(String.Empty)
                                                            Return oReturnColour
                                                        End Function
            If IsNothing(oMessageDispatcher) OrElse oMessageDispatcher.CheckAccess Then
                Return emptyDelegate.Invoke
            Else
                Try
                    ' make sure that the UI thread is not blocked with something like a task "waitall", otherwise this will hang
                    Return oMessageDispatcher.Invoke(emptyDelegate, Threading.DispatcherPriority.Render)
                Catch ex As TaskCanceledException
                End Try
            End If
        End Function
        Public Property SelectedMessageIndex() As Integer
            Get
                Return m_SelectedMessage
            End Get
            Set(value As Integer)
                Dim emptyDelegate As Action = Sub()
                                                  If m_FilteredMessages.Count > 0 AndAlso (value >= 0 And value < m_FilteredMessages.Count) Then
                                                      m_SelectedMessage = value
                                                  Else
                                                      m_SelectedMessage = 0
                                                  End If
                                                  OnPropertyChanged(String.Empty)
                                              End Sub
                InvokeDelegate(emptyDelegate)
            End Set
        End Property
        Public ReadOnly Property SelectedMessage() As Message
            Get
                If m_FilteredMessages.Count > 0 AndAlso (m_SelectedMessage >= 0 And m_SelectedMessage < m_FilteredMessages.Count) Then
                    Return m_FilteredMessages(m_SelectedMessage)
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedCount As String
            Get
                If Not IsNothing(m_FilteredMessages) AndAlso m_FilteredMessages.Count > 0 AndAlso (m_SelectedMessage >= 0 And m_SelectedMessage < m_FilteredMessages.Count) Then
                    Return (m_SelectedMessage + 1).ToString + "/" + m_FilteredMessages.Count.ToString
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedSource As String
            Get
                If Not IsNothing(m_FilteredMessages) AndAlso m_FilteredMessages.Count > 0 AndAlso (m_SelectedMessage >= 0 And m_SelectedMessage < m_FilteredMessages.Count) Then
                    Return m_FilteredMessages(m_SelectedMessage).Source
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedSourceTrimmed As String
            Get
                If Not IsNothing(m_FilteredMessages) AndAlso m_FilteredMessages.Count > 0 AndAlso (m_SelectedMessage >= 0 And m_SelectedMessage < m_FilteredMessages.Count) Then
                    Return Left(Trim(m_FilteredMessages(m_SelectedMessage).Source), 12)
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedDateTime As String
            Get
                If Not IsNothing(m_FilteredMessages) AndAlso m_FilteredMessages.Count > 0 AndAlso (m_SelectedMessage >= 0 And m_SelectedMessage < m_FilteredMessages.Count) Then
                    Return m_FilteredMessages(m_SelectedMessage).DateTimeText
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public ReadOnly Property SelectedMessageText As String
            Get
                If Not IsNothing(m_FilteredMessages) AndAlso m_FilteredMessages.Count > 0 AndAlso (m_SelectedMessage >= 0 And m_SelectedMessage < m_FilteredMessages.Count) Then
                    Return m_FilteredMessages(m_SelectedMessage).Message
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public Property Messages As ObservableCollection(Of Message)
            Get
                Return m_FilteredMessages
            End Get
            Set(value As ObservableCollection(Of Message))
                Dim emptyDelegate As Action = Sub()
                                                  m_Messages = value
                                                  m_SelectedMessage = 0
                                                  TruncateLimit()
                                                  OnPropertyChanged(String.Empty)
                                              End Sub
                InvokeDelegate(emptyDelegate)
            End Set
        End Property
        Private Sub InvokeDelegate(ByVal emptyDelegate As Action)
            ' invokes delegate on UI dispatcher
            If IsNothing(oMessageDispatcher) OrElse oMessageDispatcher.CheckAccess Then
                emptyDelegate.Invoke
            Else
                Try
                    ' make sure that the UI thread is not blocked with something like a task "waitall", otherwise this will hang
                    oMessageDispatcher.Invoke(emptyDelegate, Threading.DispatcherPriority.Render)
                Catch ex As TaskCanceledException
                End Try
            End If
        End Sub
        Private Sub TruncateLimit()
            ' removes old messages above the message limit
            Dim oCurrentMessages As List(Of Message) = (From oMessage As Message In m_Messages Order By oMessage.DateTime Descending Take MessageLimit Select oMessage).ToList
            m_Messages.Clear()
            For Each oMessage In oCurrentMessages
                m_Messages.Add(oMessage)
            Next
            UpdateFilteredMessages()
        End Sub
        Private Sub UpdateFilteredMessages()
            ' updates filtered messages
            m_FilteredMessages.Clear()
            Select Case m_FilterColour
                Case 0
                    For Each oMessage As Message In m_Messages
                        m_FilteredMessages.Add(oMessage)
                    Next
                Case 1
                    For Each oMessage As Message In m_Messages
                        If oMessage.Colour = Media.Colors.Blue Then
                            m_FilteredMessages.Add(oMessage)
                        End If
                    Next
                Case 2
                    For Each oMessage As Message In m_Messages
                        If oMessage.Colour = Media.Colors.Green Then
                            m_FilteredMessages.Add(oMessage)
                        End If
                    Next
                Case 3
                    For Each oMessage As Message In m_Messages
                        If oMessage.Colour = Media.Colors.Red Then
                            m_FilteredMessages.Add(oMessage)
                        End If
                    Next
                Case 4
                    For Each oMessage As Message In m_Messages
                        If oMessage.Colour = Media.Colors.Yellow Then
                            m_FilteredMessages.Add(oMessage)
                        End If
                    Next
                Case Else
                    For Each oMessage As Message In m_Messages
                        m_FilteredMessages.Add(oMessage)
                    Next
            End Select
        End Sub
        Public Class Message
            ' represents a single message with a source provider, a colour representing the source, a time of creation, and a message content
            Implements INotifyPropertyChanged

            Private m_Source As String
            Private m_Colour As Media.Color
            Private m_DateTime As Date
            Private m_Message As String

            Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
            Protected Sub OnPropertyChanged(ByVal sName As String)
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
            End Sub
            Public Sub New(ByVal sSource As String, oColour As Media.Color, ByVal oDateTime As Date, ByVal sMessage As String)
                m_Source = sSource
                m_Colour = oColour
                m_DateTime = oDateTime
                m_Message = sMessage
                OnPropertyChanged(String.Empty)
            End Sub
            Public Property Source As String
                Get
                    Return m_Source
                End Get
                Set(value As String)
                    m_Source = value
                    OnPropertyChanged("Source")
                End Set
            End Property
            Public Property Colour As Media.Color
                Get
                    Return m_Colour
                End Get
                Set(value As Media.Color)
                    m_Colour = value
                    OnPropertyChanged("Colour")
                    OnPropertyChanged("ColourText")
                End Set
            End Property
            Public ReadOnly Property ColourText As String
                Get
                    Return m_Colour.ToString()
                End Get
            End Property
            Public Property DateTime As Date
                Get
                    Return m_DateTime
                End Get
                Set(value As Date)
                    m_DateTime = value
                    OnPropertyChanged("DateTime")
                    OnPropertyChanged("DateTimeText")
                End Set
            End Property
            Public ReadOnly Property DateTimeText As String
                Get
                    Return m_DateTime.ToShortDateString + " " + m_DateTime.ToLongTimeString
                End Get
            End Property
            Public Property Message As String
                Get
                    Return m_Message
                End Get
                Set(value As String)
                    m_Message = value
                    OnPropertyChanged("Message")
                End Set
            End Property
        End Class
    End Class
    <DataContract> Public Structure ElementStruc
        <DataMember> Public Text As String
        <DataMember> Public ElementType As ElementTypeEnum
        <DataMember> Public Font As Enumerations.FontEnum

        Public Sub New(ByVal sText As String, ByVal oElementType As ElementTypeEnum, ByVal bBold As Boolean, ByVal bItalic As Boolean, ByVal bUnderline As Boolean)
            Text = sText
            ElementType = oElementType
            Font = (If(bBold, 1, 0) * Enumerations.FontEnum.Bold) + (If(bItalic, 1, 0) * Enumerations.FontEnum.Italic) + (If(bUnderline, 1, 0) * Enumerations.FontEnum.Underline)
        End Sub
        Public Sub New(ByVal sText As String, ByVal oElementType As ElementTypeEnum, ByVal oFont As Enumerations.FontEnum)
            Text = sText
            ElementType = oElementType
            Font = oFont
        End Sub
        Public Sub New(ByVal sText As String, ByVal oElementType As ElementTypeEnum)
            Text = sText
            ElementType = oElementType
            Font = Enumerations.FontEnum.None
        End Sub
        Public Sub New(ByVal sText As String)
            Text = sText
            ElementType = ElementTypeEnum.Text
            Font = Enumerations.FontEnum.None
        End Sub
        Public ReadOnly Property FontBold As Boolean
            Get
                If (Font And Enumerations.FontEnum.Bold).Equals(Enumerations.FontEnum.Bold) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property FontItalic As Boolean
            Get
                If (Font And Enumerations.FontEnum.Italic).Equals(Enumerations.FontEnum.Italic) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public ReadOnly Property FontUnderline As Boolean
            Get
                If (Font And Enumerations.FontEnum.Underline).Equals(Enumerations.FontEnum.Underline) Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property
        Public Enum ElementTypeEnum As Integer
            Text = 0
            Subject
            Field
            Template
        End Enum
    End Structure
    <DataContract> Public Class FieldDocumentStore
        Implements ICloneable, INotifyPropertyChanged

        <DataMember> Public WithEvents FieldCollectionStore As New TrueObservableCollection(Of FieldCollection)
        <DataMember> Public PDFTemplate As Byte()
        <DataMember> Public FieldMatrixStore As New MatrixStore
        Private m_FieldContent As New ObservableCollection(Of HighlightComboBox.HCBDisplay)

#Region "INotifyPropertyChanged"
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private Sub OnPropertyChangedLocal(ByVal sName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(sName))
        End Sub
#End Region
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oFieldDocumentStore As New FieldDocumentStore
            With oFieldDocumentStore
                .FieldCollectionStore.Clear()
                For Each oFieldCollection As FieldCollection In FieldCollectionStore
                    .FieldCollectionStore.Add(oFieldCollection.Clone)
                Next

                .PDFTemplate = PDFTemplate.Clone
                .FieldMatrixStore = FieldMatrixStore.Clone
            End With
            Return oFieldDocumentStore
        End Function
        Public Property FieldContent As ObservableCollection(Of HighlightComboBox.HCBDisplay)
            Get
                Return m_FieldContent
            End Get
            Set(value As ObservableCollection(Of HighlightComboBox.HCBDisplay))
            End Set
        End Property
        Public Sub SetMarks()
            ' set marks for all fields
            For Each oFieldCollection In FieldCollectionStore
                For Each oField In oFieldCollection.Fields
                    oField.SetMarks()
                Next
            Next
        End Sub
        Public Sub CleanMatrix()
            ' runs through matrix store and cleans out extra images
            Dim oGUIDStore As List(Of Guid) = (From oFieldCollection In FieldCollectionStore From oField As Field In oFieldCollection.Fields From oImage In oField.Images Select oImage.Item2 Distinct).ToList
            FieldMatrixStore.CleanMatrix(oGUIDStore)
        End Sub
        Private Sub LocalCollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles FieldCollectionStore.CollectionChanged
            If Not IsNothing(e.OldItems) Then
                For Each oFieldCollection As FieldCollection In e.OldItems
                    If m_FieldContent.Contains(oFieldCollection.HCBDisplay) Then
                        m_FieldContent.Remove(oFieldCollection.HCBDisplay)
                    End If
                Next
            End If
            If Not IsNothing(e.NewItems) Then
                For Each oFieldCollection As FieldCollection In e.NewItems
                    m_FieldContent.Add(oFieldCollection.HCBDisplay)
                Next
            End If
            OnPropertyChangedLocal("FieldContent")
        End Sub
        <OnDeserialized()> Private Sub OnDeserialized(Optional ByVal c As StreamingContext = Nothing)
            ' set fieldcontent
            If IsNothing(m_FieldContent) Then
                m_FieldContent = New ObservableCollection(Of HighlightComboBox.HCBDisplay)
            Else
                m_FieldContent.Clear()
            End If
            For Each oFieldCollection As FieldCollection In FieldCollectionStore
                m_FieldContent.Add(oFieldCollection.HCBDisplay)
            Next
            OnPropertyChangedLocal("FieldContent")
        End Sub
        <DataContract> Public Class FieldCollection
            Implements ICloneable

            <DataMember> Private m_GUID As Guid = Guid.NewGuid

            ' collection of input fields from a single form
            <DataMember> Public WithEvents Fields As New TrueObservableCollection(Of Field)

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

            Public Processed As Boolean = False

            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oFields As New FieldCollection
                With oFields
                    .m_GUID = m_GUID
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
                End With
                Return oFields
            End Function
            Private Sub LocalCollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles Fields.CollectionChanged
                If Not IsNothing(e.OldItems) Then
                    For Each oField As Field In e.OldItems
                        oField.Parent = Nothing
                    Next
                End If
                If Not IsNothing(e.NewItems) Then
                    For Each oField As Field In e.NewItems
                        oField.Parent = Me
                    Next
                End If
            End Sub
            <OnDeserialized()> Private Sub OnDeserialized(Optional ByVal c As StreamingContext = Nothing)
                ' set parent
                For Each oField As Field In Fields
                    oField.Parent = Me
                Next
            End Sub
            Public Property GUID As Guid
                Get
                    Return m_GUID
                End Get
                Set(value As Guid)
                    m_GUID = value
                End Set
            End Property
            Public ReadOnly Property HCBDisplay As HighlightComboBox.HCBDisplay
                Get
                    Dim iCount As Integer = Aggregate oField In Fields Where oField.DataPresent = Field.DataPresentEnum.DataPartial Into Count
                    Dim oHCBDisplay As New HighlightComboBox.HCBDisplay(SubjectName, m_GUID, iCount > 0)
                    Return oHCBDisplay
                End Get
            End Property
        End Class
        <DataContract> Public Class Field
            Implements INotifyPropertyChanged, IEditableObject, ICloneable

            ' Data for undoing canceled edits.
            Private temp_Field As Field = Nothing
            Private m_Editing As Boolean = False
            Private Shared m_SelectionChangedCommand As ICommand
            Public Event UpdateEvent()

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
            <DataMember> Private m_GUID As Guid = Guid.Empty

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
            <DataMember> Private m_MarkText As New List(Of Tuple(Of String, String, String))
            Private m_Parent As FieldCollection = Nothing
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

                    .m_MarkText.AddRange(m_MarkText)
                End With
                Return oField
            End Function
            Public Shared Property InteractionEvent() As ICommand
                Get
                    Return m_SelectionChangedCommand
                End Get
                Set(value As ICommand)
                    m_SelectionChangedCommand = value
                End Set
            End Property
#Region "Properties"
            Public Property FieldType As Enumerations.FieldTypeEnum
                Get
                    Return m_FieldType
                End Get
                Set(value As Enumerations.FieldTypeEnum)
                    m_FieldType = value
                End Set
            End Property
            Public Property FieldTypeText As String
                Get
                    Select Case m_FieldType
                        Case Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                            Return "MCQ"
                        Case Else
                            Return m_FieldType.ToString
                    End Select
                End Get
                Set(value As String)
                    If value <> m_FieldType.ToString Then
                        m_FieldType = [Enum].Parse(GetType(Enumerations.FieldTypeEnum), value)
                        OnPropertyChangedLocal("FieldType")
                        OnPropertyChangedLocal("FieldTypeText")
                    End If
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
                    Return (m_PageNumber).ToString
                End Get
                Set(value As String)
                    If value <> (m_PageNumber).ToString Then
                        m_PageNumber = CInt(Val(value))
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
            Public ReadOnly Property FormTitle As String
                Get
                    Return If(IsNothing(m_Parent), String.Empty, m_Parent.FormTitle)
                End Get
            End Property
            Public ReadOnly Property SubjectName As String
                Get
                    Return If(IsNothing(m_Parent), String.Empty, If(m_Parent.AppendText = String.Empty, m_Parent.SubjectName, m_Parent.SubjectName + " (" + m_Parent.AppendText + ")"))
                End Get
            End Property
            Public ReadOnly Property RawBarCodes As List(Of Tuple(Of String, String, String))
                Get
                    Return If(IsNothing(m_Parent), New List(Of Tuple(Of String, String, String)), m_Parent.RawBarCodes)
                End Get
            End Property
            Public Property Parent As FieldCollection
                Get
                    Return m_Parent
                End Get
                Set(value As FieldCollection)
                    m_Parent = value
                End Set
            End Property
#End Region
#Region "Images"
            Public ReadOnly Property ImageCount As Integer
                Get
                    Return m_Images.Count
                End Get
            End Property
            Public ReadOnly Property ImagesPresent As Boolean
                Get
                    Dim bImagesPresent As Boolean = False
                    For Each oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images
                        If (Not IsNothing(oImage)) AndAlso (Not oImage.Item2.Equals(Guid.Empty)) Then
                            bImagesPresent = True
                            Exit For
                        End If
                    Next
                    Return bImagesPresent
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
                Dim oGUIDSource As Guid = Guid.NewGuid
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
                                Return DataPresentEnum.DataNone
                            ElseIf (Aggregate iIndex In Enumerable.Range(0, oBoxIndexList.Count) Where ((Trim(MarkBoxChoice0(iIndex)) <> String.Empty Or Trim(MarkBoxChoice1(iIndex)) <> String.Empty) And Trim(MarkBoxChoice2(iIndex)) = String.Empty) And Not oBoxIndexList(iIndex).Item2 Into Count()) > 0 Then
                                Return DataPresentEnum.DataPartial
                            Else
                                Return DataPresentEnum.DataFull
                            End If
                        Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                            iTotalDMarkCount = Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where MarkChoice0(iIndex) Or MarkChoice1(iIndex) Or MarkChoice2(iIndex) Into Count()
                            If iTotalDMarkCount = 0 Then
                                Return DataPresentEnum.DataNone
                            ElseIf (Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where (MarkChoice0(iIndex) Or MarkChoice1(iIndex)) And (Not MarkChoice2(iIndex)) Into Count()) > 0 Then
                                Return DataPresentEnum.DataPartial
                            Else
                                Return DataPresentEnum.DataFull
                            End If
                        Case Enumerations.FieldTypeEnum.Handwriting
                            iTotalDMarkCount = Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where Trim(MarkHandwriting0(iIndex)) <> String.Empty Or Trim(MarkHandwriting1(iIndex)) <> String.Empty Or Trim(MarkHandwriting2(iIndex)) <> String.Empty And Not m_Images(iIndex).Item6 Into Count()
                            If iTotalDMarkCount = 0 Then
                                Return DataPresentEnum.DataNone
                            ElseIf (Aggregate iIndex In Enumerable.Range(0, m_Images.Count) Where (Trim(MarkHandwriting0(iIndex)) <> String.Empty Or Trim(MarkHandwriting1(iIndex)) <> String.Empty) And Trim(MarkHandwriting2(iIndex)) = String.Empty And Not m_Images(iIndex).Item6 Into Count()) > 0 Then
                                Return DataPresentEnum.DataPartial
                            Else
                                Return DataPresentEnum.DataFull
                            End If
                        Case Enumerations.FieldTypeEnum.Free
                            If Trim(MarkFree0) = String.Empty And Trim(MarkFree1) = String.Empty And Trim(MarkFree2) = String.Empty Then
                                Return DataPresentEnum.DataNone
                            ElseIf (Trim(MarkFree0) <> String.Empty Or Trim(MarkFree1) <> String.Empty) And Trim(MarkFree2) = String.Empty Then
                                Return DataPresentEnum.DataPartial
                            Else
                                Return DataPresentEnum.DataFull
                            End If
                        Case Else
                            Return DataPresentEnum.DataNone
                    End Select
                End Get
                Set(value As DataPresentEnum)
                End Set
            End Property
            Public Property DataPresentText As String
                Get
                    Select Case DataPresent
                        Case DataPresentEnum.DataFull
                            Return DataPresentFull
                        Case DataPresentEnum.DataPartial
                            Return DataPresentPartial
                        Case Else
                            Return DataPresentNone
                    End Select
                End Get
                Set(value As String)
                End Set
            End Property
            Public ReadOnly Property MarkCount As Integer
                Get
                    Return m_MarkText.Count
                End Get
            End Property
            Public Property MarkText As List(Of Tuple(Of String, String, String))
                Get
                    Return m_MarkText
                End Get
                Set(value As List(Of Tuple(Of String, String, String)))
                    m_MarkText = value
                End Set
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
                            Update()
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
                            Update()
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
                            OnPropertyChangedLocal("DataPresentText")
                            OnPropertyChangedLocal("MarkBoxChoice2")
                            Update()
                        End If
                    End If
                End Set
            End Property
            Public Property MarkBoxChoiceRow(ByVal iRow As Integer, ByVal iOrder As Integer) As String
                Get
                    If m_MarkText.Count > 0 Then
                        Dim iActualOrder As Integer = iOrder
                        If iOrder < 0 Then
                            iActualOrder = 0
                        ElseIf iOrder > 2 Then
                            iActualOrder = 2
                        End If

                        Select Case iActualOrder
                            Case 1
                                Return m_MarkText(iRow).Item2
                            Case 2
                                Return m_MarkText(iRow).Item3
                            Case Else
                                Return m_MarkText(iRow).Item1
                        End Select
                    End If
                    Return String.Empty
                End Get
                Set(value As String)
                    If m_MarkText.Count > 0 Then
                        Dim iActualOrder As Integer = iOrder
                        If iOrder < 0 Then
                            iActualOrder = 0
                        ElseIf iOrder > 2 Then
                            iActualOrder = 2
                        End If

                        Select Case iActualOrder
                            Case 1
                                If m_MarkText(iRow).Item2 <> value Then
                                    m_MarkText(iRow) = New Tuple(Of String, String, String)(m_MarkText(iRow).Item1, Left(Trim(value), 1), m_MarkText(iRow).Item3)
                                    SetFinal(iRow)
                                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                                    Update()
                                End If
                            Case 2
                                If m_MarkText(iRow).Item3 <> value Then
                                    m_MarkText(iRow) = New Tuple(Of String, String, String)(m_MarkText(iRow).Item1, m_MarkText(iRow).Item2, Left(Trim(value), 1))
                                    SetFinal(iRow)
                                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                                    Update()
                                End If
                            Case Else
                                If m_MarkText(iRow).Item1 <> value Then
                                    m_MarkText(iRow) = New Tuple(Of String, String, String)(Left(Trim(value), 1), m_MarkText(iRow).Item2, m_MarkText(iRow).Item3)
                                    SetFinal(iRow)
                                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                                    Update()
                                End If
                        End Select
                    End If
                End Set
            End Property
            Public ReadOnly Property MarkBoxChoiceCombined0 As String
                Get
                    Dim sText As String = String.Empty
                    For iIndex = 0 To m_MarkText.Count - 1
                        Dim sChar As String = MarkBoxChoice0(iIndex)
                        sText += If(sChar = String.Empty, " ", sChar)
                    Next
                    Return sText
                End Get
            End Property
            Public ReadOnly Property MarkBoxChoiceCombined1 As String
                Get
                    Dim sText As String = String.Empty
                    For iIndex = 0 To m_MarkText.Count - 1
                        Dim sChar As String = MarkBoxChoice1(iIndex)
                        sText += If(sChar = String.Empty, " ", sChar)
                    Next
                    Return sText
                End Get
            End Property
            Public ReadOnly Property MarkBoxChoiceCombined2 As String
                Get
                    Dim sText As String = String.Empty
                    For iIndex = 0 To m_MarkText.Count - 1
                        Dim sChar As String = MarkBoxChoice2(iIndex)
                        sText += If(sChar = String.Empty, " ", sChar)
                    Next
                    Return sText
                End Get
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
                            If m_TabletSingleChoiceOnly Then
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
                    Update()
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
                            If m_TabletSingleChoiceOnly Then
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
                    Update()
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
                            If m_TabletSingleChoiceOnly Then
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
                    OnPropertyChangedLocal("DataPresentText")
                    OnPropertyChangedLocal(Data.Binding.IndexerName)
                    Update()
                End Set
            End Property
            Public ReadOnly Property MarkChoiceCombined0 As String
                Get
                    Dim sMarkText As String = String.Empty
                    For i = 0 To m_MarkText.Count - 1
                        If Trim(m_MarkText(i).Item1) = String.Empty Then
                            sMarkText += "0"
                        Else
                            sMarkText += "1"
                        End If
                    Next
                    Return sMarkText
                End Get
            End Property
            Public ReadOnly Property MarkChoiceCombined1 As String
                Get
                    Dim sMarkText As String = String.Empty
                    For i = 0 To m_MarkText.Count - 1
                        If Trim(m_MarkText(i).Item2) = String.Empty Then
                            sMarkText += "0"
                        Else
                            sMarkText += "1"
                        End If
                    Next
                    Return sMarkText
                End Get
            End Property
            Public ReadOnly Property MarkChoiceCombined2 As String
                Get
                    Dim sMarkText As String = String.Empty
                    For i = 0 To m_MarkText.Count - 1
                        If Trim(m_MarkText(i).Item3) = String.Empty Then
                            sMarkText += "0"
                        Else
                            sMarkText += "1"
                        End If
                    Next
                    Return sMarkText
                End Get
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
                            Update()
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
                            Update()
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
                            OnPropertyChangedLocal("DataPresentText")
                            OnPropertyChangedLocal(Data.Binding.IndexerName)
                            Update()
                        End If
                    End If
                End Set
            End Property
            Public Property MarkHandwritingRowCol(ByVal iRow As Integer, ByVal iColumn As Integer, ByVal iOrder As Integer) As String
                Get
                    If m_MarkText.Count > 0 Then
                        For iIndex = 0 To m_Images.Count - 1
                            If m_Images(iIndex).Item4 = iRow And m_Images(iIndex).Item5 = iColumn Then
                                Dim iActualOrder As Integer = iOrder
                                If iOrder < 0 Then
                                    iActualOrder = 0
                                ElseIf iOrder > 2 Then
                                    iActualOrder = 2
                                End If

                                Select Case iActualOrder
                                    Case 1
                                        Return m_MarkText(iIndex).Item2
                                    Case 2
                                        Return m_MarkText(iIndex).Item3
                                    Case Else
                                        Return m_MarkText(iIndex).Item1
                                End Select
                            End If
                        Next
                    End If
                    Return String.Empty
                End Get
                Set(value As String)
                    If m_MarkText.Count > 0 Then
                        For iIndex = 0 To m_Images.Count - 1
                            If m_Images(iIndex).Item4 = iRow And m_Images(iIndex).Item5 = iColumn Then
                                Dim iActualOrder As Integer = iOrder
                                If iOrder < 0 Then
                                    iActualOrder = 0
                                ElseIf iOrder > 2 Then
                                    iActualOrder = 2
                                End If

                                Select Case iActualOrder
                                    Case 1
                                        If m_MarkText(iIndex).Item2 <> value Then
                                            m_MarkText(iIndex) = New Tuple(Of String, String, String)(m_MarkText(iIndex).Item1, Left(Trim(value), 1), m_MarkText(iIndex).Item3)
                                            SetFinal(iIndex)
                                            OnPropertyChangedLocal(Data.Binding.IndexerName)
                                            Update()
                                        End If
                                    Case 2
                                        If m_MarkText(iIndex).Item3 <> value Then
                                            m_MarkText(iIndex) = New Tuple(Of String, String, String)(m_MarkText(iIndex).Item1, m_MarkText(iIndex).Item2, Left(Trim(value), 1))
                                            OnPropertyChangedLocal("DataPresent")
                                            OnPropertyChangedLocal("DataPresentText")
                                            OnPropertyChangedLocal(Data.Binding.IndexerName)
                                            Update()
                                        End If
                                    Case Else
                                        If m_MarkText(iIndex).Item1 <> value Then
                                            m_MarkText(iIndex) = New Tuple(Of String, String, String)(Left(Trim(value), 1), m_MarkText(iIndex).Item2, m_MarkText(iIndex).Item3)
                                            SetFinal(iIndex)
                                            OnPropertyChangedLocal(Data.Binding.IndexerName)
                                            Update()
                                        End If
                                End Select
                            End If
                        Next
                    End If
                End Set
            End Property
            Public Property MarkHandwritingRowColText(ByVal sRowCol As String) As String
                Get
                    Dim oRowColArray As String() = sRowCol.Split(".")
                    Return MarkHandwritingRowCol(Val(oRowColArray(0)), Val(oRowColArray(1)), Val(oRowColArray(2)))
                End Get
                Set(value As String)
                    Dim oRowColArray As String() = sRowCol.Split(".")
                    MarkHandwritingRowCol(Val(oRowColArray(0)), Val(oRowColArray(1)), Val(oRowColArray(2))) = value
                End Set
            End Property
            Public ReadOnly Property MarkHandwritingCombined0 As String
                Get
                    Dim sMarkText As String = String.Empty
                    If FieldType = Enumerations.FieldTypeEnum.Handwriting Then
                        Dim iMaxRow As Integer = Aggregate oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images Into Max(oImage.Item4)
                        Dim iMaxCol As Integer = Aggregate oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images Into Max(oImage.Item5)

                        For iRow = 0 To iMaxRow
                            For iCol = 0 To iMaxCol
                                Dim iCurrentRow As Integer = iRow
                                Dim iCurrentCol As Integer = iCol
                                Dim oSelectedTextList As List(Of String) = (From iIndex As Integer In Enumerable.Range(0, m_Images.Count) Where m_Images(iIndex).Item4 = iCurrentRow And m_Images(iIndex).Item5 = iCurrentCol Select m_MarkText(iIndex).Item1).ToList
                                If oSelectedTextList.Count > 0 Then
                                    sMarkText += oSelectedTextList.First
                                End If
                            Next
                        Next
                    End If
                    Return sMarkText
                End Get
            End Property
            Public ReadOnly Property MarkHandwritingCombined1 As String
                Get
                    Dim sMarkText As String = String.Empty
                    If FieldType = Enumerations.FieldTypeEnum.Handwriting Then
                        Dim iMaxRow As Integer = Aggregate oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images Into Max(oImage.Item4)
                        Dim iMaxCol As Integer = Aggregate oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images Into Max(oImage.Item5)

                        For iRow = 0 To iMaxRow
                            For iCol = 0 To iMaxCol
                                Dim iCurrentRow As Integer = iRow
                                Dim iCurrentCol As Integer = iCol
                                Dim oSelectedTextList As List(Of String) = (From iIndex As Integer In Enumerable.Range(0, m_Images.Count) Where m_Images(iIndex).Item4 = iCurrentRow And m_Images(iIndex).Item5 = iCurrentCol Select m_MarkText(iIndex).Item2).ToList
                                If oSelectedTextList.Count > 0 Then
                                    sMarkText += oSelectedTextList.First
                                End If
                            Next
                        Next
                    End If
                    Return sMarkText
                End Get
            End Property
            Public ReadOnly Property MarkHandwritingCombined2 As String
                Get
                    Dim sMarkText As String = String.Empty
                    If FieldType = Enumerations.FieldTypeEnum.Handwriting Then
                        Dim iMaxRow As Integer = Aggregate oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images Into Max(oImage.Item4)
                        Dim iMaxCol As Integer = Aggregate oImage As Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid)) In m_Images Into Max(oImage.Item5)

                        For iRow = 0 To iMaxRow
                            For iCol = 0 To iMaxCol
                                Dim iCurrentRow As Integer = iRow
                                Dim iCurrentCol As Integer = iCol
                                Dim oSelectedTextList As List(Of String) = (From iIndex As Integer In Enumerable.Range(0, m_Images.Count) Where m_Images(iIndex).Item4 = iCurrentRow And m_Images(iIndex).Item5 = iCurrentCol Select m_MarkText(iIndex).Item3).ToList
                                If oSelectedTextList.Count > 0 Then
                                    sMarkText += oSelectedTextList.First
                                End If
                            Next
                        Next
                    End If
                    Return sMarkText
                End Get
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
                        Update()
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
                        Update()
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
                        OnPropertyChangedLocal("DataPresentText")
                        OnPropertyChangedLocal(Data.Binding.IndexerName)
                        Update()
                    End If
                End Set
            End Property
            Public Sub SetMarks()
                ' set marks for field
                Dim oMarkText As List(Of String) = Nothing
                Select Case FieldType
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        oMarkText = (From oImage In Images Where oImage.Item5 = -1 Order By oImage.Item4 Ascending Select oImage.Item3).ToList
                        Dim oBoxImageList As List(Of Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single, Guid))) = (From iIndex As Integer In Enumerable.Range(0, Images.Count) Where Images(iIndex).Item5 = -1 Select Images(iIndex)).ToList
                        m_MarkText = (From iKeyIndex As Integer In Enumerable.Range(0, oBoxImageList.Count) Select New Tuple(Of String, String, String)(If(oMarkText(iKeyIndex) = String.Empty, String.Empty, Right(oMarkText(iKeyIndex), 1)), If(oMarkText(iKeyIndex) = String.Empty, String.Empty, Left(oMarkText(iKeyIndex), 1)), If(oBoxImageList(iKeyIndex).Item5, oMarkText(iKeyIndex), String.Empty))).ToList
                    Case Else
                        oMarkText = (From oImage In Images Order By oImage.Item4, oImage.Item5 Ascending Select oImage.Item3).ToList
                        m_MarkText = (From iIndex As Integer In Enumerable.Range(0, Images.Count) Select New Tuple(Of String, String, String)(oMarkText(iIndex), If(Images(iIndex).Item5, oMarkText(iIndex), String.Empty), If(Images(iIndex).Item5, oMarkText(iIndex), String.Empty))).ToList
                End Select
                Select Case FieldType
                    Case Enumerations.FieldTypeEnum.Choice, Enumerations.FieldTypeEnum.ChoiceVertical, Enumerations.FieldTypeEnum.ChoiceVerticalMCQ
                        For j = 0 To MarkCount - 1
                            If Not Critical Then
                                MarkChoice1(j) = MarkChoice0(j)
                                MarkChoice2(j) = MarkChoice0(j)
                            End If
                        Next
                    Case Enumerations.FieldTypeEnum.Handwriting
                        For j = 0 To MarkCount - 1
                            If Not Critical Then
                                MarkHandwriting1(j) = MarkHandwriting0(j)
                                MarkHandwriting2(j) = MarkHandwriting0(j)
                            End If
                        Next
                    Case Enumerations.FieldTypeEnum.BoxChoice
                        For j = 0 To MarkCount - 1
                            If MarkBoxChoice1(j) = MarkBoxChoice0(j) Then
                                MarkBoxChoice2(j) = MarkBoxChoice0(j)
                            End If
                        Next
                    Case Enumerations.FieldTypeEnum.Free
                    Case Else
                End Select
            End Sub
            Private Sub Update()
                RaiseEvent UpdateEvent()
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
                OnPropertyChangedLocal("DataPresentText")
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
    <CollectionDataContract> Public Class TrueObservableCollection(Of T)
        Inherits ObservableCollection(Of T)

        Public Sub AddRange(ByVal oItems As IEnumerable(Of T))
            For Each oItem In oItems
                Items.Add(oItem)
                Changed(NotifyCollectionChangedAction.Add, oItem)
            Next
        End Sub
        Public Sub Reset(ByVal oItems As IEnumerable(Of T))
            Clear()
            AddRange(oItems)
        End Sub
        Public Shadows Sub Clear()
            ' remove old items
            For i = Items.Count - 1 To 0 Step -1
                Dim oOldItem As T = Items(i)
                Items.RemoveAt(i)
                Changed(NotifyCollectionChangedAction.Remove, oOldItem)
            Next
        End Sub
        Private Sub Changed(ByVal oNotifyCollectionChangedAction As NotifyCollectionChangedAction, oItem As T)
            OnPropertyChanged(New PropertyChangedEventArgs("Count"))
            OnPropertyChanged(New PropertyChangedEventArgs("Item[]"))
            Select Case oNotifyCollectionChangedAction
                Case NotifyCollectionChangedAction.Add
                    OnCollectionChanged(New NotifyCollectionChangedEventArgs(oNotifyCollectionChangedAction, oItem))
                Case NotifyCollectionChangedAction.Remove
                    OnCollectionChanged(New NotifyCollectionChangedEventArgs(oNotifyCollectionChangedAction, oItem))
            End Select
        End Sub
    End Class
    Public Class TextLayoutClass
        Private m_Lines As List(Of List(Of Tuple(Of String, Double, Double)))

        Sub New()
            m_Lines = New List(Of List(Of Tuple(Of String, Double, Double)))
        End Sub
        Public Sub AddWord(ByVal sString As String, ByVal fWidth As Double, ByVal fHeight As Double)
            If m_Lines.Count = 0 Then
                AddLine()
            End If

            Dim oCurrentLine As List(Of Tuple(Of String, Double, Double)) = m_Lines(m_Lines.Count - 1)
            oCurrentLine.Add(New Tuple(Of String, Double, Double)(sString, fWidth, fHeight))
            m_Lines(m_Lines.Count - 1) = oCurrentLine
        End Sub
        Public Sub AddLine()
            m_Lines.Add(New List(Of Tuple(Of String, Double, Double)))
        End Sub
        Public ReadOnly Property Lines As List(Of List(Of Tuple(Of String, Double, Double)))
            Get
                Return m_Lines
            End Get
        End Property
        Public ReadOnly Property LineCount As Integer
            Get
                Return m_Lines.Count
            End Get
        End Property
        Public ReadOnly Property Line(ByVal iLine As Integer) As List(Of Tuple(Of String, Double, Double))
            Get
                If iLine >= 0 And iLine < m_Lines.Count Then
                    Return m_Lines(iLine)
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property WordCount(ByVal iLine As Integer) As Integer
            Get
                If iLine >= 0 And iLine < m_Lines.Count Then
                    Return m_Lines(iLine).Count
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property Word(ByVal iLine As Integer, iWord As Integer) As Tuple(Of String, Double, Double)
            Get
                If iLine >= 0 And iLine < m_Lines.Count Then
                    If iWord >= 0 And iWord < m_Lines(iLine).Count Then
                        Return m_Lines(iLine)(iWord)
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public Structure TextStruc
            Private m_Lines As List(Of List(Of String))

            Sub New(ByVal sString As String)
                m_Lines = New List(Of List(Of String))
                Dim sCurrentString As String = sString

                ' trim off words and lines until original string is empty
                Dim oCurrentLine As New List(Of String)
                Do Until IsNothing(sCurrentString) OrElse sCurrentString.Length = 0
                    Dim iNextSpace As Integer = sCurrentString.IndexOf(" ")
                    Dim iNextCrBreak As Integer = sCurrentString.IndexOf(vbCr)
                    Dim iNextLfBreak As Integer = sCurrentString.IndexOf(vbLf)
                    Dim iNextCrLfSpace As Integer = sCurrentString.IndexOf(vbCrLf)

                    If iNextSpace = -1 Then iNextSpace = Integer.MaxValue
                    If iNextCrBreak = -1 Then iNextCrBreak = Integer.MaxValue
                    If iNextLfBreak = -1 Then iNextLfBreak = Integer.MaxValue
                    If iNextCrLfSpace = -1 Then iNextCrLfSpace = Integer.MaxValue

                    Dim iNext As Integer = Math.Min(Math.Min(iNextSpace, iNextCrBreak), Math.Min(iNextLfBreak, iNextCrLfSpace))
                    If iNext < Integer.MaxValue Then
                        If iNextSpace = iNext Then
                            ' remove word
                            oCurrentLine.Add(Left(sCurrentString, iNext))
                            sCurrentString = Right(sCurrentString, sCurrentString.Length - iNext - 1)
                            If sCurrentString = String.Empty Then
                                m_Lines.Add(oCurrentLine)
                            End If
                        ElseIf iNextCrLfSpace = iNext Then
                            ' remove line with double space
                            oCurrentLine.Add(Left(sCurrentString, iNext))
                            m_Lines.Add(oCurrentLine)
                            oCurrentLine = New List(Of String)
                            sCurrentString = Right(sCurrentString, sCurrentString.Length - iNext - 2)
                        ElseIf iNextCrBreak = iNext Then
                            ' remove line with single space
                            oCurrentLine.Add(Left(sCurrentString, iNext))
                            m_Lines.Add(oCurrentLine)
                            oCurrentLine = New List(Of String)
                            sCurrentString = Right(sCurrentString, sCurrentString.Length - iNext - 1)
                        Else
                            ' remove line with single space
                            oCurrentLine.Add(Left(sCurrentString, iNext))
                            m_Lines.Add(oCurrentLine)
                            oCurrentLine = New List(Of String)
                            sCurrentString = Right(sCurrentString, sCurrentString.Length - iNext - 1)
                        End If
                    Else
                        ' add remainder of string to current line
                        oCurrentLine.Add(sCurrentString)
                        m_Lines.Add(oCurrentLine)
                        sCurrentString = String.Empty
                    End If
                Loop

                ' remove leading and trailing spaces
                For i = 0 To m_Lines.Count - 1
                    ' remove leading spaces
                    Do Until m_Lines(i).Count = 0 OrElse m_Lines(i)(0) <> String.Empty
                        m_Lines(i).RemoveAt(0)
                    Loop

                    ' remove trailing spaces
                    Do Until m_Lines(i).Count = 0 OrElse m_Lines(i)(m_Lines(i).Count - 1) <> String.Empty
                        m_Lines(i).RemoveAt(m_Lines(i).Count - 1)
                    Loop
                Next
            End Sub
            Public ReadOnly Property Lines As List(Of List(Of String))
                Get
                    Return m_Lines
                End Get
            End Property
            Public ReadOnly Property LineCount As Integer
                Get
                    Return m_Lines.Count
                End Get
            End Property
            Public ReadOnly Property Line(ByVal iLine As Integer) As List(Of String)
                Get
                    If iLine >= 0 And iLine < m_Lines.Count Then
                        Return m_Lines(iLine)
                    Else
                        Return Nothing
                    End If
                End Get
            End Property
            Public ReadOnly Property WordCount(ByVal iLine As Integer) As Integer
                Get
                    If iLine >= 0 And iLine < m_Lines.Count Then
                        Return m_Lines(iLine).Count
                    Else
                        Return Nothing
                    End If
                End Get
            End Property
            Public ReadOnly Property Word(ByVal iLine As Integer, iWord As Integer) As String
                Get
                    If iLine >= 0 And iLine < m_Lines.Count Then
                        If iWord >= 0 And iWord < m_Lines(iLine).Count Then
                            Return m_Lines(iLine)(iWord)
                        Else
                            Return Nothing
                        End If
                    Else
                        Return Nothing
                    End If
                End Get
            End Property
        End Structure
    End Class
    Public Class LayoutClass
        Private m_Lines As List(Of List(Of Tuple(Of String, Double, Double, Boolean, ElementStruc)))

        Sub New()
            m_Lines = New List(Of List(Of Tuple(Of String, Double, Double, Boolean, ElementStruc)))
        End Sub
        Public Sub AddWord(ByVal sString As String, ByVal fWidth As Double, ByVal fHeight As Double, ByVal bSelected As Boolean, ByVal oElement As ElementStruc)
            If m_Lines.Count = 0 Then
                AddLine()
            End If

            Dim oCurrentLine As List(Of Tuple(Of String, Double, Double, Boolean, ElementStruc)) = m_Lines(m_Lines.Count - 1)
            oCurrentLine.Add(New Tuple(Of String, Double, Double, Boolean, ElementStruc)(sString, fWidth, fHeight, bSelected, oElement))
            m_Lines(m_Lines.Count - 1) = oCurrentLine
        End Sub
        Public Sub AddLine()
            m_Lines.Add(New List(Of Tuple(Of String, Double, Double, Boolean, ElementStruc)))
        End Sub
        Public ReadOnly Property Lines As List(Of List(Of Tuple(Of String, Double, Double, Boolean, ElementStruc)))
            Get
                Return m_Lines
            End Get
        End Property
        Public ReadOnly Property LineCount As Integer
            Get
                Return m_Lines.Count
            End Get
        End Property
        Public ReadOnly Property Line(ByVal iLine As Integer) As List(Of Tuple(Of String, Double, Double, Boolean, ElementStruc))
            Get
                If iLine >= 0 And iLine < m_Lines.Count Then
                    Return m_Lines(iLine)
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property WordCount(ByVal iLine As Integer) As Integer
            Get
                If iLine >= 0 And iLine < m_Lines.Count Then
                    Return m_Lines(iLine).Count
                Else
                    Return Nothing
                End If
            End Get
        End Property
        Public ReadOnly Property Word(ByVal iLine As Integer, iWord As Integer) As Tuple(Of String, Double, Double, Boolean, ElementStruc)
            Get
                If iLine >= 0 And iLine < m_Lines.Count Then
                    If iWord >= 0 And iWord < m_Lines(iLine).Count Then
                        Return m_Lines(iLine)(iWord)
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End Get
        End Property
    End Class
    Public Structure PointDouble
        Public X As Double
        Public Y As Double

        Sub New(ByVal fX As Double, ByVal fY As Double)
            X = fX
            Y = fY
        End Sub
        Public Function IsZero() As Boolean
            If Me.Equals(PointDouble.Zero) Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Function IsNaN() As Boolean
            If Double.IsNaN(Me.X) Or Double.IsNaN(Me.Y) Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Function IsInfinity() As Boolean
            If Double.IsInfinity(Me.X) Or Double.IsInfinity(Me.Y) Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Function Add(ByVal oPointDouble As PointDouble) As PointDouble
            Return New PointDouble(X + oPointDouble.X, Y + oPointDouble.Y)
        End Function
        Public Function Subtract(ByVal oPointDouble As PointDouble) As PointDouble
            Return New PointDouble(X - oPointDouble.X, Y - oPointDouble.Y)
        End Function
        Public Function Multiply(ByVal fFactor As Single) As PointDouble
            Return New PointDouble(X * fFactor, Y * fFactor)
        End Function
        Public Function Divide(ByVal fFactor As Single) As PointDouble
            If fFactor = 0 Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(X / fFactor, Y / fFactor)
            End If
        End Function
        Public Function Multiply(ByVal oPointDouble As PointDouble) As PointDouble
            Return New PointDouble(X * oPointDouble.X, Y * oPointDouble.Y)
        End Function
        Public Function Divide(ByVal oPointDouble As PointDouble) As PointDouble
            If oPointDouble.X = 0 Or oPointDouble.Y = 0 Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(X / oPointDouble.X, Y / oPointDouble.Y)
            End If
        End Function
        Public Function Invert() As PointDouble
            If Me.IsNaN Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(-X, -Y)
            End If
        End Function
        Public Function Reciprocal() As PointDouble
            If X = 0 Or Y = 0 Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(1 / X, 1 / Y)
            End If
        End Function
        Public Function DistanceTo(oPointDouble As PointDouble) As Double
            Return Math.Sqrt(((X - oPointDouble.X) * (X - oPointDouble.X)) + ((Y - oPointDouble.Y) * (Y - oPointDouble.Y)))
        End Function
        Public Function AngleTo(oPointDouble As PointDouble) As Double
            Return Math.Atan2(oPointDouble.Y - Y, oPointDouble.X - X)
        End Function
        Public Shared Function Max(ByVal oPointDouble1 As PointDouble, ByVal oPointDouble2 As PointDouble) As PointDouble
            Dim oReturnX As Double = 0
            If Double.IsNaN(oPointDouble1.X) And Double.IsNaN(oPointDouble2.X) Then
                oReturnX = Double.NaN
            ElseIf Double.IsNaN(oPointDouble1.X) Then
                oReturnX = oPointDouble2.X
            ElseIf Double.IsNaN(oPointDouble2.X) Then
                oReturnX = oPointDouble1.X
            Else
                oReturnX = Math.Max(oPointDouble1.X, oPointDouble2.X)
            End If

            Dim oReturnY As Double = 0
            If Double.IsNaN(oPointDouble1.Y) And Double.IsNaN(oPointDouble2.Y) Then
                oReturnY = Double.NaN
            ElseIf Double.IsNaN(oPointDouble1.Y) Then
                oReturnY = oPointDouble2.Y
            ElseIf Double.IsNaN(oPointDouble2.Y) Then
                oReturnY = oPointDouble1.Y
            Else
                oReturnY = Math.Max(oPointDouble1.Y, oPointDouble2.Y)
            End If

            If oReturnX = Double.NaN Or oReturnY = Double.NaN Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(oReturnX, oReturnY)
            End If
        End Function
        Public Shared Function Min(ByVal oPointDouble1 As PointDouble, ByVal oPointDouble2 As PointDouble) As PointDouble
            Dim oReturnX As Double = 0
            If Double.IsNaN(oPointDouble1.X) And Double.IsNaN(oPointDouble2.X) Then
                oReturnX = Double.NaN
            ElseIf Double.IsNaN(oPointDouble1.X) Then
                oReturnX = oPointDouble2.X
            ElseIf Double.IsNaN(oPointDouble2.X) Then
                oReturnX = oPointDouble1.X
            Else
                oReturnX = Math.Min(oPointDouble1.X, oPointDouble2.X)
            End If

            Dim oReturnY As Double = 0
            If Double.IsNaN(oPointDouble1.Y) And Double.IsNaN(oPointDouble2.Y) Then
                oReturnY = Double.NaN
            ElseIf Double.IsNaN(oPointDouble1.Y) Then
                oReturnY = oPointDouble2.Y
            ElseIf Double.IsNaN(oPointDouble2.Y) Then
                oReturnY = oPointDouble1.Y
            Else
                oReturnY = Math.Min(oPointDouble1.Y, oPointDouble2.Y)
            End If

            If oReturnX = Double.NaN Or oReturnY = Double.NaN Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(oReturnX, oReturnY)
            End If
        End Function
        Public Shared Function Zero() As PointDouble
            Return New PointDouble(0, 0)
        End Function
        Public Shared Function NaN() As PointDouble
            Return New PointDouble(Double.NaN, Double.NaN)
        End Function
        Public Function Abs() As PointDouble
            If Me.IsNaN Then
                Return PointDouble.NaN
            Else
                Return New PointDouble(Math.Abs(X), Math.Abs(Y))
            End If
        End Function
        Public Shared Function Sign(ByVal oPointDouble As PointDouble) As PointDouble
            If oPointDouble.IsNaN Then
                Return PointDouble.Zero
            Else
                Return New PointDouble(Math.Sign(oPointDouble.X), Math.Sign(oPointDouble.Y))
            End If
        End Function
        Public Overloads Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse Not Me.GetType() Is obj.GetType() Then
                Return False
            End If

            Dim p As PointDouble = CType(obj, PointDouble)
            Return Me.X = p.X And Me.Y = p.Y
        End Function
        Public Overrides Function GetHashCode() As Integer
            If Double.IsNaN(X) Or Double.IsNaN(Y) Then
                Return Integer.MinValue
            Else
                Return X Xor Y
            End If
        End Function
    End Structure
    Public Class PointList
        Inherits List(Of SinglePoint)
        Implements ICloneable, IDisposable

        Private Shared CircleAngle As Double = 2 * Math.PI

        Sub New()
        End Sub
        Sub New(ByVal oPointList As List(Of SinglePoint))
            AddRange(oPointList)
        End Sub
#Region "Shared Functions"
        Public Shared Function GetDistance(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Double
            ' returns the distance between the two points
            Return oPoint1.AForgePoint.DistanceTo(oPoint2.AForgePoint)
        End Function
        Public Shared Function InRange(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Boolean
            ' checks to see if the distance between the two points is less than the sum of their radii
            Return GetDistance(oPoint1, oPoint2) < oPoint1.Radius + oPoint2.Radius
        End Function
        Public Shared Function InCircle(ByVal oPoint1 As SinglePoint, ByVal oPoint2 As SinglePoint) As Boolean
            ' checks to see if the distance between the two points is less than the larger circle's radius
            Return GetDistance(oPoint1, oPoint2) < Math.Max(oPoint1.Radius, oPoint2.Radius)
        End Function
        Public Shared Function GetNeighbouringPoints(ByRef oPoint As SinglePoint, ByVal oPointCircleList As PointList, Optional ByVal bAscending As Boolean? = Nothing, Optional ByVal fLowerLimitMaxRadius As Single? = Nothing) As PointList
            ' gets a list of neighbouring points within range from the reference point
            ' fLowerLimitMaxRadius gives a lower limit to the max radius and helps small circles to join up
            Dim oNeighbouringPointsList As New PointList
            For Each oTestPoint As SinglePoint In oPointCircleList
                Dim fMinRadius As Single = Math.Min(oPoint.Radius, oTestPoint.Radius)
                Dim fDistance As Single = GetDistance(oTestPoint, oPoint)
                If fLowerLimitMaxRadius.HasValue Then
                    If fDistance > fMinRadius * 0.1 And (InRange(oPoint, oTestPoint) Or fDistance < fLowerLimitMaxRadius) Then
                        oNeighbouringPointsList.Add(oTestPoint.Clone)
                    End If
                Else
                    If fDistance > fMinRadius * 0.1 And InRange(oPoint, oTestPoint) Then
                        oNeighbouringPointsList.Add(oTestPoint.Clone)
                    End If
                End If
            Next

            ' reorder by distance
            If bAscending.HasValue Then
                If bAscending Then
                    oNeighbouringPointsList = (From oSinglePoint In oNeighbouringPointsList Order By oSinglePoint.Distance Ascending Select oSinglePoint).ToList
                Else
                    oNeighbouringPointsList = (From oSinglePoint In oNeighbouringPointsList Order By oSinglePoint.Distance Descending Select oSinglePoint).ToList
                End If
            End If

            Return oNeighbouringPointsList
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
        Public Shared Function CompareAngleValue(ByVal fAngle As Double, ByVal fMatchAngle As Double) As Double
            ' compares the two angles and returns the difference
            Return Math.Abs(fMatchAngle - MatchAngle(fAngle, fMatchAngle))
        End Function
#End Region
#Region "Functions"
        Public Function BoundingBox() As System.Drawing.Rectangle
            Dim oRectangle As Rect = Rect.Empty
            For Each oPoint In Me
                Dim oPointRectangle = New Rect(oPoint.Point.X - oPoint.Radius, oPoint.Point.Y - oPoint.Radius, oPoint.Radius * 2, oPoint.Radius * 2)
                If oRectangle.IsEmpty Then
                    oRectangle = oPointRectangle
                Else
                    oRectangle.Union(oPointRectangle)
                End If
            Next
            Return New System.Drawing.Rectangle(oRectangle.X, oRectangle.Y, oRectangle.Width, oRectangle.Height)
        End Function
#End Region
        Structure SinglePoint
            Implements ICloneable

            Private m_Radius As Single
            Private m_Point As AForge.Point
            Private m_Count As Integer
            Private m_Distance As Single
            Private m_GUID As Guid
            Private m_Selected As Boolean

            Sub New(ByVal fRadius As Single, ByVal oPoint As System.Drawing.Point, ByVal iCount As Integer, ByVal fDistance As Single)
                m_Radius = fRadius
                m_Point = New AForge.Point(oPoint.X, oPoint.Y)
                m_Count = iCount
                m_Distance = fDistance
                m_GUID = Guid.NewGuid
                m_Selected = False
            End Sub
            Sub New(ByVal fRadius As Single, ByVal oPoint As System.Drawing.Point, ByVal iCount As Integer, ByVal fDistance As Single, ByVal oGUID As Guid)
                m_Radius = fRadius
                m_Point = New AForge.Point(oPoint.X, oPoint.Y)
                m_Count = iCount
                m_Distance = fDistance
                m_GUID = oGUID
                m_Selected = False
            End Sub
            Sub New(ByVal fRadius As Single, ByVal oPoint As AForge.Point, ByVal bSelected As Boolean)
                m_Radius = fRadius
                m_Point = oPoint
                m_Selected = bSelected
                m_GUID = Guid.NewGuid
                m_Count = 0
                m_Distance = 0
            End Sub
            Public ReadOnly Property Radius As Single
                Get
                    Return m_Radius
                End Get
            End Property
            Public ReadOnly Property Point As System.Drawing.Point
                Get
                    Return New System.Drawing.Point(m_Point.X, m_Point.Y)
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
            Public ReadOnly Property Selected As Boolean
                Get
                    Return m_Selected
                End Get
            End Property
            Public ReadOnly Property AForgePoint As AForge.Point
                Get
                    Return New AForge.Point(m_Point.X, m_Point.Y)
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
                Return m_GUID.GetHashCode
            End Function
            Public Function Clone() As Object Implements ICloneable.Clone
                Dim oSinglePoint As New SinglePoint
                With oSinglePoint
                    .m_Radius = m_Radius
                    .m_Point = m_Point
                    .m_Count = m_Count
                    .m_Distance = m_Distance
                    .m_GUID = m_GUID
                    .m_Selected = m_Selected
                End With
                Return oSinglePoint
            End Function
        End Structure
        Structure Segment
            Implements ICloneable

            Public Segment As AForge.Math.Geometry.LineSegment

            Sub New(ByVal oSegment As AForge.Math.Geometry.LineSegment)
                Segment = oSegment
            End Sub
            Sub New(ByVal oStart As SinglePoint, ByVal oEnd As SinglePoint)
                Segment = New AForge.Math.Geometry.LineSegment(oStart.AForgePoint, oEnd.AForgePoint)
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
            Public Overrides Function Equals(obj As Object) As Boolean
                ' Check for null values and compare run-time types.
                If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                    Return False
                End If

                ' reverses the segment to check equality
                Dim oSegment As Segment = obj
                Return GetHashCode.Equals(oSegment.GetHashCode) Or (New AForge.Math.Geometry.LineSegment(Segment.End, Segment.Start)).GetHashCode.Equals(oSegment.GetHashCode)
            End Function
            Public Overrides Function GetHashCode() As Integer
                Return Segment.GetHashCode
            End Function
            Public Function Clone() As Object Implements ICloneable.Clone
                Return New Segment(New AForge.Math.Geometry.LineSegment(Segment.Start, Segment.End))
            End Function
        End Structure
        Public Overrides Function Equals(obj As Object) As Boolean
            ' Check for null values and compare run-time types.
            If obj Is Nothing OrElse [GetType]() <> obj.[GetType]() Then
                Return False
            End If

            Dim oPointList As PointList = obj
            Return GetHashCode.Equals(oPointList.GetHashCode)
        End Function
        Public Overrides Function GetHashCode() As Integer
            Dim iHashCode As Long = Aggregate oPoint As SinglePoint In Me Into Sum(CLng(oPoint.GetHashCode))
            Return iHashCode Mod Integer.MaxValue
        End Function
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim oPointList As New PointList
            For Each oSinglePoint In Me
                oPointList.Add(oSinglePoint.Clone)
            Next
            Return oPointList
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
    End Class
    Public Class CommonScanner
        Implements IDisposable

        Private Const ProcessName As String = "Twain32"
        Private Const ProcessSubDirectory As String = ""

        Private oStreamString As StreamString = Nothing
        Private oSenderPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
        Private oReceiverPipe As IO.Pipes.AnonymousPipeServerStream = Nothing
        Private DoLoop As Boolean = True

        Public Event ReturnScannedImage(ByVal oReturnMessage As Tuple(Of Twain32Enumerations.ScanProgress, Object))

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

            Dim oProcessStartInfo As New ProcessStartInfo(sFullName, senderID + " " + receiverID)
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
                    Close()
                Else
                    oSenderPipe.DisposeLocalCopyOfClientHandle()
                    oReceiverPipe.DisposeLocalCopyOfClientHandle()

                    oStreamString = New StreamString(oSenderPipe, oReceiverPipe)
                    oStreamString.WriteString(Twain32Constants.ValidateString)
                    If oStreamString.ReadString() <> Twain32Constants.Twain32GUID.ToString Then
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
            oSenderPipe = Nothing
            oReceiverPipe = Nothing
        End Sub
        Private Function SendMessage(Of T, U)(ByVal sFunctionName As String, ByVal oInputData As T) As U
            ' sends a message with input data type T and returns with data type U
            ' order of events
            ' 1) send function name
            ' 2) send data as base64 string
            ' 3) acknowledge
            ' 4) receive data as base64 string
            Dim oReturn As U = Nothing
            Dim oUnicodeEncoding As New Text.UnicodeEncoding

            If Not IsNothing(oStreamString) Then
                ' 1) send 'message string'
                oStreamString.WriteString(sFunctionName)
                Dim sMessage As String = oStreamString.ReadString
                If sMessage = Twain32Constants.OKString Then
                    ' 2) send data as base64 string
                    Dim sDataString As String = CommonFunctions.SerializeDataContractText(oInputData, Twain32Functions.GetTwainKnownTypes, False)
                    oStreamString.WriteString(sDataString)
                    sMessage = oStreamString.ReadString
                    If sMessage = Twain32Constants.OKString Then

                        ' 3) acknowledge
                        oStreamString.WriteString(Twain32Constants.OKString)

                        ' 4) receive data as base64 string
                        sMessage = oStreamString.ReadString

                        oReturn = CommonFunctions.DeserializeDataContractText(Of U)(sMessage, Twain32Functions.GetTwainKnownTypes, False)
                    End If
                End If
            End If
            Return oReturn
        End Function
        Public Function InitTwain() As List(Of String)
            ' starts twain server and returns list of scanner sources
            Dim oScannerSources As List(Of String) = SendMessage(Of String, List(Of String))(Twain32Constants.InitString, Twain32Constants.XString)
            Return oScannerSources
        End Function
        Public Sub SelectScannerSource(ByVal sScanner As String)
            ' selects scanner source
            ' if the supplied name is empty, then deselect scanner source
            SendMessage(Of String, String)(Twain32Constants.SelectScannerString, sScanner)
        End Sub
        Public Function GetTWAINScannerSources() As List(Of String)
            ' get scanner sources
            Return SendMessage(Of String, List(Of String))(Twain32Constants.GetScannerSourcesString, Twain32Constants.XString)
        End Function
        Public Sub ScannerScan()
            ' starts scan process
            Dim oStatus As Twain32Enumerations.ScanProgress = SendMessage(Of String, Twain32Enumerations.ScanProgress)(Twain32Constants.StartScanString, Twain32Constants.XString)
            If oStatus = Twain32Enumerations.ScanProgress.NoError Then
                ' fetch images
                Dim bSuccess As Boolean = True
                Do
                    Dim oReturnMessage As Tuple(Of Twain32Enumerations.ScanProgress, Object) = SendMessage(Of String, Tuple(Of Twain32Enumerations.ScanProgress, Object))(Twain32Constants.GetScanImageString, Twain32Constants.XString)
                    If Not IsNothing(oReturnMessage) Then
                        Select Case oReturnMessage.Item1
                            Case Twain32Enumerations.ScanProgress.Image
                                If (Not IsNothing(oReturnMessage.Item2)) AndAlso oReturnMessage.Item2.GetType.Equals(GetType(System.Drawing.Bitmap)) Then
                                    bSuccess = True
                                    RaiseEvent ReturnScannedImage(oReturnMessage)
                                Else
                                    bSuccess = False
                                    RaiseEvent ReturnScannedImage(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.ScanError, Nothing))
                                End If
                            Case Twain32Enumerations.ScanProgress.ScanError
                                bSuccess = False
                                RaiseEvent ReturnScannedImage(oReturnMessage)
                            Case Twain32Enumerations.ScanProgress.Complete
                                bSuccess = False
                                RaiseEvent ReturnScannedImage(oReturnMessage)
                            Case Twain32Enumerations.ScanProgress.None
                                ' no image yet, but in progress
                                bSuccess = True
                        End Select
                    End If

                    System.Threading.Thread.Sleep(100)
                Loop While bSuccess
            End If
        End Sub
        Public Sub ScannerConfigure()
            ' configure selected scanner
            SendMessage(Of String, String)(Twain32Constants.ConfigureString, Twain32Constants.XString)
        End Sub
        Public Function ProcessBarcode(ByVal oBarcodeBitmap As System.Drawing.Bitmap) As String
            ' processes the barcode bitmap using the ZBar library
            Dim sReturnMessage As String = SendMessage(Of System.Drawing.Bitmap, String)(Twain32Constants.ProcessBarcodeString, oBarcodeBitmap)
            If sReturnMessage = Twain32Constants.XString Then
                Return String.Empty
            Else
                Return sReturnMessage
            End If
        End Function
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
    Public Class DetectorFunctions
        Public Shared Function LoadMulticlassSupportVectorMachine(Of TKernel As Accord.Statistics.Kernels.IKernel(Of Double()))(ByVal sFilePath As String) As Accord.MachineLearning.VectorMachines.MulticlassSupportVectorMachine(Of TKernel)
            ' load support vector machine from compressed file
            Using oFileStream As IO.FileStream = IO.File.OpenRead(CommonFunctions.ReplaceExtension(sFilePath, "gz"))
                Using oGZipStream As New IO.Compression.GZipStream(oFileStream, IO.Compression.CompressionMode.Decompress)
                    Return Accord.IO.Serializer.Load(Of Accord.MachineLearning.VectorMachines.MulticlassSupportVectorMachine(Of TKernel))(oGZipStream)
                End Using
            End Using
        End Function
        Public Shared Function LoadMulticlassSupportVectorMachine(ByVal sFilePath As String, ByVal oDetector As Detectors) As Object
            ' load support vector machine from compressed file
            Dim oKernelType As KernelType = GetSVMType(oDetector)
            Select Case oKernelType
                Case KernelType.Linear
                    Return LoadMulticlassSupportVectorMachine(Of Accord.Statistics.Kernels.Linear)(sFilePath)
                Case KernelType.Polynomial2, KernelType.Polynomial5
                    Return LoadMulticlassSupportVectorMachine(Of Accord.Statistics.Kernels.Polynomial)(sFilePath)
                Case KernelType.Gaussian1, KernelType.Gaussian2, KernelType.Gaussian3
                    Return LoadMulticlassSupportVectorMachine(Of Accord.Statistics.Kernels.Gaussian)(sFilePath)
                Case KernelType.Sigmoid
                    Return LoadMulticlassSupportVectorMachine(Of Accord.Statistics.Kernels.Sigmoid)(sFilePath)
                Case KernelType.HistogramIntersection
                    Return LoadMulticlassSupportVectorMachine(Of Accord.Statistics.Kernels.HistogramIntersection)(sFilePath)
                Case Else
                    Return Nothing
            End Select
        End Function
        Public Shared Function LoadDeepBeliefNetwork(ByVal sFilePath As String) As Accord.Neuro.Networks.DeepBeliefNetwork
            ' load deep belief network from compressed file
            Using oFileStream As IO.FileStream = IO.File.OpenRead(CommonFunctions.ReplaceExtension(sFilePath, "gz"))
                Using oGZipStream As New IO.Compression.GZipStream(oFileStream, IO.Compression.CompressionMode.Decompress)
                    Return Accord.Neuro.Networks.DeepBeliefNetwork.Load(oGZipStream)
                End Using
            End Using
        End Function
        Public Shared Function ProcessMat(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal iSize As Integer, ByVal bTrimBox As Boolean) As Emgu.CV.Matrix(Of Byte)
            ' converts matrices to features
            Using oContourMatrix As Emgu.CV.Matrix(Of Byte) = oMatrix.SubR(Byte.MaxValue)
                Emgu.CV.CvInvoke.Normalize(oContourMatrix, oContourMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                Dim oOutputMatrix As New Emgu.CV.Matrix(Of Byte)(iSize, iSize)
                If bTrimBox Then
                    Dim oSubMatrix As Emgu.CV.Matrix(Of Byte) = GetSubMatrix(oContourMatrix, iSize)
                    Emgu.CV.CvInvoke.Resize(oSubMatrix, oOutputMatrix, oOutputMatrix.Size,,, Emgu.CV.CvEnum.Inter.Linear)
                Else
                    Emgu.CV.CvInvoke.Resize(oContourMatrix, oOutputMatrix, oOutputMatrix.Size,,, Emgu.CV.CvEnum.Inter.Linear)
                End If
                Return oOutputMatrix
            End Using
        End Function
        Public Shared Function ProcessSegment(ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal oSegmentList As Tuple(Of List(Of PointList.Segment), List(Of PointList.SinglePoint))) As Double()
            ' scans the segments and process the strokes found to give descriptors
            ' the descriptor is a histogram consisting of segment lengths and angles
            ' the angle is limited to 0-PI radians as the segments are reversible
            ' the histogram is normalised by dividing with the total stroke length captured in that layer (differs from layer to layer due to the overlap margins)
            Const HistogramBins As Integer = 20
            Dim fDiagonalLength As Double = Math.Sqrt(iWidth * iWidth + iHeight * iHeight)
            Dim oRegionDictionary As New Dictionary(Of String, System.Drawing.RectangleF)

            ' add segment array
            oRegionDictionary.Add("TL", New System.Drawing.RectangleF(0, 0, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("T", New System.Drawing.RectangleF(iWidth / 4, 0, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("TR", New System.Drawing.RectangleF(iWidth / 2, 0, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("L", New System.Drawing.RectangleF(0, iHeight / 4, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("C", New System.Drawing.RectangleF(iWidth / 4, iHeight / 4, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("R", New System.Drawing.RectangleF(iWidth / 2, iHeight / 4, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("BL", New System.Drawing.RectangleF(0, iHeight / 2, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("B", New System.Drawing.RectangleF(iWidth / 4, iHeight / 2, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("BR", New System.Drawing.RectangleF(iWidth / 2, iHeight / 2, iWidth / 2, iHeight / 2))
            oRegionDictionary.Add("W", New System.Drawing.RectangleF(0, 0, iWidth, iHeight))

            ' build up a histogram dictionary
            Dim oHistogramDictionary As New Dictionary(Of String, Double())
            For Each sKey As String In oRegionDictionary.Keys
                oHistogramDictionary.Add(sKey, Enumerable.Repeat(Of Double)(0, HistogramBins).ToArray)
            Next

            ' perform segment length extraction only if data is present
            If (Aggregate oSegment As PointList.Segment In oSegmentList.Item1 Into Sum(oSegment.Length)) > 0 Or oSegmentList.Item2.Count > 0 Then
                For Each sKey As String In oRegionDictionary.Keys
                    ' for each segment
                    For Each oSegment As PointList.Segment In oSegmentList.Item1
                        Dim fUnitLength As Double = oSegment.Length * LineInRectangle(oRegionDictionary(sKey), oSegment.Segment)
                        If fUnitLength > 0 Then
                            Dim iBin As Integer = Math.Min(Math.Max(Math.Truncate(oSegment.Angle * HistogramBins / Math.PI), 0), HistogramBins - 1)
                            oHistogramDictionary(sKey)(iBin) += (fUnitLength / fDiagonalLength)
                        End If
                    Next

                    ' for each unselected point, distribute the segmented length (ie. radius / number of bins) to each bin
                    For Each oPoint As PointList.SinglePoint In oSegmentList.Item2
                        If oRegionDictionary(sKey).Contains(New System.Drawing.PointF(oPoint.Point.X, oPoint.Point.Y)) Then
                            If oPoint.Radius > 0 Then
                                Dim fUnitLength As Double = oPoint.Radius / HistogramBins
                                For iBin = 0 To HistogramBins - 1
                                    oHistogramDictionary(sKey)(iBin) += (fUnitLength / fDiagonalLength)
                                Next
                            End If
                        End If
                    Next
                Next
            End If

            Dim oDoublesList As New List(Of Double)
            For Each sKey As String In oRegionDictionary.Keys
                oDoublesList.AddRange(oHistogramDictionary(sKey))
            Next

            Return oDoublesList.ToArray
        End Function
        Private Shared Function LineInRectangle(ByVal oRect As System.Drawing.RectangleF, ByVal oLineSegment As AForge.Math.Geometry.LineSegment) As Double
            ' returns the fraction of the line segment within the bounds of the rectangle
            ' if the line is completely out of the rectangle, then returns zero
            Dim fReturn As Double = 0
            Dim oRectLines As New List(Of AForge.Math.Geometry.LineSegment) From {New AForge.Math.Geometry.LineSegment(New AForge.Point(oRect.X, oRect.Y), New AForge.Point(oRect.X + oRect.Width - 1, oRect.Y)), New AForge.Math.Geometry.LineSegment(New AForge.Point(oRect.X + oRect.Width - 1, oRect.Y), New AForge.Point(oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1)), New AForge.Math.Geometry.LineSegment(New AForge.Point(oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1), New AForge.Point(oRect.X, oRect.Y + oRect.Height - 1)), New AForge.Math.Geometry.LineSegment(New AForge.Point(oRect.X, oRect.Y + oRect.Height - 1), New AForge.Point(oRect.X, oRect.Y))}
            Dim oIntersections As New List(Of AForge.Point)
            For Each oCurrentLineSegment In oRectLines
                Try
                    Dim oIntersectionPoint As AForge.Point? = oLineSegment.GetIntersectionWith(oCurrentLineSegment)
                    If Not IsNothing(oIntersectionPoint) Then
                        oIntersections.Add(oIntersectionPoint)
                    End If
                Catch ex As System.InvalidOperationException
                    If ex.Message = "Overlapping segments do not have a single intersection point." Then
                        ' on the border of the rectangle
                        ' do not count
                        Return 0
                    End If
                End Try
            Next

            ' check for intersections
            Select Case oIntersections.Count
                Case 2
                    If oIntersections(0).Equals(oIntersections(1)) Then
                        ' line cuts at one corner
                        fReturn = 0
                    Else
                        ' line cuts across the rectangle
                        fReturn = (New AForge.Math.Geometry.LineSegment(oIntersections(0), oIntersections(1))).Length / oLineSegment.Length
                    End If
                Case 0
                    Dim oRectContour As New Emgu.CV.Util.VectorOfPointF({New System.Drawing.PointF(oRect.X, oRect.Y), New System.Drawing.PointF(oRect.X + oRect.Width - 1, oRect.Y), New System.Drawing.PointF(oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1), New System.Drawing.PointF(oRect.X, oRect.Y + oRect.Height - 1)})
                    If Emgu.CV.CvInvoke.PointPolygonTest(oRectContour, New System.Drawing.PointF(oLineSegment.Start.X, oLineSegment.Start.Y), False) <= 0 Then
                        ' outside rectangle
                        fReturn = 0
                    Else
                        ' inside rectangle
                        fReturn = 1
                    End If
                Case Else
                    ' part of the line is inside the rectangle
                    Dim oRectContour As New Emgu.CV.Util.VectorOfPointF({New System.Drawing.PointF(oRect.X, oRect.Y), New System.Drawing.PointF(oRect.X + oRect.Width - 1, oRect.Y), New System.Drawing.PointF(oRect.X + oRect.Width - 1, oRect.Y + oRect.Height - 1), New System.Drawing.PointF(oRect.X, oRect.Y + oRect.Height - 1)})
                    If Emgu.CV.CvInvoke.PointPolygonTest(oRectContour, New System.Drawing.PointF(oLineSegment.Start.X, oLineSegment.Start.Y), False) <= 0 Then
                        ' start point is outside rectangle
                        If oLineSegment.End.Equals(oIntersections(0)) Then
                            fReturn = 0
                        Else
                            fReturn = (New AForge.Math.Geometry.LineSegment(oLineSegment.End, oIntersections(0))).Length / oLineSegment.Length
                        End If
                    Else
                        ' start point is inside rectangle
                        If oLineSegment.Start.Equals(oIntersections(0)) Then
                            fReturn = 0
                        Else
                            fReturn = (New AForge.Math.Geometry.LineSegment(oLineSegment.Start, oIntersections(0))).Length / oLineSegment.Length
                        End If
                    End If
            End Select
            Return fReturn
        End Function
        Private Shared Function GetSubMatrix(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal iSize As Integer) As Emgu.CV.Matrix(Of Byte)
            ' gets submatrix of original matrix cropped to the content
            Dim oBoundingRectangle As System.Drawing.Rectangle = GetBoundingRectangle(oMatrix)

            ' resize to box
            ' if the set the max dimension for an axis to be even/odd if the width/height is even/odd
            ' this is to ensure that the cut rectangle aligns to integral pixel borders
            Dim iMaxDimension As Integer = Math.Max(oBoundingRectangle.Width, oBoundingRectangle.Height)
            Dim iMaxX As Integer = If(oBoundingRectangle.Width Mod 2 = 0, If(iMaxDimension Mod 2 = 0, iMaxDimension, iMaxDimension + 1), If(iMaxDimension Mod 2 = 1, iMaxDimension, iMaxDimension + 1))
            Dim iMaxY As Integer = If(oBoundingRectangle.Height Mod 2 = 0, If(iMaxDimension Mod 2 = 0, iMaxDimension, iMaxDimension + 1), If(iMaxDimension Mod 2 = 1, iMaxDimension, iMaxDimension + 1))
            Dim iMaxXY As Integer = Math.Max(iMaxX, iMaxY)
            Dim oBoundingCenter As New System.Drawing.PointF(oBoundingRectangle.X + oBoundingRectangle.Width / 2, oBoundingRectangle.Y + oBoundingRectangle.Height / 2)
            Dim oSubMatrix As New Emgu.CV.Matrix(Of Byte)(iMaxXY, iMaxXY)
            oSubMatrix.SetZero()

            Dim iLeftTrim As Integer = oBoundingCenter.X - iMaxX / 2
            Dim iTopTrim As Integer = oBoundingCenter.Y - iMaxY / 2
            Dim iRightTrim As Integer = oBoundingCenter.X + iMaxX / 2 - 1
            Dim iBottomTrim As Integer = oBoundingCenter.Y + iMaxY / 2 - 1

            Dim iLeft As Integer = Math.Max(iLeftTrim, 0)
            Dim iTop As Integer = Math.Max(iTopTrim, 0)
            Dim iRight As Integer = Math.Min(iRightTrim, oMatrix.Width - 1)
            Dim iBottom As Integer = Math.Min(iBottomTrim, oMatrix.Height - 1)

            Dim iCenterDisplacementX As Single = ((iLeft - iLeftTrim) + (iRight - iRightTrim)) / 2
            Dim iCenterDisplacementY As Single = ((iTop - iTopTrim) + (iBottom - iBottomTrim)) / 2

            Dim oBoundingRectangleCut As New System.Drawing.Rectangle(iLeft, iTop, iRight - iLeft + 1, iBottom - iTop + 1)
            Using oSubMatCut As New Emgu.CV.Mat(oMatrix.Mat, oBoundingRectangleCut)
                Dim oSubMatrixRectangleCut As New System.Drawing.Rectangle(iCenterDisplacementX + (iMaxXY - oBoundingRectangleCut.Width) / 2, iCenterDisplacementY + (iMaxXY - oBoundingRectangleCut.Height) / 2, oBoundingRectangleCut.Width, oBoundingRectangleCut.Height)
                Using oSubMatrixCut As Emgu.CV.Matrix(Of Byte) = oSubMatrix.GetSubRect(oSubMatrixRectangleCut)
                    oSubMatCut.CopyTo(oSubMatrixCut)
                End Using
            End Using

            Return oSubMatrix
        End Function
        Private Shared Function GetBoundingRectangle(ByVal oMatrix As Emgu.CV.Matrix(Of Byte)) As System.Drawing.Rectangle
            ' get bounding rectangle
            Const iLowerBlobAreaLimit As Integer = 2

            ' find contours
            Dim oContours As New Emgu.CV.Util.VectorOfVectorOfPoint
            Emgu.CV.CvInvoke.FindContours(ExtractImage(oMatrix), oContours, Nothing, Emgu.CV.CvEnum.RetrType.External, Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxTc89L1)

            Dim oBoundingRectangle As New System.Drawing.Rectangle(0, 0, oMatrix.Width, oMatrix.Height)
            If oContours.Size > 0 Then
                Dim oBoundingRectangleList As New List(Of System.Drawing.Rectangle)
                For i = 0 To oContours.Size - 1
                    oBoundingRectangleList.Add(Emgu.CV.CvInvoke.BoundingRectangle(oContours(i)))
                Next
                oBoundingRectangleList = (From oRectangle In oBoundingRectangleList Order By (oRectangle.Width * oRectangle.Height) Descending Select oRectangle).ToList
                oBoundingRectangle = oBoundingRectangleList(0)
                Using oBlobDetector As New Emgu.CV.Cvb.CvBlobDetector
                    For i = 1 To oBoundingRectangleList.Count - 1
                        Using oThresholdMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                            Emgu.CV.CvInvoke.Threshold(oMatrix, oThresholdMatrix, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)
                            Using oSubRect As Emgu.CV.Matrix(Of Byte) = oThresholdMatrix.GetSubRect(oBoundingRectangleList(i))
                                Using oBlobs As New Emgu.CV.Cvb.CvBlobs
                                    oBlobDetector.Detect(oSubRect.Mat.ToImage(Of Emgu.CV.Structure.Gray, Byte), oBlobs)
                                    If oBlobs.Count > 0 Then
                                        Dim iLargestBlobArea As Integer = (From oBlob In oBlobs Order By oBlob.Value.Area Descending Select oBlob.Value.Area).First
                                        If iLargestBlobArea > iLowerBlobAreaLimit Then
                                            oBoundingRectangle = System.Drawing.Rectangle.Union(oBoundingRectangle, oBoundingRectangleList(i))
                                        End If
                                    End If
                                End Using
                            End Using
                        End Using
                    Next
                End Using
            End If
            Return oBoundingRectangle
        End Function
        Private Shared Function ExtractImage(ByRef oMatrix As Emgu.CV.Matrix(Of Byte)) As Emgu.CV.Matrix(Of Byte)
            ' two-stage threshold to pick up faint marks
            Dim oReturnMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
            Using oImageMask1 As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                Emgu.CV.CvInvoke.Threshold(oMatrix, oImageMask1, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)
                Using oImageMask2 As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                    Emgu.CV.CvInvoke.BitwiseXor(oMatrix, oImageMask1, oImageMask2)
                    Emgu.CV.CvInvoke.Threshold(oImageMask2, oImageMask2, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)
                    Emgu.CV.CvInvoke.BitwiseOr(oImageMask1, oImageMask2, oReturnMatrix)
                End Using
            End Using
            Emgu.CV.CvInvoke.Threshold(oReturnMatrix, oReturnMatrix, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)

            Return oReturnMatrix
        End Function
        Public Shared Function GetCoveredFraction(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal bInvert As Boolean) As Double
            ' returns the fraction of the matrix that is covered by the image
            Using oDetectMatrix As Emgu.CV.Matrix(Of Byte) = If(bInvert, oMatrix.SubR(Byte.MaxValue), oMatrix.Clone)
                Emgu.CV.CvInvoke.Threshold(oDetectMatrix, oDetectMatrix, Byte.MaxValue * 0.9, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.BinaryInv)
                Emgu.CV.CvInvoke.MorphologyEx(oDetectMatrix, oDetectMatrix, Emgu.CV.CvEnum.MorphOp.Open, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Rectangle, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))

                Dim oPixelBag As New ConcurrentBag(Of Double)
                Dim fCenterX As Double = (oDetectMatrix.Width - 1) / 2
                Dim fCenterY As Double = (oDetectMatrix.Height - 1) / 2

                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                            Dim oBufferIn As Byte(,) = CType(oParam.Item4, Byte(,))
                                                            Dim fCenterXLocal As Double = CType(oParam.Item5, Double)
                                                            Dim fCenterYLocal As Double = CType(oParam.Item6, Double)
                                                            Dim localwidth As Integer = CType(oParam.Item7, Integer)
                                                            Dim fPixelSumLocal As Double = 0
                                                            Dim fPixelCountLocal As Double = 0

                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                For x = 0 To localwidth - 1
                                                                    Dim oPoint As New Accord.Point(x, localY)
                                                                    Dim fLength As Double = Math.Max(Math.Sqrt((localY - fCenterYLocal) * (localY - fCenterYLocal) + (x - fCenterXLocal) * (x - fCenterXLocal)), 1)
                                                                    If oBufferIn(localY, x) <> 0 Then
                                                                        fPixelSumLocal += 1 / fLength
                                                                    End If
                                                                    fPixelCountLocal += 1 / fLength
                                                                Next
                                                            Next

                                                            If fPixelCountLocal > 0 Then
                                                                oPixelBag.Add(fPixelSumLocal / fPixelCountLocal)
                                                            End If
                                                        End Sub
                CommonFunctions.ParallelRun(oDetectMatrix.Height, oDetectMatrix.Data, fCenterX, fCenterY, oDetectMatrix.Width, TaskDelegate)

                Return oPixelBag.Sum
            End Using
        End Function
        Public Shared Function GetKnnNearestNeighbourCount(ByVal oDetector As Detectors) As Integer
            ' returns the optimal nearest neighbour count
            Select Case oDetector
                Case Detectors.knnIntensity
                    Return 5
                Case Detectors.knnHog
                    Return 3
                Case Detectors.knnStroke
                    Return 2
                Case Else
                    Return Nothing
            End Select
        End Function
        Public Shared Function NormaliseDictionary(ByVal oInputDictionary As Dictionary(Of Integer, Double)) As Dictionary(Of Integer, Double)
            ' replaces the values in the input dictionary with normalised values
            Dim oNormalisedValueArray As Double() = NormaliseResultArray(oInputDictionary.Values.ToArray)
            Return (From iIndex In Enumerable.Range(0, oInputDictionary.Count) Select New KeyValuePair(Of Integer, Double)(oInputDictionary.Keys(iIndex), oNormalisedValueArray(iIndex))).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        End Function
        Public Shared Function NormaliseResultArray(ByVal oInputArray As Double()) As Double()
            ' normalises the array by setting the maximum value to 1
            Dim fMax As Double = oInputArray.Max
            Return (From fValue As Double In oInputArray Select (fValue / fMax)).ToArray
        End Function
        Public Shared Function DeskewMat(ByVal oMatrix As Emgu.CV.Matrix(Of Byte)) As Emgu.CV.Matrix(Of Byte)
            ' deskews images
            Dim oReturnMatrix As Emgu.CV.Matrix(Of Byte) = Nothing
            Using oInverseMatrix As Emgu.CV.Matrix(Of Byte) = oMatrix.SubR(Byte.MaxValue)
                Emgu.CV.CvInvoke.Normalize(oInverseMatrix, oInverseMatrix, Byte.MinValue, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)

                Dim iSliceSize As Integer = oMatrix.Height / 8
                Dim oPointList As New List(Of System.Drawing.PointF)
                For j = 0 To 7
                    Dim iSliceStart As Integer = j * iSliceSize
                    Dim oRect As New System.Drawing.Rectangle(0, iSliceStart, oInverseMatrix.Width, Math.Min((j + 1) * iSliceSize, oInverseMatrix.Height) - iSliceStart)
                    Dim oPoint As Point = GetMatrixCenter(oInverseMatrix, oRect, 0.05)
                    If Not (oPoint.X.Equals(Double.NaN) Or oPoint.Y.Equals(Double.NaN)) Then
                        oPointList.Add(New System.Drawing.PointF(oPoint.X, oPoint.Y))
                    End If
                Next

                ' only calculate the line if more than two point present
                If oPointList.Count > 2 Then
                    Dim oDirection As System.Drawing.PointF = Nothing
                    Dim oPointOnLine As System.Drawing.PointF = Nothing
                    Emgu.CV.CvInvoke.FitLine(oPointList.ToArray, oDirection, oPointOnLine, Emgu.CV.CvEnum.DistType.L2, 0, 0.01, 0.01)

                    Dim fHalfWidth As Single = (oMatrix.Width - 1) / 2
                    Dim oTopMidPointF As New System.Drawing.PointF(fHalfWidth, 0)
                    Dim oBottomMidPointF As New System.Drawing.PointF(fHalfWidth, oMatrix.Height - 1)
                    Dim oTopPointF As System.Drawing.PointF = ProjectPointF(oDirection, oPointOnLine, oTopMidPointF.Y)
                    Dim oBottomPointF As System.Drawing.PointF = ProjectPointF(oDirection, oPointOnLine, oBottomMidPointF.Y)

                    Dim oMidPointList As New List(Of System.Drawing.PointF)
                    oMidPointList.Add(New System.Drawing.PointF(oTopMidPointF.X - fHalfWidth, oTopMidPointF.Y))
                    oMidPointList.Add(oTopMidPointF)
                    oMidPointList.Add(New System.Drawing.PointF(oTopMidPointF.X + fHalfWidth, oTopMidPointF.Y))
                    oMidPointList.Add(New System.Drawing.PointF(oBottomMidPointF.X - fHalfWidth, oBottomMidPointF.Y))
                    oMidPointList.Add(oBottomMidPointF)
                    oMidPointList.Add(New System.Drawing.PointF(oBottomMidPointF.X + fHalfWidth, oBottomMidPointF.Y))

                    oPointList.Clear()
                    oPointList.Add(New System.Drawing.PointF(oTopPointF.X - fHalfWidth, oTopPointF.Y))
                    oPointList.Add(oTopPointF)
                    oPointList.Add(New System.Drawing.PointF(oTopPointF.X + fHalfWidth, oTopPointF.Y))
                    oPointList.Add(New System.Drawing.PointF(oBottomPointF.X - fHalfWidth, oBottomPointF.Y))
                    oPointList.Add(oBottomPointF)
                    oPointList.Add(New System.Drawing.PointF(oBottomPointF.X + fHalfWidth, oBottomPointF.Y))

                    oReturnMatrix = New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                    Using oHomographyMat As New Emgu.CV.Mat(3, 3, Emgu.CV.CvEnum.DepthType.Cv64F, 1)
                        Emgu.CV.CvInvoke.FindHomography(oPointList.ToArray, oMidPointList.ToArray, oHomographyMat, Emgu.CV.CvEnum.HomographyMethod.Default)
                        Emgu.CV.CvInvoke.WarpPerspective(oMatrix, oReturnMatrix, oHomographyMat, oReturnMatrix.Size, Emgu.CV.CvEnum.Inter.Lanczos4, Emgu.CV.CvEnum.Warp.FillOutliers, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MaxValue))
                    End Using
                Else
                    oReturnMatrix = oMatrix.Clone
                End If
            End Using
            Return oReturnMatrix
        End Function
        Private Shared Function ProjectPointF(ByVal oDirection As System.Drawing.PointF, oPointOnLine As System.Drawing.PointF, ByVal fCoordinateY As Single) As System.Drawing.PointF
            ' projects a y-coordinate from a supplied direction and point on a line to an equivalent point on the same line
            Return New System.Drawing.PointF(oPointOnLine.X + ((fCoordinateY - oPointOnLine.Y) * oDirection.X / oDirection.Y), fCoordinateY)
        End Function
        Public Shared Function GetMatrixCenter(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), ByVal oRect As System.Drawing.Rectangle, ByVal fCoveredFraction As Double) As Point
            ' gets center of gravity of the submatrix in the rectangle
            ' this is expressed in relation to the parent matrix
            Using oSubMatrix As Emgu.CV.Matrix(Of Byte) = oMatrix.GetSubRect(oRect)
                If GetCoveredFraction(oSubMatrix, True) > fCoveredFraction Then
                    Dim oCenter As Emgu.CV.Structure.MCvPoint2D64f = Emgu.CV.CvInvoke.Moments(oSubMatrix).GravityCenter
                    Return New Point(oRect.X + oCenter.X, oRect.Y + oCenter.Y)
                Else
                    Return New Point(Double.NaN, Double.NaN)
                End If
            End Using
        End Function
        Public Shared Function SegmentExtractor(ByVal oMatrix As Emgu.CV.Matrix(Of Byte), Optional ByRef oPointCircleList As PointList = Nothing) As Tuple(Of List(Of PointList.Segment), List(Of PointList.SinglePoint))
            ' extracts a list of line segments connecting the points
            ' returns this list along with a list of leftover points without connecting line segments
            Const fMinFraction As Single = 0.25
            Using oImageMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Rows * ScaleFactor, oMatrix.Cols * ScaleFactor)
                Emgu.CV.CvInvoke.Resize(oMatrix, oImageMatrix, oImageMatrix.Size, 0, 0, Emgu.CV.CvEnum.Inter.Cubic)

                ' two-stage threshold to pick up faint marks
                Using oImageMask1 As New Emgu.CV.Matrix(Of Byte)(oImageMatrix.Size)
                    Emgu.CV.CvInvoke.Threshold(oImageMatrix, oImageMask1, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)
                    Using oImageMask2 As New Emgu.CV.Matrix(Of Byte)(oImageMatrix.Size)
                        Emgu.CV.CvInvoke.BitwiseAnd(oImageMatrix, oImageMask1, oImageMask2)
                        Emgu.CV.CvInvoke.Threshold(oImageMask2, oImageMask2, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)
                        Emgu.CV.CvInvoke.BitwiseAnd(oImageMask1, oImageMask2, oImageMatrix)
                    End Using
                End Using

                ' create point circle list from skeleton
                oPointCircleList = New PointList
                Using oSkeletonMatrix As Emgu.CV.Matrix(Of Byte) = SkeletonExtractor(oImageMatrix)
                    Dim oPixelBag As New ConcurrentBag(Of PointList.SinglePoint)
                    Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                                Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                                Dim oBufferIn As Emgu.CV.Matrix(Of Byte) = CType(oParam.Item4, Emgu.CV.Matrix(Of Byte))
                                                                Dim oBufferOut As ConcurrentBag(Of PointList.SinglePoint) = CType(oParam.Item5, ConcurrentBag(Of PointList.SinglePoint))
                                                                Dim localwidth As Integer = CType(oParam.Item6, Integer)
                                                                Dim oDistanceMapLocal As Emgu.CV.Matrix(Of Single) = CType(oParam.Item7, Emgu.CV.Matrix(Of Single))

                                                                For y = 0 To oParam.Item2 - 1
                                                                    Dim localY As Integer = y + oParam.Item1
                                                                    For x = 0 To localwidth - 1
                                                                        If oBufferIn(localY, x) <> Byte.MinValue Then
                                                                            oBufferOut.Add(New PointList.SinglePoint(oDistanceMapLocal(localY, x), New AForge.Point(x, localY), False))
                                                                        End If
                                                                    Next
                                                                Next
                                                            End Sub
                    Using oDistanceMap As New Emgu.CV.Matrix(Of Single)(oImageMatrix.Size)
                        Emgu.CV.CvInvoke.DistanceTransform(oImageMatrix, oDistanceMap, Nothing, Emgu.CV.CvEnum.DistType.L2, 0)
                        CommonFunctions.ParallelRun(oSkeletonMatrix.Height, oSkeletonMatrix, oPixelBag, oSkeletonMatrix.Width, oDistanceMap, TaskDelegate)
                    End Using

                    Dim oPixelList As List(Of PointList.SinglePoint) = oPixelBag.OrderByDescending(Function(x) x.Radius).ToList
                    If oPixelList.Count > 0 Then
                        Dim fSetMax As Double = oPixelList(0).Radius * fMinFraction
                        Do Until oPixelList.Count = 0 OrElse oPixelList(0).Radius < fSetMax
                            oPointCircleList.Add(oPixelList(0))
                            For iIndex = oPixelList.Count - 1 To 1 Step -1
                                If PointList.InCircle(oPixelList(0), oPixelList(iIndex)) Then
                                    oPixelList.RemoveAt(iIndex)
                                End If
                            Next
                            oPixelList.RemoveAt(0)
                        Loop
                    End If
                End Using

                ' get list of line segments
                Dim oLineSegments As New List(Of PointList.Segment)
                For i = 0 To oPointCircleList.Count - 1
                    Using oNeighbouringPoints As PointList = PointList.GetNeighbouringPoints(oPointCircleList(i), oPointCircleList)
                        If oNeighbouringPoints.Count > 0 Then
                            oPointCircleList(i) = New PointList.SinglePoint(oPointCircleList(i).Radius, oPointCircleList(i).AForgePoint, True)
                            For Each oPoint As PointList.SinglePoint In oNeighbouringPoints
                                oLineSegments.Add(New PointList.Segment(oPointCircleList(i), oPoint))
                            Next
                        End If
                    End Using
                Next
                oLineSegments = oLineSegments.Distinct.ToList

                ' get list of leftover points
                Dim oLeftoverPoints As List(Of PointList.SinglePoint) = (From oPoint As PointList.SinglePoint In oPointCircleList Where Not oPoint.Selected Select oPoint).ToList

                Return New Tuple(Of List(Of PointList.Segment), List(Of PointList.SinglePoint))(oLineSegments, oLeftoverPoints)
            End Using
        End Function
        Private Shared Function SkeletonExtractor(ByVal oMatrix As Emgu.CV.Matrix(Of Byte)) As Emgu.CV.Matrix(Of Byte)
            Dim oSkeletonMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
            Using oTempImageMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                Using oErodedMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                    Using oTempMatrix As New Emgu.CV.Matrix(Of Byte)(oMatrix.Size)
                        oSkeletonMatrix.SetZero()
                        Emgu.CV.CvInvoke.Threshold(oMatrix, oTempImageMatrix, 0, Byte.MaxValue, Emgu.CV.CvEnum.ThresholdType.Binary Or Emgu.CV.CvEnum.ThresholdType.Otsu)

                        Dim done As Boolean
                        Do
                            Emgu.CV.CvInvoke.Erode(oTempImageMatrix, oErodedMatrix, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                            Emgu.CV.CvInvoke.Dilate(oErodedMatrix, oTempMatrix, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                            Emgu.CV.CvInvoke.Subtract(oTempImageMatrix, oTempMatrix, oTempImageMatrix)
                            Emgu.CV.CvInvoke.BitwiseOr(oSkeletonMatrix, oTempImageMatrix, oSkeletonMatrix)
                            oErodedMatrix.CopyTo(oTempImageMatrix)
                            done = (Emgu.CV.CvInvoke.CountNonZero(oTempImageMatrix) = 0)
                        Loop While Not done
                    End Using
                End Using
            End Using
            Return oSkeletonMatrix
        End Function
        Public Enum Detectors As Integer
            None = 0
            knnIntensity
            knnHog
            knnStroke
            SVMIntensity
            SVMHog
            SVMStroke
            DeepIntensity
            DeepHog
            DeepStroke
        End Enum
        Public Shared Function GetSVMType(ByVal oDetector As Detectors) As KernelType
            ' returns the type of kernel used for a particular detector
            Select Case oDetector
                Case Detectors.SVMIntensity
                    Return KernelType.Linear
                Case Detectors.SVMHog
                    Return KernelType.Linear
                Case Detectors.SVMStroke
                    Return KernelType.Linear
                Case Else
                    Return Nothing
            End Select
        End Function
        Public Enum KernelType As Integer
            Linear
            Polynomial2
            Polynomial5
            Gaussian1
            Gaussian2
            Gaussian3
            Sigmoid
            HistogramIntersection
        End Enum
    End Class
    Public Class Enumerations
        Public Enum CenterAlignment As Integer
            Center = 0
            Top
            Bottom
            Left
            Right
        End Enum
        <Flags> Public Enum ScaleDirection As Integer
            None = 0
            Vertical = 1
            Horizontal = 2
            Both = 3
        End Enum
        Public Enum ScaleType As Integer
            Both = 0
            Down
            Up
        End Enum
        <Flags> Public Enum FontEnum As Integer
            None = 0
            Bold = 1
            Italic = 2
            Underline = 4
        End Enum
        Public Enum JustificationEnum As Integer
            Left = 0
            Center
            Right
            Justify
        End Enum
        Public Enum AlignmentEnum As Integer
            Left = 0
            UpperLeft
            Top
            UpperRight
            Right
            LowerRight
            Bottom
            LowerLeft
            Center
        End Enum
        Public Enum StretchEnum As Integer
            Fill = 0
            Uniform
        End Enum
        Public Enum TabletContentEnum As Integer
            None = 0
            Letter = 1
            Number = 2
        End Enum
        <Flags> Public Enum CharacterASCII As Integer
            None = 0
            Numbers = 1
            Uppercase = 2
            Lowercase = 4
            NonAlphaNumeric = 8
        End Enum
        Public Enum FieldTypeEnum As Integer
            Undefined = 0
            Border
            Background
            Numbering
            Text
            Image
            Choice
            ChoiceVertical
            ChoiceVerticalMCQ
            BoxChoice
            Handwriting
            Free
        End Enum
        Public Enum Numbering As Integer
            Number = 0
            LetterSmall = 1
            LetterBig = 2
        End Enum
        Public Enum InputDevice As Integer
            NoDevice = 0
            Mouse
            Stylus
            Touch
        End Enum
        Public Enum InputDeviceState As Integer
            StateStart = 0
            StateEnd
            StateMove
        End Enum
        Public Enum FilterData
            None = 0
            DataPresent
            DataMissing
        End Enum
        Public Enum TabletType As Integer
            Empty = 0
            Filled
            EmptyBlack
            Inner
            EmptyBlackInner
        End Enum
    End Class
#End Region
#Region "Interfaces"
    Public Interface FunctionInterface
        ' returns a list of GUIDs and variable type representing the data types that the plug-in creates
        Function GetDataTypes() As List(Of Tuple(Of Guid, Type))

        ' checks the commonvariable data store to see if the variables objects required have been properly initialised
        Function CheckDataTypes(ByRef oCommonVariables As CommonVariables) As Boolean

        ' returns the identifiers: GUID, icon, friendly name, and priority for the plugin
        Function GetIdentifier() As Tuple(Of Guid, Media.ImageSource, String, Integer)

        WriteOnly Property Identifiers As Dictionary(Of String, Tuple(Of Guid, Media.ImageSource, String, Integer))

        ' the status message event handler to update the main program status window
        Event StatusMessage(oMessage As Messages.Message)

        ' propogates the exit message to the parent page
        Event ExitButtonClick(oActivatePluginGUID As Guid)
    End Interface
#End Region
    Public Class Converter
#Region "Bitmaps"
        Public Shared Function MatrixToArray(ByVal oMatrix As Emgu.CV.Matrix(Of Byte)) As Byte()
            ' convert a matrix to a byte array
            Using oMatrixClone As Emgu.CV.Matrix(Of Byte) = oMatrix.Clone
                Return Accord.Math.Matrix.Flatten(oMatrixClone.Data)
            End Using
        End Function
        Public Shared Function ArrayToMatrix(ByVal oArray As Byte(), ByVal width As Integer, ByVal height As Integer) As Emgu.CV.Matrix(Of Byte)
            ' convert a two dimensional array to a one dimensional array equivalent
            Dim oData As Byte(,) = Accord.Math.Matrix.Reshape(Of Byte)(oArray, height, width)
            Return New Emgu.CV.Matrix(Of Byte)(oData)
        End Function
        Public Shared Function MatrixToDoubleArray(ByRef oMatrix As Emgu.CV.Matrix(Of Byte)) As Double()
            ' converts byte matrix to double jagged array
            Dim width As Integer = oMatrix.Cols
            Dim height As Integer = oMatrix.Rows
            Dim oReturnDoubleArray((width * height) - 1) As Double

            Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                        Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                        Dim oBufferIn As Emgu.CV.Matrix(Of Byte) = CType(oParam.Item4, Emgu.CV.Matrix(Of Byte))
                                                        Dim oBufferOut As Double() = CType(oParam.Item5, Double())
                                                        Dim localwidth As Integer = CType(oParam.Item6, Integer)

                                                        For y = 0 To oParam.Item2 - 1
                                                            Dim localY As Integer = y + oParam.Item1
                                                            For x = 0 To localwidth - 1
                                                                oBufferOut(localY * localwidth + x) = oBufferIn(localY, x)
                                                            Next
                                                        Next
                                                    End Sub
            CommonFunctions.ParallelRun(height, oMatrix, oReturnDoubleArray, width, Nothing, TaskDelegate)

            Return oReturnDoubleArray
        End Function
        Public Shared Function BitmapToMatrix(ByVal oImage As System.Drawing.Bitmap) As Emgu.CV.Matrix(Of Byte)
            ' convert bitmap to mat
            If IsNothing(oImage) Then
                Return Nothing
            Else
                Select Case oImage.PixelFormat
                    Case System.Drawing.Imaging.PixelFormat.Format8bppIndexed
                        Try
                            Return BitmapToMatrix8Bit(oImage)
                        Catch ex As OutOfMemoryException
                            CommonFunctions.ClearMemory()
                            Return BitmapToMatrix8Bit(oImage)
                        End Try
                    Case System.Drawing.Imaging.PixelFormat.Format32bppArgb
                        Using oGrayImage As System.Drawing.Bitmap = BitmapConvertGrayscale(oImage)
                            Try
                                Return BitmapToMatrix8Bit(oGrayImage)
                            Catch ex As OutOfMemoryException
                                CommonFunctions.ClearMemory()
                                Return BitmapToMatrix8Bit(oGrayImage)
                            End Try
                        End Using
                    Case Else
                        Return Nothing
                End Select
            End If
        End Function
        Private Shared Function BitmapToMatrix8Bit(ByRef oImage As System.Drawing.Bitmap) As Emgu.CV.Matrix(Of Byte)
            If IsNothing(oImage) OrElse oImage.PixelFormat <> System.Drawing.Imaging.PixelFormat.Format8bppIndexed Then
                Return Nothing
            Else
                ' convert jagged buffer to open cv mat
                Dim width As Integer = oImage.Width
                Dim height As Integer = oImage.Height

                Dim rect As New System.Drawing.Rectangle(0, 0, width, height)
                Dim oImageData As System.Drawing.Imaging.BitmapData = oImage.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, oImage.PixelFormat)
                Dim oImagePtr As IntPtr = oImageData.Scan0

                Dim oReturnMatrix As New Emgu.CV.Matrix(Of Byte)(height, width, 1)
                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                            Dim oBufferIn As IntPtr = CType(oParam.Item4, IntPtr)
                                                            Dim oBufferOut As IntPtr = CType(oParam.Item5, IntPtr)
                                                            Dim localwidth As Integer = CType(oParam.Item6, Integer)
                                                            Dim iStrideIn As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item1
                                                            Dim iStrideOut As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item2
                                                            Dim oBufferArrayLocal(localwidth - 1) As Byte

                                                            Dim iIndexIn, iIndexOut As Integer
                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                iIndexIn = localY * iStrideIn
                                                                iIndexOut = localY * iStrideOut
                                                                Runtime.InteropServices.Marshal.Copy(oBufferIn + iIndexIn, oBufferArrayLocal, 0, localwidth)
                                                                Runtime.InteropServices.Marshal.Copy(oBufferArrayLocal, 0, oBufferOut + iIndexOut, localwidth)
                                                            Next
                                                        End Sub
                CommonFunctions.ParallelRun(height, oImagePtr, oReturnMatrix.Mat.DataPointer, width, New Tuple(Of Integer, Integer)(oImageData.Stride, oReturnMatrix.Mat.Step), TaskDelegate)

                oImage.UnlockBits(oImageData)
                Return oReturnMatrix
            End If
        End Function
        Public Shared Function BitmapToBitmapSource(ByVal oBitmap As System.Drawing.Bitmap, Optional bForceSmall As Boolean = False) As Media.Imaging.BitmapSource
            ' converts GDI bitmap to WPF bitmapsource (small sizes)
            Const iBigThreshold As Integer = 4000000
            If IsNothing(oBitmap) Then
                Return Nothing
            Else
                If (Not bForceSmall) AndAlso oBitmap.Width * oBitmap.Height * 4 > iBigThreshold Then
                    ' big bitmap
                    Try
                        Return BitmapToBitmapSourceBig(oBitmap)
                    Catch ex As OutOfMemoryException
                        CommonFunctions.ClearMemory()
                        Return BitmapToBitmapSourceBig(oBitmap)
                    End Try
                Else
                    ' small bitmap
                    Return BitmapToBitmapSourceSmall(oBitmap)
                End If
            End If
        End Function
        Private Shared Function BitmapToBitmapSourceSmall(ByVal oBitmap As System.Drawing.Bitmap) As Media.Imaging.BitmapSource
            If oBitmap.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppArgb Then
                Dim oRect As New System.Drawing.Rectangle(0, 0, oBitmap.Width, oBitmap.Height)
                Dim oBitmapData As System.Drawing.Imaging.BitmapData = oBitmap.LockBits(oRect, System.Drawing.Imaging.ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                Dim iBufferSize As Integer = oBitmapData.Stride * oBitmap.Height
                Dim oWriteableBitmap As New Media.Imaging.WriteableBitmap(oBitmap.Width, oBitmap.Height, oBitmap.HorizontalResolution, oBitmap.VerticalResolution, Media.PixelFormats.Bgra32, Nothing)
                oWriteableBitmap.WritePixels(New Int32Rect(0, 0, oBitmap.Width, oBitmap.Height), oBitmapData.Scan0, iBufferSize, oBitmapData.Stride)
                oBitmap.UnlockBits(oBitmapData)
                Return oWriteableBitmap
            Else
                Return Nothing
            End If
        End Function
        Private Shared Function BitmapToBitmapSourceBig(ByRef oBitmap As System.Drawing.Bitmap) As Media.Imaging.BitmapSource
            ' converts GDI bitmap to WPF bitmapsource (big sizes)
            If oBitmap.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppArgb Then
                Dim rect As New System.Drawing.Rectangle(0, 0, oBitmap.Width, oBitmap.Height)

                Dim oBitmapData As System.Drawing.Imaging.BitmapData = oBitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, oBitmap.PixelFormat)
                Dim oBitmapPtr As IntPtr = oBitmapData.Scan0

                Dim iImageBitsPerPixel, iImageBytesPerPixel, iImageStride, iImageBytes As Integer
                iImageBitsPerPixel = Media.PixelFormats.Bgra32.BitsPerPixel
                iImageBytesPerPixel = Math.Truncate((iImageBitsPerPixel + 7) / 8)
                iImageStride = 4 * Math.Truncate(((oBitmap.Width * iImageBytesPerPixel) + 3) / 4)
                iImageBytes = iImageStride * oBitmap.Height
                Dim oByteBuffer(iImageBytes - 1) As Byte

                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                            Dim oBufferIn As IntPtr = CType(oParam.Item4, IntPtr)
                                                            Dim oBufferOut As Byte() = CType(oParam.Item5, Byte())
                                                            Dim localwidth As Integer = CType(oParam.Item6, Integer)
                                                            Dim iStrideIn As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item1
                                                            Dim iStrideOut As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item2

                                                            Dim iIndexIn, iIndexOut As Integer
                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                iIndexIn = localY * iStrideIn
                                                                iIndexOut = localY * iStrideOut
                                                                Runtime.InteropServices.Marshal.Copy(oBufferIn + iIndexIn, oBufferOut, iIndexOut, localwidth * 4)
                                                            Next
                                                        End Sub
                CommonFunctions.ParallelRun(oBitmap.Height, oBitmapPtr, oByteBuffer, oBitmap.Width, New Tuple(Of Integer, Integer)(oBitmapData.Stride, iImageStride), TaskDelegate)

                oBitmap.UnlockBits(oBitmapData)

                Dim oBitmapsource As Media.Imaging.BitmapSource = Media.Imaging.BitmapSource.Create(oBitmap.Width, oBitmap.Height, oBitmap.HorizontalResolution, oBitmap.VerticalResolution, Media.PixelFormats.Bgra32, Nothing, oByteBuffer, iImageStride)
                oBitmapsource.Freeze()
                Return oBitmapsource
            Else
                Return Nothing
            End If
        End Function
        Public Shared Function BitmapSourceToBitmap(ByVal oBitmapSource As Media.Imaging.BitmapSource, Optional ByVal iResolution As Integer = 0) As System.Drawing.Bitmap
            ' converts WPF bitmapsource to GDI bitmap
            Dim oFormatConvertedBitmap As New Media.Imaging.FormatConvertedBitmap
            oFormatConvertedBitmap.BeginInit()
            oFormatConvertedBitmap.Source = oBitmapSource
            oFormatConvertedBitmap.DestinationFormat = Media.PixelFormats.Bgra32
            oFormatConvertedBitmap.EndInit()

            Dim oBitmap As New System.Drawing.Bitmap(oFormatConvertedBitmap.PixelWidth, oFormatConvertedBitmap.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Dim oBitmapData = oBitmap.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, oBitmap.Size), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            oFormatConvertedBitmap.CopyPixels(Int32Rect.Empty, oBitmapData.Scan0, oBitmapData.Height * oBitmapData.Stride, oBitmapData.Stride)
            oBitmap.UnlockBits(oBitmapData)

            If iResolution = 0 Then
                ' use bitmapsource resolution
                oBitmap.SetResolution(oBitmapSource.DpiX, oBitmapSource.DpiY)
            Else
                oBitmap.SetResolution(iResolution, iResolution)
            End If

            Return oBitmap
        End Function
        Public Shared Function BitmapSourceToMatrix8Bit(ByVal oBitmapSource As Media.Imaging.BitmapSource) As Emgu.CV.Matrix(Of Byte)
            ' converts WPF bitmapsource to byte matrix
            If IsNothing(oBitmapSource) Then
                Return Nothing
            Else
                Dim oFormatConvertedBitmap As New Media.Imaging.FormatConvertedBitmap
                oFormatConvertedBitmap.BeginInit()
                oFormatConvertedBitmap.Source = oBitmapSource
                oFormatConvertedBitmap.DestinationFormat = Media.PixelFormats.Gray8
                oFormatConvertedBitmap.EndInit()

                ' convert jagged buffer to open cv mat
                Dim width As Integer = oBitmapSource.Width
                Dim height As Integer = oBitmapSource.Height

                Dim oReturnMatrix As New Emgu.CV.Matrix(Of Byte)(height, width, 1)
                oFormatConvertedBitmap.CopyPixels(Int32Rect.Empty, oReturnMatrix.Mat.DataPointer, height * oReturnMatrix.Mat.Step, oReturnMatrix.Mat.Step)

                Return oReturnMatrix
            End If
        End Function
        Public Shared Function BitmapToBitmapSourceGrayscale(ByRef oBitmap As System.Drawing.Bitmap) As Media.Imaging.BitmapSource
            ' converts GDI bitmap to WPF bitmapsource (grayscale)
            If (Not IsNothing(oBitmap)) AndAlso oBitmap.PixelFormat = System.Drawing.Imaging.PixelFormat.Format8bppIndexed Then
                Dim rect As New System.Drawing.Rectangle(0, 0, oBitmap.Width, oBitmap.Height)

                Dim oBitmapData As System.Drawing.Imaging.BitmapData = oBitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, oBitmap.PixelFormat)
                Dim oBitmapPtr As IntPtr = oBitmapData.Scan0

                Dim iImageBitsPerPixel, iImageBytesPerPixel, iImageStride, iImageBytes As Integer
                iImageBitsPerPixel = Media.PixelFormats.Gray8.BitsPerPixel
                iImageBytesPerPixel = Math.Truncate((iImageBitsPerPixel + 7) / 8)
                iImageStride = 4 * Math.Truncate(((oBitmap.Width * iImageBytesPerPixel) + 3) / 4)
                iImageBytes = iImageStride * oBitmap.Height
                Dim oByteBuffer(iImageBytes - 1) As Byte

                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                            Dim oBufferIn As IntPtr = CType(oParam.Item4, IntPtr)
                                                            Dim oBufferOut As Byte() = CType(oParam.Item5, Byte())
                                                            Dim localwidth As Integer = CType(oParam.Item6, Integer)
                                                            Dim iStrideIn As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item1
                                                            Dim iStrideOut As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item2

                                                            Dim iIndexIn, iIndexOut As Integer
                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                iIndexIn = localY * iStrideIn
                                                                iIndexOut = localY * iStrideOut
                                                                Runtime.InteropServices.Marshal.Copy(oBufferIn + iIndexIn, oBufferOut, iIndexOut, localwidth)
                                                            Next
                                                        End Sub
                CommonFunctions.ParallelRun(oBitmap.Height, oBitmapPtr, oByteBuffer, oBitmap.Width, New Tuple(Of Integer, Integer)(oBitmapData.Stride, iImageStride), TaskDelegate)

                oBitmap.UnlockBits(oBitmapData)

                Dim oBitmapsource As Media.Imaging.BitmapSource = Media.Imaging.BitmapSource.Create(oBitmap.Width, oBitmap.Height, oBitmap.HorizontalResolution, oBitmap.VerticalResolution, Media.PixelFormats.Gray8, Media.Imaging.BitmapPalettes.Gray256, oByteBuffer, iImageStride)
                oBitmapsource.Freeze()
                Return oBitmapsource
            Else
                Return Nothing
            End If
        End Function
        Public Shared Function BitmapConvertGrayscale(ByVal oImage As System.Drawing.Bitmap) As System.Drawing.Bitmap
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
        Public Shared Function BitmapConvertColour(ByVal oImage As System.Drawing.Bitmap) As System.Drawing.Bitmap
            If IsNothing(oImage) Then
                Return Nothing
            ElseIf oImage.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppArgb Then
                Return oImage.Clone
            ElseIf oImage.PixelFormat = System.Drawing.Imaging.PixelFormat.Format8bppIndexed Then
                Using oMat As New Emgu.CV.Image(Of Emgu.CV.Structure.Gray, Byte)(oImage)
                    Dim oReturnBitmap = oMat.Convert(Of Emgu.CV.Structure.Bgra, Byte).ToBitmap
                    oReturnBitmap.SetResolution(oImage.HorizontalResolution, oImage.VerticalResolution)
                    Return oReturnBitmap
                End Using
            Else
                Return Nothing
            End If
        End Function
        Public Shared Function BitmapSourceConvertGrayscale(oBitmapSource As Media.Imaging.BitmapSource, Optional ByVal alphaThreshold As Double = 0) As Media.Imaging.BitmapSource
            ' converts colour bitmapsource to grayscale
            Return New Media.Imaging.FormatConvertedBitmap(oBitmapSource, Media.PixelFormats.Gray8, Media.Imaging.BitmapPalettes.Gray256, alphaThreshold)
        End Function
        Public Shared Function MatToBitmap(ByRef oMat As Emgu.CV.Mat, ByVal iResolution As Integer, Optional iResolution2 As Integer = -1) As System.Drawing.Bitmap
            ' convert mat to bitmap
            If IsNothing(oMat) OrElse ((oMat.Depth <> Emgu.CV.CvEnum.DepthType.Cv8U And oMat.Depth <> Emgu.CV.CvEnum.DepthType.Cv32F) Or (oMat.NumberOfChannels <> 1 And oMat.NumberOfChannels <> 3)) Then
                Return Nothing
            Else
                Dim oReturnBitmap As System.Drawing.Bitmap = Nothing
                If oMat.Depth = Emgu.CV.CvEnum.DepthType.Cv8U Then
                    If oMat.NumberOfChannels = 1 Then
                        oReturnBitmap = Mat8BitToBitmap(oMat)
                    ElseIf oMat.NumberOfChannels = 3 Then
                        oReturnBitmap = oMat.ToImage(Of Emgu.CV.Structure.Bgr, Byte).ToBitmap
                    End If
                ElseIf oMat.Depth = Emgu.CV.CvEnum.DepthType.Cv32F Then
                    If oMat.NumberOfChannels = 1 Then
                        oReturnBitmap = BitmapConvertGrayscale(oMat.ToImage(Of Emgu.CV.Structure.Gray, Single).ToBitmap)
                    ElseIf oMat.NumberOfChannels = 3 Then
                        oReturnBitmap = oMat.ToImage(Of Emgu.CV.Structure.Bgr, Byte).ToBitmap
                    End If
                End If

                If Not IsNothing(oReturnBitmap) Then
                    oReturnBitmap.SetResolution(iResolution, If(iResolution2 = -1, iResolution, iResolution2))
                End If
                Return oReturnBitmap
            End If
        End Function
        Private Shared Function Mat8BitToBitmap(ByVal oMat As Emgu.CV.Mat) As System.Drawing.Bitmap
            If IsNothing(oMat) OrElse oMat.Depth <> Emgu.CV.CvEnum.DepthType.Cv8U Then
                Return Nothing
            Else
                ' convert bitmap to open cv mat
                Dim width As Integer = oMat.Cols
                Dim height As Integer = oMat.Rows

                Dim oImage As System.Drawing.Bitmap = AForge.Imaging.Image.CreateGrayscaleImage(width, height)
                Dim rect As New System.Drawing.Rectangle(0, 0, width, height)
                Dim oImageData As System.Drawing.Imaging.BitmapData = oImage.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, oImage.PixelFormat)
                Dim oImagePtr As IntPtr = oImageData.Scan0

                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                            Dim oBufferIn As IntPtr = CType(oParam.Item4, IntPtr)
                                                            Dim oBufferOut As IntPtr = CType(oParam.Item5, IntPtr)
                                                            Dim localwidth As Integer = CType(oParam.Item6, Integer)
                                                            Dim iStrideIn As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item1
                                                            Dim iStrideOut As Integer = CType(oParam.Item7, Tuple(Of Integer, Integer)).Item2
                                                            Dim oBufferArrayLocal(localwidth - 1) As Byte

                                                            Dim iIndexIn, iIndexOut As Integer
                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                iIndexIn = localY * iStrideIn
                                                                iIndexOut = localY * iStrideOut
                                                                Runtime.InteropServices.Marshal.Copy(oBufferIn + iIndexIn, oBufferArrayLocal, 0, localwidth)
                                                                Runtime.InteropServices.Marshal.Copy(oBufferArrayLocal, 0, oBufferOut + iIndexOut, localwidth)
                                                            Next
                                                        End Sub
                CommonFunctions.ParallelRun(height, oMat.DataPointer, oImagePtr, width, New Tuple(Of Integer, Integer)(oMat.Step, oImageData.Stride), TaskDelegate)

                oImage.UnlockBits(oImageData)
                Return oImage
            End If
        End Function
        Public Shared Function BitmapSourceToByteArray(ByVal oImage As Media.Imaging.BitmapSource, ByVal bConvertGrayscale As Boolean) As Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, Media.PixelFormat)
            ' converts a bitmapsource to a compressed byte array
            Dim oFormatConvertedBitmap As Media.Imaging.BitmapSource = Nothing
            If bConvertGrayscale Then
                If oImage.Format = Media.PixelFormats.Gray8 Then
                    oFormatConvertedBitmap = oImage
                Else
                    oFormatConvertedBitmap = BitmapSourceConvertGrayscale(oImage)
                End If
            Else
                oFormatConvertedBitmap = oImage
            End If

            Dim iImageBitsPerPixel, iImageBytesPerPixel, iImageStride, iImageBytes As Integer
            iImageBitsPerPixel = oFormatConvertedBitmap.Format.BitsPerPixel
            iImageBytesPerPixel = Math.Truncate((iImageBitsPerPixel + 7) / 8)
            iImageStride = 4 * Math.Truncate(((oFormatConvertedBitmap.PixelWidth * iImageBytesPerPixel) + 3) / 4)
            iImageBytes = iImageStride * oFormatConvertedBitmap.PixelHeight
            Dim oByteBuffer(iImageBytes - 1) As Byte

            oFormatConvertedBitmap.CopyPixels(Int32Rect.Empty, oByteBuffer, iImageStride, 0)
            Return New Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, Media.PixelFormat)(CommonFunctions.Compress(oByteBuffer), oFormatConvertedBitmap.PixelWidth, oFormatConvertedBitmap.PixelHeight, oFormatConvertedBitmap.DpiX, oFormatConvertedBitmap.DpiY, iImageStride, oFormatConvertedBitmap.Format)
        End Function
        Public Shared Function ByteArrayToBitmapSource(ByVal oByteArray As Tuple(Of Byte(), Integer, Integer, Single, Single, Integer, Media.PixelFormat)) As Media.Imaging.BitmapSource
            ' converts a compressed byte array to a bitmapsource
            Dim oByteBuffer As Byte() = oByteArray.Item1
            Dim iWidth As Integer = oByteArray.Item2
            Dim iHeight As Integer = oByteArray.Item3
            Dim oDpiX As Single = oByteArray.Item4
            Dim oDpiY As Single = oByteArray.Item5
            Dim iImageStride As Integer = oByteArray.Item6
            Dim oPixelFormat As Media.PixelFormat = oByteArray.Item7
            Dim oPalette As Media.Imaging.BitmapPalette = Media.Imaging.BitmapPalettes.Halftone256
            If oPixelFormat = Media.PixelFormats.Gray8 Or oPixelFormat = Media.PixelFormats.Gray16 Then
                oPalette = Media.Imaging.BitmapPalettes.Gray256
            End If

            Dim oBitmapsource As Media.Imaging.BitmapSource = Media.Imaging.BitmapSource.Create(iWidth, iHeight, oDpiX, oDpiY, oPixelFormat, oPalette, CommonFunctions.Decompress(oByteBuffer), iImageStride)
            oBitmapsource.Freeze()
            Return oBitmapsource
        End Function
#End Region
#Region "Misc"
        Public Shared Function AlphaNumericOnly(ByVal sString As String, ByVal bIncludeSpace As Boolean) As String
            Dim i As Integer
            Dim sResult As String = String.Empty

            If bIncludeSpace Then
                For i = 1 To Len(sString)
                    Select Case Asc(Mid(sString, i, 1))
                        Case 32, 48 To 57, 65 To 90, 97 To 122
                            sResult = sResult & Mid(sString, i, 1)
                    End Select
                Next
            Else
                For i = 1 To Len(sString)
                    Select Case Asc(Mid(sString, i, 1))
                        Case 48 To 57, 65 To 90, 97 To 122
                            sResult = sResult & Mid(sString, i, 1)
                    End Select
                Next
            End If
            Return sResult
        End Function
        Public Shared Function ConvertNumberToLetter(ByVal iNumber As Integer, ByVal bUpperCase As Boolean) As String
            ' converts a single or two digit number to a letter
            Dim sLetter As String = String.Empty
            Dim iAscA As Integer = 0
            If bUpperCase Then
                iAscA = Asc("A")
            Else
                iAscA = Asc("a")
            End If
            If iNumber < 26 Then
                ' one char
                sLetter = Chr(iAscA + iNumber)
            ElseIf iNumber < 26 * 26 Then
                ' two char
                sLetter = Chr(iAscA - 1 + CInt(Math.Floor(iNumber / 26))) + Chr(iAscA + CInt(Math.Floor(iNumber Mod 26)))
            End If
            Return sLetter
        End Function
        Public Shared Function ConvertLetterToNumber(ByVal sLetter As Char) As Integer
            ' converts a letter to a number starting from zero ie. A & a = 0
            If Asc(sLetter) >= 65 AndAlso Asc(sLetter) <= 90 Then
                Return Asc(sLetter) - 65
            ElseIf Asc(sLetter) >= 97 AndAlso Asc(sLetter) <= 122 Then
                Return Asc(sLetter) - 97
            Else
                Return -1
            End If
        End Function
        Public Shared Function Int32RectToRectangle(ByVal oInt32Rect As Int32Rect) As System.Drawing.Rectangle
            Return New System.Drawing.Rectangle(oInt32Rect.X, oInt32Rect.Y, oInt32Rect.Width, oInt32Rect.Height)
        End Function
#End Region
#Region "Xaml"
        Public Shared Function CombineDrawingImage(ByVal oDrawingImage1 As Media.DrawingImage, ByVal oDrawingImage2 As Media.DrawingImage, ByVal fScale As Double) As Media.DrawingImage
            ' combines two drawing images with the transform applied to image 2
            Dim oDrawingGroup1 As New Media.DrawingGroup
            oDrawingGroup1.Children.Add(oDrawingImage1.Drawing)

            Dim oDrawingGroup2 As New Media.DrawingGroup
            oDrawingGroup2.Children.Add(oDrawingImage2.Drawing)

            Dim fMatchingScale As Double = Math.Min(oDrawingGroup1.Bounds.Width / oDrawingGroup2.Bounds.Width, oDrawingGroup1.Bounds.Height / oDrawingGroup2.Bounds.Height) * fScale
            Dim fMatchingWidth As Double = oDrawingGroup2.Bounds.Width * fMatchingScale
            Dim fMatchingHeight As Double = oDrawingGroup2.Bounds.Height * fMatchingScale

            Dim fOffsetX1 As Double = 0
            Dim fOffsetY1 As Double = 0
            Dim fOffsetX2 As Double = 0
            Dim fOffsetY2 As Double = 0
            If oDrawingGroup1.Bounds.Width > fMatchingWidth Then
                fOffsetX2 = (oDrawingGroup1.Bounds.Width - fMatchingWidth) / 2
            Else
                fOffsetX1 = (fMatchingWidth - oDrawingGroup1.Bounds.Width) / 2
            End If
            If oDrawingGroup1.Bounds.Height > fMatchingHeight Then
                fOffsetY2 = (oDrawingGroup1.Bounds.Height - fMatchingHeight) / 2
            Else
                fOffsetY1 = (fMatchingHeight - oDrawingGroup1.Bounds.Height) / 2
            End If

            oDrawingGroup1.Transform = New Media.TranslateTransform(-oDrawingGroup1.Bounds.X + fOffsetX1, -oDrawingGroup1.Bounds.Y + fOffsetY1)

            Dim oTransformGroup As New Media.TransformGroup
            oTransformGroup.Children.Add(New Media.TranslateTransform(-oDrawingGroup2.Bounds.X, -oDrawingGroup2.Bounds.Y))
            oTransformGroup.Children.Add(New Media.ScaleTransform(fMatchingScale, fMatchingScale))
            oTransformGroup.Children.Add(New Media.TranslateTransform(fOffsetX2, fOffsetY2))
            oDrawingGroup2.Transform = oTransformGroup

            Dim oCombinedDrawingGroup As New Media.DrawingGroup
            oCombinedDrawingGroup.Children.Add(oDrawingGroup1)
            oCombinedDrawingGroup.Children.Add(oDrawingGroup2)

            Return New Media.DrawingImage(oCombinedDrawingGroup)
        End Function
        Public Shared Function XamlToDrawingImage(ByVal sXaml As String) As Media.DrawingImage
            ' converts text xaml image resource to DrawingImage
            Dim oDrawingImage As Media.DrawingImage = Nothing
            Using oXmlReader As Xml.XmlReader = Xml.XmlReader.Create(New IO.StringReader(sXaml))
                Dim oResourceDictionary As ResourceDictionary = CType(Markup.XamlReader.Load(oXmlReader), ResourceDictionary)
                For Each sKey As String In oResourceDictionary.Keys
                    Dim oCurrentResource As Object = oResourceDictionary(sKey)
                    If oCurrentResource.GetType.Equals(GetType(Media.DrawingBrush)) Then
                        Dim oDrawingBrush As Media.DrawingBrush = CType(oCurrentResource, Media.DrawingBrush)
                        oDrawingImage = New Media.DrawingImage(oDrawingBrush.Drawing)
                        Exit For
                    End If
                Next
            End Using

            Return oDrawingImage
        End Function
        Public Shared Function DrawingImageToBitmapSource(ByVal oDrawingImage As Media.DrawingImage, ByVal width As Double, ByVal height As Double, ByVal oStretch As Enumerations.StretchEnum, Optional ByVal oBrush As Media.Brush = Nothing, Optional ByVal fRotateAngle As Double = 0) As Media.Imaging.BitmapSource
            ' converts the supplied drawingimage to a bitmapsource
            ' determine scale factors
            Dim oBounds As Rect = oDrawingImage.Drawing.Bounds
            Dim fScaleX As Double = 1
            Dim fScaleY As Double = 1
            If oStretch = Enumerations.StretchEnum.Fill Then
                fScaleX = width / oBounds.Width
                fScaleY = height / oBounds.Height
            Else
                fScaleX = Math.Min(width / oBounds.Width, height / oBounds.Height)
                fScaleY = fScaleX
            End If

            ' draw on drawing context
            Dim oDrawingVisual As New Media.DrawingVisual
            Using oDrawingContext As Media.DrawingContext = oDrawingVisual.RenderOpen
                ' rotate
                If fRotateAngle <> 0 Then
                    oDrawingContext.PushTransform(New Media.RotateTransform(fRotateAngle, oBounds.Width * fScaleX / 2, oBounds.Height * fScaleY / 2))
                End If

                oDrawingContext.PushTransform(New Media.ScaleTransform(fScaleX, fScaleY))
                oDrawingContext.PushTransform(New Media.TranslateTransform(-oBounds.X, -oBounds.Y))

                ' draw background
                If Not IsNothing(oBrush) Then
                    oDrawingContext.DrawRectangle(oBrush, Nothing, oBounds)
                End If
                oDrawingContext.DrawDrawing(oDrawingImage.Drawing)
            End Using

            Dim oRenderTargetBitmap As New Media.Imaging.RenderTargetBitmap(oBounds.Width * fScaleX, oBounds.Height * fScaleY, ScreenResolution096, ScreenResolution096, Media.PixelFormats.Pbgra32)
            oRenderTargetBitmap.Render(oDrawingVisual)
            Return oRenderTargetBitmap
        End Function
#End Region
    End Class
    Public Class PDFHelper
        Public Const BarcodeCharLimit As Integer = 45
        Private Const CornerstoneSeperation As Single = 0.05
        Private Const MinPageContentWidth As Single = 8
        Private Const MinPageContentHeight As Single = 8
        Private Const BorderMarginLeft As Single = 3.5
        Private Const BorderMarginRight As Single = 2
        Private Const BorderMarginTop As Single = 2.5
        Private Const BorderMarginBottom As Single = 3
        Private Const BorderMarginBarcode As Single = 2
        Private Const LineSeparation As Single = 0.25
        Public Shared DefaultPageOrientation As PageOrientation = PageOrientation.Portrait
        Public Shared DefaultPageSize As PageSize = PageSize.A4
        Public Shared Cornerstone As New List(Of Tuple(Of Integer, Integer, XPoint))
        Public Shared PageDictionary As New Dictionary(Of PageOrientation, List(Of PageSize))
        Public Shared PageLimitLeft As XUnit = XUnit.Zero
        Public Shared PageLimitRight As XUnit = XUnit.Zero
        Public Shared PageLimitTop As XUnit = XUnit.Zero
        Public Shared PageLimitBottom As XUnit = XUnit.Zero
        Public Shared PageLimitWidth As XUnit = XUnit.Zero
        Public Shared PageLimitHeight As XUnit = XUnit.Zero
        Public Shared PageFullWidth As XUnit = XUnit.Zero
        Public Shared PageFullHeight As XUnit = XUnit.Zero
        Public Shared PageLimitBottomExBarcode As XUnit = XUnit.Zero
        Public Shared BarcodeBoundaries As Rect
        Public Shared BlockSpacer As XUnit = XUnit.FromCentimeter(0.25)
        Public Shared BlockWidthLimit As XUnit = XUnit.FromCentimeter(7.5)
        Public Shared BlockWidth As XUnit = XUnit.Zero
        Public Shared BlockHeight As XUnit = XUnit.FromCentimeter(0.375)
        Public Shared XMargin As New XUnit(SpacingSmall * 72 / RenderResolution300)
        Public Shared ArielFont As XFont = Nothing
        Public Shared ArielFontDescription As XFont = Nothing
        Public Shared PageBlockWidth As Integer = 0
        Public Shared SpacerBitmapSourceLine As Tuple(Of Media.Imaging.BitmapSource, XUnit, XUnit)
        Public Shared SpacerBitmapSourceNoLine As Tuple(Of Media.Imaging.BitmapSource, XUnit, XUnit)
        Public Shared LoremIpsum As List(Of String)
        Public Shared LoremIpsumWords As List(Of String)
        Private Shared CornerstoneImageStore As New Dictionary(Of Tuple(Of Integer, Integer, Integer, Boolean), System.Drawing.Bitmap)
        Private Shared CheckMatrixStore As New Dictionary(Of Tuple(Of Integer, Enumerations.TabletType), Emgu.CV.Matrix(Of Byte))
        Private Shared ChoiceMatrixStore As New Dictionary(Of Tuple(Of String, Integer, Enumerations.TabletType, Boolean), Emgu.CV.Matrix(Of Byte))

        Public Shared Function Initialise() As Boolean
            Using oPDFDocument As New Pdf.PdfDocument
                Dim oPage As Pdf.PdfPage = oPDFDocument.AddPage()
                Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                    If oXGraphics.PageUnit <> XGraphicsUnit.Point Then
                        Return False
                    End If
                End Using

                ' get list of allowable page sizes
                Dim oOrientations As PageOrientation() = [Enum].GetValues(GetType(PageOrientation))
                Dim oPageSizes As PageSize() = [Enum].GetValues(GetType(PageSize))

                Dim oMinPageXWidth As XUnit = XUnit.FromCentimeter(MinPageContentWidth + BorderMarginLeft + BorderMarginRight)
                Dim oMinPageXHeight As XUnit = XUnit.FromCentimeter(MinPageContentHeight + BorderMarginTop + BorderMarginBottom)

                For Each oOrientation In oOrientations
                    PageDictionary.Add(oOrientation, New List(Of PageSize))
                    For Each oPageSize In oPageSizes
                        If oPageSize <> PageSize.Undefined Then
                            With oPage
                                .Size = oPageSize
                                .Orientation = PageOrientation.Portrait
                                Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                                    If oXGraphics.PageSize.Width >= oMinPageXWidth.Point And oXGraphics.PageSize.Height >= oMinPageXHeight.Point Then
                                        PageDictionary(oOrientation).Add(oPageSize)
                                    End If
                                End Using
                            End With
                        End If
                    Next
                Next

                ' set lorem ipsum
                LoremIpsum = New List(Of String) From {"Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                                                   "Vivamus vitae ipsum facilisis orci cursus imperdiet.",
                                                   "Morbi fermentum viverra sapien, et pretium libero accumsan in.",
                                                   "Nam magna arcu, egestas quis fringilla ut, euismod nec augue.",
                                                   "Sed quis est vehicula, aliquam arcu sit amet, bibendum arcu.",
                                                   "Suspendisse ac augue vitae ex lobortis lobortis quis in lectus.",
                                                   "Sed pellentesque justo quis leo hendrerit dignissim.",
                                                   "Proin interdum tellus varius metus sollicitudin luctus.",
                                                   "Nam at leo quis urna efficitur tincidunt at vel ligula.",
                                                   "Nullam ipsum nibh, ultrices blandit auctor ut, accumsan ut quam.",
                                                   "Morbi ut diam et quam viverra rutrum dictum non est.",
                                                   "Vivamus sed purus eget nulla semper tempor quis quis metus.",
                                                   "Lorem ipsum dolor sit amet, consectetur adipiscing elit."}

                LoremIpsumWords = New List(Of String)
                Dim oWordList As List(Of String) = String.Join(" ", LoremIpsum).Split(" ").ToList
                For Each sWord As String In oWordList
                    LoremIpsumWords.Add(Converter.AlphaNumericOnly(sWord, False).ToLower)
                Next

                ' gets Arial font for the tablet
                Const fFontSize As Double = 10
                Const fTabletHeight As Double = 0.5
                Const fDescriptionFontScale As Double = 0.75
                Dim fSingleBlockWidth As Double = BlockHeight.Point * 2
                Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
                Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
                Dim fScaledFontSize As Double = fFontSize * 0.8 * (fSingleBlockWidth * fTabletHeight) / oTestFont.GetHeight
                ArielFont = New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)
                ArielFontDescription = New XFont(FontArial, fScaledFontSize * fDescriptionFontScale, XFontStyle.Regular, oFontOptions)

                Return True
            End Using
        End Function
        Public Shared Sub ResetPage(ByVal oPageOrientation As PageOrientation, ByVal oPageSize As PageSize)
            ' sets cornerstone locations based on page orientation and size
            Const DummyData As String = "Dummy Data"
            Using oPDFDocument As New Pdf.PdfDocument
                Dim oPage As Pdf.PdfPage = oPDFDocument.AddPage()
                With oPage
                    .Size = oPageSize
                    .Orientation = oPageOrientation
                    Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                        ' initialises cornerstones
                        Cornerstone.Clear()
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(0, 4, New XPoint(XUnit.FromCentimeter(6).Point, XUnit.FromCentimeter(1.5).Point)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(1, 5, New XPoint(oXGraphics.PageSize.Width - XUnit.FromCentimeter(4.5).Point, XUnit.FromCentimeter(1.5).Point)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(2, 6, New XPoint(oXGraphics.PageSize.Width - XUnit.FromCentimeter(1).Point, oXGraphics.PageSize.Height / 3)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(3, 7, New XPoint(oXGraphics.PageSize.Width - XUnit.FromCentimeter(1).Point, oXGraphics.PageSize.Height * 2 / 3)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(4, 0, New XPoint(oXGraphics.PageSize.Width - XUnit.FromCentimeter(4.5).Point, oXGraphics.PageSize.Height - XUnit.FromCentimeter(1.5).Point)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(5, 1, New XPoint(XUnit.FromCentimeter(6).Point, oXGraphics.PageSize.Height - XUnit.FromCentimeter(1.5).Point)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(6, 2, New XPoint(XUnit.FromCentimeter(2.5).Point, oXGraphics.PageSize.Height * 2 / 3)))
                        Cornerstone.Add(New Tuple(Of Integer, Integer, XPoint)(7, 3, New XPoint(XUnit.FromCentimeter(2.5).Point, oXGraphics.PageSize.Height / 3)))

                        ' page limit constants are in cm from the left edge
                        PageLimitLeft = XUnit.FromCentimeter(BorderMarginLeft).Point
                        PageLimitRight = oXGraphics.PageSize.Width - XUnit.FromCentimeter(BorderMarginRight).Point
                        PageLimitTop = XUnit.FromCentimeter(BorderMarginTop).Point
                        PageLimitBottom = oXGraphics.PageSize.Height - XUnit.FromCentimeter(BorderMarginBottom).Point
                        PageLimitWidth = XUnit.FromPoint(PageLimitRight.Point - PageLimitLeft.Point)
                        PageLimitHeight = XUnit.FromPoint(PageLimitBottom.Point - PageLimitTop.Point)
                        PageFullWidth = XUnit.FromPoint(oXGraphics.PageSize.Width)
                        PageFullHeight = XUnit.FromPoint(oXGraphics.PageSize.Height)
                        PageLimitBottomExBarcode = New XUnit(PageLimitBottom.Point - XUnit.FromCentimeter(BorderMarginBarcode).Point)

                        ' set barcode boundaries
                        Dim oBarcodeImage As Tuple(Of System.Drawing.Bitmap, XUnit, XUnit) = GetBarcodeImage(DummyData)
                        Dim oCenterPoint As New XPoint((oXGraphics.PageSize.Width + XUnit.FromCentimeter(1.5).Point) / 2, oXGraphics.PageSize.Height - XUnit.FromCentimeter(1.5).Point)
                        Dim XHeight As XUnit = oBarcodeImage.Item3
                        Dim XTop As New XUnit(oCenterPoint.Y - oBarcodeImage.Item3.Point / 2)

                        Dim fCornerstoneMargin As Double = (SpacingLarge + SpacingSmall) * 72 / RenderResolution300
                        Dim XLeftFullWidth As New XUnit(XUnit.FromCentimeter(6).Point + fCornerstoneMargin)
                        Dim XRightFullWidth As New XUnit(oXGraphics.PageSize.Width - XUnit.FromCentimeter(4.5).Point - fCornerstoneMargin)
                        BarcodeBoundaries = New Rect(XLeftFullWidth.Point, XTop.Point, XRightFullWidth.Point - XLeftFullWidth.Point, XHeight.Point)

                        ' set number of blocks for the page width
                        ' a block is 7.5 cm wide with 0.25 cm spacer
                        PageBlockWidth = 1
                        Dim fCurrentContentWidth As Double = PageLimitWidth.Point - BlockWidthLimit.Point - (BlockSpacer.Point + BlockWidthLimit.Point)
                        Do Until fCurrentContentWidth <= 0
                            PageBlockWidth += 1
                            fCurrentContentWidth -= BlockSpacer.Point + BlockWidthLimit.Point
                        Loop

                        BlockWidth = New XUnit((PageLimitWidth.Point - ((2 + ((PageBlockWidth - 1) * 3)) * BlockSpacer.Point)) / PageBlockWidth)
                    End Using
                End With
            End Using

            ' set spacer bitmap
            SetSpacerBitmaps()
        End Sub
        Public Shared Function GetPageLimitWidth(ByVal oPageOrientation As PageOrientation, ByVal oPageSize As PageSize) As XUnit
            ' gets the page limit width for the supplied parameters
            Using oPDFDocument As New Pdf.PdfDocument
                Dim oPage As Pdf.PdfPage = oPDFDocument.AddPage()
                With oPage
                    .Size = oPageSize
                    .Orientation = oPageOrientation
                    Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                        ' page limit constants are in cm from the left edge
                        Dim oPageLimitLeft As XUnit = XUnit.FromCentimeter(BorderMarginLeft).Point
                        Dim oPageLimitRight As XUnit = oXGraphics.PageSize.Width - XUnit.FromCentimeter(BorderMarginRight).Point
                        Dim oPageLimitWidth As XUnit = XUnit.FromPoint(PageLimitRight.Point - PageLimitLeft.Point)

                        Return PageLimitWidth
                    End Using
                End With
            End Using
        End Function
        Public Shared Sub SetSpacerBitmaps()
            ' set spacer bitmaps
            SpacerBitmapSourceLine = GetSpacerBitmapSource(True)
            SpacerBitmapSourceNoLine = GetSpacerBitmapSource(False)
        End Sub
        Private Shared Function GetSpacerBitmapSource(ByVal bLine As Boolean) As Tuple(Of Media.Imaging.BitmapSource, XUnit, XUnit)
            ' gets a spacer bitmap for the full form width
            If (Not IsNothing(oSettings)) AndAlso (Not IsNothing(oSettings.RenderResolutionValue)) Then
                Dim XSectionWidth As XUnit = PageLimitWidth.Point
                Dim XSectionHeight As XUnit = BlockSpacer

                Dim oBitmap As New System.Drawing.Bitmap(Math.Ceiling(XSectionWidth.Inch * oSettings.RenderResolutionValue), Math.Ceiling(XSectionHeight.Inch * oSettings.RenderResolutionValue), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                oBitmap.SetResolution(oSettings.RenderResolutionValue, oSettings.RenderResolutionValue)
                Dim XBitmapSize As New XSize(oBitmap.Width * 72 / oSettings.RenderResolutionValue, oBitmap.Height * 72 / oSettings.RenderResolutionValue)

                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed

                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XBitmapSize, XGraphicsUnit.Point)
                        oXGraphics.SmoothingMode = XSmoothingMode.None
                        oXGraphics.DrawRectangle(XBrushes.White, 0, 0, XBitmapSize.Width, XBitmapSize.Height)
                    End Using
                End Using

                If bLine Then
                    Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                        oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                        oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed

                        Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, XBitmapSize, XGraphicsUnit.Point)
                            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                            oXGraphics.DrawLine(New XPen(XColors.Black, 0.25), New XPoint(0, XBitmapSize.Height / 2), New XPoint(XBitmapSize.Width, XBitmapSize.Height / 2))
                        End Using
                    End Using
                End If

                Return New Tuple(Of Media.Imaging.BitmapSource, XUnit, XUnit)(Converter.BitmapToBitmapSourceGrayscale(Converter.BitmapConvertGrayscale(oBitmap)), XSectionWidth, XSectionHeight)
            Else
                Return Nothing
            End If
        End Function
        Public Shared Function LoremIpsumPhrase(ByVal iNumber As Integer, ByVal iCount As Integer, ByVal bCapitalise As Boolean) As String
            ' returns a substring of the lorem ipsum starting at iNumber without punctuation of length iCount
            Dim sPhrase As String = String.Empty
            If iNumber >= 0 And iCount > 0 Then
                sPhrase += LoremIpsumWords(iNumber Mod LoremIpsumWords.Count)
                If bCapitalise Then
                    sPhrase = Char.ToUpper(Left(sPhrase, 1)) + Right(sPhrase, sPhrase.Length - 1)
                End If

                For i = 1 To iCount - 1
                    sPhrase += " " + LoremIpsumWords((iNumber + i) Mod LoremIpsumWords.Count)
                Next
            End If

            Return sPhrase
        End Function
        Public Shared Function LoremIpsumWordList(ByVal iNumber As Integer, ByVal iCount As Integer, ByVal bCapitalise As Boolean, Optional ByVal iMaxLength As Integer = Integer.MaxValue) As List(Of String)
            ' returns a list of iCount words from the lorem ipsum starting at iNumber without punctunation, optionally with a maximum length
            Dim oWordList As New List(Of String)
            If iMaxLength < 3 Then
                iMaxLength = 3
            End If
            If iNumber >= 0 And iCount > 0 Then
                Do Until oWordList.Count >= iCount
                    Dim sWord As String = LoremIpsumWords(iNumber Mod LoremIpsumWords.Count)
                    If sWord.Count <= iMaxLength Then
                        If bCapitalise Then
                            sWord = Char.ToUpper(Left(sWord, 1)) + Right(sWord, sWord.Length - 1)
                        End If
                        oWordList.Add(sWord)
                    End If
                    iNumber += 1
                Loop
            End If

            Return oWordList
        End Function
        Public Shared Function GetCornerstoneImage(ByVal iMargin As Integer, ByVal iHorizontalResolution As Integer, ByVal iVerticalResolution As Integer, ByVal bFillBackground As Boolean) As System.Drawing.Bitmap
            ' draw concentric circles with a border in pixel width
            ' resolution is in DPI
            Dim oFoundList As List(Of Integer) = (From iIndex In Enumerable.Range(0, CornerstoneImageStore.Count) Where CornerstoneImageStore.Keys(iIndex).Item1 = iMargin And CornerstoneImageStore.Keys(iIndex).Item2 = iHorizontalResolution And CornerstoneImageStore.Keys(iIndex).Item3 = iVerticalResolution And CornerstoneImageStore.Keys(iIndex).Item4 = bFillBackground Select iIndex).ToList
            If oFoundList.Count > 0 Then
                Return CornerstoneImageStore.Values(oFoundList.First).Clone
            Else
                Const iCircleNumber As Integer = 3
                Const iStrokeWidth As Integer = 3
                Const iFill As Integer = 2
                Dim Spacing As Single = 0.05 * iStrokeWidth / 3

                Dim fPixelWidth As Single = ((Math.Abs(iCircleNumber) * 2) - 1) * Spacing * iHorizontalResolution / 2.54
                Dim fPixelHeight As Single = ((Math.Abs(iCircleNumber) * 2) - 1) * Spacing * iVerticalResolution / 2.54
                Dim fPixelSeperationHorizontal As Single = Spacing * iHorizontalResolution / 2.54
                Dim fPixelSeperationVertical As Single = Spacing * iVerticalResolution / 2.54

                Dim oTabletImage As New System.Drawing.Bitmap(fPixelWidth + (iMargin * 2), fPixelHeight + (iMargin * 2), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                oTabletImage.SetResolution(iHorizontalResolution, iVerticalResolution)

                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oTabletImage)
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    oGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                    oGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic

                    If bFillBackground Then
                        oGraphics.FillRectangle(System.Drawing.Brushes.White, -1, -1, oTabletImage.Width + 2, oTabletImage.Height + 2)
                    End If

                    Dim oPen As New System.Drawing.Pen(System.Drawing.Color.Black, iStrokeWidth * iVerticalResolution / 300)
                    For i = 1 To Math.Abs(iCircleNumber)
                        oGraphics.DrawEllipse(oPen, iMargin + (fPixelSeperationHorizontal * (i - 1)), iMargin + (fPixelSeperationVertical * (i - 1)), fPixelSeperationHorizontal * (((Math.Abs(iCircleNumber) - i) * 2) + 1), fPixelSeperationVertical * (((Math.Abs(iCircleNumber) - i) * 2) + 1))
                    Next

                    ' fill center
                    If iFill > 0 Then
                        oGraphics.FillEllipse(If(iCircleNumber < 0, System.Drawing.Brushes.White, System.Drawing.Brushes.Black), iMargin + (fPixelSeperationHorizontal * (iFill - 1)), iMargin + (fPixelSeperationVertical * (iFill - 1)), fPixelSeperationHorizontal * (((Math.Abs(iCircleNumber) - iFill) * 2) + 1), fPixelSeperationVertical * (((Math.Abs(iCircleNumber) - iFill) * 2) + 1))
                    End If

                    ' invert circle
                    If iCircleNumber < 0 Then
                        Using oTabletMask As System.Drawing.Bitmap = oTabletImage.Clone
                            Using oMaskGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oTabletMask)
                                oMaskGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                                oMaskGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                                oMaskGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic

                                oMaskGraphics.FillEllipse(System.Drawing.Brushes.Black, iMargin, iMargin, fPixelSeperationHorizontal * (((Math.Abs(iCircleNumber) - 1) * 2) + 1), fPixelSeperationVertical * (((Math.Abs(iCircleNumber) - 1) * 2) + 1))
                            End Using

                            Using oImageMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(oTabletImage)
                                Using oMaskMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(oTabletMask)
                                    Emgu.CV.CvInvoke.MorphologyEx(oMaskMatrix, oMaskMatrix, Emgu.CV.CvEnum.MorphOp.Dilate, Emgu.CV.CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Rectangle, New System.Drawing.Size(3, 3), New System.Drawing.Point(-1, -1)), New System.Drawing.Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Constant, New Emgu.CV.Structure.MCvScalar(Byte.MinValue))
                                    Emgu.CV.CvInvoke.Max(oImageMatrix.SubR(Byte.MaxValue), oMaskMatrix, oImageMatrix)
                                    oTabletImage = Converter.BitmapConvertColour(Converter.MatToBitmap(oImageMatrix.Mat, iHorizontalResolution))
                                    oTabletImage.SetResolution(iHorizontalResolution, iVerticalResolution)
                                End Using
                            End Using
                        End Using
                    End If
                End Using

                CornerstoneImageStore.Add(New Tuple(Of Integer, Integer, Integer, Boolean)(iMargin, iHorizontalResolution, iVerticalResolution, bFillBackground), oTabletImage.Clone)

                Return oTabletImage
            End If
        End Function
        Public Shared Function GetBarcodeImage(ByVal sBarcodeData As String) As Tuple(Of System.Drawing.Bitmap, XUnit, XUnit)
            ' gets barcode image
            Dim oImage As System.Drawing.Bitmap = BarcodeGenerator.Code128Rendering.MakeBarcodeImage(Trim(sBarcodeData))
            oImage.SetResolution(RenderResolution300, RenderResolution300)
            Dim XWidth As New XUnit(oImage.Width * 72 / oImage.HorizontalResolution)
            Dim XHeight As New XUnit(oImage.Height * 72 / oImage.VerticalResolution)

            Return New Tuple(Of System.Drawing.Bitmap, XUnit, XUnit)(oImage, XWidth, XHeight)
        End Function
        Public Shared Function GetBarcode(ByVal sBarcodeData As String, ByVal sBarcodeTopText As String, ByVal sBarcodeBottomText As String, Optional ByVal oBitmapWidth As Double = 11, Optional ByVal oBitmapHeight As Double = 4) As System.Drawing.Bitmap
            ' creates a 11 x 4 cm bitmap to contain the drawn barcode
            ' the barcode data is in the form "Title|Page Number|Subject Name"
            Dim oBackgroundSize As New XSize(XUnit.FromCentimeter(oBitmapWidth).Point, XUnit.FromCentimeter(oBitmapHeight).Point)

            Dim fBitmapWidth As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Width).Inch * RenderResolution300)
            Dim fBitmapHeight As Integer = Math.Ceiling(XUnit.FromPoint(oBackgroundSize.Height).Inch * RenderResolution300)
            Dim oBitmap As New System.Drawing.Bitmap(fBitmapWidth, fBitmapHeight)
            oBitmap.SetResolution(RenderResolution300, RenderResolution300)

            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                oGraphics.PageUnit = System.Drawing.GraphicsUnit.Point
                oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, oBackgroundSize, XGraphicsUnit.Point)
                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                    oXGraphics.DrawRectangle(XBrushes.White, New XRect(oBackgroundSize))
                    Dim oCenterPoint As New XPoint(oXGraphics.PageSize.Width / 2, oXGraphics.PageSize.Height / 2)
                    DrawBarcode(oXGraphics, oCenterPoint, sBarcodeData, sBarcodeTopText, sBarcodeBottomText)
                End Using
            End Using

            Return oBitmap
        End Function
        Public Shared Sub DrawBarcode(ByRef oPage As Pdf.PdfPage, ByVal sBarcodeData As String, ByVal sBarcodeTopText As String, ByVal sBarcodeBottomText As String)
            ' draws barcode on PDF page
            Using oXGraphics As XGraphics = XGraphics.FromPdfPage(oPage)
                If oXGraphics.PageUnit = XGraphicsUnit.Point Then
                    oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                    Dim oCenterPoint As New XPoint((oXGraphics.PageSize.Width + XUnit.FromCentimeter(1.5).Point) / 2, oXGraphics.PageSize.Height - XUnit.FromCentimeter(1.5).Point)
                    DrawBarcode(oXGraphics, oCenterPoint, sBarcodeData, sBarcodeTopText, sBarcodeBottomText)
                End If
            End Using
        End Sub
        Private Shared Sub DrawBarcode(ByVal oXGraphics As XGraphics, ByVal oCenterPoint As XPoint, ByVal sBarcodeData As String, ByVal sBarcodeTopText As String, ByVal sBarcodeBottomText As String)
            ' draws barcode on XGraphics
            Dim oBarcodeImage As Tuple(Of System.Drawing.Bitmap, XUnit, XUnit) = GetBarcodeImage(sBarcodeData)

            Dim XWidth As XUnit = oBarcodeImage.Item2
            Dim XHeight As XUnit = oBarcodeImage.Item3
            Dim XLeft As New XUnit(oCenterPoint.X - oBarcodeImage.Item2.Point / 2)
            Dim XTop As New XUnit(oCenterPoint.Y - oBarcodeImage.Item3.Point / 2)
            oXGraphics.DrawImage(XImage.FromGdiPlusImage(oBarcodeImage.Item1), XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point)

            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oArielFont As New XFont(FontArial, 8, XFontStyle.Regular, oFontOptions)
            Dim oStringFormat As New XStringFormat
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            Dim oBottomPoint As New XPoint(oCenterPoint.X, oCenterPoint.Y - oBarcodeImage.Item3.Point / 2)
            oXGraphics.DrawString(Left(Trim(sBarcodeTopText), BarcodeCharLimit), oArielFont, XBrushes.Black, GetTextXRect(Left(Trim(sBarcodeTopText), BarcodeCharLimit), oArielFont, oXGraphics, oBottomPoint, Enumerations.CenterAlignment.Bottom), oStringFormat)

            Dim oTopPoint As New XPoint(oCenterPoint.X, oCenterPoint.Y + oBarcodeImage.Item3.Point / 2)
            oXGraphics.DrawString(Left(Trim(sBarcodeBottomText), BarcodeCharLimit), oArielFont, XBrushes.Black, GetTextXRect(Left(Trim(sBarcodeBottomText), BarcodeCharLimit), oArielFont, oXGraphics, oTopPoint, Enumerations.CenterAlignment.Top), oStringFormat)
        End Sub
        Public Shared Function ProcessBarcode(ByVal oCommonScanner As CommonScanner, ByVal oExtractedMatrix As Emgu.CV.Matrix(Of Byte), ByVal oBarcodeDataList As List(Of String)) As Tuple(Of String, String, String, String)
            ' extracts the barcode from the image
            ' the first three items of the return tuple indicate the three fields on the original bar code, the fourth item is the append text
            Dim oProcessedBarcode As Tuple(Of String, String, String, String) = Nothing
            Using oBarcodeImage As Emgu.CV.Matrix(Of Byte) = oExtractedMatrix.Clone
                Emgu.CV.CvInvoke.Normalize(oBarcodeImage, oBarcodeImage, 0, Byte.MaxValue, Emgu.CV.CvEnum.NormType.MinMax)
                MedianBarcode(oBarcodeImage)
                Using oBarcodeBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oBarcodeImage.Mat, ScreenResolution096)
                    Dim sReturnMessage As String = oCommonScanner.ProcessBarcode(oBarcodeBitmap)
                    If sReturnMessage <> String.Empty Then
                        oProcessedBarcode = ProcessBarcodeText(sReturnMessage)

                        ' run a check to see if part of supplied list, and if not present to try spire detection
                        If (Not IsNothing(oProcessedBarcode)) AndAlso (Not BarcodeInList(oBarcodeDataList, oProcessedBarcode)) Then
                            oProcessedBarcode = Nothing
                        End If
                    End If

                    If IsNothing(oProcessedBarcode) Then
                        Dim oBarcodeData As New ArrayList
                        BarcodeScanner.ScanPage(oBarcodeData, oBarcodeBitmap, 1, BarcodeScanner.ScanDirection.Vertical, BarcodeScanner.BarcodeType.Code128)

                        If oBarcodeData.Count = 3 Then
                            oProcessedBarcode = ProcessBarcodeText(String.Join("|", oBarcodeData.ToArray))
                        End If
                    End If
                End Using
            End Using

            Return oProcessedBarcode
        End Function
        Private Shared Function BarcodeInList(ByVal oBarcodeDataList As List(Of String), ByVal oProcessedBarcode As Tuple(Of String, String, String, String)) As Boolean
            ' checks through barcode list to see if it contains the supplied barcode
            ' if the supplied barcode is a template, then just check the first two items
            Dim bFound As Boolean = False

            If Left(oProcessedBarcode.Item3, 1) = "[" And Right(oProcessedBarcode.Item3, 1) = "]" Then
                For i = 0 To oBarcodeDataList.Count - 1
                    Dim oCurrentProcessedBarcode As Tuple(Of String, String, String, String) = ProcessBarcodeText(oBarcodeDataList(i))
                    If oCurrentProcessedBarcode.Item1 = oProcessedBarcode.Item1 And oCurrentProcessedBarcode.Item2 = oProcessedBarcode.Item2 Then
                        bFound = True
                        Exit For
                    End If
                Next
            Else
                For i = 0 To oBarcodeDataList.Count - 1
                    Dim oCurrentProcessedBarcode As Tuple(Of String, String, String, String) = ProcessBarcodeText(oBarcodeDataList(i))
                    If oCurrentProcessedBarcode.Item1 = oProcessedBarcode.Item1 And oCurrentProcessedBarcode.Item2 = oProcessedBarcode.Item2 And oCurrentProcessedBarcode.Item3 = oProcessedBarcode.Item3 And oCurrentProcessedBarcode.Item4 = oProcessedBarcode.Item4 Then
                        bFound = True
                        Exit For
                    End If
                Next
            End If

            Return bFound
        End Function
        Private Shared Function ProcessBarcodeText(ByVal sBarcodeData As String) As Tuple(Of String, String, String, String)
            ' separates barcode text
            Dim oData As String() = sBarcodeData.Split(SeparatorChar)
            If oData.Count = 3 Then
                Return New Tuple(Of String, String, String, String)(oData(0), oData(1), oData(2), String.Empty)
            ElseIf oData.Count = 4 Then
                Return New Tuple(Of String, String, String, String)(oData(0), oData(1), oData(2), oData(3))
            Else
                Return Nothing
            End If
        End Function
        Private Shared Sub MedianBarcode(ByRef oMatrix As Emgu.CV.Matrix(Of Byte))
            ' processes the barcode to get the median value of each column
            Const iMargin As Integer = 6
            Using oSubMatrix As Emgu.CV.Matrix(Of Byte) = oMatrix.GetSubRect(New System.Drawing.Rectangle(0, iMargin, oMatrix.Width, oMatrix.Height - (iMargin * 2)))
                Dim TaskDelegate As Action(Of Object) = Sub(oParamIn As Object)
                                                            Dim oParam As Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object) = oParamIn

                                                            Dim oMatrixLocal As Emgu.CV.Matrix(Of Byte) = CType(oParam.Item4, Emgu.CV.Matrix(Of Byte))
                                                            Dim localheight As Integer = CType(oParam.Item5, Integer)

                                                            Dim oByteArray(localheight - 1) As Byte
                                                            Dim iMedianLow As Integer = Math.Floor((localheight - 1) / 2)
                                                            Dim iMedianHigh As Integer = Math.Ceiling((localheight - 1) / 2)
                                                            For y = 0 To oParam.Item2 - 1
                                                                Dim localY As Integer = y + oParam.Item1
                                                                Using oCol As Emgu.CV.Matrix(Of Byte) = oMatrixLocal.GetCol(localY)
                                                                    oCol.Mat.CopyTo(oByteArray)
                                                                    Array.Sort(oByteArray)
                                                                    Dim iMedian As Byte = (CShort(oByteArray(iMedianLow)) + CShort(oByteArray(iMedianHigh))) / 2
                                                                    oCol.SetValue(iMedian)
                                                                End Using
                                                            Next
                                                        End Sub
                CommonFunctions.ParallelRun(oSubMatrix.Width, oSubMatrix, oSubMatrix.Height, Nothing, Nothing, TaskDelegate)
            End Using
        End Sub
        Public Shared Function GetTextXRect(ByVal sText As String, ByVal oFont As XFont, ByVal oXGraphics As XGraphics, ByVal oCenter As XPoint, ByVal oCenterAlignment As Enumerations.CenterAlignment) As XRect
            Dim oSize As XSize = oXGraphics.MeasureString(sText, oFont)
            Select Case oCenterAlignment
                Case Enumerations.CenterAlignment.Center
                    Return New XRect(oCenter.X - oSize.Width / 2, oCenter.Y - oSize.Height / 2, oSize.Width, oSize.Height)
                Case Enumerations.CenterAlignment.Top
                    Return New XRect(oCenter.X - oSize.Width / 2, oCenter.Y, oSize.Width, oSize.Height)
                Case Enumerations.CenterAlignment.Bottom
                    Return New XRect(oCenter.X - oSize.Width / 2, oCenter.Y - oSize.Height, oSize.Width, oSize.Height)
                Case Enumerations.CenterAlignment.Left
                    Return New XRect(oCenter.X, oCenter.Y - oSize.Height / 2, oSize.Width, oSize.Height)
                Case Enumerations.CenterAlignment.Right
                    Return New XRect(oCenter.X - oSize.Width, oCenter.Y - oSize.Height / 2, oSize.Width, oSize.Height)
            End Select
        End Function
        Public Shared Function GetScaledFontSize(ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal oXSingleBoxHeight As XUnit) As Double
            ' gets the font size scaled to the container size
            Const fFontSize As Double = 10
            Dim oXGraphics As XGraphics = XGraphics.CreateMeasureContext(New XSize(oXWidth.Point, oXHeight.Point), XGraphicsUnit.Point, XPageDirection.Downwards)
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * oXSingleBoxHeight.Point / oTestFont.GetHeight
            Return fScaledFontSize
        End Function
        Public Shared Function DrawFieldTablets(ByRef oXGraphics As XGraphics, ByVal XDisplacement As XPoint, ByVal iTabletCount As Integer, ByVal iTabletStart As Integer, ByVal iTabletGroups As Integer, ByVal oTabletContent As Enumerations.TabletContentEnum, ByRef oField As FieldDocumentStore.Field, Optional ByVal oTabletImages As List(Of XImage) = Nothing, Optional ByVal oTabletDescriptionTop As List(Of Tuple(Of Rect, String)) = Nothing, Optional ByVal oTabletDescriptionBottom As List(Of Tuple(Of Rect, String)) = Nothing, Optional ByVal iRow As Integer = 0) As Rect
            ' draws tablets for fields
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim oReturnRect As New Rect
            Dim fSingleBlockWidth As Double = BlockHeight.Point * 2

            ' draw letters
            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            ' draws tablets
            Dim iTabletColumns As Integer = Math.Ceiling(iTabletCount / iTabletGroups)
            For i = 0 To iTabletCount - 1
                Dim sContentText As String = String.Empty
                Select Case oTabletContent
                    Case Enumerations.TabletContentEnum.Number
                        sContentText = (i + If(iTabletStart = -2, 0, iTabletStart + 1)).ToString
                    Case Enumerations.TabletContentEnum.Letter
                        sContentText = Converter.ConvertNumberToLetter(i + Math.Max(iTabletStart, 0), True)
                End Select

                Dim fDisplacementLeft As Double = XDisplacement.X + fSingleBlockWidth / 2 + ((i Mod iTabletColumns) * fSingleBlockWidth)
                Dim fDisplacementTop As Double = XDisplacement.Y + If(iTabletStart < -1, fSingleBlockWidth / 2, 0) + Math.Floor(i / iTabletColumns) * fSingleBlockWidth * 3 / 2
                Dim XCheckEmptyBitmapSize As XSize = DrawTablet(oXGraphics, New XPoint(fDisplacementLeft, fDisplacementTop), sContentText, ArielFont, oStringFormat, If(IsNothing(oTabletImages) OrElse i >= oTabletImages.Count, Nothing, oTabletImages(i)))

                Dim XLeft As New XUnit(fDisplacementLeft - (XCheckEmptyBitmapSize.Width / 2))
                Dim XTop As New XUnit(fDisplacementTop - (XCheckEmptyBitmapSize.Height / 2))
                Dim XWidth As New XUnit(XCheckEmptyBitmapSize.Width)
                Dim XHeight As New XUnit(XCheckEmptyBitmapSize.Height)
                Dim oImageRect As New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point)

                If Not IsNothing(oField) Then
                    If IsNothing(oTabletImages) Then
                        oField.AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oImageRect, Nothing, String.Empty, iRow, i, False, 0, New Tuple(Of Single)(0)))
                    End If
                End If

                If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                    oReturnRect = oImageRect
                Else
                    oReturnRect.Union(oImageRect)
                End If
            Next

            If iTabletStart >= -1 Then
                ' draws descriptions
                If (Not IsNothing(oTabletDescriptionTop)) AndAlso oTabletDescriptionTop.Count > 0 Then
                    For i = 0 To Math.Min(iTabletCount, oTabletDescriptionTop.Count) - 1
                        Dim fCenterX As Double = XDisplacement.X + (fSingleBlockWidth / 2) + ((i Mod iTabletColumns) * fSingleBlockWidth)
                        Dim fCenterY As Double = XDisplacement.Y + (Math.Floor(i / iTabletColumns) * fSingleBlockWidth * 3 / 2) - (fSingleBlockWidth / 2)
                        oXGraphics.DrawString(oTabletDescriptionTop(i).Item2, ArielFontDescription, XBrushes.Black, fCenterX, fCenterY, oStringFormat)

                        Dim oXRect As XSize = oXGraphics.MeasureString(oTabletDescriptionTop(i).Item2, ArielFontDescription, oStringFormat)
                        Dim oDescriptionRect As New Rect(fCenterX - oXRect.Width / 2, fCenterY - oXRect.Height / 2, oXRect.Width, oXRect.Height)

                        If Not IsNothing(oField) Then
                            If (Not IsNothing(oField.TabletDescriptionTop)) AndAlso oField.TabletDescriptionTop.Count = Math.Min(iTabletCount, oTabletDescriptionTop.Count) AndAlso oField.TabletDescriptionTop(i).Item1.IsEmpty Then
                                oField.TabletDescriptionTop(i) = New Tuple(Of Rect, String)(oDescriptionRect, oField.TabletDescriptionTop(i).Item2)
                            End If
                        End If

                        If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                            oReturnRect = oDescriptionRect
                        Else
                            oReturnRect.Union(oDescriptionRect)
                        End If
                    Next
                End If

                If (Not IsNothing(oTabletDescriptionBottom)) AndAlso oTabletDescriptionBottom.Count > 0 Then
                    For i = 0 To Math.Min(iTabletCount, oTabletDescriptionBottom.Count) - 1
                        Dim fCenterX As Double = XDisplacement.X + (fSingleBlockWidth / 2) + ((i Mod iTabletColumns) * fSingleBlockWidth)
                        Dim fCenterY As Double = XDisplacement.Y + (Math.Floor(i / iTabletColumns) * fSingleBlockWidth * 3 / 2) + (fSingleBlockWidth / 2)
                        oXGraphics.DrawString(oTabletDescriptionBottom(i).Item2, ArielFontDescription, XBrushes.Black, fCenterX, fCenterY, oStringFormat)

                        Dim oXRect As XSize = oXGraphics.MeasureString(oTabletDescriptionBottom(i).Item2, ArielFontDescription, oStringFormat)
                        Dim oDescriptionRect As New Rect(fCenterX - oXRect.Width / 2, fCenterY - oXRect.Height / 2, oXRect.Width, oXRect.Height)

                        If Not IsNothing(oField) Then
                            If (Not IsNothing(oField.TabletDescriptionBottom)) AndAlso oField.TabletDescriptionBottom.Count = Math.Min(iTabletCount, oTabletDescriptionBottom.Count) AndAlso oField.TabletDescriptionBottom(i).Item1.IsEmpty Then
                                oField.TabletDescriptionBottom(i) = New Tuple(Of Rect, String)(oDescriptionRect, oField.TabletDescriptionBottom(i).Item2)
                            End If
                        End If

                        If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                            oReturnRect = oDescriptionRect
                        Else
                            oReturnRect.Union(oDescriptionRect)
                        End If
                    Next
                End If
            End If

            oXGraphics.SmoothingMode = oSmoothingMode
            If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                Return New Rect
            Else
                Return oReturnRect
            End If
        End Function
        Public Shared Function DrawFieldTabletsVertical(ByRef oXGraphics As XGraphics, ByVal XBlockWidth As XUnit, ByVal oXWidthLimit As XUnit, ByVal XDisplacement As XPoint, ByVal iTabletCount As Integer, ByVal iTabletStart As Integer, ByVal iTabletGroups As Integer, ByVal oTabletContent As Enumerations.TabletContentEnum, ByRef oField As FieldDocumentStore.Field, ByVal oAlignment As Enumerations.AlignmentEnum, Optional ByVal oTabletImages As List(Of XImage) = Nothing, Optional ByVal oTabletDescriptionTop As List(Of Tuple(Of Rect, String)) = Nothing, Optional ByVal oTabletDescriptionBottom As List(Of Tuple(Of Rect, String)) = Nothing, Optional ByVal oTabletDisplacements As List(Of Tuple(Of XPoint, XRect, XRect)) = Nothing) As Tuple(Of Rect, Double)
            ' draws tablets for fields, vertical configuration
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim oReturnRect As New Rect
            Dim fSingleBlockWidth As Double = BlockHeight.Point * 2

            ' draw letters
            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            ' draws tablets
            Dim oTabletRect As New Rect
            Dim iTabletRows As Integer = Math.Ceiling(iTabletCount / iTabletGroups)
            Dim XCheckEmptyBitmapSize As XSize = Nothing
            For i = 0 To iTabletCount - 1
                Dim sContentText As String = String.Empty
                Select Case oTabletContent
                    Case Enumerations.TabletContentEnum.Number
                        sContentText = (i + If(iTabletStart = -2, 0, iTabletStart + 1)).ToString
                    Case Enumerations.TabletContentEnum.Letter
                        sContentText = Converter.ConvertNumberToLetter(i + Math.Max(iTabletStart, 0), True)
                End Select

                Dim fDisplacementLeft As Double = 0
                Dim fDisplacementTop As Double = 0

                ' use supplied tablet displacements if given
                If (Not IsNothing(oTabletDisplacements)) AndAlso oTabletDisplacements.Count = iTabletCount Then
                    fDisplacementLeft = oTabletDisplacements(i).Item1.X
                    fDisplacementTop = oTabletDisplacements(i).Item1.Y
                Else
                    fDisplacementTop = XDisplacement.Y + (fSingleBlockWidth / 2) + ((i Mod iTabletRows) * fSingleBlockWidth)
                    Select Case oAlignment
                        Case Enumerations.AlignmentEnum.Left
                            fDisplacementLeft = XDisplacement.X + (Math.Floor(i / iTabletRows) * (XBlockWidth.Point / iTabletGroups)) + (fSingleBlockWidth / 2)
                        Case Enumerations.AlignmentEnum.Center
                            fDisplacementLeft = XDisplacement.X + (Math.Floor(i / iTabletRows) * (XBlockWidth.Point / iTabletGroups)) + (XBlockWidth.Point / 2) - ((iTabletGroups - 1) * XBlockWidth.Point / (2 * iTabletGroups))
                        Case Enumerations.AlignmentEnum.Right
                            fDisplacementLeft = XDisplacement.X + ((Math.Floor(i / iTabletRows) + 1) * XBlockWidth.Point / iTabletGroups) - (fSingleBlockWidth / 2)
                    End Select
                End If

                XCheckEmptyBitmapSize = DrawTablet(oXGraphics, New XPoint(fDisplacementLeft, fDisplacementTop), sContentText, ArielFont, oStringFormat, If(IsNothing(oTabletImages) OrElse i >= oTabletImages.Count, Nothing, oTabletImages(i)))

                Dim XLeft As New XUnit(fDisplacementLeft - (XCheckEmptyBitmapSize.Width / 2))
                Dim XTop As XUnit = New XUnit(fDisplacementTop - (XCheckEmptyBitmapSize.Height / 2))
                Dim XWidth As New XUnit(XCheckEmptyBitmapSize.Width)
                Dim XHeight As New XUnit(XCheckEmptyBitmapSize.Height)
                Dim oImageRect As New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point)

                If Not IsNothing(oField) Then
                    If IsNothing(oTabletImages) Then
                        oField.AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oImageRect, Nothing, String.Empty, 0, i, False, 0, New Tuple(Of Single)(0)))
                    End If
                End If

                If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                    oReturnRect = oImageRect
                Else
                    oReturnRect.Union(oImageRect)
                End If
                If oTabletRect.Width = 0 Or oTabletRect.Height = 0 Then
                    oTabletRect = oImageRect
                Else
                    oTabletRect.Union(oImageRect)
                End If
            Next

            ' draws descriptions
            Const fDescriptionScale As Single = 0.75
            Dim XLeftDescriptionDisplacementX As XUnit = Nothing
            Dim XRightDescriptionDisplacementX As XUnit = Nothing
            Dim XDescriptionDisplacementY As XUnit = Nothing
            Dim XDescriptionWidth As XUnit = Nothing
            Dim XDescriptionHeight As XUnit = Nothing
            If (Not IsNothing(oTabletDescriptionTop)) AndAlso oTabletDescriptionTop.Count > 0 Then
                For i = 0 To Math.Min(iTabletCount, oTabletDescriptionTop.Count) - 1
                    If Trim(oTabletDescriptionTop(i).Item2) <> String.Empty Then
                        ' use supplied tablet displacements if given
                        If (Not IsNothing(oTabletDisplacements)) AndAlso oTabletDisplacements.Count = iTabletCount AndAlso (Not (oTabletDisplacements(i).Item2.X = 0 And oTabletDisplacements(i).Item2.Y = 0)) Then
                            XDescriptionDisplacementY = New XUnit(oTabletDisplacements(i).Item2.Y)
                            XLeftDescriptionDisplacementX = New XUnit(oTabletDisplacements(i).Item2.X)
                            XDescriptionHeight = New XUnit(oTabletDisplacements(i).Item2.Height)
                            XDescriptionWidth = New XUnit(oTabletDisplacements(i).Item2.Width)
                        Else
                            XDescriptionHeight = New XUnit(fSingleBlockWidth * fDescriptionScale)
                            Dim fDisplacementLeft As Double = 0
                            Dim fDisplacementTop As Double = 0

                            Select Case oAlignment
                                Case Enumerations.AlignmentEnum.Left
                                Case Enumerations.AlignmentEnum.Center
                                    XDescriptionWidth = New XUnit((XBlockWidth.Point / (2 * iTabletGroups)) - (fSingleBlockWidth / 2))
                                    fDisplacementLeft = XDisplacement.X + (Math.Floor(i / iTabletRows) * (XBlockWidth.Point / iTabletGroups)) + (XBlockWidth.Point / 2) - ((iTabletGroups - 1) * XBlockWidth.Point / (2 * iTabletGroups))
                                Case Enumerations.AlignmentEnum.Right
                                    XDescriptionWidth = New XUnit((XBlockWidth.Point / iTabletGroups) - fSingleBlockWidth)
                                    fDisplacementLeft = XDisplacement.X + ((Math.Floor(i / iTabletRows) + 1) * XBlockWidth.Point / iTabletGroups) - (fSingleBlockWidth / 2)
                            End Select
                            fDisplacementTop = XDisplacement.Y + (fSingleBlockWidth / 2) + ((i Mod iTabletRows) * fSingleBlockWidth)
                            XDescriptionDisplacementY = New XUnit(fDisplacementTop - (fDescriptionScale * fSingleBlockWidth / 2))
                            XLeftDescriptionDisplacementX = New XUnit(fDisplacementLeft - (fSingleBlockWidth / 2) - XDescriptionWidth.Point)
                        End If

                        Dim oDescriptionRect As Rect = DrawText(oXGraphics, XDescriptionWidth, oXWidthLimit, XDescriptionHeight, New XPoint(XLeftDescriptionDisplacementX, XDescriptionDisplacementY), oTabletDescriptionTop(i).Item2, False, False, False, MigraDoc.DocumentObjectModel.ParagraphAlignment.Right, MigraDoc.DocumentObjectModel.Colors.Black, 2)

                        If Not IsNothing(oField) Then
                            If (Not IsNothing(oField.TabletDescriptionTop)) AndAlso oField.TabletDescriptionTop.Count = Math.Min(iTabletCount, oTabletDescriptionTop.Count) AndAlso oField.TabletDescriptionTop(i).Item1.IsEmpty Then
                                oField.TabletDescriptionTop(i) = New Tuple(Of Rect, String)(New Rect(XLeftDescriptionDisplacementX, XDescriptionDisplacementY, XDescriptionWidth, XDescriptionHeight), oField.TabletDescriptionTop(i).Item2)
                            End If
                        End If

                        If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                            oReturnRect = oDescriptionRect
                        Else
                            oReturnRect.Union(oDescriptionRect)
                        End If
                    End If
                Next
            End If

            If (Not IsNothing(oTabletDescriptionBottom)) AndAlso oTabletDescriptionBottom.Count > 0 Then
                For i = 0 To Math.Min(iTabletCount, oTabletDescriptionBottom.Count) - 1
                    If Trim(oTabletDescriptionBottom(i).Item2) <> String.Empty Then
                        ' use supplied tablet displacements if given
                        If (Not IsNothing(oTabletDisplacements)) AndAlso oTabletDisplacements.Count = iTabletCount AndAlso (Not (oTabletDisplacements(i).Item3.X = 0 And oTabletDisplacements(i).Item3.Y = 0)) Then
                            XDescriptionDisplacementY = New XUnit(oTabletDisplacements(i).Item3.Y)
                            XRightDescriptionDisplacementX = New XUnit(oTabletDisplacements(i).Item3.X)
                            XDescriptionHeight = New XUnit(oTabletDisplacements(i).Item3.Height)
                            XDescriptionWidth = New XUnit(oTabletDisplacements(i).Item3.Width)
                        Else
                            XDescriptionHeight = New XUnit(fSingleBlockWidth * fDescriptionScale)
                            Dim fDisplacementLeft As Double = 0
                            Dim fDisplacementTop As Double = 0

                            Select Case oAlignment
                                Case Enumerations.AlignmentEnum.Left
                                    XDescriptionWidth = New XUnit((XBlockWidth.Point / iTabletGroups) - fSingleBlockWidth)
                                    fDisplacementLeft = XDisplacement.X + (Math.Floor(i / iTabletRows) * (XBlockWidth.Point / iTabletGroups)) + (fSingleBlockWidth / 2)
                                Case Enumerations.AlignmentEnum.Center
                                    XDescriptionWidth = New XUnit((XBlockWidth.Point / (2 * iTabletGroups)) - (fSingleBlockWidth / 2))
                                    fDisplacementLeft = XDisplacement.X + (Math.Floor(i / iTabletRows) * (XBlockWidth.Point / iTabletGroups)) + (XBlockWidth.Point / 2) - ((iTabletGroups - 1) * XBlockWidth.Point / (2 * iTabletGroups))
                                Case Enumerations.AlignmentEnum.Right
                            End Select
                            fDisplacementTop = XDisplacement.Y + (fSingleBlockWidth / 2) + ((i Mod iTabletRows) * fSingleBlockWidth)
                            XDescriptionDisplacementY = New XUnit(fDisplacementTop - (fDescriptionScale * fSingleBlockWidth / 2))
                            XRightDescriptionDisplacementX = New XUnit(fDisplacementLeft + (fSingleBlockWidth / 2))
                        End If

                        Dim oDescriptionRect As Rect = DrawText(oXGraphics, XDescriptionWidth, oXWidthLimit, XDescriptionHeight, New XPoint(XRightDescriptionDisplacementX, XDescriptionDisplacementY), oTabletDescriptionBottom(i).Item2, False, False, False, MigraDoc.DocumentObjectModel.ParagraphAlignment.Left, MigraDoc.DocumentObjectModel.Colors.Black, 2)

                        If Not IsNothing(oField) Then
                            If (Not IsNothing(oField.TabletDescriptionBottom)) AndAlso oField.TabletDescriptionBottom.Count = Math.Min(iTabletCount, oTabletDescriptionBottom.Count) AndAlso oField.TabletDescriptionBottom(i).Item1.IsEmpty Then
                                oField.TabletDescriptionBottom(i) = New Tuple(Of Rect, String)(New Rect(XRightDescriptionDisplacementX, XDescriptionDisplacementY, XDescriptionWidth, XDescriptionHeight), oField.TabletDescriptionBottom(i).Item2)
                            End If
                        End If

                        If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                            oReturnRect = oDescriptionRect
                        Else
                            oReturnRect.Union(oDescriptionRect)
                        End If
                    End If
                Next
            End If

            oXGraphics.SmoothingMode = oSmoothingMode
            If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                Return New Tuple(Of Rect, Double)(New Rect, 0)
            Else
                Return New Tuple(Of Rect, Double)(oReturnRect, XDescriptionWidth.Point)
            End If
        End Function
        Public Shared Function DrawFieldTabletsMCQ(ByRef oXGraphics As XGraphics, ByVal XImageWidth As XUnit, ByVal XImageHeight As XUnit, ByVal XDisplacement As XPoint, ByVal oGridRect As Int32Rect, ByVal iGridHeight As Integer, ByVal iFontSizeMultiplier As Integer, ByVal oOverflowRows As List(Of Integer), ByRef oField As FieldDocumentStore.Field, Optional ByVal oTabletImages As List(Of XImage) = Nothing, Optional ByVal oTabletDisplacements As List(Of Tuple(Of XRect, Integer, Integer, XRect, List(Of ElementStruc))) = Nothing) As Rect
            ' draws tablets for MCQ field
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim fSingleBlockWidth As Double = BlockHeight.Point * 2
            Dim XColumnWidth As New XUnit((XImageWidth.Point - (oField.TabletGroups * fSingleBlockWidth)) / oField.TabletGroups)
            Dim oReturnRect As New Rect

            Dim oMCQHeight As Tuple(Of Double, Integer, List(Of Tuple(Of Integer, Integer))) = GetFieldTabletsMCQHeight(XImageWidth, iFontSizeMultiplier, oField.TabletGroups, oField.TabletDescriptionMCQ)
            Dim oLayoutList As List(Of Tuple(Of Integer, Integer)) = oMCQHeight.Item3

            ' draw letters
            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Center
            oStringFormat.LineAlignment = XLineAlignment.Center

            Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Subject, MigraDoc.DocumentObjectModel.Colors.DarkViolet)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Field, MigraDoc.DocumentObjectModel.Colors.Green)

            Dim XCheckEmptyBitmapSize As XSize = Nothing
            Dim iRowCount As Integer = Math.Ceiling(oField.TabletDescriptionMCQ.Count / oField.TabletGroups)
            For iColumn = 0 To oField.TabletGroups - 1
                Dim iCurrentRowHeight As Integer = 0
                For iRow = 0 To iRowCount - 1
                    Dim iIndex As Integer = (iColumn * iRowCount) + iRow
                    If iIndex < oField.TabletDescriptionMCQ.Count Then
                        Dim sContentText As String = String.Empty
                        Select Case oField.TabletContent
                            Case Enumerations.TabletContentEnum.Number
                                sContentText = (iIndex + If(oField.TabletStart = -2, 0, oField.TabletStart + 1)).ToString
                            Case Enumerations.TabletContentEnum.Letter
                                sContentText = Converter.ConvertNumberToLetter(iIndex + Math.Max(oField.TabletStart, 0), True)
                        End Select

                        Dim fDisplacementLeft As Double = 0
                        Dim fDisplacementTop As Double = 0

                        ' use supplied tablet displacements if given
                        If (Not IsNothing(oTabletDisplacements)) Then
                            fDisplacementLeft = oTabletDisplacements(iIndex).Item1.X + oTabletDisplacements(iIndex).Item1.Width / 2
                            fDisplacementTop = oTabletDisplacements(iIndex).Item1.Y + oTabletDisplacements(iIndex).Item1.Height / 2
                        Else
                            fDisplacementLeft = XDisplacement.X + (fSingleBlockWidth / 2) + ((XColumnWidth.Point + fSingleBlockWidth) * iColumn)
                            fDisplacementTop = XDisplacement.Y + (oLayoutList(iIndex).Item1 * iFontSizeMultiplier * BlockHeight.Point / 2) + (iCurrentRowHeight * BlockHeight.Point) + ((iCurrentRowHeight + 1) * BlockHeight.Point * (LineSeparation / 2))
                        End If

                        XCheckEmptyBitmapSize = DrawTablet(oXGraphics, New XPoint(fDisplacementLeft, fDisplacementTop), sContentText, ArielFont, oStringFormat, If(IsNothing(oTabletImages) OrElse iIndex >= oTabletImages.Count, Nothing, oTabletImages(iIndex)))
                        Dim XLeft As New XUnit(fDisplacementLeft - (XCheckEmptyBitmapSize.Width / 2))
                        Dim XTop As XUnit = New XUnit(fDisplacementTop - (XCheckEmptyBitmapSize.Height / 2))
                        Dim XWidth As New XUnit(XCheckEmptyBitmapSize.Width)
                        Dim XHeight As New XUnit(XCheckEmptyBitmapSize.Height)
                        Dim oImageRect As New Rect(XLeft.Point, XTop.Point, XWidth.Point, XHeight.Point)

                        If IsNothing(oTabletImages) Then
                            oField.AddImage(New Tuple(Of Rect, Guid, String, Integer, Integer, Boolean, Single, Tuple(Of Single))(oImageRect, Nothing, String.Empty, 0, iIndex, False, 0, New Tuple(Of Single)(0)))
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
                            oAdjustedWidth = oTabletDisplacements(iIndex).Item4.Width
                            oAdjustedHeight = oTabletDisplacements(iIndex).Item4.Height
                            fDisplacementLeft = oTabletDisplacements(iIndex).Item4.X
                            fDisplacementTop = oTabletDisplacements(iIndex).Item4.Y
                            iLinesInBox = oTabletDisplacements(iIndex).Item2
                            iLinesPerLine = oTabletDisplacements(iIndex).Item3
                            oElements = oTabletDisplacements(iIndex).Item5
                        Else
                            oAdjustedWidth = XColumnWidth
                            oAdjustedHeight = XImageHeight
                            fDisplacementLeft = XDisplacement.X + fSingleBlockWidth + ((XColumnWidth.Point + fSingleBlockWidth) * iColumn)
                            fDisplacementTop = XDisplacement.Y + ((oLayoutList(iIndex).Item1 - oLayoutList(iIndex).Item2) * iFontSizeMultiplier * BlockHeight.Point / 2) + (iCurrentRowHeight * BlockHeight.Point) + ((iCurrentRowHeight + 1) * BlockHeight.Point * (LineSeparation / 2))
                            iLinesInBox = oGridRect.Height
                            iLinesPerLine = iFontSizeMultiplier
                            oElements = oField.TabletDescriptionMCQ(iIndex).Item4
                        End If

                        DrawFieldText(oXGraphics, oAdjustedWidth, oAdjustedHeight, New XPoint(fDisplacementLeft, fDisplacementTop), iLinesInBox, iLinesPerLine, oElements, -1, MigraDoc.DocumentObjectModel.ParagraphAlignment.Left, oFontColourDictionary)

                        If IsNothing(oTabletDisplacements) Then
                            If (Not IsNothing(oField.TabletDescriptionMCQ)) AndAlso oField.TabletDescriptionMCQ(iIndex).Item1.IsEmpty Then
                                oField.TabletDescriptionMCQ(iIndex) = New Tuple(Of Rect, Integer, Integer, List(Of ElementStruc))(New Rect(fDisplacementLeft, fDisplacementTop, XColumnWidth, XImageHeight), oGridRect.Height, iFontSizeMultiplier, oField.TabletDescriptionMCQ(iIndex).Item4)
                            End If

                            Dim oTextRect As New Rect(fDisplacementLeft, fDisplacementTop, XColumnWidth.Point, oLayoutList(iIndex).Item2 * iFontSizeMultiplier * BlockHeight.Point)
                            If oReturnRect.Width = 0 Or oReturnRect.Height = 0 Then
                                oReturnRect = oTextRect
                            Else
                                oReturnRect.Union(oTextRect)
                            End If
                        End If

                        iCurrentRowHeight += oLayoutList(iIndex).Item1
                    End If
                Next
            Next

            oXGraphics.SmoothingMode = oSmoothingMode
            Return oReturnRect
        End Function
        Public Shared Function GetFieldTabletsMCQHeight(ByVal XImageWidth As XUnit, ByVal iFontSizeMultiplier As Integer, ByVal oTabletGroups As Integer, ByVal oTabletDescriptionMCQ As List(Of Tuple(Of Rect, Integer, Integer, List(Of ElementStruc)))) As Tuple(Of Double, Integer, List(Of Tuple(Of Integer, Integer)))
            ' gets the height of the field
            ' give a line spacer between items
            Dim fSingleBlockWidth As Double = BlockHeight.Point * 2
            Dim XColumnWidth As New XUnit((XImageWidth.Point - (oTabletGroups * fSingleBlockWidth)) / oTabletGroups)

            Dim oLayoutList As New List(Of Tuple(Of Integer, Integer))
            For Each oTabletElements As Tuple(Of Rect, Integer, Integer, List(Of ElementStruc)) In oTabletDescriptionMCQ
                Dim oLayout As LayoutClass = GetLayout(XColumnWidth, BlockHeight.Point, 1, iFontSizeMultiplier, oTabletElements.Item4, -1).Item1
                oLayoutList.Add(New Tuple(Of Integer, Integer)(Math.Max(oLayout.LineCount, 2), oLayout.LineCount))
            Next

            Dim iRowCount As Integer = Math.Ceiling(oTabletDescriptionMCQ.Count / oTabletGroups)
            Dim oRowList As New List(Of Integer)
            For iColumn = 0 To oTabletGroups - 1
                Dim iCurrentColumnSum As Integer = 0
                For iRow = 0 To iRowCount - 1
                    Dim iIndex As Integer = (iColumn * iRowCount) + iRow
                    If iIndex < oTabletDescriptionMCQ.Count Then
                        iCurrentColumnSum += oLayoutList(iIndex).Item1
                    End If
                Next
                oRowList.Add(iCurrentColumnSum)
            Next

            Dim fHeight As Double = (oRowList.Max + (Math.Max(iRowCount - 1, 0) * LineSeparation)) * BlockHeight.Point
            Dim iSetRowCount As Integer = Math.Ceiling(fHeight / BlockHeight.Point)

            Return New Tuple(Of Double, Integer, List(Of Tuple(Of Integer, Integer)))(fHeight, iSetRowCount, oLayoutList)
        End Function
        Public Shared Function GetLayout(ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal iLinesInBox As Integer, ByVal iLinesPerLine As Integer, ByVal oElements As List(Of ElementStruc), ByVal iSelectedElement As Integer) As Tuple(Of LayoutClass, Double)
            ' initialise layout class
            Const fFontSize As Double = 10
            Dim oXGraphics As XGraphics = XGraphics.CreateMeasureContext(New XSize(oXWidth.Point, oXHeight.Point), XGraphicsUnit.Point, XPageDirection.Downwards)
            Dim oLayoutClass As New LayoutClass
            Dim oSingleBoxSize As New XSize(oXWidth.Point, oXHeight.Point * iLinesPerLine / iLinesInBox)
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * oSingleBoxSize.Height / oTestFont.GetHeight

            Dim oStringFormat As New XStringFormat()
            oStringFormat.Alignment = XStringAlignment.Near
            oStringFormat.LineAlignment = XLineAlignment.Center
            oStringFormat.FormatFlags = XStringFormatFlags.MeasureTrailingSpaces

            Dim fCurrentWidth As Double = 0
            For i = 0 To oElements.Count - 1
                Dim oElement As ElementStruc = oElements(i)
                If IsNothing(oElement.Text) OrElse oElement.Text.Length > 0 Then
                    Dim oElementTextStruc As New TextLayoutClass.TextStruc(oElement.Text)
                    For j = 0 To oElementTextStruc.LineCount - 1
                        For k = 0 To oElementTextStruc.WordCount(j) - 1
                            Dim sWord As String = oElementTextStruc.Word(j, k)

                            Dim oFontStyle As XFontStyle = XFontStyle.Regular
                            If oElement.FontBold Then
                                oFontStyle = oFontStyle Or XFontStyle.Bold
                            End If
                            If oElement.FontItalic Then
                                oFontStyle = oFontStyle Or XFontStyle.Italic
                            End If
                            If oElement.FontUnderline Then
                                oFontStyle = oFontStyle Or XFontStyle.Underline
                            End If

                            Dim oArielFont As New XFont(FontArial, fScaledFontSize, oFontStyle, oFontOptions)
                            Dim fSpaceWidth As Double = oXGraphics.MeasureString(" ", oArielFont, oStringFormat).Width
                            Dim oMeasure As XSize = oXGraphics.MeasureString(sWord, oArielFont, oStringFormat)
                            Dim fWordWidth As Double = oMeasure.Width
                            Dim fWordHeight As Double = oMeasure.Height

                            If fCurrentWidth = 0 Then
                                ' first word
                                oLayoutClass.AddWord(sWord, fWordWidth, fWordHeight, If(i = iSelectedElement, True, False), oElement)
                                If fWordWidth >= oXWidth.Point Then
                                    oLayoutClass.AddLine()
                                    fCurrentWidth = 0
                                Else
                                    fCurrentWidth += fWordWidth
                                End If
                            Else
                                ' subsequent words in line
                                If fCurrentWidth + fSpaceWidth + fWordWidth >= oXWidth.Point Then
                                    oLayoutClass.AddLine()

                                    oLayoutClass.AddWord(sWord, fWordWidth, fWordHeight, If(i = iSelectedElement, True, False), oElement)
                                    If fWordWidth >= oXWidth.Point Then
                                        oLayoutClass.AddLine()
                                        fCurrentWidth = 0
                                    Else
                                        fCurrentWidth = fWordWidth
                                    End If
                                Else
                                    oLayoutClass.AddWord(sWord, fWordWidth, fWordHeight, If(i = iSelectedElement, True, False), oElement)
                                    fCurrentWidth += fSpaceWidth + fWordWidth
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            Return New Tuple(Of LayoutClass, Double)(oLayoutClass, fScaledFontSize)
        End Function
        Public Shared Sub DrawFieldBackground(ByRef oXGraphics As XGraphics, ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal XDisplacement As XPoint, ByVal oBrush As XBrush, Optional ByVal iGridWidth As Integer = 0, Optional ByVal iGridHeight As Integer = 0, Optional ByVal oGridRect As Int32Rect = Nothing)
            ' draws background for fields
            Dim oCurrentSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.None

            If IsNothing(oGridRect) Or iGridWidth = 0 Or iGridHeight = 0 Then
                oXGraphics.DrawRectangle(oBrush, XDisplacement.X, XDisplacement.Y, oXWidth.Point, oXHeight.Point)
            Else
                Dim oDispX As Double = oXWidth.Point * oGridRect.X / iGridWidth
                Dim oDispY As Double = oXHeight.Point * oGridRect.Y / iGridHeight
                Dim oDispWidth As Double = oXWidth.Point * oGridRect.Width / iGridWidth
                Dim oDispHeight As Double = oXHeight.Point * oGridRect.Height / iGridHeight

                oXGraphics.DrawRectangle(oBrush, XDisplacement.X + oDispX, XDisplacement.Y + oDispY, oDispWidth, oDispHeight)
            End If

            oXGraphics.SmoothingMode = oCurrentSmoothingMode
        End Sub
        Public Shared Sub DrawFieldBorder(ByRef oXGraphics As XGraphics, ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal XDisplacement As XPoint, ByVal fBorderWidth As Double, Optional ByVal iGridWidth As Integer = 0, Optional ByVal iGridHeight As Integer = 0, Optional ByVal oGridRect As Int32Rect = Nothing)
            ' draws border for fields
            Dim oXPen As XPen = New XPen(XColors.Black, fBorderWidth)
            If IsNothing(oGridRect) Or iGridWidth = 0 Or iGridHeight = 0 Then
                DrawFieldBorderPen(oXGraphics, oXPen, oXWidth, oXHeight, XDisplacement)
            Else
                Dim oDispX As Double = oXWidth.Point * oGridRect.X / iGridWidth
                Dim oDispY As Double = oXHeight.Point * oGridRect.Y / iGridHeight
                Dim oDispWidth As Double = oXWidth.Point * oGridRect.Width / iGridWidth
                Dim oDispHeight As Double = oXHeight.Point * oGridRect.Height / iGridHeight

                DrawFieldBorderPen(oXGraphics, oXPen, oDispWidth, oDispHeight, New XPoint(XDisplacement.X + oDispX, XDisplacement.Y + oDispY))
            End If
        End Sub
        Public Shared Sub DrawFieldBorderPen(ByRef oXGraphics As XGraphics, ByVal oXPen As XPen, ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal XDisplacement As XPoint)
            ' draws border for fields
            Dim oCurrentSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.None
            oXGraphics.DrawRectangle(oXPen, XDisplacement.X, XDisplacement.Y, oXWidth.Point, oXHeight.Point)
            oXGraphics.SmoothingMode = oCurrentSmoothingMode
        End Sub
        Public Shared Sub DrawFieldNumbering(ByRef oXGraphics As XGraphics, ByVal XDisplacement As XPoint, ByVal iNumber As Integer, ByVal iSubNumber As Integer, ByVal oNumberingType As Enumerations.Numbering, ByVal bBackground As Boolean, ByVal bBorder As Boolean, ByVal oBorderBackground As XBrush)
            ' draw numbering for fields
            Dim XContentWidth As New XUnit(BlockHeight.Point * 2)
            Dim XContentHeight As XUnit = BlockHeight

            If bBackground Then
                DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XDisplacement, System.Drawing.Brushes.LightGray)
            Else
                If bBorder Then
                    DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XDisplacement, XBrushes.White)
                Else
                    DrawFieldBackground(oXGraphics, XContentWidth, XContentHeight, XDisplacement, oBorderBackground)
                End If
            End If
            If bBorder Then
                DrawFieldBorder(oXGraphics, XContentWidth, XContentHeight, XDisplacement, 0.5)
            End If

            Dim sNumbering As String = GetNumbering(iNumber, iSubNumber, oNumberingType)
            Dim oElements As New List(Of ElementStruc)
            oElements.Add(New ElementStruc(sNumbering, ElementStruc.ElementTypeEnum.Text, Enumerations.FontEnum.Bold))
            Dim oFontColourDictionary As New Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color)
            oFontColourDictionary.Add(ElementStruc.ElementTypeEnum.Text, MigraDoc.DocumentObjectModel.Colors.Black)
            DrawFieldText(oXGraphics, XContentWidth, XContentHeight, XDisplacement, 1, 1, oElements, -1, MigraDoc.DocumentObjectModel.ParagraphAlignment.Center, oFontColourDictionary)
        End Sub
        Public Shared Function GetNumbering(ByVal iNumber As Integer, ByVal iSubNumber As Integer, ByVal oNumberingType As Enumerations.Numbering) As String
            ' gets the actual numbering
            Dim sNumbering As String = String.Empty
            Select Case oNumberingType
                Case Enumerations.Numbering.Number
                    sNumbering = (iNumber + 1).ToString + If(iSubNumber >= 0, Converter.ConvertNumberToLetter(iSubNumber, False), String.Empty)
                Case Enumerations.Numbering.LetterSmall
                    sNumbering = Converter.ConvertNumberToLetter(iNumber, False) + If(iSubNumber >= 0, (iSubNumber + 1).ToString, String.Empty)
                Case Enumerations.Numbering.LetterBig
                    sNumbering = Converter.ConvertNumberToLetter(iNumber, True) + If(iSubNumber >= 0, (iSubNumber + 1).ToString, String.Empty)
            End Select
            Return sNumbering
        End Function
        Public Shared Sub DrawFieldText(ByRef oXGraphics As XGraphics, ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal XDisplacement As XPoint, ByVal iLinesInBox As Integer, ByVal iLinesPerLine As Integer, ByVal oElements As List(Of ElementStruc), ByVal iSelectedElement As Integer, ByVal oAlignment As MigraDoc.DocumentObjectModel.ParagraphAlignment, ByVal oFontColourDictionary As Dictionary(Of ElementStruc.ElementTypeEnum, MigraDoc.DocumentObjectModel.Color), Optional ByVal oLayout As Tuple(Of LayoutClass, Double) = Nothing)
            ' draws text with fields differentiated by colour
            ' this contains both static text as well as dynamic fields which are indicated by a different brush colour
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim sCombinedText As String = Trim(String.Concat((From oElement As ElementStruc In oElements Select oElement.Text).ToArray))
            If sCombinedText <> String.Empty Then
                If IsNothing(oLayout) Then
                    oLayout = PDFHelper.GetLayout(oXWidth, oXHeight, iLinesInBox, iLinesPerLine, oElements, iSelectedElement)
                End If
                Dim oLayoutClass As LayoutClass = oLayout.Item1
                Dim fScaledFontSize As Double = oLayout.Item2

                Dim oMigraDocument As New MigraDoc.DocumentObjectModel.Document
                Dim oMigraSection As MigraDoc.DocumentObjectModel.Section = oMigraDocument.AddSection

                Dim oMigraParagraph As MigraDoc.DocumentObjectModel.Paragraph = oMigraSection.AddParagraph
                oMigraParagraph.Format.Borders.Distance = 0
                oMigraParagraph.Format.Font.Name = FontArial
                oMigraParagraph.Format.Alignment = oAlignment
                oMigraParagraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(fScaledFontSize)
                oMigraParagraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Transparent

                Dim sFontName As String = FontArial
                Dim oColor As MigraDoc.DocumentObjectModel.Color = Nothing
                Dim oUnderline As MigraDoc.DocumentObjectModel.Underline = Nothing
                Dim bBold As Boolean = False
                Dim bItalic As Boolean = False
                For i = 0 To oLayoutClass.LineCount - 1
                    For j = 0 To oLayoutClass.WordCount(i) - 1
                        Dim oWord As Tuple(Of String, Double, Double, Boolean, ElementStruc) = oLayoutClass.Word(i, j)

                        If oWord.Item4 Then
                            oColor = MigraDoc.DocumentObjectModel.Colors.Red
                            oUnderline = MigraDoc.DocumentObjectModel.Underline.Dotted
                        Else
                            oColor = oFontColourDictionary(oWord.Item5.ElementType)
                            If oWord.Item5.FontUnderline Then
                                oUnderline = MigraDoc.DocumentObjectModel.Underline.Single
                            Else
                                oUnderline = MigraDoc.DocumentObjectModel.Underline.None
                            End If
                        End If
                        bBold = oWord.Item5.FontBold
                        bItalic = oWord.Item5.FontItalic
                        If j = 0 Then
                            If i = 0 Then
                                oMigraParagraph.Add(GetFormattedText(oWord.Item1, sFontName, fScaledFontSize, oColor, oUnderline, bBold, bItalic, False))
                            Else
                                oMigraParagraph.Add(GetFormattedText(oWord.Item1, sFontName, fScaledFontSize, oColor, oUnderline, bBold, bItalic, True))
                            End If
                        Else
                            Dim oPreviousWord As Tuple(Of String, Double, Double, Boolean, ElementStruc) = oLayoutClass.Word(i, j - 1)
                            oMigraParagraph.Add(GetFormattedText(" ", sFontName, fScaledFontSize, oColor, If((Not oWord.Item4) And (Not oPreviousWord.Item4) And oWord.Item5.FontUnderline And oPreviousWord.Item5.FontUnderline, MigraDoc.DocumentObjectModel.Underline.Single, MigraDoc.DocumentObjectModel.Underline.None), bBold, bItalic, False))
                            oMigraParagraph.Add(GetFormattedText(oWord.Item1, sFontName, fScaledFontSize, oColor, oUnderline, bBold, bItalic, False))
                        End If
                    Next
                Next

                Dim oMigraRenderer As New MigraDoc.Rendering.DocumentRenderer(oMigraDocument)
                oMigraRenderer.PrepareDocument()
                oMigraRenderer.RenderObject(oXGraphics, XDisplacement.X, XDisplacement.Y, oXWidth, oMigraParagraph)
            End If
            oXGraphics.SmoothingMode = oSmoothingMode
        End Sub
        Private Shared Function GetFormattedText(ByVal sText As String, ByVal sFontName As String, ByVal fFontSize As Double, ByVal oColor As MigraDoc.DocumentObjectModel.Color, ByVal oUnderline As MigraDoc.DocumentObjectModel.Underline, ByVal bBold As Boolean, ByVal bItalic As Boolean, ByVal bLeadingLineBreak As Boolean) As MigraDoc.DocumentObjectModel.FormattedText
            ' returns formatted migradoc text
            Dim oFormattedText As New MigraDoc.DocumentObjectModel.FormattedText()
            With oFormattedText
                .FontName = sFontName
                .Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(fFontSize)
                .Color = oColor
                .Underline = oUnderline
                .Bold = bBold
                .Italic = bItalic

                If bLeadingLineBreak Then
                    .AddLineBreak()
                End If
                .AddText(sText)
            End With
            Return oFormattedText
        End Function
        Public Shared Function DrawTablet(ByRef oXGraphics As XGraphics, ByVal XCenterPoint As XPoint, ByVal sTabletContent As String, ByVal oFont As XFont, ByVal oStringFormat As XStringFormat, Optional ByVal oTabletImage As XImage = Nothing) As XSize
            ' draws a single tablet with the center point given
            ' the content is limited to two characters only
            ' draws tablet
            Dim iResolution As Integer = MarkResolution600
            Dim fSingleBlockWidth As Double = BlockHeight.Point * 2
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            If IsNothing(oTabletImage) Then
                oXGraphics.SmoothingMode = XSmoothingMode.None

                Dim XChoiceBitmapSize As XSize = Nothing
                Using oChoiceMatrix As Emgu.CV.Matrix(Of Byte) = GetChoiceMatrix(Left(sTabletContent, 2), iResolution, Enumerations.TabletType.Empty, fSingleBlockWidth)
                    Using oChoiceBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oChoiceMatrix.Mat, iResolution)
                        Using oXChoiceImage As XImage = XImage.FromGdiPlusImage(oChoiceBitmap)
                            XChoiceBitmapSize = New XSize(oXChoiceImage.PointWidth, oXChoiceImage.PointHeight)
                            oXGraphics.DrawImage(oXChoiceImage, XCenterPoint.X - XChoiceBitmapSize.Width / 2, XCenterPoint.Y - XChoiceBitmapSize.Height / 2)
                        End Using
                    End Using
                End Using

                oXGraphics.SmoothingMode = oSmoothingMode

                Using oTabletEmpty As Emgu.CV.Matrix(Of Byte) = GetCheckMatrix(iResolution, Enumerations.TabletType.Empty, fSingleBlockWidth)
                    Return New XSize(XUnit.FromInch(oTabletEmpty.Width / iResolution).Point, XUnit.FromInch(oTabletEmpty.Height / iResolution).Point)
                End Using
            Else
                oXGraphics.SmoothingMode = XSmoothingMode.None
                Dim XTabletImageSize As New XSize(oTabletImage.PointWidth, oTabletImage.PointHeight)

                oXGraphics.DrawImage(oTabletImage, XCenterPoint.X - oTabletImage.PointWidth / 2, XCenterPoint.Y - oTabletImage.PointHeight / 2)

                oXGraphics.SmoothingMode = oSmoothingMode
                Return XTabletImageSize
            End If
        End Function
        Public Shared Function GetTabletFont(ByVal fResolution As Single) As XFont
            ' gets Arial font for the tablet
            Dim fSingleBlockWidth As Double = PDFHelper.BlockHeight.Point * 2
            Const fTabletHeight As Double = 0.5

            Const fFontSize As Double = 10
            Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
            Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
            Dim fScaledFontSize As Double = fFontSize * 0.8 * (fSingleBlockWidth * fTabletHeight) * fResolution / (oTestFont.GetHeight * 72)
            Dim oArielFont As New XFont(FontArial, fScaledFontSize, XFontStyle.Regular, oFontOptions)

            Return oArielFont
        End Function
        Public Shared Function DrawText(ByRef oXGraphics As XGraphics, ByVal oXWidth As XUnit, ByVal oXWidthLimit As XUnit, ByVal oXHeight As XUnit, ByVal XDisplacement As XPoint, ByVal sText As String, ByVal bBold As Boolean, ByVal bItalic As Boolean, ByVal bUnderline As Boolean, ByVal oAlignment As MigraDoc.DocumentObjectModel.ParagraphAlignment, ByVal oFontColour As MigraDoc.DocumentObjectModel.Color, Optional ByVal iLines As Integer = 1) As Rect
            ' draws text with fields differentiated by colour
            Dim oSmoothingMode As XSmoothingMode = oXGraphics.SmoothingMode
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

            Dim oLayout As Tuple(Of TextLayoutClass, Double, XSize) = GetTextLayout(New XUnit(Math.Max(oXWidth.Point, oXWidthLimit.Point)), oXHeight, iLines, 1, sText, bBold, bItalic, bUnderline)
            If Not IsNothing(oLayout) Then
                Dim oLayoutClass As TextLayoutClass = oLayout.Item1
                Dim fScaledFontSize As Double = oLayout.Item2
                Dim fDisplacementShift As Double = (iLines - oLayoutClass.LineCount) * (oXHeight.Point / iLines) / 2

                Dim oMigraDocument As New MigraDoc.DocumentObjectModel.Document
                Dim oMigraSection As MigraDoc.DocumentObjectModel.Section = oMigraDocument.AddSection

                Dim oMigraParagraph As MigraDoc.DocumentObjectModel.Paragraph = oMigraSection.AddParagraph
                oMigraParagraph.Format.Borders.Distance = 0
                oMigraParagraph.Format.Font.Name = FontArial
                oMigraParagraph.Format.Alignment = oAlignment
                oMigraParagraph.Format.Font.Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(fScaledFontSize)
                oMigraParagraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Transparent
                Dim oFormattedText As New MigraDoc.DocumentObjectModel.FormattedText()
                With oFormattedText
                    .FontName = FontArial
                    .Size = MigraDoc.DocumentObjectModel.Unit.FromPoint(fScaledFontSize)
                    .Color = oFontColour
                    If bUnderline Then
                        .Underline = MigraDoc.DocumentObjectModel.Underline.Single
                    Else
                        .Underline = MigraDoc.DocumentObjectModel.Underline.None
                    End If
                    .Bold = bBold
                    .Italic = bItalic
                    .AddText(sText)
                End With
                oMigraParagraph.Add(oFormattedText)

                Dim fShift As Double = Math.Max(oXWidth.Point, oXWidthLimit.Point) - oXWidth.Point
                Dim oMigraRenderer As New MigraDoc.Rendering.DocumentRenderer(oMigraDocument)
                oMigraRenderer.PrepareDocument()
                oMigraRenderer.RenderObject(oXGraphics, XDisplacement.X - fShift, XDisplacement.Y + fDisplacementShift, New XUnit(oXWidth.Point + fShift), oMigraParagraph)

                Select Case oAlignment
                    Case MigraDoc.DocumentObjectModel.ParagraphAlignment.Left, MigraDoc.DocumentObjectModel.ParagraphAlignment.Justify
                        Return New Rect(XDisplacement.X, XDisplacement.Y + fDisplacementShift, oLayout.Item3.Width, oLayout.Item3.Height)
                    Case MigraDoc.DocumentObjectModel.ParagraphAlignment.Right
                        Return New Rect(XDisplacement.X + (oXWidth.Point - oLayout.Item3.Width), XDisplacement.Y + fDisplacementShift, oLayout.Item3.Width, oLayout.Item3.Height)
                    Case Else
                        ' center alignment
                        Return New Rect(XDisplacement.X + (oXWidth.Point - oLayout.Item3.Width) / 2, XDisplacement.Y + fDisplacementShift, oLayout.Item3.Width, oLayout.Item3.Height)
                End Select
            End If
            oXGraphics.SmoothingMode = oSmoothingMode
        End Function
        Public Shared Function GetTextLayout(ByVal oXWidth As XUnit, ByVal oXHeight As XUnit, ByVal iLinesInBox As Integer, ByVal iLinesPerLine As Integer, ByVal sText As String, ByVal bBold As Boolean, ByVal bItalic As Boolean, ByVal bUnderline As Boolean) As Tuple(Of TextLayoutClass, Double, XSize)
            ' initialise layout class
            If oXWidth.Point > 0 And oXHeight.Point > 0 Then
                Const fFontSize As Double = 10
                Dim oXGraphics As XGraphics = XGraphics.CreateMeasureContext(New XSize(oXWidth.Point, oXHeight.Point), XGraphicsUnit.Point, XPageDirection.Downwards)
                Dim oLayoutClass As New TextLayoutClass
                Dim oSingleBoxSize As New XSize(oXWidth.Point, oXHeight.Point * iLinesPerLine / iLinesInBox)
                Dim oFontOptions As New XPdfFontOptions(Pdf.PdfFontEncoding.Unicode)
                Dim oTestFont As New XFont(FontArial, fFontSize, XFontStyle.Regular, oFontOptions)
                Dim fScaledFontSize As Double = fFontSize * oSingleBoxSize.Height / oTestFont.GetHeight

                Dim oStringFormat As New XStringFormat()
                oStringFormat.Alignment = XStringAlignment.Near
                oStringFormat.LineAlignment = XLineAlignment.Center
                oStringFormat.FormatFlags = XStringFormatFlags.MeasureTrailingSpaces

                Dim fTextWidth As Double = 0
                Dim fTextHeight As Double = 0
                Dim fCurrentWidth As Double = 0
                Dim oTextStruc As New TextLayoutClass.TextStruc(sText)
                For j = 0 To oTextStruc.LineCount - 1
                    For k = 0 To oTextStruc.WordCount(j) - 1
                        Dim sWord As String = oTextStruc.Word(j, k)

                        Dim oFontStyle As XFontStyle = XFontStyle.Regular
                        If bBold Then
                            oFontStyle = oFontStyle Or XFontStyle.Bold
                        End If
                        If bItalic Then
                            oFontStyle = oFontStyle Or XFontStyle.Italic
                        End If
                        If bUnderline Then
                            oFontStyle = oFontStyle Or XFontStyle.Underline
                        End If

                        Dim oArielFont As New XFont(FontArial, fScaledFontSize, oFontStyle, oFontOptions)
                        Dim fSpaceWidth As Double = oXGraphics.MeasureString(" ", oArielFont, oStringFormat).Width
                        Dim oMeasure As XSize = oXGraphics.MeasureString(sWord, oArielFont, oStringFormat)
                        Dim fWordWidth As Double = oMeasure.Width
                        Dim fWordHeight As Double = oMeasure.Height
                        fTextHeight = Math.Max(fTextHeight, fWordHeight * oTextStruc.LineCount)

                        If fCurrentWidth = 0 Then
                            ' first word
                            oLayoutClass.AddWord(sWord, fWordWidth, fWordHeight)
                            If fWordWidth >= oXWidth.Point Then
                                oLayoutClass.AddLine()
                                fCurrentWidth = 0
                            Else
                                fCurrentWidth += fWordWidth
                                fTextWidth = Math.Max(fTextWidth, fCurrentWidth)
                            End If
                        Else
                            ' subsequent words in line
                            If fCurrentWidth + fSpaceWidth + fWordWidth >= oXWidth.Point Then
                                oLayoutClass.AddLine()

                                oLayoutClass.AddWord(sWord, fWordWidth, fWordHeight)
                                If fWordWidth >= oXWidth.Point Then
                                    oLayoutClass.AddLine()
                                    fCurrentWidth = 0
                                Else
                                    fCurrentWidth = fWordWidth
                                    fTextWidth = Math.Max(fTextWidth, fCurrentWidth)
                                End If
                            Else
                                oLayoutClass.AddWord(sWord, fWordWidth, fWordHeight)
                                fCurrentWidth += fSpaceWidth + fWordWidth
                                fTextWidth = Math.Max(fTextWidth, fCurrentWidth)
                            End If
                        End If
                    Next
                Next

                Return New Tuple(Of TextLayoutClass, Double, XSize)(oLayoutClass, fScaledFontSize, New XSize(fTextWidth, fTextHeight))
            Else
                Return Nothing
            End If
        End Function
        Public Shared Sub DrawStringRotated(ByRef oXGraphics As XGraphics, ByVal XCenterPoint As XPoint, ByVal oFont As XFont, ByVal oStringFormat As XStringFormat, ByVal sText As String, ByVal fAngle As Double)
            ' draws a string that is rotated by fAngle degrees
            Dim oContainer As XGraphicsContainer = oXGraphics.BeginContainer()
            oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
            oXGraphics.RotateAtTransform(fAngle, XCenterPoint)
            oXGraphics.DrawString(sText, oFont, XBrushes.Black, XCenterPoint.X, XCenterPoint.Y, oStringFormat)
            oXGraphics.EndContainer(oContainer)
        End Sub
        Public Shared Function GetCheckMatrix(ByVal iResolution As Integer, ByVal oTabletType As Enumerations.TabletType, Optional ByVal fSingleBlockWidth As Double = 0) As Emgu.CV.Matrix(Of Byte)
            ' gets the bitmap of a single tablet without the text
            Dim oFoundList As List(Of Integer) = (From iIndex In Enumerable.Range(0, CheckMatrixStore.Count) Where CheckMatrixStore.Keys(iIndex).Item1 = iResolution And CheckMatrixStore.Keys(iIndex).Item2 = oTabletType Select iIndex).ToList
            If oFoundList.Count > 0 Then
                Return CheckMatrixStore.Values(oFoundList.First).Clone
            Else
                If fSingleBlockWidth = 0 Then
                    fSingleBlockWidth = BlockHeight.Point * 2
                End If
                Const fTabletHeight As Double = 0.5

                Dim oCheckMatrix As Emgu.CV.Matrix(Of Byte) = Nothing
                Dim sCheckString As String = String.Empty
                Select Case oTabletType
                    Case Enumerations.TabletType.Empty
                        sCheckString = My.Resources.CCMCheckEmpty
                    Case Enumerations.TabletType.Filled
                        sCheckString = My.Resources.CCMCheckFilled
                    Case Enumerations.TabletType.EmptyBlack
                        sCheckString = My.Resources.CCMCheckEmptyBlack
                    Case Enumerations.TabletType.Inner, Enumerations.TabletType.EmptyBlackInner
                        sCheckString = My.Resources.CCMCheckInner
                End Select

                Dim oCheck As Media.DrawingImage = Converter.XamlToDrawingImage(sCheckString)
                oCheckMatrix = Converter.BitmapSourceToMatrix8Bit(Converter.DrawingImageToBitmapSource(oCheck, Double.MaxValue, fSingleBlockWidth * fTabletHeight * iResolution / 72, Enumerations.StretchEnum.Uniform, Media.Brushes.White))

                Dim oReturnMatrix As Emgu.CV.Matrix(Of Byte) = Nothing
                If oTabletType = Enumerations.TabletType.EmptyBlackInner Then
                    oReturnMatrix = oCheckMatrix.SubR(Byte.MaxValue)
                Else
                    oReturnMatrix = oCheckMatrix
                End If

                CheckMatrixStore.Add(New Tuple(Of Integer, Enumerations.TabletType)(iResolution, oTabletType), oReturnMatrix.Clone)

                Return oReturnMatrix
            End If
        End Function
        Public Shared Function GetChoiceMatrix(ByVal sTabletContent As String, ByVal iResolution As Integer, ByVal oTabletType As Enumerations.TabletType, Optional ByVal fSingleBlockWidth As Double = 0, Optional ByVal bMarginBlack As Boolean = False) As Emgu.CV.Matrix(Of Byte)
            ' gets the bitmap of a single tablet
            Dim oFoundList As List(Of Integer) = (From iIndex In Enumerable.Range(0, ChoiceMatrixStore.Count) Where ChoiceMatrixStore.Keys(iIndex).Item1 = Left(sTabletContent, 2) And ChoiceMatrixStore.Keys(iIndex).Item2 = iResolution And ChoiceMatrixStore.Keys(iIndex).Item3 = oTabletType And ChoiceMatrixStore.Keys(iIndex).Item4 = bMarginBlack Select iIndex).ToList
            If oFoundList.Count > 0 Then
                Return ChoiceMatrixStore.Values(oFoundList.First).Clone
            Else
                If fSingleBlockWidth = 0 Then
                    fSingleBlockWidth = BlockHeight.Point * 2
                End If

                Using oCheckMatrix As Emgu.CV.Matrix(Of Byte) = GetCheckMatrix(iResolution, oTabletType, fSingleBlockWidth)
                    Dim iMargin As Integer = SpacingSmall
                    Dim oArielFont As XFont = GetTabletFont(iResolution)

                    Dim oStringFormat As New XStringFormat()
                    oStringFormat.Alignment = XStringAlignment.Center
                    oStringFormat.LineAlignment = XLineAlignment.Center

                    Dim oReturnBitmap As System.Drawing.Bitmap = Nothing
                    Using oBitmap As New System.Drawing.Bitmap(oCheckMatrix.Width + (iMargin * 2), oCheckMatrix.Height + (iMargin * 2), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                        oBitmap.SetResolution(iResolution, iResolution)
                        Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                            oGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                            oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                            If bMarginBlack Then
                                oGraphics.FillRectangle(System.Drawing.Brushes.Black, 0, 0, oBitmap.Width, oBitmap.Height)
                                oGraphics.FillRectangle(System.Drawing.Brushes.White, iMargin, iMargin, oCheckMatrix.Width, oCheckMatrix.Height)
                            Else
                                oGraphics.FillRectangle(System.Drawing.Brushes.White, 0, 0, oBitmap.Width, oBitmap.Height)
                            End If
                            Using oCheckBitmap As System.Drawing.Bitmap = Converter.MatToBitmap(oCheckMatrix.Mat, iResolution)
                                oGraphics.DrawImage(oCheckBitmap, iMargin, iMargin)
                            End Using
                            Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, New XSize(oCheckMatrix.Width * 72 / iResolution, oCheckMatrix.Height * 72 / iResolution), XGraphicsUnit.Point)
                                oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                                Select Case oTabletType
                                    Case Enumerations.TabletType.Empty, Enumerations.TabletType.EmptyBlack, Enumerations.TabletType.EmptyBlackInner
                                        ' draws tablet text
                                        oXGraphics.DrawString(Left(sTabletContent, 2), oArielFont, XBrushes.Black, (oCheckMatrix.Width / 2) + iMargin, (oCheckMatrix.Height / 2) + iMargin, oStringFormat)
                                End Select
                            End Using
                        End Using
                        oReturnBitmap = Converter.BitmapConvertGrayscale(oBitmap)
                    End Using

                    Dim oReturnMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(oReturnBitmap)
                    oReturnBitmap.Dispose()

                    ChoiceMatrixStore.Add(New Tuple(Of String, Integer, Enumerations.TabletType, Boolean)(Left(sTabletContent, 2), iResolution, oTabletType, bMarginBlack), oReturnMatrix.Clone)

                    Return oReturnMatrix
                End Using
            End If
        End Function
        Public Shared Function GetHandwritingMatrix(ByVal sHandwritingContent As String, ByVal fResolution As Single, ByVal fMargin As Single, Optional ByVal fBorderWidth As Double = 0.5) As Tuple(Of Emgu.CV.Matrix(Of Byte), Emgu.CV.Matrix(Of Byte))
            ' gets the bitmap for the border box of a single letter
            Dim fSingleBlockWidth As Double = BlockHeight.Point * 2
            Dim iBitmapWidth As Integer = Math.Ceiling((fSingleBlockWidth * fResolution / 72) + (fMargin * 2))
            Dim iInnerBitmapWidth As Integer = Math.Ceiling((fSingleBlockWidth * fResolution / 72) - (fMargin * 2))
            Dim XSingleBlockDimension As New XUnit(fSingleBlockWidth)

            Dim oReturnBitmap As System.Drawing.Bitmap = Nothing
            Using oBitmap As New System.Drawing.Bitmap(iBitmapWidth, iBitmapWidth, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                oBitmap.SetResolution(fResolution, fResolution)
                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    oGraphics.FillRectangle(System.Drawing.Brushes.White, 0, 0, oBitmap.Width, oBitmap.Height)
                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, New XSize(oBitmap.Width * 72 / fResolution, oBitmap.Height * 72 / fResolution), XGraphicsUnit.Point)
                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias
                        If sHandwritingContent = String.Empty Then
                            DrawFieldBorder(oXGraphics, XSingleBlockDimension.Inch * fResolution, XSingleBlockDimension.Inch * fResolution, New XPoint(fMargin, fMargin), fBorderWidth)
                        Else
                            ' draw text
                            DrawText(oXGraphics, XSingleBlockDimension.Inch * fResolution, XUnit.Zero, XSingleBlockDimension.Inch * fResolution, New XPoint(fMargin, fMargin), sHandwritingContent, False, False, False, MigraDoc.DocumentObjectModel.ParagraphAlignment.Center, MigraDoc.DocumentObjectModel.Colors.Black)
                        End If
                    End Using
                End Using

                oReturnBitmap = Converter.BitmapConvertGrayscale(oBitmap)
            End Using

            Dim oReturnMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(oReturnBitmap)
            oReturnBitmap.Dispose()

            Dim oInnerMatrix As Emgu.CV.Matrix(Of Byte) = Nothing
            If iInnerBitmapWidth > 0 Then
                oInnerMatrix = New Emgu.CV.Matrix(Of Byte)(iInnerBitmapWidth, iInnerBitmapWidth)
                oInnerMatrix.SetValue(Byte.MaxValue)
            End If

            Return New Tuple(Of Emgu.CV.Matrix(Of Byte), Emgu.CV.Matrix(Of Byte))(oReturnMatrix, oInnerMatrix)
        End Function
        Public Shared Function GetFreeMatrix(ByVal oSize As Size, ByVal fResolution As Single, ByVal fMargin As Single) As Emgu.CV.Matrix(Of Byte)
            ' gets the bitmap for the border box of the free field
            Dim XWidth As New XUnit(oSize.Width)
            Dim XHeight As New XUnit(oSize.Height)
            Dim XSize As New XSize(XWidth.Point, XHeight.Point)
            Dim iBitmapWidth As Integer = Math.Ceiling((XWidth.Inch * fResolution) + (fMargin * 2))
            Dim iBitmapHeight As Integer = Math.Ceiling((XHeight.Inch * fResolution) + (fMargin * 2))

            Dim oReturnBitmap As System.Drawing.Bitmap = Nothing
            Using oBitmap As New System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                oBitmap.SetResolution(fResolution, fResolution)
                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oBitmap)
                    oGraphics.PageUnit = System.Drawing.GraphicsUnit.Pixel
                    oGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    oGraphics.FillRectangle(System.Drawing.Brushes.White, 0, 0, oBitmap.Width, oBitmap.Height)
                    Using oXGraphics As XGraphics = XGraphics.FromGraphics(oGraphics, New XSize(oBitmap.Width * 72 / fResolution, oBitmap.Height * 72 / fResolution), XGraphicsUnit.Point)
                        oXGraphics.SmoothingMode = XSmoothingMode.AntiAlias

                        DrawFieldBorder(oXGraphics, XWidth.Inch * fResolution, XHeight.Inch * fResolution, New XPoint(fMargin, fMargin), 3)
                    End Using
                End Using

                oReturnBitmap = Converter.BitmapConvertGrayscale(oBitmap)
            End Using

            Dim oReturnMatrix As Emgu.CV.Matrix(Of Byte) = Converter.BitmapToMatrix(oReturnBitmap)
            oReturnBitmap.Dispose()

            Return oReturnMatrix
        End Function
    End Class
    Public Class BarcodeGenerator
        Public Enum CodeSet
            CodeA
            CodeB
        End Enum

        ' Represent the set of code values to be output into barcode form
        Public Class Code128Content
            Private mCodeList As Integer()

            Public Sub New(AsciiData As String)
                ' Create content based on a string of ASCII data
                mCodeList = StringToCode128(AsciiData)
            End Sub
            Public ReadOnly Property Codes() As Integer()
                ' Provides the Code128 code values representing the object's string
                Get
                    Return mCodeList
                End Get
            End Property
            Private Shared Function StringToCode128(AsciiData As String) As Integer()
                ' Transform the string into integers representing the Code128 codes necessary to represent it
                ' turn the string into ascii byte data
                Dim asciiBytes As Byte() = Text.Encoding.ASCII.GetBytes(AsciiData)

                ' decide which codeset to start with
                Dim csa1 As Code128Code.CodeSetAllowed = If(asciiBytes.Length > 0, Code128Code.CodesetAllowedForChar(asciiBytes(0)), Code128Code.CodeSetAllowed.CodeAorB)
                Dim csa2 As Code128Code.CodeSetAllowed = If(asciiBytes.Length > 0, Code128Code.CodesetAllowedForChar(asciiBytes(1)), Code128Code.CodeSetAllowed.CodeAorB)
                Dim currcs As CodeSet = GetBestStartSet(csa1, csa2)

                ' set up the beginning of the barcode
                Dim codes As New ArrayList(asciiBytes.Length + 3)

                ' assume no codeset changes, account for start, checksum, and stop
                codes.Add(Code128Code.StartCodeForCodeSet(currcs))

                ' add the codes for each character in the string
                For i As Integer = 0 To asciiBytes.Length - 1
                    Dim thischar As Integer = asciiBytes(i)
                    Dim nextchar As Integer = If(asciiBytes.Length > (i + 1), asciiBytes(i + 1), -1)
                    codes.AddRange(Code128Code.CodesForChar(thischar, nextchar, currcs))
                Next

                ' calculate the check digit
                Dim checksum As Integer = CInt(codes(0))
                For i As Integer = 1 To codes.Count - 1
                    checksum += i * CInt(codes(i))
                Next
                codes.Add(checksum Mod 103)

                codes.Add(Code128Code.StopCode())

                Dim result As Integer() = TryCast(codes.ToArray(GetType(Integer)), Integer())
                Return result
            End Function
            Private Shared Function GetBestStartSet(csa1 As Code128Code.CodeSetAllowed, csa2 As Code128Code.CodeSetAllowed) As CodeSet
                ' Determines the best starting code set based on the the first two characters of the string to be encoded
                Dim vote As Integer = 0
                vote += If((csa1 = Code128Code.CodeSetAllowed.CodeA), 1, 0)
                vote += If((csa1 = Code128Code.CodeSetAllowed.CodeB), -1, 0)
                vote += If((csa2 = Code128Code.CodeSetAllowed.CodeA), 1, 0)
                vote += If((csa2 = Code128Code.CodeSetAllowed.CodeB), -1, 0)

                Return If((vote > 0), CodeSet.CodeA, CodeSet.CodeB)
            End Function
        End Class
        Public Class Code128Code
            ' Static tools for determining codes for individual characters in the content
            Private Sub New()
            End Sub
#Region "Constants"
            Private Const cSHIFT As Integer = 98
            Private Const cCODEA As Integer = 101
            Private Const cCODEB As Integer = 100
            Private Const cSTARTA As Integer = 103
            Private Const cSTARTB As Integer = 104
            Private Const cSTOP As Integer = 106
#End Region
            Public Shared Function CodesForChar(CharAscii As Integer, LookAheadAscii As Integer, ByRef CurrCodeSet As CodeSet) As Integer()
                ' Get the Code128 code value(s) to represent an ASCII character, with optional look-ahead for length optimization
                ' returns an array of integers representing the codes that need to be output to produce the given character
                Dim result As Integer()
                Dim shifter As Integer = -1

                If Not CharCompatibleWithCodeset(CharAscii, CurrCodeSet) Then
                    ' if we have a lookahead character AND if the next character is ALSO not compatible
                    If (LookAheadAscii <> -1) AndAlso Not CharCompatibleWithCodeset(LookAheadAscii, CurrCodeSet) Then
                        ' we need to switch code sets
                        Select Case CurrCodeSet
                            Case CodeSet.CodeA
                                shifter = cCODEB
                                CurrCodeSet = CodeSet.CodeB
                                Exit Select
                            Case CodeSet.CodeB
                                shifter = cCODEA
                                CurrCodeSet = CodeSet.CodeA
                                Exit Select
                        End Select
                    Else
                        ' no need to switch code sets, a temporary SHIFT will suffice
                        shifter = cSHIFT
                    End If
                End If

                If shifter <> -1 Then
                    result = New Integer(1) {}
                    result(0) = shifter
                    result(1) = CodeValueForChar(CharAscii)
                Else
                    result = New Integer(0) {}
                    result(0) = CodeValueForChar(CharAscii)
                End If

                Return result
            End Function
            Public Shared Function CodesetAllowedForChar(CharAscii As Integer) As CodeSetAllowed
                ' Tells us which codesets a given character value is allowed in
                If CharAscii >= 32 AndAlso CharAscii <= 95 Then
                    Return CodeSetAllowed.CodeAorB
                Else
                    Return If((CharAscii < 32), CodeSetAllowed.CodeA, CodeSetAllowed.CodeB)
                End If
            End Function
            Public Shared Function CharCompatibleWithCodeset(CharAscii As Integer, currcs As CodeSet) As Boolean
                ' Determine if a character can be represented in a given codeset
                Dim csa As CodeSetAllowed = CodesetAllowedForChar(CharAscii)
                Return csa = CodeSetAllowed.CodeAorB OrElse (csa = CodeSetAllowed.CodeA AndAlso currcs = CodeSet.CodeA) OrElse (csa = CodeSetAllowed.CodeB AndAlso currcs = CodeSet.CodeB)
            End Function
            Public Shared Function CodeValueForChar(CharAscii As Integer) As Integer
                ' Gets the integer code128 code value for a character (assuming the appropriate code set)
                Return If((CharAscii >= 32), CharAscii - 32, CharAscii + 64)
            End Function
            Public Shared Function StartCodeForCodeSet(cs As CodeSet) As Integer
                ' Return the appropriate START code depending on the codeset we want to be in
                Return If(cs = CodeSet.CodeA, cSTARTA, cSTARTB)
            End Function
            Public Shared Function StopCode() As Integer
                ' Return the Code128 stop code
                Return cSTOP
            End Function
            Public Enum CodeSetAllowed
                ' Indicates which code sets can represent a character -- CodeA, CodeB, or either
                CodeA
                CodeB
                CodeAorB
            End Enum
        End Class
        Public Class Code128Rendering
            Private Sub New()
            End Sub
#Region "Code patterns"
            Private Shared ReadOnly cPatterns As Integer(,) = {{2, 1, 2, 2, 2, 2,
                0, 0}, {2, 2, 2, 1, 2, 2,
                0, 0}, {2, 2, 2, 2, 2, 1,
                0, 0}, {1, 2, 1, 2, 2, 3,
                0, 0}, {1, 2, 1, 3, 2, 2,
                0, 0}, {1, 3, 1, 2, 2, 2,
                0, 0},
                {1, 2, 2, 2, 1, 3,
                0, 0}, {1, 2, 2, 3, 1, 2,
                0, 0}, {1, 3, 2, 2, 1, 2,
                0, 0}, {2, 2, 1, 2, 1, 3,
                0, 0}, {2, 2, 1, 3, 1, 2,
                0, 0}, {2, 3, 1, 2, 1, 2,
                0, 0},
                {1, 1, 2, 2, 3, 2,
                0, 0}, {1, 2, 2, 1, 3, 2,
                0, 0}, {1, 2, 2, 2, 3, 1,
                0, 0}, {1, 1, 3, 2, 2, 2,
                0, 0}, {1, 2, 3, 1, 2, 2,
                0, 0}, {1, 2, 3, 2, 2, 1,
                0, 0},
                {2, 2, 3, 2, 1, 1,
                0, 0}, {2, 2, 1, 1, 3, 2,
                0, 0}, {2, 2, 1, 2, 3, 1,
                0, 0}, {2, 1, 3, 2, 1, 2,
                0, 0}, {2, 2, 3, 1, 1, 2,
                0, 0}, {3, 1, 2, 1, 3, 1,
                0, 0},
                {3, 1, 1, 2, 2, 2,
                0, 0}, {3, 2, 1, 1, 2, 2,
                0, 0}, {3, 2, 1, 2, 2, 1,
                0, 0}, {3, 1, 2, 2, 1, 2,
                0, 0}, {3, 2, 2, 1, 1, 2,
                0, 0}, {3, 2, 2, 2, 1, 1,
                0, 0},
                {2, 1, 2, 1, 2, 3,
                0, 0}, {2, 1, 2, 3, 2, 1,
                0, 0}, {2, 3, 2, 1, 2, 1,
                0, 0}, {1, 1, 1, 3, 2, 3,
                0, 0}, {1, 3, 1, 1, 2, 3,
                0, 0}, {1, 3, 1, 3, 2, 1,
                0, 0},
                {1, 1, 2, 3, 1, 3,
                0, 0}, {1, 3, 2, 1, 1, 3,
                0, 0}, {1, 3, 2, 3, 1, 1,
                0, 0}, {2, 1, 1, 3, 1, 3,
                0, 0}, {2, 3, 1, 1, 1, 3,
                0, 0}, {2, 3, 1, 3, 1, 1,
                0, 0},
                {1, 1, 2, 1, 3, 3,
                0, 0}, {1, 1, 2, 3, 3, 1,
                0, 0}, {1, 3, 2, 1, 3, 1,
                0, 0}, {1, 1, 3, 1, 2, 3,
                0, 0}, {1, 1, 3, 3, 2, 1,
                0, 0}, {1, 3, 3, 1, 2, 1,
                0, 0},
                {3, 1, 3, 1, 2, 1,
                0, 0}, {2, 1, 1, 3, 3, 1,
                0, 0}, {2, 3, 1, 1, 3, 1,
                0, 0}, {2, 1, 3, 1, 1, 3,
                0, 0}, {2, 1, 3, 3, 1, 1,
                0, 0}, {2, 1, 3, 1, 3, 1,
                0, 0},
                {3, 1, 1, 1, 2, 3,
                0, 0}, {3, 1, 1, 3, 2, 1,
                0, 0}, {3, 3, 1, 1, 2, 1,
                0, 0}, {3, 1, 2, 1, 1, 3,
                0, 0}, {3, 1, 2, 3, 1, 1,
                0, 0}, {3, 3, 2, 1, 1, 1,
                0, 0},
                {3, 1, 4, 1, 1, 1,
                0, 0}, {2, 2, 1, 4, 1, 1,
                0, 0}, {4, 3, 1, 1, 1, 1,
                0, 0}, {1, 1, 1, 2, 2, 4,
                0, 0}, {1, 1, 1, 4, 2, 2,
                0, 0}, {1, 2, 1, 1, 2, 4,
                0, 0},
                {1, 2, 1, 4, 2, 1,
                0, 0}, {1, 4, 1, 1, 2, 2,
                0, 0}, {1, 4, 1, 2, 2, 1,
                0, 0}, {1, 1, 2, 2, 1, 4,
                0, 0}, {1, 1, 2, 4, 1, 2,
                0, 0}, {1, 2, 2, 1, 1, 4,
                0, 0},
                {1, 2, 2, 4, 1, 1,
                0, 0}, {1, 4, 2, 1, 1, 2,
                0, 0}, {1, 4, 2, 2, 1, 1,
                0, 0}, {2, 4, 1, 2, 1, 1,
                0, 0}, {2, 2, 1, 1, 1, 4,
                0, 0}, {4, 1, 3, 1, 1, 1,
                0, 0},
                {2, 4, 1, 1, 1, 2,
                0, 0}, {1, 3, 4, 1, 1, 1,
                0, 0}, {1, 1, 1, 2, 4, 2,
                0, 0}, {1, 2, 1, 1, 4, 2,
                0, 0}, {1, 2, 1, 2, 4, 1,
                0, 0}, {1, 1, 4, 2, 1, 2,
                0, 0},
                {1, 2, 4, 1, 1, 2,
                0, 0}, {1, 2, 4, 2, 1, 1,
                0, 0}, {4, 1, 1, 2, 1, 2,
                0, 0}, {4, 2, 1, 1, 1, 2,
                0, 0}, {4, 2, 1, 2, 1, 1,
                0, 0}, {2, 1, 2, 1, 4, 1,
                0, 0},
                {2, 1, 4, 1, 2, 1,
                0, 0}, {4, 1, 2, 1, 2, 1,
                0, 0}, {1, 1, 1, 1, 4, 3,
                0, 0}, {1, 1, 1, 3, 4, 1,
                0, 0}, {1, 3, 1, 1, 4, 1,
                0, 0}, {1, 1, 4, 1, 1, 3,
                0, 0},
                {1, 1, 4, 3, 1, 1,
                0, 0}, {4, 1, 1, 1, 1, 3,
                0, 0}, {4, 1, 1, 3, 1, 1,
                0, 0}, {1, 1, 3, 1, 4, 1,
                0, 0}, {1, 1, 4, 1, 3, 1,
                0, 0}, {3, 1, 1, 1, 4, 1,
                0, 0},
                {4, 1, 1, 1, 3, 1,
                0, 0}, {2, 1, 1, 4, 1, 2,
                0, 0}, {2, 1, 1, 2, 1, 4,
                0, 0}, {2, 1, 1, 2, 3, 2,
                0, 0}, {2, 3, 3, 1, 1, 1,
                2, 0}}
#End Region
            Private Const TopMargin As Integer = 5
            Private Const BottomMargin As Integer = 4
            Private Const BarcodeHeight As Integer = 62
            Private Const SideMargin As Integer = 20
            Private Const BarWeight As Integer = 2
            Public Shared Function MakeBarcodeImage(InputData As String) As System.Drawing.Bitmap
                ' get the Code128 codes to represent the message
                Dim content As New Code128Content(InputData)
                Dim codes As Integer() = content.Codes

                Dim width As Integer, height As Integer
                width = (((codes.Length - 3) * 11 + 35) * BarWeight) + (2 * SideMargin)
                height = BarcodeHeight + TopMargin + BottomMargin

                ' get surface to draw on
                Dim myimg As System.Drawing.Bitmap = New System.Drawing.Bitmap(width, height)
                Using gr As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(myimg)

                    ' set to white so we don't have to fill the spaces with white
                    gr.FillRectangle(System.Drawing.Brushes.White, 0, 0, width, height)

                    ' skip quiet zone
                    Dim cursor As Integer = SideMargin

                    For codeidx As Integer = 0 To codes.Length - 1
                        Dim code As Integer = codes(codeidx)

                        ' take the bars two at a time: a black and a white
                        For bar As Integer = 0 To 7 Step 2
                            Dim barwidth As Integer = cPatterns(code, bar) * BarWeight
                            Dim spcwidth As Integer = cPatterns(code, bar + 1) * BarWeight

                            ' if width is zero, don't try to draw it
                            If barwidth > 0 Then
                                gr.FillRectangle(System.Drawing.Brushes.Black, cursor, TopMargin, barwidth, BarcodeHeight)
                            End If

                            ' advance cursor beyond this pair
                            cursor += (barwidth + spcwidth)
                        Next
                    Next
                End Using

                Return myimg
            End Function
        End Class
    End Class
    Public Class CommonFunctions
#Region "UI"
        Public Shared Function GetChildObjects(Of T As FrameworkElement)(ByRef oObject As FrameworkElement, Optional ByVal sName As String = "") As List(Of T)
            ' recurses through each level of a framework element to get all of its children, optionally filtered by type
            Dim oObjectList As New List(Of T)
            Dim iObjectCount As Integer = Media.VisualTreeHelper.GetChildrenCount(oObject)
            Dim oDependencyObjectList As List(Of DependencyObject) = GetChildDependencyObjects(oObject, GetType(T))
            For Each oChildObject In oDependencyObjectList
                Dim oConvertedObject As T = Convert.ChangeType(oChildObject, GetType(T))
                If sName = "" OrElse oConvertedObject.Name = sName Then
                    oObjectList.Add(oConvertedObject)
                End If
            Next
            Return oObjectList
        End Function
        Private Shared Function GetChildDependencyObjects(ByRef oObject As DependencyObject, Optional ByVal oType As Type = Nothing) As List(Of DependencyObject)
            ' recurses through each level of a dependency object to get all of its children, optionally filtered by type
            Dim oObjectList As New List(Of DependencyObject)
            Dim iObjectCount As Integer = Media.VisualTreeHelper.GetChildrenCount(oObject)
            If iObjectCount > 0 Then
                For i = 0 To iObjectCount - 1
                    Dim oChildObject As DependencyObject = Media.VisualTreeHelper.GetChild(oObject, i)
                    If IsNothing(oType) Then
                        oObjectList.Add(oChildObject)
                        oObjectList.AddRange(GetChildDependencyObjects(oChildObject))
                    Else
                        If oChildObject.GetType.Equals(oType) Then
                            oObjectList.Add(oChildObject)
                        End If
                        oObjectList.AddRange(GetChildDependencyObjects(oChildObject, oType))
                    End If
                Next
            End If
            Return oObjectList
        End Function
        Public Shared Function GetParentObject(Of T As FrameworkElement)(ByRef oObject As FrameworkElement, Optional ByVal sName As String = "") As T
            ' recurses through each parent level of a framework element to get a parent of a particular type, optionally filtered by name
            Dim oParent As DependencyObject = oObject
            Do
                oParent = Media.VisualTreeHelper.GetParent(oParent)
                If IsNothing(oParent) Then
                    Exit Do
                Else
                    If oParent.GetType.Equals(GetType(T)) AndAlso (sName = "" OrElse CType(oParent, FrameworkElement).Name = sName) Then
                        Exit Do
                    End If
                End If
            Loop

            Return CType(oParent, T)
        End Function
        Public Shared Sub RefreshControl(ByRef oFrameworkElement As FrameworkElement)
            ' refresh element
            Dim oFrame As New Threading.DispatcherFrame()
            oFrameworkElement.Dispatcher.BeginInvoke(Threading.DispatcherPriority.Background, New Threading.DispatcherOperationCallback(AddressOf ExitFrame), oFrame)
            Threading.Dispatcher.PushFrame(oFrame)
        End Sub
        Private Shared Function ExitFrame(ByVal f As Object) As Object
            CType(f, Threading.DispatcherFrame).Continue = False
            Return Nothing
        End Function
        Public Shared Sub IterateElementDictionary(ByRef oElementDictionary As Dictionary(Of String, FrameworkElement), ByRef oFrameworkElement As FrameworkElement, Optional ByVal sPrefix As String = "", Optional ByVal bFirstPass As Boolean = True)
            ' iterates through the logical tree and populates the framework element dictionary with named elements only
            If bFirstPass Then
                oElementDictionary.Clear()
            End If

            If oFrameworkElement.Name <> String.Empty AndAlso Left(oFrameworkElement.Name, sPrefix.Length) = sPrefix AndAlso Not oElementDictionary.ContainsKey(oFrameworkElement.Name) Then
                oElementDictionary.Add(oFrameworkElement.Name, oFrameworkElement)
            End If
            Dim oFrameworkElementCollection = LogicalTreeHelper.GetChildren(oFrameworkElement).OfType(Of FrameworkElement)()
            For Each oFrameworkElementChild In oFrameworkElementCollection
                IterateElementDictionary(oElementDictionary, oFrameworkElementChild, sPrefix, False)
            Next
        End Sub
        Public Shared Sub RepeatCheck(sender As Object, e As RoutedEventArgs, ByVal oAction As Action)
            ' ignore stylus and touch events
            If (Not e.RoutedEvent.Equals(Input.Stylus.StylusDownEvent)) AndAlso (Not e.RoutedEvent.Equals(UIElement.TouchDownEvent)) Then
                oAction.Invoke()
            End If
        End Sub
#End Region
#Region "Serialisation"
        Public Shared Function GetKnownTypes(Optional ByVal oTypes As List(Of Type) = Nothing) As List(Of Type)
            ' gets list of known types
            Dim oKnownTypes As New List(Of Type)

            If Not IsNothing(oTypes) Then
                oKnownTypes.AddRange(oTypes)
            End If
            Return oKnownTypes
        End Function
        Public Shared Sub SerializeDataContractStream(Of T)(ByRef oStream As IO.Stream, ByVal data As T, Optional ByVal oAdditionalTypes As List(Of Type) = Nothing, Optional ByVal bUseKnownTypes As Boolean = True)
            ' serialise to stream
            Dim oKnownTypes As New List(Of Type)
            If bUseKnownTypes Then
                oKnownTypes.AddRange(GetKnownTypes)
            End If
            If Not IsNothing(oAdditionalTypes) Then
                oKnownTypes.AddRange(oAdditionalTypes)
            End If

            Dim oDataContractSerializer As New DataContractSerializer(GetType(T), oKnownTypes)
            oDataContractSerializer.WriteObject(oStream, data)
        End Sub
        Public Shared Function DeserializeDataContractStream(Of T)(ByRef oStream As IO.Stream, Optional ByVal oAdditionalTypes As List(Of Type) = Nothing, Optional ByVal bUseKnownTypes As Boolean = True, Optional ByVal oDataContractSurrogate As IDataContractSurrogate = Nothing) As T
            ' deserialise from stream
            Dim oXmlDictionaryReaderQuotas As New Xml.XmlDictionaryReaderQuotas()
            oXmlDictionaryReaderQuotas.MaxArrayLength = 100000000
            Dim oXmlDictionaryReader As Xml.XmlDictionaryReader = Xml.XmlDictionaryReader.CreateTextReader(oStream, oXmlDictionaryReaderQuotas)

            Dim theObject As T = Nothing
            Try
                Dim oKnownTypes As New List(Of Type)
                If bUseKnownTypes Then
                    oKnownTypes.AddRange(GetKnownTypes)
                End If
                If Not IsNothing(oAdditionalTypes) Then
                    oKnownTypes.AddRange(oAdditionalTypes)
                End If

                Dim oDataContractSerializer As New DataContractSerializer(GetType(T), oKnownTypes, Integer.MaxValue, False, True, oDataContractSurrogate)
                theObject = oDataContractSerializer.ReadObject(oXmlDictionaryReader, True)
            Catch ex As SerializationException
            End Try

            oXmlDictionaryReader.Close()
            Return theObject
        End Function
        Public Shared Sub SerializeDataContractFile(Of T)(ByVal sFilePath As String, ByVal data As T, Optional ByVal oAdditionalTypes As List(Of Type) = Nothing, Optional ByVal bUseKnownTypes As Boolean = True, Optional ByVal sExtension As String = "gz")
            ' serialise using data contract serialiser
            ' compress using gzip
            Using oFileStream As IO.FileStream = IO.File.Create(If(sExtension = String.Empty, sFilePath, ReplaceExtension(sFilePath, sExtension)))
                Using oGZipStream As New IO.Compression.GZipStream(oFileStream, IO.Compression.CompressionMode.Compress)
                    SerializeDataContractStream(oGZipStream, data, oAdditionalTypes, bUseKnownTypes)
                End Using
            End Using
        End Sub
        Public Shared Function DeserializeDataContractFile(Of T)(ByVal sFilePath As String, Optional ByVal oAdditionalTypes As List(Of Type) = Nothing, Optional ByVal bUseKnownTypes As Boolean = True, Optional ByVal oDataContractSurrogate As IDataContractSurrogate = Nothing, Optional ByVal sExtension As String = "gz") As T
            ' deserialise using data contract serialiser
            Using oFileStream As IO.FileStream = IO.File.OpenRead(If(sExtension = String.Empty, sFilePath, ReplaceExtension(sFilePath, sExtension)))
                Using oGZipStream As New IO.Compression.GZipStream(oFileStream, IO.Compression.CompressionMode.Decompress)
                    Return DeserializeDataContractStream(Of T)(oGZipStream, oAdditionalTypes, bUseKnownTypes, oDataContractSurrogate)
                End Using
            End Using
        End Function
        Public Shared Sub SerializeBinaryFile(Of T)(ByVal sFilePath As String, ByVal data As T)
            ' binary serialisation
            Using oFileStream As IO.FileStream = IO.File.Create(ReplaceExtension(sFilePath, "bin"))
                Using oGZipStream As New IO.Compression.GZipStream(oFileStream, IO.Compression.CompressionMode.Compress)
                    Dim oBinaryFormatter As New Formatters.Binary.BinaryFormatter
                    oBinaryFormatter.Serialize(oGZipStream, data)
                    oGZipStream.Close()
                End Using
            End Using
        End Sub
        Public Shared Function DeserializeBinaryFile(Of T)(ByVal sFilePath As String) As T
            ' binary deserialisation
            Using oFileStream As IO.FileStream = IO.File.OpenRead(ReplaceExtension(sFilePath, "bin"))
                Using oGZipStream As New IO.Compression.GZipStream(oFileStream, IO.Compression.CompressionMode.Decompress)
                    Dim oBinaryFormatter As New Formatters.Binary.BinaryFormatter
                    Dim theObject As T = oBinaryFormatter.Deserialize(oGZipStream)
                    oGZipStream.Close()
                    Return theObject
                End Using
            End Using
        End Function
        Public Shared Function SerializeDataContractText(Of T)(ByVal data As T, Optional ByVal oAdditionalTypes As List(Of Type) = Nothing, Optional ByVal bUseKnownTypes As Boolean = True) As String
            ' serialise using data contract serialiser
            ' returns base64 text
            Using oMemoryStream As New IO.MemoryStream
                SerializeDataContractStream(oMemoryStream, data, oAdditionalTypes, bUseKnownTypes)
                Dim bByteArray As Byte() = oMemoryStream.ToArray
                Return Convert.ToBase64String(bByteArray)
            End Using
        End Function
        Public Shared Function DeserializeDataContractText(Of T)(ByVal sBase64String As String, Optional ByVal oAdditionalTypes As List(Of Type) = Nothing, Optional ByVal bUseKnownTypes As Boolean = True, Optional ByVal oDataContractSurrogate As IDataContractSurrogate = Nothing) As T
            ' deserialise from base64 text
            Using oMemoryStream As New IO.MemoryStream(Convert.FromBase64String(sBase64String))
                Return DeserializeDataContractStream(Of T)(oMemoryStream, oAdditionalTypes, bUseKnownTypes, oDataContractSurrogate)
            End Using
        End Function
#End Region
#Region "Utility"
        Public Shared Sub ParallelRun(ByVal height As Integer, ByRef oObject1 As Object, ByRef oObject2 As Object, ByRef oObject3 As Object, ByRef oObject4 As Object, ByVal TaskDelegate As Action(Of Object), Optional ByVal iTaskCount As Integer = -1)
            ' runs tasks in parallel
            If iTaskCount = -1 Then
                iTaskCount = TaskCount
            End If
            Dim oTaskList As New List(Of Task)
            Dim iTopList As List(Of Integer)
            Dim iHeightList As List(Of Integer)

            If ParallelDictionary.ContainsKey(height) Then
                iTopList = ParallelDictionary(height).Item1
                iHeightList = ParallelDictionary(height).Item2
            Else
                iTopList = New List(Of Integer)
                iHeightList = New List(Of Integer)
                For i = 0 To iTaskCount - 2
                    iTopList.Add(Math.Truncate(height / iTaskCount) * i)
                    iHeightList.Add(Math.Truncate(height / iTaskCount))
                Next
                iTopList.Add(Math.Truncate(height / iTaskCount) * (iTaskCount - 1))
                iHeightList.Add(height - (Math.Truncate(height / iTaskCount) * (iTaskCount - 1)))

                ParallelDictionary.Add(height, New Tuple(Of List(Of Integer), List(Of Integer))(iTopList, iHeightList))
            End If

            For i = 0 To iTaskCount - 1
                If iHeightList(i) > 0 Then
                    oTaskList.Add(New Task(TaskDelegate, New Tuple(Of Integer, Integer, Integer, Object, Object, Object, Object)(iTopList(i), iHeightList(i), i, oObject1, oObject2, oObject3, oObject4)))
                End If
            Next
            For Each oTask In oTaskList
                oTask.Start()
            Next
            Task.WaitAll(oTaskList.ToArray)
            oTaskList.Clear()
        End Sub
        Public Shared Sub ClearMemory()
            ' clear up memory
            Runtime.GCSettings.LargeObjectHeapCompactionMode = Runtime.GCLargeObjectHeapCompactionMode.CompactOnce
            GC.Collect(2, GCCollectionMode.Forced, True, True)
            GC.WaitForPendingFinalizers()
        End Sub
        Public Shared Sub ProtectedRunTasks(ByVal oActions As List(Of Action))
            ' protects tasks from memory exceptions by clearing memory if an error occurs and rerunning the task
            Dim oActionDictionary As New Dictionary(Of Guid, Action)
            For Each oAction As Action In oActions
                oActionDictionary.Add(Guid.NewGuid, oAction)
            Next

            Do Until oActionDictionary.Count = 0
                ' create task dictionary from actions
                Dim oTaskDictionary As New ConcurrentDictionary(Of Guid, Task)
                For Each oGUID As Guid In oActionDictionary.Keys
                    oTaskDictionary.TryAdd(oGUID, Task.Run(oActionDictionary(oGUID)))
                Next

                Try
                    Task.WaitAll(oTaskDictionary.Values.ToArray)
                Catch ae As AggregateException
                End Try

                For Each oGUID In oTaskDictionary.Keys
                    If oTaskDictionary(oGUID).IsCompleted Then
                        oActionDictionary.Remove(oGUID)
                    ElseIf oTaskDictionary(oGUID).IsFaulted Then
                        For Each ex In oTaskDictionary(oGUID).Exception.Flatten.InnerExceptions
                            If TypeOf ex Is OutOfMemoryException Then
                                ClearMemory()
                            Else
                                Throw ex
                            End If
                        Next
                    End If
                Next

                oTaskDictionary.Clear()
            Loop
        End Sub
        Public Shared Function Compress(ByVal oByteArray As Byte()) As Byte()
            ' compress byte array
            If IsNothing(oByteArray) Then
                Return Nothing
            Else
                Using oOutputMemoryStream As New IO.MemoryStream()
                    Using oGZipStream As New IO.Compression.GZipStream(oOutputMemoryStream, IO.Compression.CompressionMode.Compress)
                        oGZipStream.Write(oByteArray, 0, oByteArray.Length)
                    End Using
                    Return oOutputMemoryStream.ToArray
                End Using
            End If
        End Function
        Public Shared Function Decompress(ByVal oByteArray As Byte()) As Byte()
            ' decompress byte array
            If IsNothing(oByteArray) Then
                Return Nothing
            Else
                Const iBufferSize As Integer = 4096
                Using oInputMemoryStream = New IO.MemoryStream(oByteArray)
                    Using oGZipStream As New IO.Compression.GZipStream(oInputMemoryStream, IO.Compression.CompressionMode.Decompress)
                        Using oOutputMemoryStream As New IO.MemoryStream()
                            Dim oBuffer As Byte() = New Byte(iBufferSize - 1) {}
                            While True
                                Dim iReadBytes As Integer = oGZipStream.Read(oBuffer, 0, oBuffer.Length)
                                If iReadBytes > 0 Then
                                    oOutputMemoryStream.Write(oBuffer, 0, iReadBytes)
                                End If
                                If iReadBytes < iBufferSize Then
                                    Exit While
                                End If
                            End While
                            Return oOutputMemoryStream.ToArray()
                        End Using
                    End Using
                End Using
            End If
        End Function
        Public Shared Function CheckCharacterASCII(ByVal sChar As String, oCharacterASCII As Enumerations.CharacterASCII) As Boolean
            ' checks if the supplied character is a type of the specified enumeration
            Return CheckCharacterASCII(Asc(sChar), oCharacterASCII)
        End Function
        Public Shared Function CheckCharacterASCII(ByVal iChar As Integer, oCharacterASCII As Enumerations.CharacterASCII) As Boolean
            ' checks if the supplied character is a type of the specified enumeration
            If oCharacterASCII = Enumerations.CharacterASCII.None Then
                Return False
            Else
                Select Case iChar
                    Case 48 To 57
                        If (oCharacterASCII And Enumerations.CharacterASCII.Numbers) = Enumerations.CharacterASCII.Numbers Then
                            Return True
                        Else
                            Return False
                        End If
                    Case 65 To 90
                        If (oCharacterASCII And Enumerations.CharacterASCII.Uppercase) = Enumerations.CharacterASCII.Uppercase Then
                            Return True
                        Else
                            Return False
                        End If
                    Case 97 To 122
                        If (oCharacterASCII And Enumerations.CharacterASCII.Lowercase) = Enumerations.CharacterASCII.Lowercase Then
                            Return True
                        Else
                            Return False
                        End If
                    Case Else
                        If (oCharacterASCII And Enumerations.CharacterASCII.NonAlphaNumeric) = Enumerations.CharacterASCII.NonAlphaNumeric Then
                            Return True
                        Else
                            Return False
                        End If
                End Select
            End If
        End Function
        Public Shared Sub DispatcherInvoke(ByVal oUIDispatcher As Threading.Dispatcher, ByVal oAction As Action)
            ' invokes the action via the dispatcher if present
            If IsNothing(oUIDispatcher) Then
                oAction.Invoke
            Else
                oUIDispatcher.Invoke(oAction)
            End If
        End Sub
#End Region
#Region "I/O"
        Public Shared Function ReplaceExtension(ByVal sFilePath As String, ByVal sExtension As String) As String
            Try
                Return IO.Path.ChangeExtension(sFilePath, sExtension)
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function
        Public Shared Function LoadBitmap(ByVal sFilename As String, Optional ByVal oPixelFormat As System.Drawing.Imaging.PixelFormat = System.Drawing.Imaging.PixelFormat.Format8bppIndexed) As System.Drawing.Bitmap
            ' a non-locking bitmap load
            Dim oReturnBitmap As System.Drawing.Bitmap = Nothing
            Using oFileStream As New IO.FileStream(sFilename, IO.FileMode.Open, IO.FileAccess.Read)
                Using oBitmap As New System.Drawing.Bitmap(oFileStream)
                    oReturnBitmap = New System.Drawing.Bitmap(oBitmap.Width, oBitmap.Height, oPixelFormat)
                    oReturnBitmap.SetResolution(oBitmap.HorizontalResolution, oBitmap.VerticalResolution)
                    Select Case oPixelFormat
                        Case System.Drawing.Imaging.PixelFormat.Format32bppArgb
                            Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oReturnBitmap)
                                oGraphics.DrawImage(oBitmap, System.Drawing.Point.Empty)
                                oGraphics.Flush()
                            End Using
                        Case System.Drawing.Imaging.PixelFormat.Format8bppIndexed
                            Using oColourBitmap As New System.Drawing.Bitmap(oBitmap.Width, oBitmap.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                                oColourBitmap.SetResolution(oBitmap.HorizontalResolution, oBitmap.VerticalResolution)
                                Using oGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(oColourBitmap)
                                    oGraphics.DrawImage(oBitmap, System.Drawing.Point.Empty)
                                    oGraphics.Flush()
                                    oReturnBitmap = Converter.BitmapConvertGrayscale(oColourBitmap)
                                End Using
                            End Using
                    End Select
                End Using
            End Using
            Return oReturnBitmap
        End Function
        Public Shared Sub SaveBitmap(ByVal sFilename As String, ByVal oBitmap As System.Drawing.Bitmap, ByVal bDelete As Boolean)
            ' saves bitmap as tiff file
            If IO.File.Exists(sFilename) Then
                If bDelete Then
                    IO.File.Delete(sFilename)
                Else
                    Exit Sub
                End If
            End If
            oBitmap.Save(sFilename, System.Drawing.Imaging.ImageFormat.Tiff)
        End Sub
        Public Shared Function SafeFileName(ByVal sFileName As String, Optional ByVal sReplaceChar As String = "", Optional ByVal sRemoveChar As String = "") As String
            ' removes unsafe characters from a file name
            Dim sReturnFileName As String = sFileName
            sReturnFileName = Text.RegularExpressions.Regex.Replace(sReturnFileName, "[^\w\ !@#$%^&*()_+-={}[]|\:;'<>,.?/]", String.Empty)

            If sRemoveChar <> String.Empty Then
                Dim sRemoveChars As Char() = sRemoveChar.ToCharArray
                For Each sChar In sRemoveChars
                    sReturnFileName = sReturnFileName.Replace(sChar.ToString, String.Empty)
                Next
            End If

            Dim sInvalidFileNameChars As Char() = IO.Path.GetInvalidFileNameChars
            For Each sChar In sInvalidFileNameChars
                sReturnFileName = sReturnFileName.Replace(sChar.ToString, sReplaceChar)
            Next
            Return sReturnFileName
        End Function
        Public Shared Function SafeFrameworkName(ByVal sName As String) As String
            ' removes all characters except alphanumberic characters
            Return Text.RegularExpressions.Regex.Replace(sName, "[^a-zA-Z0-9]", New Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceMatch))
        End Function
        Private Shared Function ReplaceMatch(ByVal m As Text.RegularExpressions.Match) As String
            ' match delegate
            Dim bytes = Text.Encoding.UTF8.GetBytes(m.Value)
            Dim hexs = bytes.[Select](Function(b) String.Format("_{0:x2}", b))
            Return String.Concat(hexs)
        End Function
#End Region
    End Class
    Public Class NativeMethods
        <Runtime.InteropServices.DllImport("gdi32.dll")>
        Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
        End Function
    End Class
End Class
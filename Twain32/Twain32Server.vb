Imports Twain32Shared.Twain32Shared
Imports System.Collections.Concurrent
Imports System.IO.Pipes
Imports System.Runtime.Serialization
Imports System.Windows
Imports System.Windows.Media

Module Twain32Server
    Private oStreamString As StreamString = Nothing
    Private oSenderPipe As AnonymousPipeClientStream = Nothing
    Private oReceiverPipe As AnonymousPipeClientStream = Nothing
    Private Blocker As System.Threading.ManualResetEvent
    Private ErrorCount As Integer = 0
    Private DoLoop As Boolean = True

    ' Twain variables
    Private ScannerIsActive As Boolean = False
    Private ScanQueue As New ConcurrentQueue(Of Tuple(Of Twain32Enumerations.ScanProgress, Object))
    Private AppID As NTwain.Data.TWIdentity
    Private SelectedScannerSource As String = String.Empty
    Private WithEvents Session As NTwain.TwainSession

    Sub Main(ByVal args As String())
        ' attach a debugger
        'If Not Debugger.IsAttached Then
        'Debugger.Launch()
        'End If

        Dim parentSenderID As String = args(0)
        Dim parentReceiverID As String = args(1)

        oSenderPipe = New AnonymousPipeClientStream(PipeDirection.Out, parentReceiverID)
        oReceiverPipe = New AnonymousPipeClientStream(PipeDirection.In, parentSenderID)

        ' initialises streamstring if not done
        Dim bSuccess As Boolean = True
        Try
            If IsNothing(oStreamString) Then
                oStreamString = New StreamString(oSenderPipe, oReceiverPipe)
            End If
        Catch e As IO.IOException
            ' if connection interrupted, then exit
            Cleanup()
            bSuccess = False
        End Try

        If bSuccess Then
            Blocker = New System.Threading.ManualResetEvent(False)
            Task.Run(AddressOf Listener)
            Blocker.WaitOne()
            Cleanup()

            If Not IsNothing(Blocker) Then
                Blocker.Dispose()
            End If
        End If
    End Sub
    Private Sub Cleanup()
        CloseTwain()

        DoLoop = False

        oSenderPipe.Dispose()
        oReceiverPipe.Dispose()
        oSenderPipe = Nothing
        oReceiverPipe = Nothing
    End Sub
    Private Sub Listener()
        ' main messaging loop
        Do
            Try
                If IsNothing(oStreamString) Then
                    DoLoop = False
                Else
                    Dim sMessage As String = oStreamString.ReadString
                    Select Case sMessage
                        Case Twain32Constants.ValidateString
                            ErrorCount = 0
                            oStreamString.WriteString(Twain32Constants.Twain32GUID.ToString)
                        Case Twain32Constants.EmptyString
                            ' communication error
                            ErrorCount += 1
                        Case Else
                            ProcessMessage(sMessage)
                    End Select
                End If
            Catch e As IO.IOException
                ' if connection interrupted, then exit
                DoLoop = False
            End Try

            If ErrorCount > Twain32Constants.MaxErrorCount Then
                ' communication interrupted
                DoLoop = False
            End If
        Loop While DoLoop

        If Not IsNothing(Blocker) Then
            Blocker.Set()
        End If
    End Sub
    Private Sub ProcessMessage(ByVal sMessage As String)
        ' processes message loop
        Select Case sMessage
            Case Twain32Constants.InitString
                GetData(Of String)()
                Dim oScannerList As List(Of String) = InitTwain()
                SendData(oScannerList)
            Case Twain32Constants.SelectScannerString
                SelectedScannerSource = GetData(Of String)()
                SendData(Twain32Constants.XString)
            Case Twain32Constants.GetScannerSourcesString
                GetData(Of String)()
                Dim oScannerList As List(Of String) = GetTWAINScannerSources()
                SendData(oScannerList)
            Case Twain32Constants.ConfigureString
                GetData(Of String)()
                ScannerConfigure()
                SendData(Twain32Constants.XString)
            Case Twain32Constants.StartScanString
                GetData(Of String)()
                If ScannerIsActive Then
                    SendData(Twain32Enumerations.ScanProgress.ScanError)
                Else
                    ' clear queue
                    Do Until ScanQueue.Count = 0
                        ScanQueue.TryDequeue(Nothing)
                    Loop

                    ' starts scan process
                    ScannerScan()
                    If ScannerIsActive Then
                        SendData(Twain32Enumerations.ScanProgress.NoError)
                    Else
                        SendData(Twain32Enumerations.ScanProgress.ScanError)
                    End If
                End If
            Case Twain32Constants.GetScanImageString
                GetData(Of String)()
                If ScanQueue.Count = 0 Then
                    SendData(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.None, Nothing))
                Else
                    Dim oReturnMessage As Tuple(Of Twain32Enumerations.ScanProgress, Object) = Nothing
                    ScanQueue.TryDequeue(oReturnMessage)
                    If IsNothing(oReturnMessage) Then
                        SendData(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.None, Nothing))
                    Else
                        SendData(oReturnMessage)
                    End If
                End If
            Case Twain32Constants.ProcessBarcodeString
                Using oBarcodeBitmap As System.Drawing.Bitmap = GetData(Of System.Drawing.Bitmap)()
                    If IsNothing(oBarcodeBitmap) Then
                        SendData(Twain32Constants.XString)
                    Else
                        Using oImageScanner As New ZBar.ImageScanner
                            Dim oScanResult As List(Of ZBar.Symbol) = oImageScanner.Scan(oBarcodeBitmap)
                            If oScanResult.Count > 0 Then
                                SendData(String.Concat((From oSymbol As ZBar.Symbol In oScanResult Select oSymbol.Data).ToArray))
                            Else
                                SendData(Twain32Constants.XString)
                            End If
                        End Using
                    End If
                End Using
            Case Else
                ErrorCount += 1
        End Select
    End Sub
    Private Function GetData(Of T)() As T
        ' gets data
        oStreamString.WriteString(Twain32Constants.OKString)
        Dim sDataString As String = oStreamString.ReadString
        Return DeserializeDataContractText(Of T)(sDataString)
    End Function
    Private Sub SendData(Of T)(ByVal oData As T)
        ' returns data
        oStreamString.WriteString(Twain32Constants.OKString)
        Dim sMessage As String = oStreamString.ReadString
        If sMessage = Twain32Constants.OKString Then
            Dim sDataString As String = SerializeDataContractText(oData)
            oStreamString.WriteString(sDataString)
        End If
    End Sub
    Private Function InitTwain() As List(Of String)
        ' starts twain server and returns list of scanner sources
        ScannerIsActive = False

        If IsNothing(AppID) Then
            AppID = NTwain.Data.TWIdentity.CreateFromAssembly(NTwain.Data.DataGroups.Image, Reflection.Assembly.GetExecutingAssembly)
        End If

        Dim oScannerSources As List(Of String) = GetTWAINScannerSources()

        Return oScannerSources
    End Function
    Private Sub CloseTwain()
        ' closes twain server
        If (Not IsNothing(AppID)) AndAlso (Not IsNothing(Session)) AndAlso Session.IsDsmOpen Then
            Session.Close()
        End If
    End Sub
    Private Sub OpenSession()
        ' opens new TWAIN session
        Session = New NTwain.TwainSession(AppID)
        Session.SynchronizationContext = System.Threading.SynchronizationContext.Current
        Session.Open()
    End Sub
    Private Sub CloseSession()
        ScannerIsActive = False
        Session.Close()
        Session = Nothing
    End Sub
    Private Sub ScannerConfigure()
        ' configure scanner
        If SelectedScannerSource <> String.Empty Then
            OpenSession()

            Try
                Dim oSource As NTwain.DataSource = GetTWAINDataSource(SelectedScannerSource)
                If Not IsNothing(oSource) Then
                    oSource.Open()

                    ScannerIsActive = True
                    If oSource.IsOpen Then
                        Dim oHandle As IntPtr = Process.GetCurrentProcess().MainWindowHandle
                        oSource.Enable(NTwain.SourceEnableMode.ShowUIOnly, False, oHandle)
                    End If
                End If
            Catch ex As Exception
            Finally
                CloseSession()
            End Try
        End If
    End Sub
    Private Function GetTWAINScannerSources() As List(Of String)
        ' gets scanner sources
        OpenSession()

        Dim oScannerSources As New List(Of String)
        Dim oSources As List(Of NTwain.DataSource) = Session.GetSources.ToList

        For Each oSource As NTwain.DataSource In oSources
            Dim oReturnCode As NTwain.Data.ReturnCode = oSource.Open()
            If oReturnCode = NTwain.Data.ReturnCode.Success Then
                oScannerSources.Add(oSource.Name)
            End If
            If oSource.IsOpen Then
                oSource.Close()
            End If
        Next

        CloseSession()

        Return oScannerSources
    End Function
    Private Function GetTWAINDataSource(ByVal sName As String) As NTwain.DataSource
        ' gets a scanner data source based on the supplied name
        Dim oSources As List(Of NTwain.DataSource) = Session.GetSources.ToList
        For Each oSource As NTwain.DataSource In oSources
            If oSource.Name = sName Then
                Return oSource
            End If
        Next
        Return Nothing
    End Function
    Public Sub ScannerScan()
        ' starts scan process
        If SelectedScannerSource <> String.Empty Then
            OpenSession()

            Dim oSource As NTwain.DataSource = GetTWAINDataSource(SelectedScannerSource)
            If Not IsNothing(oSource) Then
                oSource.Open()

                If oSource.IsOpen Then
                    ScannerIsActive = True
                    Dim oHandle As IntPtr = Process.GetCurrentProcess().MainWindowHandle
                    oSource.Enable(NTwain.SourceEnableMode.ShowUI, False, oHandle)
                Else
                    CloseSession()
                End If
            Else
                CloseSession()
            End If
        End If
    End Sub
    Private Sub Twain_TransferReadyHandler(sender As Object, e As NTwain.TransferReadyEventArgs) Handles Session.TransferReady
        ' sets up the data transfer mechanism
        Dim oSource As NTwain.DataSource = GetTWAINDataSource(SelectedScannerSource)

        If oSource.Capabilities.ICapXferMech.GetValues.Contains(NTwain.Data.XferMech.Native) Then
            oSource.Capabilities.ICapXferMech.SetValue(NTwain.Data.XferMech.Native)

        ElseIf oSource.Capabilities.ICapXferMech.GetValues.Contains(NTwain.Data.XferMech.File) Then
            oSource.Capabilities.ICapXferMech.SetValue(NTwain.Data.XferMech.File)

            If oSource.Capabilities.ICapImageFileFormat.GetValues.Contains(NTwain.Data.FileFormat.Bmp) Then
                Dim oFileXfer As New NTwain.Data.TWSetupFileXfer
                With oFileXfer
                    .Format = NTwain.Data.FileFormat.Bmp
                    .FileName = IO.Path.GetTempPath + "TwainCapture.bmp"
                    If IO.File.Exists(.FileName) Then
                        IO.File.Delete(.FileName)
                    End If
                End With

                oSource.DGControl.SetupFileXfer.Set(oFileXfer)
            End If
        End If
    End Sub
    Private Sub Twain_DataTransferredHandler(sender As Object, e As NTwain.DataTransferredEventArgs) Handles Session.DataTransferred
        ' converts the scanned data to a bitmapsource
        Dim oImage As Imaging.BitmapSource = Nothing

        If e.NativeData <> IntPtr.Zero Then
            Using oStream As IO.Stream = e.GetNativeImageStream()
                If Not IsNothing(oStream) Then
                    Dim oBitmapImage As New Imaging.BitmapImage
                    oBitmapImage.BeginInit()
                    oBitmapImage.CacheOption = Imaging.BitmapCacheOption.OnLoad
                    oBitmapImage.DecodePixelHeight = 0
                    oBitmapImage.DecodePixelWidth = 0
                    oBitmapImage.StreamSource = oStream
                    oBitmapImage.EndInit()
                    If (oBitmapImage.CanFreeze) Then
                        oBitmapImage.Freeze()
                    End If

                    oImage = oBitmapImage
                End If
            End Using
        ElseIf Not String.IsNullOrEmpty(e.FileDataPath) Then
            oImage = New Imaging.BitmapImage(New Uri(e.FileDataPath))

            If IO.File.Exists(e.FileDataPath) Then
                IO.File.Delete(e.FileDataPath)
            End If
        End If

        Using oBitmap As System.Drawing.Bitmap = BitmapSourceToBitmap(oImage, CInt(oImage.DpiX))
            ScanQueue.Enqueue(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.Image, BitmapConvertGrayscale(oBitmap)))
        End Using
    End Sub
    Private Sub Twain_TransferErrorHandler(sender As Object, e As NTwain.TransferErrorEventArgs) Handles Session.TransferError
        ' scan error
        ScanQueue.Enqueue(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.ScanError, "Scan error: " + e.Exception.Message))
    End Sub
    Private Sub Twain_StateChangedHandler(senser As Object, e As EventArgs) Handles Session.StateChanged
        ' change in state
        If ScannerIsActive And Session.State = 4 Then
            ScannerIsActive = False
            Dim oSource As NTwain.DataSource = GetTWAINDataSource(SelectedScannerSource)
            oSource.Close()
            Session.Close()
            ScanQueue.Enqueue(New Tuple(Of Twain32Enumerations.ScanProgress, Object)(Twain32Enumerations.ScanProgress.Complete, Nothing))
        End If
    End Sub
#Region "Local Functions"
    Private Function BitmapSourceToBitmap(ByVal oBitmapSource As Imaging.BitmapSource, ByVal oResolution As Single) As System.Drawing.Bitmap
        ' converts WPF bitmapsource to GDI bitmap
        Dim oFormatConvertedBitmap As New Imaging.FormatConvertedBitmap
        oFormatConvertedBitmap.BeginInit()
        oFormatConvertedBitmap.Source = oBitmapSource
        oFormatConvertedBitmap.DestinationFormat = PixelFormats.Bgra32
        oFormatConvertedBitmap.EndInit()

        Dim oBitmap As New System.Drawing.Bitmap(oFormatConvertedBitmap.PixelWidth, oFormatConvertedBitmap.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim oBitmapData = oBitmap.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, oBitmap.Size), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        oFormatConvertedBitmap.CopyPixels(Int32Rect.Empty, oBitmapData.Scan0, oBitmapData.Height * oBitmapData.Stride, oBitmapData.Stride)
        oBitmap.UnlockBits(oBitmapData)
        oBitmap.SetResolution(oResolution, oResolution)
        Return oBitmap
    End Function
    Public Function BitmapConvertGrayscale(ByVal oImage As System.Drawing.Bitmap) As System.Drawing.Bitmap
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
    Private Sub SerializeDataContractStream(Of T)(ByRef oStream As IO.Stream, ByVal data As T)
        ' serialise to stream
        Dim oKnownTypes As New List(Of Type)
        oKnownTypes.AddRange(Twain32Functions.GetTwainKnownTypes)

        Dim oDataContractSerializer As New DataContractSerializer(GetType(T), oKnownTypes)
        oDataContractSerializer.WriteObject(oStream, data)
    End Sub
    Private Function DeserializeDataContractStream(Of T)(ByRef oStream As IO.Stream) As T
        ' deserialise from stream
        Dim oXmlDictionaryReaderQuotas As New Xml.XmlDictionaryReaderQuotas()
        oXmlDictionaryReaderQuotas.MaxArrayLength = 100000000
        Dim oXmlDictionaryReader As Xml.XmlDictionaryReader = Xml.XmlDictionaryReader.CreateTextReader(oStream, oXmlDictionaryReaderQuotas)

        Dim theObject As T = Nothing
        Try
            Dim oKnownTypes As New List(Of Type)
            oKnownTypes.AddRange(Twain32Functions.GetTwainKnownTypes)

            Dim oDataContractSerializer As New DataContractSerializer(GetType(T), oKnownTypes, Integer.MaxValue, False, True, Nothing)
            theObject = oDataContractSerializer.ReadObject(oXmlDictionaryReader, True)
        Catch ex As SerializationException
        End Try

        oXmlDictionaryReader.Close()
        Return theObject
    End Function
    Private Function SerializeDataContractText(Of T)(ByVal data As T) As String
        ' serialise using data contract serialiser
        ' returns base64 text
        Using oMemoryStream As New IO.MemoryStream
            SerializeDataContractStream(oMemoryStream, data)
            Dim bByteArray As Byte() = oMemoryStream.ToArray
            Return Convert.ToBase64String(bByteArray)
        End Using
    End Function
    Private Function DeserializeDataContractText(Of T)(ByVal sBase64String As String) As T
        ' deserialise from base64 text
        Using oMemoryStream As New IO.MemoryStream(Convert.FromBase64String(sBase64String))
            Return DeserializeDataContractStream(Of T)(oMemoryStream)
        End Using
    End Function
#End Region
End Module

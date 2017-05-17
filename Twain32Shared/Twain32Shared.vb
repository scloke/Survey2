Imports System.Runtime.Serialization

Public Class Twain32Shared
    Public Class Twain32Constants
        Public Const ValidateString As String = "Validate"
        Public Const EmptyString As String = ""
        Public Const XString As String = "X"
        Public Const OKString As String = "OK"
        Public Const InitString As String = "Init"
        Public Const SelectScannerString As String = "SelectScanner"
        Public Const GetScannerSourcesString As String = "GetScannerSources"
        Public Const ConfigureString As String = "Configure"
        Public Const StartScanString As String = "StartScan"
        Public Const GetScanImageString As String = "GetScanImage"
        Public Const ProcessBarcodeString As String = "ProcessBarcode"
        Public Const MaxErrorCount As Integer = 10
        Public Shared Twain32GUID As New Guid("{d73b4483-1b6c-4031-ad36-75e1b490bbab}")
    End Class
    Public Class Twain32Enumerations
        <DataContract> Public Enum ScannerSource As Integer
            <EnumMember> None = 0
            <EnumMember> TWAIN
            <EnumMember> WIA
        End Enum
        <DataContract> Public Enum ScanProgress As Integer
            <EnumMember> None = 0
            <EnumMember> NoError
            <EnumMember> ScanError
            <EnumMember> Image
            <EnumMember> Complete
        End Enum
    End Class
    Public Class Twain32Functions
        Public Shared Function GetTwainKnownTypes() As List(Of Type)
            ' gets list of known types
            Dim oKnownTypes As New List(Of Type)
            oKnownTypes.Add(GetType(System.Drawing.Bitmap))
            Return oKnownTypes
        End Function
    End Class
    Public Class StreamString
        Private m_WriteStream As IO.Stream
        Private m_ReadStream As IO.Stream
        Private m_Encoding As Text.UnicodeEncoding

        Public Sub New(ByRef oWriteStream As IO.Stream, ByRef oReadStream As IO.Stream)
            m_WriteStream = oWriteStream
            m_ReadStream = oReadStream
            m_Encoding = New Text.UnicodeEncoding
        End Sub
        Public Function ReadString() As String
            Try
                Dim byteArray(3) As Byte
                For i = 0 To 3
                    Dim iByte As Integer = m_ReadStream.ReadByte()
                    If iByte = -1 Then
                        Return String.Empty
                    Else
                        byteArray(i) = iByte
                    End If
                Next
                Dim iBufferLength As Integer = BitConverter.ToInt32(byteArray, 0)
                Dim bBuffer As Array = Array.CreateInstance(GetType(Byte), iBufferLength)
                Dim iBufferReadCount As Integer = m_ReadStream.Read(bBuffer, 0, iBufferLength)

                If iBufferReadCount = iBufferLength Then
                    Return m_Encoding.GetString(bBuffer)
                Else
                    Return String.Empty
                End If
            Catch ex As OverflowException
                Return String.Empty
            End Try
        End Function
        Public Sub WriteString(outString As String)
            Dim bBuffer As Byte() = m_Encoding.GetBytes(outString)
            Dim iBufferLength As Integer = bBuffer.Length
            Dim byteArray As Byte() = BitConverter.GetBytes(iBufferLength)

            For i = 0 To 3
                m_WriteStream.WriteByte(byteArray(i))
            Next
            m_WriteStream.Write(bBuffer, 0, iBufferLength)
            m_WriteStream.Flush()
        End Sub
    End Class
End Class
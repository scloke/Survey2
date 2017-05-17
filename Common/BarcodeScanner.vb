Imports System
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text
Imports System.Collections.Specialized
Imports System.Collections

''' <summary>
''' VB.Net translation of BarcodeImaging class provided by Alessandro Gubbiotti (princes@tiscali.it).
''' 
''' Translated from the c# class by Berend Engelbrecht (b.engelbrecht@gmail.com).
''' See http://www.codeproject.com/KB/graphics/BarcodeImaging3.aspx
''' 
''' Parts of this class are derived from an earlier code project by James Fitch (qlipoth).
''' Used and published with permission of the original author.
''' 
''' Licensed under The Code Project Open License (CPOL). 
''' See http://www.codeproject.com/info/cpol10.aspx
''' </summary>
Public Class BarcodeScanner
#Region "Public types (used in public function parameters)"
    ''' <summary>
    ''' Used to specify what barcode type(s) to detect.
    ''' </summary>
    '''
    Public Enum BarcodeType
        ''' <summary>Not specified</summary>
        None = 0
        ''' <summary>Code39</summary>
        Code39 = 1
        ''' <summary>EAN/UPC</summary>
        EAN = 2
        ''' <summary>Code128</summary>
        Code128 = 4
        ''' <summary>Use BarcodeType.All for all supported types</summary>
        All = Code39 Or EAN Or Code128

        ' Note: Extend this enum with new types numbered as 8, 16, 32 ... ,
        '       so that we can use bitwise logic: All = Code39 | EAN | <your favorite type here> | ...
    End Enum

    ''' <summary>
    ''' Used to specify whether to scan a page in vertical direction,
    ''' horizontally, or both.
    ''' </summary>
    Public Enum ScanDirection
        ''' <summary>Scan top-to-bottom</summary>
        Vertical = 1
        ''' <summary>Scan left-to-right</summary>
        Horizontal = 2
    End Enum
#End Region

#Region "Private constants and types"
    Private Structure BarcodeZone
        Public Start As Integer
        Public [End] As Integer
    End Structure

    ''' <summary>
    ''' Structure used to return the processed data from an image
    ''' </summary>
    Private Class histogramResult
        ''' <summary>Averaged image brightness values over one scanned band</summary>
        Public histogram As Byte()
        ''' <summary>Minimum brightness (darkest)</summary>
        Public min As Byte
        ' 
        ''' <summary>Maximum brightness (lightest)</summary>
        Public max As Byte

        Public threshold As Byte
        ' threshold brightness to detect change from "light" to "dark" color
        Public lightnarrowbarwidth As Single
        ' narrow bar width for light bars
        Public darknarrowbarwidth As Single
        ' narrow bar width for dark bars
        Public lightwiderbarwidth As Single
        ' width of most common wider bar for light bars
        Public darkwiderbarwidth As Single
        ' width of most common wider bar for dark bars
        Public zones As BarcodeZone()
        ' list of zones on the current band that might contain barcode data
    End Class

    ' General
    Private Const GAPFACTOR As Integer = 48
    ' width of quiet zone compared to narrow bar
    Private Const MINNARROWBARCOUNT As Integer = 4
    ' minimum occurence of a narrow bar width
    ' Code39
    Private Const STARPATTERN As String = "nwnnwnwnn"
    ' the pattern representing a star in Code39
    Private Const WIDEFACTOR As Single = 2.0F
    ' minimum width of wide bar compared to narrow bar
    Private Const MINPATTERNLENGTH As Integer = 10
    ' length of one barcode digit + gap
    ' Code128
    Private Const CODE128START As Integer = 103
    ' Startcodes have index >= 103
    Private Const CODE128STOP As Integer = 106
    ' Stopcode has index 106
    Private Const CODE128C As Integer = 99
    ' Switch to code page C
    Private Const CODE128B As Integer = 100
    ' Switch to code page B
    Private Const CODE128A As Integer = 101
    ' Switch to code page A
#End Region

#Region "Private member variables"
    Private Shared m_bUseBarcodeZones As Boolean = True
#End Region

#Region "Public Shared properties used in configuration"
    ''' <summary>
    ''' Set UseBarcodeZones to false if you do not need this feature.
    ''' Barcode regions improve detection of multiple barcodes on one scan line,
    ''' but have a significant performance impact.
    ''' </summary>
    Public Shared Property UseBarcodeZones() As Boolean
        Get
            Return m_bUseBarcodeZones
        End Get
        Set(ByVal value As Boolean)
            m_bUseBarcodeZones = value
        End Set
    End Property
#End Region

#Region "Public Shared methods"
    ''' <summary>
    ''' FullScanPage does a full scan of the active frame in the passed bitmap. This function
    ''' will scan both vertically and horizontally. 
    ''' </summary>
    ''' <remarks>
    ''' 
    ''' Use ScanPage instead of FullScanPage if you want to scan in one direction only, 
    ''' or only for specific barcode types.
    ''' 
    ''' For a multi-page tiff only one page is scanned. By default, the first page is used, but
    ''' you can scan other pages by calling bmp.SelectActiveFrame(FrameDimension.Page, pagenumber)
    ''' before calling FullScanPage.
    ''' </remarks>
    ''' <param name="CodesRead">Will contain detected barcode strings when the function returns</param>
    ''' <param name="bmp">Input bitmap</param>
    ''' <param name="numscans">Number of passes that must be made over the page. 
    ''' 50 - 100 usually gives a good result.</param>
    Public Shared Sub FullScanPage(ByRef CodesRead As System.Collections.ArrayList, ByVal bmp As Bitmap, ByVal numscans As Integer, ByVal types As BarcodeType)
        ScanPage(CodesRead, bmp, numscans, ScanDirection.Vertical, types)
        ScanPage(CodesRead, bmp, numscans, ScanDirection.Horizontal, types)
    End Sub

    ''' <summary>
    ''' Scans the active frame in the passed bitmap for barcodes.
    ''' </summary>
    ''' <param name="CodesRead">Will contain detected barcode strings when the function returns</param>
    ''' <param name="bmp">Input bitmap</param>
    ''' <param name="numscans">Number of passes that must be made over the page. 
    ''' 50 - 100 usually gives a good result.</param>
    ''' <param name="direction">Scan direction</param>
    ''' <param name="types">Barcode types. Pass BarcodeType.All, or you can specify a list of types,
    ''' e.g., BarcodeType.Code39 | BarcodeType.EAN</param>
    Public Shared Sub ScanPage(ByRef CodesRead As System.Collections.ArrayList, ByVal bmp As Bitmap, ByVal numscans As Integer, ByVal direction As ScanDirection, ByVal types As BarcodeType)
        Dim iHeight As Integer, iWidth As Integer
        If direction = ScanDirection.Horizontal Then
            iHeight = bmp.Width
            iWidth = bmp.Height
        Else
            iHeight = bmp.Height
            iWidth = bmp.Width
        End If
        If numscans > iHeight Then numscans = iHeight ' fix for doing full scan on small images
        For i As Integer = 0 To numscans - 1
            Dim y1 As Integer = (i * iHeight) \ numscans
            Dim y2 As Integer = ((i + 1) * iHeight) \ numscans
            Dim sCodesRead As String = ReadBarcodes(bmp, y1, y2, direction, types)

            If (sCodesRead <> Nothing) AndAlso (sCodesRead.Length > 0) Then
                Dim asCodes As String() = sCodesRead.Split("|")
                For Each sCode As String In asCodes
                    If Not CodesRead.Contains(sCode) Then
                        CodesRead.Add(sCode)
                    End If
                Next
            End If
        Next i
    End Sub

    ''' <summary>
    ''' Scans one band in the passed bitmap for barcodes. 
    ''' </summary>
    ''' <param name="bmp">Input bitmap</param>
    ''' <param name="start">Start coordinate</param>
    ''' <param name="end">End coordinate</param>
    ''' <param name="direction">
    ''' ScanDirection.Vertical: a horizontal band across the page will be examined 
    ''' and start,end should be valid y-coordinates.
    ''' ScanDirection.Horizontal: a vertical band across the page will be examined 
    ''' and start,end should be valid x-coordinates.
    ''' </param>
    ''' <param name="types">Barcode types to be found</param>
    ''' <returns>Pipe-separated list of barcodes, empty string if none were detected</returns>
    Public Shared Function ReadBarcodes(ByVal bmp As Bitmap, ByVal start As Integer, ByVal [end] As Integer, ByVal direction As ScanDirection, ByVal types As BarcodeType) As String
        Dim sBarCodes As String = "|" ' will hold return values

        ' To find a horizontal barcode, find the vertical histogram to find individual barcodes, 
        ' then get the vertical histogram to decode each
        Dim vertHist As histogramResult = verticalHistogram(bmp, start, [end], direction)

        ' Get the light/dark bar patterns.
        ' GetBarPatterns returns the bar pattern in 2 formats: 
        '
        '   sbCode39Pattern: for Code39 (only distinguishes narrow bars "n" and wide bars "w")
        '   sbEANPattern: for EAN (distinguishes bar widths 1, 2, 3, 4 and L/G-code)
        '
        Dim sbCode39Pattern As New StringBuilder
        Dim sbEANPattern As New StringBuilder
        GetBarPatterns(vertHist, sbCode39Pattern, sbEANPattern)

        ' We now have a barcode in terms of narrow & wide bars... Parse it!
        If (sbCode39Pattern.Length > 0) OrElse (sbEANPattern.Length > 0) Then
            For iPass As Integer = 0 To 1
                If (types And BarcodeType.Code39) <> BarcodeType.None Then
                    ' if caller wanted Code39
                    Dim sCode39 As String = ParseCode39(sbCode39Pattern)
                    If sCode39.Length > 0 Then
                        sBarCodes += sCode39 + "|"
                    End If
                End If
                If (types And BarcodeType.EAN) <> BarcodeType.None Then
                    ' if caller wanted EAN
                    Dim sEAN As String = ParseEAN(sbEANPattern)
                    If sEAN.Length > 0 Then
                        sBarCodes += sEAN + "|"
                    End If
                End If
                If (types And BarcodeType.Code128) <> BarcodeType.None Then
                    ' if caller wanted Code128
                    ' Note: Code128 uses same bar width measurement data as EAN
                    Dim sCode128 As String = ParseCode128(sbEANPattern)
                    If sCode128.Length > 0 Then
                        sBarCodes += sCode128 + "|"
                    End If
                End If

                ' Reverse the bar pattern arrays to read again in the mirror direction
                If iPass = 0 Then
                    sbCode39Pattern = ReversePattern(sbCode39Pattern)
                    sbEANPattern = ReversePattern(sbEANPattern)
                End If
            Next iPass
        End If

        ' Return pipe-separated list of found barcodes, if any
        If sBarCodes.Length > 2 Then
            Return sBarCodes.Substring(1, sBarCodes.Length - 2)
        End If
        Return String.Empty
    End Function
#End Region
#Region "Private functions"
#Region "General"
    ''' <summary>
    ''' Scans for patterns of bars and returns them encoded as strings in the passed
    ''' string builder parameters.
    ''' </summary>
    ''' <param name="hist">Input data containing picture information for the scan line</param>
    ''' <param name="sbCode39Pattern">Returns string containing "w" for wide bars and "n" for narrow bars</param>
    ''' <param name="sbEANPattern">Returns string with numbers designating relative bar widths compared to 
    ''' narrowest bar: "1" to "4" are valid widths that can be present in an EAN barcode</param>
    ''' <remarks>In both output strings, "|"-characters will be inserted to indicate gaps 
    ''' in the input data.</remarks>
    Private Shared Sub GetBarPatterns(ByRef hist As histogramResult, ByRef sbCode39Pattern As StringBuilder, ByRef sbEANPattern As StringBuilder)
        ' Initialize return data
        sbCode39Pattern = New StringBuilder()
        sbEANPattern = New StringBuilder()

        If Not hist.zones Is Nothing Then
            ' if barcode zones were found along the scan line
            For iZone As Integer = 0 To hist.zones.Length - 1
                ' Recalculate bar width distribution if more than one zone is present, it could differ per zone
                If hist.zones.Length > 1 Then
                    GetBarWidthDistribution(hist, hist.zones(iZone).Start, hist.zones(iZone).[End])
                End If

                ' Check the calculated narrow bar widths. If they are very different, the pattern is
                ' unlikely to be a bar code
                If ValidBars(hist) Then
                    ' add gap separator to output patterns
                    sbCode39Pattern.Append("|")
                    sbEANPattern.Append("|")

                    ' Variables needed to check for
                    Dim iBarStart As Integer = 0
                    Dim bDarkBar As Boolean = (hist.histogram(0) <= hist.threshold)

                    ' Find the narrow and wide bars
                    For i As Integer = 1 To hist.histogram.Length - 1
                        Dim bDark As Boolean = (hist.histogram(i) <= hist.threshold)
                        If bDark <> bDarkBar Then
                            Dim iBarWidth As Integer = i - iBarStart
                            Dim fNarrowBarWidth As Single = If(bDarkBar, hist.darknarrowbarwidth, hist.lightnarrowbarwidth)
                            Dim fWiderBarWidth As Single = If(bDarkBar, hist.darkwiderbarwidth, hist.lightwiderbarwidth)
                            If IsWideBar(iBarWidth, fNarrowBarWidth, fWiderBarWidth) Then
                                ' The bar was wider than the narrow bar width, it's a wide bar or a gap
                                If iBarWidth > (GAPFACTOR * fNarrowBarWidth) Then
                                    sbCode39Pattern.Append("|")
                                    sbEANPattern.Append("|")
                                Else
                                    sbCode39Pattern.Append("w")
                                    AppendEAN(sbEANPattern, iBarWidth, fNarrowBarWidth)
                                End If
                            Else
                                ' The bar is a narrow bar
                                sbCode39Pattern.Append("n")
                                AppendEAN(sbEANPattern, iBarWidth, fNarrowBarWidth)
                            End If
                            bDarkBar = bDark
                            iBarStart = i
                        End If
                    Next i
                End If
            Next iZone
        End If
    End Sub

    ''' <summary>
    ''' Returns true if the bar appears to be "wide".
    ''' </summary>
    ''' <param name="iBarWidth">measured bar width in pixels</param>
    ''' <param name="fNarrowBarWidth">average narrow bar width</param>
    ''' <param name="fWiderBarWidth">average width of next wider bar</param>
    ''' <returns></returns>
    Private Shared Function IsWideBar(ByVal iBarWidth As Integer, ByVal fNarrowBarWidth As Single, ByVal fWiderBarWidth As Single) As Boolean
        If fNarrowBarWidth < 4 Then
            Return (iBarWidth > (WIDEFACTOR * fNarrowBarWidth))
        End If
        Return (iBarWidth >= fWiderBarWidth) OrElse ((fWiderBarWidth - iBarWidth) < (iBarWidth - fNarrowBarWidth))
    End Function

    ''' <summary>
    ''' Checks if dark and light narrow bar widths are in agreement.
    ''' </summary>
    ''' <param name="hist">barcode data</param>
    ''' <returns>true if barcode data is valid</returns>
    Private Shared Function ValidBars(ByRef hist As histogramResult) As Boolean
        Dim fCompNarrowBarWidths As Single = hist.lightnarrowbarwidth / hist.darknarrowbarwidth
        Dim fCompWiderBarWidths As Single = hist.lightwiderbarwidth / hist.darkwiderbarwidth
        Return ((fCompNarrowBarWidths >= 0.5) AndAlso (fCompNarrowBarWidths <= 2) AndAlso (fCompWiderBarWidths >= 0.5) AndAlso (fCompWiderBarWidths <= 2) AndAlso (hist.darkwiderbarwidth / hist.darknarrowbarwidth >= 1.5) AndAlso (hist.lightwiderbarwidth / hist.lightnarrowbarwidth >= 1.5))
    End Function

    ''' <summary>
    ''' Used by ReadBarcodes to reverse a bar pattern. 
    ''' </summary>
    ''' <param name="sbPattern">String builder containing a bar pattern string</param>
    ''' <returns>String builder containing the reverse of the input string</returns>
    Private Shared Function ReversePattern(ByVal sbPattern As StringBuilder) As StringBuilder
        If sbPattern.Length > 0 Then
            Dim acPattern As Char() = sbPattern.ToString().ToCharArray()
            Array.Reverse(acPattern)
            sbPattern = New StringBuilder(acPattern.Length)
            sbPattern.Append(acPattern)
        End If
        Return sbPattern
    End Function

    ''' <summary>
    ''' Vertical histogram of an image
    ''' </summary>
    ''' <param name="bmp">Bitmap</param>
    ''' <param name="start">Start coordinate of band to be scanned</param>
    ''' <param name="end">End coordinate of band to be scanned</param>
    ''' <param name="direction">
    ''' ScanDirection.Vertical: start and end denote y-coordinates.
    ''' ScanDirection.Horizontal: start and end denote x-coordinates.
    ''' </param>
    ''' <returns>histogramResult, containing average brightness values across the scan line</returns>
    Private Shared Function verticalHistogram(ByVal bmp As Bitmap, ByVal start As Integer, ByVal [end] As Integer, ByVal direction As ScanDirection) As histogramResult
        ' convert the pixel format of the bitmap to something that we can handle
        Dim pf As PixelFormat = CheckSupportedPixelFormat(bmp.PixelFormat)
        Dim bmData As BitmapData
        Dim xMax As Integer, yMax As Integer

        If direction = ScanDirection.Horizontal Then
            bmData = bmp.LockBits(New Rectangle(start, 0, [end] - start, bmp.Height), ImageLockMode.[ReadOnly], pf)
            xMax = bmData.Height
            yMax = [end] - start
        Else
            bmData = bmp.LockBits(New Rectangle(0, start, bmp.Width, [end] - start), ImageLockMode.[ReadOnly], pf)
            xMax = bmp.Width
            yMax = bmData.Height
        End If

        ' Create the return value
        Dim histResult As Byte() = New Byte(0 To xMax + 1) {}
        ' add 2 to simulate light-colored background pixels at start and end of scanline
        Dim vertSum As UInteger() = New UInteger(0 To xMax - 1) {}

        ' stride is offset between horizontal lines in p 
        Dim stride As Integer = bmData.Stride

        For y As Integer = 0 To yMax - 1
            ' Add up all the pixel values vertically
            For x As Integer = 0 To xMax - 1
                If direction = ScanDirection.Horizontal Then
                    vertSum(x) += getpixelbrightness(bmData, pf, y, x)
                Else
                    vertSum(x) += getpixelbrightness(bmData, pf, x, y)
                End If
            Next x
        Next y

        bmp.UnlockBits(bmData)

        ' Now get the average of the row by dividing the pixel by num pixels
        Dim iDivider As Integer = [end] - start
        If pf <> PixelFormat.Format1bppIndexed Then
            iDivider *= IIf(pf = PixelFormat.Format24bppRgb, 3, 4)
        End If

        Dim maxValue As Byte = Byte.MinValue
        ' Start the max value at zero
        Dim minValue As Byte = Byte.MaxValue
        ' Start the min value at the absolute maximum
        For i As Integer = 1 To xMax
            ' note: intentionally skips first pixel in histResult
            histResult(i) = vertSum(i - 1) \ iDivider
            'Save the max value for later
            If histResult(i) > maxValue Then
                maxValue = histResult(i)
            End If
            ' Save the min value for later
            If histResult(i) < minValue Then
                minValue = histResult(i)
            End If
        Next i

        ' Set first and last pixel to "white", i.e., maximum intensity
        histResult(0) = maxValue
        histResult(xMax + 1) = maxValue

        Dim retVal As New histogramResult()
        retVal.histogram = histResult
        retVal.max = maxValue
        retVal.min = minValue

        ' Now we have the brightness distribution along the scan band, try to find the distribution of bar widths.
        retVal.threshold = minValue + ((maxValue - minValue) >> 1)
        GetBarWidthDistribution(retVal, 0, retVal.histogram.Length)

        ' Now that we know the narrow bar width, lets look for barcode zones.
        ' The image could have more than one barcode in the same band, with 
        ' different bar widths.
        FindBarcodeZones(retVal)
        Return retVal
    End Function

    ''' <summary>
    ''' Gets the bar width distribution and calculates narrow bar width over the specified
    ''' range of the histogramResult. A histogramResult could have multiple ranges, separated 
    ''' by quiet zones.
    ''' </summary>
    ''' <param name="hist">histogramResult data</param>
    ''' <param name="iStart">start coordinate to be considered</param>
    ''' <param name="iEnd">end coordinate + 1</param>
    Private Shared Sub GetBarWidthDistribution(ByRef hist As histogramResult, ByVal iStart As Integer, ByVal iEnd As Integer)
        Dim hdLightBars As New HybridDictionary()
        Dim hdDarkBars As New HybridDictionary()
        Dim bDarkBar As Boolean = (hist.histogram(iStart) <= hist.threshold)
        Dim iBarStart As Integer = 0

        For i As Integer = iStart + 1 To iEnd - 1
            Dim bDark As Boolean = (hist.histogram(i) <= hist.threshold)
            If bDark <> bDarkBar Then
                Dim iBarWidth As Integer = i - iBarStart
                If bDarkBar Then
                    If Not hdDarkBars.Contains(iBarWidth) Then
                        hdDarkBars.Add(iBarWidth, 1)
                    Else
                        hdDarkBars(iBarWidth) = DirectCast(hdDarkBars(iBarWidth), Integer) + 1
                    End If
                Else
                    If Not hdLightBars.Contains(iBarWidth) Then
                        hdLightBars.Add(iBarWidth, 1)
                    Else
                        hdLightBars(iBarWidth) = DirectCast(hdLightBars(iBarWidth), Integer) + 1
                    End If
                End If
                bDarkBar = bDark
                iBarStart = i
            End If
        Next i

        ' Now get the most common bar widths
        CalcNarrowBarWidth(hdLightBars, hist.lightnarrowbarwidth, hist.lightwiderbarwidth)
        CalcNarrowBarWidth(hdDarkBars, hist.darknarrowbarwidth, hist.darkwiderbarwidth)
    End Sub

    Private Shared Sub CalcNarrowBarWidth(ByVal hdBarWidths As HybridDictionary, ByRef fNarrowBarWidth As Single, ByRef fWiderBarWidth As Single)
        fNarrowBarWidth = 1.0F
        fWiderBarWidth = 2.0F
        If hdBarWidths.Count > 1 Then
            ' we expect at least two different bar widths in supported barcodes
            Dim aiWidths As Integer() = New Integer(0 To hdBarWidths.Count - 1) {}
            Dim aiCounts As Integer() = New Integer(0 To hdBarWidths.Count - 1) {}
            Dim i As Integer = 0
            For Each iKey As Integer In hdBarWidths.Keys
                aiWidths(i) = iKey
                aiCounts(i) = hdBarWidths(iKey)
                i += 1
            Next
            Array.Sort(aiWidths, aiCounts)

            ' walk from lowest to highest width. The narrowest bar should occur at least 4 times
            fNarrowBarWidth = aiWidths(0)
            fWiderBarWidth = WIDEFACTOR * fNarrowBarWidth
            For i = 0 To aiCounts.Length - 1
                If aiCounts(i) >= MINNARROWBARCOUNT Then
                    fNarrowBarWidth = aiWidths(i)
                    If fNarrowBarWidth < 3 Then
                        fWiderBarWidth = WIDEFACTOR * fNarrowBarWidth
                    Else
                        ' if the width is not singular, look for the most common width in the neighbourhood
                        Dim fCount As Single
                        FindPeakWidth(i, aiWidths, aiCounts, fNarrowBarWidth, fCount)
                        fWiderBarWidth = WIDEFACTOR * fNarrowBarWidth

                        If fNarrowBarWidth >= 6 Then
                            ' ... and for the next wider common bar width if the barcode is fairly large
                            Dim fMaxCount As Single = 0.0F
                            For j As Integer = i + 1 To aiCounts.Length - 1
                                Dim fNextWidth As Single, fNextCount As Single
                                FindPeakWidth(j, aiWidths, aiCounts, fNextWidth, fNextCount)
                                If (fNextWidth / fNarrowBarWidth) > 1.5 Then
                                    If fNextCount > fMaxCount Then
                                        fWiderBarWidth = fNextWidth
                                        fMaxCount = fNextCount
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next j
                        End If
                    End If
                    Exit For
                End If
            Next i
        End If
    End Sub

    Shared Sub FindPeakWidth(ByVal i As Integer, ByRef aiWidths As Integer(), ByRef aiCounts As Integer(), ByRef fWidth As Single, ByRef fCount As Single)
        fWidth = 0.0F
        fCount = 0.0F
        Dim iSamples As Integer = 0
        For j As Integer = i - 1 To i + 1
            If (j >= 0) AndAlso (j < aiWidths.Length) AndAlso (Math.Abs(aiWidths(j) - aiWidths(i)) = Math.Abs(j - i)) Then
                iSamples += 1
                fCount += aiCounts(j)
                fWidth += aiWidths(j) * aiCounts(j)
            End If
        Next j
        fWidth /= fCount
        fCount /= iSamples
    End Sub

    ''' <summary>
    ''' FindBarcodeZones looks for barcode zones in the current band. 
    ''' We look for white space that is more than GAPFACTOR * narrowbarwidth
    ''' separating two zones. For narrowbarwidth we take the maximum of the 
    ''' dark and light narrow bar width.
    ''' </summary>
    ''' <param name="hist">Data for current image band</param>
    Private Shared Sub FindBarcodeZones(ByRef hist As histogramResult)
        If Not ValidBars(hist) Then
            hist.zones = Nothing
        ElseIf Not UseBarcodeZones Then
            hist.zones = New BarcodeZone(0) {}
            hist.zones(0).Start = 0
            hist.zones(0).[End] = hist.histogram.Length
        Else
            Dim alBarcodeZones As New ArrayList()
            Dim bDarkBar As Boolean = (hist.histogram(0) <= hist.threshold)
            Dim iBarStart As Integer = 0
            Dim iZoneStart As Integer = -1
            Dim iZoneEnd As Integer = -1
            Dim fQuietZoneWidth As Single = GAPFACTOR * (hist.darknarrowbarwidth + hist.lightnarrowbarwidth) / 2
            Dim fMinZoneWidth As Single = fQuietZoneWidth

            For i As Integer = 1 To hist.histogram.Length - 1
                Dim bDark As Boolean = (hist.histogram(i) <= hist.threshold)
                If bDark <> bDarkBar Then
                    Dim iBarWidth As Integer = i - iBarStart
                    If Not bDarkBar Then
                        ' This ends a light area
                        If (iZoneStart = -1) OrElse (iBarWidth > fQuietZoneWidth) Then
                            ' the light area can be seen as a quiet zone
                            iZoneEnd = i - (iBarWidth >> 1)

                            ' Check if the active zone is big enough to contain a barcode
                            If (iZoneStart >= 0) AndAlso (iZoneEnd > (iZoneStart + fMinZoneWidth)) Then
                                ' record the barcode zone that ended in the detected quiet zone ...
                                Dim bz As New BarcodeZone()
                                bz.Start = iZoneStart
                                bz.[End] = iZoneEnd
                                alBarcodeZones.Add(bz)

                                ' .. and start a new barcode zone
                                iZoneStart = iZoneEnd
                            End If
                            If iZoneStart = -1 Then
                                iZoneStart = iZoneEnd
                                ' first zone starts here
                            End If
                        End If
                    End If
                    bDarkBar = bDark
                    iBarStart = i
                End If
            Next i
            If iZoneStart >= 0 Then
                Dim bz As New BarcodeZone()
                bz.Start = iZoneStart
                bz.[End] = hist.histogram.Length
                alBarcodeZones.Add(bz)
            End If
            If alBarcodeZones.Count > 0 Then
                hist.zones = DirectCast(alBarcodeZones.ToArray(GetType(BarcodeZone)), BarcodeZone())
            Else
                hist.zones = Nothing
            End If
        End If
    End Sub

    ''' <summary>
    ''' Checks if the supplied pixelFormat is supported, returns the default
    ''' pixel format (PixelFormat.Format24bppRgb) if it isn't supported.
    ''' </summary>
    ''' <param name="pixelFormat">Input pixel format</param>
    ''' <returns>Input pixel format if it is supported, else default.</returns>
    Private Shared Function CheckSupportedPixelFormat(ByVal pixelFormat As PixelFormat) As PixelFormat
        Select Case pixelFormat
            Case PixelFormat.Format1bppIndexed, PixelFormat.Format32bppArgb, PixelFormat.Format32bppRgb
                Return pixelFormat
        End Select
        Return PixelFormat.Format24bppRgb
    End Function

    ''' <summary>
    ''' Calculates pixel brightness for specified pixel in byte array of locked bitmap rectangle.
    ''' For RGB  : returns sum of the three color values.
    ''' For 1bpp : returns 255 for a white pixel, 0 for a black pixel.
    ''' </summary>
    ''' <param name="bmd">Bitmap data</param>
    ''' <param name="pf">Pixel format used in the byte array</param>
    ''' <param name="x">Horizontal coordinate, relative to upper left corner of locked rectangle</param>
    ''' <param name="y">Vertical coordinate, relative to upper left corner of locked rectangle</param>
    ''' <returns></returns>
    Private Shared Function getpixelbrightness(ByVal bmd As BitmapData, ByVal pf As PixelFormat, ByVal x As Integer, ByVal y As Integer) As UShort
        Dim uBrightness As UShort = 0
        Dim offset As Integer
        Dim pixelrange As Integer
        Dim currentpixel As Byte

        Select Case pf
            Case PixelFormat.Format1bppIndexed
                offset = (y * bmd.Stride) + (x >> 3)
                currentpixel = Marshal.ReadByte(bmd.Scan0, offset)

                If ((currentpixel << (x Mod 8)) And 128) <> 0 Then
                    uBrightness = 255
                End If
                Exit Select
            Case Else

                ' 24bpp RGB, 32bpp formats
                pixelrange = (y * bmd.Stride) + (x * If(pf = PixelFormat.Format24bppRgb, 3, 4))
                For offset = pixelrange To pixelrange + 2
                    uBrightness += Marshal.ReadByte(bmd.Scan0, offset)
                Next offset
        End Select
        Return uBrightness
    End Function
#End Region

#Region "Code39-specific"
    ''' <summary>
    ''' Parses Code39 barcodes from the input pattern.
    ''' </summary>
    ''' <param name="sbPattern">Input pattern, should contain "n"-characters to
    ''' indicate narrow bars and "w" to indicate wide bars.</param>
    ''' <returns>Pipe-separated list of barcodes, empty string if none were detected</returns>
    Private Shared Function ParseCode39(ByVal sbPattern As StringBuilder) As String
        ' Each pattern within code 39 is nine bars with one white bar between each pattern
        If sbPattern.Length > 9 Then
            Dim sbBarcodes As New StringBuilder()
            Dim sPattern As String = sbPattern.ToString()
            Dim iStarPattern As Integer = sPattern.IndexOf(STARPATTERN)
            Dim iDimension As Integer = sbPattern.Length \ 10
            ' index of first star barcode in pattern
            While (iStarPattern >= 0) AndAlso (iStarPattern <= sbPattern.Length - 9)
                Dim iPos As Integer = iStarPattern
                Dim iNoise As Integer = 0
                Dim sbData As New StringBuilder(iDimension)
                While iPos <= sbPattern.Length - 9
                    ' Test the next 9 characters from the pattern string
                    Dim sData As String = ParseCode39Pattern(sbPattern.ToString(iPos, 9))

                    If sData = Nothing Then
                        ' no recognizeable data
                        iPos += 1
                        iNoise += 1
                    Else
                        ' record if the data contained a lot of noise before the next valid data
                        If iNoise >= 2 Then
                            sbData.Append("|")
                        End If
                        iNoise = 0
                        ' reset noise counter
                        sbData.Append(sData)
                        iPos += 10
                    End If
                End While
                If sbData.Length > 0 Then
                    ' look for valid Code39 patterns in the data.
                    ' A valid Code39 pattern starts and ends with "*" and does not contain a noise character "|".
                    ' We return a pipe-separated list of these patterns.
                    Dim asBarcodes As String() = sbData.ToString().Split("|")
                    For Each sBarcode As String In asBarcodes
                        If sBarcode.Length > 2 Then
                            Dim iFirstStar As Integer = sBarcode.IndexOf("*")
                            If (iFirstStar >= 0) AndAlso (iFirstStar < sBarcode.Length - 1) Then
                                Dim iSecondStar As Integer = sBarcode.IndexOf("*", iFirstStar + 1)
                                If iSecondStar > iFirstStar + 1 Then
                                    sbBarcodes.Append(sBarcode.Substring(iFirstStar + 1, (iSecondStar - iFirstStar - 1)) + "|")
                                End If
                            End If
                        End If
                    Next
                End If
                ' "nwnnwnwnn" pattern can not occur again before current index + 5
                iStarPattern = sPattern.IndexOf(STARPATTERN, iStarPattern + 5)
            End While
            If sbBarcodes.Length > 1 Then
                Return sbBarcodes.ToString(0, sbBarcodes.Length - 1)
            End If
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Parses bar pattern for one Code39 character.
    ''' </summary>
    ''' <param name="pattern">Pattern to be examined, should be 9 characters</param>
    ''' <returns>Detected character or null</returns>
    Private Shared Function ParseCode39Pattern(ByVal pattern As String) As String
        Select Case pattern
            Case "wnnwnnnnw"
                Return "1"
            Case "nnwwnnnnw"
                Return "2"
            Case "wnwwnnnnn"
                Return "3"
            Case "nnnwwnnnw"
                Return "4"
            Case "wnnwwnnnn"
                Return "5"
            Case "nnwwwnnnn"
                Return "6"
            Case "nnnwnnwnw"
                Return "7"
            Case "wnnwnnwnn"
                Return "8"
            Case "nnwwnnwnn"
                Return "9"
            Case "nnnwwnwnn"
                Return "0"
            Case "wnnnnwnnw"
                Return "A"
            Case "nnwnnwnnw"
                Return "B"
            Case "wnwnnwnnn"
                Return "C"
            Case "nnnnwwnnw"
                Return "D"
            Case "wnnnwwnnn"
                Return "E"
            Case "nnwnwwnnn"
                Return "F"
            Case "nnnnnwwnw"
                Return "G"
            Case "wnnnnwwnn"
                Return "H"
            Case "nnwnnwwnn"
                Return "I"
            Case "nnnnwwwnn"
                Return "J"
            Case "wnnnnnnww"
                Return "K"
            Case "nnwnnnnww"
                Return "L"
            Case "wnwnnnnwn"
                Return "M"
            Case "nnnnwnnww"
                Return "N"
            Case "wnnnwnnwn"
                Return "O"
            Case "nnwnwnnwn"
                Return "P"
            Case "nnnnnnwww"
                Return "Q"
            Case "wnnnnnwwn"
                Return "R"
            Case "nnwnnnwwn"
                Return "S"
            Case "nnnnwnwwn"
                Return "T"
            Case "wwnnnnnnw"
                Return "U"
            Case "nwwnnnnnw"
                Return "V"
            Case "wwwnnnnnn"
                Return "W"
            Case "nwnnwnnnw"
                Return "X"
            Case "wwnnwnnnn"
                Return "Y"
            Case "nwwnwnnnn"
                Return "Z"
            Case "nwnnnnwnw"
                Return "-"
            Case "wwnnnnwnn"
                Return "."
            Case "nwwnnnwnn"
                Return " "
            Case STARPATTERN
                Return "*"
            Case "nwnwnwnnn"
                Return "$"
            Case "nwnwnnnwn"
                Return "/"
            Case "nwnnnwnwn"
                Return "+"
            Case "nnnwnwnwn"
                Return "%"
            Case Else
                Return Nothing
        End Select
    End Function
#End Region

#Region "EAN-specific"
    ''' <summary>
    ''' Parses EAN-barcodes from the input pattern.
    ''' </summary>
    ''' <param name="sbPattern">Input pattern, should contain characters
    ''' "1" .. "4" to indicate valid EAN bar widths.</param>
    ''' <returns>Pipe-separated list of barcodes, empty string if none were detected</returns>
    Private Shared Function ParseEAN(ByVal sbPattern As StringBuilder) As String
        Dim sbEANData As New StringBuilder(32)
        Dim iEANSeparators As Integer = 0
        Dim sEANCode As String = String.Empty

        Dim iPos As Integer = 0
        sbPattern.Append("|")
        ' append one extra "gap" character because separator has only 3 bands
        While iPos <= (sbPattern.Length - 4)
            Dim sEANDigit As String = ParseEANPattern(sbPattern.ToString(iPos, 4), sEANCode, iEANSeparators)
            Select Case sEANDigit
                Case Nothing
                    ' reset on invalid code
                    'iEANSeparators = 0;
                    sEANCode = String.Empty
                    iPos += 1
                    Exit Select
                Case "|"
                    ' EAN separator found. Each EAN code contains three separators.
                    If iEANSeparators >= 3 Then
                        iEANSeparators = 1
                    Else
                        iEANSeparators += 1
                    End If
                    iPos += 3
                    If iEANSeparators = 2 Then
                        ' middle separator has 5 bars
                        iPos += 2
                    ElseIf iEANSeparators = 3 Then
                        ' end of EAN code detected
                        Dim sFirstDigit As String = GetEANFirstDigit(sEANCode)
                        If (sFirstDigit <> Nothing) AndAlso (sEANCode.Length = 12) Then
                            sEANCode = sFirstDigit + sEANCode
                            If sbEANData.Length > 0 Then
                                sbEANData.Append("|")
                            End If
                            sbEANData.Append(sEANCode)
                        End If
                        ' reset after end of code
                        'iEANSeparators = 0;
                        sEANCode = String.Empty
                    End If
                    Exit Select
                Case "S"
                    ' Start of supplemental code after EAN code
                    iPos += 3
                    sEANCode = "S"
                    iEANSeparators = 1
                    Exit Select
                Case Else
                    If iEANSeparators > 0 Then
                        sEANCode += sEANDigit
                        iPos += 4
                        If sEANCode.StartsWith("S") Then
                            ' Each digit of the supplemental code is followed by an additional "11"
                            ' We assume that the code ends if that is no longer the case.
                            If (sbPattern.Length > iPos + 2) AndAlso (sbPattern.ToString(iPos, 2) = "11") Then
                                iPos += 2
                            Else
                                ' Supplemental code ends. It must be either 2 or 5 digits.
                                sEANCode = CheckEANSupplement(sEANCode)
                                If sEANCode.Length > 0 Then
                                    If sbEANData.Length > 0 Then
                                        sbEANData.Append("|")
                                    End If
                                    sbEANData.Append(sEANCode)
                                End If
                                ' reset after end of code
                                iEANSeparators = 0
                                sEANCode = String.Empty
                            End If
                        End If
                    Else
                        iPos += 1
                    End If
                    ' no EAN digit expected before first separator
                    Exit Select
            End Select
        End While
        Return sbEANData.ToString()
    End Function

    ''' <summary>
    ''' Used by GetBarPatterns to derive bar character from bar width.
    ''' </summary>
    ''' <param name="sbEAN">Output pattern</param>
    ''' <param name="nBarWidth">Measured bar width in pixels</param>
    ''' <param name="fNarrowBarWidth">Narrow bar width in pixels</param>
    Private Shared Sub AppendEAN(ByVal sbEAN As StringBuilder, ByVal nBarWidth As Integer, ByVal fNarrowBarWidth As Single)
        Dim nEAN As Integer = Math.Round(nBarWidth / fNarrowBarWidth)
        If nEAN = 5 Then nEAN = 4
        ' bar width could be slightly off due to distortion
        If nEAN < 10 Then
            sbEAN.Append(nEAN.ToString())
        Else
            sbEAN.Append("|")
        End If
    End Sub

    ''' <summary>
    ''' Parses the EAN pattern for one digit or separator
    ''' </summary>
    ''' <param name="sPattern">Pattern to be parsed</param>
    ''' <param name="sEANCode">EAN code found so far</param>
    ''' <param name="iEANSeparators">Number of separators found so far</param>
    ''' <returns>Detected digit type (L/G/R) and digit, "|" for separator
    ''' or null if the pattern was not recognized.</returns>
    Private Shared Function ParseEANPattern(ByVal sPattern As String, ByVal sEANCode As String, ByVal iEANSeparators As Integer) As String
        Dim LRCodes As String() = {"3211", "2221", "2122", "1411", "1132", "1231",
     "1114", "1312", "1213", "3112"}
        Dim GCodes As String() = {"1123", "1222", "2212", "1141", "2311", "1321",
     "4111", "2131", "3121", "2113"}
        If (sPattern <> Nothing) AndAlso (sPattern.Length >= 3) Then
            If sPattern.StartsWith("111") AndAlso ((iEANSeparators * 12) = sEANCode.Length) Then
                Return "|"
            End If
            ' found separator
            If sPattern.StartsWith("112") AndAlso (iEANSeparators = 3) AndAlso (sEANCode.Length = 0) Then
                Return "S"
            End If
            ' found EAN supplemental code
            For i As Integer = 0 To 9
                If sPattern.StartsWith(LRCodes(i)) Then
                    Return (If((iEANSeparators = 2), "R", "L")) + i.ToString()
                End If
                If sPattern.StartsWith(GCodes(i)) Then
                    Return "G" + i.ToString()
                End If
            Next i
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Decodes the L/G-pattern for the left half of the EAN code 
    ''' to derive the first digit. See table in
    ''' http://en.wikipedia.org/wiki/European_Article_Number
    ''' </summary>
    ''' <param name="sEANPattern">
    ''' IN: EAN pattern with digits and L/G/R codes.
    ''' OUT: EAN digits only.
    ''' </param>
    ''' <returns>Detected first digit or null.</returns>
    Private Shared Function GetEANFirstDigit(ByRef sEANPattern As String) As String
        Dim LGPatterns As String() = {"LLLLLL", "LLGLGG", "LLGGLG", "LLGGGL", "LGLLGG", "LGGLLG",
     "LGGGLL", "LGLGLG", "LGLGGL", "LGGLGL"}
        Dim sLG As String = String.Empty
        Dim sDigits As String = String.Empty
        If (sEANPattern <> Nothing) AndAlso (sEANPattern.Length >= 24) Then
            For i As Integer = 0 To 11
                sLG += sEANPattern(2 * i)
                sDigits += sEANPattern(2 * i + 1)
            Next i
            For i = 0 To 9
                If sLG.StartsWith(LGPatterns(i)) Then
                    sEANPattern = sDigits + sEANPattern.Substring(24)
                    Return i.ToString()
                End If
            Next i
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Checks if EAN supplemental code is valid.
    ''' </summary>
    ''' <param name="sEANPattern">Parse result</param>
    ''' <returns>Supplemental code or empty string</returns>
    Private Shared Function CheckEANSupplement(ByVal sEANPattern As String) As String
        Try
            If sEANPattern.StartsWith("S") Then
                Dim sDigits As String = String.Empty
                Dim sLG As String = String.Empty

                For i As Integer = 1 To sEANPattern.Length - 2 Step 2
                    sLG += sEANPattern(i)
                    sDigits += sEANPattern(i + 1)
                Next i

                ' Supplemental code must be either 2 or 5 digits.
                Select Case sDigits.Length
                    Case 2
                        ' Do EAN-2 parity check
                        Dim EAN2Parity As String() = {"LL", "LG", "GL", "GG"}
                        Dim iParity As Integer = Convert.ToInt32(sDigits) Mod 4
                        If sLG <> EAN2Parity(iParity) Then
                            Return String.Empty
                        End If
                        ' parity check failed
                        Exit Select
                    Case 5
                        ' Do EAN-5 checksum validation
                        Dim uCheckSum As UInteger = 0
                        For i = 0 To sDigits.Length - 1
                            uCheckSum += Convert.ToUInt32(sDigits.Substring(i, 1)) * (If(((i And 1) = 0), 3, 9))
                        Next i
                        Dim EAN5CheckSumPattern As String() = {"GGLLL", "GLGLL", "GLLGL", "GLLLG", "LGGLL", "LLGGL",
             "LLLGG", "LGLGL", "LGLLG", "LLGLG"}
                        If sLG <> EAN5CheckSumPattern(uCheckSum Mod 10) Then
                            Return String.Empty
                        End If
                        ' Checksum validation failed
                        Exit Select
                    Case Else
                        Return String.Empty
                End Select
                Return "S" + sDigits
            End If
        Catch ex As Exception
            System.Diagnostics.Trace.Write(ex)
        End Try
        Return String.Empty
    End Function
#End Region

#Region "Code128-specific"
    ''' <summary>
    ''' Parses Code128 barcodes.
    ''' </summary>
    ''' <param name="sbPattern">Input pattern, should contain characters
    ''' "1" .. "4" to indicate valid bar widths.</param>
    ''' <returns>Pipe-separated list of barcodes, empty string if none were detected</returns>
    Private Shared Function ParseCode128(ByVal sbPattern As StringBuilder) As String
        Dim sbCode128Data As New StringBuilder(32)
        Dim sCode128Code As String = String.Empty
        Dim uCheckSum As UInteger = 0
        Dim iCodes As Integer = 0
        Dim iPos As Integer = 0
        Dim cCodePage As Char = "B"

        While iPos <= (sbPattern.Length - 6)
            Dim iResult As Integer = ParseCode128Pattern(sbPattern.ToString(iPos, 6), sCode128Code, uCheckSum, cCodePage, iCodes)
            Select Case iResult
                Case -1
                    ' unrecognized pattern
                    iPos += 1
                    Exit Select
                Case -2
                    ' stop condition, but failed to recognize barcode
                    iPos += 7
                    Exit Select
                Case CODE128STOP
                    iPos += 7
                    If sCode128Code.Length > 0 Then
                        If sbCode128Data.Length > 0 Then
                            sbCode128Data.Append("|")
                        End If
                        sbCode128Data.Append(sCode128Code)
                    End If
                    Exit Select
                Case Else
                    iPos += 6
                    Exit Select
            End Select
        End While
        Return sbCode128Data.ToString()
    End Function

    ''' <summary>
    ''' Parses the Code128 pattern for one barcode character.
    ''' </summary>
    ''' <param name="sPattern">Pattern to be parsed, should be 6 characters</param>
    ''' <param name="sResult">Resulting barcode up to current character</param>
    ''' <param name="uCheckSum">Checksum up to current character</param>
    ''' <param name="cCodePage">Current code page</param>
    ''' <param name="iCodes">Count of barcode characters already parsed (needed for checksum)</param>
    ''' <returns>
    ''' CODE128STOP: end of barcode detected, barcode recognized.
    '''          -2: end of barcode, recognition failed.
    '''          -1: unrecognized pattern.
    '''       other: code 128 character index
    ''' </returns>
    Private Shared Function ParseCode128Pattern(ByVal sPattern As String, ByRef sResult As String, ByRef uCheckSum As UInteger, ByRef cCodePage As Char, ByRef iCodes As Integer) As Integer
        Dim Code128 As String() = {"212222", "222122", "222221", "121223", "121322", "131222",
     "122213", "122312", "132212", "221213", "221312", "231212",
     "112232", "122132", "122231", "113222", "123122", "123221",
     "223211", "221132", "221231", "213212", "223112", "312131",
     "311222", "321122", "321221", "312212", "322112", "322211",
     "212123", "212321", "232121", "111323", "131123", "131321",
     "112313", "132113", "132311", "211313", "231113", "231311",
     "112133", "112331", "132131", "113123", "113321", "133121",
     "313121", "211331", "231131", "213113", "213311", "213131",
     "311123", "311321", "331121", "312113", "312311", "332111",
     "314111", "221411", "431111", "111224", "111422", "121124",
     "121421", "141122", "141221", "112214", "112412", "122114",
     "122411", "142112", "142211", "241211", "221114", "413111",
     "241112", "134111", "111242", "121142", "121241", "114212",
     "124112", "124211", "411212", "421112", "421211", "212141",
     "214121", "412121", "111143", "111341", "131141", "114113",
     "114311", "411113", "411311", "113141", "114131", "311141",
     "411131", "211412", "211214", "211232", "233111"}

        If (sPattern <> Nothing) AndAlso (sPattern.Length >= 6) Then
            For i As Integer = 0 To Code128.Length - 1
                If sPattern.StartsWith(Code128(i)) Then
                    If i = CODE128STOP Then
                        Try
                            Dim iLength As Integer = sResult.Length
                            If iLength > 1 Then
                                Dim cCheckDigit As Char
                                If cCodePage = "C" Then
                                    cCheckDigit = Microsoft.VisualBasic.ChrW(Microsoft.VisualBasic.AscW(sResult.Substring(iLength - 2)) + 32)
                                    sResult = sResult.Substring(0, iLength - 2)
                                Else
                                    cCheckDigit = sResult(iLength - 1)
                                    sResult = sResult.Substring(0, iLength - 1)
                                End If
                                Dim uCheckDigit As UInteger = (Microsoft.VisualBasic.AscW(cCheckDigit) - 32) * iCodes
                                If uCheckSum > uCheckDigit Then
                                    uCheckSum = (uCheckSum - uCheckDigit) Mod 103
                                    If cCheckDigit = Microsoft.VisualBasic.ChrW(uCheckSum + 32) Then
                                        Return CODE128STOP
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            System.Diagnostics.Trace.Write(ex)
                        End Try
                        ' If reach this point, some check failed.
                        ' Reset everything and return error.
                        sResult = String.Empty
                        uCheckSum = 0
                        iCodes = 0
                        Return -2
                    ElseIf i >= CODE128START Then
                        ' Start new code 128 sequence
                        sResult = String.Empty
                        uCheckSum = i
                        cCodePage = Microsoft.VisualBasic.ChrW(Microsoft.VisualBasic.AscW("A") + (i - CODE128START))
                    ElseIf uCheckSum > 0 Then
                        Dim bSkip As Boolean = False
                        Dim cNewCodePage As Char = cCodePage
                        Select Case i
                            Case CODE128C
                                cNewCodePage = "C"
                                Exit Select
                            Case CODE128B
                                cNewCodePage = "B"
                                Exit Select
                            Case CODE128A
                                cNewCodePage = "A"
                                Exit Select
                        End Select
                        If cCodePage <> cNewCodePage Then
                            cCodePage = cNewCodePage
                            bSkip = True
                        End If
                        If Not bSkip Then
                            Select Case cCodePage
                                Case "C"
                                    sResult += i.ToString("00")
                                    Exit Select
                                Case Else
                                    ' Regular character
                                    Dim c As Char = Microsoft.VisualBasic.ChrW(i + 32)
                                    sResult += c
                                    Exit Select
                            End Select
                        End If
                        iCodes += 1
                        uCheckSum += i * iCodes
                    End If
                    Return i
                End If
            Next i
        End If
        sResult = String.Empty
        uCheckSum = 0
        iCodes = 0
        Return -1
    End Function
#End Region
#End Region
End Class
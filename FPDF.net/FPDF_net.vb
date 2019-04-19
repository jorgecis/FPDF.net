Imports System.Runtime.InteropServices

Friend Class FPDFnet
    ' *******************************************************************************
    '  Software: FPDF for VB.NET (Port of FPDF Ver 1.53 by Olivier PLATHEY)         *
    ' * Version:  0.12 Beta                                                         *
    ' * Date:     2010-12-28                                                        *
    ' * Author:   Jorge Cisneros jorgecis@gmail.com                                 *
    ' * License:  Freeware                                                          *
    ' * WebPage:  http://sourceforge.net/projects/fpdfnet/                          *
    ' *                                                                             *
    ' * You may use and modify this software as you wish.                           *
    ' * If you like this software, please make a donation                           *
    ' *******************************************************************************


    ' Private properties
    Private Structure NewArray
        Dim Key As String
        Dim Value As String
        Dim i As Short
        Dim Type As String
        Dim Name As String
        Dim up As Short
        Dim ut As Short
        Dim cw() As String
        Dim n As Short
        Dim diff As String
        Dim file As String
        Dim nf As Short
        Dim desc As String
        Dim originalsize As Integer
        Dim enc As String
        Dim length1 As Integer
        Dim length2 As Integer
    End Structure

    Private Structure ArrayImages
        Dim w As Double
        Dim h As Double
        Dim cs As String
        Dim bpc As String
        Dim F As String
        Dim parms As String
        Dim pal As String
        Dim trns As String
        Dim imgdata As String
        Dim file As String
        Dim i As Short
        Dim n As Short
    End Structure

    Private page As Integer  ' current page number
    Private n As Integer ' current object number
    Private offsets() As String ' array of object offsets
    Private Buffer As String ' buffer holding in-memory PDF
    Private pages() As String ' array containing pages
    Private State As Integer ' current document state
    Private compress As Boolean = True  ' compression flag
    Private DefOrientation As String ' default orientation
    Private CurOrientation As String ' current orientation
    Private OrientationChanges() As String  ' array indicating orientation changes
    Private K As Double ' scale factor (number of points in user unit)
    Private fwPt As Double
    Private fhPt As Double ' dimensions of page format in points
    Private fw As Double
    Private fh As Double ' dimensions of page format in user unit
    Private wPt As Double
    Private hPt As Double ' current dimensions of page in points
    Private w As Double
    Private h As Double ' current dimensions of page in user unit
    Private lMargin As Double ' left margin
    Private tMargin As Double ' top margin
    Private rMargin As Double ' right margin
    Private bMargin As Double ' page break margin
    Private cMargin As Double ' cell margin
    Private X As Double
    Private Y As Double ' current position in user unit for cell positioning
    Private lasth As Double ' height of last cell printed
    Private LineWidth As Double ' line width in user unit
    Private CoreFonts(13) As NewArray  ' array of standard font names
    Private fonts() As NewArray ' array of used fonts
    Private CurrentFont As NewArray ' current font info
    Private FontFiles As Array ' array of font files
    Private diffs() As String ' array of encoding differences
    Private images() As ArrayImages  ' array of used images
    Private PageLinks() As Array ' array of links in pages
    Private links As Array ' array of internal links
    Private FontFamily As String ' current font family
    Private FontStyle As String ' current font style
    Private underline As Boolean ' underlining flag
    Private FontSizePt As Integer ' current font size in points
    Private FontSize As Integer ' current font size in user unit
    Private DrawColor As String ' commands for drawing color
    Private FillColor As String ' commands for filling color
    Private TextColor As String ' commands for text color
    Private ColorFlag As Boolean ' indicates whether fill and text colors are different
    Private ws As Double ' word spacing
    Private AutoPageBreak As Boolean ' automatic page breaking
    Private PageBreakTrigger As String ' threshold used to trigger page breaks
    Private InFooter As Boolean ' flag set when processing footer
    Private ZoomMode As String ' zoom display mode
    Private LayoutMode As String ' layout display mode
    Private title As String ' title
    Private subject As String ' subject
    Private author As String ' author
    Private keywords As String ' keywords
    Private creator As String ' creator
    Private AliasNbPages As String ' alias for total number of pages
    Private PDFVersion As String ' PDF version number
    Private angle As Integer


    'Evento Header
    Public Event header(ByRef pdf As FPDFnet)
    Public Event footer(ByRef pdf As FPDFnet)

#If PLATFORM = "x86" Then
    Public Const DLLname As String = "zlib32.dll"
    <DllImport("zlib32.dll", EntryPoint:="compress")> _
    Private Shared Function CompressByteArray(ByVal dest As Byte(), ByRef destLen As Integer, ByVal src As Byte(), ByVal srcLen As Integer) As Integer
    End Function
#Else
    Public Const DLLname As String = "zlib64.dll"
    <DllImport("zlib64.dll", EntryPoint:="compress")> _
    Private Shared Function CompressByteArray(ByVal dest As Byte(), ByRef destLen As Integer, ByVal src As Byte(), ByVal srcLen As Integer) As Integer
    End Function
#End If


    '/*******************************************************************************
    '*                                                                              *
    '*                               Public methods                                 *
    '*                                                                              *
    '*******************************************************************************/
    Public Sub New(Optional ByRef orientation As String = "P", Optional ByRef unit As String = "mm", Optional ByRef page_format As String = "A4")
        Dim margin As Double
        Dim formatp() As Double


        ' Initialization of properties

        'Dim fpdf_charwidths(0) As Object
        ReDim fonts(0)
        ReDim PageLinks(0)
        ReDim images(0)
        page = 0
        n = 2
        Buffer = ""
        ReDim OrientationChanges(0)
        OrientationChanges(0) = ""
        State = 0
        'FontFiles = New Object() {}
        ReDim diffs(0)
        diffs(0) = ""
        'links = New Object() {}
        InFooter = False
        lasth = 0
        FontFamily = ""
        FontStyle = ""
        FontSizePt = 12
        underline = False
        DrawColor = "0 G"
        FillColor = "0 g"
        TextColor = "0 g"
        ColorFlag = False
        ws = 0

        ' Standard fonts el array cambiado al siguiente
        CoreFonts(0).Key = "courier"
        CoreFonts(0).Value = "Courier"
        CoreFonts(1).Key = "courierB"
        CoreFonts(1).Value = "Courier-Bold"
        CoreFonts(2).Key = "courierI"
        CoreFonts(2).Value = "Courier-Oblique"
        CoreFonts(3).Key = "courierBI"
        CoreFonts(3).Value = "Courier-BoldOblique"
        CoreFonts(4).Key = "helvetica"
        CoreFonts(4).Value = "Helvetica"
        CoreFonts(5).Key = "helveticaB"
        CoreFonts(5).Value = "Helvetica-Bold"
        CoreFonts(6).Key = "helveticaI"
        CoreFonts(6).Value = "Helvetica-Oblique"
        CoreFonts(7).Key = "helveticaBI"
        CoreFonts(7).Value = "Helvetica-BoldOblique"
        CoreFonts(8).Key = "times"
        CoreFonts(8).Value = "Times-Roman"
        CoreFonts(9).Key = "timesB"
        CoreFonts(9).Value = "Times-Bold"
        CoreFonts(10).Key = "timesI"
        CoreFonts(10).Value = "Times-Bold"
        CoreFonts(11).Key = "timesBI"
        CoreFonts(11).Value = "Times-BoldOblique"
        CoreFonts(12).Key = "symbol"
        CoreFonts(12).Value = "Symbol"
        CoreFonts(13).Key = "zapfdingbats"
        CoreFonts(13).Value = "ZapfDingbats"

        ' Scale factor
        If unit = "pt" Then
            K = 1
        ElseIf unit = "mm" Then
            K = 72 / 25.4
        ElseIf unit = "cm" Then
            K = 72 / 2.54
        ElseIf unit = "in" Then
            K = 72
        Else
            Show_Error("Incorrect unit: " & unit)
        End If
        ' Page format
        If page_format <> "" Then
            page_format = LCase(page_format)
            If page_format = "a3" Then
                formatp = New Double() {595.28, 841.89}
            ElseIf (page_format = "a4") Then
                formatp = New Double() {595.28, 841.89}
            ElseIf (page_format = "a5") Then
                formatp = New Double() {420.94, 595.28}
            ElseIf (page_format = "letter") Then
                formatp = New Double() {612, 792}
            ElseIf (page_format = "legal") Then
                formatp = New Double() {612, 1008}
            Else
                Show_Error("Unknown page format: " & page_format)
            End If
            fwPt = formatp(0)
            fhPt = formatp(1)
        End If

        fw = fwPt / K
        fh = fhPt / K

        ' Page orientation
        orientation = LCase(orientation)
        If (orientation = "p" Or orientation = "portrait") Then
            DefOrientation = "P"
            wPt = fwPt
            hPt = fhPt
        ElseIf (orientation = "l" Or orientation = "landscape") Then
            DefOrientation = "L"
            wPt = fhPt
            hPt = fwPt
        Else
            Show_Error("Incorrect orientation: " & orientation)
        End If
        CurOrientation = DefOrientation
        w = wPt / K
        h = hPt / K
        ' Page margins (1 cm)
        margin = 28.35 / K
        SetMargins(margin, margin)
        ' Interior cell margin (1 mm)
        cMargin = margin / 10
        ' Line width (0.2 mm)
        LineWidth = 0.567 / K
        ' Automatic page break
        SetAutoPageBreak(True, 2 * margin)
        ' Full width display mode
        SetDisplayMode("fullwidth")
        ' Enable compression
        SetCompression(True)
        ' Set default PDF version number
        PDFVersion = "1.6"
    End Sub

    Sub SetMargins(ByRef Left_Margin As Double, ByRef Top As Double, Optional ByRef Right_Margin As Double = -1)
        ' Set left, top and right margins
        lMargin = Left_Margin
        tMargin = Top
        If (Right_Margin = -1) Then
            Right_Margin = Left_Margin
        End If
        rMargin = Right_Margin
    End Sub

    Public Sub SetLeftMargin(ByRef margin As Double)
        ' Set left margin
        lMargin = margin
        If (page > 0 And X < margin) Then
            X = margin
        End If
    End Sub

    Public Sub SetTopMargin(ByRef margin As Double)
        ' Set top margin
        tMargin = margin
    End Sub

    Public Sub SetRightMargin(ByRef margin As Double)
        ' Set right margin
        rMargin = margin
    End Sub

    Public Sub SetAutoPageBreak(ByRef auto_break As Boolean, Optional ByRef margin As Double = 0)
        ' Set auto page break mode and triggering margin
        AutoPageBreak = auto_break
        bMargin = margin
        PageBreakTrigger = h - margin
    End Sub

    Public Sub SetDisplayMode(ByRef zoom As String, Optional ByRef layout As String = "continuous")
        ' Set display mode in viewer
        If (zoom = "fullpage" Or zoom = "fullwidth" Or zoom = "real" Or zoom = "default") Then
            ZoomMode = zoom
        Else
            Show_Error("Incorrect zoom display mode: " & zoom)
        End If

        If (layout = "single" Or layout = "continuous" Or layout = "two" Or layout = "default") Then
            LayoutMode = layout
        Else
            Show_Error("Incorrect layout display mode: " & layout)
        End If
    End Sub

    Sub SetCompression(ByRef st_compress As Boolean)
        ' Set page compression
        If IO.File.Exists(Environment.SystemDirectory & "\" & DLLname) Or IO.File.Exists(Application.StartupPath & "\" & DLLname) Then
            compress = st_compress
        Else
            compress = False
        End If
    End Sub

    Sub SetTitle(ByRef st_title As String)
        ' Title of document
        title = st_title
    End Sub

    Sub SetSubject(ByRef st_subject As String)
        ' Subject of document
        subject = st_subject
    End Sub

    Sub Setauthor(ByRef st_author As String)
        ' Author of document
        author = st_author
    End Sub

    Sub SetKeywords(ByRef st_keywords As String)
        ' Keywords of document
        keywords = st_keywords
    End Sub

    Sub SetCreator(ByRef st_creator As String)
        ' Creator of document
        creator = st_creator
    End Sub

    Sub AliasNbPage(Optional ByRef alias_Renamed As String = "{nb}")
        ' Define an alias for total number of pages
        AliasNbPages = alias_Renamed
    End Sub

    Sub Show_Error(ByRef MSG As String)
        ' Fatal error
        MsgBox("<B>FPDF error: </B>" & MSG, MsgBoxStyle.Critical, "Error")
        End
    End Sub

    Sub OpenPDF()
        ' Begin document
        State = 1
    End Sub

    Sub ClosePDF()
        ' Terminate document
        If State = 3 Then Exit Sub
        If page = 0 Then AddPage()
        ' Page footer
        InFooter = True
        footer_event()
        InFooter = False
        ' Close page
        endpage()
        ' Close document
        enddoc()
    End Sub

    Sub AddPage(Optional ByRef orientation As String = "")
        Dim cf As Boolean
        Dim tc As String
        Dim fc As String
        Dim dc As String
        Dim lW As Double
        Dim Size As Integer
        Dim style As String
        Dim family As String

        ' Start a new page
        If (State = 0) Then
            Call OpenPDF()
        End If
        family = FontFamily
        style = FontStyle & IIf(underline = True, "U", "")
        Size = FontSizePt
        lW = LineWidth
        dc = DrawColor
        fc = FillColor
        tc = TextColor
        cf = ColorFlag
        If (page > 0) Then
            ' Page footer
            InFooter = True
            footer_event()
            InFooter = False
            ' Close page
            endpage()
        End If
        ' Start new page
        beginpage(orientation)
        ' Set line cap style to square
        out("2 J")
        ' Set line width
        LineWidth = lW
        out(sprintf("%.2f w", lW * K))
        ' Set font
        If (family <> "") Then
            SetFont(family, style, Size)
        End If
        ' Set colors
        DrawColor = dc
        If (dc <> "0 G") Then
            out(dc)
        End If
        FillColor = fc
        If (fc <> "0 g") Then
            out(fc)
        End If
        TextColor = tc
        ColorFlag = cf
        ' Page header
        header_event()
        ' Restore line width
        If (LineWidth <> lW) Then
            LineWidth = lW
            out(sprintf("%.2f w", lW * K))
        End If
        ' Restore font
        If (family <> "") Then
            SetFont(family, style, Size)
        End If
        ' Restore colors
        If (DrawColor <> dc) Then
            DrawColor = dc
            out(dc)
        End If
        If (FillColor <> fc) Then
            FillColor = fc
            out(fc)
        End If
        TextColor = tc
        ColorFlag = cf
    End Sub
    Sub header_event()
        ' To be implemented in your own inherited class
        RaiseEvent header(Me)
    End Sub

    Sub footer_event()
        ' To be implemented in your own inherited class
        RaiseEvent footer(Me)
    End Sub

    Function PageNo() As Integer
        PageNo = page
    End Function

    Sub SetDrawColor(ByRef R As Double, Optional ByRef G As Double = -1, Optional ByRef b As Double = -1)

        ' Set color for all stroking operations
        If (R = 0 And G = 0 And b = 0) Or G = -1 Then
            DrawColor = sprintf("%.3f G", R / 255)
        Else
            DrawColor = sprintf("%.3f %.3f %.3f RG", R / 255, G / 255, b / 255)
        End If
        If page > 0 Then
            out(DrawColor)
        End If
    End Sub

    Sub SetFillColor(ByRef R As Double, Optional ByRef G As Double = -1, Optional ByRef b As Double = -1)

        ' Set color for all filling operations
        If (R = 0 And G = 0 And b = 0) Or G = -1 Then
            FillColor = sprintf("%.3f g", R / 255)
        Else
            FillColor = sprintf("%.3f %.3f %.3f rg", R / 255, G / 255, b / 255)
        End If
        ColorFlag = (FillColor <> TextColor)
        If page > 0 Then
            Call out(FillColor)
        End If
    End Sub

    Sub SetTextColor(ByRef R As Double, Optional ByRef G As Double = -1, Optional ByRef b As Double = -1)

        ' Set color for text
        If ((R = 0 And G = 0 And b = 0) Or G = -1) Then
            TextColor = sprintf("%.3f g", R / 255)
        Else
            TextColor = sprintf("%.3f %.3f %.3f rg", R / 255, G / 255, b / 255)
        End If
        ColorFlag = (FillColor <> TextColor)
    End Sub

    Function GetStringWidth(ByRef s As String) As Double
        Dim a As Integer
        Dim letra As String
        Dim i As Integer
        Dim l As Integer
        Dim w1 As Double
        Dim cw As Array

        ' Get width of a string in the current font

        If IsNumeric(s) Then s = Trim(s)
        cw = CurrentFont.cw
        w1 = 0
        l = Len(s)
        For i = 1 To l
            letra = Mid(s, i, 1)
            For a = 0 To UBound(cw) Step 2
                If cw(a) = letra Then
                    w1 = w1 + Val(cw(a + 1))
                    Exit For
                End If
            Next
        Next
        GetStringWidth = w1 * FontSize / 1000
    End Function

    Sub SetLineWidth(ByRef Width As Double)

        ' Set line width
        LineWidth = Width
        If page > 0 Then
            out(sprintf("%.2f w", Width * K))
        End If
    End Sub

    Sub Lines(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double)

        ' Draw a line
        out(sprintf("%.2f %.2f m %.2f %.2f l S", x1 * K, (h - y1) * K, x2 * K, (h - y2) * K))
    End Sub

    Sub RECT(ByRef x1 As Double, ByRef y1 As Double, ByRef w1 As Double, ByRef h1 As Double, Optional ByRef style As String = "")
        Dim op As String

        ' Draw a rectangle
        If style = "F" Then
            op = "f"
        ElseIf style = "FD" Or style = "DF" Then
            op = "B"
        Else
            op = "S"
        End If
        out(sprintf("%.2f %.2f %.2f %.2f re %s", x1 * K, (h - y1) * K, w1 * K, -h1 * K, op))
    End Sub

    Sub AddFont(ByRef family As String, Optional ByRef style As String = "", Optional ByRef file As String = "")
        Dim diff As Integer
        Dim nb As Integer
        Dim d As Integer
        Dim i As Integer
        Dim fontkey As String

        Dim info As NewArray

        family = LCase(family)
        If file = "" Then
            file = Replace(family, " ", "") & LCase(style) & ".vbf"
        End If
        If family = "arial" Then
            family = "helvetica"
        End If
        style = UCase(style)
        If style = "IB" Then
            style = "BI"
        End If
        fontkey = family & style
        If get_array_value(fonts, fontkey) = Not False Then
            Show_Error("Font already added: " & family & " " & style)
        End If
        info = Read_font(getfontpath() & file)
        info.Key = fontkey
        If (info.Name = "") Then
            Show_Error("Could not include font definition file")
        End If

        i = UBound(fonts) + 1
        ReDim Preserve fonts(UBound(fonts) + 1)
        fonts(i) = info
        fonts(i).i = i
        If info.diff <> "" Then
            ' Search existing encodings
            d = 0
            nb = UBound(diffs)
            For i = 1 To nb
                If (diffs(i) = diff) Then
                    d = i
                    Exit For
                End If
            Next
            If (d = 0) Then
                d = nb + 1
                'diffs [d] = diff
            End If
            fonts(UBound(fonts)).diff = d
        End If
        If (info.file <> "") Then
            If (info.Type = "TrueType") Then
                fonts(UBound(fonts)).length1 = fonts(UBound(fonts)).originalsize
                'FontFiles [file] = Array("length1" >= OriginalSize)
            Else
                'FontFiles [file] = Array("length1" >= size1, "length2" >= size2)
            End If
        End If
    End Sub

    Sub SetFont(ByRef family As String, Optional ByRef style As String = "", Optional ByRef Size As Integer = 0)
        Dim i As Integer
        Dim file As String
        Dim a As Integer
        Dim fontkey As String

        ' Select a font size given in points
        family = LCase(family)
        If (family = "") Then
            family = FontFamily
        End If
        If (family = "arial") Then
            family = "helvetica"
        ElseIf (family = "symbol" Or family = "zapfdingbats") Then
            style = ""
        End If
        style = UCase(style)
        If (InStr(style, "U") <> False) Then
            underline = True
            style = Replace("U", "", style)
        Else
            underline = False
        End If
        If (style = "IB") Then
            style = "BI"
        End If
        If (Size = 0) Then
            Size = FontSizePt
        End If
        ' Test if font is already selected
        If (FontFamily = family And FontStyle = style And FontSizePt = Size) Then
            Exit Sub
        End If
        ' Test if used for the first time
        fontkey = family & style
        Dim info As NewArray = Nothing
        If (find_array(fonts, fontkey) = -1) Then
            ' Check if one of the standard fonts
            a = find_array(CoreFonts, fontkey)
            If (a >= 0) Then

                If (CoreFonts(a).cw Is Nothing) Then
                    ' Load metric file
                    file = family
                    If (family = "times" Or family = "helvetica") Then
                        file = file & LCase(style)
                    End If

                    info = Read_font(getfontpath() & file & ".vbf")
                    If info.cw(0) = "" Then
                        Show_Error("Could not include font metric file")
                    End If
                End If
                i = UBound(fonts) + 1
                ReDim Preserve fonts(UBound(fonts) + 1)
                fonts(UBound(fonts)).Key = fontkey
                fonts(UBound(fonts)).i = i
                fonts(UBound(fonts)).Type = "core"
                fonts(UBound(fonts)).Name = get_array_value(CoreFonts, fontkey)
                fonts(UBound(fonts)).Value = get_array_value(CoreFonts, fontkey)
                fonts(UBound(fonts)).up = -100
                fonts(UBound(fonts)).ut = 50
                fonts(UBound(fonts)).cw = info.cw
            Else
                Show_Error("Undefined font: " & family & " " & style)
            End If
        End If
        ' Select it
        FontFamily = family
        FontStyle = style
        FontSizePt = Size
        FontSize = Size / K
        CurrentFont = fonts(get_array_index(fonts, fontkey))
        If (page > 0) Then
            Call out(sprintf("BT /F%d %.2f Tf ET", CurrentFont.i, FontSizePt))
        End If
    End Sub

    Sub SetFontSize(ByRef Size As Integer)

        ' Set font size in points
        If (FontSizePt = Size) Then
            Exit Sub
        End If
        FontSizePt = Size
        FontSize = Size / K
        If page > 0 Then
            Call out(sprintf("BT /F%d %.2f Tf ET", CurrentFont.i, FontSizePt))
        End If
    End Sub

    Function AddLink() As Object

        ' Create a new internal link
        n = UBound(links) + 1
        'links [n] = Array(0, 0)
        AddLink = n
    End Function

    Sub SetLink(ByRef link As Object, Optional ByRef Y As Double = 0, Optional ByRef page As Object = -1)

        ' Set destination of internal link
        If (Y = -1) Then Y = Y
        If (page = -1) Then page = page
        'links [link] = Array(page, y)
    End Sub

    'Function link(x1, y1, w1, h1, link)
    ' Put a link on the page
    'PageLinks[page][]=array(x*k,hPt-y*k,w*k,h*k,link)
    'End Function

    Sub text(ByRef X As Double, ByRef Y As Double, ByRef txt As String)
        Dim s As String

        ' Output a string
        s = sprintf("BT %.2f %.2f Td (%s) Tj ET", X * K, (h - Y) * K, escape(txt))
        If underline And txt <> "" Then
            s = s & " " & dounderline(X, Y, txt)
        End If
        If ColorFlag Then
            s = "q " & TextColor & " " & s & " Q"
        End If
        Call out(s)
    End Sub

    Function AcceptPageBreak() As Boolean

        ' Accept automatic page break or not
        Debug.Print(AutoPageBreak)
        'UPGRADE_WARNING: Couldn't resolve default property of object AutoPageBreak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AcceptPageBreak = AutoPageBreak
    End Function

    Sub Cell(ByRef w1 As Double, Optional ByRef h1 As Double = 0, Optional ByRef txt As String = "", Optional ByRef border As String = "", Optional ByRef Ln As Double = 0, Optional ByRef align As String = "", Optional ByRef fill As Double = 0, Optional ByRef link As String = "")
        Dim dx As Double
        Dim y1 As Double
        Dim op As String
        Dim s As String
        Dim ws1 As Double
        Dim x1 As Double
        Dim k1 As Double

        ' Output a cell
        k1 = K
        If (Y + h1 > PageBreakTrigger And Not InFooter And AcceptPageBreak()) Then
            ' Automatic page break
            x1 = X
            ws1 = ws
            If (ws1 > 0) Then
                ws = 0
                out("0 Tw")
            End If
            AddPage(CurOrientation)
            X = x1
            If (ws1 > 0) Then
                ws = ws1
                out(sprintf("%.3f Tw", ws1 * k1))
            End If
        End If
        If (w1 = 0) Then
            w1 = w - rMargin - X
        End If
        s = ""
        If (fill = 1 Or border = "1") Then
            If (fill = 1) Then
                op = IIf(border = 1, "B", "f")
            Else
                op = "S"
            End If
            s = sprintf("%.2f %.2f %.2f %.2f re %s ", X * k1, (h - Y) * k1, w1 * k1, -h1 * k1, op)
        End If
        If (border <> "0") Then
            x1 = X
            y1 = Y
            If (InStr(border, "L") > 0) Then
                s = s & sprintf("%.2f %.2f m %.2f %.2f l S ", x1 * k1, (h - y1) * k1, x1 * k1, (h - (y1 + h1)) * k1)
            End If
            If (InStr(border, "T") > 0) Then
                s = s & sprintf("%.2f %.2f m %.2f %.2f l S ", x1 * k1, (h - y1) * k1, (x1 + w1) * K, (h - y1) * k1)
            End If
            If (InStr(border, "R") > 0) Then
                s = s & sprintf("%.2f %.2f m %.2f %.2f l S ", (x1 + w1) * k1, (h - y1) * k1, (x1 + w1) * K, (h - (y1 + h1)) * K)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object border. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (InStr(border, "B") > 0) Then
                s = s & sprintf("%.2f %.2f m %.2f %.2f l S ", x1 * k1, (h - (y1 + h1)) * k1, (x1 + w1) * k1, (h - (y1 + h1)) * k1)
            End If
        End If
        If (txt <> "") Then
            If (align = "R") Then
                dx = w1 - cMargin - GetStringWidth(txt)
            ElseIf (align = "C") Then
                dx = (w1 - GetStringWidth(txt)) / 2
            Else
                dx = cMargin
            End If
            If (ColorFlag) Then
                s = s & "q " & TextColor & " "
            End If
            txt = Replace(Replace(Replace(txt, "\", "\\"), "(", "\("), ")", "\)")
            s = s & sprintf("BT %.2f %.2f Td (%s) Tj ET", (X + dx) * K, (h - (Y + 0.5 * h1 + 0.3 * FontSize)) * K, txt)
            If (underline) Then
                s = s & " " & dounderline(X + dx, Y + 0.5 * h + 0.3 * FontSize, txt)
            End If
            If (ColorFlag) Then
                s = s & " Q"
            End If
            If (link <> "") Then
                'Call link(x + dx, y + 0.5 * h - 0.5 * FontSize, GetStringWidth(txt), FontSize, link)
            End If
        End If
        If (s <> "") Then
            out(s)
        End If
        lasth = h1
        If (Ln > 0) Then
            ' Go to next line
            Y = Y + h1
            If (Ln = 1) Then
                X = lMargin
            End If
        Else
            X = X + w1
        End If
    End Sub

    Sub MultiCell(ByRef w1 As Double, ByRef h1 As Double, ByRef txt As String, Optional ByRef border As String = "", Optional ByRef align As String = "J", Optional ByRef fill As Integer = 0)
        Dim ls As Double
        Dim c As String
        Dim nl As Integer
        Dim ns As Integer
        Dim l As Integer
        Dim j As Integer
        Dim i As Integer
        Dim sep As Integer
        Dim b2 As String = ""
        Dim b As String = ""
        Dim nb As Integer
        Dim s As String
        Dim wmax As Double
        Dim cw As Array

        ' Output text with automatic or explicit line breaks
        cw = CurrentFont.cw
        If (w1 = 0) Then
            w1 = w - rMargin - X
        End If
        wmax = (w1 - 2 * cMargin) * 1000 / FontSize
        If txt = "" Then txt = " "
        s = Replace(txt, vbCrLf, vbLf)
        s = txt

        nb = Len(s)
        If (nb > 0 And Right(s, 1) = vbCrLf) Then
            nb = nb - 1
        End If
        If (border <> "") Then
            If (border = "1") Then
                border = "LTRB"
                b = "LRT"
                b2 = "LR"
            Else
                b2 = ""
                If (InStr(border, "L") > 0) Then
                    b2 = b2 & "L"
                End If
                If (InStr(border, "R") > 0) Then
                    b2 = b2 & "R"
                End If
                b = IIf(InStr(border, "T") > 0, b2 & "T", b2)
            End If
        End If
        sep = -1
        i = 1
        j = 1
        l = 0
        ns = 0
        nl = 1
        Do While i < nb
            ' Get next character
            c = Mid(s, i, 1)
            If (c = vbLf) Then
                ' Explicit line break
                If (ws > 0) Then
                    ws = 0
                    out("0 Tw")
                End If
                Cell(w1, h1, Mid(s, j, i - j), b, 2, align, fill)
                i = i + 1
                sep = -1
                j = i
                l = 0
                ns = 0
                nl = nl + 1
                If (border <> "" And nl = 2) Then
                    b = b2
                End If
                'continue
            End If
            If (c = " ") Then
                sep = i
                ls = l
                ns = ns + 1
            End If
            l = l + cw((Asc(c) * 2) + 1)
            If (l > wmax) Then
                ' Automatic line break
                If (sep = -1) Then
                    If (i = j) Then i = i + 1
                    If (ws > 0) Then
                        ws = 0
                        Call out("0 Tw")
                    End If
                    Call Cell(w1, h1, Mid(s, j, i - j), b, 2, align, fill)
                Else
                    If (align = "J") Then
                        ws = IIf(ns > 1, (wmax - ls) / 1000 * FontSize / (ns - 1), 0)
                        Call out(sprintf("%.3f Tw", ws * K))
                    End If
                    Call Cell(w1, h1, Mid(s, j, sep - j), b, 2, align, fill)
                    i = sep + 1
                End If
                sep = -1
                j = i
                l = 0
                ns = 0
                nl = nl + 1
                If (border <> "" And nl = 2) Then
                    b = b2
                End If
            Else
                i = i + 1
            End If
        Loop
        ' Last chunk
        If (ws > 0) Then
            ws = 0
            Call out("0 Tw")
        End If
        If (border <> "" And InStr(border, "B") > 0) Then
            b = b & "B"
        End If
        Call Cell(w1, h1, Mid(s, j, i - j), b, 2, align, fill)
        X = lMargin
    End Sub

    Sub WritePDF(ByRef h As Double, ByRef txt As Object, Optional ByRef link As Object = "")
        Dim c As String = ""
        Dim nl As Integer
        Dim l As Integer
        Dim j As Integer
        Dim i As Integer
        Dim sep As Integer
        Dim nb As Integer
        Dim s As String
        Dim wmax As Double
        Dim cw As Array

        ' Output text in flowing mode
        cw = CurrentFont.cw
        w = w - rMargin - X
        wmax = (w - 2 * cMargin) * 1000 / FontSize
        s = Replace("\r", "", txt)
        nb = Len(s)
        sep = -1
        i = 0
        j = 0
        l = 0
        nl = 1
        Do While (i < nb)
            ' Get next character
            'c=s{i}
            If (c = "\n") Then
                ' Explicit line break
                Call Cell(w, h, Mid(s, j, i - j), 0, 2, "", 0, link)
                i = i + 1
                sep = -1
                j = i
                l = 0
                If (nl = 1) Then
                    X = lMargin
                    w = w - rMargin - X
                    wmax = (w - 2 * cMargin) * 1000 / FontSize
                End If
                nl = nl + 1
                'continue
            End If
            If (c = " ") Then
                sep = i
            End If
            l = l + cw(c)
            If (l > wmax) Then
                ' Automatic line break
                If (sep = -1) Then
                    If (X > lMargin) Then
                        ' Move to next line
                        X = lMargin
                        Y = Y + h
                        w = w - rMargin - X
                        wmax = (w - 2 * cMargin) * 1000 / FontSize
                        i = i + 1
                        nl = nl + 1
                        'continue
                    End If
                    If (i = j) Then i = i + 1
                    Call Cell(w, h, Mid(s, j, i - j), 0, 2, "", 0, link)
                Else
                    Call Cell(w, h, Mid(s, j, sep - j), 0, 2, "", 0, link)
                    i = sep + 1
                End If
                sep = -1
                j = i
                l = 0
                If (nl = 1) Then
                    X = lMargin
                    w = w - rMargin - X
                    wmax = (w - 2 * cMargin) * 1000 / FontSize
                End If
                nl = nl + 1
            Else
                i = i + 1
            End If
        Loop
        ' Last chunk
        If (i <> j) Then
            Call Cell(l / 1000 * FontSize, h, Mid(s, j), 0, 0, "", 0, link)
        End If
    End Sub

    Sub image(ByRef file As String, ByRef xi As Double, ByRef yi As Double, Optional ByRef wi As Double = 0, Optional ByRef hi As Double = 0, Optional ByRef typeimg As String = "", Optional ByRef link As String = "")
        Dim pos As Integer

        Dim info As ArrayImages = Nothing

        ' Put an image on the page
        If get_img_array(images, file) = False Then
            ' First use of image, get info
            If typeimg = "" Then
                pos = InStr(file, ".")
                If (pos = 0) Then
                    Show_Error("Image file has no extension and no type was specified: " & file)
                End If
                typeimg = Right(file, 3)
            End If
            typeimg = LCase(typeimg)
            If (typeimg = "jpg" Or typeimg = "jpeg") Then
                info = parsejpg(file)
            ElseIf (typeimg = "png") Then
                info = parsepng(file)
            Else
                ' Allow for additional formats
                Show_Error("Unsupported image type: " & typeimg & " " & file)
            End If
            info.i = UBound(images) + 1
            ReDim Preserve images(info.i)
            images(info.i) = info
        Else
            info = images(get_img_array(images, file))
        End If
        ' Automatic width and height calculation if needed
        If wi = 0 And hi = 0 Then
            ' Put image at 72 dpi
            wi = info.w / K
            hi = info.h / K
        End If
        If wi = 0 Then
            wi = hi * info.w / info.h
        End If
        If hi = 0 Then
            hi = wi * info.h / info.w
        End If
        Call out(sprintf("q %.2f 0 0 %.2f %.2f %.2f cm /I%d Do Q", wi * K, hi * K, xi * K, (h - (yi + hi)) * K, info.i))
        If (link <> "") Then
            'Call link(x1, y1, w1, h1, link)
        End If
    End Sub

    Sub Ln(Optional ByRef h1 As String = "")

        ' Line feed default value is last cell height
        X = lMargin
        If h1 = "" Then
            Y = Y + lasth
        Else
            Y = Y + Val(h1)
        End If
    End Sub

    Function GetX() As Double
        ' Get x position
        GetX = X
    End Function

    Sub SetX(ByRef x1 As Double)

        ' Set x position
        If x1 >= 0 Then
            X = x1
        Else
            X = w + x1
        End If
    End Sub

    Function GetY() As Double
        ' Get y position
        GetY = Y
    End Function

    Sub SetY(ByRef y1 As Double)

        ' Set y position and reset x
        'x = lMargin
        If y1 >= 0 Then
            Y = y1
        Else
            Y = h + y1
        End If
    End Sub

    Sub SetXY(ByRef x1 As Double, ByRef y1 As Double)

        ' Set x and y positions
        SetY(y1)
        SetX(x1)
    End Sub

    Function Output(ByRef Name As String, Optional ByRef Dest As String = "F") As String

        ' Output PDF to some destination
        ' Finish document if necessary
        If State < 3 Then
            Call ClosePDF()
        End If

        If Dest = "F" Then
            ' Save to local file
            FileOpen(1, Name, OpenMode.Output)
            PrintLine(1, Buffer)
            FileClose()
            Return "Ok"
        ElseIf Dest = "S" Then
            Output = Buffer
        Else
            Show_Error("Incorrect output destination: " & Dest)
            Return "Error"
        End If
    End Function

    '/*******************************************************************************
    '*                                                                              *
    '*                              Protected methods                               *
    '*                                                                              *
    '*******************************************************************************/


    Private Function getfontpath() As String
        Dim FPDF_FONTPATH As String = ""
        If FPDF_FONTPATH = "" Then
            getfontpath = My.Application.Info.DirectoryPath & "\font\"
        Else
            getfontpath = FPDF_FONTPATH
        End If
    End Function

    Private Sub putpages()
        Dim i As Integer
        Dim kids As String
        Dim p As Object
        Dim annots As String = ""
        Dim a As Integer
        Dim nb As Integer

        nb = page
        If Not IsNothing(AliasNbPages) Then
            ' Replace number of pages
            For a = 1 To nb
                pages(a) = pages(a).ToString.Replace(AliasNbPages, nb)
            Next
        End If
        If DefOrientation = "P" Then
            wPt = fwPt
            hPt = fhPt
        Else
            wPt = fhPt
            hPt = fwPt
        End If

        For a = 1 To nb
            ' Page
            newobj()
            out("<</Type /Page")
            out("/Parent 1 0 R")
            If UBound(OrientationChanges) >= a Then
                If OrientationChanges(a) <> "" Then
                    out(sprintf("/MediaBox [0 0 %.2f %.2f]", hPt, wPt))
                End If
            End If
            out("/Resources 2 0 R")
            If UBound(PageLinks) >= a Then
                'If PageLinks(a). <> "" Then
                ' Links
                'annots = "/Annots ["
                'foreach(PageLinks[a] as pl)
                '{
                '    rect=sprintf("%.2f %.2f %.2f %.2f",pl[0],pl[1],pl[0]+pl[2],pl[1]-pl[3])
                '    annots.="<</Type /Annot /Subtype /Link /Rect [".rect."] /Border [0 0 0] "
                '    if(is_string(pl[4]))
                '        annots.="/A <</S /URI /URI "._textstring(pl[4]).">>>>"
                '    Else
                '    {
                '        l=links[pl[4]]
                '        h=isset(OrientationChanges[l[0]]) ? wPt : hPt
                '        annots.=sprintf("/Dest [%d 0 R /XYZ 0 %.2f null]>>",1+2*l[0],h-l[1]*k)
                '    }
                '}
                out(annots & "]")
                'End If
            End If
            out("/Contents " & (n + 1) & " 0 R>>")
            out("endobj")
            ' Page content
            p = IIf(compress, gzcompress(pages(a)), pages(a))
            Dim Filter As String = IIf(compress, "/Filter /FlateDecode ", "")
            'p = pages(a)
            newobj()
            out("<<" & Filter & "/Length " & Len(p) & ">>")
            putstream(p)
            out("endobj")
        Next
        ' Pages root
        offsets(1) = Len(Buffer)
        out("1 0 obj")
        out("<</Type /Pages")
        kids = "/Kids ["
        For i = 0 To nb - 1
            kids = kids & (3 + 2 * i) & " 0 R "
        Next
        out(kids & "]")
        out("/Count " & nb)
        out(sprintf("/MediaBox [0 0 %.2f %.2f]", wPt, hPt))
        out(">>")
        out("endobj")
    End Sub

    Private Sub putfonts()
        Dim mtd As String
        Dim file As String
        Dim desc As Array
        Dim ii As Integer
        Dim i As Integer
        Dim s As String
        Dim cw As Array
        Dim Name As String
        Dim typefont As String
        Dim compressed As Boolean
        Dim diff As String = ""
        Dim a As Integer
        Dim nf As Integer

        nf = n
        For a = 1 To UBound(diffs)
            ' Encodings
            newobj()
            out("<</Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [" & diff & "]>>")
            out("endobj")
        Next

        Dim fontdata As String
        For a = 1 To UBound(fonts)
            If fonts(a).Type <> "core" Then
                ' Font file embedding
                newobj()
                fonts(a).nf = n

                FileOpen(1, getfontpath() & fonts(a).file, OpenMode.Binary)
                'if(!f)
                '    Show_Error ("Font file not found")
                fontdata = Space(LOF(1))
                FileGet(1, fontdata, 1)
                FileClose(1)
                compressed = (Right(fonts(a).file, 2) = ".z")
                Debug.Print(compressed)
                If (Not compressed And fonts(a).length2 > 0) Then
                    '    header=(ord(font{0})=128)
                    '    if(header)
                    '    {
                    '        ' Strip first binary header
                    '        Font = substr(Font, 6)
                    '    }
                    '    if(header and ord(font{info["length1"]})=128)
                    '    {
                    '        ' Strip second binary header
                    '        font=substr(font,0,info["length1"]).substr(font,info["length1"]+6)
                    '    }
                End If
                out("<</Length " & Len(fontdata))
                If compressed Then
                    Call out("/Filter /FlateDecode")
                End If
                out("/Length1 " & fonts(a).length1)
                If (fonts(a).length2 > 0) Then
                    Call out("/Length2 " & fonts(a).length2 & " /Length3 0")
                End If
                out(">>")
                putstream(fontdata)
                out("endobj")
            End If
        Next

        For a = 1 To UBound(fonts)
            ' Font objects
            fonts(a).n = n + 1
            typefont = fonts(a).Type
            Name = fonts(a).Name
            If typefont = "core" Then
                ' Standard font
                newobj()
                out("<</Type /Font")
                out("/BaseFont /" & Name)
                out("/Subtype /Type1")
                If Name <> "Symbol" And Name <> "ZapfDingbats" Then
                    out("/Encoding /WinAnsiEncoding")
                End If
                Call out(">>")
                Call out("endobj")
            ElseIf typefont = "Type1" Or typefont = "TrueType" Then
                ' Additional Type1 or TrueType font
                newobj()
                Call out("<</Type /Font")
                Call out("/BaseFont /" & Name)
                Call out("/Subtype /" & typefont)
                Call out("/FirstChar 32 /LastChar 255")
                Call out("/Widths " & (n + 1) & " 0 R")
                Call out("/FontDescriptor " & (n + 2) & " 0 R")
                'If fonts(a).name = "ENC" Then
                If fonts(a).enc <> "" Then
                    If (fonts(a).diff <> "") Then
                        Call out("/Encoding " & (nf + fonts(a).diff) & " 0 R")
                    Else
                        Call out("/Encoding /WinAnsiEncoding")
                    End If
                End If
                Call out(">>")
                Call out("endobj")
                ' Widths
                newobj()
                cw = fonts(a).cw

                s = "["
                For i = 32 To 255
                    For ii = i * 2 To UBound(cw) Step 2
                        If cw(ii) = Chr(i) Then
                            s = s & cw(ii + 1) & " "
                            Exit For
                        Else
                            Debug.Print("Error@")
                        End If
                    Next
                Next
                Call out(s & "]")
                Call out("endobj")
                ' Descriptor
                newobj()
                s = "<</Type /FontDescriptor /FontName /" & Name
                If fonts(a).desc <> "" Then
                    desc = Split(fonts(a).desc, ",")
                    For i = 0 To UBound(desc) Step 2
                        s = s & " /" & desc(i) & " " & desc(i + 1)
                    Next
                End If
                file = fonts(a).file
                If (file <> "") Then
                    s = s & " /FontFile" & IIf(typefont = "Type1", "", "2") & " " & fonts(a).nf & " 0 R"
                End If
                Call out(s & ">>")
                Call out("endobj")
            Else
                ' Allow for additional types
                mtd = "_put" & LCase(typefont)
                'if(!method_exists(this,mtd))
                '    Show_Error("Unsupported font type: ".type)
                'mtd (Font)
            End If
        Next
    End Sub

    Private Sub putimages()
        Dim a As Integer

        'filter=(compress) ? "/Filter /FlateDecode " : ""
        'Filter = ""
        'Reset (images)
        For a = 1 To UBound(images)
            newobj()
            With images(a)
                .n = n
                Call out("<</Type /XObject")
                Call out("/Subtype /Image")
                Call out("/Width " & .w)
                Call out("/Height " & .h)
                If (.cs = "Indexed") Then
                    Call out("/ColorSpace [/Indexed /DeviceRGB " & (Len(.pal) / 3 - 1) & " " & (n + 1) & " 0 R]")
                Else
                    Call out("/ColorSpace /" & .cs)
                    If (.cs = "DeviceCMYK") Then
                        Call out("/Decode [1 0 1 0 1 0 1 0]")
                    End If
                End If
                Call out("/BitsPerComponent " & .bpc)
                If (.F <> "") Then
                    Call out("/Filter /" & .F)
                End If
                If (.parms <> "") Then
                    Call out(.parms)
                End If
                'if(.trns <> "") and is_array(info["trns"]))
                '{
                '    trns = ""
                '    for(i=0i<count(info["trns"])i++)
                '        trns.=info["trns"][i]." ".info["trns"][i]." "
                '    call out("/Mask [".trns."]")
                '}
                Call out("/Length " & Len(.imgdata) & ">>")
                putstream(.imgdata)
                'unset(images[file]["data"])
                Call out("endobj")
                ' Palette
                If (.cs = "Indexed") Then
                    newobj()
                    'pal=(compress) ? gzcompress(info["pal"]) : info["pal"]
                    Call out("<< " & "/Length " & Len(.pal) & ">>")
                    putstream(.pal)
                    Call out("endobj")
                End If
            End With
        Next
    End Sub

    Private Sub putxobjectdict()
        For a As Integer = 1 To UBound(images)
            Call out("/I" & images(a).i & " " & images(a).n & " 0 R")
        Next
    End Sub

    Private Sub putresourcedict()

        Call out("/ProcSet [/PDF /Text /ImageB /ImageC /ImageI]")
        Call out("/Font <<")
        For a As Integer = 1 To UBound(fonts)
            out("/F" & fonts(a).i & " " & fonts(a).n & " 0 R")
        Next
        Call out(">>")
        Call out("/XObject <<")
        putxobjectdict()
        Call out(">>")
    End Sub

    Private Sub putresources()

        putfonts()
        putimages()
        ' Resource dictionary
        offsets(2) = Len(Buffer)
        out("2 0 obj")
        out("<<")
        putresourcedict()
        out(">>")
        out("endobj")
    End Sub

    Private Sub putinfo()
        out("/Producer " & textstring("FPDF " & PDFVersion))
        If title <> "" Then Call out("/Title " & textstring(title))
        If subject <> "" Then Call out("/Subject " & textstring(subject))
        If author <> "" Then Call out("/Author " & textstring(author))
        If keywords <> "" Then Call out("/Keywords " & textstring(keywords))
        If creator <> "" Then Call out("/Creator " & textstring(creator))
        Call out("/CreationDate " & textstring("D:" & Today))
    End Sub

    Private Sub putcatalog()
        Call out("/Type /Catalog")
        Call out("/Pages 1 0 R")
        If (ZoomMode = "fullpage") Then
            Call out("/OpenAction [3 0 R /Fit]")
        ElseIf (ZoomMode = "fullwidth") Then
            Call out("/OpenAction [3 0 R /FitH null]")
        ElseIf (ZoomMode = "real") Then
            Call out("/OpenAction [3 0 R /XYZ null null 1]")
        ElseIf (ZoomMode <> "") Then
            Call out("/OpenAction [3 0 R /XYZ null null " & (ZoomMode / 100) & "]")
        End If
        If (LayoutMode = "single") Then
            Call out("/PageLayout /SinglePage")
        ElseIf (LayoutMode = "continuous") Then
            Call out("/PageLayout /OneColumn")
        ElseIf (LayoutMode = "two") Then
            Call out("/PageLayout /TwoColumnLeft")
        End If
    End Sub

    Private Sub putheader()
        Call out("%PDF-" & PDFVersion)
    End Sub

    Private Sub puttrailer()
        Call out("/Size " & (n + 1))
        Call out("/Root " & n & " 0 R")
        Call out("/Info " & (n - 1) & " 0 R")
    End Sub

    Private Sub enddoc()
        Dim i As Integer
        Dim o As Integer

        putheader()
        putpages()
        putresources()
        ' Info
        newobj()
        out("<<")
        putinfo()
        out(">>")
        out("endobj")
        ' Catalog
        newobj()
        out("<<")
        putcatalog()
        out(">>")
        out("endobj")
        ' Cross-ref
        o = Len(Buffer)
        out("xref")
        out("0 " & (n + 1))
        out("0000000000 65535 f ")
        For i = 1 To n
            out(sprintf("%010d 00000 n ", offsets(i)))
        Next
        ' Trailer
        out("trailer")
        out("<<")
        puttrailer()
        out(">>")
        out("startxref")
        out(o)
        out("%%EOF")
        State = 3
    End Sub

    Private Sub beginpage(ByRef orientation As Object)

        page = page + 1
        ReDim Preserve pages(page)
        pages(page) = ""
        State = 2
        X = lMargin
        Y = tMargin
        FontFamily = ""
        ' Page orientation
        If (orientation = "") Then
            orientation = DefOrientation
        Else
            orientation = UCase(orientation)
            If (orientation <> DefOrientation) Then
                OrientationChanges(page) = True
            End If
        End If
        If (orientation <> CurOrientation) Then
            ' Change orientation
            If (orientation = "P") Then
                wPt = fwPt
                hPt = fhPt
                w = fw
                h = fh
            Else
                wPt = fhPt
                hPt = fwPt
                w = fh
                h = fw
            End If
            PageBreakTrigger = h - bMargin
            CurOrientation = orientation
        End If
    End Sub

    Private Sub endpage()

        If (angle <> 0) Then
            angle = 0
            Call out("Q")
        End If

        ' End of page contents
        State = 1
    End Sub

    Private Sub newobj()
        ' Begin a new object
        n = n + 1
        ReDim Preserve offsets(n)
        offsets(n) = Len(Buffer)
        out(n & " 0 obj")
    End Sub

    Private Function dounderline(ByRef X As Double, ByRef Y As Double, ByRef txt As Double) As String
        Dim ut As Short
        Dim up As Short

        ' Underline text
        up = CurrentFont.up
        ut = CurrentFont.ut
        w = GetStringWidth(txt) + ws * substr_count(txt, " ")
        dounderline = sprintf("%.2f %.2f %.2f %.2f re f", X * K, (h - (Y - up / 1000 * FontSize)) * K, w * K, -ut / 1000 * FontSizePt)
    End Function

    Private Function parsejpg(ByRef file As Object) As ArrayImages
        Dim bpc As Short
        Dim colspace As String
        parsejpg = Nothing
        Const BUFFERSIZE As Integer = 65535

        Dim bBuf(0) As Byte

        ' Extract info from a JPG file

        Dim fs As IO.Stream
        Try
            fs = New IO.FileStream(file, IO.FileMode.Open, IO.FileAccess.Read)
            ReDim bBuf(fs.Length)
            fs.Read(bBuf, 0, fs.Length)
            fs.Close()
        Catch exFile As Exception
            MsgBox(("Cannot open " & exFile.ToString()))
            Exit Function
        End Try

        Dim lPos As Integer

        Do
            ' loop through looking for the byte sequence FF,D8,FF
            ' which marks the begining of a JPEG file
            ' lPos will be left at the postion of the start
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 And bBuf(lPos + 2) = &HFF) Or (lPos >= BUFFERSIZE - 10) Then Exit Do

            ' move our pointer up
            lPos = lPos + 1

            ' and continue
        Loop

        lPos = lPos + 2
        If lPos >= BUFFERSIZE - 10 Then
            Return Nothing
            Exit Function
        End If




        Do
            ' loop through the markers until we find the one
            'starting with FF,C0 which is the block containing the
            'image information

            Do
                ' loop until we find the beginning of the next marker
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= BUFFERSIZE - 10 Then
                    Return Nothing
                    Exit Function
                End If

            Loop

            ' move pointer up
            lPos = lPos + 1

            Select Case bBuf(lPos)
                Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, &HCD To &HCF
                    ' we found the right block
                    Exit Do
            End Select

            ' otherwise keep looking
            lPos = lPos + Mult(bBuf(lPos + 2), bBuf(lPos + 1))

            ' check for end of buffer
            If lPos >= BUFFERSIZE - 10 Then
                Return Nothing
                Exit Function
            End If


        Loop





        If bBuf(lPos + 8) = 1 Or bBuf(lPos + 8) = 3 Then
            colspace = "DeviceRGB"
        ElseIf bBuf(lPos + 8) = 4 Then
            colspace = "DeviceCMYK"
        Else
            colspace = "DeviceGray"
        End If

        'bpc = IIf(bBuf(lPos + 8) > 0, bBuf(lPos + 8), 8)
        bpc = 8



        With parsejpg
            .w = Mult(bBuf(lPos + 7), bBuf(lPos + 6))
            .h = Mult(bBuf(lPos + 5), bBuf(lPos + 4))
            .cs = colspace
            .bpc = bpc
            .F = "DCTDecode"
            .file = file
        End With

        Dim data As String
        parsejpg.imgdata = ""

        FileOpen(1, file, OpenMode.Binary)
        data = Space(LOF(1))
        Do While Not EOF(1)
            FileGet(1, data)
            parsejpg.imgdata = parsejpg.imgdata & data
        Loop
        FileClose(1)

    End Function

    Private Function parsepng(ByRef file As String) As ArrayImages
        Dim typ As Object
        Dim ni As Object
        Dim parms As Object
        Dim colspace As String = ""
        Dim ct As Object
        Dim bpc As Short
        Dim datos As String


        ' Extract info from a PNG file
        FileOpen(1, file, OpenMode.Binary)
        '    if(!f)
        '        Show_Error("Can\"t open image file: ".file)
        ' Check signature
        datos = Space(8)
        FileGet(1, datos)
        If datos <> Chr(137) & "PNG" & Chr(13) & Chr(10) & Chr(26) & Chr(10) Then
            Show_Error("Not a PNG file: " & file)
        End If

        ' Read header chunk
        datos = Space(4)
        FileGet(1, datos)
        FileGet(1, datos)

        If datos <> "IHDR" Then
            Show_Error("Incorrect PNG file: " & file)
        End If

        parsepng.w = freadint(1)
        parsepng.h = freadint(1)
        parsepng.file = file

        datos = Space(1)
        FileGet(1, datos)
        bpc = Asc(datos)
        If bpc > 8 Then
            Show_Error("16-bit depth not supported: " & file)
        End If

        FileGet(1, datos)
        ct = Asc(datos)
        If (ct = 0) Then
            colspace = "DeviceGray"
        ElseIf (ct = 2) Then
            colspace = "DeviceRGB"
        ElseIf (ct = 3) Then
            colspace = "Indexed"
        Else
            Show_Error("Alpha channel not supported: " & file)
        End If

        FileGet(1, datos)
        If (Asc(datos) <> 0) Then
            Show_Error("Unknown compression method: " & file)
        End If

        FileGet(1, datos)
        If (Asc(datos) <> 0) Then
            Show_Error("Unknown filter method: " & file)
        End If

        FileGet(1, datos)
        If (Asc(datos) <> 0) Then
            Show_Error("Interlacing not supported: " & file)
        End If

        datos = Space(4)
        FileGet(1, datos)

        parms = "/DecodeParms <</Predictor 15 /Colors " & IIf(ct = 2, 3, 1) & " /BitsPerComponent " & bpc & " /Columns " & parsepng.w & ">>"

        ' Scan chunks looking for palette, transparency and image data
        Dim pal As String = ""
        Dim trns As String = ""
        Dim imgdata As String = ""
        Do
            ni = freadint(1)
            datos = Space(4)
            FileGet(1, datos)
            typ = datos
            If (typ = "PLTE") Then
                ' Read palette
                pal = Space(ni)
                FileGet(1, pal)
                Seek(1, Seek(1) + 4)
            ElseIf (typ = "tRNS") Then
                ' Read transparency info
                '        t = fread(f, n)
                '        if(ct=0)
                '            trns = Array(ord(substr(t, 1, 1)))
                '        elseif(ct=2)
                '            trns = Array(ord(substr(t, 1, 1)), ord(substr(t, 3, 1)), ord(substr(t, 5, 1)))
                '        Else
                '        {
                '            pos = InStr(t, Chr(0))
                '            if(pos!=false)
                '                trns = Array(pos)
                '        }
                '        fread(f,4)
                '    }
                'UPGRADE_WARNING: Couldn't resolve default property of object typ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (typ = "IDAT") Then
                ' Read image data block
                datos = Space(ni)
                FileGet(1, datos)
                imgdata = imgdata & datos
                Seek(1, Seek(1) + 4)
            ElseIf (typ = "IEND") Then
                Exit Do
            Else
                Seek(1, Seek(1) + ni + 4)
            End If
        Loop
        If (colspace = "Indexed" And pal = "") Then
            Show_Error("Missing palette in " & file)
        End If
        FileClose(1)

        With parsepng
            .cs = colspace
            .bpc = bpc
            .F = "FlateDecode"
            .parms = parms
            .pal = pal
            .trns = trns
            .imgdata = imgdata
            .file = file
        End With
    End Function

    Private Function freadint(ByRef F As Integer) As Integer
        ' Read a 4-byte integer from file
        Dim byteArr(3) As Byte
        FileGet(F, byteArr)
        freadint = (CInt(byteArr(2)) * 256) + CInt(byteArr(3))
    End Function

    Private Function textstring(ByRef s As String) As String

        ' Format a text string
        ' Return "("._escape(s).")"
        textstring = "(" & escape(s) & ")"

    End Function

    Private Function escape(ByRef s As String) As String

        ' Add \ before \, ( and )
        escape = Replace(Replace(Replace(s, "\\", "\\\\"), "(", "\\("), ")", "\\)")

    End Function

    Private Sub putstream(ByRef s As String)

        out("stream")
        out(s)
        out("endstream")
    End Sub

    Private Sub out(ByRef s As String)

        ' Add a line to the document
        If (State = 2) Then
            pages(page) = pages(page) & s & vbLf
        Else
            Buffer = Buffer & s & vbLf
        End If
    End Sub

    Private Function substr_count(ByRef texto As String, ByRef paja As String) As Short
        substr_count = 0
        For a As Integer = 1 To Len(texto)
            If Mid(texto, a, 1) = paja Then
                substr_count = substr_count + 1
            End If
        Next
    End Function


    Private Function find_array(ByVal narray() As NewArray, ByRef akey As String) As Short
        find_array = -1
        For a As Integer = 1 To UBound(narray)
            If narray(a).Key = akey Then
                find_array = a
                Exit Function
            End If
        Next
        Exit Function
    End Function


    Private Function get_array_value(ByRef narray() As NewArray, ByRef akey As String) As String
        get_array_value = False
        For a As Integer = 1 To UBound(narray)
            If narray(a).Key = akey Then
                get_array_value = narray(a).Value
                Exit Function
            End If
        Next
        Exit Function
    End Function

    Private Function get_array_index(ByRef narray() As NewArray, ByRef akey As String) As Integer
        get_array_index = 0
        For a As Integer = 1 To UBound(narray)
            If narray(a).Key = akey Then
                get_array_index = a
                Exit Function
            End If
        Next
        Exit Function
    End Function

    Private Function Read_font(ByRef file As String) As NewArray
        Dim a As Integer
        Dim fx As String

        Dim sTemp As String
        Dim atemp() As String

        Read_font = New NewArray

        FileOpen(1, file, OpenMode.Input)
        'Read key
        Do Until EOF(1)
            sTemp = LineInput(1)
            If sTemp <> "" Then
                fx = Trim(Left(sTemp, InStr(1, sTemp, "=") - 1))
                Select Case LCase(fx)
                    Case "key"
                        Read_font.Key = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                    Case "type"
                        Read_font.Type = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                    Case "name"
                        Read_font.Name = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                    Case "desc"
                        Read_font.desc = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                    Case "enc"
                        Read_font.enc = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                    Case "up"
                        Read_font.up = CShort(Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "="))))
                    Case "ut"
                        Read_font.ut = CShort(Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "="))))
                    Case "diff"
                        If InStr(1, sTemp, "=") <> Len(Trim(sTemp)) Then
                            Read_font.diff = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=") - 1))
                        End If
                    Case "file"
                        Read_font.file = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                    Case "originalsize"
                        Read_font.originalsize = CInt(Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "="))))
                    Case "cw"
                        sTemp = Trim(Right(sTemp, Len(sTemp) - InStr(1, sTemp, "=")))
                        atemp = Split(sTemp, ",")
                        ReDim Read_font.cw(UBound(atemp))
                        For a = 0 To UBound(atemp)
                            If InStr(1, atemp(a), "chr") > 0 Then
                                Read_font.cw(a) = Chr(Val(Mid(atemp(a), InStr(1, atemp(a), "(") + 1, InStr(1, atemp(a), ")") - InStr(1, atemp(a), "("))))
                            ElseIf Len(Trim(atemp(a))) > 1 And a Mod 2 = 0 Then
                                MsgBox("Error Font file character=> " & atemp(a))
                            Else
                                Read_font.cw(a) = atemp(a)
                            End If
                        Next
                    Case Else
                        MsgBox("Wrong parameter font file: " & fx)
                End Select
            End If
        Loop
        FileClose(1)
    End Function



    Private Function get_img_array(ByRef narray() As ArrayImages, ByRef file As String) As Integer
        get_img_array = False
        For a As Integer = 1 To UBound(narray)
            If narray(a).file = file Then
                get_img_array = a
                Exit Function
            End If
        Next
        Exit Function
    End Function

    Private Function Mult(ByRef lsb As Byte, ByRef msb As Byte) As Integer
        Mult = lsb + (msb * CInt(256))
    End Function

    Private Sub Rotate(ByRef angle1 As Double, Optional ByRef x1 As Double = -1, Optional ByRef y1 As Double = -1)
        Dim c As Double
        Dim cx As Double
        Dim s As Double
        Dim cy As Double

        If (x1 = -1) Then x1 = X
        If (y1 = -1) Then y1 = Y
        If (angle <> 0) Then
            Call out("Q")
        End If
        angle = angle1
        If (angle1 <> 0) Then

            angle1 = angle1 * 3.141592 / 180
            c = System.Math.Cos(angle1)
            s = System.Math.Sin(angle1)
            cx = x1 * K
            cy = (h - y1) * K
            Call out(sprintf("q %.5f %.5f %.5f %.5f %.2f %.2f cm 1 0 0 1 %.2f %.2f cm", c, s, -s, c, cx, cy, -cx, -cy))
        End If
    End Sub

    Sub RotatedText(ByRef x1 As Double, ByRef y1 As Double, ByRef txt As String, ByRef angle1 As Double)
        'Text rotated around its origin
        Call Rotate(angle1, x1, y1)
        Call text(x1, y1, txt)
        Call Rotate(0)
    End Sub

    Sub RoundedRect(ByRef x1 As Double, ByRef y1 As Double, ByRef w1 As Double, ByRef h1 As Double, ByRef R As Double, Optional ByRef style As String = "")
        Dim MyArc As Double
        Dim op As String
        Dim hp As Double
        Dim k1 As Double
        Dim yc As Double
        Dim xc As Double

        k1 = K
        hp = h
        If (style = "F") Then
            op = "f"
        ElseIf (style = "FD" Or style = "DF") Then
            op = "B"
        Else
            op = "S"
        End If
        MyArc = 4 / 3 * (System.Math.Sqrt(2) - 1)
        Call out(sprintf("%.2f %.2f m", (x1 + R) * k1, (hp - y1) * k1))
        xc = x1 + w1 - R
        yc = y1 + R
        Call out(sprintf("%.2f %.2f l", xc * k1, (hp - y1) * k1))

        Call Arc(xc + R * MyArc, yc - R, xc + R, yc - R * MyArc, xc + R, yc)
        xc = x1 + w1 - R
        yc = y1 + h1 - R
        Call out(sprintf("%.2f %.2f l", (x1 + w1) * k1, (hp - yc) * k1))
        Call Arc(xc + R, yc + R * MyArc, xc + R * MyArc, yc + R, xc, yc + R)
        xc = x1 + R
        yc = y1 + h1 - R
        Call out(sprintf("%.2f %.2f l", xc * k1, (hp - (y1 + h1)) * k1))
        Call Arc(xc - R * MyArc, yc + R, xc - R, yc + R * MyArc, xc - R, yc)
        xc = x1 + R
        yc = y1 + R
        Call out(sprintf("%.2f %.2f l", (x1) * k1, (hp - yc) * k1))
        Call Arc(xc - R, yc - R * MyArc, xc - R * MyArc, yc - R, xc, yc - R)
        Call out(op)

    End Sub

    Private Sub Arc(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, ByRef x3 As Double, ByRef y3 As Double)
        Dim h1 As Double
        h1 = h
        Call out(sprintf("%.2f %.2f %.2f %.2f %.2f %.2f c ", x1 * K, (h1 - y1) * K, x2 * K, (h1 - y2) * K, x3 * K, (h1 - y3) * K))
    End Sub

    Private Function sprintf(ByVal cadena As String, ByVal ParamArray data() As Object) As String

        Dim f As Boolean = False
        Dim i As Integer = 0
        Dim a As Integer = 0
        Dim param As String = ""
        Dim BeginPos As Integer

        Do
            If f Then
                If InStr("gGxXeEpnfcdiosu%", cadena.Substring(a, 1)) Then
                    param = cadena.Substring(BeginPos, a - BeginPos + 1)
                    Dim p As String
                    If param.Substring(1, 1) = "." Then
                        p = "{" & i & ":0."
                        For c As Integer = 1 To CInt(param.Substring(2, 1))
                            p = p & "0"
                        Next
                        p = p & "}"
                    Else
                        Select Case param.Substring(param.Length - 1)
                            Case "f"
                                p = "{" & i & ":.3f}"
                            Case "d"
                                p = "{" & i & ":d}"
                            Case "s"
                                p = "{" & i & "}"
                            Case Else
                                p = ""
                        End Select
                    End If
                    i += 1
                    cadena = cadena.Remove(BeginPos, a - BeginPos + 1)
                    cadena = cadena.Insert(BeginPos, p)
                    f = False
                End If
            End If
            If Not f Then
                If cadena.Substring(a, 1) = "%" Then
                    f = True
                    BeginPos = a
                End If
            End If
            a += 1
        Loop Until a = cadena.Length
        Return String.Format(cadena, data)
    End Function



    Public Sub show_preview(ByVal filename As String, Optional ByVal text As String = "")

        Try


            Dim f As New Form
            'Dim AxAcroPDF1 = New AxAcroPDFLib.AxAcroPDF
            With f

                '   CType(AxAcroPDF1, System.ComponentModel.ISupportInitialize).BeginInit()
                .SuspendLayout()
                '  AxAcroPDF1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                '        Or System.Windows.Forms.AnchorStyles.Left) _
                '        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                ' AxAcroPDF1.Enabled = True
                'AxAcroPDF1.Location = New System.Drawing.Point(12, 12)
                'AxAcroPDF1.Name = "AxAcroPDF1"
                'AxAcroPDF1.OcxState = CType(resources.GetObject("AxAcroPDF1.OcxState"), System.Windows.Forms.AxHost.State)
                ' AxAcroPDF1.Size = New System.Drawing.Size(268, 242)
                ' AxAcroPDF1.TabIndex = 0
                .Name = "AxPDF"
                'CType(AxAcroPDF1, System.ComponentModel.ISupportInitialize).EndInit()
                '.Controls.Add(AxAcroPDF1)
                .ResumeLayout(False)
                '.MdiParent = main
            End With
            f.Size = New Size(600, 600)
            f.Text = text
            f.Show()
            '            AxAcroPDF1.LoadFile(filename)

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

    End Sub

    Private Function gzcompress(ByVal text As String) As String

        Dim data() As Byte = System.Text.Encoding.Default.GetBytes(text)

        'Original data size
        Dim OriginalSize As Long = UBound(data) + 1

        'Fields for the temporary buffer
        Dim result As Integer

        'Resizes buffers
        Dim BufferSize As Integer = UBound(data) + 1
        BufferSize = CInt(BufferSize + (BufferSize * 0.01) + 12)
        Dim tempBuffer(BufferSize) As Byte
        Try
            result = CompressByteArray(tempBuffer, BufferSize, data, UBound(data) + 1)
        Catch ex As Exception
            result = -1
        End Try

        If result = 0 Then
            ReDim data(BufferSize - 1)
            Array.Copy(tempBuffer, data, BufferSize)
            Return System.Text.Encoding.Default.GetString(data)
        Else
            compress = False
            Return text
        End If

    End Function

End Class
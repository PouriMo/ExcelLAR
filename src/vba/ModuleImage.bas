Attribute VB_Name = "ModuleImage"
Option Explicit

' ============================================================================
'  Images for the userform panels and the PNG export.
'
'  Panels: the old code exported the chart as a GIF (256 colours, dithered) and let the
'  Image control stretch it (nearest-neighbour), which made it look pixelated. Now the chart
'  is exported as PNG and resampled with GDI+ (high-quality bicubic) to exactly the pixel
'  size of the panel, so the control shows it 1:1.
'
'  Export: for a scale factor k > 1 the chart is redrawn on a temporary sheet k times larger
'  (fonts, lines and markers too), exported, and the temporary sheet is deleted. The result
'  is a real high-resolution image, not an enlarged one.
' ============================================================================

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hImage As LongPtr
    hPal As LongPtr
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef image As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal scan0 As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, ByRef graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal mode As Long) As Long
Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal mode As Long) As Long
Private Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As LongPtr, ByVal color As Long) As Long
Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal image As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32" (ByVal clr As Long, ByVal hpal As LongPtr, ByRef lpcolorref As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long

Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const InterpolationModeHighQualityBicubic As Long = 7
Private Const PixelOffsetModeHighQuality As Long = 2
Private Const LOGPIXELSX As Long = 88
Private Const PICTYPE_BITMAP As Long = 1

' Screen resolution in pixels per inch (96 at 100 % Windows scaling).
Public Function ScreenDpi() As Long
    Dim hdc As LongPtr
    hdc = GetDC(0)
    ScreenDpi = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    If ScreenDpi <= 0 Then ScreenDpi = 96
End Function

' Converts an OLE colour (which may be a system colour such as &H8000000F) to an RGB value.
Public Function OleColorToRGB(ByVal oleColor As Long) As Long
    Dim c As Long
    If OleTranslateColor(oleColor, 0, c) = 0 Then OleColorToRGB = c Else OleColorToRGB = vbWhite
End Function

' Pixel size of an image file (PNG, GIF, JPG, BMP).
Public Function GetImageSizePx(ByVal path As String, ByRef w As Long, ByRef h As Long) As Boolean
    Dim si As GdiplusStartupInput, token As LongPtr, img As LongPtr
    w = 0: h = 0
    si.GdiplusVersion = 1
    If GdiplusStartup(token, si, 0) <> 0 Then Exit Function
    If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
        GdipGetImageWidth img, w
        GdipGetImageHeight img, h
        GdipDisposeImage img
        GetImageSizePx = (w > 0 And h > 0)
    End If
    GdiplusShutdown token
End Function

' Loads an image file and resamples it (high-quality bicubic) to fit inside maxW x maxH pixels,
' keeping its proportions, on a background of colour bgRGB. Returns Nothing on failure.
Public Function LoadPictureHQ(ByVal path As String, ByVal maxW As Long, ByVal maxH As Long, _
                              Optional ByVal bgRGB As Long = vbWhite) As StdPicture
    Dim si As GdiplusStartupInput, token As LongPtr, img As LongPtr, bmp As LongPtr, gfx As LongPtr, hbm As LongPtr
    Dim w As Long, h As Long, tw As Long, th As Long, s As Double
    Dim pd As PICTDESC, iid As GUID, pic As IPicture

    si.GdiplusVersion = 1
    If GdiplusStartup(token, si, 0) <> 0 Then Exit Function
    If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
        GdipGetImageWidth img, w
        GdipGetImageHeight img, h
        If w > 0 And h > 0 And maxW > 0 And maxH > 0 Then
            s = maxW / w
            If maxH / h < s Then s = maxH / h
            tw = CLng(w * s): th = CLng(h * s)
            If tw < 1 Then tw = 1
            If th < 1 Then th = 1
            If GdipCreateBitmapFromScan0(tw, th, 0, PixelFormat32bppARGB, 0, bmp) = 0 Then
                If GdipGetImageGraphicsContext(bmp, gfx) = 0 Then
                    GdipGraphicsClear gfx, ToARGB(bgRGB)
                    GdipSetInterpolationMode gfx, InterpolationModeHighQualityBicubic
                    GdipSetPixelOffsetMode gfx, PixelOffsetModeHighQuality
                    GdipDrawImageRectI gfx, img, 0, 0, tw, th
                    GdipDeleteGraphics gfx
                    If GdipCreateHBITMAPFromBitmap(bmp, hbm, ToARGB(bgRGB)) = 0 Then
                        pd.cbSizeOfStruct = LenB(pd)
                        pd.picType = PICTYPE_BITMAP
                        pd.hImage = hbm
                        ' IID_IPicture = {7BF80980-BF32-101A-8BBB-00AA00300CAB}
                        iid.Data1 = &H7BF80980: iid.Data2 = &HBF32: iid.Data3 = &H101A
                        iid.Data4(0) = &H8B: iid.Data4(1) = &HBB: iid.Data4(2) = &H0: iid.Data4(3) = &HAA
                        iid.Data4(4) = &H0: iid.Data4(5) = &H30: iid.Data4(6) = &HC: iid.Data4(7) = &HAB
                        If OleCreatePictureIndirect(pd, iid, 1, pic) = 0 Then Set LoadPictureHQ = pic
                    End If
                End If
                GdipDisposeImage bmp
            End If
        End If
        GdipDisposeImage img
    End If
    GdiplusShutdown token
End Function

Private Function ToARGB(ByVal rgbColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = rgbColor And &HFF
    g = (rgbColor \ &H100) And &HFF
    b = (rgbColor \ &H10000) And &HFF
    ToARGB = &HFF000000 Or (r * &H10000) Or (g * &H100) Or b
End Function

' Exports the thrust-line (plotKind = "Thrust") or collapse-mechanism ("Deformed") chart as PNG.
' scaleK = 1 exports the chart already on the ThrustPlot / DeformedPlot sheet (wsData may be
' Nothing); scaleK > 1 redraws it k times larger on a temporary sheet from data sheet wsData.
Public Function ExportPlotPNG(wbUser As Workbook, wsData As Worksheet, ByVal plotKind As String, _
                              ByVal scaleK As Double, ByVal outPath As String) As Boolean
    Dim srcName As String, tmpName As String, ws As Worksheet, prevSheet As Object
    Dim prevSU As Boolean, prevAlerts As Boolean

    If plotKind = "Thrust" Then srcName = "ThrustPlot" Else srcName = "DeformedPlot"
    On Error GoTo Fail

    If scaleK <= 1 Or wsData Is Nothing Then
        wbUser.Sheets(srcName).ChartObjects(1).Chart.Export Filename:=outPath, FilterName:="PNG"
        ExportPlotPNG = True
        Exit Function
    End If

    tmpName = srcName & "_HiRes"
    prevSU = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    Set prevSheet = wbUser.ActiveSheet
    Application.ScreenUpdating = False

    If plotKind = "Thrust" Then
        PlotThrustLine wsData, scaleK, tmpName
    Else
        PlotDeformedShape wsData, scaleK, tmpName, True
    End If
    Set ws = wbUser.Sheets(tmpName)
    ws.ChartObjects(1).Chart.Export Filename:=outPath, FilterName:="PNG"
    ExportPlotPNG = True

Cleanup:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = prevAlerts
    If Not prevSheet Is Nothing Then prevSheet.Activate
    Application.ScreenUpdating = prevSU
    Exit Function

Fail:
    ExportPlotPNG = False
    If Len(tmpName) = 0 Then Exit Function
    Resume FindTemp
FindTemp:
    On Error Resume Next
    Set ws = Nothing
    Set ws = wbUser.Sheets(tmpName)
    GoTo Cleanup
End Function

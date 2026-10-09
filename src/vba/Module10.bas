Attribute VB_Name = "Module10"
Option Explicit

' ============================================================================
'  Equal-scale axes for the XY charts (thrust line and collapse mechanism).
'  Without this, Excel scales x and y independently and the arch is distorted.
' ============================================================================

' Makes 1 drawing unit (mm) on x as long as 1 drawing unit on y and gives both axes the same
' round major unit (square grid). [xMin..xMax] x [yMin..yMax] is the bounding box of everything
' that has to be visible; padFraction is the free margin on each side.
Public Sub SetEqualAspectAxes(ch As Chart, ByVal xMin As Double, ByVal xMax As Double, _
                              ByVal yMin As Double, ByVal yMax As Double, _
                              Optional ByVal padFraction As Double = 0.06)
    Dim dx As Double, dy As Double, cx As Double, cy As Double
    Dim w As Double, h As Double, s As Double, hx As Double, hy As Double
    Dim stepU As Double, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim pass As Long

    dx = xMax - xMin: dy = yMax - yMin
    If dx <= 0 Then dx = 1
    If dy <= 0 Then dy = 1
    cx = (xMin + xMax) / 2: cy = (yMin + yMax) / 2
    dx = dx * (1 + 2 * padFraction): dy = dy * (1 + 2 * padFraction)

    ' The inside size of the plot area changes slightly when the axis labels change, so iterate.
    For pass = 1 To 3
        w = ch.PlotArea.InsideWidth: h = ch.PlotArea.InsideHeight
        If w <= 1 Or h <= 1 Then Exit For
        s = w / dx                                   ' points per drawing unit
        If h / dy < s Then s = h / dy                ' the tighter direction decides
        hx = w / s / 2: hy = h / s / 2               ' half ranges that fill the plot area exactly
        If hx > hy Then
            stepU = NiceStep(2 * hx / 8)
        Else
            stepU = NiceStep(2 * hy / 8)
        End If
        Call FitWindow(xMin, xMax, cx, hx, stepU, x0, x1)
        Call FitWindow(yMin, yMax, cy, hy, stepU, y0, y1)

        On Error Resume Next
        Call SetAxisRange(ch.Axes(xlCategory), x0, x1)
        Call SetAxisRange(ch.Axes(xlValue), y0, y1)
        ch.Axes(xlCategory).MajorUnit = stepU
        ch.Axes(xlValue).MajorUnit = stepU
        On Error GoTo 0
        DoEvents
    Next pass
End Sub

' Window [lo, hi] of fixed length 2*half centred on c; its lower limit is moved to a multiple of
' stepU (so tick labels are round numbers) provided the data [dMin, dMax] stay inside.
Private Sub FitWindow(ByVal dMin As Double, ByVal dMax As Double, ByVal c As Double, ByVal half As Double, _
                      ByVal stepU As Double, ByRef lo As Double, ByRef hi As Double)
    Dim cand As Double
    lo = c - half: hi = c + half
    cand = Int(lo / stepU) * stepU
    If cand <= dMin And cand + 2 * half >= dMax Then
        lo = cand: hi = cand + 2 * half
    Else
        cand = cand + stepU
        If cand <= dMin And cand + 2 * half >= dMax Then
            lo = cand: hi = cand + 2 * half
        End If
    End If
End Sub

' Sets min and max in an order that never makes Excel reject the new pair.
Private Sub SetAxisRange(ax As Axis, ByVal lo As Double, ByVal hi As Double)
    If hi > ax.MinimumScale Then
        ax.MaximumScale = hi
        ax.MinimumScale = lo
    Else
        ax.MinimumScale = lo
        ax.MaximumScale = hi
    End If
End Sub

Private Function NiceStep(ByVal rawStep As Double) As Double
    Dim p As Double, m As Double
    If rawStep <= 0 Then NiceStep = 1: Exit Function
    p = 10 ^ Int(Log(rawStep) / Log(10#))
    m = rawStep / p
    If m <= 1 Then
        NiceStep = p
    ElseIf m <= 2 Then
        NiceStep = 2 * p
    ElseIf m <= 5 Then
        NiceStep = 5 * p
    Else
        NiceStep = 10 * p
    End If
End Function

' Extends a bounding box with the numeric (x, y) pairs found in two worksheet columns.
' Blank cells and #N/A separators are skipped. gotAny must be False before the first call.
Public Sub ExtendBoundsFromColumns(ws As Worksheet, ByVal xCol As Long, ByVal yCol As Long, _
                                   ByVal r1 As Long, ByVal r2 As Long, _
                                   ByRef xMin As Double, ByRef xMax As Double, _
                                   ByRef yMin As Double, ByRef yMax As Double, ByRef gotAny As Boolean)
    Dim r As Long, vx As Variant, vy As Variant, px As Double, py As Double
    For r = r1 To r2
        vx = ws.Cells(r, xCol).Value
        vy = ws.Cells(r, yCol).Value
        If Not IsError(vx) Then
            If Not IsError(vy) Then
                If Not IsEmpty(vx) And Not IsEmpty(vy) Then
                    If IsNumeric(vx) And IsNumeric(vy) Then
                        px = CDbl(vx): py = CDbl(vy)
                        If Not gotAny Then
                            xMin = px: xMax = px: yMin = py: yMax = py: gotAny = True
                        Else
                            If px < xMin Then xMin = px
                            If px > xMax Then xMax = px
                            If py < yMin Then yMin = py
                            If py > yMax Then yMax = py
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub

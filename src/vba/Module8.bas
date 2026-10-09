Attribute VB_Name = "Module8"
Option Explicit

' ---- settings for the collapse-mechanism drawing -----------------------------------------
' Largest movement of any block vertex, as a fraction of the larger arch dimension.
' 0.04 = the deformed shape is drawn with a maximum movement of 4 % of the span.
Private Const DEF_FRACTION As Double = 0.04
' Backfill is only a load. TRUE re-enables the old trick of rotating backfill pieces away
' from the blocks (not recommended: the backfill ends up detached from the arch).
Private Const NUDGE_BACKFILL As Boolean = False
Private Const MECH_PI As Double = 3.14159265358979

' scaleK   : multiplies the chart size, fonts, line weights and markers (used for hi-res PNG export)
' plotName : name of the sheet that receives the chart
' silent   : True suppresses the informational message boxes (used when re-drawing for an export)
Sub PlotDeformedShape(Optional wsDataIn As Worksheet, Optional ByVal scaleK As Double = 1, _
                      Optional ByVal plotName As String = "DeformedPlot", Optional ByVal silent As Boolean = False)
    Dim wsData As Worksheet, wsForces As Worksheet, wsPlot As Worksheet
    Dim wbUser As Workbook
    Dim i As Long, j As Long, NI As Long, NB As Long, lam As Double

    If Not wsDataIn Is Nothing Then
        Set wsData = wsDataIn
        Set wbUser = wsDataIn.Parent
    Else
        Set wbUser = ActiveWorkbook
        Set wsData = wbUser.ActiveSheet
    End If

    On Error Resume Next
    Set wsForces = wbUser.Sheets("ForcesTable")
    On Error GoTo ErrorHandler
    If wsForces Is Nothing Then MsgBox "Run analysis first!", vbCritical: Exit Sub

    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant
    Call ReadCSVGeometry(wsData, bx, by, fx, fy, NB)
    If NB = 0 Then
        MsgBox "No block geometry found on sheet '" & wsData.Name & "'." & vbCrLf & _
               "Make sure your CSV data sheet is selected/active before plotting.", vbExclamation
        Exit Sub
    End If
    NI = NB + 1

    Dim tArr() As Double: tArr = GetInterfaceThickness(wsData, NI)
    Dim Nvec() As Double, Mvec() As Double: ReDim Nvec(1 To NI): ReDim Mvec(1 To NI)
    For i = 1 To NI
        Nvec(i) = wsForces.Cells(i + 1, 2).Value
        Mvec(i) = wsForces.Cells(i + 1, 4).Value
    Next i
    lam = wsForces.Range("G2").Value

    Dim p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double
    ReDim p1x(1 To NI), p1y(1 To NI), p2x(1 To NI), p2y(1 To NI)

    Dim FiArr() As Double: ReDim FiArr(1 To NI)
    For i = 1 To NI: FiArr(i) = wsData.Cells(i + 1, 8).Value: Next i

    For i = 1 To NI
        Dim curX() As Double, curY() As Double
        Dim prevX() As Double, prevY() As Double

        If i = 1 Then
            curX = bx(1): curY = by(1)
            FindEdgeByNormal curX, curY, FiArr(1), p1x(1), p1y(1), p2x(1), p2y(1)
        ElseIf i = NI Then
            curX = bx(NB): curY = by(NB)
            FindEdgeByNormal curX, curY, FiArr(NI), p1x(NI), p1y(NI), p2x(NI), p2y(NI)
        Else
            prevX = bx(i - 1): prevY = by(i - 1)
            curX = bx(i): curY = by(i)
            FindSharedEdge prevX, prevY, curX, curY, p1x(i), p1y(i), p2x(i), p2y(i)
        End If
        OrientEdgeByFi p1x(i), p1y(i), p2x(i), p2y(i), FiArr(i)
    Next i

    Dim thrX() As Double, thrY() As Double: ReDim thrX(1 To NI): ReDim thrY(1 To NI)
    Dim eVec() As Double: ReDim eVec(1 To NI)

    Dim Nref As Double: Nref = 0
    For i = 1 To NI
        If Abs(Nvec(i)) > Nref Then Nref = Abs(Nvec(i))
    Next i
    If Nref = 0 Then Nref = 1

    For i = 1 To NI
        eVec(i) = 0
        If Abs(Nvec(i)) > 0.000001 * Nref Then eVec(i) = -1 * Mvec(i) / Nvec(i)
        Dim dx As Double, dy As Double, L As Double
        dx = p2x(i) - p1x(i): dy = p2y(i) - p1y(i): L = Sqr(dx * dx + dy * dy)
        If L > 0 Then dx = dx / L: dy = dy / L
        Dim e As Double: e = eVec(i)
        If e > tArr(i) / 2 Then e = tArr(i) / 2
        If e < -tArr(i) / 2 Then e = -tArr(i) / 2
        thrX(i) = (p1x(i) + p2x(i)) / 2 + e * dx: thrY(i) = (p1y(i) + p2y(i)) / 2 + e * dy
    Next i

    ' Four hinges = the admissible mechanism with the smallest kinematic multiplier
    ' (FindCollapseMechanism in Module9). It also gives the sense of the collapse.
    Dim hIfaces() As Long, hSide() As Long, hX() As Double, hY() As Double, lamKin As Double, mechSign As Long
    If Not FindCollapseMechanism(wsData, NB, bx, by, p1x, p1y, p2x, p2y, tArr, eVec, _
                                 hIfaces, hSide, hX, hY, lamKin, mechSign) Then
        MsgBox "No admissible four-hinge collapse mechanism was found: no set of four interfaces where the " & _
               "thrust line touches the intrados/extrados can open as a mechanism under the applied load." & vbCrLf & _
               "Check the thrust line and run DumpHingeDiagnostics (sheet 'Diag').", vbExclamation
        Exit Sub
    End If
    Dim boundaries(1 To 4) As Long
    For i = 1 To 4: boundaries(i) = hIfaces(i): Next i

    ' hinge points on the intrados/extrados edge the thrust line touches
    Dim PLx As Double, PLy As Double, H1x As Double, H1y As Double, H2x As Double, H2y As Double, PRx As Double, PRy As Double
    PLx = hX(1): PLy = hY(1)
    H1x = hX(2): H1y = hY(2)
    H2x = hX(3): H2y = hY(3)
    PRx = hX(4): PRy = hY(4)

    Dim segOf() As Long: ReDim segOf(1 To NB)
    Dim segTmp As Long: segTmp = 0
    For i = 1 To NB
        If i = boundaries(1) Then segTmp = 1
        If i = boundaries(2) Then segTmp = 2
        If i = boundaries(3) Then segTmp = 3
        If i = boundaries(4) Then segTmp = 4
        segOf(i) = segTmp
    Next i

    Dim chains(0 To 4) As String, seg As Long: seg = 0
    For i = 1 To NB
        If i = boundaries(1) Then seg = 1
        If i = boundaries(2) Then seg = 2
        If i = boundaries(3) Then seg = 3
        If i = boundaries(4) Then seg = 4
        chains(seg) = chains(seg) & i & ","
    Next i

    ' =====================================================================
    '  COLLAPSE MECHANISM  (rewritten)
    '
    '  * The sense of rotation at each hinge is NOT searched for by trial and
    '    error: it follows from the thrust line. At a hinge the thrust line
    '    touches the intrados or the extrados and the joint must open on the
    '    OPPOSITE edge. A rotation of the other sign closes the joint and crushes
    '    the masonry at the contact point (compression failure) - this is what
    '    produced the "seagull" shapes.
    '  * The four-bar branch is the one continuous with the undeformed shape.
    '  * The amplitude is set from a target displacement (DEF_FRACTION of the arch
    '    size) instead of taking the largest rotation that does not interpenetrate.
    ' =====================================================================
    Dim theta1 As Double, theta3 As Double, mcTheta2 As Double
    Dim H1nx As Double, H1ny As Double, H2nx As Double, H2ny As Double
    Dim Rmid(1 To 2, 1 To 2) As Double, TmX As Double, TmY As Double

    ' (1) admissible sense of the relative rotation at each hinge (+1 = counter-clockwise):
    '     a hinge on side s of the interface opens on the other edge only if s * rotation > 0
    '     (0 = pin, either sense)
    Dim mcReq(1 To 4) As Long
    For i = 1 To 4
        mcReq(i) = hSide(i)
    Next i

    ' (2) target displacement from the size of the arch
    Dim mcXmin As Double, mcXmax As Double, mcYmin As Double, mcYmax As Double
    Dim mcGot As Boolean, mcX As Double, mcY As Double
    mcGot = False
    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            For j = 1 To UBound(bx(i))
                mcX = bx(i)(j): mcY = by(i)(j)
                If Not mcGot Then
                    mcXmin = mcX: mcXmax = mcX: mcYmin = mcY: mcYmax = mcY: mcGot = True
                Else
                    If mcX < mcXmin Then mcXmin = mcX
                    If mcX > mcXmax Then mcXmax = mcX
                    If mcY < mcYmin Then mcYmin = mcY
                    If mcY > mcYmax Then mcYmax = mcY
                End If
            Next j
        End If
    Next i
    Dim mcTarget As Double
    mcTarget = mcXmax - mcXmin
    If mcYmax - mcYmin > mcTarget Then mcTarget = mcYmax - mcYmin
    mcTarget = DEF_FRACTION * mcTarget

    ' (3) crank angle: start small and rescale until the largest vertex displacement equals the target
    Dim mcSign As Long, mcIter As Long, mcOK As Boolean, mcDisp As Double, mcPen As Double, mcTry As Long
    mcSign = mechSign                    ' sense in which the live load does positive work
    theta1 = mcSign * 0.5 * MECH_PI / 180#
    For mcIter = 1 To 5
        If Not MechanismState(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy, theta1, _
                              H1nx, H1ny, H2nx, H2ny, mcTheta2, theta3, Rmid, TmX, TmY) Then Exit For
        mcDisp = MaxVertexDisplacement(bx, by, NB, segOf, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
        If mcDisp < 0.000000001 Then Exit For
        If Abs(mcDisp - mcTarget) < 0.01 * mcTarget Then Exit For
        theta1 = theta1 * mcTarget / mcDisp
    Next mcIter

    ' (4) final state; if the blocks of different links interpenetrate, halve the crank angle
    mcOK = MechanismState(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy, theta1, _
                          H1nx, H1ny, H2nx, H2ny, mcTheta2, theta3, Rmid, TmX, TmY)
    mcTry = 0
    Do
        If mcOK Then
            mcPen = MaxPenetrationFull(bx, by, fx, fy, NB, segOf, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY, False)
            If mcPen <= 0 Then Exit Do
        End If
        mcTry = mcTry + 1
        If mcTry > 8 Then Exit Do
        theta1 = theta1 / 2
        mcOK = MechanismState(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy, theta1, _
                              H1nx, H1ny, H2nx, H2ny, mcTheta2, theta3, Rmid, TmX, TmY)
    Loop
    If Not mcOK Then
        MsgBox "The four-hinge linkage cannot be closed for the detected hinges, so no mechanism can be drawn." & _
               vbCrLf & "Check the hinge detection (DumpHingeDiagnostics, sheet 'Diag').", vbExclamation
        Exit Sub
    End If
    If mcPen > 0 And Not silent Then
        MsgBox "Note: a small residual overlap (about " & Format(mcPen, "0.00") & " mm) remains between " & _
               "blocks of different links in the drawn mechanism.", vbInformation
    End If

    ' (5) consistency check: every hinge must open on the side required by the thrust line
    Dim mcW(1 To 4) As Double, mcBad As Boolean
    mcW(1) = theta1
    mcW(2) = NormAngle(mcTheta2 - theta1)
    mcW(3) = NormAngle(theta3 - mcTheta2)
    mcW(4) = NormAngle(-theta3)
    mcBad = False
    For i = 1 To 4
        If mcReq(i) <> 0 Then
            If Sgn(mcW(i)) <> mcReq(i) Then mcBad = True
        End If
    Next i
    If mcBad And Not silent Then
        MsgBox "Warning: the rotations of the four-bar mechanism do not open all four hinges on the side " & _
               "required by the thrust line. The hinge set is probably not a true collapse mechanism " & _
               "(check the hinge detection on sheet 'Diag').", vbExclamation
    End If

    Application.DisplayAlerts = False
    On Error Resume Next: wbUser.Sheets(plotName).Delete: On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    Set wsPlot = wbUser.Sheets.Add: wsPlot.Name = plotName

    wsPlot.Cells(1, 1).Value = "OrigX": wsPlot.Cells(1, 2).Value = "OrigY"
    Dim origRow As Long: origRow = 2

    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            Dim oax() As Double, oay() As Double
            oax = bx(i): oay = by(i)
            For j = 1 To UBound(oax)
                wsPlot.Cells(origRow, 1).Value = oax(j): wsPlot.Cells(origRow, 2).Value = oay(j): origRow = origRow + 1
            Next j
            wsPlot.Cells(origRow, 1).Value = oax(1): wsPlot.Cells(origRow, 2).Value = oay(1): origRow = origRow + 1
            wsPlot.Cells(origRow, 1).Value = CVErr(xlErrNA): wsPlot.Cells(origRow, 2).Value = CVErr(xlErrNA): origRow = origRow + 1
        End If
    Next i
    Dim origRows As Long: origRows = origRow - 1

    wsPlot.Cells(1, 3).Value = "DefX": wsPlot.Cells(1, 4).Value = "DefY"
    ' Moved DefBackX/DefBackY from cols 11-12 to cols 7-8 (right after
    ' the hinge cols 5-6), so they sit in the left "data area" of the
    ' sheet instead of under the chart. The chart is placed starting
    ' at Left:=300 which is roughly column K, so cols 7-8 are safely
    ' clear of it.
    wsPlot.Cells(1, 7).Value = "DefBackX": wsPlot.Cells(1, 8).Value = "DefBackY"

    Dim defRow As Long: defRow = 2
    Dim defBackRow As Long: defBackRow = 2
    Dim segIdx As Long, bIdx As Long, bnum As Long, parts() As String
    Dim npx As Double, npy As Double, px As Double, py As Double

    Dim blockObstacles As Collection: Set blockObstacles = New Collection
    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            blockObstacles.Add TransformPolyBySeg(bx(i), by(i), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
        End If
    Next i
    Dim backfillObstacles As Collection: Set backfillObstacles = New Collection

    For segIdx = 0 To 4
        If Len(chains(segIdx)) > 0 Then
            parts = Split(Left(chains(segIdx), Len(chains(segIdx)) - 1), ",")
            For bIdx = 0 To UBound(parts)
                bnum = CLng(parts(bIdx))

                If Not IsEmpty(bx(bnum)) Then
                    Dim ax() As Double, ay() As Double
                    ax = bx(bnum): ay = by(bnum)
                    For j = 1 To UBound(ax)
                        px = ax(j): py = ay(j)
                        TransformPoint px, py, segIdx, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY, npx, npy
                        wsPlot.Cells(defRow, 3).Value = npx: wsPlot.Cells(defRow, 4).Value = npy: defRow = defRow + 1
                    Next j
                    px = ax(1): py = ay(1)
                    TransformPoint px, py, segIdx, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY, npx, npy
                    wsPlot.Cells(defRow, 3).Value = npx: wsPlot.Cells(defRow, 4).Value = npy: defRow = defRow + 1
                    wsPlot.Cells(defRow, 3).Value = CVErr(xlErrNA): wsPlot.Cells(defRow, 4).Value = CVErr(xlErrNA): defRow = defRow + 1
                End If

                If Not fx(bnum) Is Nothing Then
                    Dim k As Long
                    For k = 1 To fx(bnum).Count
                        Dim bfx() As Double, bfy() As Double
                        bfx = fx(bnum).Item(k): bfy = fy(bnum).Item(k)
                        Dim bfPoly() As Double, bfNudged() As Double
                        bfPoly = TransformPolyBySeg(bfx, bfy, segIdx, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
                        If NUDGE_BACKFILL Then
                            Dim clearAgainst As Collection: Set clearAgainst = New Collection
                            Dim ob As Variant
                            For Each ob In blockObstacles: clearAgainst.Add ob: Next ob
                            For Each ob In backfillObstacles: clearAgainst.Add ob: Next ob
                            bfNudged = NudgeClearByRotation(bfPoly, clearAgainst)
                        Else
                            bfNudged = bfPoly          ' backfill is a load: it simply follows its block
                        End If
                        backfillObstacles.Add bfNudged

                        Dim bn As Long: bn = UBound(bfNudged, 1)
                                                For j = 1 To bn
                            wsPlot.Cells(defBackRow, 7).Value = bfNudged(j, 1): wsPlot.Cells(defBackRow, 8).Value = bfNudged(j, 2): defBackRow = defBackRow + 1
                        Next j
                        wsPlot.Cells(defBackRow, 7).Value = bfNudged(1, 1): wsPlot.Cells(defBackRow, 8).Value = bfNudged(1, 2): defBackRow = defBackRow + 1
                        wsPlot.Cells(defBackRow, 7).Value = CVErr(xlErrNA): wsPlot.Cells(defBackRow, 8).Value = CVErr(xlErrNA): defBackRow = defBackRow + 1
                    Next k
                End If
            Next bIdx
        End If
    Next segIdx

    wsPlot.Cells(2, 5).Value = PLx: wsPlot.Cells(2, 6).Value = PLy
    wsPlot.Cells(3, 5).Value = H1nx: wsPlot.Cells(3, 6).Value = H1ny
    wsPlot.Cells(4, 5).Value = H2nx: wsPlot.Cells(4, 6).Value = H2ny
    wsPlot.Cells(5, 5).Value = PRx: wsPlot.Cells(5, 6).Value = PRy

        ' Chart moved further right (Left:=550) so columns A..J are fully
    ' available for the analysis data without being obscured.
        ' Chart is larger AND moved further right so it never overlaps the
    ' data columns (which now include BackX/BackY at H..I and force
    ' arrow data at O..R after Fix 3 below).
        Dim ch As ChartObject: Set ch = wsPlot.ChartObjects.Add(Left:=800, Top:=10, Width:=750 * scaleK, Height:=420 * scaleK)
    With ch.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True: .ChartTitle.Text = "Collapse Mechanism - lambda = " & Format$(lam, "0.00") & " [N]"
        .ChartTitle.Font.Size = 16 * scaleK: .ChartTitle.Font.Bold = True

        ' Chart background WHITE (visible), plot area transparent, no
        ' borders. White background is essential so exported GIF renders
        ' properly on the userform image control (else it's fully
        ' transparent and appears blank).
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop

        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True: .AxisTitle.Text = "x [mm]": .AxisTitle.Font.Size = 12 * scaleK
            .TickLabels.Font.Size = 11 * scaleK
            .Format.Line.Weight = 0.75 * scaleK
            .HasMajorGridlines = True
            With .MajorGridlines.Format.Line
                .Visible = msoTrue: .ForeColor.RGB = RGB(190, 190, 190): .Weight = 0.75 * scaleK: .DashStyle = msoLineDash
            End With
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True: .AxisTitle.Text = "y [mm]": .AxisTitle.Font.Size = 12 * scaleK
            .TickLabels.Font.Size = 11 * scaleK
            .Format.Line.Weight = 0.75 * scaleK
            .HasMajorGridlines = True
            With .MajorGridlines.Format.Line
                .Visible = msoTrue: .ForeColor.RGB = RGB(190, 190, 190): .Weight = 0.75 * scaleK: .DashStyle = msoLineDash
            End With
        End With
        .Legend.Font.Size = 11 * scaleK

        With .SeriesCollection.NewSeries
            .Name = "Original": .XValues = wsPlot.Range("A2:A" & origRows): .Values = wsPlot.Range("B2:B" & origRows)
            .Format.Line.ForeColor.RGB = RGB(180, 180, 180): .Format.Line.Weight = 1 * scaleK: .MarkerStyle = xlMarkerStyleNone
        End With

        With .SeriesCollection.NewSeries
            .Name = "Deformed": .XValues = wsPlot.Range("C2:C" & defRow - 1): .Values = wsPlot.Range("D2:D" & defRow - 1)
            .Format.Line.ForeColor.RGB = RGB(220, 20, 60): .Format.Line.Weight = 2 * scaleK: .MarkerStyle = xlMarkerStyleNone
        End With
        
    End With

    ' =====================================================================
    '  EQUAL-SCALE AXES: 1 mm on x occupies the same length on screen as 1 mm
    '  on y, so the deformed shape is not distorted. The box covers the original
    '  arch, the deformed arch and the deformed backfill (the last two are drawn
    '  as freeform shapes, not as chart series, so Excel's autoscale ignores them).
    ' =====================================================================
    Dim eqXmin As Double, eqXmax As Double, eqYmin As Double, eqYmax As Double, eqGot As Boolean
    eqGot = False
    Call ExtendBoundsFromColumns(wsPlot, 1, 2, 2, origRows, eqXmin, eqXmax, eqYmin, eqYmax, eqGot)
    Call ExtendBoundsFromColumns(wsPlot, 3, 4, 2, defRow - 1, eqXmin, eqXmax, eqYmin, eqYmax, eqGot)
    Call ExtendBoundsFromColumns(wsPlot, 7, 8, 2, defBackRow - 1, eqXmin, eqXmax, eqYmin, eqYmax, eqGot)
    If eqGot Then Call SetEqualAspectAxes(ch.Chart, eqXmin, eqXmax, eqYmin, eqYmax)
    On Error Resume Next
    ch.Chart.Axes(xlCategory).TickLabels.NumberFormat = "0"
    ch.Chart.Axes(xlValue).TickLabels.NumberFormat = "0"
    On Error GoTo ErrorHandler
    DoEvents

    Call DrawBlocksOnChart(ch.Chart.Parent, wsPlot, 3, 4, 2, defRow - 1, RGB(242, 136, 89), RGB(220, 20, 60), 2 * scaleK, False)
        If defBackRow > 2 Then Call DrawBlocksOnChart(ch.Chart.Parent, wsPlot, 7, 8, 2, defBackRow - 1, RGB(0, 0, 0), RGB(100, 100, 100), 2 * scaleK, True)
    Call DrawHingesOnChart(ch.Chart.Parent, wsPlot, 5, 6, 2, 5, RGB(0, 0, 0), scaleK)
    Exit Sub

ErrorHandler:
    MsgBox "Error in PlotDeformedShape: " & Err.Description, vbCritical
End Sub

' ============================================================================
'  Helper functions for PlotDeformedShape
' ============================================================================

Private Function NudgeClearByRotation(poly() As Double, obstacles As Collection) As Double()
    Dim n As Long
    Dim pivX As Double
    Dim pivY As Double
    Dim i As Long

    n = UBound(poly, 1)

    pivX = poly(1, 1)
    pivY = poly(1, 2)

    For i = 2 To n
        If poly(i, 2) < pivY Then
            pivY = poly(i, 2)
            pivX = poly(i, 1)
        End If
    Next i

    Dim maxDeg As Double
    Dim stepDeg As Double
    maxDeg = 15#
    stepDeg = 0.25

    Dim s As Double
    Dim testSign As Variant
    Dim testPoly() As Double
    Dim ok As Boolean
    Dim ob As Variant
    Dim rad As Double

    For s = 0 To maxDeg Step stepDeg

        For Each testSign In Array(1, -1)

            If Not (s = 0 And testSign = -1) Then

                rad = CDbl(testSign) * s * 3.14159265358979 / 180#

                ReDim testPoly(1 To n, 1 To 2)

                For i = 1 To n
                    RotateAboutPt poly(i, 1), poly(i, 2), _
                                  pivX, pivY, rad, _
                                  testPoly(i, 1), testPoly(i, 2)
                Next i

                ok = True

                For Each ob In obstacles
                    If N_ConvexPolysOverlap(testPoly, ob) Then
                        ok = False
                        Exit For
                    End If
                Next ob

                If ok Then
                    NudgeClearByRotation = testPoly
                    Exit Function
                End If

            End If

        Next testSign

    Next s

    NudgeClearByRotation = testPoly

End Function

Private Sub TransformPoint(px As Double, py As Double, segIdx As Long, PLx As Double, PLy As Double, theta1 As Double, PRx As Double, PRy As Double, theta3 As Double, Rmid() As Double, TmX As Double, TmY As Double, ByRef npx As Double, ByRef npy As Double)
    Select Case segIdx
        Case 0: npx = px: npy = py
        Case 1: Call RotateAboutPt(px, py, PLx, PLy, theta1, npx, npy)
        Case 2: npx = Rmid(1, 1) * px + Rmid(1, 2) * py + TmX: npy = Rmid(2, 1) * px + Rmid(2, 2) * py + TmY
        Case 3: Call RotateAboutPt(px, py, PRx, PRy, theta3, npx, npy)
        Case 4: npx = px: npy = py
    End Select
End Sub

Private Sub RotateAboutPt(px As Double, py As Double, cx As Double, cy As Double, angle As Double, ByRef nx As Double, ByRef ny As Double)
    Dim c As Double, s As Double: c = Cos(angle): s = Sin(angle)
    nx = c * (px - cx) - s * (py - cy) + cx: ny = s * (px - cx) + c * (py - cy) + cy
End Sub

Private Function Atn2b(Y As Double, X As Double) As Double
    Const PI As Double = 3.14159265358979
    If X > 0 Then
        Atn2b = Atn(Y / X)
    ElseIf X < 0 Then
        If Y >= 0 Then
            Atn2b = Atn(Y / X) + PI
        Else
            Atn2b = Atn(Y / X) - PI
        End If
    ElseIf Y > 0 Then
        Atn2b = PI / 2
    ElseIf Y < 0 Then
        Atn2b = -PI / 2
    Else
        Atn2b = 0
    End If
End Function

Private Sub RigidTransform2(a1x As Double, a1y As Double, a2x As Double, a2y As Double, b1x As Double, b1y As Double, b2x As Double, b2y As Double, ByRef r() As Double, ByRef Tx As Double, ByRef Ty As Double)
    Dim ang As Double: ang = Atn2b(b2y - b1y, b2x - b1x) - Atn2b(a2y - a1y, a2x - a1x)
    Dim c As Double, s As Double: c = Cos(ang): s = Sin(ang)
    r(1, 1) = c: r(1, 2) = -s: r(2, 1) = s: r(2, 2) = c
    Tx = b1x - (c * a1x - s * a1y): Ty = b1y - (s * a1x + c * a1y)
End Sub

Private Function TransformPolyBySeg(ax As Variant, ay As Variant, segIdx As Long, _
                                     PLx As Double, PLy As Double, theta1 As Double, _
                                     PRx As Double, PRy As Double, theta3 As Double, _
                                     Rmid() As Double, TmX As Double, TmY As Double) As Double()
    Dim n As Long: n = UBound(ax)
    Dim result() As Double
    ReDim result(1 To n, 1 To 2)
    Dim c As Long, nx As Double, ny As Double
    For c = 1 To n
        Select Case segIdx
            Case 1: Call RotateAboutPt(CDbl(ax(c)), CDbl(ay(c)), PLx, PLy, theta1, nx, ny)
            Case 2: nx = Rmid(1, 1) * ax(c) + Rmid(1, 2) * ay(c) + TmX: ny = Rmid(2, 1) * ax(c) + Rmid(2, 2) * ay(c) + TmY
            Case 3: Call RotateAboutPt(CDbl(ax(c)), CDbl(ay(c)), PRx, PRy, theta3, nx, ny)
            Case Else: nx = ax(c): ny = ay(c)
        End Select
        result(c, 1) = nx: result(c, 2) = ny
    Next c
    TransformPolyBySeg = result
End Function

Private Function MaxPenetrationFull(bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, _
                                     NB As Long, segOf() As Long, _
                                     PLx As Double, PLy As Double, theta1 As Double, _
                                     PRx As Double, PRy As Double, theta3 As Double, _
                                     Rmid() As Double, TmX As Double, TmY As Double, _
                                     Optional includeBackfill As Boolean = True) As Double
    Dim worst As Double: worst = 0
    Dim allPolys As Collection: Set allPolys = New Collection
    Dim blkSeg() As Long: ReDim blkSeg(1 To NB)
    Dim i As Long

    For i = 1 To NB
        Dim polys As Collection: Set polys = New Collection
        blkSeg(i) = segOf(i)
        If Not IsEmpty(bx(i)) Then
            polys.Add TransformPolyBySeg(bx(i), by(i), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
        End If
        If includeBackfill Then
            If Not fx(i) Is Nothing Then
                Dim kf As Long
                For kf = 1 To fx(i).Count
                    polys.Add TransformPolyBySeg(fx(i)(kf), fy(i)(kf), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
                Next kf
            End If
        End If
        allPolys.Add polys
    Next i

    Dim iBlk As Long, jBlk As Long
    For iBlk = 1 To NB - 1
        For jBlk = iBlk + 1 To NB
            If blkSeg(iBlk) <> blkSeg(jBlk) Then
                Dim pa As Variant, pb As Variant, d As Double
                For Each pa In allPolys(iBlk)
                    For Each pb In allPolys(jBlk)
                        d = PolyPenetrationDepth(pa, pb)
                        If d > worst Then worst = d
                    Next pb
                Next pa
            End If
        Next jBlk
    Next iBlk

    MaxPenetrationFull = worst
End Function


' ============================================================================
'  Collapse-mechanism helpers (new)
' ============================================================================

Private Function NormAngle(ByVal a As Double) As Double
    Do While a > MECH_PI: a = a - 2 * MECH_PI: Loop
    Do While a <= -MECH_PI: a = a + 2 * MECH_PI: Loop
    NormAngle = a
End Function

' Four-bar linkage PL - H1 - H2 - PR.  The left link turns by th1 about PL (H1 -> H1n); H2n is the
' intersection of the circles (H1n, |H1H2|) and (PR, |H2PR|).  Of the two intersections the one that
' lies on the SAME side of the line H1n->PR as H2 does of H1->PR in the undeformed state is taken, i.e.
' the branch that is continuous with the undeformed configuration (no jump to the "mirrored" linkage).
Private Function FourBarSolve3(PLx As Double, PLy As Double, H1x As Double, H1y As Double, _
                               H2x As Double, H2y As Double, PRx As Double, PRy As Double, _
                               theta1 As Double, ByRef H1nx As Double, ByRef H1ny As Double, _
                               ByRef H2nx As Double, ByRef H2ny As Double) As Boolean
    Dim Lc As Double, Lr As Double, d As Double, A As Double, hSq As Double, h As Double
    Dim exx As Double, eyy As Double, mx As Double, my As Double
    Dim c1x As Double, c1y As Double, c2x As Double, c2y As Double
    Dim side0 As Double, s1 As Double

    FourBarSolve3 = False
    Call RotateAboutPt(H1x, H1y, PLx, PLy, theta1, H1nx, H1ny)
    Lc = Sqr((H2x - H1x) ^ 2 + (H2y - H1y) ^ 2)
    Lr = Sqr((H2x - PRx) ^ 2 + (H2y - PRy) ^ 2)
    d = Sqr((PRx - H1nx) ^ 2 + (PRy - H1ny) ^ 2)
    If d < 0.000000000001 Then Exit Function
    If d > Lc + Lr - 0.000000001 Or d < Abs(Lc - Lr) + 0.000000001 Then Exit Function

    A = (Lc ^ 2 - Lr ^ 2 + d ^ 2) / (2 * d)
    hSq = Lc ^ 2 - A ^ 2
    If hSq <= 0 Then Exit Function
    h = Sqr(hSq)
    exx = (PRx - H1nx) / d: eyy = (PRy - H1ny) / d
    mx = H1nx + A * exx: my = H1ny + A * eyy
    c1x = mx - h * eyy: c1y = my + h * exx
    c2x = mx + h * eyy: c2y = my - h * exx

    side0 = (PRx - H1x) * (H2y - H1y) - (PRy - H1y) * (H2x - H1x)
    s1 = (PRx - H1nx) * (c1y - H1ny) - (PRy - H1ny) * (c1x - H1nx)
    If side0 = 0 Then
        ' hinges aligned in the undeformed state: take the intersection nearest to H2
        If (c1x - H2x) ^ 2 + (c1y - H2y) ^ 2 <= (c2x - H2x) ^ 2 + (c2y - H2y) ^ 2 Then
            H2nx = c1x: H2ny = c1y
        Else
            H2nx = c2x: H2ny = c2y
        End If
    ElseIf s1 * side0 > 0 Then
        H2nx = c1x: H2ny = c1y
    Else
        H2nx = c2x: H2ny = c2y
    End If
    FourBarSolve3 = True
End Function

' Complete kinematic state of the mechanism for a given crank angle theta1:
' hinge positions, rotation of the middle link (th2) and of the right link (th3),
' and the rigid transform (Rmid, TmX, TmY) of the middle link.
Private Function MechanismState(PLx As Double, PLy As Double, H1x As Double, H1y As Double, _
                                H2x As Double, H2y As Double, PRx As Double, PRy As Double, _
                                theta1 As Double, ByRef H1nx As Double, ByRef H1ny As Double, _
                                ByRef H2nx As Double, ByRef H2ny As Double, _
                                ByRef th2 As Double, ByRef th3 As Double, _
                                Rmid() As Double, ByRef TmX As Double, ByRef TmY As Double) As Boolean
    MechanismState = False
    If Not FourBarSolve3(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy, theta1, H1nx, H1ny, H2nx, H2ny) Then Exit Function
    Call RigidTransform2(H1x, H1y, H2x, H2y, H1nx, H1ny, H2nx, H2ny, Rmid, TmX, TmY)
    th2 = NormAngle(Atn2b(H2ny - H1ny, H2nx - H1nx) - Atn2b(H2y - H1y, H2x - H1x))
    th3 = NormAngle(Atn2b(H2ny - PRy, H2nx - PRx) - Atn2b(H2y - PRy, H2x - PRx))
    MechanismState = True
End Function

' Largest displacement of any block vertex in the current mechanism state.
Private Function MaxVertexDisplacement(bx() As Variant, by() As Variant, NB As Long, segOf() As Long, _
                                       PLx As Double, PLy As Double, theta1 As Double, _
                                       PRx As Double, PRy As Double, theta3 As Double, _
                                       Rmid() As Double, TmX As Double, TmY As Double) As Double
    Dim i As Long, j As Long
    Dim px As Double, py As Double, npx As Double, npy As Double, d As Double, worst As Double
    worst = 0
    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            For j = 1 To UBound(bx(i))
                px = bx(i)(j): py = by(i)(j)
                Call TransformPoint(px, py, segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY, npx, npy)
                d = Sqr((npx - px) ^ 2 + (npy - py) ^ 2)
                If d > worst Then worst = d
            Next j
        End If
    Next i
    MaxVertexDisplacement = worst
End Function



Attribute VB_Name = "Module9"
Option Explicit

' ============================================================================
'  COLLAPSE MECHANISM
'
'  The hinges are NOT taken from a tolerance band on |e|/(t/2) alone (that picks
'  neighbouring interfaces on fine meshes, misses hinges when the solver result is
'  slightly off, and can return a hinge set that is not a mechanism at all).
'
'  Instead:
'   1. candidates = interfaces where the thrust line (nearly) touches intrados or
'      extrados, plus the local peaks of |e|/(t/2);
'   2. every set of 4 candidates is checked as a four-hinge mechanism
'      (ground - link 1 - link 2 - link 3 - ground). Each hinge sits on the edge
'      the thrust line touches and must OPEN on the opposite edge, which fixes the
'      sense of every relative rotation;
'   3. for each admissible set the kinematic load multiplier
'          lambda_kin = - (work of dead loads) / (work of live loads)
'      is computed. By the upper-bound theorem the true collapse mechanism is the
'      admissible one with the SMALLEST lambda_kin, and it must equal the static
'      lambda from the LP (a built-in check of the whole analysis).
'
'  An interface of zero length (t = 0) can only transmit a force through a point,
'  so it acts as a pin: it is always a candidate and may rotate either way.
' ============================================================================
Private Const HINGE_RATIO_HI As Double = 0.98     ' |e|/(t/2) counted as "touching"
Private Const HINGE_RATIO_PEAK As Double = 0.5    ' smallest local peak still considered
Private Const MAX_PER_CLUSTER As Long = 3         ' candidates kept per run of touching interfaces
Private Const MAX_CANDIDATES As Long = 30         ' C(30,4) = 27405 sets at most
Private Const MAX_FALLBACK As Long = 40           ' second pass: highest ratios only

' working data of the kinematic search (filled by FindCollapseMechanism)
Private kmHx() As Double, kmHy() As Double, kmSide() As Long
Private kmDx() As Double, kmDy() As Double, kmDm() As Double   ' prefix sums of dead loads (Fx, Fy, moment about origin)
Private kmLx() As Double, kmLy() As Double, kmLm() As Double   ' prefix sums of live loads
Private kmLiveTol As Double


Public Function GetInterfaceThickness(wsData As Worksheet, NI As Long) As Double()
    Dim tArr() As Double
    ReDim tArr(1 To NI)
    Dim i As Long, v As Variant, lastGood As Double

    For i = 1 To NI
        v = wsData.Cells(i + 1, 6).Value
        If IsNumeric(v) And Not IsEmpty(v) Then
            tArr(i) = CDbl(v)
            lastGood = tArr(i)
        ElseIf lastGood > 0 Then
            tArr(i) = lastGood
        End If
    Next i
    GetInterfaceThickness = tArr
End Function

' Interface end points p1 -> p2, oriented so that the normal (Fi) points from block i-1 to block i.
Public Sub GetInterfaceEdges(wsData As Worksheet, NB As Long, bx() As Variant, by() As Variant, _
                             ByRef p1x() As Double, ByRef p1y() As Double, ByRef p2x() As Double, ByRef p2y() As Double)
    Dim NI As Long, i As Long, fi As Double, v As Variant
    Dim curX() As Double, curY() As Double, prevX() As Double, prevY() As Double
    NI = NB + 1
    ReDim p1x(1 To NI): ReDim p1y(1 To NI): ReDim p2x(1 To NI): ReDim p2y(1 To NI)
    For i = 1 To NI
        fi = 0
        v = wsData.Cells(i + 1, 8).Value
        If IsNumeric(v) And Not IsEmpty(v) Then fi = CDbl(v)
        If i = 1 Then
            curX = bx(1): curY = by(1)
            FindEdgeByNormal curX, curY, fi, p1x(1), p1y(1), p2x(1), p2y(1)
        ElseIf i = NI Then
            curX = bx(NB): curY = by(NB)
            FindEdgeByNormal curX, curY, fi, p1x(NI), p1y(NI), p2x(NI), p2y(NI)
        Else
            prevX = bx(i - 1): prevY = by(i - 1)
            curX = bx(i): curY = by(i)
            FindSharedEdge prevX, prevY, curX, curY, p1x(i), p1y(i), p2x(i), p2y(i)
        End If
        OrientEdgeByFi p1x(i), p1y(i), p2x(i), p2y(i), fi
    Next i
End Sub

' Eccentricity e = -M/N of every interface (same sign convention as the thrust-line plot).
Public Function GetEccentricities(wsForces As Worksheet, NI As Long) As Double()
    Dim eVec() As Double, Nv() As Double, Mv() As Double, Nref As Double, i As Long
    ReDim eVec(1 To NI): ReDim Nv(1 To NI): ReDim Mv(1 To NI)
    For i = 1 To NI
        Nv(i) = wsForces.Cells(i + 1, 2).Value
        Mv(i) = wsForces.Cells(i + 1, 4).Value
        If Abs(Nv(i)) > Nref Then Nref = Abs(Nv(i))
    Next i
    If Nref = 0 Then Nref = 1
    For i = 1 To NI
        If Abs(Nv(i)) > 0.000001 * Nref Then eVec(i) = -1 * Mv(i) / Nv(i)
    Next i
    GetEccentricities = eVec
End Function

' Area centroid of a polygon.
Public Sub BlockCentroid(ax As Variant, ay As Variant, ByRef cx As Double, ByRef cy As Double)
    Dim j As Long, n As Long, x0 As Double, y0 As Double, x1 As Double, y1 As Double, cr As Double, a2 As Double
    n = UBound(ax)
    a2 = 0: cx = 0: cy = 0
    For j = 1 To n
        x0 = ax(j): y0 = ay(j)
        If j = n Then
            x1 = ax(1): y1 = ay(1)
        Else
            x1 = ax(j + 1): y1 = ay(j + 1)
        End If
        cr = x0 * y1 - x1 * y0
        a2 = a2 + cr
        cx = cx + (x0 + x1) * cr
        cy = cy + (y0 + y1) * cr
    Next j
    If Abs(a2) > 0.000000001 Then
        cx = cx / (3 * a2): cy = cy / (3 * a2)          ' a2 = 2*Area
    Else
        cx = 0: cy = 0
        For j = 1 To n
            cx = cx + ax(j): cy = cy + ay(j)
        Next j
        cx = cx / n: cy = cy / n
    End If
End Sub

' Convenience wrapper: reads geometry and the ForcesTable, then finds the mechanism.
Public Function FindCollapseMechanismForSheet(wsData As Worksheet, wsForces As Worksheet, _
        ByRef hIdx() As Long, ByRef hSide() As Long, ByRef hX() As Double, ByRef hY() As Double, _
        ByRef lamKin As Double, ByRef mechSign As Long) As Boolean
    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, NB As Long
    Dim p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double
    Dim tArr() As Double, eVec() As Double
    FindCollapseMechanismForSheet = False
    Call ReadCSVGeometry(wsData, bx, by, fx, fy, NB)
    If NB < 3 Then Exit Function
    Call GetInterfaceEdges(wsData, NB, bx, by, p1x, p1y, p2x, p2y)
    tArr = GetInterfaceThickness(wsData, NB + 1)
    eVec = GetEccentricities(wsForces, NB + 1)
    FindCollapseMechanismForSheet = FindCollapseMechanism(wsData, NB, bx, by, p1x, p1y, p2x, p2y, _
            tArr, eVec, hIdx, hSide, hX, hY, lamKin, mechSign)
End Function

' Finds the four-hinge collapse mechanism (see the header of this section).
'   hIdx(1..4)  interfaces of the hinges, left to right
'   hSide(k)    +1 / -1 = edge the hinge sits on (sign of e), 0 = pin (zero-length interface)
'   hX, hY      hinge points (on the intrados/extrados edge)
'   lamKin      kinematic load multiplier of the mechanism
'   mechSign    +1 / -1: sense of rotation of link 1 about hinge 1 (+1 = counter-clockwise)
' Returns False if no admissible four-hinge mechanism exists among the candidates.
Public Function FindCollapseMechanism(wsData As Worksheet, NB As Long, bx() As Variant, by() As Variant, _
        p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double, _
        tArr() As Double, eVec() As Double, _
        ByRef hIdx() As Long, ByRef hSide() As Long, ByRef hX() As Double, ByRef hY() As Double, _
        ByRef lamKin As Double, ByRef mechSign As Long) As Boolean
    Dim NI As Long, i As Long, k As Long
    Dim ux As Double, uy As Double, L As Double, tRef As Double
    Dim ratio() As Double, v As Variant, w As Double, cx As Double, cy As Double
    Dim FxL() As Double, FyL() As Double, liveSum As Double, archSize As Double

    FindCollapseMechanism = False
    NI = NB + 1
    ReDim hIdx(1 To 4): ReDim hSide(1 To 4): ReDim hX(1 To 4): ReDim hY(1 To 4)
    If NB < 3 Then Exit Function

    ' ---- hinge point, side and |e|/(t/2) of every interface ----
    ReDim kmHx(1 To NI): ReDim kmHy(1 To NI): ReDim kmSide(1 To NI): ReDim ratio(1 To NI)
    tRef = 0
    For i = 1 To NI
        If tArr(i) > tRef Then tRef = tArr(i)
    Next i
    For i = 1 To NI
        ux = p2x(i) - p1x(i): uy = p2y(i) - p1y(i)
        L = Sqr(ux * ux + uy * uy)
        If L > archSize Then archSize = L
        If L > 0 Then ux = ux / L: uy = uy / L
        If tArr(i) <= 0.000000001 * tRef Then
            kmSide(i) = 0: ratio(i) = 1                  ' zero-length interface: pin
        Else
            If eVec(i) >= 0 Then kmSide(i) = 1 Else kmSide(i) = -1
            ratio(i) = Abs(eVec(i)) / (tArr(i) / 2)
        End If
        kmHx(i) = (p1x(i) + p2x(i)) / 2 + kmSide(i) * tArr(i) / 2 * ux
        kmHy(i) = (p1y(i) + p2y(i)) / 2 + kmSide(i) * tArr(i) / 2 * uy
    Next i

    ' ---- loads at the block centroids, as prefix sums (work of a link in O(1)) ----
    Call GetBlockForces(wsData, NB, FxL, FyL)
    ReDim kmDx(0 To NB): ReDim kmDy(0 To NB): ReDim kmDm(0 To NB)
    ReDim kmLx(0 To NB): ReDim kmLy(0 To NB): ReDim kmLm(0 To NB)
    For i = 1 To NB
        If IsEmpty(bx(i)) Then Exit Function
        Call BlockCentroid(bx(i), by(i), cx, cy)
        w = 0
        v = wsData.Cells(i + 1, 10).Value                ' weight (column J), as in CreateFassSheet
        If IsNumeric(v) And Not IsEmpty(v) Then w = CDbl(v)
        kmDx(i) = kmDx(i - 1)
        kmDy(i) = kmDy(i - 1) - w
        kmDm(i) = kmDm(i - 1) + cx * (-w)
        kmLx(i) = kmLx(i - 1) + FxL(i)
        kmLy(i) = kmLy(i - 1) + FyL(i)
        kmLm(i) = kmLm(i - 1) + cx * FyL(i) - cy * FxL(i)
        liveSum = liveSum + Sqr(FxL(i) ^ 2 + FyL(i) ^ 2)
        If Abs(cx) > archSize Then archSize = Abs(cx)
        If Abs(cy) > archSize Then archSize = Abs(cy)
    Next i
    If liveSum = 0 Then Exit Function                    ' no live load: lambda is undefined
    kmLiveTol = 0.000000001 * liveSum * archSize

    ' ---- search: first the thrust-line candidates, then the highest ratios ----
    Dim cand() As Long, nCand As Long, pass As Long, found As Boolean
    For pass = 1 To 2
        If pass = 1 Then
            Call PickHingeCandidates(ratio, NI, cand, nCand)
        Else
            Call PickTopRatios(ratio, NI, MAX_FALLBACK, cand, nCand)
        End If
        found = SearchMechanism(cand, nCand, hIdx, lamKin, mechSign)
        If found Then Exit For
    Next pass
    If Not found Then Exit Function

    For k = 1 To 4
        hSide(k) = kmSide(hIdx(k))
        hX(k) = kmHx(hIdx(k)): hY(k) = kmHy(hIdx(k))
    Next k
    FindCollapseMechanism = True
End Function

' Runs of consecutive "touching" interfaces on the same edge keep their MAX_PER_CLUSTER
' best members; local peaks of the ratio are added; at most MAX_CANDIDATES, sorted by index.
Private Sub PickHingeCandidates(ratio() As Double, NI As Long, ByRef cand() As Long, ByRef nCand As Long)
    Dim isC() As Boolean, i As Long, a As Long, b As Long, k As Long, best As Long, j As Long
    Dim taken() As Boolean
    ReDim isC(1 To NI): ReDim taken(1 To NI)

    i = 1
    Do While i <= NI
        If ratio(i) >= HINGE_RATIO_HI Then
            a = i: b = i
            Do While b < NI
                If ratio(b + 1) < HINGE_RATIO_HI Or kmSide(b + 1) <> kmSide(a) Then Exit Do
                b = b + 1
            Loop
            For k = 1 To MAX_PER_CLUSTER
                best = 0
                For j = a To b
                    If Not taken(j) Then
                        If best = 0 Then
                            best = j
                        ElseIf ratio(j) > ratio(best) Then
                            best = j
                        End If
                    End If
                Next j
                If best = 0 Then Exit For
                taken(best) = True: isC(best) = True
            Next k
            i = b + 1
        Else
            i = i + 1
        End If
    Loop

    For i = 1 To NI
        If ratio(i) >= HINGE_RATIO_PEAK Then
            If i = 1 Then
                a = i
            Else
                a = i - 1
            End If
            If i = NI Then
                b = i
            Else
                b = i + 1
            End If
            If ratio(i) >= ratio(a) And ratio(i) >= ratio(b) Then isC(i) = True
        End If
    Next i

    Call CollectCandidates(isC, ratio, NI, MAX_CANDIDATES, cand, nCand)
End Sub

' The n interfaces with the largest ratio (fallback when the first pass finds nothing).
Private Sub PickTopRatios(ratio() As Double, NI As Long, ByVal n As Long, ByRef cand() As Long, ByRef nCand As Long)
    Dim isC() As Boolean, i As Long
    ReDim isC(1 To NI)
    For i = 1 To NI: isC(i) = True: Next i
    Call CollectCandidates(isC, ratio, NI, n, cand, nCand)
End Sub

' Flags -> list; keeps the maxN largest ratios; result sorted by interface index.
Private Sub CollectCandidates(isC() As Boolean, ratio() As Double, NI As Long, ByVal maxN As Long, _
                              ByRef cand() As Long, ByRef nCand As Long)
    Dim i As Long, j As Long, tmp As Long
    ReDim cand(1 To NI)
    nCand = 0
    For i = 1 To NI
        If isC(i) Then nCand = nCand + 1: cand(nCand) = i
    Next i
    If nCand > maxN Then
        For i = 1 To maxN                                ' partial selection sort by ratio
            For j = i + 1 To nCand
                If ratio(cand(j)) > ratio(cand(i)) Then tmp = cand(i): cand(i) = cand(j): cand(j) = tmp
            Next j
        Next i
        nCand = maxN
        For i = 1 To nCand - 1                           ' back to index order
            For j = i + 1 To nCand
                If cand(j) < cand(i) Then tmp = cand(i): cand(i) = cand(j): cand(j) = tmp
            Next j
        Next i
    End If
End Sub

' Tries every set of 4 candidates; keeps the admissible one with the smallest lambda_kin.
Private Function SearchMechanism(cand() As Long, nCand As Long, ByRef hIdx() As Long, _
                                 ByRef bestLam As Double, ByRef bestSign As Long) As Boolean
    Dim a As Long, b As Long, c As Long, d As Long, lam As Double, sg As Long, found As Boolean
    found = False
    For a = 1 To nCand - 3
        For b = a + 1 To nCand - 2
            For c = b + 1 To nCand - 1
                For d = c + 1 To nCand
                    If KinematicMultiplier(cand(a), cand(b), cand(c), cand(d), lam, sg) Then
                        If Not found Or lam < bestLam Then
                            found = True
                            bestLam = lam: bestSign = sg
                            hIdx(1) = cand(a): hIdx(2) = cand(b): hIdx(3) = cand(c): hIdx(4) = cand(d)
                        End If
                    End If
                Next d
            Next c
        Next b
    Next a
    SearchMechanism = found
End Function

' Infinitesimal four-hinge mechanism with hinges at interfaces h1 < h2 < h3 < h4:
'   link 1 = blocks h1..h2-1 turns by w1 about hinge 1 (fixed to the ground),
'   link 2 = blocks h2..h3-1, link 3 = blocks h3..h4-1 turns by w3 about hinge 4.
'   Closure (hinge 4 does not move): w1 (H2-H1) + w2 (H3-H2) + w3 (H4-H3) = 0, with w1 = 1.
' Relative rotations: phi1 = w1, phi2 = w2-w1, phi3 = w3-w2, phi4 = -w3. A hinge on side s
' opens on the opposite edge only if s*phi > 0 (all four with the same overall sign sg).
' Returns lambda_kin = -Wdead/Wlive for the admissible sense, False if not admissible.
Private Function KinematicMultiplier(ByVal h1 As Long, ByVal h2 As Long, ByVal h3 As Long, ByVal h4 As Long, _
                                     ByRef lam As Double, ByRef sg As Long) As Boolean
    Dim ax As Double, ay As Double, bxx As Double, byy As Double, rx As Double, ry As Double, det As Double
    Dim w2 As Double, w3 As Double, phi(1 To 4) As Double, hh(1 To 4) As Long
    Dim k As Long, sk As Long, phiMax As Double, nSided As Long
    Dim v0x As Double, v0y As Double, Wd As Double, Wl As Double

    KinematicMultiplier = False
    ax = kmHx(h3) - kmHx(h2): ay = kmHy(h3) - kmHy(h2)
    bxx = kmHx(h4) - kmHx(h3): byy = kmHy(h4) - kmHy(h3)
    rx = -(kmHx(h2) - kmHx(h1)): ry = -(kmHy(h2) - kmHy(h1))
    det = ax * byy - ay * bxx
    If Abs(det) <= 0.000000001 * Sqr((ax * ax + ay * ay) * (bxx * bxx + byy * byy)) Or det = 0 Then Exit Function
    w2 = (rx * byy - ry * bxx) / det
    w3 = (ax * ry - ay * rx) / det

    phi(1) = 1: phi(2) = w2 - 1: phi(3) = w3 - w2: phi(4) = -w3
    hh(1) = h1: hh(2) = h2: hh(3) = h3: hh(4) = h4
    For k = 1 To 4
        If Abs(phi(k)) > phiMax Then phiMax = Abs(phi(k))
    Next k
    sg = 0: nSided = 0
    For k = 1 To 4
        If Abs(phi(k)) <= 0.000001 * phiMax Then Exit Function    ' a "hinge" that does not turn
        If kmSide(hh(k)) <> 0 Then
            nSided = nSided + 1
            sk = Sgn(kmSide(hh(k)) * phi(k))
            If sg = 0 Then
                sg = sk
            ElseIf sk <> sg Then
                Exit Function                                       ' some hinge would close/crush
            End If
        End If
    Next k
    If sg = 0 Then sg = 1                                           ' only pins: either sense

    ' work done = sum over links of  F_sum . v0 + w * (Mc - O x F_sum)
    v0x = -(kmHy(h2) - kmHy(h1)): v0y = kmHx(h2) - kmHx(h1)        ' velocity of hinge 2 for w1 = 1
    Wd = LinkWork(kmDx, kmDy, kmDm, h1, h2 - 1, 0, 0, 1, kmHx(h1), kmHy(h1)) _
       + LinkWork(kmDx, kmDy, kmDm, h2, h3 - 1, v0x, v0y, w2, kmHx(h2), kmHy(h2)) _
       + LinkWork(kmDx, kmDy, kmDm, h3, h4 - 1, 0, 0, w3, kmHx(h4), kmHy(h4))
    Wl = LinkWork(kmLx, kmLy, kmLm, h1, h2 - 1, 0, 0, 1, kmHx(h1), kmHy(h1)) _
       + LinkWork(kmLx, kmLy, kmLm, h2, h3 - 1, v0x, v0y, w2, kmHx(h2), kmHy(h2)) _
       + LinkWork(kmLx, kmLy, kmLm, h3, h4 - 1, 0, 0, w3, kmHx(h4), kmHy(h4))
    Wd = sg * Wd: Wl = sg * Wl
    If Wl <= kmLiveTol And nSided = 0 Then sg = -sg: Wd = -Wd: Wl = -Wl
    If Wl <= kmLiveTol Then Exit Function               ' the live load cannot drive this mechanism
    lam = -Wd / Wl
    KinematicMultiplier = True
End Function

' Virtual work of the loads on blocks b1..b2 for the rigid velocity field v(P) = v0 + w k x (P - O).
Private Function LinkWork(sx() As Double, sy() As Double, sm() As Double, ByVal b1 As Long, ByVal b2 As Long, _
                          ByVal v0x As Double, ByVal v0y As Double, ByVal w As Double, _
                          ByVal ox As Double, ByVal oy As Double) As Double
    Dim fxs As Double, fys As Double, mcs As Double
    If b2 < b1 Then Exit Function
    fxs = sx(b2) - sx(b1 - 1): fys = sy(b2) - sy(b1 - 1): mcs = sm(b2) - sm(b1 - 1)
    LinkWork = fxs * v0x + fys * v0y + w * (mcs - (ox * fys - oy * fxs))
End Function

' Data check and repair, run BEFORE the analysis.
' The exporter writes the length (column F) and the normal angle Fi (column H) of every
' interface. For the internal interfaces both follow from the block geometry: the edge shared
' by block i-1 and block i, with the normal pointing from block i-1 to block i. Wrong values
' (e.g. Len = 0 / Fi = -90 when the exporter matched a zero-length segment instead of the real
' edge) are overwritten with the geometric ones and highlighted in yellow.
' Returns "" if nothing was changed, a description of the corrections otherwise, or a text
' starting with "ERROR:" if the data cannot be used (no shared edge, zero-length abutment).
Public Function RepairInterfaceData(wsData As Worksheet) As String
    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, NB As Long, i As Long
    Dim ax() As Double, ay() As Double, cx() As Double, cy() As Double
    Dim p1x As Double, p1y As Double, p2x As Double, p2y As Double
    Dim L As Double, nx As Double, ny As Double, fiGeo As Double, dFi As Double
    Dim c1x As Double, c1y As Double, c2x As Double, c2y As Double
    Dim vLen As Variant, vFi As Variant, oldLen As Double, oldFi As Double
    Dim msg As String, nFix As Long

    Call ReadCSVGeometry(wsData, bx, by, fx, fy, NB)
    If NB < 1 Then Exit Function
    For i = 1 To NB
        If IsEmpty(bx(i)) Then
            RepairInterfaceData = "ERROR: block " & i & " has no geometry in the data sheet."
            Exit Function
        End If
    Next i

    For i = 2 To NB
        ax = bx(i - 1): ay = by(i - 1): cx = bx(i): cy = by(i)
        If Not FindExactSharedEdge(ax, ay, cx, cy, p1x, p1y, p2x, p2y) Then
            RepairInterfaceData = "ERROR: blocks " & (i - 1) & " and " & i & " do not share an edge " & _
                                  "(interface " & i & "). Check the CAD geometry: neighbouring blocks must have " & _
                                  "two common vertices."
            Exit Function
        End If
        L = Sqr((p2x - p1x) ^ 2 + (p2y - p1y) ^ 2)
        nx = (p2y - p1y) / L: ny = -(p2x - p1x) / L
        Call BlockCentroid(bx(i - 1), by(i - 1), c1x, c1y)
        Call BlockCentroid(bx(i), by(i), c2x, c2y)
        If (c2x - c1x) * nx + (c2y - c1y) * ny < 0 Then nx = -nx: ny = -ny
        fiGeo = WorksheetFunction.Atan2(nx, ny) * 180 / WorksheetFunction.PI()

        vLen = wsData.Cells(i + 1, 6).Value: vFi = wsData.Cells(i + 1, 8).Value
        oldLen = 0: oldFi = 0
        If IsNumeric(vLen) And Not IsEmpty(vLen) Then oldLen = CDbl(vLen)
        If IsNumeric(vFi) And Not IsEmpty(vFi) Then oldFi = CDbl(vFi)
        dFi = Abs(oldFi - fiGeo)
        Do While dFi > 180: dFi = Abs(dFi - 360): Loop
        If Abs(oldLen - L) > 0.001 * L Or dFi > 0.01 Then
            nFix = nFix + 1
            If nFix <= 5 Then msg = msg & IIf(Len(msg) > 0, "; ", "") & "interface " & i & _
                                    " Len " & Format(oldLen, "0.0##") & " -> " & Format(L, "0.0##") & _
                                    ", Fi " & Format(oldFi, "0.0##") & " -> " & Format(fiGeo, "0.0##")
            wsData.Cells(i + 1, 6).Value = L
            wsData.Cells(i + 1, 8).Value = fiGeo
            wsData.Range(wsData.Cells(i + 1, 6), wsData.Cells(i + 1, 8)).Interior.Color = RGB(255, 235, 156)
        End If
    Next i

    ' Abutments (interfaces 1 and NB+1): the ground face of the end block is an edge that does
    ' not touch the edge shared with the neighbouring block (for a voussoir: the opposite face);
    ' if there are several, the lowest one. The exporter takes the edge nearest to the point
    ' clicked in CAD, which can be the intrados/extrados face instead (test 2: Len 204, Fi 31.6).
    Dim k As Long, eb As Long, nb2 As Long, cbx As Double, cby As Double
    If NB >= 2 Then
        For k = 1 To 2
            If k = 1 Then
                i = 1: eb = 1: nb2 = 2
            Else
                i = NB + 1: eb = NB: nb2 = NB - 1
            End If
            ax = bx(eb): ay = by(eb): cx = bx(nb2): cy = by(nb2)
            If FindAbutmentEdge(ax, ay, cx, cy, p1x, p1y, p2x, p2y) Then
                L = Sqr((p2x - p1x) ^ 2 + (p2y - p1y) ^ 2)
                nx = (p2y - p1y) / L: ny = -(p2x - p1x) / L
                Call BlockCentroid(bx(eb), by(eb), cbx, cby)
                ' interface 1: normal from the ground INTO block 1; interface NB+1: OUT of block NB
                If ((cbx - (p1x + p2x) / 2) * nx + (cby - (p1y + p2y) / 2) * ny < 0) = (k = 1) Then nx = -nx: ny = -ny
                fiGeo = WorksheetFunction.Atan2(nx, ny) * 180 / WorksheetFunction.PI()

                vLen = wsData.Cells(i + 1, 6).Value: vFi = wsData.Cells(i + 1, 8).Value
                oldLen = 0: oldFi = 0
                If IsNumeric(vLen) And Not IsEmpty(vLen) Then oldLen = CDbl(vLen)
                If IsNumeric(vFi) And Not IsEmpty(vFi) Then oldFi = CDbl(vFi)
                dFi = Abs(oldFi - fiGeo)
                Do While dFi > 180: dFi = Abs(dFi - 360): Loop
                If Abs(oldLen - L) > 0.001 * L Or dFi > 0.01 Then
                    nFix = nFix + 1
                    If nFix <= 5 Then msg = msg & IIf(Len(msg) > 0, "; ", "") & "abutment interface " & i & _
                                            " Len " & Format(oldLen, "0.0##") & " -> " & Format(L, "0.0##") & _
                                            ", Fi " & Format(oldFi, "0.0##") & " -> " & Format(fiGeo, "0.0##")
                    wsData.Cells(i + 1, 6).Value = L
                    wsData.Cells(i + 1, 8).Value = fiGeo
                    wsData.Range(wsData.Cells(i + 1, 6), wsData.Cells(i + 1, 8)).Interior.Color = RGB(255, 235, 156)
                End If
            End If
        Next k
    End If

    For i = 1 To NB + 1 Step NB
        vLen = wsData.Cells(i + 1, 6).Value
        If Not IsNumeric(vLen) Or IsEmpty(vLen) Then vLen = 0
        If CDbl(vLen) <= 0 Then
            RepairInterfaceData = "ERROR: the abutment interface " & i & " has zero length in column F. " & _
                                  "Re-export the CSV and click exactly on the abutment edge."
            Exit Function
        End If
    Next i

    If nFix > 5 Then msg = msg & " and " & (nFix - 5) & " more"
    If nFix > 0 Then RepairInterfaceData = "corrected " & nFix & " interface(s) from the block geometry: " & msg
End Function

' Ground face of an end block (ax, ay) whose neighbour is (bx, by): among the edges that touch
' neither end of the shared edge, the one with the lowest mid-height. False if there is none.
Private Function FindAbutmentEdge(ax() As Double, ay() As Double, bx() As Double, by() As Double, _
                                  ByRef p1x As Double, ByRef p1y As Double, _
                                  ByRef p2x As Double, ByRef p2y As Double) As Boolean
    Dim s1x As Double, s1y As Double, s2x As Double, s2y As Double
    Dim i As Long, i2 As Long, n As Long, tol As Double, sz As Double, midY As Double, bestY As Double
    If Not FindExactSharedEdge(ax, ay, bx, by, s1x, s1y, s2x, s2y) Then Exit Function
    n = UBound(ax)
    For i = 1 To n
        i2 = (i Mod n) + 1
        sz = sz + Sqr((ax(i2) - ax(i)) ^ 2 + (ay(i2) - ay(i)) ^ 2)
    Next i
    tol = 0.001 * sz / n
    FindAbutmentEdge = False
    For i = 1 To n
        i2 = (i Mod n) + 1
        If Sqr((ax(i2) - ax(i)) ^ 2 + (ay(i2) - ay(i)) ^ 2) > 10 * tol Then
            If Not (NearPt(ax(i), ay(i), s1x, s1y, tol) Or NearPt(ax(i), ay(i), s2x, s2y, tol) Or _
                    NearPt(ax(i2), ay(i2), s1x, s1y, tol) Or NearPt(ax(i2), ay(i2), s2x, s2y, tol)) Then
                midY = (ay(i) + ay(i2)) / 2
                If Not FindAbutmentEdge Or midY < bestY Then
                    FindAbutmentEdge = True
                    bestY = midY
                    p1x = ax(i): p1y = ay(i): p2x = ax(i2): p2y = ay(i2)
                End If
            End If
        End If
    Next i
End Function

Private Function NearPt(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                        ByVal tol As Double) As Boolean
    NearPt = (Abs(x1 - x2) < tol And Abs(y1 - y2) < tol)
End Function

' Edge common to two polygons (both end points equal within a tolerance relative to the block
' size). Zero-length edges (duplicated vertices) are ignored; the longest common edge is taken.
Private Function FindExactSharedEdge(ax() As Double, ay() As Double, bx() As Double, by() As Double, _
                                     ByRef p1x As Double, ByRef p1y As Double, _
                                     ByRef p2x As Double, ByRef p2y As Double) As Boolean
    Dim i As Long, j As Long, i2 As Long, j2 As Long, nA As Long, nB As Long
    Dim tol As Double, L As Double, bestL As Double, sz As Double
    nA = UBound(ax): nB = UBound(bx)
    For i = 1 To nA
        i2 = (i Mod nA) + 1
        sz = sz + Sqr((ax(i2) - ax(i)) ^ 2 + (ay(i2) - ay(i)) ^ 2)
    Next i
    tol = 0.001 * sz / nA
    If tol < 0.0000001 Then tol = 0.0000001
    bestL = 0
    For i = 1 To nA
        i2 = (i Mod nA) + 1
        L = Sqr((ax(i2) - ax(i)) ^ 2 + (ay(i2) - ay(i)) ^ 2)
        If L > 10 * tol Then
            For j = 1 To nB
                j2 = (j Mod nB) + 1
                If (Abs(ax(i) - bx(j)) < tol And Abs(ay(i) - by(j)) < tol And _
                    Abs(ax(i2) - bx(j2)) < tol And Abs(ay(i2) - by(j2)) < tol) Or _
                   (Abs(ax(i) - bx(j2)) < tol And Abs(ay(i) - by(j2)) < tol And _
                    Abs(ax(i2) - bx(j)) < tol And Abs(ay(i2) - by(j)) < tol) Then
                    If L > bestL Then
                        bestL = L
                        p1x = ax(i): p1y = ay(i): p2x = ax(i2): p2y = ay(i2)
                    End If
                End If
            Next j
        End If
    Next i
    FindExactSharedEdge = (bestL > 0)
End Function

' True if at least one block carries a live load (Fx_lam / Fy_lam, columns L and M).
Public Function HasLiveLoad(wsData As Worksheet, NB As Long) As Boolean
    Dim FxL() As Double, FyL() As Double, i As Long
    If NB < 1 Then Exit Function
    Call GetBlockForces(wsData, NB, FxL, FyL)
    For i = 1 To NB
        If FxL(i) <> 0 Or FyL(i) <> 0 Then HasLiveLoad = True: Exit Function
    Next i
End Function

Public Sub ReadCSVGeometry(ws As Worksheet, ByRef bx() As Variant, ByRef by() As Variant, ByRef fx() As Variant, ByRef fy() As Variant, ByRef NB As Long)
    Dim lastRow As Long, i As Long, s As String, bType As String, bIdx As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    NB = 0
    
    For i = 2 To lastRow
        s = CStr(ws.Cells(i, 1).Value)
        If Left(s, 5) = "Block" Then
            bIdx = CLng(Trim(Replace(s, "Block", "")))
            If bIdx > NB Then NB = bIdx
        End If
    Next i
    
    If NB = 0 Then
        ReDim bx(1 To 1): ReDim by(1 To 1): ReDim fx(1 To 1): ReDim fy(1 To 1)
        Exit Sub
    End If
    
    ReDim bx(1 To NB)
    ReDim by(1 To NB)
    ReDim fx(1 To NB)
    ReDim fy(1 To NB)
    
    For i = 1 To NB
        Set fx(i) = New Collection
        Set fy(i) = New Collection
    Next i
    
    Dim curBlock As Long, isBackfill As Boolean
    Dim curX As New Collection, curY As New Collection
    
    For i = 2 To lastRow + 1
        s = CStr(ws.Cells(i, 1).Value)
        
        If s <> "" Or i > lastRow Then
            Dim px As Double, py As Double
            
            If i <= lastRow Then
                If IsNumeric(ws.Cells(i, 2).Value) Then px = CDbl(ws.Cells(i, 2).Value) Else px = 0
                If IsNumeric(ws.Cells(i, 3).Value) Then py = CDbl(ws.Cells(i, 3).Value) Else py = 0
            End If
            
            If s <> "" Then
                If Left(s, 5) = "Block" Then
                    bType = "Block"
                    bIdx = CLng(Trim(Replace(s, "Block", "")))
                End If
                If Left(s, 8) = "Backfill" Then
                    bType = "Backfill"
                    bIdx = CLng(Trim(Replace(s, "Backfill", "")))
                End If
            End If
            
            If curBlock = 0 And bIdx > 0 Then
                curBlock = bIdx
                isBackfill = (bType = "Backfill")
            ElseIf (bIdx <> curBlock Or bType <> IIf(isBackfill, "Backfill", "Block") Or i > lastRow) Then
                FlushPoly curX, curY, isBackfill, curBlock, bx, by, fx, fy
                curBlock = bIdx
                isBackfill = (bType = "Backfill")
            End If
            
            If i <= lastRow And s <> "" Then
                curX.Add px
                curY.Add py
                If curX.Count > 1 Then
                    If Abs(px - curX(1)) < 0.001 And Abs(py - curY(1)) < 0.001 Then
                        FlushPoly curX, curY, isBackfill, curBlock, bx, by, fx, fy
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub FlushPoly(ByRef cx As Collection, ByRef cy As Collection, isBackfill As Boolean, cBlock As Long, ByRef bx() As Variant, ByRef by() As Variant, ByRef fx() As Variant, ByRef fy() As Variant)
    If cx.Count = 0 Or cBlock = 0 Then Exit Sub
    Dim n As Long: n = cx.Count
    If n > 1 Then
        If Abs(cx(n) - cx(1)) < 0.001 And Abs(cy(n) - cy(1)) < 0.001 Then
            n = n - 1
        End If
    End If
    Dim ax() As Double, ay() As Double, j As Long
    ReDim ax(1 To n)
    ReDim ay(1 To n)
    
    For j = 1 To n
        ax(j) = cx(j)
        ay(j) = cy(j)
    Next j
    
    If isBackfill Then
        fx(cBlock).Add ax
        fy(cBlock).Add ay
    Else
        bx(cBlock) = ax
        by(cBlock) = ay
    End If
    
    Set cx = New Collection
    Set cy = New Collection
End Sub

Public Sub DrawBlocksOnChart(chObj As ChartObject, ws As Worksheet, xCol As Long, yCol As Long, startRow As Long, endRow As Long, fillColor As Long, lineRgb As Long, lineWt As Double, isHatch As Boolean)
    Dim ch As Chart
    Set ch = chObj.Chart
    
    Dim axX As Axis, axY As Axis, pArea As PlotArea
    Set axX = ch.Axes(xlCategory)
    Set axY = ch.Axes(xlValue)
    Set pArea = ch.PlotArea
    
    Dim scaleX As Double, scaleY As Double
    scaleX = pArea.InsideWidth / (axX.MaximumScale - axX.MinimumScale)
    scaleY = pArea.InsideHeight / (axY.MaximumScale - axY.MinimumScale)
    
    Dim builder As FreeformBuilder, shp As Shape
    Dim r As Long, isNew As Boolean
    isNew = True
    
    For r = startRow To endRow
        If IsError(ws.Cells(r, xCol).Value) Or IsEmpty(ws.Cells(r, xCol).Value) Then
            If Not builder Is Nothing Then
                Set shp = builder.ConvertToShape
                shp.Fill.Visible = msoTrue
                                If isHatch Then
                    ' FIX: patterned fills in Excel don't natively support the
                    ' .Transparency property, so instead of a solid diagonal
                    ' hatch we use a much sparser, lighter grey pattern and
                    ' also lighten the line/border. This keeps the "backfill
                    ' looks different from blocks" cue without hiding the
                    ' chart title, legend or axis labels behind it.
                    shp.Fill.Patterned msoPatternWideUpwardDiagonal
                    shp.Fill.ForeColor.RGB = RGB(140, 140, 140)
                    shp.Fill.BackColor.RGB = RGB(245, 245, 245)
                    ' Try to apply transparency (works on newer Excel, silently
                    ' skipped on older versions where patterns are opaque)
                    On Error Resume Next
                    shp.Fill.Transparency = 0.6
                    On Error GoTo 0
                Else
                    shp.Fill.Solid
                    shp.Fill.ForeColor.RGB = fillColor
                    shp.Fill.Transparency = 0.5
                End If
                shp.Line.Visible = msoTrue
                shp.Line.ForeColor.RGB = lineRgb
                shp.Line.Weight = lineWt
                shp.ZOrder msoSendToBack
                Set builder = Nothing
            End If
            isNew = True
        Else
            Dim px As Double, py As Double
            px = pArea.InsideLeft + (ws.Cells(r, xCol).Value - axX.MinimumScale) * scaleX
            py = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yCol).Value) * scaleY
            
            If isNew Then
                Set builder = ch.Shapes.BuildFreeform(msoEditingAuto, px, py)
                isNew = False
            Else
                builder.AddNodes msoSegmentLine, msoEditingAuto, px, py
            End If
        End If
    Next r
End Sub

Public Sub DrawHingesOnChart(chObj As ChartObject, ws As Worksheet, xCol As Long, yCol As Long, startRow As Long, endRow As Long, borderRgb As Long, Optional ByVal scaleK As Double = 1)
    Dim ch As Chart, pArea As PlotArea, axX As Axis, axY As Axis
    Set ch = chObj.Chart
    Set pArea = ch.PlotArea
    Set axX = ch.Axes(xlCategory)
    Set axY = ch.Axes(xlValue)
    
    Dim scaleX As Double
    scaleX = pArea.InsideWidth / (axX.MaximumScale - axX.MinimumScale)
    Dim scaleY As Double
    scaleY = pArea.InsideHeight / (axY.MaximumScale - axY.MinimumScale)
    
    Dim r As Long, radius As Double
    radius = 5 * scaleK
    
    For r = startRow To endRow
        If IsNumeric(ws.Cells(r, xCol).Value) And Not IsEmpty(ws.Cells(r, xCol).Value) Then
            Dim px As Double, py As Double
            px = pArea.InsideLeft + (ws.Cells(r, xCol).Value - axX.MinimumScale) * scaleX
            py = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yCol).Value) * scaleY
            
            Dim shp As Shape
            Set shp = ch.Shapes.AddShape(msoShapeOval, px - radius, py - radius, radius * 2, radius * 2)
            shp.Fill.Visible = msoTrue
            shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = borderRgb
            shp.Line.Weight = 1.5 * scaleK
        End If
    Next r
End Sub

Public Sub FindEdgeByNormal(ax() As Double, ay() As Double, targetFiDeg As Double, ByRef p1x As Double, ByRef p1y As Double, ByRef p2x As Double, ByRef p2y As Double)
    Dim i As Long, n As Long
    n = UBound(ax)
    Dim bestDiff As Double
    bestDiff = 1000
    
    For i = 1 To n
        Dim px1 As Double, py1 As Double, px2 As Double, py2 As Double
        px1 = ax(i)
        py1 = ay(i)
        
        If i = n Then
            px2 = ax(1)
            py2 = ay(1)
        Else
            px2 = ax(i + 1)
            py2 = ay(i + 1)
        End If
        
        Dim ang As Double, fi As Double
        ang = WorksheetFunction.Atan2(px2 - px1, py2 - py1) * 180 / WorksheetFunction.PI()
        fi = ang - 90
        Do While fi > 180
            fi = fi - 360
        Loop
        Do While fi <= -180
            fi = fi + 360
        Loop
        
        Dim diff1 As Double, diff2 As Double
        diff1 = Abs(fi - targetFiDeg)
        Do While diff1 > 180
            diff1 = Abs(diff1 - 360)
        Loop
        
        diff2 = Abs((fi + 180) - targetFiDeg)
        Do While diff2 > 180
            diff2 = Abs(diff2 - 360)
        Loop
        
        If diff1 < bestDiff Then
            bestDiff = diff1
            p1x = px1
            p1y = py1
            p2x = px2
            p2y = py2
        End If
        
        If diff2 < bestDiff Then
            bestDiff = diff2
            p1x = px2
            p1y = py2
            p2x = px1
            p2y = py1
        End If
    Next i
End Sub

Public Sub FindSharedEdge(ax() As Double, ay() As Double, bx() As Double, by() As Double, ByRef p1x As Double, ByRef p1y As Double, ByRef p2x As Double, ByRef p2y As Double)
    Dim i As Long, j As Long
    Dim nA As Long, NB As Long
    nA = UBound(ax)
    NB = UBound(bx)

    ' --- FIX: tol used to be a fixed 0.1 (drawing units, e.g. mm). That's
    ' fine at ~30mm block scale (~0.3% of an edge) but becomes ~2%+ of an
    ' edge at ~5mm scale, which can match the WRONG pair of edges as the
    ' "shared interface" between two blocks (or miss the true match and
    ' fall through to the nearest-edge fallback below). Make it relative
    ' to the average edge length of the two polygons being compared instead,
    ' so behaviour doesn't change with how big/small you draw the arch.
    Dim avgEdge As Double, edgeCount As Long, i2 As Long, j2 As Long
    avgEdge = 0: edgeCount = 0
    For i = 1 To nA
        i2 = (i Mod nA) + 1
        avgEdge = avgEdge + Sqr((ax(i2) - ax(i)) ^ 2 + (ay(i2) - ay(i)) ^ 2)
        edgeCount = edgeCount + 1
    Next i
    For j = 1 To NB
        j2 = (j Mod NB) + 1
        avgEdge = avgEdge + Sqr((bx(j2) - bx(j)) ^ 2 + (by(j2) - by(j)) ^ 2)
        edgeCount = edgeCount + 1
    Next j
    If edgeCount > 0 Then avgEdge = avgEdge / edgeCount

    Dim tol As Double
    tol = 0.01 * avgEdge                      ' 1% of the average edge length
    If tol < 0.0000001 Then tol = 0.0000001   ' guard against degenerate/zero-size input
    
    Dim a1x As Double, a1y As Double, a2x As Double, a2y As Double
    Dim b1x As Double, b1y As Double, b2x As Double, b2y As Double

    For i = 1 To nA
        a1x = ax(i)
        a1y = ay(i)
        If i = nA Then
            a2x = ax(1)
            a2y = ay(1)
        Else
            a2x = ax(i + 1)
            a2y = ay(i + 1)
        End If
        
        For j = 1 To NB
            b1x = bx(j)
            b1y = by(j)
            If j = NB Then
                b2x = bx(1)
                b2y = by(1)
            Else
                b2x = bx(j + 1)
                b2y = by(j + 1)
            End If
            
            If (Abs(a1x - b1x) < tol And Abs(a1y - b1y) < tol And Abs(a2x - b2x) < tol And Abs(a2y - b2y) < tol) Or _
               (Abs(a1x - b2x) < tol And Abs(a1y - b2y) < tol And Abs(a2x - b1x) < tol And Abs(a2y - b1y) < tol) Then
                p1x = a1x
                p1y = a1y
                p2x = a2x
                p2y = a2y
                Exit Sub
            End If
        Next j
    Next i

    Dim bestDist As Double: bestDist = 1E+30
    Dim bp1x As Double, bp1y As Double, bp2x As Double, bp2y As Double
    Dim mAx As Double, mAy As Double, mBx As Double, mBy As Double, d As Double

    For i = 1 To nA
        a1x = ax(i)
        a1y = ay(i)
        If i = nA Then
            a2x = ax(1)
            a2y = ay(1)
        Else
            a2x = ax(i + 1)
            a2y = ay(i + 1)
        End If
        mAx = (a1x + a2x) / 2
        mAy = (a1y + a2y) / 2
        
        For j = 1 To NB
            b1x = bx(j)
            b1y = by(j)
            If j = NB Then
                b2x = bx(1)
                b2y = by(1)
            Else
                b2x = bx(j + 1)
                b2y = by(j + 1)
            End If
            mBx = (b1x + b2x) / 2
            mBy = (b1y + b2y) / 2
            
            d = (mAx - mBx) ^ 2 + (mAy - mBy) ^ 2
            If d < bestDist Then
                bestDist = d
                bp1x = a1x: bp1y = a1y: bp2x = a2x: bp2y = a2y
            End If
        Next j
    Next i

    p1x = bp1x: p1y = bp1y: p2x = bp2x: p2y = bp2y
End Sub

Public Sub OrientEdgeByFi(ByRef p1x As Double, ByRef p1y As Double, ByRef p2x As Double, ByRef p2y As Double, targetFiDeg As Double)
    Dim ang As Double, fi As Double
    ang = WorksheetFunction.Atan2(p2x - p1x, p2y - p1y) * 180 / WorksheetFunction.PI()
    fi = ang - 90
    Do While fi > 180
        fi = fi - 360
    Loop
    Do While fi <= -180
        fi = fi + 360
    Loop
    
    Dim diffSame As Double, diffFlip As Double
    diffSame = Abs(fi - targetFiDeg)
    Do While diffSame > 180
        diffSame = Abs(diffSame - 360)
    Loop
    diffFlip = Abs((fi + 180) - targetFiDeg)
    Do While diffFlip > 180
        diffFlip = Abs(diffFlip - 360)
    Loop
    
    If diffFlip < diffSame Then
        Dim Tx As Double, Ty As Double
        Tx = p1x: Ty = p1y
        p1x = p2x: p1y = p2y
        p2x = Tx: p2y = Ty
    End If
End Sub

Public Function N_ConvexPolysOverlap(polyA As Variant, polyB As Variant) As Boolean
    N_ConvexPolysOverlap = True
    Dim shrink As Double
    shrink = 0.94
    
    Dim cax As Double, cay As Double, cbx As Double, cby As Double
    Dim i As Long, p As Long, poly As Long
    Dim nA As Long, NB As Long
    
    nA = UBound(polyA, 1)
    NB = UBound(polyB, 1)
    
    For i = 1 To nA
        cax = cax + polyA(i, 1)
        cay = cay + polyA(i, 2)
    Next i
    cax = cax / nA
    cay = cay / nA
    
    For i = 1 To NB
        cbx = cbx + polyB(i, 1)
        cby = cby + polyB(i, 2)
    Next i
    cbx = cbx / NB
    cby = cby / NB
    
    Dim A() As Double, b() As Double
    ReDim A(1 To nA, 1 To 2)
    ReDim b(1 To NB, 1 To 2)
    
    For i = 1 To nA
        A(i, 1) = cax + (polyA(i, 1) - cax) * shrink
        A(i, 2) = cay + (polyA(i, 2) - cay) * shrink
    Next i
    
    For i = 1 To NB
        b(i, 1) = cbx + (polyB(i, 1) - cbx) * shrink
        b(i, 2) = cby + (polyB(i, 2) - cby) * shrink
    Next i
    
    Dim edgeX As Double, edgeY As Double, axisX As Double, axisY As Double
    Dim minA As Double, maxA As Double, minB As Double, maxB As Double, proj As Double
    
    For poly = 1 To 2
        Dim nEdges As Long
        If poly = 1 Then
            nEdges = nA
        Else
            nEdges = NB
        End If
        
        For p = 1 To nEdges
            Dim p2 As Long
            p2 = (p Mod nEdges) + 1
            
            If poly = 1 Then
                edgeX = A(p2, 1) - A(p, 1)
                edgeY = A(p2, 2) - A(p, 2)
            Else
                edgeX = b(p2, 1) - b(p, 1)
                edgeY = b(p2, 2) - b(p, 2)
            End If
            
            axisX = -edgeY
            axisY = edgeX
            
            If Not (axisX = 0 And axisY = 0) Then
                minA = 1E+30
                maxA = -1E+30
                
                For i = 1 To nA
                    proj = A(i, 1) * axisX + A(i, 2) * axisY
                    If proj < minA Then minA = proj
                    If proj > maxA Then maxA = proj
                Next i
                
                minB = 1E+30
                maxB = -1E+30
                
                For i = 1 To NB
                    proj = b(i, 1) * axisX + b(i, 2) * axisY
                    If proj < minB Then minB = proj
                    If proj > maxB Then maxB = proj
                Next i
                
                If maxA < minB Or maxB < minA Then
                    N_ConvexPolysOverlap = False
                    Exit Function
                End If
            End If
        Next p
    Next poly
End Function

Public Sub GetBlockForces(wsData As Worksheet, NB As Long, ByRef FxArr() As Double, ByRef FyArr() As Double)
    ReDim FxArr(1 To NB)
    ReDim FyArr(1 To NB)
    Dim lastRow As Long, r As Long, s As String, bIdx As Long, v As Variant
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    ' FIX: read the load from EVERY row of the block (same rule as CreateFassSheet). The LISP
    ' export sometimes writes Fx_lam/Fy_lam on the closing vertex row instead of the first one;
    ' reading only the first row drew no arrow and gave the mechanism check a zero live load.
    For r = 2 To lastRow
        s = Trim(CStr(wsData.Cells(r, 1).Value))
        If Left(s, 5) = "Block" Then
            bIdx = CLng(Trim(Replace(s, "Block", "")))
            If bIdx >= 1 And bIdx <= NB Then
                v = wsData.Cells(r, 12).Value
                If Len(CStr(v)) > 0 And IsNumeric(v) Then FxArr(bIdx) = CDbl(v)
                v = wsData.Cells(r, 13).Value
                If Len(CStr(v)) > 0 And IsNumeric(v) Then FyArr(bIdx) = CDbl(v)
            End If
        End If
    Next r
End Sub

Public Sub DrawForceArrows(chObj As ChartObject, ws As Worksheet, xStartCol As Long, yStartCol As Long, xEndCol As Long, yEndCol As Long, startRow As Long, endRow As Long, lineRgb As Long, Optional ByVal scaleK As Double = 1)
    Dim ch As Chart, pArea As PlotArea, axX As Axis, axY As Axis
    Set ch = chObj.Chart
    Set pArea = ch.PlotArea
    Set axX = ch.Axes(xlCategory)
    Set axY = ch.Axes(xlValue)
    
    Dim scaleX As Double, scaleY As Double
    scaleX = pArea.InsideWidth / (axX.MaximumScale - axX.MinimumScale)
    scaleY = pArea.InsideHeight / (axY.MaximumScale - axY.MinimumScale)
    
    Dim r As Long
    For r = startRow To endRow
        If IsNumeric(ws.Cells(r, xStartCol).Value) And Not IsEmpty(ws.Cells(r, xStartCol).Value) Then
            Dim sx As Double, sy As Double, ex As Double, ey As Double
            sx = pArea.InsideLeft + (ws.Cells(r, xStartCol).Value - axX.MinimumScale) * scaleX
            sy = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yStartCol).Value) * scaleY
            ex = pArea.InsideLeft + (ws.Cells(r, xEndCol).Value - axX.MinimumScale) * scaleX
            ey = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yEndCol).Value) * scaleY
            
            Dim shp As Shape
            Set shp = ch.Shapes.AddLine(sx, sy, ex, ey)
            shp.Line.ForeColor.RGB = lineRgb
            shp.Line.Weight = 2.25 * scaleK
            shp.Line.EndArrowheadStyle = msoArrowheadTriangle
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide

            Dim dotR As Double: dotR = 3.5 * scaleK
            Dim dot As Shape
            Set dot = ch.Shapes.AddShape(msoShapeOval, sx - dotR, sy - dotR, dotR * 2, dotR * 2)
            dot.Fill.Visible = msoTrue
            dot.Fill.ForeColor.RGB = lineRgb
            dot.Line.Visible = msoFalse

            Dim lbl As Shape
            Set lbl = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, sx + 4 * scaleK, sy - 16 * scaleK, 16 * scaleK, 14 * scaleK)
            lbl.TextFrame.Characters.Text = "P"
            lbl.TextFrame.Characters.Font.Bold = True
            lbl.TextFrame.Characters.Font.Size = 10 * scaleK
            lbl.TextFrame.Characters.Font.Color = lineRgb
            lbl.Fill.Visible = msoFalse
            lbl.Line.Visible = msoFalse
            lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0
            lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        End If
    Next r
End Sub
Sub DumpHingeDiagnostics()
    Dim wbUser As Workbook: Set wbUser = ActiveWorkbook
    Dim wsData As Worksheet: Set wsData = wbUser.ActiveSheet
    Dim wsForces As Worksheet: Set wsForces = wbUser.Sheets("ForcesTable")

    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, NB As Long
    Call ReadCSVGeometry(wsData, bx, by, fx, fy, NB)
    Dim NI As Long: NI = NB + 1

    Dim tArr() As Double: tArr = GetInterfaceThickness(wsData, NI)
    Dim Nvec() As Double, Mvec() As Double: ReDim Nvec(1 To NI): ReDim Mvec(1 To NI)
    Dim i As Long
    For i = 1 To NI
        Nvec(i) = wsForces.Cells(i + 1, 2).Value
        Mvec(i) = wsForces.Cells(i + 1, 4).Value
    Next i

    Dim FiArr() As Double: ReDim FiArr(1 To NI)
    For i = 1 To NI: FiArr(i) = wsData.Cells(i + 1, 8).Value: Next i

    Dim p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double
    ReDim p1x(1 To NI), p1y(1 To NI), p2x(1 To NI), p2y(1 To NI)
    Dim curX() As Double, curY() As Double, prevX() As Double, prevY() As Double
    For i = 1 To NI
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

    On Error Resume Next: wbUser.Sheets("Diag").Delete: On Error GoTo 0
    Dim ws As Worksheet: Set ws = wbUser.Sheets.Add: ws.Name = "Diag"
    ws.Range("A1:J1").Value = Array("i", "N", "M", "t", "e", "ratio", "p1x", "p1y", "p2x", "p2y")
    For i = 1 To NI
        Dim eV As Double: eV = 0
        If Abs(Nvec(i)) > 0.000000001 Then eV = -1 * Mvec(i) / Nvec(i)
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = Nvec(i)
        ws.Cells(i + 1, 3).Value = Mvec(i)
        ws.Cells(i + 1, 4).Value = tArr(i)
        ws.Cells(i + 1, 5).Value = eV
        If tArr(i) > 0 Then ws.Cells(i + 1, 6).Value = Abs(eV) / (tArr(i) / 2)
        ws.Cells(i + 1, 7).Value = p1x(i): ws.Cells(i + 1, 8).Value = p1y(i)
        ws.Cells(i + 1, 9).Value = p2x(i): ws.Cells(i + 1, 10).Value = p2y(i)
    Next i
    ws.Columns.AutoFit
    MsgBox "Diagnostics written to 'Diag' sheet."
End Sub

' Returns 0 if the two convex polygons are separated (or exactly touching),
' otherwise the magnitude of the Minimum Translation Vector needed to
' separate them - i.e. HOW BADLY they overlap, not just whether they do.
' Used to rank candidate crank angles by severity of interpenetration,
' instead of only knowing pass/fail (see PlotDeformedShape's search loop).
Public Function PolyPenetrationDepth(polyA As Variant, polyB As Variant) As Double
    Dim shrink As Double: shrink = 0.98
    Dim nA As Long, NB As Long, i As Long
    nA = UBound(polyA, 1): NB = UBound(polyB, 1)

    Dim cax As Double, cay As Double, cbx As Double, cby As Double
    For i = 1 To nA: cax = cax + polyA(i, 1): cay = cay + polyA(i, 2): Next i
    cax = cax / nA: cay = cay / nA
    For i = 1 To NB: cbx = cbx + polyB(i, 1): cby = cby + polyB(i, 2): Next i
    cbx = cbx / NB: cby = cby / NB

    Dim A() As Double, b() As Double
    ReDim A(1 To nA, 1 To 2): ReDim b(1 To NB, 1 To 2)
    For i = 1 To nA
        A(i, 1) = cax + (polyA(i, 1) - cax) * shrink
        A(i, 2) = cay + (polyA(i, 2) - cay) * shrink
    Next i
    For i = 1 To NB
        b(i, 1) = cbx + (polyB(i, 1) - cbx) * shrink
        b(i, 2) = cby + (polyB(i, 2) - cby) * shrink
    Next i

    Dim minOverlap As Double: minOverlap = 1E+30
    Dim poly As Long, p As Long, p2 As Long, nEdges As Long
    Dim edgeX As Double, edgeY As Double, axisX As Double, axisY As Double, axisLen As Double
    Dim minA As Double, maxA As Double, minB As Double, maxB As Double, proj As Double, ov As Double

    For poly = 1 To 2
        If poly = 1 Then nEdges = nA Else nEdges = NB
        For p = 1 To nEdges
            p2 = (p Mod nEdges) + 1
            If poly = 1 Then
                edgeX = A(p2, 1) - A(p, 1): edgeY = A(p2, 2) - A(p, 2)
            Else
                edgeX = b(p2, 1) - b(p, 1): edgeY = b(p2, 2) - b(p, 2)
            End If
            axisX = -edgeY: axisY = edgeX
            axisLen = Sqr(axisX ^ 2 + axisY ^ 2)
            If axisLen > 0.000000001 Then
                axisX = axisX / axisLen: axisY = axisY / axisLen

                minA = 1E+30: maxA = -1E+30
                For i = 1 To nA
                    proj = A(i, 1) * axisX + A(i, 2) * axisY
                    If proj < minA Then minA = proj
                    If proj > maxA Then maxA = proj
                Next i
                minB = 1E+30: maxB = -1E+30
                For i = 1 To NB
                    proj = b(i, 1) * axisX + b(i, 2) * axisY
                    If proj < minB Then minB = proj
                    If proj > maxB Then maxB = proj
                Next i

                If maxA < minB Or maxB < minA Then
                    PolyPenetrationDepth = 0   ' separating axis exists -> no overlap
                    Exit Function
                End If
                ov = WorksheetFunction.Min(maxA, maxB) - WorksheetFunction.mAx(minA, minB)
                If ov < minOverlap Then minOverlap = ov
            End If
        Next p
    Next poly

    PolyPenetrationDepth = minOverlap
End Function

Attribute VB_Name = "Module6"
Option Explicit

Sub RunSolverAndBackSubstitute()
    Dim wsOpt As Worksheet
    Dim wsGass As Worksheet, wsFass As Worksheet
    Dim wsForces As Worksheet
    Dim wbUser As Workbook
    Dim lastRow As Long
    Dim i As Long, j As Long

    On Error GoTo ErrorHandler
    Set wbUser = ActiveWorkbook
    Set wsOpt = wbUser.Sheets("optimization")
    wsOpt.Activate

    lastRow = wsOpt.Cells(wsOpt.Rows.Count, "A").End(xlUp).row

    Dim t0 As Double, solveStatus As Long
    t0 = Timer
    solveStatus = SolveWithOpenSolver(wsOpt, lastRow)
    gResult.secLP = Timer - t0
    If solveStatus <> OS_OPTIMAL Then
        Err.Raise vbObjectError + 516, "RunSolverAndBackSubstitute", OpenSolverStatusText(solveStatus)
    End If

    Dim N_last As Double, V_last As Double, M_last As Double, lam As Double
    ' The sheet holds the scaled variables (see Module 5)
    N_last = wsOpt.Range("B2").Value / wsOpt.Range("B4").Value
    V_last = wsOpt.Range("C2").Value / wsOpt.Range("C4").Value
    M_last = wsOpt.Range("D2").Value / wsOpt.Range("D4").Value
    lam = wsOpt.Range("E2").Value / wsOpt.Range("E4").Value

    ' The solution must satisfy the constraints as they stand on the sheet
    Dim resMax As Double, r As Long, lhsRhs As Variant
    lhsRhs = wsOpt.Range("G5:I" & lastRow).Value
    For r = 1 To UBound(lhsRhs, 1)
        If lhsRhs(r, 1) - lhsRhs(r, 3) > resMax Then resMax = lhsRhs(r, 1) - lhsRhs(r, 3)
    Next r
    gResult.residual = resMax
    If resMax > 0.000001 Then
        Err.Raise vbObjectError + 520, "RunSolverAndBackSubstitute", _
                  "The solver result violates the constraints on sheet 'optimization' by " & Format(resMax, "0.0E+00") & "."
    End If

    Set wsGass = wbUser.Sheets("Gass")
    Set wsFass = wbUser.Sheets("Fass")

    Dim gassRows As Long, gassCols As Long
    gassRows = wsGass.Cells(wsGass.Rows.Count, "A").End(xlUp).row
    gassCols = wsGass.Cells(1, wsGass.Columns.Count).End(xlToLeft).Column

    Dim G1 As Variant, G2 As Variant
    G1 = wsGass.Range(wsGass.Cells(1, 1), wsGass.Cells(gassRows, gassCols - 3)).Value
    G2 = wsGass.Range(wsGass.Cells(1, gassCols - 2), wsGass.Cells(gassRows, gassCols)).Value

    Dim fassRows As Long
    fassRows = wsFass.Cells(wsFass.Rows.Count, "A").End(xlUp).row
    Dim Fw As Variant, Flam As Variant
    Fw = wsFass.Range("A2:A" & fassRows).Value
    Flam = wsFass.Range("B2:B" & fassRows).Value

    Dim x_last(1 To 3, 1 To 1) As Double
    x_last(1, 1) = N_last: x_last(2, 1) = V_last: x_last(3, 1) = M_last

    Dim G2x() As Double
    ReDim G2x(1 To UBound(G2, 1), 1 To 1)
    Dim sumG2 As Double
    For i = 1 To UBound(G2, 1)
        sumG2 = 0
        For j = 1 To 3
            sumG2 = sumG2 + G2(i, j) * x_last(j, 1)
        Next j
        G2x(i, 1) = sumG2
    Next i

    Dim rhs() As Double
    ReDim rhs(1 To UBound(Fw, 1), 1 To 1)
    For i = 1 To UBound(Fw, 1)
        rhs(i, 1) = Fw(i, 1) + lam * Flam(i, 1) - G2x(i, 1)
    Next i

    Dim xFullVar As Variant
    xFullVar = SolveLinearSystem(G1, rhs)

    Dim xFull() As Double
    ReDim xFull(1 To UBound(xFullVar, 1), 1 To 1)
    For i = 1 To UBound(xFullVar, 1)
        xFull(i, 1) = xFullVar(i, 1)
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    wbUser.Sheets("ForcesTable").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsForces = wbUser.Sheets.Add
    wsForces.Name = "ForcesTable"

    ' Columns: A interface, B N, C V, D M, E e, F |V/N| (minimum friction coefficient required
    ' by this solution, since sliding is not checked), G status, H lambda_u, I |F_ref|, J P_u.
    wsForces.Cells(1, 1).Value = "Interface"
    wsForces.Cells(1, 2).Value = "N [" & UnitForce() & "]"
    wsForces.Cells(1, 3).Value = "V [" & UnitForce() & "]"
    wsForces.Cells(1, 4).Value = "M [" & UnitMoment() & "]"
    wsForces.Cells(1, 5).Value = "e=M/N [" & UnitLength() & "]"
    wsForces.Cells(1, 6).Value = "|V/N| [-]"
    wsForces.Cells(1, 7).Value = "Status"
    wsForces.Range("A1:G1").Font.Bold = True
    wsForces.Range("A1:G1").Interior.Color = RGB(200, 200, 200)

    Dim NI As Long
    NI = UBound(xFull, 1) / 3 + 1

    Dim Nref As Double: Nref = 0
    For i = 1 To NI - 1
        If Abs(xFull((i - 1) * 3 + 1, 1)) > Nref Then Nref = Abs(xFull((i - 1) * 3 + 1, 1))
    Next i
    If Abs(N_last) > Nref Then Nref = Abs(N_last)
    If Nref = 0 Then Nref = 1

    Dim row As Long: row = 2
    Dim Nval As Double, Vval As Double, Mval As Double, eVal As Double, muMax As Double
    For i = 1 To NI - 1
        Nval = xFull((i - 1) * 3 + 1, 1)
        Vval = xFull((i - 1) * 3 + 2, 1)
        Mval = xFull((i - 1) * 3 + 3, 1)
        eVal = 0
        If Abs(Nval) > 0.000001 * Nref Then eVal = Mval / Nval
        wsForces.Cells(row, 1).Value = i
        wsForces.Cells(row, 2).Value = Nval
        wsForces.Cells(row, 3).Value = Vval
        wsForces.Cells(row, 4).Value = Mval
        wsForces.Cells(row, 5).Value = eVal
        If Abs(Nval) > 0.000001 * Nref Then
            wsForces.Cells(row, 6).Value = Abs(Vval / Nval)
            If Abs(Vval / Nval) > muMax Then muMax = Abs(Vval / Nval)
        End If
        wsForces.Cells(row, 7).Value = "Computed"
        row = row + 1
    Next i

    eVal = 0
    If Abs(N_last) > 0.000001 * Nref Then eVal = M_last / N_last
    wsForces.Cells(row, 1).Value = NI
    wsForces.Cells(row, 2).Value = N_last
    wsForces.Cells(row, 3).Value = V_last
    wsForces.Cells(row, 4).Value = M_last
    wsForces.Cells(row, 5).Value = eVal
    If Abs(N_last) > 0.000001 * Nref Then
        wsForces.Cells(row, 6).Value = Abs(V_last / N_last)
        If Abs(V_last / N_last) > muMax Then muMax = Abs(V_last / N_last)
    End If
    wsForces.Cells(row, 7).Value = "Solver (last)"

    ' F_ref is the live-load vector of the data sheet and |F_ref| the magnitude of its resultant.
    ' lambda_u is dimensionless and P_u = lambda_u * |F_ref| is the collapse force.
    Dim Fref As Double, fh As Double, fv As Double
    For i = 1 To UBound(Flam, 1) Step 3
        fh = fh + Flam(i, 1): fv = fv + Flam(i + 1, 1)
    Next i
    Fref = Sqr(fh * fh + fv * fv)
    wsForces.Cells(1, 8).Value = "lambda_u [-]"
    wsForces.Cells(2, 8).Value = lam
    wsForces.Cells(1, 9).Value = "|F_ref| [" & UnitForce() & "]"
    wsForces.Cells(2, 9).Value = Fref
    wsForces.Cells(1, 10).Value = "P_u [" & UnitForce() & "]"
    wsForces.Cells(2, 10).Value = lam * Fref
    wsForces.Range("H1:J1").Font.Bold = True
    wsForces.Range("H2:J2").NumberFormat = "0.0000"
    wsForces.Columns("A:J").AutoFit

    gResult.lam = lam
    gResult.Fref = Fref
    gResult.Pu = lam * Fref
    gResult.muMax = muMax
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 517, "RunSolverAndBackSubstitute", Err.Description
End Sub

' ============================================================
' Solves the LP on sheet "optimization" through the public API of
' OpenSolver. The calls are late bound (Application.Run), so the add-in
' has no VBA reference to OpenSolver or to the Excel Solver add-in, and
' needs neither "Trust access to the VBA project object model" nor a
' fixed installation path. Returns the OpenSolver status code.
' ============================================================
Private Function SolveWithOpenSolver(wsOpt As Worksheet, lastRow As Long) As Long
    Const OS As String = "OpenSolver.xlam!"
    Const OBJ_MAXIMISE As Long = 1, REL_LE As Long = 1
    wsOpt.Activate
    Application.Run OS & "SetChosenSolver", "CBC", wsOpt     ' linear solver, whatever the last choice was
    Application.Run OS & "SetObjectiveFunctionCell", wsOpt.Range("$F$3"), wsOpt
    Application.Run OS & "SetObjectiveSense", OBJ_MAXIMISE, wsOpt
    Application.Run OS & "SetDecisionVariables", wsOpt.Range("$B$2:$E$2"), wsOpt
    Application.Run OS & "SetNumConstraints", 1&, wsOpt
    Application.Run OS & "SetConstraintLhs", 1&, wsOpt.Range("$G$5:$G$" & lastRow), wsOpt
    Application.Run OS & "SetConstraintRel", 1&, REL_LE, wsOpt
    Application.Run OS & "SetConstraintRhs", 1&, wsOpt.Range("$I$5:$I$" & lastRow), "", wsOpt
    Application.Run OS & "SetNonNegativity", False, wsOpt
    Application.Run OS & "SetLinearityCheck", False, wsOpt
    Application.Run OS & "SetShowSolverProgress", False, wsOpt
    Application.Run OS & "SetTolerance", 0#, wsOpt
    Application.Run OS & "SetPrecision", 0.0000001, wsOpt
    ' RunOpenSolver(SolveRelaxation, MinimiseUserInteraction, LinearityOffset, sheet)
    SolveWithOpenSolver = Application.Run(OS & "RunOpenSolver", False, True, 10.423, wsOpt)
End Function

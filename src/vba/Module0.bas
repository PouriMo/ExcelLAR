Attribute VB_Name = "Module0"
Option Explicit

' ============================================================================
'  Start-up checks and analysis driver.
'
'  The add-in has no VBA reference to OpenSolver or to the Excel Solver add-in and does not
'  use the VBA project object model, so it needs neither a fixed Office path nor the option
'  "Trust access to the VBA project object model".
'
'  RunAnalysis is the only entry point of the calculation. The window and scripted runs
'  (RunAnalysisSilent, used for the tests and the run times of the paper) share it.
' ============================================================================
Public Const EXCELLAR_VERSION As String = "1.1"
Public Const OS_OPTIMAL As Long = 0

Public Type AnalysisResult
    lam As Double          ' dimensionless collapse multiplier lambda_u
    Fref As Double         ' |F_ref|: magnitude of the resultant of the reference forces
    Pu As Double           ' collapse force lambda_u * |F_ref|
    lamKin As Double       ' kinematic multiplier of the mechanism
    nHinges As Long
    hinges As String
    muMax As Double        ' largest |V/N|
    residual As Double     ' largest constraint violation of the LP solution (scaled rows)
    secInput As Double
    secAssembly As Double  ' Modules 1 to 5 and back-substitution
    secLP As Double
    secMech As Double
    secPlots As Double
    secTotal As Double
End Type

Public gResult As AnalysisResult
Public gSilent As Boolean          ' True in scripted runs: no message boxes
Public gLogTarget As Object        ' the window while it runs an analysis
Private gUnitL As String, gUnitF As String

' ---- units: labels only, the analysis works in the units of the data sheet ----
Public Function UnitLength() As String
    UnitLength = IIf(Len(gUnitL) > 0, gUnitL, "mm")
End Function

Public Function UnitForce() As String
    UnitForce = IIf(Len(gUnitF) > 0, gUnitF, "N")
End Function

Public Function UnitMoment() As String
    UnitMoment = UnitForce() & ChrW(183) & UnitLength()
End Function

Public Sub SetUnits(ByVal lengthUnit As String, ByVal forceUnit As String)
    gUnitL = Trim(lengthUnit): gUnitF = Trim(forceUnit)
End Sub

' Unit written between brackets in a header cell of the data sheet, e.g. "X [mm]" gives "mm".
Public Function UnitFromHeader(ws As Worksheet, ByVal col As Long, ByVal dflt As String) As String
    Dim s As String, a As Long, b As Long
    UnitFromHeader = dflt
    s = CStr(ws.Cells(1, col).Value)
    a = InStr(s, "["): b = InStr(s, "]")
    If a > 0 And b > a + 1 Then UnitFromHeader = Mid$(s, a + 1, b - a - 1)
End Function

' ---- dependencies ----
Public Function OpenSolverLoaded() As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks("OpenSolver.xlam")
    OpenSolverLoaded = Not wb Is Nothing
End Function

' "" if the analysis can run, otherwise what the user has to do.
Public Function CheckDependencies() As String
    If InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) = 0 Then
        CheckDependencies = "ExcelLAR runs on Excel for Windows only (the plots use Windows libraries)."
    ElseIf Not OpenSolverLoaded() Then
        CheckDependencies = "OpenSolver is not loaded." & vbCrLf & vbCrLf & _
            "1. Download OpenSolver from https://opensolver.org." & vbCrLf & _
            "2. Right-click the ZIP file, choose Properties, tick Unblock, click OK, then extract it." & vbCrLf & _
            "3. In Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse, " & _
            "select OpenSolver.xlam and make sure it is ticked." & vbCrLf & vbCrLf & _
            "No macro or Trust Center setting has to be lowered."
    End If
End Function

Public Function OpenSolverStatusText(ByVal status As Long) As String
    Select Case status
        Case 5: OpenSolverStatusText = "The linear program is infeasible: no thrust line in equilibrium with these " & _
                                       "loads fits inside the arch. Check the geometry and the block weights."
        Case 4: OpenSolverStatusText = "The linear program is unbounded: the live load does not drive any mechanism. " & _
                                       "Check its position and direction."
        Case Else: OpenSolverStatusText = "OpenSolver did not return an optimal solution (status " & status & ")."
    End Select
End Function

Private Sub LogRun(ByVal msg As String)
    If Not gLogTarget Is Nothing Then gLogTarget.LogMsg msg
End Sub

' ---- analysis driver ----
' Runs the whole analysis on a data sheet of wbUser. Returns "" on success, otherwise the error.
Public Function RunAnalysis(wbUser As Workbook, ByVal dataSheet As String, ByVal doPlots As Boolean) As String
    Dim wsData As Worksheet, wsForces As Worksheet, msg As String
    Dim tStart As Double, t0 As Double, nb As Long
    Dim cbx() As Variant, cby() As Variant, cfx() As Variant, cfy() As Variant

    On Error GoTo Failed
    Dim blank As AnalysisResult
    gResult = blank
    tStart = Timer: t0 = tStart
    Set wsData = wbUser.Sheets(dataSheet)
    wsData.Activate
    LogRun "=== Starting full analysis on '" & wbUser.Name & "' ==="

    msg = CheckDependencies()
    If Len(msg) > 0 Then RunAnalysis = msg: LogRun "ERROR: " & Replace(msg, vbCrLf, " "): Exit Function

    ' Input checks: interface data against the block geometry, stereotomy, live load
    msg = RepairInterfaceData(wsData)
    If Left$(msg, 6) = "ERROR:" Then
        RunAnalysis = Mid$(msg, 8): LogRun msg: Exit Function
    ElseIf Len(msg) > 0 Then
        LogRun "Data check: " & msg & " (cells highlighted in yellow)."
    End If
    msg = CheckStereotomy(wsData)
    If Left$(msg, 6) = "ERROR:" Then
        RunAnalysis = Mid$(msg, 8): LogRun msg: Exit Function
    ElseIf Len(msg) > 0 Then
        LogRun "Stereotomy check: " & msg
    End If
    Call ReadCSVGeometry(wsData, cbx, cby, cfx, cfy, nb)
    If Not HasLiveLoad(wsData, nb) Then
        RunAnalysis = "The data sheet has no live load: all Fx_lam / Fy_lam values (columns L and M) are zero." & vbCrLf & _
                      "Without a load to scale, the collapse multiplier is unbounded." & vbCrLf & _
                      "Re-export the CSV and enter the loaded block number(s) and the load vector."
        LogRun "ERROR: no live load in the data sheet (Fx_lam / Fy_lam are all zero)."
        Exit Function
    End If
    gResult.secInput = Timer - t0

    t0 = Timer
    LogRun "Step 1/6: Generating Gass_in matrix..."
    Generate_Gass_in
    wsData.Activate
    LogRun "Step 2/6: Generating Gass matrix..."
    GenerateGassMatrix
    wsData.Activate
    LogRun "Step 3/6: Creating Fass sheet..."
    CreateFassSheet
    LogRun "Step 4/6: Creating Ain matrix and Bin vector..."
    CreateAinSheet
    CreateBinSheet
    LogRun "Step 5/6: Building optimization sheet..."
    CreateOptimizationSheet
    LogRun "Step 6/6: Running OpenSolver and back-substitution..."
    RunSolverAndBackSubstitute
    gResult.secAssembly = Timer - t0 - gResult.secLP

    Set wsForces = wbUser.Sheets("ForcesTable")
    If gResult.lam < 0 Then LogRun "Warning: lambda_u is negative: the arch does not stand under the dead loads alone."

    ' Collapse mechanism: its kinematic multiplier must equal the static multiplier
    t0 = Timer
    Dim hIdx() As Long, hSide() As Long, hX() As Double, hY() As Double, mechSign As Long
    If FindCollapseMechanismForSheet(wsData, wsForces, hIdx, hSide, hX, hY, gResult.lamKin, mechSign) Then
        gResult.nHinges = 4
        gResult.hinges = hIdx(1) & ", " & hIdx(2) & ", " & hIdx(3) & ", " & hIdx(4)
        LogRun "Collapse mechanism: hinges at interfaces " & gResult.hinges & _
               " (kinematic multiplier = " & Format(gResult.lamKin, "0.0000") & ")."
        If gResult.lam <> 0 Then
            If Abs(gResult.lamKin - gResult.lam) > 0.01 * Abs(gResult.lam) Then
                LogRun "Warning: the kinematic multiplier differs from the static one by " & _
                       Format(Abs(gResult.lamKin - gResult.lam) / Abs(gResult.lam) * 100, "0.0") & _
                       "% - the solver result may be inaccurate."
            End If
        End If
    Else
        LogRun "Warning: no admissible four-hinge collapse mechanism could be formed from the thrust line."
    End If
    gResult.secMech = Timer - t0

    LogRun "=== Analysis complete: lambda_u = " & Format(gResult.lam, "0.0000") & ", P_u = " & _
           Format(gResult.Pu, "0.0000") & " " & UnitForce() & ", max |V/N| = " & Format(gResult.muMax, "0.000") & " ==="

    t0 = Timer
    If doPlots Then
        LogRun "Generating thrust line plot..."
        PlotThrustLine wsData
        LogRun "Generating deformed shape plot..."
        PlotDeformedShape wsData
        LogRun "Plots created on 'ThrustPlot' and 'DeformedPlot' sheets."
    End If
    gResult.secPlots = Timer - t0
    gResult.secTotal = Timer - tStart

    Call WriteRunInfo(wbUser, wsData, nb)
    Exit Function

Failed:
    RunAnalysis = Err.Description
    LogRun "ERROR: " & Err.Description
End Function

' Sheet RunInfo: what was run, with which software, the results and the time of each stage.
Private Sub WriteRunInfo(wbUser As Workbook, wsData As Worksheet, ByVal nb As Long)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = wbUser.Worksheets("RunInfo")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wbUser.Worksheets.Add(After:=wbUser.Sheets(wbUser.Sheets.Count))
        ws.Name = "RunInfo"
    Else
        ws.Cells.Clear
    End If
    Dim v As Variant, item As Variant
    v = Array( _
        Array("ExcelLAR version", EXCELLAR_VERSION), _
        Array("Excel version", Application.Version & " build " & Application.Build), _
        Array("Operating system", Application.OperatingSystem), _
        Array("Data sheet", wsData.Name), _
        Array("Blocks", nb), _
        Array("Interfaces", nb + 1), _
        Array("LP constraints", 2 * (nb + 1)), _
        Array("lambda_u [-]", gResult.lam), _
        Array("|F_ref| [" & UnitForce() & "]", gResult.Fref), _
        Array("P_u [" & UnitForce() & "]", gResult.Pu), _
        Array("Kinematic multiplier [-]", gResult.lamKin), _
        Array("Hinge interfaces", gResult.hinges), _
        Array("Max |V/N| [-]", gResult.muMax), _
        Array("LP residual [-]", gResult.residual), _
        Array("Input checks [s]", gResult.secInput), _
        Array("Assembly + reduction [s]", gResult.secAssembly), _
        Array("LP, OpenSolver [s]", gResult.secLP), _
        Array("Mechanism search [s]", gResult.secMech), _
        Array("Plots [s]", gResult.secPlots), _
        Array("Total [s]", gResult.secTotal))
    r = 1
    For Each item In v
        ws.Cells(r, 1).Value = item(0)
        If VarType(item(1)) = vbString Then ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = item(1)
        r = r + 1
    Next item
    ws.Columns("A:B").AutoFit
End Sub

' Entry point for scripted runs (Application.Run from a test script): no window, no dialogs.
' Returns "" on success, otherwise the error text; the results are on the sheet RunInfo.
Public Function RunAnalysisSilent(ByVal dataSheet As String, Optional ByVal doPlots As Boolean = False, _
                                  Optional ByVal lengthUnit As String = "", Optional ByVal forceUnit As String = "") As String
    Dim ws As Worksheet, prevScreen As Boolean
    gSilent = True
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = ActiveWorkbook.Sheets(dataSheet)
    Call SetUnits(IIf(Len(lengthUnit) > 0, lengthUnit, UnitFromHeader(ws, 2, "mm")), _
                  IIf(Len(forceUnit) > 0, forceUnit, UnitFromHeader(ws, 10, "N")))
    RunAnalysisSilent = RunAnalysis(ActiveWorkbook, dataSheet, doPlots)
    Application.ScreenUpdating = prevScreen
    gSilent = False
End Function

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelLAR 
   Caption         =   "Excel Limit Analysis Toolkit"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "ExcelLAR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelLAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' PNG export quality selectors, created at run time next to the Export buttons
Private cmbQualityThrust As MSForms.ComboBox
Private cmbQualityDeformed As MSForms.ComboBox

' ==================== Helper: Get the user's workbook ====================
Private Function UserWorkbook() As Workbook
    If ActiveWorkbook Is Nothing Then
        Set UserWorkbook = Workbooks.Add
    ElseIf ActiveWorkbook Is ThisWorkbook Then
        Set UserWorkbook = Workbooks.Add
    Else
        Set UserWorkbook = ActiveWorkbook
    End If
End Function

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

' ==================== Form Initialize ====================
Private Sub UserForm_Initialize()
    Call RefreshSheetList
    txtLambda.Value = "-"
    txtHinges.Value = "-"
    btnShowThrust.Enabled = False
    btnShowDeformed.Enabled = False

    ' Configure text log to handle its own scrolling properly
    txtLog.Value = ""
    txtLog.MultiLine = True
    txtLog.ScrollBars = 2 ' fmScrollBarsVertical
    txtLog.Enabled = True ' Required to allow scrolling interaction
    txtLog.Locked = True  ' Prevents user from typing in the log

    ' Make image panel backgrounds transparent
    imgThrust.BackStyle = 0 ' fmBackStyleTransparent
    imgDeformed.BackStyle = 0 ' fmBackStyleTransparent

    Set cmbQualityThrust = AddQualityPicker(btnExportThrust)
    Set cmbQualityDeformed = AddQualityPicker(btnExportDeformed)

    Call LogMsg("Ready. Load or select a CSV data sheet to begin.")
End Sub

' Adds "PNG quality: [list]" on the same row as an Export button, left of the buttons there.
Private Function AddQualityPicker(btn As MSForms.CommandButton) As MSForms.ComboBox
    Dim host As Object, ctl As MSForms.Control, rowLeft As Single
    Dim cmb As MSForms.ComboBox, lbl As MSForms.Label

    On Error GoTo Fail
    Set host = btn.Parent
    rowLeft = btn.Left
    For Each ctl In host.Controls
        If Abs((ctl.Top + ctl.height / 2) - (btn.Top + btn.height / 2)) < btn.height / 2 Then
            If ctl.Left < rowLeft Then rowLeft = ctl.Left
        End If
    Next ctl

    Set cmb = host.Controls.Add("Forms.ComboBox.1", "cmbQuality_" & btn.Name, True)
    With cmb
        .Style = fmStyleDropDownList
        .AddItem "Screen (1x)"
        .AddItem "High (2x)"
        .AddItem "Very high (3x)"
        .AddItem "Print (4x)"
        .ListIndex = 1
        .Font.Size = 9
        .width = 96
        .height = 18
        .Top = btn.Top + (btn.height - .height) / 2
        .Left = rowLeft - .width - 8
        If .Left < 6 Then .Left = 6
        .ControlTipText = "Resolution of the exported PNG: the chart is redrawn k times larger (1x = as on the sheet)"
    End With

    Set lbl = host.Controls.Add("Forms.Label.1", "lblQuality_" & btn.Name, True)
    With lbl
        .Caption = "PNG quality:"
        .Font.Size = 9
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .AutoSize = True
        .Top = cmb.Top + (cmb.height - .height) / 2
        .Left = cmb.Left - .width - 4
    End With

    Set AddQualityPicker = cmb
    Exit Function
Fail:
    Set AddQualityPicker = Nothing
End Function

' Scale factor selected in a quality list (1 when the list could not be created).
Private Function QualityScale(cmb As MSForms.ComboBox) As Double
    QualityScale = 1
    If cmb Is Nothing Then Exit Function
    If cmb.ListIndex >= 0 Then QualityScale = cmb.ListIndex + 1
End Function

' ==================== Data Section ====================
Private Sub btnRefreshSheets_Click()
    Call RefreshSheetList
End Sub

Private Sub RefreshSheetList()
    Dim ws As Worksheet
    Dim wbUser As Workbook
    Set wbUser = UserWorkbook()

    cmbDataSheet.Clear
    For Each ws In wbUser.Worksheets
        Select Case ws.Name
            Case "Gass_in", "Gass", "Fass", "Ain", "Bin", "optimization", _
                 "ForcesTable", "ThrustPlot", "DeformedPlot"
                ' skip generated sheets
            Case Else
                If Left(CStr(ws.Cells(2, 1).Value), 5) = "Block" Then
                    cmbDataSheet.AddItem ws.Name
                End If
        End Select
    Next ws
    If cmbDataSheet.ListCount > 0 Then cmbDataSheet.ListIndex = 0
End Sub

Private Sub btnImportCSV_Click()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Select the CSV file from the LISP export")
    If filePath = False Then Exit Sub

    Call LogMsg("Importing CSV: " & filePath)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wbUser As Workbook
    Set wbUser = UserWorkbook()

    Dim wbCSV As Workbook
    Dim sheetName As String
    sheetName = Left(GetFilenameWithoutExt(CStr(filePath)), 31)

    On Error Resume Next
    wbUser.Sheets(sheetName).Delete
    On Error GoTo 0

    Set wbCSV = Workbooks.Open(CStr(filePath))
    wbCSV.Sheets(1).Copy Before:=wbUser.Sheets(1)
    wbUser.Sheets(1).Name = sheetName
    wbCSV.Close SaveChanges:=False

    wbUser.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Call RefreshSheetList
    cmbDataSheet.Value = sheetName
    Call LogMsg("Imported '" & sheetName & "' into '" & wbUser.Name & "'")
End Sub

Private Function GetFilenameWithoutExt(fullPath As String) As String
    Dim fname As String
    fname = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    If InStrRev(fname, ".") > 0 Then
        fname = Left(fname, InStrRev(fname, ".") - 1)
    End If
    GetFilenameWithoutExt = fname
End Function

' ==================== Run Full Analysis ====================
Private Sub btnRunAll_Click()
    Dim dataSheet As String
    dataSheet = CStr(cmbDataSheet.Value)
    If dataSheet = "" Then
        MsgBox "Please select or import a CSV data sheet first.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim wbUser As Workbook
    Set wbUser = UserWorkbook()
    wbUser.Sheets(dataSheet).Activate

    Call LogMsg("=== Starting full analysis on '" & wbUser.Name & "' ===")

    ' Input checks before the solver: interface data against the block geometry, live load present
    Dim dataMsg As String, nbChk As Long
    Dim cbx() As Variant, cby() As Variant, cfx() As Variant, cfy() As Variant
    dataMsg = RepairInterfaceData(wbUser.Sheets(dataSheet))
    If Left(dataMsg, 6) = "ERROR:" Then
        Application.ScreenUpdating = True
        Call LogMsg(dataMsg)
        MsgBox Mid(dataMsg, 8), vbCritical, "Invalid geometry data"
        Exit Sub
    ElseIf Len(dataMsg) > 0 Then
        Call LogMsg("Data check: " & dataMsg & " (cells highlighted in yellow).")
    End If
    Call ReadCSVGeometry(wbUser.Sheets(dataSheet), cbx, cby, cfx, cfy, nbChk)
    If Not HasLiveLoad(wbUser.Sheets(dataSheet), nbChk) Then
        Application.ScreenUpdating = True
        Call LogMsg("ERROR: no live load in the data sheet (Fx_lam / Fy_lam are all zero).")
        MsgBox "The data sheet has no live load: all Fx_lam / Fy_lam values (columns L and M) are zero." & vbCrLf & _
               "Without a load to scale, the collapse multiplier is unbounded and the solver fails." & vbCrLf & _
               "Re-export the CSV and enter the loaded block number(s) and the load vector.", _
               vbCritical, "No live load"
        Exit Sub
    End If

    Call LogMsg("Step 1/6: Generating Gass_in matrix...")
    Application.Run "Generate_Gass_in"

    wbUser.Sheets(dataSheet).Activate
    Call LogMsg("Step 2/6: Generating Gass matrix...")
    Application.Run "GenerateGassMatrix"

    wbUser.Sheets(dataSheet).Activate
    Call LogMsg("Step 3/6: Creating Fass sheet...")
    Application.Run "CreateFassSheet"

    Call LogMsg("Step 4a/6: Creating Ain matrix...")
    Application.Run "CreateAinSheet"

    Call LogMsg("Step 4b/6: Creating Bin vector...")
    Application.Run "CreateBinSheet"

    Call LogMsg("Step 5/6: Building optimization sheet...")
    Application.Run "CreateOptimizationSheet"

    Call LogMsg("Step 6/6: Running Solver + back-substitution...")
    Application.Run "RunSolverAndBackSubstitute"

    Dim wsForces As Worksheet
    Set wsForces = wbUser.Sheets("ForcesTable")
    txtLambda.Value = Format(wsForces.Range("G2").Value, "0.0000")

    ' Collapse mechanism: four hinges checked kinematically (FindCollapseMechanism, Module9).
    ' Its kinematic multiplier must equal the static lambda of the solver.
    Dim hIdx() As Long, hSide() As Long, hX() As Double, hY() As Double
    Dim lamKin As Double, mechSign As Long, nFinal As Long, lamStat As Double
    lamStat = wsForces.Range("G2").Value
    If FindCollapseMechanismForSheet(wbUser.Sheets(dataSheet), wsForces, hIdx, hSide, hX, hY, lamKin, mechSign) Then
        nFinal = 4
        Call LogMsg("Collapse mechanism: hinges at interfaces " & hIdx(1) & ", " & hIdx(2) & ", " & _
                    hIdx(3) & ", " & hIdx(4) & " (kinematic lambda = " & Format(lamKin, "0.0000") & ").")
        If lamStat <> 0 Then
            If Abs(lamKin - lamStat) > 0.01 * Abs(lamStat) Then
                Call LogMsg("Warning: the kinematic lambda of the mechanism differs from the solver lambda by " & _
                            Format(Abs(lamKin - lamStat) / Abs(lamStat) * 100, "0.0") & _
                            "% - the solver result may be inaccurate.")
            End If
        End If
    Else
        nFinal = 0
        Call LogMsg("Warning: no admissible four-hinge collapse mechanism could be formed from the thrust line.")
    End If
    txtHinges.Value = CStr(nFinal)


    btnShowThrust.Enabled = True
    btnShowDeformed.Enabled = True

    Call LogMsg("=== Analysis complete! Lambda = " & _
                Format(wsForces.Range("G2").Value, "0.0000") & " N ===")

    If chkAutoPlots.Value Then
        Call LogMsg("Generating thrust line plot...")
        Application.Run "PlotThrustLine", wbUser.Sheets(dataSheet)
        Call LogMsg("Generating deformed shape plot...")
        Application.Run "PlotDeformedShape", wbUser.Sheets(dataSheet)
        Call LogMsg("Plots created on 'ThrustPlot' and 'DeformedPlot' sheets.")
        Call UpdateImagePanels(wbUser, wbUser.Sheets(dataSheet))
        MultiPage1.Value = 1
    End If

    Application.ScreenUpdating = True
    MsgBox "Analysis complete!" & vbCrLf & _
           "Lambda = " & Format(wsForces.Range("G2").Value, "0.0000") & " N" & vbCrLf & _
           "Hinges: " & nFinal, vbInformation, "Done"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call LogMsg("ERROR: " & Err.Description)
    MsgBox "Analysis failed: " & Err.Description, vbCritical
End Sub

' ==================== Show Plots ====================
Private Sub btnShowThrust_Click()
    On Error GoTo ErrorHandler
    Dim wbUser As Workbook: Set wbUser = UserWorkbook()
    If CStr(cmbDataSheet.Value) = "" Then
        MsgBox "Please select a data sheet first.", vbExclamation
        Exit Sub
    End If
    Application.Run "PlotThrustLine", wbUser.Sheets(CStr(cmbDataSheet.Value))
    Call UpdateImagePanels(wbUser, wbUser.Sheets(CStr(cmbDataSheet.Value)))
    MultiPage1.Value = 1
    Exit Sub
ErrorHandler:
    MsgBox "Could not show thrust line: " & Err.Description, vbCritical
End Sub

Private Sub btnShowDeformed_Click()
    On Error GoTo ErrorHandler
    Dim wbUser As Workbook: Set wbUser = UserWorkbook()
    If CStr(cmbDataSheet.Value) = "" Then
        MsgBox "Please select a data sheet first.", vbExclamation
        Exit Sub
    End If
    Application.Run "PlotDeformedShape", wbUser.Sheets(CStr(cmbDataSheet.Value))
    Call UpdateImagePanels(wbUser, wbUser.Sheets(CStr(cmbDataSheet.Value)))
    MultiPage1.Value = 2
    Exit Sub
ErrorHandler:
    MsgBox "Could not show deformed shape: " & Err.Description, vbCritical
End Sub

' ==================== Close ====================
Private Sub btnClose_Click()
    Unload Me
End Sub

' ==================== Logger ====================
Private Sub LogMsg(msg As String)
    Dim timestamp As String
    timestamp = Format(Now, "hh:mm:ss")
    If Len(txtLog.Value) = 0 Then
        txtLog.Value = "[" & timestamp & "] " & msg
    Else
        txtLog.Value = txtLog.Value & vbCrLf & "[" & timestamp & "] " & msg
    End If
    ' Force scroll to bottom
    txtLog.SelStart = Len(txtLog.Value)
    DoEvents
End Sub

' ==================== Image Panels ====================
' The charts are exported as PNG and resampled (GDI+, high-quality bicubic) to the exact pixel
' size of the panels, so the Image controls show them 1:1 instead of stretching a GIF.
Private Sub UpdateImagePanels(wbUser As Workbook, Optional wsData As Worksheet)
    Call ShowPlotOnPanel(imgThrust, wbUser, wsData, "Thrust")
    Call ShowPlotOnPanel(imgDeformed, wbUser, wsData, "Deformed")
End Sub

Private Sub ShowPlotOnPanel(img As MSForms.image, wbUser As Workbook, wsData As Worksheet, ByVal plotKind As String)
    Dim srcName As String, path As String, pic As StdPicture
    Dim dpi As Long, tw As Long, th As Long, w As Long, h As Long

    If plotKind = "Thrust" Then srcName = "ThrustPlot" Else srcName = "DeformedPlot"
    If Not SheetExists(wbUser, srcName) Then Exit Sub

    ' pixel size of the panel (minus its border)
    dpi = ScreenDpi()
    tw = CLng((img.width - 4) * dpi / 72)
    th = CLng((img.height - 4) * dpi / 72)
    If tw < 10 Or th < 10 Then Exit Sub

    path = Environ("TEMP") & "\excella_" & LCase(plotKind) & "_panel.png"
    On Error Resume Next
    If Not ExportPlotPNG(wbUser, Nothing, plotKind, 1, path) Then Exit Sub
    ' Chart smaller than the panel (high-DPI screen): redraw it at 2x rather than enlarging it
    If Not wsData Is Nothing And GetImageSizePx(path, w, h) Then
        If w < 0.95 * tw And h < 0.95 * th Then Call ExportPlotPNG(wbUser, wsData, plotKind, 2, path)
    End If

    Set pic = LoadPictureHQ(path, tw, th, OleColorToRGB(MultiPage1.BackColor))
    If pic Is Nothing Then
        ' fallback (GDI+ not available): previous behaviour
        If ExportChartGif(wbUser.Sheets(srcName).ChartObjects(1), Environ("TEMP") & "\excella_panel.gif") Then
            Set pic = LoadPicture(Environ("TEMP") & "\excella_panel.gif")
        End If
        img.PictureSizeMode = 3 ' fmPictureSizeModeZoom
    Else
        img.PictureSizeMode = 0 ' fmPictureSizeModeClip: already the right size, no resampling
        img.PictureAlignment = 2 ' fmPictureAlignmentCenter
    End If
    If Not pic Is Nothing Then Set img.Picture = pic
    If Err.Number <> 0 Then
        Call LogMsg("Warning: could not show the " & plotKind & " image (" & Err.Description & ").")
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function ExportChartGif(ch As ChartObject, path As String) As Boolean
    On Error GoTo Fail
    ch.Chart.Export filename:=path, FilterName:="GIF"
    ExportChartGif = True
    Exit Function
Fail:
    ExportChartGif = False
End Function

Private Function SheetExists(wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

' ==================== PNG Export ====================
Private Sub btnExportThrust_Click()
    Call ExportPlotWithDialog("Thrust", QualityScale(cmbQualityThrust))
End Sub

Private Sub btnExportDeformed_Click()
    Call ExportPlotWithDialog("Deformed", QualityScale(cmbQualityDeformed))
End Sub

Private Sub ExportPlotWithDialog(ByVal plotKind As String, ByVal scaleK As Double)
    Dim wbUser As Workbook, wsData As Worksheet, srcName As String, defName As String
    Dim savePath As Variant, w As Long, h As Long, sizeTxt As String

    Set wbUser = UserWorkbook()
    If plotKind = "Thrust" Then
        srcName = "ThrustPlot": defName = "ThrustLine.png"
    Else
        srcName = "DeformedPlot": defName = "DeformedShape.png"
    End If
    If Not SheetExists(wbUser, srcName) Then
        MsgBox IIf(plotKind = "Thrust", "Thrust line", "Deformed shape") & _
               " plot not found. Run the analysis first.", vbExclamation
        Exit Sub
    End If

    ' a redraw at scale k > 1 needs the data sheet; without it the existing chart is exported
    If scaleK > 1 Then
        If Len(CStr(cmbDataSheet.Value)) > 0 Then
            If SheetExists(wbUser, CStr(cmbDataSheet.Value)) Then Set wsData = wbUser.Sheets(CStr(cmbDataSheet.Value))
        End If
        If wsData Is Nothing Then
            Call LogMsg("Note: no data sheet selected - exporting the chart at screen resolution (1x).")
            scaleK = 1
        End If
    End If

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defName, _
        FileFilter:="PNG Image (*.png), *.png", _
        Title:="Save " & IIf(plotKind = "Thrust", "Thrust Line", "Deformed Shape") & " Plot")
    If VarType(savePath) = vbBoolean Then Exit Sub

    Application.Cursor = xlWait
    If ExportPlotPNG(wbUser, wsData, plotKind, scaleK, CStr(savePath)) Then
        Application.Cursor = xlDefault
        If GetImageSizePx(CStr(savePath), w, h) Then sizeTxt = " (" & w & " x " & h & " px)"
        Call LogMsg("Exported " & savePath & sizeTxt)
        MsgBox "Plot saved successfully" & sizeTxt & " to:" & vbCrLf & savePath, vbInformation
    Else
        Application.Cursor = xlDefault
        MsgBox "The plot could not be exported to:" & vbCrLf & savePath, vbExclamation
    End If
End Sub

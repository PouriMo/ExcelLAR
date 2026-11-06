"""Builds a clean ExcelLAR.xlam from src/vba (no leftover sheets, no VBA reference to Solver)."""
import win32com.client as w, os, sys
vba=os.path.abspath(sys.argv[1]); out=os.path.abspath(sys.argv[2])
app=w.DispatchEx('Excel.Application'); app.Visible=False; app.DisplayAlerts=False
try:
    wb=app.Workbooks.Add(1)
    vp=wb.VBProject
    order=['LinearSolver','Module0','Module1','Module2','Module3','Module4','Module5','Module6','Module7','Module8','Module9','Module10','ModuleImage','Module_Launcher']
    for n in order: vp.VBComponents.Import(os.path.join(vba,n+'.bas'))
    vp.VBComponents.Import(os.path.join(vba,'ExcelLAR.frm'))
    wb.BuiltinDocumentProperties('Title').Value='ExcelLAR'
    wb.BuiltinDocumentProperties('Subject').Value='Static limit analysis of masonry arches'
    wb.BuiltinDocumentProperties('Author').Value='G. Milani, N. Pingaro, M. Pourfouladi'
    wb.BuiltinDocumentProperties('Comments').Value='Version 1.1. Press Ctrl+Shift+A to open the ExcelLAR window.'
    wb.IsAddin=True
    if os.path.exists(out): os.remove(out)
    wb.SaveAs(out,55)
    print([r.Name for r in vp.References])
    print([c.Name for c in vp.VBComponents])
    wb.Close(False)
finally:
    app.Quit()

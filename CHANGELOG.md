# Changelog

## 1.1

### Installation and security

An installer is added: `Install_ExcelLAR.bat` checks Excel, verifies the checksum, copies and unblocks the add-in, finds or helps to install OpenSolver, registers both add-ins and runs a test example. `Uninstall_ExcelLAR.bat` removes the add-in. `INSTALL.md` documents the installer, the manual installation and the troubleshooting.


The add-in is rebuilt from the source files and contains no leftover worksheets. The VBA reference to `SOLVER.XLAM` at a fixed Office path is removed; it could stop the project from compiling on another computer. No routine uses the VBA project object model any more, so "Trust access to the VBA project object model" is not needed. The Excel Solver add-in is no longer a dependency. OpenSolver is called through its public API without a VBA reference, and the window explains what to do if it is missing.

### Linear program

The model is passed to OpenSolver with the CBC linear solver selected explicitly. Version 1.0 went through the Solver interface, which left Bonmin selected, a nonlinear solver.

OpenSolver discards coefficients smaller than 1E-5 when it builds a model. The row normalization of version 1.0 included the right-hand side in the row factor, which pushed the moment coefficients of the benchmarks below that limit: with CBC, the benchmarks gave 1144.85 and 266.77 instead of 1259.81 and 704.86. Module 5 now scales columns, rows and right-hand side, and the column factors are written on row 4 of sheet `optimization`. Module 6 checks the solution against the constraints on the sheet.

### Results and units

The dimensionless multiplier λu and the collapse force Pu = λu |F_ref| are separated in the window, the log, the plots and `ForcesTable`. |F_ref| is the magnitude of the resultant of the reference forces. `ForcesTable` has a new column |V/N|, the minimum friction coefficient required by the solution. The columns after M changed: F is |V/N|, G the status, H λu, I |F_ref| and J Pu.

Unit labels can be set in the window and follow the CSV headers. The AutoLISP routine has a new Metre option (drawing in m, specific weight in kN/m³, forces in kN), and it writes block weights with nine decimals. Coincidence tolerances of the CSV reader are now relative to the size of the drawing.

### Input checks

A stereotomy check stops the analysis if blocks overlap, if the joints of a block touch or the block is twisted, or if joints cross. Non-convex blocks produce a warning.

### Structure

Modules are numbered in execution order. Former Module 9 (mechanism) is Module 7, former Module 7 (thrust line) is Module 8, former Module 8 (deformed shape) is Module 9. Procedure names are unchanged. `RunAnalysis` in Module 0 is the single driver of the calculation, `RunAnalysisSilent` is its scripted entry point, and the new sheet `RunInfo` records versions, results and the time of each stage.

The hinge thresholds can be varied with `TestHingeThresholds` for sensitivity studies. Their default values are unchanged.

### Examples and validation

`55BlocksArch.csv` is replaced. The former file applied the load 0.79 mm away from the load point of the 11-block model; the new one keeps the same line of action with statically equivalent forces on two blocks (λu = 12.16, previously 11.67). New files: the refinement series in `examples/convergence/`, the drawings in `examples/cad/`, `Minimal_5Blocks.csv`, the minimum-thickness models of `examples/stereotomy/`, two files with the load moved, macro-free workbooks in `results/`, and an independent Python implementation in `validation/`.

## 1.0

First public release.

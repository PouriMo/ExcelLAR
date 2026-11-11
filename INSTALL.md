# Installing ExcelLAR

This guide installs ExcelLAR 1.1 and its solver, OpenSolver, for the current Windows user. It keeps the default security settings of Windows and Excel: no macro setting, Trust Center option or administrator right has to be changed.

## What is needed

| Component | Notes |
|---|---|
| Windows 10 or 11 | Excel for Mac is not supported, since the plots use Windows libraries. |
| Microsoft Excel for Windows | Tested with Microsoft 365, 64-bit. The code also supports 32-bit Excel, which has not been tested. |
| OpenSolver | Free solver for Excel (GPL), from https://opensolver.org. Tested with version 2.9.4. |
| AutoCAD | Optional, only to export geometry. Any CSV file in the documented format can be used instead. |

## Option A: installer (recommended)

1. On the GitHub page, click the green **Code** button, then **Download ZIP**.
2. Right-click the downloaded ZIP file, choose **Properties**, tick **Unblock** (if present) and click **OK**. Extract the ZIP file to a folder of your choice.
3. Close every Excel window.
4. Double-click `Install_ExcelLAR.bat` in the extracted folder. A console window opens and shows each step.

The installer performs these steps and stops with a clear message if one fails:

| Step | What happens |
|---|---|
| 1. Excel | Checks that Excel for Windows is installed and closed, and reports whether it is 32-bit or 64-bit. |
| 2. Checksum | Compares the SHA-256 of `src\ExcelLAR.xlam` with `SHA256SUMS.txt`, so a damaged or modified download is rejected. |
| 3. Copy | Copies the add-in to the Excel add-in folder, `%APPDATA%\Microsoft\AddIns`, and removes the download mark of that copy only. This is the same as ticking **Unblock** in its Properties. An earlier copy is kept as `ExcelLAR.xlam.bak`. |
| 4. OpenSolver | Looks for an OpenSolver already registered in Excel or extracted in `C:\OpenSolver`, Downloads, Desktop or Documents. If none is found, it opens the official OpenSolver page and asks for the path of `OpenSolver.xlam` after you have downloaded and extracted it. |
| 5. Registration | Starts a hidden Excel, registers OpenSolver and ExcelLAR in the add-in list (as **File** > **Options** > **Add-ins** would), and offers to deactivate an older ExcelLAR add-in, without deleting it. |
| 6. Test | Runs `examples\Minimal_5Blocks.csv` and compares the collapse multiplier with the expected value, 1226.1996. |

At the end, open Excel and press **Ctrl+Shift+A** in any workbook. A log of the installation is written to `%TEMP%\ExcelLAR_install.log`.

The `.bat` file starts PowerShell with `-ExecutionPolicy Bypass` for this script only; no system setting is changed. Advanced users can run the script directly, for example `install\Install-ExcelLAR.ps1 -OpenSolverPath C:\OpenSolver\OpenSolver.xlam`. The options are listed at the top of the script.

To remove ExcelLAR, double-click `Uninstall_ExcelLAR.bat`. OpenSolver is not changed.

## Option B: manual installation

1. Download the OpenSolver ZIP file for Windows from https://opensolver.org/installing-opensolver/.
2. Right-click the ZIP file, choose **Properties**, tick **Unblock** and click **OK**. Extract it to a simple folder such as `C:\OpenSolver`. Keep all its files together.
3. Download `src/ExcelLAR.xlam` from this repository and save it in a permanent folder, for example `%APPDATA%\Microsoft\AddIns`.
4. Right-click `ExcelLAR.xlam`, choose **Properties**, tick **Unblock** and click **OK**.
5. Open Excel, go to **File** > **Options** > **Add-ins**, choose **Excel Add-ins** in the **Manage** list and click **Go**.
6. Click **Browse**, select `OpenSolver.xlam` and click **OK**. Repeat for `ExcelLAR.xlam`. Both must be ticked.
7. Press **Ctrl+Shift+A** in any workbook to open ExcelLAR.

To verify the download, run `Get-FileHash ExcelLAR.xlam -Algorithm SHA256` in PowerShell and compare the result with `SHA256SUMS.txt`.

## First analysis

1. Press **Ctrl+Shift+A**. The ExcelLAR window opens.
2. Click **Import CSV File** and select `examples\11BlocksArch.csv`.
3. Click **RUN ANALYSIS**. The collapse multiplier is 13.19 and the hinges form at interfaces 1, 3, 8 and 12.

The README gives the expected results of all the examples of the paper.

## Troubleshooting

| Symptom | Cause and remedy |
|---|---|
| "Microsoft has blocked macros from running because the source of this file is untrusted", or a red Security Risk bar | The file still carries the download mark. Close Excel, tick **Unblock** in the Properties of the file, or run the installer, which does it for the installed copy. A Trusted Location of your own (**Trust Center Settings** > **Trusted Locations**) also works. Do not enable all macros. |
| Compile error, or "Can't find project or library" | This occurred with version 1.0, which referenced the Excel Solver add-in at a fixed path. Version 1.1 has no such reference: remove the old add-in from the list and install `src\ExcelLAR.xlam` of this release. |
| "OpenSolver is not loaded" | OpenSolver is missing or unticked. Install it as above and run the installer again, or tick it in the Add-ins list. |
| OpenSolver starts but fails | Its folder must not contain accents or non-English characters, nor be a network path. Extract it again to `C:\OpenSolver`. |
| The installer says Excel is open | Close every Excel window. Check the Task Manager for a hidden Excel process left by another program. |
| The installer cannot start PowerShell | Run `install\Install-ExcelLAR.ps1` from a PowerShell window, or follow Option B. |
| Ctrl+Shift+A does nothing | The add-in is not ticked, or an older add-in uses the same shortcut. Open the Add-ins list, tick ExcelLAR and untick older versions. |
| The test step fails | Read the message in the console and in the log file. Report it as below. |
| "do not share an edge", "overlap" or "joints touch" | The drawing does not form a chain of blocks. See Input checks in the README. |
| "infeasible" | No thrust line in equilibrium with the dead loads fits inside the arch. |

To report a problem, open an issue on GitHub or write to mohammad.pourfouladi@polimi.it. Please include the Excel version and build (**File** > **Account** > **About Excel**), the text of the message and the log file.

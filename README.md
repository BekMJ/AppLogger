## Lab App Logger

This PowerShell-based launcher prompts users for Name, Advisor, and Experiment before opening the target application. It logs entries with timestamp, Windows username, and computer name to CSV (monthly files) and, if Microsoft Excel is installed, also to an `.xlsx` workbook.

### Files
- `AppLogger.ps1`: GUI logger and launcher
 - `Launch-ToupView.cmd`: Launcher for ToupView
- `Launch-Opus.cmd`: Launcher for Opus
- `Launch-ThorCam.cmd`: Launcher for ThorCam
- `Launch-FlowVision.cmd`: Example launcher for FlowVision
- `Logs/`: Created automatically to store monthly CSV/XLSX (e.g., `LabUsage-202510.csv` and `.xlsx`)

### How it works
1. User double-clicks the app launcher (e.g., `Launch-ToupView.cmd`, `Launch-Opus.cmd`, `Launch-ThorCam.cmd`, or `Launch-FlowVision.cmd`).
2. A small form appears asking for Name, Advisor, Experiment.
3. On OK, a row is appended to `Logs/LabUsage-YYYYMM.csv` (and to `.xlsx` if Excel is available).
4. The target app is started.

### Setup
1. Copy this folder to the lab PC (e.g., `C:\LabTools\AppLogger`).
2. Edit the launcher(s) you need:
   - `Launch-ToupView.cmd`, `Launch-Opus.cmd`, `Launch-ThorCam.cmd` (and `Launch-FlowVision.cmd` as an example)
   - Set `APP_PATH` to the actual application path.
   - Optionally set `APP_ARGS`.
3. Create a desktop shortcut for the chosen `.cmd` and instruct users to launch the app via that shortcut.

### Creating more launchers
Copy `Launch-FlowVision.cmd` to a new file (e.g., `Launch-YourApp.cmd`) and change:
```
set "APP_PATH=C:\Path\To\YourApp.exe"
set "APP_ARGS=--optional"
...
powershell -NoProfile -ExecutionPolicy Bypass -File "%LOGGER_PS%" -AppPath "%APP_PATH%" -AppArgs "%APP_ARGS%" -AppName "YourApp"
```

### Direct PowerShell usage
You can call the logger directly if you wrap your own shortcuts:
```
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\LabTools\AppLogger\AppLogger.ps1" -AppPath "C:\Path\App.exe" -AppArgs "--any" -AppName "App"
```

### Permissions and policy
- The script writes to the `Logs` folder next to the script. Ensure users have write permission.
- The example `.cmd` uses `-ExecutionPolicy Bypass` so you do not need to change the system Execution Policy.

### CSV-only mode
If you want to skip Excel entirely, pass `-ForceCsvOnly` to `AppLogger.ps1` in your launcher.

### Data columns
`Timestamp, Name, Advisor, Experiment, ComputerName, WindowsUser, App`

### Notes
- Files rotate monthly (`LabUsage-YYYYMM.*`).
- Excel logging uses COM automation if Excel is installed. If not, CSV is still recorded.


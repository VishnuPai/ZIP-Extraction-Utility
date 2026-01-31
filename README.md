# ZIP Extraction Utility



Windows PowerShell GUI tool for batch ZIP extraction with progress tracking, overwrite protection, and server-safe processing. Designed for Windows Server and enterprise environments.



🔍 Project Overview



ZIP Extraction Utility is a lightweight Windows GUI application built with PowerShell that enables controlled batch extraction of ZIP files from a selected source directory into a specified output folder.



Designed for internal enterprise and server environments, this tool simplifies bulk ZIP processing through a user-friendly interface while maintaining stability and overwrite protection.



🚀 Features



📁 Browse and select Source and Output folders



🗂 Batch extract multiple ZIP files



📅 Optional filtering by date \& time range



🏷 Optional filtering by specific ZIP file names



🔁 Automatic overwrite protection (auto-rename duplicates)



📊 Real-time progress tracking with percentage



🛑 Cancel support during processing



📝 Execution logging for auditing



🖥 Windows Server 2019 compatible



💻 Windows 11 compatible



📦 Packaged as standalone .exe using ps2exe



🏢 Intended Use



This utility is designed for:



Internal enterprise environments



Server-side file processing workflows



IT operations teams



Batch data extraction scenarios



Non-technical users requiring GUI-based ZIP handling



🏗 Architecture

Technology Stack



PowerShell 5.1



System.Windows.Forms (WinForms GUI)



ps2exe (for standalone executable packaging)



Processing Model



Sequential ZIP extraction for maximum stability



Temporary extraction folder per ZIP



Safe file copy with duplicate name handling



UI-driven execution loop with progress updates



Logging to local directory for traceability



Design Principles



Stability over complexity



Server-safe execution



Minimal dependencies



Clear and maintainable code



GUI-first user interaction



⚙ Installation



Option 1 — Use EXE (Recommended)



Download the compiled .exe and run directly.



Option 2 — Run PowerShell Script



.\\ZipUtility.ps1



Option 3 — Build EXE Yourself



🔧 Building the Executable (ps2exe)



This project can be packaged into a standalone Windows .exe file using the ps2exe PowerShell module.



1️⃣ Install ps2exe (One-Time Setup)



Open PowerShell as a normal user and run:



Install-Module ps2exe -Scope CurrentUser



If prompted about an untrusted repository, choose:



A  (Yes to All)



2️⃣ Build the Executable



Navigate to the folder containing the script and run:



Invoke-ps2exe .\\ZipUtility.ps1 ZipUtility.exe -noConsole



🔹 What -noConsole Does



Removes the black PowerShell console window and launches only the GUI.



3️⃣ Run the Executable



After building, you will see:



ZipUtility.exe



Double-click to run.



No PowerShell window will appear.



⚠ Notes



Designed for PowerShell 5.1



Works on Windows Server 2019 and Windows 11



No administrative privileges required



If execution policy blocks installation, run:



Set-ExecutionPolicy RemoteSigned -Scope CurrentUser





📝 Logging



Execution logs are stored in:



C:\\ZipUtilityLogs\\



Log entries include:



Execution timestamp



Username



Number of ZIP files processed



🔐 Security Considerations



Designed for internal usage



No internet connectivity required



No external dependencies



No system tray dependency



Runs without administrative privileges



📌 Version



Current Version: 1.0.0



📜 License



This project is licensed under the MIT License.



Why MIT?



Simple



Permissive



Suitable for internal enterprise tools



Allows modification and redistribution



You can create a LICENSE file with standard MIT license text.



🤝 Contributing



This project is intended for internal usage.

Pull requests and improvements are welcome.




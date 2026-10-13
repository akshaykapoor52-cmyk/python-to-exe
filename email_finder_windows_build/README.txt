Email Finder — Windows EXE build

1. Extract this folder on a Windows computer.
2. Install Python for Windows from https://www.python.org/downloads/windows/
   During setup, enable "Add Python to PATH".
3. Double-click build_exe.bat.
4. When it finishes, find the program at dist\email_finder.exe.

The build requires an internet connection to install openpyxl and PyInstaller.
Keep the EXE in a writable folder when using it. The program edits and saves
the selected workbook, so make a backup of your Excel file first.

Note: This bundle builds a Windows EXE on Windows; it is not itself the EXE.

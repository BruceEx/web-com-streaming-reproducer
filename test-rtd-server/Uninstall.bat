regedit /s "%~dp0open_entry_delete.reg"
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe "%~dp0ExcelRtdTest.dll" /u /codebase /verbose 
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "%~dp0ExcelRtdTest.dll" /u /codebase /verbose 
pause

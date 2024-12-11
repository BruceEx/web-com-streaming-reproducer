regedit /s "%~dp0open_entry_add.reg"
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe "%~dp0ExcelRtdTest.dll" /codebase /verbose 
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "%~dp0ExcelRtdTest.dll" /codebase /verbose 
pause

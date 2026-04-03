@echo off
echo Unregistering old version...
regsvr32 /u /s "C:\Program Files\VBE AddIn\VBEAddIn.dll"

echo Registering new version...
regsvr32 /s "C:\Program Files\VBE AddIn\VBEAddIn.dll"

echo Done! Start Excel now.
pause

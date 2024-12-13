@echo off
echo Please ensure all Excel instances are closed.
pause
del "C:\Users\Ehsan.mokhtari\AppData\Roaming\Microsoft\AddIns\SRG_AddIn.xlam"
if exist "C:\Users\Ehsan.mokhtari\AppData\Roaming\Microsoft\AddIns\SRG_AddIn.xlam" (
    echo Failed to delete the add-in file. Please check permissions.
) else (
    echo Add-in file deleted successfully!
)
pause
exit

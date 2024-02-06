@echo off
setlocal EnableDelayedExpansion

set "interpreterScript=simulateKeyboard.vbs"

for %%f in (*.txt) do (
    cscript //NoLogo %interpreterScript% "%%f"
)

exit /b 0

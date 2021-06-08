
call "D:\output_to_csv\output_to_json\serialize.bat"
@echo off
setlocal EnableDelayedExpansion
set "json="
for /f "tokens=*" %%x in ('type "D:\output_to_csv\output_to_json\temp.json"') do (
    set "mod=%%x"
    set "mod=!mod:"=\"!"
    set "json=!json! !mod!" )
@echo on
curl ^
-d "{ \"action\":\"xrf_service\", \"pass\":\"efbuy3uy42ub429d\", \"csv\": %json% }" ^
http://192.168.17.61:8085/dyn/dash/cc/defos/xrf_service.php
endlocal

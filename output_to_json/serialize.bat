
@echo off
echo [ > D:\output_to_csv\output_to_json\temp.json
setlocal EnableDelayedExpansion
for /f "tokens=*" %%x in ('type "D:\output_to_csv\data\temp.csv"') do (
    set "mod=%%x"
    set "mod=!mod:,=","!"
    echo ["!mod!"], >> D:\output_to_csv\output_to_json\temp.json )
echo null] >> D:\output_to_csv\output_to_json\temp.json
endlocal

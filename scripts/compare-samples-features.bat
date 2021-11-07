@echo off

setlocal

::set "WINMERGE=WinMergeU"
set "WINMERGE=winmerge"

echo:Plan 1: Maximize the minified example file
cscript //nologo test-maximize.wsf <..\data\samples\features.min.example.vbs >..\data\samples\features.max.example.vbs

echo:Plan 2: Minify the original file then maximize it
cscript //nologo ..\bin\vbsmin.wsf <..\data\samples\features.vbs | cscript //nologo test-maximize.wsf >..\data\samples\features.max.vbs

echo:Plan 3: Compare using WinMerge
call %WINMERGE% ..\data\samples\features.max.vbs ..\data\samples\features.max.example.vbs ..\data\samples\features.vbs

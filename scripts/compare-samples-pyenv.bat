@echo off

setlocal

::set "WINMERGE=WinMergeU"
set "WINMERGE=winmerge"

echo:Plan 1: Maximize the minified example file
cscript //nologo test-maximize.wsf <..\data\samples\pyenv.min.example.vbs >..\data\samples\pyenv.max.example.vbs

echo:Plan 2: Minify the original file then maximize it
cscript //nologo ..\bin\vbsmin.wsf <..\data\samples\pyenv.vbs | cscript //nologo test-maximize.wsf >..\data\samples\pyenv.max.vbs

echo:Plan 3: Compare using WinMerge
call %WINMERGE% ..\data\samples\pyenv.max.vbs ..\data\samples\pyenv.max.example.vbs ..\data\samples\pyenv.vbs

@echo off

if "%~1" == "" (
	echo:Usage: %~n0 FILE
	goto :EOF
)

ruby -r "%~dp0vbsmin.rb" -e "VBSMin.new.minify(ARGV[0])" %*

@echo off

set "tail=\build\vc14-x64"

if exist ..%tail% rd /q /s ..%tail%

mkdir ..%tail%

pushd ..%tail%

set errorlevel = 0

cmake ^
	-G"Visual Studio 15 2017 Win64" -T "v141" ^
	"-DBUILD_WITH_STATIC_CRT:BOOL=ON" ^
	"-Wno-dev" ^
	../..

if errorlevel 1 goto :error

echo [*] Project generation succeeded!

popd
goto :eof

:error
echo [!] ERROR: Failed to generate project - see above and correct.
popd
exit /b 1

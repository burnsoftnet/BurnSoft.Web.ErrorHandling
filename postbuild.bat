@echo off
REM ******************************************************
REM * VS Buid Script when using GhostScript for chm file *
REM * and building nuget packages from Solution 	 *
REM ******************************************************
REM ******************************************************
REM USE Build Script below for Solution
REM $(SolutionDir)postbuild.bat $(SolutionDir) $(ProjectDir) $(ConfigurationName) $(ProjectName)
REM 
REM USE Build Script below for Project Only
REM $(ProjectDir)postbuild.bat $(SolutionDir) $(ProjectDir) $(ConfigurationName) $(ProjectName)
REM ******************************************************
SET SolutionDir=%1
SET ProjectDir=%2
SET ConfigurationName=%3
SET HELPFILENAME=%4
SET DEBUG="Debug"
SET RELEASE="Release"
SET NUGETEXT=nupkg
SET USENUGETSERVER=""
cd "%ProjectDir%"
copy /Y "%SolutionDir%Help\%HELPFILENAME%.chm" "%ProjectDir%bin\%ConfigurationName%\%HELPFILENAME%.chm"
cd "%ProjectDir%"
del /Q %ProjectDir%*.nupkg

if "%ConfigurationName%" == %DEBUG% (
	echo "nuget Dev packing"
	SET USENUGETSERVER="http://nuget.burnsoft.test"
	nuget pack
)

if "%ConfigurationName%" == %RELEASE% (
	echo "nuget Production Packing"
	SET USENUGETSERVER="http://nuget.burnsoft.prod"
	nuget pack -Prop Configuration=Release
)

SET nupak=""
for /R "%CD%" %%f in (*.%NUGETEXT%) do (
	SET nupak=%%~nf
	echo %%~nf
)
echo "removing %HELPFILENAME%. from %nupak% "

FOR /F "tokens=* USEBACKQ" %%F IN (`%SolutionDir%StringReplace %nupak% %HELPFILENAME%.`) DO (
	SET ver=%%F
)
echo "PackageVersion: %ver%"
echo "deleting %HELPFILENAME% %ver%"
nuget delete %HELPFILENAME% %ver% burnsoft -Source %USENUGETSERVER% -NonInteractive
echo "Uploading %nupak%.%NUGETEXT%"
nuget push %nupak%.%NUGETEXT% burnsoft -Source %USENUGETSERVER%

if "%ConfigurationName%" == %RELEASE% (
	echo "nuget guthub push"
	nuget push %nupak%.%NUGETEXT% -source "github"
)
cd ..
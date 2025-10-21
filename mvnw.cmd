@echo off
SETLOCAL
set BASEDIR=%~dp0
if not defined MAVEN_HOME (
  if not defined M2_HOME (
    set MAVEN_WRAPPER_JAR=%BASEDIR%\.mvn\wrapper\maven-wrapper.jar
  ) else (
    set MAVEN_EXEC=%M2_HOME%\bin\mvn
  )
) else (
  set MAVEN_EXEC=%MAVEN_HOME%\bin\mvn
)

if defined MAVEN_EXEC (
  "%MAVEN_EXEC%" %*
  exit /b %errorlevel%
)

if exist "%MAVEN_WRAPPER_JAR%" (
  java -jar "%MAVEN_WRAPPER_JAR%" %*
  exit /b %errorlevel%
)

echo Maven wrapper JAR not found. Run the included PowerShell script to download it.
exit /b 1

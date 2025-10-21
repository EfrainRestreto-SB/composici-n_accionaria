#!/usr/bin/env bash
# ----------------------------------------------------------------------------
# Maven wrapper
# Auto-generated wrapper script.
# ----------------------------------------------------------------------------

basedir="$(dirname "$0")"

if [ -z "$MAVEN_HOME" ] ; then
  if [ -z "$M2_HOME" ] ; then
    MAVEN_EXEC="$basedir/.mvn/wrapper/maven-wrapper.jar"
  else
    MAVEN_EXEC="$M2_HOME/bin/mvn"
  fi
else
  MAVEN_EXEC="$MAVEN_HOME/bin/mvn"
fi

if [ -f "$MAVEN_EXEC" ] ; then
  java -jar "$MAVEN_EXEC" "$@"
else
  echo "Maven wrapper JAR not found. Run the included download script: powershell -File download-maven-wrapper.ps1" >&2
  exit 1
fi

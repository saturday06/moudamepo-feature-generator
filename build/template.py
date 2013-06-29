#!/usr/bin/env python
# -*- coding: utf-8-unix; mode: javascript -*-
# vim: set ft=javascript
#
# Automatic feature generation
#
# Usage: 
#   ./spec.py inputdirectory outputdirectory
#

generateFeatureJs = r""" //" // magic comment for editor's syntax highlighing
%%%GENERATE_FEATURE_JS%%%
"""
#" /* magic comment for editor's syntax highlighting

odBase64 = """
%%%GENERATE_FEATURE_DUMMY_OD%%%
"""

import uno
import unohelper

import atexit
import base64
import datetime
import os
import signal
import sys
import tempfile
import zipfile
from time import sleep
from subprocess import Popen
from com.sun.star.script.provider import XScriptContext
from com.sun.star.connection import NoConnectException
from com.sun.star.util import Date
from com.sun.star.beans import PropertyValue

odFile = tempfile.NamedTemporaryFile()
odFile.write(base64.b64decode(odBase64))
#print(odFile.name)
odFile.flush()

scriptOdFile = tempfile.NamedTemporaryFile()
#print(scriptOdFile.name)
scriptOdFile.flush()

generateFeatureJsPath = "Scripts/javascript/Library/GenerateFeature.js"
zin = zipfile.ZipFile (odFile.name, 'r')
zout = zipfile.ZipFile (scriptOdFile.name, 'w')
for item in zin.infolist():
    buffer = zin.read(item.filename)
    #print(item.filename)
    if (item.filename == generateFeatureJsPath):
        zout.writestr(item, generateFeatureJs)
    else:
        zout.writestr(item, buffer)
zout.close()
zin.close()

pipeName = "generatefeaturepipe"
acceptArg = "--accept=pipe,name=%s;urp;StarOffice.ServiceManager" % pipeName
url = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipeName
officePath = "soffice"
process = Popen([officePath, acceptArg
                 , "--nologo"
                 , "--norestore"
                 #, "--invisible"
                 #, "--minimized"
                 #, "--headless"                 
])

ctx = None
for i in range(5):
    try:
        localctx = uno.getComponentContext()
        resolver = localctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localctx)
        ctx = resolver.resolve(url)
    except NoConnectException:
        sleep(1)
    if ctx:
        break

desktop = ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
macroExecutionModeArg = PropertyValue()
macroExecutionModeArg.Name = "MacroExecutionMode"
macroExecutionModeArg.Value = 4

readOnlyArg = PropertyValue()
readOnlyArg.Name = "ReadOnly"
readOnlyArg.Value = True

print("GenerateFeature")
print("script: " + scriptOdFile.name)
url = "file://" + scriptOdFile.name
#print(url)
document = desktop.loadComponentFromURL(url, "_blank", 0, (macroExecutionModeArg, readOnlyArg));
macroUrl = "vnd.sun.star.script:Library.GenerateFeature.js?language=JavaScript&location=document"

scriptProvider = document.getScriptProvider();
script = scriptProvider.getScript(macroUrl)

try:
  script.invoke(tuple(sys.argv), (), ())
finally:
  try:
      document.dispose()
  except Exception: # __main__.DisposeException
      None
  try:
      desktop.terminate()
  except Exception: # __main__.DisposeException
      None
  process.terminate()

# magic comment terminator */

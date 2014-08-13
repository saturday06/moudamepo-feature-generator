#!/usr/bin/env python
# -*- coding: utf-8-unix; mode: javascript -*-
# vim: set ft=javascript
#
# Automatic feature generation
#
# Usage: 
#   ./generatefeature.py inputdirectory outputdirectory
#

generateFeatureJs = r""" //" // magic comment for editor's syntax highlighing
%%%GENERATE_FEATURE_JS%%%
"""
#" /* magic comment for editor's syntax highlighting

import uno
import unohelper

import atexit
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

pipeName = "generatefeaturepipe"
acceptArg = "-accept=pipe,name=%s;urp;StarOffice.ServiceManager" % pipeName
url = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipeName
officePath = "soffice"
process = Popen([officePath, acceptArg
                 #, "-nologo"
                 , "-norestore"
                 , "-invisible"
                 #, "-minimized"
                 #, "-headless"
])

ctx = None
for i in range(20):
    print("Connectiong...")
    try:
        localctx = uno.getComponentContext()
        resolver = localctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localctx)
        ctx = resolver.resolve(url)
    except NoConnectException:
        sleep(i * 2 + 1)
    if ctx:
        break
    if process.poll():
        raise Exception("Process exited")
if not ctx:
    raise Exception("Connection failure")

desktop = ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

tempDir = tempfile.mkdtemp()
emptyOdPath = tempDir + "/empty.odg"
emptyOdExtractPath = tempDir + "/empty.odg.extract"
emptyOdUrl = "file://" + emptyOdPath
hiddenArg = PropertyValue()
hiddenArg.Name = "Hidden"
hiddenArg.Value = True
emptyDocument = desktop.loadComponentFromURL("private:factory/sdraw", "_blank", 0, (hiddenArg,));
emptyDocument.storeToURL(emptyOdUrl, ())
emptyDocument.dispose()

#scriptOdFile = tempfile.NamedTemporaryFile("w+b", -1, "ods")
scriptOdPath = "/home/saturday06/tmpx/asdf.odg"
scriptOdUrl = "file://" + scriptOdPath

with zipfile.ZipFile(emptyOdPath, "r") as zin:
    zin.extractall(emptyOdExtractPath)

manifest = None
with open(emptyOdExtractPath + "/META-INF/manifest.xml", "r") as f:
    manifest = f.read()

with open(emptyOdExtractPath + "/META-INF/manifest.xml", "w") as f:
    f.write(manifest.replace("</manifest:manifest>", r"""
  <manifest:file-entry manifest:full-path="Scripts/javascript/Library/GenerateFeature.js" manifest:media-type=""/>
  <manifest:file-entry manifest:full-path="Scripts/javascript/Library/parcel-descriptor.xml" manifest:media-type=""/>
  <manifest:file-entry manifest:full-path="Scripts/javascript/Library/" manifest:media-type="application/binary"/>
  <manifest:file-entry manifest:full-path="Scripts/javascript/" manifest:media-type="application/binary"/>
  <manifest:file-entry manifest:full-path="Scripts/" manifest:media-type="application/binary"/>
</manifest:manifest>
""".strip()))

scriptDir = emptyOdExtractPath + "/Scripts/javascript/Library"
if not os.path.exists(scriptDir):
    os.makedirs(scriptDir)

with open(scriptDir + "/GenerateFeature.js", "w") as f:
    f.write(generateFeatureJs)

with open(scriptDir + "/parcel-descriptor.xml", "w") as f:
    f.write(r"""
<?xml version="1.0" encoding="UTF-8"?>
<parcel language="JavaScript" xmlns:parcel="scripting.dtd">
  <script language="JavaScript">
    <locale lang="en">
      <displayname value="GenerateFeature.js"/>
      <description>GenerateFeature.js</description>
    </locale>
    <logicalname value="GenerateFeature.js"/>
    <functionname value="GenerateFeature.js"/>
  </script>
</parcel>
""".strip())

with zipfile.ZipFile(scriptOdPath, "w") as zout:
    for dir, subdirs, files in os.walk(emptyOdExtractPath):
        arcdir = os.path.relpath(dir, emptyOdExtractPath)
        if not arcdir == ".":
            zout.write(dir, arcdir)
        for file in files:
            arcfile = os.path.join(os.path.relpath(dir, emptyOdExtractPath), file)
            zout.write(os.path.join(dir, file), arcfile)

macroExecutionModeArg = PropertyValue()
macroExecutionModeArg.Name = "MacroExecutionMode"
macroExecutionModeArg.Value = 4

readOnlyArg = PropertyValue()
readOnlyArg.Name = "ReadOnly"
readOnlyArg.Value = True

#print(url)
document = desktop.loadComponentFromURL(scriptOdUrl, "_blank", 0, (macroExecutionModeArg, readOnlyArg, hiddenArg));
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

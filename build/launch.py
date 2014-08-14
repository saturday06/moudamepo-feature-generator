# Launcher fragment
# -*- coding: us-ascii-dos -*-

import uno
import unohelper

import datetime
import os
import re
import signal
import sys
import tempfile
import threading
import zipfile
import codecs
from time import sleep
from subprocess import Popen
from com.sun.star.script.provider import XScriptContext
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue

if not 'generateFeatureJs' in locals():
    generateFeatureJs = None
    with codecs.open(sys.argv[3], encoding='utf-8') as f:
        generateFeatureJs = f.read()

pipeName = "generatefeaturepipe"
acceptArg = "-accept=pipe,name=%s;urp;StarOffice.ServiceManager" % pipeName
url = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipeName
officePath = "soffice"
process = Popen([officePath, acceptArg
                 , "-nologo"
                 , "-norestore"
                 , "-invisible"
                 #, "-minimized"
                 #, "-headless"
])

ctx = None
for i in range(20):
    print("Connectiong...")
    sys.stdout.flush()
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

tempDir = os.path.abspath(tempfile.mkdtemp()).replace("\\", "/")
emptyOdPath = tempDir + "/empty.ods"
emptyOdExtractPath = tempDir + "/empty.ods.extract"
emptyOdUrl = "file://" + re.sub(r'^/?', "/", emptyOdPath)
hiddenArg = PropertyValue()
hiddenArg.Name = "Hidden"
hiddenArg.Value = True
emptyDocument = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (hiddenArg,));
emptyDocument.storeToURL(emptyOdUrl, ())
emptyDocument.dispose()

scriptOdPath = tempDir + "/script.ods"
scriptOdUrl = "file://" + re.sub(r'^/?', "/", scriptOdPath)

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

if os.name == 'nt':
    encoding = sys.stdin.encoding
else:
    encoding = "utf-8"
with codecs.open(scriptDir + "/GenerateFeature.js", "w", encoding) as f:
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

document = desktop.loadComponentFromURL(scriptOdUrl, "_blank", 0, (macroExecutionModeArg, readOnlyArg, hiddenArg));
macroUrl = "vnd.sun.star.script:Library.GenerateFeature.js?language=JavaScript&location=document"

scriptProvider = document.getScriptProvider();
script = scriptProvider.getScript(macroUrl)
logPath = tempDir + "/script.log"
print("log=" + logPath);
args = (sys.argv[0], logPath) + tuple(sys.argv[1:])
print("args=%s" % (args,));
sys.stdout.flush()

def tailF():
    print("tailf=" + logPath);
    pos = 0
    while True:
        sleep(0.5)
        if not os.path.exists(logPath):
            continue
        with codecs.open(logPath, encoding='utf-8') as f:
            if f.seek(0, 2) == pos:
                continue
            f.seek(pos)
            data = f.read()
            sys.stdout.write(data)
            sys.stdout.flush()
            pos += len(data.encode('utf-8'))

t = threading.Thread(target=tailF)
t.daemon = True
t.start()
try:
    script.invoke(args, (), ())
finally:
    t.join(3)
    try:
        document.dispose()
    except Exception: # __main__.DisposeException
        None
    try:
        desktop.terminate()
    except Exception: # __main__.DisposeException
        None
    process.terminate()

# Javascript comment terminator */

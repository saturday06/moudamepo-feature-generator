# Launcher fragment
# -*- coding: us-ascii-dos -*-

import uno
import unohelper

import datetime
import os
import re
import signal
import shutil
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

def CreateUnoService(context, name):
    return context.getServiceManager().createInstanceWithContext(name, context)

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

context = None
for i in range(20):
    print("Connectiong...")
    sys.stdout.flush()
    try:
        localContext = uno.getComponentContext()
        resolver = CreateUnoService(localContext, "com.sun.star.bridge.UnoUrlResolver")
        context = resolver.resolve(url)
    except NoConnectException:
        sleep(i * 2 + 1)
    if context:
        break
    if process.poll():
        raise Exception("Process exited")
if not context:
    raise Exception("Connection failure")

desktop = CreateUnoService(context, "com.sun.star.frame.Desktop")

tempDir = os.path.abspath(tempfile.mkdtemp()).replace("\\", "/")
odPath = tempDir + "/script.ods"
print("Working document: " + odPath)
odUrl = "file://" + re.sub(r'^/?', "/", odPath)
hiddenArg = PropertyValue()
hiddenArg.Name = "Hidden"
hiddenArg.Value = True

document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (hiddenArg,));
tddcf = CreateUnoService(context, "com.sun.star.frame.TransientDocumentsDocumentContentFactory")
fileAccess = CreateUnoService(context, "com.sun.star.ucb.SimpleFileAccess")
content = tddcf.createDocumentContent(document)
scriptsDir = content.getIdentifier().getContentIdentifier() + "/Scripts"
jsDir = scriptsDir + "/javascript"
libraryDir = jsDir + "/Library"
fileAccess.createFolder(scriptsDir)
fileAccess.createFolder(jsDir)
fileAccess.createFolder(libraryDir)

scriptPipe = CreateUnoService(context, "com.sun.star.io.Pipe")
scriptOut = CreateUnoService(context, "com.sun.star.io.TextOutputStream")
scriptOut.setOutputStream(scriptPipe)
if os.name == 'nt':
    scriptOut.setEncoding(sys.stdin.encoding)
else:
    scriptOut.setEncoding("UTF-8")
scriptOut.writeString(generateFeatureJs)
scriptOut.flush()
scriptOut.closeOutput()
fileAccess.writeFile(libraryDir + "/GenerateFeature.js", scriptPipe)
scriptPipe.closeInput()

descriptorPipe = CreateUnoService(context, "com.sun.star.io.Pipe")
descriptorOut = CreateUnoService(context, "com.sun.star.io.TextOutputStream")
descriptorOut.setOutputStream(descriptorPipe)
descriptorOut.setEncoding("UTF-8")
descriptorOut.writeString(r"""
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
descriptorOut.flush()
descriptorOut.closeOutput()
fileAccess.writeFile(libraryDir + "/parcel-descriptor.xml", descriptorPipe)
descriptorPipe.closeInput()

document.storeToURL(odUrl, ())
document.dispose()

macroExecutionModeArg = PropertyValue()
macroExecutionModeArg.Name = "MacroExecutionMode"
macroExecutionModeArg.Value = 4

readOnlyArg = PropertyValue()
readOnlyArg.Name = "ReadOnly"
readOnlyArg.Value = True

document = desktop.loadComponentFromURL(odUrl, "_blank", 0, (macroExecutionModeArg, readOnlyArg, hiddenArg));
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
        encoding = 'utf-8'
        with codecs.open(logPath, encoding=encoding) as f:
            if f.seek(0, 2) == pos:
                continue
            f.seek(pos)
            data = f.read().encode(encoding)
            if sys.version < '3':
                sys.stdout.write(data)
            else:
                sys.stdout.write(data.decode(encoding))
            sys.stdout.flush()
            pos += len(data)

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

# if no error
shutil.rmtree(tempDir)

# Javascript comment terminator */

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
// -*- coding: shift_jis-dos -*-
/**
 * Main
 *
 */

var WINDOWS = (typeof WScript != 'undefined');
var STAR = (typeof XSCRIPTCONTEXT != 'undefined');

if (STAR) {
    var JFile = Packages.java.io.File;
    var FILE_SEPARATOR = JFile.separator;
    importClass(Packages.java.lang.System);
    importClass(Packages.java.io.PrintStream);
    importClass(Packages.com.sun.star.beans.PropertyValue);
    importClass(Packages.com.sun.star.container.XIndexAccess);
    importClass(Packages.com.sun.star.container.XNamed);
    importClass(Packages.com.sun.star.frame.XComponentLoader);
    importClass(Packages.com.sun.star.frame.XModel);
    importClass(Packages.com.sun.star.lang.XComponent);
    importClass(Packages.com.sun.star.sheet.XSpreadsheet);
    importClass(Packages.com.sun.star.sheet.XSpreadsheetDocument);
    importClass(Packages.com.sun.star.sheet.XSpreadsheetView);
    importClass(Packages.com.sun.star.sheet.XSpreadsheets);
    importClass(Packages.com.sun.star.sheet.XViewFreezable);
    importClass(Packages.com.sun.star.sheet.XViewSplitable);
    importClass(Packages.com.sun.star.table.XCell);
    importClass(Packages.com.sun.star.table.XCellRange);
    importClass(Packages.com.sun.star.text.XText);
    importClass(Packages.com.sun.star.ui.dialogs.XFolderPicker);

    System.out.println("StarOffice !");
    function qi(interfaceClass, object) {
        return Packages.com.sun.star.uno.UnoRuntime.queryInterface(interfaceClass, object);
    }
    var desktop = XSCRIPTCONTEXT.getDesktop();
    var componentContext = XSCRIPTCONTEXT.getComponentContext();
    var serviceManager = componentContext.getServiceManager();
    var ExecutableDialogResults = {
        CANCEL: 0,
        OK: 1
    };
} else {
    function qi(interfaceClass, object) {
        return object;
    }

    var FILE_SEPARATOR = "\\";
}

function getStarPath(nativePath) {
    return ("file://" + (WINDOWS ? "/" : "") + nativePath).replace(/\\/, "/");
}

function getNativePath(starPath) {
    var regexp = WINDOWS ? /^file:\/\/\// : /^file:\/\//
    return decodeURI((starPath + "").replace(regexp, ""));
}

// -*- coding: shift_jis-dos -*-

var DECIDION_TABLE_MAX_RIGHT = 200;
var DECIDION_TABLE_MAX_BOTTOM = 200;
var DECIDION_TABLE_MAX_IGNORED_LINES = 20;

var POSITIVE_REGEXP = /[○ＹY]/;
var NEGATIVE_REGEXP = /[×ＮN]/;
var BRACKET_REGEXP = /[“”「」『』【】]/;

function WindowsFilesystem() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var BIF_NONEWFOLDERBUTTON = 512;
    var StreamTypeEnum = {
        adTypeText: 2
    };
    var SaveOptionsEnum = {
        adSaveCreateOverWrite: 2
    };

    this.write = function (path, content) {
        var stream = new ActiveXObject("ADODB.Stream");
        try {
            stream.Open();
            stream.Type = StreamTypeEnum.adTypeText;
            stream.Charset = "UTF-8";
            stream.WriteText(content.replace(/\r?\n/g, "\r\n"));
            stream.SaveToFile(path, SaveOptionsEnum.adSaveCreateOverWrite);
        } finally {
            stream.Close();
        }
    };

    /**
     * 入力フォルダ取得
     */
    this.getInputFolder = function () {
        if (WScript.Arguments.Length > 0) {
            return WScript.Arguments(0);
        }
        var shellApplication = new ActiveXObject("Shell.Application");
        var title = "試験仕様書があるフォルダを選択してください";
        console.log(title);
        var folder = shellApplication.BrowseForFolder(0, title, BIF_NONEWFOLDERBUTTON);
        if (folder == null) {
            console.log("キャンセルされました。");
            application.exit();
        }
        try {
            var path = folder.Items().Item().Path
        } catch (e) {
            console.log("フォルダの取得に失敗しました。");
            throw e;
        }
        return path;
    };

    /**
     * 出力フォルダ取得
     */
    this.getOutputFolder = function () {
        if (WScript.Arguments.Length > 1) {
            return WScript.Arguments(1);
        }
        var outputFolder = fso.GetParentFolderName(WScript.ScriptFullName) + "\\feature"
        return outputFolder;
    };

    /**
     * フォルダーからExcelファイルを探してパスの配列を返す
     */
    this.getSpreadsheetFiles = function (baseFolderPath) {
        if (!fso.FolderExists(baseFolderPath)) {
            return [];
        }
        var baseFolder = fso.GetFolder(baseFolderPath);
        var filePaths = [];
        var files = new Enumerator(baseFolder.Files);
        var progress = 0;
        for (; !files.atEnd(); files.moveNext(), ++progress) {
            var path = files.item().Path;
            if (path.match(/\.(xls|xlsx|xlsm)$/i)
                && !fso.GetBaseName(path).match(/^~\$/)) {
                console.write("*");
                filePaths.push(path);
            }
            if (progress > 10) {
                console.write(".");
                progress = 0;
            }
        }
        var folders = new Enumerator(baseFolder.SubFolders);
        for (; !folders.atEnd(); folders.moveNext(), ++progress) {
            filePaths = filePaths.concat(this.getSpreadsheetFiles(folders.item()));
            if (progress > 10) {
                console.write(".");
                progress = 0;
            }
        }
        return filePaths;
    };

    this.createFolder = function () {
        fso.FolderExists(outputFolder) || fso.CreateFolder(outputFolder);
    };

    this.getBaseName = function (path) {
        return fso.GetBaseName(path);
    };
}

function WindowsConsole() {
    this.write = function (message) {
        WScript.StdOut.Write(message);
    };

    this.log = function (message) {
        WScript.StdOut.Write(message + "\r\n");
    };
}

function WindowsApplication() {
    this.exit = function () {
        WScript.Quit();
    };
}

function StarFilesystem() {
    this.write = function (path, content) {
        var printStream;
        try {
            printStream = new PrintStream(new JFile(path), "UTF-8");
            printStream.write(0xef);
            printStream.write(0xbb);
            printStream.write(0xbf);
            printStream.print(content.replace(/\r?\n/g, "\r\n"));
        } finally {
            if (printStream) {
                printStream.close();
            }
        }
    };

    /**
     * 入力フォルダ取得
     */
    this.getInputFolder = function () {
        if (ARGUMENTS[1]) {
            return "" + ARGUMENTS[1];
        }
        var folderPicker = qi(XFolderPicker, 
            serviceManager.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", componentContext));
        var message = "入力フォルダ取得";
        folderPicker.setTitle(message);
        folderPicker.setDescription(message);
        if (folderPicker.execute() == ExecutableDialogResults.OK) {
            return getNativePath(folderPicker.getDirectory());
        } else {
            console.log("キャンセルされました。");
            application.exit();
        }
    };

    /**
     * 出力フォルダ取得
     */
    this.getOutputFolder = function () {
        if (ARGUMENTS[2]) {
            return "" + ARGUMENTS[2];
        }
        var folderPicker = qi(XFolderPicker, 
            serviceManager.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", componentContext));
        var message = "出力フォルダ取得";
        folderPicker.setTitle(message);
        folderPicker.setDescription(message);
        if (folderPicker.execute() == ExecutableDialogResults.OK) {
            return getNativePath(folderPicker.getDirectory());
        } else {
            console.log("キャンセルされました。");
            application.exit();
        }
    };

    /**
     * フォルダーからExcelファイルを探してパスの配列を返す
     */
    this.getSpreadsheetFiles = function (baseFolderPath) {
        var spreadsheetFiles = [];
        var files = new JFile(baseFolderPath).listFiles() || [];
        for (var i = 0; i < files.length; ++i) {
            var path = files[i].getAbsolutePath();
            if (files[i].isDirectory()) {
                spreadsheetFiles = spreadsheetFiles.concat(this.getSpreadsheetFiles(path));
            } else {
                if (path.match(/\.(xls|xlsx|ods)$/i)) {
                    System.out.println(path);
                    spreadsheetFiles.push(path);
                }
            }
        }
        return spreadsheetFiles;
    };

    this.createFolder = function (path) {
        (new JFile(path)).mkdirs();
    };

    this.getBaseName = function (path) {
        return ((new JFile(path)).getName() + "").replace(/\..*$/i, "");
    };
}

function StarConsole() {
    this.write = function (message) {
        System.out.print("" + message);
    };

    this.log = function (message) {
        System.out.println("" + message);
    };
}

function StarApplication() {
    this.exit = function () {
        throw new Error("Exit!");
    };
}

var console;
var filesystem;
var application;

if (STAR) {
    console = new StarConsole();
    filesystem = new StarFilesystem();
    application = new StarApplication();
} else {
    console = new WindowsConsole();
    filesystem = new WindowsFilesystem();
    application = new WindowsApplication();
}

/**
 * LibreOffice Calc
 */
function StarBook(path) {
    var readOnly = new PropertyValue();
    readOnly.Name = "ReadOnly";
    readOnly.Value = true;

    var hidden = new PropertyValue();
    hidden.Name = "Hidden";
    hidden.Value = true;

    var properties = [
        readOnly,
        hidden
    ];

    var book = this;
    var url = getStarPath(path);
    var starBook = qi(XSpreadsheetDocument, qi(XComponentLoader, desktop).loadComponentFromURL(url, "_blank", 0, properties));

    function Sheet(starSheet) {
        function Cell(starCell) {
            this.getValue = function () {
                return qi(XText, starCell).getString() + "";
            };

            this.activate = function() {
                // qi(XXX, qi(XModel, starBook).getCurrentController()).select(starCell);
            };
        };

        this.getCell = function (x, y) {
            return new Cell(qi(XCell, starSheet.getCellByPosition(y - 1, x - 1)));
        };

        this.getName = function () {
            return qi(XNamed, starSheet).getName();
        };
        
        this.activate = function () {
            qi(XSpreadsheetView, qi(XModel, starBook).getCurrentController()).setActiveSheet(starSheet);
        };
        
        this.getTableTop = function () {
            var currentController = qi(XModel, starBook).getCurrentController();
            return qi(XViewFreezable, currentController).hasFrozenPanes() ?
                qi(XViewSplitable, currentController).getSplitRow() : 0;
        };

        this.getTableLeft = function () {
            var currentController = qi(XModel, starBook).getCurrentController();
            return qi(XViewFreezable, currentController).hasFrozenPanes() ?
                qi(XViewSplitable, currentController).getSplitColumn() : 0;
        };

        this.getBook = function() {
            return book;
        };
    }

    this.dispose = function() {
        qi(XModel, starBook).dispose();
    };

    this.getSheets = function() {
        var results = [];
        var sheets = qi(XIndexAccess, starBook.getSheets());
        for (var i = 0; i < sheets.getCount(); ++i) {
            results.push(new Sheet(qi(XSpreadsheet, sheets.getByIndex(i))));
        }
        return results;
    };

    this.getBaseName = function() {
        return filesystem.getBaseName(path);
    };
};

/**
 * Excel
 */
function ExcelBook(path) {
    var excelApplication;
    var excelBook;
    var book = this;
    
    function Sheet(excelSheet) {
        var sheet = this;

        function Cell(excelCell) {
            this.getValue = function () {
                return (excelCell.Value || "") + "";
            };

            this.activate = function() {
            };

            this.getSheet = function() {
                return sheet;
            };
        };
        
        this.getCell = function (x, y) {
            return new Cell(excelSheet.Cells(x, y));
        };
        
        this.getName = function () {
            return excelSheet.Name;
        };
        
        this.activate = function () {
            excelSheet.Activate();
        };
        
        this.getTableTop = function () {
            return excelApplication.ActiveWindow.SplitRow + 1
        };

        this.getTableLeft = function () {
            return excelApplication.ActiveWindow.SplitColumn + 1;
        };

        this.getBook = function() {
            return book;
        };
    }

    this.dispose = function() {
        excelBook.Close();
        excelApplication.Quit();
    };

    this.getSheets = function() {
        var sheets = [];
        var enumerator = new Enumerator(excelBook.WorkSheets);
        for (; !enumerator.atEnd(); enumerator.moveNext()) {
            sheets.push(new Sheet(enumerator.item()));
        }
        return sheets;
    };

    this.getBaseName = function() {
        return filesystem.getBaseName(path);
    };

    try {
        excelApplication = new ActiveXObject("Excel.Application");
    } catch (e) {
        console.log("Excelの実行に失敗しました");
        application.exit();
    }

    excelApplication.Visible = false
    excelApplication.DisplayAlerts = false;
    excelBook = excelApplication.Workbooks.Open(path);
};

/**
 * 二値のセルからステップを取得
 */
function GetStepFromBooleanCell(command, condition) {
    if (condition.match(POSITIVE_REGEXP)) {
        console.write("o");
        return " * " + NormalizeStep(command) + "\n";
    } else {
        console.write(".");
        return "";
    }
};

/**
 * テキストのセルからステップを取得
 */
function GetStepFromTextCell(command, condition) {
    console.write("*");
    return " * " + NormalizeStep(command) + " \""
        + NormalizeStep(condition).replace(/\"/g /* escape &quot; for syntax highlighting */, "\\\"")
        + "\"\n";
}

/**
 * ステップの文字を整形
 */
function NormalizeStep(step) {
    var rules = [
        [/^[　 \t]+/, ""],
        [/[　 \t]+$/, ""],
        [/　/g, " "],
        [new RegExp(BRACKET_REGEXP.source, "g"), "\""],
        [/(\r\n|\r|\n)+/g, "/"],
    ];
    for (var i = 0; i < rules.length && rules[i]; ++i) {
        step = ((step || "") + "").replace(rules[i][0], rules[i][1]);
    }
    return step;
}

/**
 * ワークシートの列からCucumberのシナリオデータを出力する
 */
function CreateScenarioFromWorkSheetColumn(sheet, topRow, bottomRow, commandColumn, conditionColumn, getStepFunctions) {
    var excelColumnName = GetExcelColumnName(conditionColumn);
    var range = excelColumnName + topRow + ":" + excelColumnName + bottomRow;
    var scenario = "シナリオ: " + sheet.getBook().getBaseName() + "_" + sheet.getName() + "_"  + ("0000" + (conditionColumn - commandColumn)).slice(-5) + "_" + range + "\n";
    for (var row = topRow; row <= bottomRow; ++row) {
        var condition = sheet.getCell(row, conditionColumn).getValue();
        var command = sheet.getCell(row, commandColumn).getValue();
        if (command.length == 0) {
            continue;
        }
        scenario += getStepFunctions[row](command, condition);
    }
    console.write("\n");
    return scenario + "\n";
}

/**
 * ワークシートからCucumberのfeatureデータを出力する
 */
function CreateFeatureFromWorkSheet(sheet, featureName) {
    console.log(sheet.getName());
    sheet.activate();
    sheet.getCell(1, 1).activate();

    // Freezeされているセルをデシジョンテーブルの左上とする
    var top = sheet.getTableTop();
    var left = sheet.getTableLeft();
    if (top < 2 || left < 2) {
        return;
    }

    // 右限を検索
    var right = 0;
    var ignoredXLines = 0;
    for (var x = left; x < DECIDION_TABLE_MAX_RIGHT; ++x) {
        var valueFound = false;
        for (var y = top; y < DECIDION_TABLE_MAX_BOTTOM; ++y) {
            if (sheet.getCell(y, x).getValue().length > 0) {
                valueFound = true;
                break;
            }
        }
        if (valueFound) {
            ignoredXLines = 0;
            right = x;
        } else if (++ignoredXLines > DECIDION_TABLE_MAX_IGNORED_LINES) {
            break;
        }
    }
    if (!right) {
        return;
    }

    // 下限を検索
    var bottom = 0;
    var ignoredYLines = 0;
    for (var y = top; y < DECIDION_TABLE_MAX_RIGHT; ++y) {
        var valueFound = false;
        for (var x = left; x <= right; ++x) {
            if (sheet.getCell(y, x).getValue().length > 0) {
                valueFound = true;
                break;
            }
        }
        if (valueFound) {
            ignoredYLines = 0;
            bottom = y;
        } else if (++ignoredYLines > DECIDION_TABLE_MAX_IGNORED_LINES) {
            break;
        }
    }
    if (!bottom) {
        return;
    }
    console.log("デシジョンテーブルの範囲: (top=" + top + ", left=" + left + ") - (bottom=" + bottom + ", right=" + right + ")");

    var getStepFunctions = new Array(bottom + 1);
    for (var y = top; y <= bottom; ++y) {
        var textFound = false;
        for (var x = left; x <= right; ++x) {
            var value = sheet.getCell(y, x).getValue();
            if (!value.match(POSITIVE_REGEXP) && !value.match(NEGATIVE_REGEXP) && value.length > 0) {
                textFound = true;
                break;
            }
        }
        getStepFunctions[y] = textFound ? GetStepFromTextCell : GetStepFromBooleanCell;
    }
    var feature = "フィーチャ: " + featureName + "\n\n";
    for (var x = left; x <= right; ++x) {
        feature += CreateScenarioFromWorkSheetColumn(sheet, top, bottom, left - 1, x, getStepFunctions);
    }
    console.log("OK");
    return feature;
}

/**
 * ExcelファイルからCucumberのfeatureファイルを出力する
 */
function CreateFeature(path, outputFolder) {
    console.log("□ " + path);
    filesystem.createFolder(outputFolder);
    var book;
    try {
        if (WINDOWS) {
            try {
                book = new ExcelBook(path);
            } catch (e) {
                book = new StarBook(path);
            }
        } else {
            book = new StarBook(path);
        }
    } catch (e) {
        // TODO
        // return;
        throw e;
    }

    try {
        var sheets = book.getSheets();
        for (var i = 0; i < sheets.length; ++i) {
            var sheet = sheets[i];
            var featureName = book.getBaseName() + " " + sheet.getName();
            var feature = CreateFeatureFromWorkSheet(sheet, featureName);
            if (!feature) {
                console.log("  skip")
                continue;
            }
            var outputPath = outputFolder + FILE_SEPARATOR + filesystem.getBaseName(path) + "_" + sheet.getName() + ".feature";
            filesystem.write(outputPath, feature);
        }
    } finally {
        try {
            book.dispose();
        } catch (e) {
            // TODO
            throw e;
        }
    }
}

/**
 * Excelのアルファベットのカラム名を取得
 */
function GetExcelColumnName(Index) {
    var letters = 26;
    var excelColumnName = (Index - 1).toString(letters);
    var result = "";
    for (var i = 0; i < excelColumnName.length; ++i) {
        var offset = (i == 0 && excelColumnName.length != 1) ? 1 : 0;
        result += String.fromCharCode(
            "A".charCodeAt(0) + parseInt(excelColumnName.charAt(i), letters) - offset);
    }
    return result;
}

/**
 * http://stackoverflow.com/a/20260831
 */
function objToString(obj, level)
{
    if (level > 10) {
        return "!!! level too deep"
    }
    var out = '';
    for (var i in obj) {
        for (loop = level; loop > 0; loop--) {
            out += "  ";
        }
        if (obj[i] instanceof Object) {
            out += i + " (Object):\n";
            out += objToString(obj[i], level + 1);
        } else {
            out += i + ": " + obj[i] + "\n";
        }
    }
    return out;
}

try {
    var inputFolder = filesystem.getInputFolder();
    var outputFolder = filesystem.getOutputFolder();

    console.log(inputFolder + "から試験仕様書を検索しています");
    var message = "\n";
    var filePaths = filesystem.getSpreadsheetFiles(inputFolder);
    for (var i = 0; i < filePaths.length; ++i) {
        message += filePaths[i] + "\n";
    }
    message += filePaths.length + "件見つかりました"
    console.log(message);

    for (var i = 0; i < filePaths.length; ++i) {
        CreateFeature(filePaths[i], outputFolder);
    }
} catch (e) {
    console.log(e);
    console.log(objToString(e, 1));
    if (typeof e.rhinoException != 'undefined') {
        e.rhinoException.printStackTrace();
    } else if (typeof e.javaException != 'undefined') {
        e.javaException.printStackTrace();
    }

    if (typeof e.stack != 'undefined') {
        console.log(e.stack);
    }

    throw e;
}

"""
#" /* magic comment for editor's syntax highlighting

odBase64 = """
UEsDBBQACAgIAKpm3UIAAAAAAAAAAAAAAAAnAAAAQ29uZmlndXJhdGlvbnMyL2FjY2VsZXJhdG9y
L2N1cnJlbnQueG1sAwBQSwcIAAAAAAIAAAAAAAAAUEsDBBQAAAgAAKpm3UIAAAAAAAAAAAAAAAAY
AAAAQ29uZmlndXJhdGlvbnMyL2Zsb2F0ZXIvUEsDBBQAAAgAAKpm3UIAAAAAAAAAAAAAAAAfAAAA
Q29uZmlndXJhdGlvbnMyL2ltYWdlcy9CaXRtYXBzL1BLAwQUAAAIAACqZt1CAAAAAAAAAAAAAAAA
GAAAAENvbmZpZ3VyYXRpb25zMi9tZW51YmFyL1BLAwQUAAAIAACqZt1CAAAAAAAAAAAAAAAAGgAA
AENvbmZpZ3VyYXRpb25zMi9wb3B1cG1lbnUvUEsDBBQAAAgAAKpm3UIAAAAAAAAAAAAAAAAcAAAA
Q29uZmlndXJhdGlvbnMyL3Byb2dyZXNzYmFyL1BLAwQUAAAIAACqZt1CAAAAAAAAAAAAAAAAGgAA
AENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIvUEsDBBQAAAgAAKpm3UIAAAAAAAAAAAAAAAAYAAAA
Q29uZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsDBBQAAAgAAKpm3UIAAAAAAAAAAAAAAAAaAAAAQ29u
ZmlndXJhdGlvbnMyL3Rvb2xwYW5lbC9QSwMEFAAICAgAqmbdQgAAAAAAAAAAAAAAAAsAAABjb250
ZW50LnhtbO1azW7jNhC+9ylcLXqUZctxagux97a9JEDQpECvtETJRChRoOi/vRVoj8XupS36Br30
BYr6ZXrYq1+hQ1KiaTtylMZou8bmEFsz3wxH3ww5JJOr18uUtuaYF4RlI6fb7jgtnIUsIlkycr65
f+MOnNfjz65YHJMQBxELZynOhBuyTMBnC6yzItDakTPjWcBQQYogQykuAhEGLMdZZRXY6ECNpSWF
WNHG5gpsWwu8FE2NJXbHFk2aj6zAtnXE0aKpscQCqbZ5zJoaLwvqxgxYT3MkyF4US0qyh5EzFSIP
PG+xWLQXvTbjidcdDoee0pqAQ4PLZ5wqVBR6mGI5WOF1212vwqZYoKbxSawdUjZLJ5g3pgYJdJDV
nOMCIPC6sjCbObJtduprnjSurnlSQ3M4RbxxnSnwbqn0oual0ots2xSJaU1+B94NKNWvm+ttXfG0
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
EgATBQAA2R8AAAAA
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
                 , "--invisible"
                 #, "--minimized"
                 #, "--headless"                 
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

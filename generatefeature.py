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
// -*- coding: shift_jis-dos -*-
/**
 * Main
 *
 */

var WINDOWS = (typeof(WScript) != 'undefined');
var STAR = (typeof(XSCRIPTCONTEXT) != 'undefined');

if (STAR) {
    importClass(Packages.java.lang.System);
    importClass(Packages.java.io.File);
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

    function File() {
        this.separator = "\\";
    };
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
            filePaths = filePaths.concat(GetExcelFiles(folders.item()));
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
            printStream = new PrintStream(new File(path), "UTF-8");
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
        // return "/home/i_mogi"
        var folderPicker = qi(XFolderPicker, 
            serviceManager.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", componentContext));
        if (folderPicker.execute() == ExecutableDialogResults.OK) {
            return getNativePath(folderPicker.getDirectory());
        } else {
            foo.bar();
        }
    };

    /**
     * 出力フォルダ取得
     */
    this.getOutputFolder = function () {
        if (ARGUMENTS[2]) {
            return "" + ARGUMENTS[2];
        }
        return "/home/i_mogi";
    };

    /**
     * フォルダーからExcelファイルを探してパスの配列を返す
     */
    this.getSpreadsheetFiles = function (baseFolderPath) {
        var spreadsheetFiles = [];
        var files = new File(baseFolderPath).listFiles() || [];
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
        (new File(path)).mkdirs();
    };

    this.getBaseName = function (path) {
        return ((new File(path)).getName() + "").replace(/\..*$/i, "");
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
};

/**
 * Excel
 */
function ExcelBook(path) {
    var excelApplication;
    var excelBook;
    
    function Sheet(excelSheet) {
        function Cell(excelCell) {
            this.getValue = function () {
                return (excelCell.Value || "") + "";
            };

            this.activate = function() {
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
    var scenario = "シナリオ: No." + (conditionColumn - commandColumn) + " " + range + "\n";
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
    console.log(10);
    console.log(sheet.getName());
    console.log("1x");
    sheet.activate();
    console.log(11);
    sheet.getCell(1, 1).activate();
    console.log(12);

    // Freezeされているセルをデシジョンテーブルの左上とする
    var top = sheet.getTableTop();
    var left = sheet.getTableLeft();
    if (top < 2 || left < 2) {
        return;
    }
    console.log(13);

    // 右限を検索
    var right = 0;
    var ignoredXLines = 0;
    for (var x = left; x < DECIDION_TABLE_MAX_RIGHT; ++x) {
        var valueFound = false;
        for (var y = top; y < DECIDION_TABLE_MAX_BOTTOM; ++y) {
            console.log(20);
            if (sheet.getCell(y, x).getValue().length > 0) {
                console.log(22);
                valueFound = true;
                break;
            }
            console.log(21);
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
    console.log(14);

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
    console.log(15);

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
    var feature = "";
    feature += "# language: ja\n\n"; // 手元のSpecFlowで言語指定の引数が使えないため
    feature += "フィーチャ: " + featureName + "\n\n";
    for (var x = left; x <= right; ++x) {
        feature += CreateScenarioFromWorkSheetColumn(sheet, top, bottom, left - 1, x, getStepFunctions);
    }
    console.log(16);
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
    }

    try {
        console.log(1);
        var sheets = book.getSheets();
        for (var i = 0; i < sheets.length; ++i) {
            var sheet = sheets[i];
            console.log(2);
            var featureName = filesystem.getBaseName(path) + " " + sheet.getName();
            console.log(3);
            var feature = CreateFeatureFromWorkSheet(sheet, featureName);
            console.log(4);
            if (!feature) {
                console.log("  skip")
                continue;
            }
            console.log(5);
            var outputPath = outputFolder + File.separator + filesystem.getBaseName(path) + "_" + sheet.getName() + ".feature";
            console.log(6);
            console.log(outputPath);
            console.log(7);
            filesystem.write(outputPath, feature);
            console.log(8);
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

"""
#" /* magic comment for editor's syntax highlighting

odBase64 = """
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
etqqaDm+Y3wccHP3hz+JBGkQAKRBAJAGAUAaBABpEACkQQCQBgFAGgQAaRAApEEAkAYBQBoEAGkQ
AKRBAJAGAUAaBABpEACkQQCQBgFAGgQAaRAApEEAkAYBQBoEAGkQAKRBAJAGAUAaBABpEACkQQCQ
BgFAGgQAaRAApEEAkAYBQBoEAGkQAKRBAJAGAUAaBABpEACkQQCQBgFAGgQAaRAApEEAkAYBQBoE
AGkQAKRBAJAGAUAaBABpEACkQQCQBgFAGgQAaRAApEEAkAYBQBoEAGkQAKRBAJAGAUAaBABpEACk
QQCQBgFAGgQAaRAApEEAkAYBQBoEAGkQAKRBAJAGAUAaBABpEACkQQCQBgFAGgQAaRAApEEAkAYB
QBoEAGkQAKRBAJAGAUAaBABpEACkQQCQBgFAGgQAaRAApEEAkAYBQBoEAGkQAKRBAJAGAUAaBABp
EACkQQCQBgFAGgQAaRAApPkfF/g0uJl/kuYAAAAASUVORK5CYIJQSwMEFAAACAAAUWXdQgAAAAAA
AAAAAAAAABoAAABDb25maWd1cmF0aW9uczIvcG9wdXBtZW51L1BLAwQUAAAIAABRZd1CAAAAAAAA
AAAAAAAAHwAAAENvbmZpZ3VyYXRpb25zMi9pbWFnZXMvQml0bWFwcy9QSwMEFAAACAAAUWXdQgAA
AAAAAAAAAAAAABoAAABDb25maWd1cmF0aW9uczIvdG9vbHBhbmVsL1BLAwQUAAAIAABRZd1CAAAA
AAAAAAAAAAAAGgAAAENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIvUEsDBBQAAAgAAFFl3UIAAAAA
AAAAAAAAAAAYAAAAQ29uZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsDBBQAAAgAAFFl3UIAAAAAAAAA
AAAAAAAcAAAAQ29uZmlndXJhdGlvbnMyL3Byb2dyZXNzYmFyL1BLAwQUAAAIAABRZd1CAAAAAAAA
AAAAAAAAGAAAAENvbmZpZ3VyYXRpb25zMi9tZW51YmFyL1BLAwQUAAAIAABRZd1CAAAAAAAAAAAA
AAAAGAAAAENvbmZpZ3VyYXRpb25zMi9mbG9hdGVyL1BLAwQUAAgICABRZd1CAAAAAAAAAAAAAAAA
JwAAAENvbmZpZ3VyYXRpb25zMi9hY2NlbGVyYXRvci9jdXJyZW50LnhtbAMAUEsHCAAAAAACAAAA
AAAAAFBLAwQUAAgICABRZd1CAAAAAAAAAAAAAAAACgAAAHN0eWxlcy54bWzdXM2O5LYRvucpGjKc
m0ZS/22rszOGcwiQIGMY3vUDUBKlplcSFYqant7j+p7cEhtI7rkECZAAAYJ9miyw132FFElJLXVL
PaKnZzSZacDYZhXFqo8fi8Wi3C+/uE3iyQ1mOaHppeFc2MYEpz4NSBpdGt++/pW5Mr64+tlLGobE
x+uA+kWCU27mfBfjfAKd03ythJdGwdI1RTnJ1ylKcL7m/ppmOK06rZvaazmUapEPG9pdKjd7c3zL
h3YWuq2+yBs+slRu9g4Y2g7tLHQB02b3kA7tfJvHZkhNnyYZ4uTAituYpG8ujQ3n2dqyttvtxXZ2
QVlkOa7rWlJaG+zXelnBYqkV+BaOsRgst5wLx6p0E8zRUPuEbtOktEg8zAZDgzg6mtWM4RxUwF3B
y2EPavZp8esmGsyum6gHZn+D2GCeSeU2VWbBcKrMgmbfBPFNz/yurGsQyv9c/3bPK5YMHUvotqDy
GckGu6m0m/0ppbWpooNa7NLcqW3PLfW9ob09qb5lhGPWUPdPqvso9mvEadIFGug5FmiY+EZQvvY7
IfFgr0G3hyQoJYOhF7pHVGUC/F4PFxbDGWW8BiQcHnRhlGkdMjY8iftDhpBWqhELgk5VMGdmQfiA
xWveELz9zGjtBqeJ4B4QQYbWu7pIpWbsPdnBsS2hUy9foMZ+o2BRvZWFtEgDNQ8KQHybYUaECMWy
27r1hFZMyPMZ7wLn9TeWkJliv4GIWm55jW12alxVe6raSq9eCmthrbM3mE3kv4U5l8aXjFHwAgLT
WsD8S3p7adgTezK1JzNbtUNkSRxoMx3Rtpnabw3r6qUKqQEOURGXG/ZEtYUIGLy7NCKGsg3xjUq3
/G5mDNBknMAGLx6fc0bfYGB8TCGmfzabLxcoNJSNIYnjWvJi6oY+SEK63sKjTJqp6J1SU3zfG5Uh
huRgraGkSCBmooLTPEMip0hpCnNeditSnxdyruQDL42cJFlcy2Grw6bHMILtEIwmPq8kIo7A7msm
NIBnxszkXiUKKaQzJA2wWDQitZFPEaPLFChEcY5rhICCgCTNcnDF6velVhfOHLlY5NiEXCCgW1MO
XsLHWYEldrKxmqOfo4zmv3gNfMwnX+Ht5BuaoFQ1tjxQ+maEUyAvREkm9FoaGeE+7CM3iBG1kKqh
cvIWQJnOMy7bYpRGBYqgCaeywYcVwhnY8u2rriEh9qC0MvTT+799ev/Pyaf3//j4w+8//vkvd1pa
dc93OcfJscGVfG92Q0NYXiko+0silC5Usu9QJSl9qQS/+brLNBHbYwzr7DXaAIynrK9Ve+2vNfo9
qFU6failG3LoRS369VfGnpGtJV9Rsbn+VWDJOYLAxwJDPyrIpa/CAjyHxiQwmpFiSwKRsdh+YtwZ
QFTAA1shYao7XkxF12O5D+tT5JVqUbYUYPGe6C6kXZ1F+Ko96A9oUiIWM4LgYm4oI2+p2B5MFJNI
8KvIOQl3crVkKBCHJxOihDDFmS6EMQ2BRzkXuUmXLMYhlw4cChiJNg2JmoANCmT+RIJALNVGowl7
S465edtGoy3cdQor31e2+AAL5JkpJnl7DzngUFMthtQqVsqmV8Qx5hMlFO2w+RnqqxKZIlW+NP77
pz/UhGs8pME52SchqRkjDyT1XC/BAasv4B5G01cwg692iUdjQycYqwg5X3wuV9kpTzVwmN4DB7k5
mh6G9B1XIDw/hGZnQ8hRPH92CM3PiNDqWSK0OBtC04v5s0RoeTaEZs8Snxfnw+eZxunV2RCaP9M4
7Z4RoecZpx37bBAt/r8CdUNcntisvoPXyfoJjA2nnUg4rE4R5RGibKxOEO1WdUppt9UHlLJZFkQ2
WPV3bPtz2SqRI7JkolR7MR2AW6MjLbgYsD6nNUTSEL5htIg2ZnlppOpDh9P2JRy14/PWR5xVWR9R
bdXoLBEDVXhUR7Jm1UoKCgCKqcpSl93bEt3qeSXjMYcTqwkH2FTWoppYDajAfHj3rw/v/v3h++8/
vPv749VhFE5NoWR7KW072PD9hMJhVeYapdHJ2T1zVabHo1re71OPimQETrINUtRqskV2ZzgmOOyg
Eb3pYlGXQrm89kutUSC6ozBEve+wz7eEQwhWtefO+lBVlEWsvhA2j6oC9y0hXTiLniqSLf86qkR1
yfxEfenF6fqSWmi95aWZina6aKrg8JBwltHnhuREUvpBKkI/wXOI6aKu9jCu38UNVeOTK2Ww6WIt
PQLv1fK9l5UeDXZjWno6jSq3ziUEUG2vRFE1JLi3Pv143vWlW+quql0C1nZTpVBP10mNnLKdElaJ
9kA4CG9swk+UyPO5FpGFS84YPvVeaNgqhD/ydjF8FakNeOh0TPWnY/rge1CdJdj7m5rWHISh77vu
05qDIYtcjXJ4dLxweg6PtaBjfvXOj+1Jn2ltJhuMApEPjx9Z7oC+RHM+7YNz6gzHSG9hVBiNEqrG
A6l9oK9O4B6NA33oHiisPFno5h21EMJhkfv3RTTBKC/YY6zX40PnWY6WXe8edPbreWOhMbdVSIeA
XqSEl+fSoXM01YkB5Tsjh7DLV+7gWAxnt7tgvSPz8OEvHPAqhdiCDrVuxDj+Xge43Pmqxaz/TYtD
0cAXLc6ZHT+tiuujVUyV/vMoZT6JSuTgNQ183R3t6Y0V3R9Lq2hwzzWPl+KjZ/DRTvqYBvvyT8/g
2ZgGezPxGW6wtx2VEKH80zF3VDro8tfbjkqGMm0ZXp9lEL6Os/7HJYQ8C2uaPCopwtB1lxqkUCaP
SowwXC5nGlGCF+x3BcnHpcb8hbfyVvpGj0qO2QzBR9/oUenhuCu0QhphLi7GpYbryrOElsGj0sK2
XVfX4JG3kuVSJxHKi3T0TFMHYLB39ERT095R+eC6egTGiPHN6KmmDoOlxSMnFnq5kLR4VFb4vl5e
ETGMx40TuqyQFo/KCnfuBQuNdwKkxaOyYuGvFlONBDnHmEOKPCovRDTWisfK5lGZoZsglzaPyg3d
/DgWJdDR803dskVt9ejVi59k9eiZhk6E3oFVdPyqlr7JI+caooyha/LIRYzu9Mg6+BmF8qv4CQHx
8xq+WQkqAyNsxmhHC97y8Otr2+jQOXGZ2vPiw0GzuhlyOm+GnPruKarfG5m66uXUqrm6BZLXsxXI
jADGlJH9D+TEKA1yH2XNTKHhRv/MXgfZ0eIR4Iu7MPGAGpRm49Hcech/EzHxSxrl3Y9HWSBe9Ljz
3Uard75KQYJyLm9Zd/tfyQCnRJP4vyf239VYZSyTXovRuuV7e4foqBdZ835Vn6ac0fiERnmtLS72
pJZ16IcCpXRWYNyapY9//fHjf/64X2R7fu7ZW91u71eenFtrj/EBlFb3j3pd/Q9QSwcIOD6Bsc8J
AAAUTAAAUEsDBBQACAgIAFFl3UIAAAAAAAAAAAAAAAAVAAAATUVUQS1JTkYvbWFuaWZlc3QueG1s
tZTPbsIwDMbvPEWVe5ON01RROExil93GHsCkpgSlTpQ4iL79Wib+TBOICrg5sfP9PiuWJ7NdY7Mt
hmgcleJVvogMSbvKUF2K78U8fxOz6WjSAJkVRi4OQda9o3g8liIFKhxEEwuCBmPBunAeqXI6NUhc
/K0v9qTj6czAWExH2Ym3Mhbz7n1oT9WrZG3ugdelUJdETtcNVgZybj2WAry3RgN3ZWpLldwbluc+
ZR3Ar42OQg3x0SCD7Fq8wGXcserTg0QjMnf/EB8urB1x3+ujdb90MJ6j2sAW4j5Wn2YZILTqAwkD
MM4ROAWUGy838QL8YVAPQaPNK/y9d+FKxw+D3jB6S0Nd5d3EJ5OeJL9Yp2ZJYGxUfAilp/oCzjRQ
o+rzgyjvjlamTmHvM44V6G4S+gl0QekUwvXxv4914/KJiXoLMhmpzxUGLgluLd66Iibq3yKf/gBQ
SwcIyUvNGFUBAAADBgAAUEsBAhQAFAAACAAAUWXdQp8DLsQrAAAAKwAAAAgAAAAAAAAAAAAAAAAA
AAAAAG1pbWV0eXBlUEsBAhQAFAAICAgAUWXdQmwvwlXBAQAANQQAAAgAAAAAAAAAAAAAAAAAUQAA
AG1ldGEueG1sUEsBAhQAFAAICAgAUWXdQi1IKa1iBgAAKyYAAAwAAAAAAAAAAAAAAAAASAIAAHNl
dHRpbmdzLnhtbFBLAQIUABQACAgIAFFl3ULXoFT/8AUAAF0lAAALAAAAAAAAAAAAAAAAAOQIAABj
b250ZW50LnhtbFBLAQIUABQACAgIAFFl3UIAAAAAAgAAAAAAAAAwAAAAAAAAAAAAAAAAAA0PAABT
Y3JpcHRzL2phdmFzY3JpcHQvTGlicmFyeS9HZW5lcmF0ZUZlYXR1cmUuanAuanNQSwECFAAUAAgI
CABRZd1CdDwvwr8AAACsAQAAMAAAAAAAAAAAAAAAAABtDwAAU2NyaXB0cy9qYXZhc2NyaXB0L0xp
YnJhcnkvcGFyY2VsLWRlc2NyaXB0b3IueG1sUEsBAhQAFAAACAAAUWXdQgJTYYasCQAArAkAABgA
AAAAAAAAAAAAAAAAihAAAFRodW1ibmFpbHMvdGh1bWJuYWlsLnBuZ1BLAQIUABQAAAgAAFFl3UIA
AAAAAAAAAAAAAAAaAAAAAAAAAAAAAAAAAGwaAABDb25maWd1cmF0aW9uczIvcG9wdXBtZW51L1BL
AQIUABQAAAgAAFFl3UIAAAAAAAAAAAAAAAAfAAAAAAAAAAAAAAAAAKQaAABDb25maWd1cmF0aW9u
czIvaW1hZ2VzL0JpdG1hcHMvUEsBAhQAFAAACAAAUWXdQgAAAAAAAAAAAAAAABoAAAAAAAAAAAAA
AAAA4RoAAENvbmZpZ3VyYXRpb25zMi90b29scGFuZWwvUEsBAhQAFAAACAAAUWXdQgAAAAAAAAAA
AAAAABoAAAAAAAAAAAAAAAAAGRsAAENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIvUEsBAhQAFAAA
CAAAUWXdQgAAAAAAAAAAAAAAABgAAAAAAAAAAAAAAAAAURsAAENvbmZpZ3VyYXRpb25zMi90b29s
YmFyL1BLAQIUABQAAAgAAFFl3UIAAAAAAAAAAAAAAAAcAAAAAAAAAAAAAAAAAIcbAABDb25maWd1
cmF0aW9uczIvcHJvZ3Jlc3NiYXIvUEsBAhQAFAAACAAAUWXdQgAAAAAAAAAAAAAAABgAAAAAAAAA
AAAAAAAAwRsAAENvbmZpZ3VyYXRpb25zMi9tZW51YmFyL1BLAQIUABQAAAgAAFFl3UIAAAAAAAAA
AAAAAAAYAAAAAAAAAAAAAAAAAPcbAABDb25maWd1cmF0aW9uczIvZmxvYXRlci9QSwECFAAUAAgI
CABRZd1CAAAAAAIAAAAAAAAAJwAAAAAAAAAAAAAAAAAtHAAAQ29uZmlndXJhdGlvbnMyL2FjY2Vs
ZXJhdG9yL2N1cnJlbnQueG1sUEsBAhQAFAAICAgAUWXdQjg+gbHPCQAAFEwAAAoAAAAAAAAAAAAA
AAAAhBwAAHN0eWxlcy54bWxQSwECFAAUAAgICABRZd1CyUvNGFUBAAADBgAAFQAAAAAAAAAAAAAA
AACLJgAATUVUQS1JTkYvbWFuaWZlc3QueG1sUEsFBgAAAAASABIA8gQAACMoAAAAAA==
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

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
VZ+mnNH4hEZ5rS0u9qSWdeiHAqV0VmDcmqWPf/3x43/+uF9ke37u2Vvdbu9Xnpxba4/xAZRW9496
Xf0PUEsHCDg+gbHPCQAAFEwAAFBLAwQUAAgICACqZt1CAAAAAAAAAAAAAAAAJwAAAENvbmZpZ3Vy
YXRpb25zMi9hY2NlbGVyYXRvci9jdXJyZW50LnhtbAMAUEsHCAAAAAACAAAAAAAAAFBLAwQUAAAI
AACqZt1CAAAAAAAAAAAAAAAAGgAAAENvbmZpZ3VyYXRpb25zMi90b29scGFuZWwvUEsDBBQAAAgA
AKpm3UIAAAAAAAAAAAAAAAAYAAAAQ29uZmlndXJhdGlvbnMyL2Zsb2F0ZXIvUEsDBBQAAAgAAKpm
3UIAAAAAAAAAAAAAAAAYAAAAQ29uZmlndXJhdGlvbnMyL21lbnViYXIvUEsDBBQAAAgAAKpm3UIA
AAAAAAAAAAAAAAAYAAAAQ29uZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsDBBQAAAgAAKpm3UIAAAAA
AAAAAAAAAAAcAAAAQ29uZmlndXJhdGlvbnMyL3Byb2dyZXNzYmFyL1BLAwQUAAAIAACqZt1CAAAA
AAAAAAAAAAAAGgAAAENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIvUEsDBBQAAAgAAKpm3UIAAAAA
AAAAAAAAAAAaAAAAQ29uZmlndXJhdGlvbnMyL3BvcHVwbWVudS9QSwMEFAAACAAAqmbdQgAAAAAA
AAAAAAAAAB8AAABDb25maWd1cmF0aW9uczIvaW1hZ2VzL0JpdG1hcHMvUEsDBBQACAgIAKpm3UIA
AAAAAAAAAAAAAAALAAAAY29udGVudC54bWztWs1u4zYQvvcpXC16lGXLcWoLsfe2vSRA0KRAr7RE
yUQoUaDov70VaI/F7qUt+ga99AWK+mV62KtfoUNSomk7cpTGaLvG5hBbM98MR98MOSSTq9fLlLbm
mBeEZSOn2+44LZyFLCJZMnK+uX/jDpzX48+uWByTEAcRC2cpzoQbskzAZwussyLQ2pEz41nAUEGK
IEMpLgIRBizHWWUV2OhAjaUlhVjRxuYKbFsLvBRNjSV2xxZNmo+swLZ1xNGiqbHEAqm2ecyaGi8L
6sYMWE9zJMheFEtKsoeRMxUiDzxvsVi0F70244nXHQ6HntKagEODy2ecKlQUephiOVjhddtdr8Km
WKCm8UmsHVI2SyeYN6YGCXSQ1ZzjAiDwurIwmzmybXbqa540rq55UkNzOEW8cZ0p8G6p9KLmpdKL
bNsUiWlNfgfeDSjVr5vrbV3xtOlYErtDVchJ3vg1Ndq2Z4yZUKWBnuwqXL/TufD0s4VeHIUvOBGY
W/DwKDxENDSMs/Qx0gDX9QDh4rkseTOJJBFFjYHvabUBF1Gt629vru/CKU7RFkyeBrskKwTKtswU
KaGNswDYmqJFGWlcChJ7MHW4LIZaxvsexznjwiQobt4EYBTfcDQVKa1fwqS2giY8ih6FQjg9D5Yz
WEzcOcGLV85OdzpemMO9wlRL/VMmCmT3gqMG3Y4nMWY5gVLdNi6emN4as1kW6TxoAvEyx5xIFaLK
LNjxsNNRCKbVMmPGf8wNROqmBVQdzC6WB5b1boPi6bKZOzmjWBTve9xbXcKi6InHknf/tSd1ruzP
0IHKkax9ie+Mq02IXnQKzwjQTDBZuaGrukgxvtLdRP1u6e8y6JET5V2nFMQIZs0KRLovuzlKsOPV
myb8wDThKJ+SsBLniMutkXpwtZGc0xHikVP5LU3cHDjCXBBctGQAMB5nD2CQsQwqCppQKYFpTRk0
0lcd9eNodEworbBGYKCx+ik1klCIC7lTxslbJrPlIkoSIJXiWOyj5jKocIsRLC8hkmQ34WzhTjFJ
ppBFwWf4QLkgkWxWWhezICWZMei0h/0wNeIS2u23ex0f5EC+Z7F/LBX+p1S8LBV+u/NIKny/ffmc
RNweTAkgHiliDc3yfWyOYcQY3t2tLO7RA2K3NyQLp8x4k4CciBBimiNO9DpbWRbkLYx9eZkLS6ZP
EAQ4lVVQiRfl604YjXac69FdaFMoeyqGCrWNxELIYCqADslWqvortVVsll6HVwEOglS+ZV+neFnr
3ejr/RuIGqF5dg9m2WF2jWQvxSrrZeGGWPaEZ4zbO/W4/2Yl1pbc0Up5WSW8LM/3B7NYMvVpAn/0
E/j+YAL/B4n9/84HdRdESVEmYoe6665jAygcFmmZr8mMUixaWinlsIVx9KNWufLwP3L++uWdodpy
YhGubGT7pWgCmrIJd6oe3CxLd7BbuFul8IoVBbMCgy84cC1chSw3I2YfYKXqov+F4uXYmz6DB/8F
PBQ5CsEphiMDrkg4P4Z6J2MIjkRnydDFCRkanCVD/ZMx5LcvzpKhy5Mx1DtLfr48HT9nuk4PTsbQ
xZmu08MTMnSe63S3czKK+h/XQm2py722V3tVWyomLFqZh/JOdnylbrTkzay+29J7c/ncdarLuu0N
n7rcVdIUFXDyV1e6pe7Db79++ONnp/QYcxAeOlBXvOYSbuf28LbSULSSf9iEDzYT+npwe485UPeY
Sri9aBt0LyshHEUMYKW/lxGpASdsWZZXXn5CFWRlPVjBwIF9vPnz98363Wb902b9/Wb93Wb942b9
w2b9HoQl/dK2SgX48/bG8bZMPMWKX8tK70lW/It2p9d/Div9elYOqbj1neNUgf4NRmLGcesrnGGO
BOOtz/8BRZ6pxW01mzr1dqrYq/lvifHfUEsHCFs//2uPBQAAbiEAAFBLAwQUAAgICACqZt1CAAAA
AAAAAAAAAAAALQAAAFNjcmlwdHMvamF2YXNjcmlwdC9MaWJyYXJ5L0dlbmVyYXRlRmVhdHVyZS5q
cwMAUEsHCAAAAAACAAAAAAAAAFBLAwQUAAgICACqZt1CAAAAAAAAAAAAAAAAMAAAAFNjcmlwdHMv
amF2YXNjcmlwdC9MaWJyYXJ5L3BhcmNlbC1kZXNjcmlwdG9yLnhtbI2QSw6CMBCG95yimb1Ud8a0
uMPErXqASRlJTRlIC0Rvb6Um4mNhl/+rX0Ztr40TI/lgW9awypcgiE1bWa41nI7lYg3bIlMdekNO
OOR6wJo07HHEg/G260HECQ6bFNEQJjX286qvoMhEfCqJv/spMsVca9DRFNNAPLMmu7Khc3hjbEiM
6Ia4syMmjz2VhP3gKb8EkJ8teiK1XHzHlZz7LxSZWN7gahul/39X54HNY/a/ipKJI55bpmMW2R1Q
SwcI22PX9bsAAACgAQAAUEsDBBQACAgIAKpm3UIAAAAAAAAAAAAAAAAVAAAATUVUQS1JTkYvbWFu
aWZlc3QueG1stVTLbsIwELzzFZHvsVtOVUTgUIleeiv9gMXZBCO/5Acif18nFY+qCiIlve16d2fG
O5YXq6OS2QGdF0aX5Jk+kQw1N5XQTUk+N+v8hayWs4UCLWr0oTgFWZrT/pyWJDpdGPDCFxoU+iLw
wljUleFRoQ7Fz/6iZzpnVwLmZDnLLny1kJineddeuusoZW4h7ErChkAuxworAXloLZYErJWCQ0ht
7KAr2gum1zpp48DuBPeEjdGx2UW11SCkZ+EUUqubAR1CQYOsq49iURiApkUOoAY8BtaVR4F6DCG5
7acHDq3E6WFfja5FE13vop8z4BwlptQ4xqNznYnDnI9x3fmsfNSdBBoF5dcI48jTaLh9mb8t8IM7
YYNneziA72P2LrYOXMveUHebxDVCiA7p3k+yxhuMFlyyL6/w+9y4yby7QXqHjVuhU+fDjP/MNAp+
wX794ssvUEsHCNzKUFpTAQAAAAYAAFBLAQIUABQAAAgAAKpm3UKfAy7EKwAAACsAAAAIAAAAAAAA
AAAAAAAAAAAAAABtaW1ldHlwZVBLAQIUABQAAAgAAKpm3UKdugMx6B0AAOgdAAAYAAAAAAAAAAAA
AAAAAFEAAABUaHVtYm5haWxzL3RodW1ibmFpbC5wbmdQSwECFAAUAAgICACqZt1CV085dK8BAAAZ
BAAACAAAAAAAAAAAAAAAAABvHgAAbWV0YS54bWxQSwECFAAUAAgICACqZt1CCjLuQpcHAAB0KAAA
DAAAAAAAAAAAAAAAAABUIAAAc2V0dGluZ3MueG1sUEsBAhQAFAAICAgAqmbdQjg+gbHPCQAAFEwA
AAoAAAAAAAAAAAAAAAAAJSgAAHN0eWxlcy54bWxQSwECFAAUAAgICACqZt1CAAAAAAIAAAAAAAAA
JwAAAAAAAAAAAAAAAAAsMgAAQ29uZmlndXJhdGlvbnMyL2FjY2VsZXJhdG9yL2N1cnJlbnQueG1s
UEsBAhQAFAAACAAAqmbdQgAAAAAAAAAAAAAAABoAAAAAAAAAAAAAAAAAgzIAAENvbmZpZ3VyYXRp
b25zMi90b29scGFuZWwvUEsBAhQAFAAACAAAqmbdQgAAAAAAAAAAAAAAABgAAAAAAAAAAAAAAAAA
uzIAAENvbmZpZ3VyYXRpb25zMi9mbG9hdGVyL1BLAQIUABQAAAgAAKpm3UIAAAAAAAAAAAAAAAAY
AAAAAAAAAAAAAAAAAPEyAABDb25maWd1cmF0aW9uczIvbWVudWJhci9QSwECFAAUAAAIAACqZt1C
AAAAAAAAAAAAAAAAGAAAAAAAAAAAAAAAAAAnMwAAQ29uZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsB
AhQAFAAACAAAqmbdQgAAAAAAAAAAAAAAABwAAAAAAAAAAAAAAAAAXTMAAENvbmZpZ3VyYXRpb25z
Mi9wcm9ncmVzc2Jhci9QSwECFAAUAAAIAACqZt1CAAAAAAAAAAAAAAAAGgAAAAAAAAAAAAAAAACX
MwAAQ29uZmlndXJhdGlvbnMyL3N0YXR1c2Jhci9QSwECFAAUAAAIAACqZt1CAAAAAAAAAAAAAAAA
GgAAAAAAAAAAAAAAAADPMwAAQ29uZmlndXJhdGlvbnMyL3BvcHVwbWVudS9QSwECFAAUAAAIAACq
Zt1CAAAAAAAAAAAAAAAAHwAAAAAAAAAAAAAAAAAHNAAAQ29uZmlndXJhdGlvbnMyL2ltYWdlcy9C
aXRtYXBzL1BLAQIUABQACAgIAKpm3UJbP/9rjwUAAG4hAAALAAAAAAAAAAAAAAAAAEQ0AABjb250
ZW50LnhtbFBLAQIUABQACAgIAKpm3UIAAAAAAgAAAAAAAAAtAAAAAAAAAAAAAAAAAAw6AABTY3Jp
cHRzL2phdmFzY3JpcHQvTGlicmFyeS9HZW5lcmF0ZUZlYXR1cmUuanNQSwECFAAUAAgICACqZt1C
22PX9bsAAACgAQAAMAAAAAAAAAAAAAAAAABpOgAAU2NyaXB0cy9qYXZhc2NyaXB0L0xpYnJhcnkv
cGFyY2VsLWRlc2NyaXB0b3IueG1sUEsBAhQAFAAICAgAqmbdQtzKUFpTAQAAAAYAABUAAAAAAAAA
AAAAAAAAgjsAAE1FVEEtSU5GL21hbmlmZXN0LnhtbFBLBQYAAAAAEgASAO8EAAAYPQAAAAA=
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

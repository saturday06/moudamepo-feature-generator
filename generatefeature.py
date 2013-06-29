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
kokjnlumYxbZHVBLBwjbY9f1uwAAAKABAABQSwMEFAAICAgAKWbdQgAAAAAAAAAAAAAAAC0AAABT
Y3JpcHRzL2phdmFzY3JpcHQvTGlicmFyeS9HZW5lcmF0ZUZlYXR1cmUuanMDAFBLBwgAAAAAAgAA
AAAAAABQSwMEFAAACAAAKWbdQgAAAAAAAAAAAAAAABoAAABDb25maWd1cmF0aW9uczIvcG9wdXBt
ZW51L1BLAwQUAAAIAAApZt1CAAAAAAAAAAAAAAAAHwAAAENvbmZpZ3VyYXRpb25zMi9pbWFnZXMv
Qml0bWFwcy9QSwMEFAAACAAAKWbdQgAAAAAAAAAAAAAAABoAAABDb25maWd1cmF0aW9uczIvc3Rh
dHVzYmFyL1BLAwQUAAAIAAApZt1CAAAAAAAAAAAAAAAAGAAAAENvbmZpZ3VyYXRpb25zMi9tZW51
YmFyL1BLAwQUAAAIAAApZt1CAAAAAAAAAAAAAAAAGAAAAENvbmZpZ3VyYXRpb25zMi9mbG9hdGVy
L1BLAwQUAAAIAAApZt1CAAAAAAAAAAAAAAAAGAAAAENvbmZpZ3VyYXRpb25zMi90b29sYmFyL1BL
AwQUAAAIAAApZt1CAAAAAAAAAAAAAAAAHAAAAENvbmZpZ3VyYXRpb25zMi9wcm9ncmVzc2Jhci9Q
SwMEFAAACAAAKWbdQgAAAAAAAAAAAAAAABoAAABDb25maWd1cmF0aW9uczIvdG9vbHBhbmVsL1BL
AwQUAAgICAApZt1CAAAAAAAAAAAAAAAAJwAAAENvbmZpZ3VyYXRpb25zMi9hY2NlbGVyYXRvci9j
dXJyZW50LnhtbAMAUEsHCAAAAAACAAAAAAAAAFBLAwQUAAgICAApZt1CAAAAAAAAAAAAAAAAFQAA
AE1FVEEtSU5GL21hbmlmZXN0LnhtbLWUz27CMAzG7zxFlXuTjdNUUThMYpfdxh7ApG4Jap0ocRB9
+7Wd+DNNIBDl5iT29/sUO5kt9k2d7NAHYykXr/JFJEjaFoaqXHyvlumbWMwnswbIlBg4OwRJV0fh
uMxF9JRZCCZkBA2GjHVmHVJhdWyQOPubnw2k4+rMwFTMJ8mJV5oa067et6fsMtZ16oA3uVCXRE7b
DRYGUm4d5gKcq40G7tLUjgo5GJbnPmXlwW2MDkLd42O1ic2awNRB8SGUjqoLPkwDFar+/C5Kgwyy
u8gLqox7Vv3xXaIBmbtuh/GFua1xfFltiftGja37pb1xHNQWdhCGWH2atQffKgdeY50W+Ltv/RX4
aNAPJPTAuETg6FFuw7OJNzyataEu82Hik0lPkn+3VJoq+kEhTBXobir6HlmvdPT++lQ+xrrxQwuR
egsyGqnPFXr4TP37xec/UEsHCEBDBuRPAQAAAAYAAFBLAQIUABQAAAgAAClm3UKfAy7EKwAAACsA
AAAIAAAAAAAAAAAAAAAAAAAAAABtaW1ldHlwZVBLAQIUABQAAAgAAClm3UJQ+rpE6B0AAOgdAAAY
AAAAAAAAAAAAAAAAAFEAAABUaHVtYm5haWxzL3RodW1ibmFpbC5wbmdQSwECFAAUAAgICAApZt1C
qZUa9q4BAAAZBAAACAAAAAAAAAAAAAAAAABvHgAAbWV0YS54bWxQSwECFAAUAAgICAApZt1CLP+r
Z4gHAABkKAAADAAAAAAAAAAAAAAAAABTIAAAc2V0dGluZ3MueG1sUEsBAhQAFAAICAgAKWbdQjg+
gbHPCQAAFEwAAAoAAAAAAAAAAAAAAAAAFSgAAHN0eWxlcy54bWxQSwECFAAUAAgICAApZt1Cr5ja
oqwFAAAHIwAACwAAAAAAAAAAAAAAAAAcMgAAY29udGVudC54bWxQSwECFAAUAAgICAApZt1C22PX
9bsAAACgAQAAMAAAAAAAAAAAAAAAAAABOAAAU2NyaXB0cy9qYXZhc2NyaXB0L0xpYnJhcnkvcGFy
Y2VsLWRlc2NyaXB0b3IueG1sUEsBAhQAFAAICAgAKWbdQgAAAAACAAAAAAAAAC0AAAAAAAAAAAAA
AAAAGjkAAFNjcmlwdHMvamF2YXNjcmlwdC9MaWJyYXJ5L0dlbmVyYXRlRmVhdHVyZS5qc1BLAQIU
ABQAAAgAAClm3UIAAAAAAAAAAAAAAAAaAAAAAAAAAAAAAAAAAHc5AABDb25maWd1cmF0aW9uczIv
cG9wdXBtZW51L1BLAQIUABQAAAgAAClm3UIAAAAAAAAAAAAAAAAfAAAAAAAAAAAAAAAAAK85AABD
b25maWd1cmF0aW9uczIvaW1hZ2VzL0JpdG1hcHMvUEsBAhQAFAAACAAAKWbdQgAAAAAAAAAAAAAA
ABoAAAAAAAAAAAAAAAAA7DkAAENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIvUEsBAhQAFAAACAAA
KWbdQgAAAAAAAAAAAAAAABgAAAAAAAAAAAAAAAAAJDoAAENvbmZpZ3VyYXRpb25zMi9tZW51YmFy
L1BLAQIUABQAAAgAAClm3UIAAAAAAAAAAAAAAAAYAAAAAAAAAAAAAAAAAFo6AABDb25maWd1cmF0
aW9uczIvZmxvYXRlci9QSwECFAAUAAAIAAApZt1CAAAAAAAAAAAAAAAAGAAAAAAAAAAAAAAAAACQ
OgAAQ29uZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsBAhQAFAAACAAAKWbdQgAAAAAAAAAAAAAAABwA
AAAAAAAAAAAAAAAAxjoAAENvbmZpZ3VyYXRpb25zMi9wcm9ncmVzc2Jhci9QSwECFAAUAAAIAAAp
Zt1CAAAAAAAAAAAAAAAAGgAAAAAAAAAAAAAAAAAAOwAAQ29uZmlndXJhdGlvbnMyL3Rvb2xwYW5l
bC9QSwECFAAUAAgICAApZt1CAAAAAAIAAAAAAAAAJwAAAAAAAAAAAAAAAAA4OwAAQ29uZmlndXJh
dGlvbnMyL2FjY2VsZXJhdG9yL2N1cnJlbnQueG1sUEsBAhQAFAAICAgAKWbdQkBDBuRPAQAAAAYA
ABUAAAAAAAAAAAAAAAAAjzsAAE1FVEEtSU5GL21hbmlmZXN0LnhtbFBLBQYAAAAAEgASAO8EAAAh
PQAAAAA=
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

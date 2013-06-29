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

odsBase64 = """
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
/e9/fPZ1tv8q23+X7T/P9t9k+y+z/SczXLJm+DTXiQlk9tVdVTz/haZxEsFbkhJF4dEzHl1cXT65
aFkc/vcZewYzLRVFB3ERPJPgh51+xUhjijZaCjeaDnf8u2LxF1BLBwjwf9cxewQAAO8QAABQSwME
FAAICAgAAHLbQgAAAAAAAAAAAAAAAC0AAABTY3JpcHRzL2phdmFzY3JpcHQvTGlicmFyeS9HZW5l
cmF0ZUZlYXR1cmUuanMDAFBLBwgAAAAAAgAAAAAAAABQSwMEFAAICAgAAHLbQgAAAAAAAAAAAAAA
ADAAAABTY3JpcHRzL2phdmFzY3JpcHQvTGlicmFyeS9wYXJjZWwtZGVzY3JpcHRvci54bWyNkEsO
gjAQhvecopm9VHfGtLjDxK16gEkZSU0ZSAtEb2+lJuJjYZf/q19Gba+NEyP5YFvWsMqXIIhNW1mu
NZyO5WIN2yJTHXpDTjjkesCaNOxxxIPxtutBxAkOmxTRECY19vOqr6DIRHwqib/7KTLFXGvQ0RTT
QDyzJruyoXN4Y2xIjOiGuLMjJo89lYT94Cm/BJCfLXoitVx8x5Wc+y8UmVje4Gobpf9/V+eBzWP2
v4qSiSOeW6ZjFtkdUEsHCNtj1/W7AAAAoAEAAFBLAwQUAAAIAAAActtCAAAAAAAAAAAAAAAAGgAA
AENvbmZpZ3VyYXRpb25zMi9wb3B1cG1lbnUvUEsDBBQAAAgAAABy20IAAAAAAAAAAAAAAAAfAAAA
Q29uZmlndXJhdGlvbnMyL2ltYWdlcy9CaXRtYXBzL1BLAwQUAAAIAAAActtCAAAAAAAAAAAAAAAA
GgAAAENvbmZpZ3VyYXRpb25zMi90b29scGFuZWwvUEsDBBQAAAgAAABy20IAAAAAAAAAAAAAAAAY
AAAAQ29uZmlndXJhdGlvbnMyL2Zsb2F0ZXIvUEsDBBQAAAgAAABy20IAAAAAAAAAAAAAAAAYAAAA
Q29uZmlndXJhdGlvbnMyL21lbnViYXIvUEsDBBQAAAgAAABy20IAAAAAAAAAAAAAAAAYAAAAQ29u
ZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsDBBQAAAgAAABy20IAAAAAAAAAAAAAAAAcAAAAQ29uZmln
dXJhdGlvbnMyL3Byb2dyZXNzYmFyL1BLAwQUAAgICAAActtCAAAAAAAAAAAAAAAAJwAAAENvbmZp
Z3VyYXRpb25zMi9hY2NlbGVyYXRvci9jdXJyZW50LnhtbAMAUEsHCAAAAAACAAAAAAAAAFBLAwQU
AAAIAAAActtCAAAAAAAAAAAAAAAAGgAAAENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIvUEsDBBQA
CAgIAABy20IAAAAAAAAAAAAAAAAVAAAATUVUQS1JTkYvbWFuaWZlc3QueG1stZTPbsIwDMbvPEWV
e5ON01RROExil93GHsCkbglqnShxEH37tZ34M00gEOXmxM73+xInmS32TZ3s0AdjKRev8kUkSNoW
hqpcfK+W6ZtYzCezBsiUGDg7BEm3jsJxmIvoKbMQTMgIGgwZ68w6pMLq2CBx9rc+G0jH0ZmBqZhP
khOvNDWm3XrfnqrLWNepA97kQl0SOU03WBhIuXWYC3CuNhq4K1M7KuRgWJ77lMF5hCJsEFmoe6ys
NrFZE5g6KD6E0lF1wYppoELV5++iBGTuehNkd54XlBn3rPr0fcLc1ji+bIMMo4tqS9z3amzdL+2N
46C2sIMwxOrTrD34Vn0goQfGJQJHj3IbLpBHIzrwGuu0wN95669sdzToDe9mbairfJj4ZNKT5N8t
laaKflAIUwW661F/MaxXOnp//VY+xrrxTwuRegsyGqnPFXr4TP37yOc/UEsHCITfUCJTAQAAAwYA
AFBLAQIUABQAAAgAAABy20KFbDmKLgAAAC4AAAAIAAAAAAAAAAAAAAAAAAAAAABtaW1ldHlwZVBL
AQIUABQAAAgAAABy20LEL9IBWwoAAFsKAAAYAAAAAAAAAAAAAAAAAFQAAABUaHVtYm5haWxzL3Ro
dW1ibmFpbC5wbmdQSwECFAAUAAgICAAActtCDit0EMAEAABvHgAADAAAAAAAAAAAAAAAAADlCgAA
c2V0dGluZ3MueG1sUEsBAhQAFAAICAgAAHLbQnRuF8NeBgAAahwAAAoAAAAAAAAAAAAAAAAA3w8A
AHN0eWxlcy54bWxQSwECFAAUAAgICAAActtCoCCVnI8BAAA5AwAACAAAAAAAAAAAAAAAAAB1FgAA
bWV0YS54bWxQSwECFAAUAAgICAAActtC8H/XMXsEAADvEAAACwAAAAAAAAAAAAAAAAA6GAAAY29u
dGVudC54bWxQSwECFAAUAAgICAAActtCAAAAAAIAAAAAAAAALQAAAAAAAAAAAAAAAADuHAAAU2Ny
aXB0cy9qYXZhc2NyaXB0L0xpYnJhcnkvR2VuZXJhdGVGZWF0dXJlLmpzUEsBAhQAFAAICAgAAHLb
Qttj1/W7AAAAoAEAADAAAAAAAAAAAAAAAAAASx0AAFNjcmlwdHMvamF2YXNjcmlwdC9MaWJyYXJ5
L3BhcmNlbC1kZXNjcmlwdG9yLnhtbFBLAQIUABQAAAgAAABy20IAAAAAAAAAAAAAAAAaAAAAAAAA
AAAAAAAAAGQeAABDb25maWd1cmF0aW9uczIvcG9wdXBtZW51L1BLAQIUABQAAAgAAABy20IAAAAA
AAAAAAAAAAAfAAAAAAAAAAAAAAAAAJweAABDb25maWd1cmF0aW9uczIvaW1hZ2VzL0JpdG1hcHMv
UEsBAhQAFAAACAAAAHLbQgAAAAAAAAAAAAAAABoAAAAAAAAAAAAAAAAA2R4AAENvbmZpZ3VyYXRp
b25zMi90b29scGFuZWwvUEsBAhQAFAAACAAAAHLbQgAAAAAAAAAAAAAAABgAAAAAAAAAAAAAAAAA
ER8AAENvbmZpZ3VyYXRpb25zMi9mbG9hdGVyL1BLAQIUABQAAAgAAABy20IAAAAAAAAAAAAAAAAY
AAAAAAAAAAAAAAAAAEcfAABDb25maWd1cmF0aW9uczIvbWVudWJhci9QSwECFAAUAAAIAAAActtC
AAAAAAAAAAAAAAAAGAAAAAAAAAAAAAAAAAB9HwAAQ29uZmlndXJhdGlvbnMyL3Rvb2xiYXIvUEsB
AhQAFAAACAAAAHLbQgAAAAAAAAAAAAAAABwAAAAAAAAAAAAAAAAAsx8AAENvbmZpZ3VyYXRpb25z
Mi9wcm9ncmVzc2Jhci9QSwECFAAUAAgICAAActtCAAAAAAIAAAAAAAAAJwAAAAAAAAAAAAAAAADt
HwAAQ29uZmlndXJhdGlvbnMyL2FjY2VsZXJhdG9yL2N1cnJlbnQueG1sUEsBAhQAFAAACAAAAHLb
QgAAAAAAAAAAAAAAABoAAAAAAAAAAAAAAAAARCAAAENvbmZpZ3VyYXRpb25zMi9zdGF0dXNiYXIv
UEsBAhQAFAAICAgAAHLbQoTfUCJTAQAAAwYAABUAAAAAAAAAAAAAAAAAfCAAAE1FVEEtSU5GL21h
bmlmZXN0LnhtbFBLBQYAAAAAEgASAO8EAAASIgAAAAA=
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

odsFile = tempfile.NamedTemporaryFile()
odsFile.write(base64.b64decode(odsBase64))
#print(odsFile.name)
odsFile.flush()

scriptOdsFile = tempfile.NamedTemporaryFile()
#print(scriptOdsFile.name)
scriptOdsFile.flush()

generateFeatureJsPath = "Scripts/javascript/Library/GenerateFeature.js"
zin = zipfile.ZipFile (odsFile.name, 'r')
zout = zipfile.ZipFile (scriptOdsFile.name, 'w')
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
print("script: " + scriptOdsFile.name)
url = "file://" + scriptOdsFile.name
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

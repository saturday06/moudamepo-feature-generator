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
    importClass(Packages.java.io.PrintWriter);
    importClass(Packages.java.io.FileOutputStream);
    importClass(Packages.java.io.OutputStreamWriter);
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
    importClass(Packages.com.sun.star.view.XSelectionSupplier);

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
    return "file://" + (nativePath + "").replace(/\\/g, "/").replace(/^\/?/, "/");
}

function getNativePath(starPath) {
    return decodeURI((starPath + "").replace(/^file:\/\//, "").replace(/^\/?([a-z]:)/i, "\1"));
}

var POSITIVE_REGEXP = /[○ＹY]/;
var NEGATIVE_REGEXP = /[×ＮN]/;
var IGNORE_REGEXP = /[ -‐]/;
var BRACKET_REGEXP = /[“”「」『』【】]/;

function WindowsFilesystem() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var BIF_NONEWFOLDERBUTTON = 512;
    var StreamTypeEnum = {
        adTypeBinary: 1,
        adTypeText: 2
    };
    var SaveOptionsEnum = {
        adSaveCreateOverWrite: 2
    };

    this.write = function (path, content, bom) {
        var stream = new ActiveXObject("ADODB.Stream");
        try {
            stream.Open();
            stream.Type = StreamTypeEnum.adTypeText;
            stream.Charset = "UTF-8";
            stream.WriteText(content.replace(/\r?\n/g, "\r\n"));
            if (!bom) {
                stream.Position = 0;
                stream.Type = StreamTypeEnum.adTypeBinary;
                stream.Position = 3; // skip BOM
                binary = stream.Read();
                stream.Close();
                stream.Open();
                stream.Write(binary);
            }
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

    this.createFolder = function (path) {
        fso.FolderExists(path) || fso.CreateFolder(path);
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
    this.write = function (path, content, bom) {
        var printStream;
        try {
            printStream = new PrintStream(new JFile(path), "UTF-8");
            if (bom) {
                printStream.write(0xef);
                printStream.write(0xbb);
                printStream.write(0xbf);
            }
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
        if (ARGUMENTS[2]) {
            return "" + ARGUMENTS[2];
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
        if (ARGUMENTS[3]) {
            return "" + ARGUMENTS[3];
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
            if (i % 10 == 0) {
                console.write(".");
            }
            var path = files[i].getAbsolutePath();
            if (files[i].isDirectory()) {
                spreadsheetFiles = spreadsheetFiles.concat(this.getSpreadsheetFiles(path));
            } else {
                if (path.match(/\.(xls|xlsx|ods)$/i)) {
                    console.write("*");
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
    var logPath = ARGUMENTS[1];
    this.write = function (message) {
        // UTF-8文字列をファイルに追記するもっと良い方法
        var fos;
        var osw;
        try {
            fos = new FileOutputStream(logPath, true);
            osw = new OutputStreamWriter(fos, "UTF-8");
            osw.write("" + message);
        } finally {
            try {
                if (osw) {
                    osw.close();
                }
            } catch (e) {
            }
            try {
                if (fos) {
                    fos.close();
                }
            } catch (e) {
            }
        }
    };

    this.log = function (message) {
        this.write(message + "\n");
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
    var DECIDION_TABLE_MAX_RIGHT = 200;
    var DECIDION_TABLE_MAX_BOTTOM = 200;
    var DECIDION_TABLE_MAX_IGNORED_LINES = 20;
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
        var sheet = this;

        function Cell(starCell) {
            this.getValue = function () {
                return qi(XText, starCell).getString() + "";
            };

            this.activate = function() {
                qi(XSelectionSupplier, qi(XModel, starBook).getCurrentController()).select(starCell);
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

        this.getWidth = function() {
            var right = undefined;
            var ignoredXLines = 0;
            for (var x = this.getTableLeft(); x < DECIDION_TABLE_MAX_RIGHT; ++x) {
                var valueFound = false;
                for (var y = this.getTableTop(); y < DECIDION_TABLE_MAX_BOTTOM; ++y) {
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
            return right;
        };

        this.getHeight = function() {
            var right = this.getWidth();
            var bottom = 0;
            var ignoredYLines = 0;
            for (var y = this.getTableTop(); y < DECIDION_TABLE_MAX_RIGHT; ++y) {
                var valueFound = false;
                for (var x = this.getTableLeft(); x <= right; ++x) {
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
            return bottom;
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
    var xlByRows = 1;
    var xlByColumns = 2;
    var xlPrevious = 2;
    var xlFormulas = -4123;
    var xlPart = 2;

    function Sheet(excelSheet) {
        var sheet = this;

        function Cell(excelCell) {
            this.getValue = function () {
                return (excelCell.Value || "") + "";
            };

            this.activate = function() {
                excelCell.Activate();
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

        this.getWidth = function() {
            return excelSheet.UsedRange.Find(
                "*", excelSheet.Cells(1, 1), xlFormulas, xlPart, xlByColumns, xlPrevious).Column;
        };

        this.getHeight = function() {
            return excelSheet.UsedRange.Find(
                "*", excelSheet.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious).Row;
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
        return "  " + NormalizeStep(command) + "\n";
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
    return "  " + NormalizeStep(command) + " \""
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

    var right = sheet.getWidth();
    if (!right) {
        return;
    }
    var bottom = sheet.getHeight();
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
            book = new ExcelBook(path);
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
            filesystem.write(outputPath, feature, true);
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
function ObjToString(obj, level)
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
            out += ObjToString(obj[i], level + 1);
        } else {
            out += i + ": " + obj[i] + "\n";
        }
    }
    return out;
}

function UseStarOfficeVariantInWindows(inputFolder, outputFolder) {
    var keys = [
        "HKEY_CLASSES_ROOT\\Software\\LibreOffice\\LibreOffice\\Path",
        "HKEY_CLASSES_ROOT\\Software\\OpenOffice\\OpenOffice\\Path"
    ];
    var installDir;
    var shell = new ActiveXObject("WScript.Shell");
    for (var i = 0; i < keys.length; i++) {
        try {
            installDir = shell.RegRead(keys[i]);
            break;
        } catch (e) {
        }
    }
    console.log("python=" + installDir + "program\\python.exe");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var TemporaryFolder = 2;
    var tempAppDir = fso.GetSpecialFolder(TemporaryFolder) + "\\SOfficeFeatureGenerator";
    var tempProcessDir = tempAppDir + "\\" + fso.GetTempName();
    if (!fso.FolderExists(tempAppDir)) {
        fso.CreateFolder(tempAppDir);
    }
    var ForReading = 1;
    var ForWriting = 2;
    var scriptFile = fso.OpenTextFile(WScript.ScriptFullName, ForReading);
    var scriptText = scriptFile.ReadAll();
    scriptFile.Close();

    fso.CreateFolder(tempProcessDir);
    console.log(tempProcessDir);
    var separator = "<<<<<" + "SEPARATOR" + ">>>>>";

    var jsPath = tempProcessDir + "\\GenerateFeature.js";
    var jsText = scriptText.replace(new RegExp(separator + "[\\s\\S]*", "m"), "")
    jsText = jsText.replace(new RegExp("^[\\s\\S]*?/\\*", "m"), "/*");
    filesystem.write(jsPath, jsText);

    var pyPath = tempProcessDir + "\\RunSOfficeScript.py";
    var pyText = scriptText.replace(new RegExp("[\\s\\S]*" + separator + "[\\s\\S]*?/\\*", "m"), "");
    filesystem.write(pyPath, pyText);

    // OpenOffice4.lは、なぜかパイプでスクリプトを渡さないとimport unoが失敗する
    var commandText = "type \"" + pyPath + "\" | \"" + installDir + "program\\python.exe\" - \"" + inputFolder + "\" \"" + outputFolder + "\" \"" + jsPath + "\"";
    var commandPath = tempProcessDir + "\\RunPython.bat";
    var commandFile = fso.CreateTextFile(commandPath, true);
    commandFile.WriteLine("cd /d \"" + installDir + "program\"");
    commandFile.WriteLine(commandText);
    commandFile.Close();

    var env = shell.Environment;
    // stdoutとstderrをマージするためにバッチファイルを作っているが、unicode必須パス上では動かない
    var command = "\"" + env("COMSPEC") + "\" /c \"" + commandPath + "\" 2>&1";
    console.log(command);
    var process = shell.Exec(command);
    while (!process.StdOut.AtEndOfStream) {
        console.write(process.StdOut.Read(1));
    }
}

function Main() {
    try {
        var inputFolder = filesystem.getInputFolder();
        var outputFolder = filesystem.getOutputFolder();
        if (WINDOWS) {
            try {
                new ActiveXObject("Excel.Application");
            } catch (e) {
                UseStarOfficeVariantInWindows(inputFolder, outputFolder);
                return;
            }
        }
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
        console.log(ObjToString(e, 1));
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
}

Main();

// <<<<<SEPARATOR>>>>>

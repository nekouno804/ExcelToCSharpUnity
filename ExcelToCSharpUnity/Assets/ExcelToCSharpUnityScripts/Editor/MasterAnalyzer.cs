using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEngine;

namespace Core.MasterScripts.MasterSystemsScripts
{
    /// <summary>
    /// マスタを解析し開始位置を取得し指定した処理を実行する
    /// </summary>
    public static class MasterAnalyzer {
        /// <summary>
        /// 各マスタを解析し各マスタに指定した処理を実行する
        /// </summary>
        /// <param name="excelFiles">処理するマスタのパス</param>
        /// <param name="action">読み込んだマスタに対して行う処理(xlsxファイル名、シート、マスタ開始行、マスタ開始列)</param>
        public static void Analyze(List<string> excelFiles, Action<string, ISheet, int, int> action) {
            // 順にエクセルファイルを取得する
            foreach (var path in excelFiles) {
                Debug.Log($"対象：{path}");
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    IWorkbook book = new XSSFWorkbook(fs);
                    var fileName = Path.GetFileName(path);

                    // シートを取得
                    int sheetCount = book.NumberOfSheets;
                    for (int sheetNum = 0; sheetNum < sheetCount; sheetNum++) {
                        var sheet = book.GetSheetAt(sheetNum);
                        // シートの行を取得
                        for (int rowCount = 0; rowCount <= sheet.LastRowNum; rowCount++) {
                            var row = sheet.GetRow(rowCount);
                            if (row == null) { continue; }

                            // シートの列を取得
                            for (int column = 0; column <= row.LastCellNum; column++) {
                                ICell cell = sheet.GetRow(rowCount).GetCell(column);
                                if (cell == null) { continue; }

                                // マスタデータ開始シンボルを検出した
                                if (cell.CellType == CellType.String) {
                                    if (cell.StringCellValue == "@") {
                                        action(fileName, sheet, rowCount, column);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;
using Object = UnityEngine.Object;

namespace ExcelToCSharpUnityScripts {
    /// <summary>
    /// マスタのベース
    /// </summary>
    public abstract class MasterBase<TKey, TValue> : ScriptableObject where TValue : IMasterData<TKey>
    {
        /// <summary> シリアライズデータ </summary>
        [SerializeField] protected List<TValue> SerializeData = new();

        /// <summary> マスタデータ（辞書） </summary>
        private static readonly Dictionary<TKey, TValue> masterDictionaries = new();

        /// <inheritdoc cref="masterDictionaries"/>
        public static IReadOnlyDictionary<TKey, TValue> MasterDictionaries => masterDictionaries;

        /// <summary> マスタ初期化 </summary>
        public void Init()
        {
            // データリストを辞書形式に変換
            foreach (var data in SerializeData)
            {
                // 追加
                var key = data.GetKey();
                if (masterDictionaries.TryAdd(key, data))
                {
                    continue;
                }

                // 追加に失敗(既に値が入っている)なら値を更新
                Debug.LogError($"{this}:マスタのキーが被っています({key}). 値を更新します");
                masterDictionaries[key] = data;
            }

            OnInit();
        }

        /// <summary> 辞書作成後に呼ばれる初期化時処理 </summary>
        /// <remarks> 初期化時したいことあればここで </remarks>
        protected virtual void OnInit()
        {
        }

#if UNITY_EDITOR
        /// <summary>
        /// 各マスタを解析し各マスタに指定した処理を実行する
        /// </summary>
        /// <param name="path">処理するマスタのパス</param>
        /// <returns>何もなければsheetがnull</returns>
        public static (ISheet sheet, int startRow, int lastRow) GetMasterSheetAndPoints(string path)
        {
            path = Directory.GetCurrentDirectory() + path;
            // 順にエクセルファイルを取得する
            Debug.Log($"対象：{path}");
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook book = new XSSFWorkbook(fs);

                // シートを取得
                var sheetCount = book.NumberOfSheets;
                for (int sheetNum = 0; sheetNum < sheetCount; sheetNum++)
                {
                    var sheet = book.GetSheetAt(sheetNum);
                    // シートの行を取得
                    for (int rowCount = 0; rowCount <= sheet.LastRowNum; rowCount++)
                    {
                        var row = sheet.GetRow(rowCount);
                        if (row == null)
                        {
                            continue;
                        }

                        // シートの列を取得
                        for (int column = 0; column <= row.LastCellNum; column++)
                        {
                            ICell cell = sheet.GetRow(rowCount).GetCell(column);
                            if (cell == null)
                            {
                                continue;
                            }

                            // マスタデータ開始シンボルを検出した
                            if (cell.CellType == CellType.String)
                            {
                                if (cell.StringCellValue == "@")
                                {
                                    var startRow = rowCount + (int)MasterRowDefinition.Data;
                                    return (sheet, startRow, sheet.LastRowNum);
                                }
                            }
                        }
                    }
                }

                return (null, 0, 0);
            }
        }

        /// <summary>
        /// プリミティブ型かどうかの判定
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsPrimitive(Type type) =>
            type.IsPrimitive ||
            type.IsEnum ||
            type.IsArray;

        /// <summary>
        /// プリミティブ型以外のマスタデータ取得
        /// </summary>
        /// <param name="type"></param>
        /// <param name="data"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static object GetNotPrimitiveMasterData(Type type, object data)
        {
            if (type.IsEnum) return Enum.Parse(type, (string)data);
            if (type.IsArray)
            {
                Type elementType = type.GetElementType();
                string str = data.ToString();
                string[] elements = str.Split(',');
                Array array = Array.CreateInstance(elementType, elements.Length);
                for (int i = 0; i < elements.Length; i++)
                {
                    object elem = Convert.ChangeType(elements[i], elementType);
                    array.SetValue(elem, i);
                }

                return array;
            }

            return default;
        }

        /// <summary>
        /// エクセルシート(Sheet)からセル取得
        /// </summary>
        /// <param name="sheet">シート</param>
        /// <param name="rownum">行</param>
        /// <param name="cellnum">列</param>
        /// <returns></returns>
        protected static ICell GetCell(ISheet sheet, int rownum, int cellnum)
        {
            ICell cell = sheet.GetRow(rownum).GetCell(cellnum);
            return cell;
        }
        
        protected static string GetCellString(ISheet sheet, int rownum, int cellnum)
        {
            ICell cell = sheet.GetRow(rownum).GetCell(cellnum);
            if (cell == null) return string.Empty;
            if (cell.CellType != CellType.String) return string.Empty;
            switch(cell.CellType) {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// ScriptableObjectアセットの保存
        /// 既存のmetaファイルを上書きしないように上書きする https://kan-kikuchi.hatenablog.com/entry/CreateAssetWithOverwrite
        /// </summary>
        /// <param name="asset">保存するインスタンス</param>
        /// <param name="masterName">マスタ名</param>
        public static void CreateAssetWithOverwrite(Object asset, string masterName)
        {
            var exportPath = $"{MasterDataHelper.scriptableObjectMastersPath}{masterName}.asset";

            // ディレクトリがなければ作成
            var outPutResourcesFolderPath =
                $@"{Directory.GetCurrentDirectory()}\{MasterDataHelper.scriptableObjectMastersPath}";
            if (!Directory.Exists(outPutResourcesFolderPath)) Directory.CreateDirectory(outPutResourcesFolderPath);

            // 未作成であれば作成
            if (!File.Exists(exportPath))
            {
                AssetDatabase.CreateAsset(asset, exportPath);
                return;
            }

            //仮ファイルを作るためのディレクトリを作成
            var fileName = Path.GetFileName(exportPath);
            var tmpDirectoryPath = Path.Combine(exportPath.Replace(fileName, ""), "tmpDirectory");
            Directory.CreateDirectory(tmpDirectoryPath);

            //仮ファイルを保存
            var tmpFilePath = Path.Combine(tmpDirectoryPath, fileName);
            AssetDatabase.CreateAsset(asset, tmpFilePath);

            //仮ファイルを既存のファイルに上書き(metaデータはそのまま)
            FileUtil.ReplaceFile(tmpFilePath, exportPath);

            //仮ディレクトリとファイルを削除
            AssetDatabase.DeleteAsset(tmpDirectoryPath);

            //データ変更をUnityに伝えるためインポートしなおし
            AssetDatabase.ImportAsset(exportPath);
        }
        
        /// <summary>
        /// スクリプタブルオブジェクト書き出しパスを取得する
        /// </summary>
        /// <param name="masterName">マスタ名</param>
        /// <returns></returns>
        public static string GetScriptableObjectMasterPath(string masterName) => $"{MasterDataHelper.scriptableObjectMastersPath}{masterName}.asset";
#endif
    }
}
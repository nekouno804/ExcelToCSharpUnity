using System.Collections.Generic;
using System.IO;
using System.Linq;
using Core.MasterScripts.MasterSystemsScripts;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

namespace ExcelToCSharpUnityScripts.Editor
{
    /// <summary>
    /// ExcelファイルからC#定義を生成する
    /// </summary>
    public class ConvertExcelToMaster : MonoBehaviour
    {
        private static List<string> _masterNames = new();

        /// <summary>
        /// エクセルから型定義を生成
        /// </summary>
        [MenuItem("Master/1.エクセルファイルからスクリプトを生成")]
        private static void ConvertExcelToCsharp()
        {
            var fileNameList = new List<string>();
            var currentDirectory = Directory.GetCurrentDirectory() + MasterDataHelper.excelFilePath;
            Debug.Log($"下記のフォルダにあるマスタデータからスクリプトを生成します。\n{MasterDataHelper.excelFilePath}");
            try
            {
                var excels = Directory.GetFiles(currentDirectory, "*.xlsx");

                // エクセルファイルを取得する
                foreach (var fileName in excels) if (!fileName.Contains("~")) { fileNameList.Add(fileName); }
            
                Debug.Log($"マスタを構造に変換します。");
                MasterAnalyzer.Analyze(fileNameList,
                    (fileName, sheet, row, column) => GenerateMasterFile(fileName, sheet, row, column));

                CreateMasterDataLoader(_masterNames);

                AssetDatabase.Refresh();
            }
            catch (System.Exception e)
            {
                Debug.LogError($"エラーが発生しました。{e.Message}");
            }
        }

        /// <summary>
        /// マスタ構造をC#化
        /// </summary>
        /// <param name="fileName">xlsxファイル名</param>
        /// <param name="sheet">シート</param>
        /// <param name="rowNum">マスタ開始シンボル行</param>
        /// <param name="columnNum">マスタ開始シンボル列</param>
        private static void GenerateMasterFile(string fileName, ISheet sheet, int rowNum, int columnNum)
        {
            var masterStrings = new List<string>();
            (int row, int column) initLocate = (rowNum, columnNum);

            // 型定義列の列の長さを見るためセルを1つ取得
            var cell = GetCell(sheet, initLocate.row + (int)MasterRowDefinition.Type, initLocate.column);
            if (cell == null) { return; }

            // @の列のデータがインデックスとして利用される(辞書型のKey)
            var keyType = GetCell(sheet, initLocate.row + (int)MasterRowDefinition.Type, initLocate.column);
            var keyName = GetCell(sheet, initLocate.row + (int)MasterRowDefinition.Name, initLocate.column);

            var masterNameCell = GetCell(sheet, 0, 0);
            var masterSummaryCell = GetCell(sheet, 0, 2);
            if (masterNameCell == null) { return; }
            var masterName = masterNameCell.StringCellValue;
            _masterNames.Add(masterName);

            // 中身生成
            masterStrings.Add("using System;");
            masterStrings.Add("using ExcelToCSharpUnityScripts;");
            masterStrings.Add("using UnityEngine;");
            masterStrings.Add($"#if UNITY_EDITOR");
            masterStrings.Add("using UnityEditor;");
            masterStrings.Add("#endif");
            masterStrings.Add("namespace MasterScripts {");
            masterStrings.Add($"\t/// <summary>\n\t/// {masterSummaryCell?.StringCellValue}\n\t/// </summary>");
            masterStrings.Add($"\t[CreateAssetMenu(fileName = \"{masterName}\", menuName = \"Master/GenerateMaster/{masterNameCell.StringCellValue}\")]");
            masterStrings.Add($"\tpublic partial class {masterNameCell.StringCellValue} : MasterBase<{keyType}, {masterNameCell.StringCellValue}.MasterData> {{");
            masterStrings.Add($"\t\tprivate const string masterName = \"{masterName}\";");
            masterStrings.Add($"\t\tprivate const string excelFilePath = @\"{MasterDataHelper.excelFilePath}{fileName}\";\n");
            masterStrings.Add($"\t\t[Serializable] public class MasterData : IMasterData<{keyType}> {{");
            // 列をずらしながら構造を定義していく
            var argTypes = new List<string>();
            var argNames = new List<string>();
            var dataSetters = new List<string>();
            for (int currentColumn = initLocate.column; currentColumn <= cell.Row.LastCellNum; currentColumn++) {
                // マスタの変数型、変数名、変数説明を読み込む
                var typeCell = GetCell(sheet, initLocate.row + (int)MasterRowDefinition.Type, currentColumn);
                var nameCell = GetCell(sheet, initLocate.row + (int)MasterRowDefinition.Name, currentColumn);
                var summaryCell = GetCell(sheet, initLocate.row + (int)MasterRowDefinition.Summary, currentColumn);
                if (typeCell == null || nameCell == null) { continue; }

                // コンストラクタ用
                var memberType = typeCell.StringCellValue;
                argTypes.Add(memberType);
                argNames.Add(nameCell.StringCellValue);
                // マスタのある列でEnumかどうか判別
                var dataSetter = memberType switch {
                    "string" => $"GetCellString(data.sheet, row, {currentColumn})",
                    "bool" => $"GetCell(data.sheet, row, {currentColumn}).BooleanCellValue",
                    _ => MasterDataHelper.defaultTypes.Contains(memberType)
                        ? $"({memberType})GetCell(data.sheet, row, {currentColumn}).NumericCellValue"
                        : $"({memberType})(IsPrimitive(typeof({memberType}))\n\t\t\t\t\t\t" + 
                          $"? GetNotPrimitiveMasterData(typeof({memberType}), GetCellString(data.sheet, row, {currentColumn}))\n\t\t\t\t\t\t\t" +
                          $": (asset.SerializeData.Count + data.startRow < row && asset.SerializeData[row - data.startRow].{nameCell} == null)\n\t\t\t\t\t\t\t" +
                          $"? default\n\t\t\t\t\t\t\t" +
                          $": asset.SerializeData[row - data.startRow].{nameCell})"};
                dataSetters.Add(dataSetter);

                masterStrings.Add($"\t\t\t/// <summary> {summaryCell?.StringCellValue} </summary>\n" +
                                  $"\t\t\t[SerializeField] private {typeCell.StringCellValue} _{nameCell.StringCellValue};\n" +
                                  $"\t\t\t/// <inheritdoc cref=\"_{nameCell.StringCellValue}\"/>\n" +
                                  $"\t\t\tpublic {typeCell.StringCellValue} {nameCell.StringCellValue} => _{nameCell.StringCellValue};\n");
            }
            masterStrings.Add("\t\t\t/// <inheritdoc/>");
            masterStrings.Add($"\t\t\t{keyType} IMasterData<{keyType}>.GetKey() => {keyName};");

            var args = argTypes.Zip(argNames,(t,n) => new {Type = t, Name = n});
            var argsStrList = string.Join(", ", args.Select(a => $"{a.Type} {a.Name}"));
            var variableStr = string.Join("\n", args.Select(a => $"\t\t\t\t_{a.Name} = {a.Name};"));

            // コンストラクタ
            masterStrings.Add($"\n\t\t\t/// <summary> コンストラクタ </summary>\n\t\t\tpublic MasterData({argsStrList}) {{");
            masterStrings.Add($"{variableStr}");
            masterStrings.Add("\t\t\t}\n");
            masterStrings.Add("\t\t}\n");

            // マスタ自動読み込み
            masterStrings.Add($"#if UNITY_EDITOR");
            masterStrings.Add($"\t\t/// <summary> マスタデータをExcelから読み込み </summary>");
            masterStrings.Add($"\t\tpublic static void SetDataFromExcel() {{");
            masterStrings.Add($"\t\t\tvar data = GetMasterSheetAndPoints(excelFilePath);");
            masterStrings.Add($"\t\t\tif (data.sheet == null) return;\n");
            masterStrings.Add($"\t\t\tvar asset = ({masterName})AssetDatabase.LoadAssetAtPath(GetScriptableObjectMasterPath(masterName), typeof({masterName}));\n");
            masterStrings.Add($"\t\t\tvar instance = CreateInstance<{masterName}>();");
            masterStrings.Add($"\t\t\tfor (var row = data.startRow; row <= data.lastRow; row++) {{");
            masterStrings.Add($"\t\t\t\tvar master = new MasterData(");
            masterStrings.Add($"{string.Join(",\n", dataSetters.Select(a => $"\t\t\t\t\t{a}"))}");
            masterStrings.Add($"\t\t\t\t);");
            masterStrings.Add($"\t\t\t\tinstance.SerializeData.Add(master);");
            masterStrings.Add($"\t\t\t}}");
            masterStrings.Add($"\t\t\tCreateAssetWithOverwrite(instance, masterName);");

            masterStrings.Add("\t\t}\n#endif");
            masterStrings.Add("\t}");
            masterStrings.Add("}");

            // 書き出し
            var newFolderPath = $@"{Directory.GetCurrentDirectory()}\{MasterDataHelper.outputMasterPath}\{masterName}\";
            var outputFile = string.Join("\n", masterStrings.ToArray());
            if (!Directory.Exists(newFolderPath)) { Directory.CreateDirectory(newFolderPath); }
            File.WriteAllText($@"{newFolderPath}{masterName}.cs", outputFile);
        }

        /// <summary>
        /// マスタデータをエクセルから読み込む機能を提供する機能を生成する
        /// </summary>
        /// <param name="masterNames">マスタ名のリスト</param>
        private static void CreateMasterDataLoader(List<string> masterNames)
        {
            var masterStrings = new List<string>();
            masterStrings.Add($"using UnityEditor;");
            masterStrings.Add($"using UnityEngine;");
            masterStrings.Add($"using MasterScripts;");
            masterStrings.Add($"/// <summary>");
            masterStrings.Add($"/// エクセルにあるマスタのデータをスクリプタブルオブジェクトに入れる");
            masterStrings.Add($"/// </summary>");
            masterStrings.Add($"public class MasterDataSetter : MonoBehaviour {{");
            masterStrings.Add($"\t[MenuItem(\"Master/2.Excelデータをスクリプタブルオブジェクトに挿入\")]");
            masterStrings.Add($"\tprivate static void Execute() {{");
            masterStrings.Add($"{string.Join("\n", masterNames.Select(x => $"\t\t{x}.SetDataFromExcel();"))}");
            masterStrings.Add($"\t}}\n}}");

            // 書き出し
            var newFolderPath = $@"{Directory.GetCurrentDirectory()}\{MasterDataHelper.outputMasterPath}\MasterSystemsGenerated\";
            if (!Directory.Exists(newFolderPath)) { Directory.CreateDirectory(newFolderPath); }
            var newFilePath = $@"{newFolderPath}\MasterDataSetter.cs";
            var outputFile = string.Join("\n", masterStrings.ToArray());
            File.WriteAllText($@"{newFilePath}", outputFile);
        }

        /// <summary>
        /// エクセルシート(Sheet)からセル取得
        /// </summary>
        /// <param name="sheet">シート</param>
        /// <param name="rownum">行</param>
        /// <param name="cellnum">列</param>
        /// <returns></returns>
        public static ICell GetCell(ISheet sheet, int rownum, int cellnum)
        {
            ICell cell = sheet.GetRow(rownum).GetCell(cellnum);
            return cell;
        }
    }
}

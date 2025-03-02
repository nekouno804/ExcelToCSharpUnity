using System;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace ExcelToCSharpUnityScripts.Editor
{
    /// <summary>
    /// マスタ一括初期化処理生成
    /// </summary>
    public class MasterInitializerGenerator : MonoBehaviour
    {
        /// <summary> ジェネレータの出力先 </summary>
        private const string outputPath = @"\Assets\CommonScripts\MasterScripts\MasterSystemsGenerated\";

        [MenuItem("Master/3.マスタ一括初期化スクリプトを生成")]
        private static void Generate() {
            Debug.Log($"下記のフォルダにあるマスタデータを読み込みます。\n{MasterDataHelper.scriptableObjectMastersPath}");
            // パスを取得
            try
            {
                var scriptableObjectPaths = Directory.GetFiles(
                    $@"{Directory.GetCurrentDirectory()}\{MasterDataHelper.scriptableObjectMastersPath}", "*.asset",
                    SearchOption.AllDirectories);

                var csharpScripts = new List<string>();
                csharpScripts.Add("using UnityEngine;\n");
                csharpScripts.Add("namespace MasterScripts {");
                csharpScripts.Add("\t/// <summary>\n\t/// マスタ一括初期化\n\t/// </summary>");
                csharpScripts.Add("\tpublic static class MasterInitializer {");

                csharpScripts.Add("\t\t/// <summary>");
                csharpScripts.Add("\t\t/// マスタ一括初期化");
                csharpScripts.Add("\t\t/// Resourcesフォルダーにあるマスター(スクリプタブルオブジェクト)を読み込む初期化処理");
                csharpScripts.Add("\t\t/// </summary>");
                csharpScripts.Add("\t\tpublic static void InitMasters() {");

                // Init()呼び出しを定義していく
                foreach (var scriptPath in scriptableObjectPaths)
                {
                    var className = Path.GetFileNameWithoutExtension(scriptPath);
                    csharpScripts.Add($"\t\t\tvar {className} = Resources.Load<{className}>($@\"MasterData\\{className}\");");
                    csharpScripts.Add($"\t\t\t{className}.Init();");
                }
                csharpScripts.Add("\t\t}\n\t}\n}");

                // 書き出し
                var newFolderPath = $@"{Directory.GetCurrentDirectory()}{outputPath}\";
                if (!Directory.Exists(newFolderPath))
                {
                    Directory.CreateDirectory(newFolderPath);
                }

                var outputFile = string.Join("\n", csharpScripts.ToArray());
                File.WriteAllText($@"{newFolderPath}\MasterInitializer.cs", outputFile);

                AssetDatabase.Refresh();
            }
            catch (Exception e)
            {
                Debug.LogError($"エラーが発生しました。{e.Message}");
            }
        }
    }
}

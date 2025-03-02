namespace ExcelToCSharpUnityScripts
{
    /// <summary>
    /// マスタデータヘルパー
    /// </summary>
    public static class MasterDataHelper
    {
#if UNITY_EDITOR
        /// <summary> Excelを入れているフォルダパス </summary>
        public static string excelFilePath = @"\..\ExcelResources\";
        /// <summary> マスタ構造ファイルを書き出すフォルダパス </summary>
        public static string outputMasterPath = @"Assets\CommonScripts\MasterScripts\";
        /// <summary> ScriptableObject化したマスタのパス </summary>
        public static string scriptableObjectMastersPath = @"Assets\Resources\MasterData\";
#endif
    
        /// <summary> マスタ構造化時の変数型 </summary>
        public static string[] defaultTypes = {/*"bool",*/ "byte", "short", "ushort", "char", "int", "uint", /*"string",*/ "float", "long", "ulong", "double"};
    }
}

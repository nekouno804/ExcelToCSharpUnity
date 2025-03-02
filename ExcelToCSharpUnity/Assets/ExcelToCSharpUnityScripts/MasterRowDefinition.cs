namespace ExcelToCSharpUnityScripts
{
    /// <summary>
    /// Excelマスタファイルの行定義(n行目がなにを示すか)
    /// </summary>
    public enum MasterRowDefinition
    {
        // 変数型
        Type = 1,
        // 変数名
        Name = 2,
        // 変数説明
        Summary = 3,
        // マスタデータが始まる行
        Data = 4,
    }
}

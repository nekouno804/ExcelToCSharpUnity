namespace ExcelToCSharpUnityScripts {
    /// <summary>
    /// マスターデータ用インタフェース
    /// </summary>
    public interface IMasterData<out TKeyType> {
        /// <summary> キー取得 </summary>
        public TKeyType GetKey();
    }
}
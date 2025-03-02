using UnityEditor;
using UnityEngine;
using MasterScripts;
/// <summary>
/// エクセルにあるマスタのデータをスクリプタブルオブジェクトに入れる
/// </summary>
public class MasterDataSetter : MonoBehaviour {
	[MenuItem("Master/2.Excelデータをスクリプタブルオブジェクトに挿入")]
	private static void Execute() {
		CharacterMaster.SetDataFromExcel();
	}
}
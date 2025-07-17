using System;
using ExcelToCSharpUnityScripts;
using UnityEngine;
#if UNITY_EDITOR
using UnityEditor;
#endif
namespace MasterScripts {
	/// <summary>
	/// キャラクターマスタ
	/// </summary>
	[CreateAssetMenu(fileName = "CharacterMaster", menuName = "Master/GenerateMaster/CharacterMaster")]
	public partial class CharacterMaster : MasterBase<int, CharacterMaster.MasterData> {
		private const string masterName = "CharacterMaster";
		private const string excelFilePath = @"\..\ExcelResources\SampleMaster.xlsx";

		[Serializable] public class MasterData : IMasterData<int> {
			/// <summary> キャラクターId </summary>
			[SerializeField] private int _Id;
			/// <inheritdoc cref="_Id"/>
			public int Id => _Id;

			/// <summary> キャラクター名 </summary>
			[SerializeField] private string _Name;
			/// <inheritdoc cref="_Name"/>
			public string Name => _Name;

			/// <summary> 体力 </summary>
			[SerializeField] private int _HP;
			/// <inheritdoc cref="_HP"/>
			public int HP => _HP;

			/// <summary> 魔力 </summary>
			[SerializeField] private int _MP;
			/// <inheritdoc cref="_MP"/>
			public int MP => _MP;

			/// <summary> 物理攻撃力 </summary>
			[SerializeField] private int _Attack;
			/// <inheritdoc cref="_Attack"/>
			public int Attack => _Attack;

			/// <summary> 魔法攻撃力 </summary>
			[SerializeField] private int _MagicAttack;
			/// <inheritdoc cref="_MagicAttack"/>
			public int MagicAttack => _MagicAttack;

			/// <summary> 防御力 </summary>
			[SerializeField] private int _Defense;
			/// <inheritdoc cref="_Defense"/>
			public int Defense => _Defense;

			/// <summary> 行動速度 </summary>
			[SerializeField] private float _Speed;
			/// <inheritdoc cref="_Speed"/>
			public float Speed => _Speed;

			/// <summary> 属性 </summary>
			[SerializeField] private CharacterAttribute _Attribute;
			/// <inheritdoc cref="_Attribute"/>
			public CharacterAttribute Attribute => _Attribute;

			/// <inheritdoc/>
			int IMasterData<int>.GetKey() => Id;

			/// <summary> コンストラクタ </summary>
			public MasterData(int Id, string Name, int HP, int MP, int Attack, int MagicAttack, int Defense, float Speed, CharacterAttribute Attribute) {
				_Id = Id;
				_Name = Name;
				_HP = HP;
				_MP = MP;
				_Attack = Attack;
				_MagicAttack = MagicAttack;
				_Defense = Defense;
				_Speed = Speed;
				_Attribute = Attribute;
			}

		}

#if UNITY_EDITOR
		/// <summary> マスタデータをExcelから読み込み </summary>
		public static void SetDataFromExcel() {
			var data = GetMasterSheetAndPoints(excelFilePath);
			if (data.sheet == null) return;

			var asset = (CharacterMaster)AssetDatabase.LoadAssetAtPath(GetScriptableObjectMasterPath(masterName), typeof(CharacterMaster));

			var instance = CreateInstance<CharacterMaster>();
			for (var row = data.startRow; row <= data.lastRow; row++) {
				var master = new MasterData(
					(int)GetCell(data.sheet, row, 1).NumericCellValue,
					GetCellString(data.sheet, row, 2),
					(int)GetCell(data.sheet, row, 3).NumericCellValue,
					(int)GetCell(data.sheet, row, 4).NumericCellValue,
					(int)GetCell(data.sheet, row, 5).NumericCellValue,
					(int)GetCell(data.sheet, row, 6).NumericCellValue,
					(int)GetCell(data.sheet, row, 7).NumericCellValue,
					(float)GetCell(data.sheet, row, 8).NumericCellValue,
					(CharacterAttribute)(IsPrimitive(typeof(CharacterAttribute))
						? GetNotPrimitiveMasterData(typeof(CharacterAttribute), GetCellString(data.sheet, row, 9))
							: (asset.SerializeData.Count + data.startRow < row && asset.SerializeData[row - data.startRow].Attribute == null)
							? default
							: asset.SerializeData[row - data.startRow].Attribute)
				);
				instance.SerializeData.Add(master);
			}
			CreateAssetWithOverwrite(instance, masterName);
		}
#endif
	}
}
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuzuOffice
{
	using VBE = Microsoft.Vbe.Interop;

	public class VBEOperater
	{
		/// <summary>
		/// モジュールをクリアする。
		/// </summary>
		/// <param name="_Project">モジュールをクリアするプロジェクト</param>
		public void ClearModules(VBE.VBProject _Project)
		{
			if (_Project is null)	throw new ArgumentNullException(nameof(_Project));

			foreach (VBE.VBComponent component in _Project.VBComponents)
			{
				//標準モジュール(.bas) / クラスモジュール(.cls)を全て削除
				if ((component.Type == VBE.vbext_ComponentType.vbext_ct_StdModule) || (component.Type == VBE.vbext_ComponentType.vbext_ct_ClassModule))
				{
					_Project.VBComponents.Remove(component);
				}
			}

			//消去の成否の確認
			if (IsModuleClearSuccess(_Project)) throw new Exception("標準モジュール,クラスモジュールの削除に失敗しました");
		}

		/// <summary>
		/// モジュールを取り込む
		/// </summary>
		/// <param name="_Project">モジュールを取り込ませるプロジェクト</param>
		/// <param name="_PathList">取り込みたいモジュールの絶対パス</param>
		public void ImportModules(VBE.VBProject _Project, List<string> _PathList)
		{
			if (_Project is null)		throw new ArgumentNullException(nameof(_Project));
			if (_PathList is null)	throw new ArgumentNullException(nameof(_PathList));

			foreach (string module_path in _PathList)
			{
				ImportModule(_Project, module_path);
			}
		}

		/// <summary>
		/// モジュールを外部に書き出す
		/// </summary>
		/// <param name="_Project">書き出すディレクトリ</param>
		/// <param name="_Path">書き出し先のディレクトリ</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public List<string> ExportModules(VBE.VBProject _Project, in string _Path)
		{
			if (_Project is null)	throw new ArgumentNullException(nameof(_Project));
			if (string.IsNullOrWhiteSpace(_Path))	throw new ArgumentException("無効なパスです", nameof(_Path));

			List<string> module_pathes = new List<string>();

			string file_name = "";
			foreach (VBE.VBComponent component in _Project.VBComponents)
			{
				if ((component.Type == VBE.vbext_ComponentType.vbext_ct_StdModule) || (component.Type == VBE.vbext_ComponentType.vbext_ct_ClassModule))
				{
					file_name = _Path + "\\" + component.Name;

					switch (component.Type)
					{
						case VBE.vbext_ComponentType.vbext_ct_StdModule:	file_name += VBA_MODULE_EXTENSION;	break;
						case VBE.vbext_ComponentType.vbext_ct_ClassModule:	file_name += VBA_CLASS_EXTENSION;	break;
					}

					component.Export(file_name);
					module_pathes.Add(file_name);
				}
			}

			file_name = _Path + "//" + THIS_WORKBOOK + VBA_CLASS_EXTENSION;
			_Project.VBComponents.Item(THIS_WORKBOOK).Export(file_name);
			module_pathes.Add(file_name);

			return module_pathes;
		}

		//*************************************************************************************************//
		//private method
		//*************************************************************************************************//

		/// <summary>
		/// ThisWorkBookをクリアする
		/// </summary>
		/// <param name="_Project"></param>
		/// <returns>クリア前のコード</returns>
		private string ClearThisWorkbookModule(VBE.VBProject _Project)
		{
			var conponent = _Project.VBComponents.Item(THIS_WORKBOOK);
			int line_count = conponent.CodeModule.CountOfLines;
			string code = conponent.CodeModule.get_Lines(1, line_count);
			conponent.CodeModule.DeleteLines(1, line_count);

			return code;
		}

		private bool IsModuleClearSuccess(VBE.VBProject _Project)
		{
			foreach (VBE.VBComponent component in _Project.VBComponents)
			{
				//標準モジュール(.bas) / クラスモジュール(.cls)を全て削除
				if ((component.Type == VBE.vbext_ComponentType.vbext_ct_StdModule) || (component.Type == VBE.vbext_ComponentType.vbext_ct_ClassModule))
				{
					return false;
				}
			}

			return true;
		}

		private void ImportModule(VBE.VBProject _Project, in string _Path)
		{
			if (string.IsNullOrWhiteSpace(_Path))	throw new ArgumentException("無効なパスです", nameof(_Path));

			if (File.Exists(_Path))
			{
				if (Path.GetFileName(_Path) == THIS_WORKBOOK + VBA_CLASS_EXTENSION)
				{
					ImportThisWorkbookModule(_Project, _Path);
				}
				else
				{
					_Project.VBComponents.Import(_Path);
				}
			}
			else
			{
				throw new FileNotFoundException("[" + _Path + "]は無効なパスです");
			}
		}

		private void ImportThisWorkbookModule(VBE.VBProject _Project, in string _Path)
		{
			string original_code = ClearThisWorkbookModule(_Project);

			try
			{
				StreamReader _Reader = new StreamReader(_Path, Encoding.GetEncoding("Shift_JIS"));

				var conponent = _Project.VBComponents.Item(THIS_WORKBOOK);
				conponent.CodeModule.AddFromString(_Reader.ReadToEnd());
			}
			catch
			{
				if (original_code != "")
				{
					var conponent = _Project.VBComponents.Item(THIS_WORKBOOK);
					conponent.CodeModule.AddFromString(original_code);
				}

				throw new Exception("ThisWorkbookの更新に失敗しました");
			}
		}

		private const string VBA_MODULE_EXTENSION = ".bas";
		private const string VBA_CLASS_EXTENSION = ".cls";

		private const string THIS_WORKBOOK = "ThisWorkbook";
	}
}

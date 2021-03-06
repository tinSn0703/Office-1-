﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office
{
	using VBE = Microsoft.Vbe.Interop;

	class VBProjectOperater
	{
		/// <summary>
		/// モジュールをクリアする。
		/// </summary>
		/// <param name="_Project">モジュールをクリアするプロジェクト</param>
		public void ClearModules(VBE.VBProject _Project)
		{
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
		/// <param name="module_pathes">取り込みたいモジュールの絶対パス</param>
		public void ImportModules(VBE.VBProject _Project, List<string> module_pathes)
		{
			foreach (string module_path in module_pathes)
			{
				ImportModule(_Project, module_path);
			}
		}

		/// <summary>
		/// モジュールを外部に書き出す
		/// </summary>
		/// <param name="_Project">書き出すディレクトリ</param>
		/// <param name="path">書き出し先のディレクトリ</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public List<string> ExportModules(VBE.VBProject _Project, in string path)
		{
			List<string> module_pathes = new List<string>();

			foreach (VBE.VBComponent component in _Project.VBComponents)
			{
				if ((component.Type == VBE.vbext_ComponentType.vbext_ct_StdModule) || (component.Type == VBE.vbext_ComponentType.vbext_ct_ClassModule))
				{
					string file_name = path + "\\" + component.Name;

					switch (component.Type)
					{
						case VBE.vbext_ComponentType.vbext_ct_StdModule:	file_name += VBA_MODULE_EXTENSION;	break;
						case VBE.vbext_ComponentType.vbext_ct_ClassModule:	file_name += VBA_CLASS_EXTENSION;	break;
					}

					component.Export(file_name);
					module_pathes.Add(file_name);
				}
			}

			return module_pathes;
		}

		private string ClearThisWorkbookModule(VBE.VBProject _Project)
		{
			var conponent = _Project.VBComponents.Item("ThisWorkbook");
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

		private void ImportModule(VBE.VBProject _Project, string module_path)
		{
			if (File.Exists(module_path))
			{
				if (Path.GetFileName(module_path) == "ThisWorkbook.cls")
				{
					LoadModuleThisWorkbook(_Project, module_path);
				}
				else
				{
					_Project.VBComponents.Import(module_path);
				}
			}
			else
			{
				throw new Exception(module_path + "は存在しません");
			}
		}

		private void LoadModuleThisWorkbook(VBE.VBProject _Project, string module_path)
		{
			string original_code = ClearThisWorkbookModule(_Project);

			try
			{
				StreamReader _Reader = new StreamReader(module_path, Encoding.GetEncoding("Shift_JIS"));

				var conponent = _Project.VBComponents.Item("ThisWorkbook");
				conponent.CodeModule.AddFromString(_Reader.ReadToEnd());
			}
			catch
			{
				if (original_code != "")
				{
					var conponent = _Project.VBComponents.Item("ThisWorkbook");
					conponent.CodeModule.AddFromString(original_code);
				}

				throw new Exception("ThisWorkbookのマクロの更新に失敗しました");
			}
		}

		private const string VBA_MODULE_EXTENSION = ".bas";
		private const string VBA_CLASS_EXTENSION = ".cls";
	}
}

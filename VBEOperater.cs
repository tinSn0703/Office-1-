using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuzuOffice
{
	using VBE = Microsoft.Vbe.Interop;

	public class ModuleExtension
	{
		public string VBA_MODULE => ".bas";
		public string VBA_CLASS => ".cls";
		public string VBA_FORM => ".frm";

		public string Convert(VBE.vbext_ComponentType _Type)
		{
			switch (_Type)
			{
				case VBE.vbext_ComponentType.vbext_ct_StdModule:	return VBA_MODULE;
				case VBE.vbext_ComponentType.vbext_ct_Document:		return VBA_CLASS;
				case VBE.vbext_ComponentType.vbext_ct_ClassModule:	return VBA_CLASS;
				case VBE.vbext_ComponentType.vbext_ct_MSForm:		return VBA_FORM;
			}

			return "";
		}
	}

	public class VBEOperater
	{
		public VBEOperater()
		{
			_ComponentTypes.Add(VBE.vbext_ComponentType.vbext_ct_StdModule);
			_ComponentTypes.Add(VBE.vbext_ComponentType.vbext_ct_ClassModule);
		}

		/// <summary>
		/// モジュールをクリアする。
		/// </summary>
		/// <param name="_Project">モジュールをクリアするプロジェクト</param>
		public void Clear(VBE.VBProject _Project)
		{
			if (_Project is null)	throw new ArgumentNullException(nameof(_Project));
			      
			foreach (VBE.VBComponent _Component in _Project.VBComponents)
			{
				ClearModule(_Project, _Component);
			}

			//消去の成否の確認
			if (IsModuleClearSuccess(_Project)) throw new Exception("モジュールの削除に失敗しました");
		}

		/// <summary>
		/// モジュールを取り込む
		/// </summary>
		/// <param name="_Project">モジュールを取り込ませるプロジェクト</param>
		/// <param name="_PathList">取り込みたいモジュールの絶対パスの一覧</param>
		public void Import(VBE.VBProject _Project, List<string> _PathList)
		{
			if (_Project is null)	throw new ArgumentNullException(nameof(_Project));
			if (_PathList is null)	throw new ArgumentNullException(nameof(_PathList));

			foreach (string _Path in _PathList)
			{
				ImportModule(_Project, _Path);
			}
		}

		/// <summary>
		/// モジュールを取り込む
		/// </summary>
		/// <param name="_Project">モジュールを取り込ませるプロジェクト</param>
		/// <param name="_Path">取り込みたいモジュールの絶対パス</param>
		public void Import(VBE.VBProject _Project, in string _Path)
		{
			if (_Project is null) throw new ArgumentNullException(nameof(_Project));

			ImportModule(_Project, _Path);
		}


		/// <summary>
		/// モジュールを外部に書き出す
		/// </summary>
		/// <param name="_Project">書き出すディレクトリ</param>
		/// <param name="_Path">書き出し先のディレクトリ</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public List<string> Export(VBE.VBProject _Project, in string _Path)
		{
			if (_Project is null)					throw new ArgumentNullException(nameof(_Project));
			if (_Project.VBComponents.Count < 1)	throw new ArgumentException("モジュールが存在しないプロジェクトです", nameof(_Project));
			if (string.IsNullOrWhiteSpace(_Path))	throw new ArgumentException("無効なパスです", nameof(_Path));
			if (Directory.Exists(_Path))			throw new DirectoryNotFoundException("[" + _Path + "]は無効なパスです");

			List<string> _ModulePathes = new List<string>();
			
			foreach (VBE.VBComponent _Component in _Project.VBComponents)
			{
				string _FileName = ExportModule(_Component, _Path);

				if (_FileName != "") _ModulePathes.Add(_FileName);
			}

			return _ModulePathes;
		}

		/// <summary>
		/// モジュールを外部に書き出す
		/// </summary>
		/// <param name="_Project">書き出すディレクトリ</param>
		/// <param name="_Path">書き出し先のディレクトリ</param>
		/// <param name="_ModuleName">書き出したいモジュール名</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public string Export(VBE.VBProject _Project, in string _Path, in string _ModuleName)
		{
			if (_Project is null)						throw new ArgumentNullException(nameof(_Project), "プロジェクトが存在しません");
			if (_Project.VBComponents.Count < 1)		throw new ArgumentException("モジュールが存在しないプロジェクトです", nameof(_Project));
			if (string.IsNullOrWhiteSpace(_Path))		throw new ArgumentNullException(nameof(_Path), "無効なパスです");
			if (Directory.Exists(_Path))				throw new DirectoryNotFoundException("[" + _Path + "]は無効なパスです");
			if (string.IsNullOrWhiteSpace(_ModuleName))	throw new ArgumentException("無効なモジュール名です", nameof(_ModuleName));

			string _FileName;
			try
			{
				_FileName = ExportModule(_Project.VBComponents.Item(_ModuleName), _Path);
			}
			catch (IndexOutOfRangeException e)
			{
				throw new IndexOutOfRangeException("[" + _ModuleName + "]は存在しないモジュールです。", e);
			}

			return ((_FileName != "") ? _FileName : throw new Exception("[" + _ModuleName + "]のExportに失敗しました。"));
		}

		//*************************************************************************************************//
		//private method
		//*************************************************************************************************//

		private void ClearModule(VBE.VBProject _Project, VBE.VBComponent _Component)
		{
			if (_ComponentTypes.Contains(_Component.Type))
			{
				if (VBE.vbext_ComponentType.vbext_ct_Document == _Component.Type)
				{
					ClearDocumentModule(_Component);
				}
				else
				{
					_Project.VBComponents.Remove(_Component);
				}
			}
		}

		private void ClearDocumentModule(VBE.VBComponent _Component)
		{
			int line_count = _Component.CodeModule.CountOfLines;
			_Component.CodeModule.DeleteLines(1, line_count);
		}

		/// <summary>
		/// モジュールのクリアに成功しましたか?
		/// </summary>
		/// <param name="_Project">確認するプロジェクト</param>
		/// <returns>Yes or No</returns>
		private bool IsModuleClearSuccess(VBE.VBProject _Project)
		{
			foreach (VBE.VBComponent component in _Project.VBComponents)
			{
				if (_ComponentTypes.Contains(component.Type)) return false;
			}

			return true;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="_Project"></param>
		/// <param name="_ModuleName"></param>
		/// <returns></returns>
		private bool IsModuleExits(VBE.VBProject _Project, in string _ModuleName)
		{
			try
			{
				_Project.VBComponents.Item(_ModuleName);
			}
			catch (IndexOutOfRangeException)
			{
				return false;
			}

			return true;
		}

		/// <summary>
		/// モジュールのコードを切り取る。
		/// </summary>
		/// <param name="_Component">切り取るモジュール</param>
		/// <returns>クリア前のコード</returns>
		private string CutDoucumentModule(VBE.VBComponent _Component)
		{
			int _LineCount = _Component.CodeModule.CountOfLines;
			string _Code = _Component.CodeModule.get_Lines(1, _LineCount);
			_Component.CodeModule.DeleteLines(1, _LineCount);

			return _Code;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="_Project"></param>
		/// <param name="_Path"></param>
		private void ImportModule(VBE.VBProject _Project, in string _Path)
		{
			if (string.IsNullOrWhiteSpace(_Path))	throw new ArgumentException("無効なパスです", nameof(_Path));
			if (!File.Exists(_Path))				throw new FileNotFoundException("[" + _Path + "]は無効なパスです");

			var _ModuleName = Path.GetFileNameWithoutExtension(_Path);
			if (IsModuleExits(_Project, _ModuleName))
			{
				var _Component = _Project.VBComponents.Item(_ModuleName);

				if (_Component.Type == VBE.vbext_ComponentType.vbext_ct_Document)
				{
					if (!_ComponentTypes.Contains(VBE.vbext_ComponentType.vbext_ct_Document)) return;

					ImportDocumentModule(_Component, _Path);
				}
				else
				{
					ClearModule(_Project, _Component);
				}
			}

			_Project.VBComponents.Import(_Path);
		}

		/// <summary>
		/// ドキュメントモジュールをインポートする
		/// </summary>
		/// <param name="_Project"></param>
		/// <param name="_Path"></param>
		private void ImportDocumentModule(VBE.VBComponent _Component, in string _Path)
		{
			string _Code = CutDoucumentModule(_Component);
			try
			{
				_Component.CodeModule.AddFromString(new StreamReader(_Path, Encoding.GetEncoding("Shift_JIS")).ReadToEnd());
			}
			catch
			{
				if (_Code != "") _Component.CodeModule.AddFromString(_Code);
				
				throw new Exception("[" + _Path + "]のインポートに失敗しました");
			}
		}

		/// <summary>
		/// モジュールをエクスポートする
		/// </summary>
		/// <param name="_Component">エクスポートするモジュール</param>
		/// <param name="_Path">エクスポート先のディレクトリ</param>
		/// <returns>エクスポートしたモジュールのパス</returns>
		private string ExportModule(VBE.VBComponent _Component, in string _Path)
		{
			if (_ComponentTypes.Contains(_Component.Type))
			{
				string _FileName = _Path + "\\" + _Component.Name + _Extension.Convert(_Component.Type);

				_Component.Export(_FileName);
				
				return _FileName;
			}

			return "";
		}

		private List<VBE.vbext_ComponentType> _ComponentTypes = new List<VBE.vbext_ComponentType>();
		private ModuleExtension _Extension = new ModuleExtension();

		private const string THIS_WORKBOOK = "ThisWorkbook";
	}
}

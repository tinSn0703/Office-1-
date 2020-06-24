using System.Collections.Generic;

namespace SuzuOffice
{
	/************************************************************************/

	class ExtensionChecker
	{
		private List<string> _ExtensionList;

		public ExtensionChecker(List<string> _List)
		{
			_ExtensionList = _List;
		}

		public bool IsList(string _Extension)
		{
			return _ExtensionList.Contains(_Extension);
		}

		/// <summary>
		/// 正しい拡張子ですか?
		/// </summary>
		/// <param name="_Checker">正しい拡張子のリスト</param>
		/// <param name="_FilePath">調べるファイルのパス</param>
		/// <returns></returns>
		public bool IsExtensionCorrect(string _FilePath)
		{
			return IsList(System.IO.Path.GetExtension(_FilePath));
		}
	}

	/************************************************************************/
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuzuOffice.Excel
{
	using Excel = Microsoft.Office.Interop.Excel;
	using VBE = Microsoft.Vbe.Interop;
	public class VBEOperator
	{
		public VBEOperator()
		{
			_Project = new VBEOperater();
		}

		/// <summary>
		/// モジュールをクリアする。
		/// </summary>
		/// <param name="_Book">モジュールをクリアするブック</param>
		public void ClearModules(Excel.Workbook _Book)
		{
			if (_Book is null)	throw new ArgumentNullException(nameof(_Book));

			this._Project.ClearModules(_Book.VBProject);
		}

		/// <summary>
		/// モジュールを取り込む
		/// </summary>
		/// <param name="_Book">モジュールを取り込ませるブック</param>
		/// <param name="_PathList">取り込みたいモジュールの絶対パス</param>
		/// <exception cref="ArgumentNullException"></exception>
		public void ImportModules(Excel.Workbook _Book, List<string> _PathList)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));

			this._Project.ImportModules(_Book.VBProject, _PathList);
		}

		/// <summary>
		/// モジュールを外部に書き出す
		/// </summary>
		/// <param name="_Book">モジュールを書き出したいブック</param>
		/// <param name="_Path">書き出し先のディレクトリ</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public List<string> ExportModules(Excel.Workbook _Book, in string _Path)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));

			return this._Project.ExportModules(_Book.VBProject, _Path);
		}

		private VBEOperater _Project;
	}
}

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

		/// <summary>モジュールをクリアする。</summary>
		/// <param name="_Book">モジュールをクリアするブック</param>
		public void Clear(in Excel.Workbook _Book)
		{
			if (_Book is null)	throw new ArgumentNullException(nameof(_Book));

			this._Project.Clear(_Book.VBProject);
		}

		/// <summary>モジュールを取り込む</summary>
		/// <param name="_Book">モジュールを取り込ませるブック</param>
		/// <param name="_PathList">取り込みたいモジュールの絶対パスの一覧</param>
		/// <exception cref="ArgumentNullException"></exception>
		/// <exception cref="ArgumentException"></exception>
		public void Import(in Excel.Workbook _Book, IReadOnlyCollection<string> _PathList)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));
			if (_PathList is null) throw new ArgumentNullException(nameof(_PathList));
			if (_PathList.Count() < 1) throw new ArgumentException("一覧が空です", nameof(_PathList));

			this._Project.Import(_Book.VBProject, _PathList);
		}

		/// <summary>モジュールを取り込む</summary>
		/// <param name="_Book">モジュールを取り込ませるブック</param>
		/// <param name="_Path">取り込みたいモジュールの絶対パス</param>
		/// <exception cref="ArgumentNullException"></exception>
		public void Import(in Excel.Workbook _Book, in string _Path)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));

			this._Project.Import(_Book.VBProject, _Path);
		}

		/// <summary>モジュールを外部に書き出す</summary>
		/// <param name="_Book">モジュールを書き出したいブック</param>
		/// <param name="_Path">書き出し先のディレクトリ</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public List<string> Export(in Excel.Workbook _Book, in string _Path)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));

			return this._Project.Export(_Book.VBProject, _Path);
		}

		/// <summary>モジュールを外部に書き出す</summary>
		/// <param name="_Book">モジュールを書き出したいブック</param>
		/// <param name="_Path">書き出し先のディレクトリ</param>
		/// <param name="_ModuleName">書き出したいモジュール名</param>
		/// <returns>書き出したディレクトリのパス</returns>
		public string Export(in Excel.Workbook _Book, in string _Path, in string _ModuleName)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));

			return this._Project.Export(_Book.VBProject, _Path, _ModuleName);
		}

		private VBEOperater _Project;
	}
}

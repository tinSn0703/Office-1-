using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace SuzuOffice.Excel
{
	using Excel = Microsoft.Office.Interop.Excel;
	public class ExcelOperator
	{
		public ExcelOperator()	{}

		/// <summary>Excel Bookを開く</summary>
		/// <param name="_Excel">Excel Application</param>
		/// <param name="_FilePath">開くBookのパス</param>
		/// <returns></returns>
		public ExcelBookAccessor Open(ExcelAppAccessor _Excel, in string _FilePath)
		{
			try
			{
				if (this.IsFileOpened(_FilePath)) return GetRunningApp(_Excel, _FilePath);
			}
			catch (Exception e)
			{
				if (e.Message.IndexOf(_FilePath) < 0) throw;
				Console.WriteLine(e);
			}

			return new ExcelBookAccessor(_Excel.Books.Open(_FilePath));
		}

		/// <summary>Excel Bookを新規作成する</summary>
		/// <param name="_Excel">Excel Application</param>
		/// <param name="_WorkBookPath">新規作成先のパス</param>
		/// <returns></returns>
		public ExcelBookAccessor Add(ExcelAppAccessor _Excel, in string _WorkBookPath)
		{
			Excel.Workbook _Book = null;

			try
			{
				_Book = _Excel.Books.Add();
				_Book.SaveAs(_WorkBookPath);
			}
			catch
			{
				this.ReleaseObject(_Book);
				throw;
			}

			return new ExcelBookAccessor(_Book);
		}

		public ExcelSheetAccessor GetSheet(ExcelBookAccessor _Book, object _SheetIndex)
		{
			return new ExcelSheetAccessor(_Book.Sheets[_SheetIndex]);
		}

		public ExcelSheetAccessor AddSheet(ExcelBookAccessor _Book, string _SheetName)
		{
			Excel.Worksheet _Sheet = null;

			try
			{
				_Sheet = _Book.Sheets.Add();
				_Sheet.Name = _SheetName;
			}
			catch
			{
				this.ReleaseObject(_Sheet);
				throw;
			}

			return new ExcelSheetAccessor(_Sheet);
		}

		public void Save(ExcelBookAccessor _Book, string _FilePath = "")
		{
			if (_FilePath != "")
			{
				_Book.Book.SaveAs(_FilePath);
				return;
			}

			if (_Book.Book.Path != "")
			{
				_Book.Book.Save();
				return;
			}
		}

		public void Close(Excel.Workbook _Book)
		{
			_Book.Close();
		}

		//*************************************************************************************************//
		//private method
		//*************************************************************************************************//

		/// <summary>指定したファイルは、既に開かれていますか?</summary>
		/// <param name="_FilePath">調べるファイルのパス</param>
		/// <returns></returns>
		private bool IsFileOpened(in string _FilePath)
		{
			string _FileName = System.IO.Path.GetFileName(_FilePath); //ファイル名を取り出す
			foreach (Process _Process in Process.GetProcesses())
			{
				//関係ないプロセス。スキップ
				if (_Process.MainWindowTitle.Length == 0) continue;

				//現在開かれているプロセス名と比較し、ファイルが開かれているか確認する
				if (_Process.MainWindowTitle.IndexOf(_FileName) >= 0) return true;
			}

			return false;
		}

		/// <summary>実行中のExcelブックを取得する</summary>
		/// <param name="_FilePath">実行中ブックのパス</param>
		/// <returns>実行中のブック</returns>
		private ExcelBookAccessor GetRunningApp(ExcelAppAccessor _Excel, in string _FilePath)
		{
			Excel.Workbook _Book = null;

			try
			{
				_Book = Marshal.BindToMoniker(_FilePath) as Excel.Workbook;

				if (_Book == null) throw new System.Exception("[" + _FilePath + "]の確保に失敗しました");

				_Excel.Open(_Book.Application);
			}
			catch
			{
				ReleaseObject(_Book);
				throw;
			}

			return new ExcelBookAccessor(_Book);
		}

		/// <summary>Objectを開放する</summary>
		private void ReleaseObject(object _Obj)
		{
			if (_Obj != null)
			{
				while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_Obj) > 0) ;
				_Obj = null;
			}
		}

		//*************************************************************************************************//
		//private method
		//*************************************************************************************************//

	}
}

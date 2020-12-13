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

		/// <summary>
		/// Excel Bookを開く。
		/// </summary>
		/// <param name="_Excel">Excel Application</param>
		/// <param name="_WorkBookPath">開くBookのパス</param>
		/// <returns></returns>
		public Excel.Workbook Open(ref ExcelAppAccessor _Excel, in string _WorkBookPath)
		{
			try
			{
				if (this.IsFileOpened(_WorkBookPath)) return GetRunningApp(ref _Excel, _WorkBookPath);
			}
			catch (Exception e)
			{
				if (e.Message.IndexOf(_WorkBookPath) < 0) throw;
				Console.WriteLine(e);
			}

			return _Excel.Books.Open(_WorkBookPath);
		}

		/// <summary>
		/// Excel Bookを新規作成する。
		/// </summary>
		/// <param name="_Excel">Excel Application</param>
		/// <param name="_WorkBookPath">新規作成先のパス</param>
		/// <returns></returns>
		public Excel.Workbook Add(ExcelAppAccessor _Excel, in string _WorkBookPath)
		{
			Microsoft.Office.Interop.Excel.Workbook _ExcelBook = null;

			try
			{
				_ExcelBook = _Excel.Books.Add();
				_ExcelBook.SaveAs(_WorkBookPath);

				return _ExcelBook;
			}
			catch (Exception e)
			{
				this.ReleaseObject(_ExcelBook);
				throw e;
			}
		}

		public void GetSheet(Excel.Workbook _Book)
		{
			_Book.Close();
		}

		public void Close(Excel.Workbook _Book)
		{
			_Book.Close();
		}

		//*************************************************************************************************//
		//private method
		//*************************************************************************************************//

		/// <summary>
		/// 指定したファイルは、既に開かれていますか?
		/// </summary>
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

		/// <summary>
		/// 実行中のExcelブックを取得する
		/// </summary>
		/// <param name="_FilePath">実行中ブックのパス</param>
		/// <returns>実行中のブック</returns>
		private Excel.Workbook GetRunningApp(ref ExcelAppAccessor _Excel, in string _FilePath)
		{
			Excel.Workbook _ExcelBook = null;

			try
			{
				_ExcelBook = Marshal.BindToMoniker(_FilePath) as Excel.Workbook;

				if (_ExcelBook == null) throw new System.Exception("[" + _FilePath + "]の確保に失敗しました");

				_Excel.Open(_ExcelBook.Application);

				return _ExcelBook;
			}
			catch
			{
				ReleaseObject(_ExcelBook);
				throw;
			}
		}

		/// <summary>
		/// Objectを開放する
		/// </summary>
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

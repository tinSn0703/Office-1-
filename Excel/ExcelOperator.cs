using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuzuOffice.Excel
{
	using Excel = Microsoft.Office.Interop.Excel;
	public class ExcelOperator : IDisposable
	{
		public ExcelOperator()
		{
			if (_Application == null)
			{
				_Application = new Excel.Application();
				_WorkBooks = _Application.Workbooks;
			}
		}

		public Excel.Workbook Open(string _WorkBookPath)
		{
			return _WorkBooks.Open(_WorkBookPath);
		}

		public Excel.Workbook Add(string _WorkBookPath)
		{
			Microsoft.Office.Interop.Excel.Workbook _ExcelBook = null;

			try
			{
				_ExcelBook = _WorkBooks.Add();
				_ExcelBook.SaveAs(_WorkBookPath);

				return _ExcelBook;
			}
			catch (Exception e)
			{
				this.ReleaseObject(_ExcelBook);
				throw e;
			}
		}

		public void Close(Excel.Workbook _Document)
		{
			_Document.Close();
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

		/// <summary>
		/// 確保されたリソースの解放
		/// </summary>
		/// <param name="_Disposing">GCが解放してくれるリソースを開放するかしないか</param>
		protected virtual void Dispose(bool _Disposing)
		{
			if (!_DisposeValue)
			{
				if (_Disposing) { }

				if (_ReferenceCounter < 2)
				{
					this.ReleaseObject(_WorkBooks);
					this.ReleaseObject(_Application);
				}
				else
				{
					_ReferenceCounter -= 1;
				}

				_DisposeValue = true;
			}
		}

		/// <summary>
		/// 確保されたリソースの解放
		/// </summary>
		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// デストラクタではなくファイナライザ。
		/// C++と構文が同じだから勘違いしてたぞ
		/// </summary>
		~ExcelOperator()
		{
			this.Dispose(false);
		}

		private bool _DisposeValue = false;

		static private int _ReferenceCounter = 0;
		static private Excel.Application _Application = null;
		static private Excel.Workbooks _WorkBooks = null;
	}
}

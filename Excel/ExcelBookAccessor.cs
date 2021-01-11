using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace SuzuOffice.Excel
{
	using Excel = Microsoft.Office.Interop.Excel;

	/// <summary>
	/// Excel book へのアクセスを担当する。
	/// </summary>
	public class ExcelBookAccessor : IDisposable
    {
		//*************************************************************************************************//
		//public method
		//*************************************************************************************************//
		
		public ExcelBookAccessor()
		{
			this._Book = null;
			this._Sheets = null;
		}

		public ExcelBookAccessor(Excel.Workbook _ExcelBook)
		{
			this.Open(_ExcelBook);
		}

		/// <summary>アクセス先のExcelブックを設定する</summary>
		/// <param name="_Book">アクセス先のブック</param>
		/// <exception cref="ArgumentNullException">アクセス先のブックがnullだった場合</exception>
		public void Open(Excel.Workbook _Book)
		{
			if (_Book is null) throw new ArgumentNullException(nameof(_Book));

			this.Close();

			this._Book = _Book;
			this._Sheets = this._Book.Sheets;
		}
		
		/// <summary>Excelブックを閉じる</summary>
		public void Close()
		{
			if (_Book != null)
			{
				_Book.Close();
				this.Dispose();
			}
		}

		/// <summary>リソースを開放する</summary>
		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		~ExcelBookAccessor()
		{
			this.Dispose(false);
		}

		public ref Excel.Workbook Book { get => ref _Book; }
		public ref Excel.Sheets Sheets { get => ref _Sheets; }

		//*************************************************************************************************//
		//private method
		//*************************************************************************************************//

		protected virtual void Dispose(bool _Disposing)
		{
			if (!_DisposeValue)
			{
				if (_Disposing)	{}

				// Sheets解放
				if (_Sheets != null)
				{
					while (Marshal.ReleaseComObject(_Sheets) > 0);
					_Sheets = null;
				}

				// Book解放
				if (_Book != null)
				{
					while (Marshal.ReleaseComObject(_Book) > 0);
					_Book = null;
				}
			}
		}

		//*************************************************************************************************//
		//private field
		//*************************************************************************************************//

		private bool _DisposeValue = false;
		
		private Excel.Workbook _Book = null;
		private Excel.Sheets _Sheets = null;
	}
}

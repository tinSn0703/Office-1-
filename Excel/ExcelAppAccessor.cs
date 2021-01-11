using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace SuzuOffice.Excel
{
	using Excel = Microsoft.Office.Interop.Excel;

	/// <summary>
	/// Excelのアプリケーションをポンポン開くとおいしくないから、一度だけにしたい。
	/// とりあえず作った。まだよくわからないので試行していく。
	/// Excelのリリースが問題にならないようにラッピングする。
	/// </summary>
	public class ExcelAppAccessor : IDisposable
	{
		//----------------------------------------------------------------------//
		//const
		//----------------------------------------------------------------------//

		readonly string[] EXCEL_EXTENSION = { "xls", "xlsx", "xlsm", "xlt", "xltx", "xltm" };

		//----------------------------------------------------------------------//
		//function
		//----------------------------------------------------------------------//

		public ExcelAppAccessor()
		{
			_ReferenceCounter += 1;
		}

		/// <summary>Excelを開く</summary>
		/// <param name="_SecuredApp">既に開かれているApplicationがある場合</param>
		public void Open(Excel.Application _SecuredApp = null)
		{
			if (_App != null) return;

			this.Release();

			if (_SecuredApp == null) _App = new Excel.Application();
			else _App = _SecuredApp;

			_Books = _App.Workbooks;
		}

		/// <summary>Excelを閉じる</summary>
		public void Close()
		{
			if (_App != null)
			{
				_App.Quit();

				this.Dispose();
			}
		}

		/// <summary>確保されたリソースの解放</summary>
		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// デストラクタではなくファイナライザ。
		/// C++と構文が同じだから勘違いしてたぞ
		/// </summary>
		~ExcelAppAccessor()
		{
			this.Dispose(false);
		}

		public ref Application Application { get => ref _App; }
		public ref Workbooks Books { get => ref _Books; }

		//--------------------------------------------------------------------------------------//

		/// <summary>
		/// オブジェクトを開放する
		/// </summary>
		private void Release()
		{
			if (_Books != null)
			{
				while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_Books) > 0);
				_Books = null;
			}

			if (_App != null)
			{
				while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_App) > 0);
				_App = null;
			}
		}

		/// <summary>確保されたリソースの解放</summary>
		/// <param name="_Disposing">GCが解放してくれるリソースを開放するかしないか</param>
		protected virtual void Dispose(in bool _Disposing)
		{
			if (_Disposed) return;
			if (_Disposing) { }

			if (_ReferenceCounter <= 1)	this.Release();
			else						_ReferenceCounter -= 1;

			_Disposed = true;
		}

		private bool _Disposed = false;

		static private int _ReferenceCounter = 0;
		static private Excel.Application _App = null;
		static private Excel.Workbooks _Books = null;
	}
}

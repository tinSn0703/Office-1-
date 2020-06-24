using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Visio = Microsoft.Office.Interop.Visio;

namespace SuzuOffice
{
	/************************************************************************/	

	class VisioOpenerSelecter
	{
		public VisioOpenerSelecter()
		{
		}
	}

	class VisioAppAccessor : IDisposable
	{
		//----------------------------------------------------------------------//
		//public function
		//----------------------------------------------------------------------//

		public VisioAppAccessor()
		{
			_ReferenceCounter += 1;
		}

		/// <summary>既存のドキュメントを開く</summary>
		/// <param name="_FilePath">開きたいドキュメントの保存先のパス</param>
		/// <returns>開いたドキュメント</returns>
		public Visio.Document Open(string _FilePath)
		{
			try
			{
				if (IsDocumnetOpen(Path.GetFileName(_FilePath)))
				{
					return _Docs[Path.GetFileName(_FilePath)];
				}

				return _Opener.Open(ref _App, ref _Docs, _FilePath);
			}
			catch (Exception e)
			{
				ReleaseApplication();
				throw e;
			}
		}
		
		/// <summary>新しいドキュメントを開く</summary>
		/// <param name="_FilePath">追加したドキュメントの保存先のパス</param>
		/// <returns>追加したドキュメント</returns>
		public Visio.Document Add(string _FilePath)
		{
			try
			{
				return _Opener.Add(ref _App, ref _Docs, _FilePath);
			}
			catch (Exception e)
			{
				ReleaseApplication();
				throw e;
			}
		}

		/// <summary>
		/// アプリを閉じる
		/// </summary>
		public void Close()
		{
			if (_App != null) _App.Quit();
		}

		/// <summary>
		/// 表示する
		/// </summary>
		/// <param name="_IsVisuble"></param>
		public void Visible(bool _IsVisuble)
		{
			_App.Visible = _IsVisuble;
		}

		/// <summary>
		/// 確保されたリソースの解放
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		~VisioAppAccessor()
		{
			this.Dispose(false);
		}

		//----------------------------------------------------------------------//
		//propaty
		//----------------------------------------------------------------------//

		public Visio.Application Application => _App;

		public Visio.Documents Documents => _Docs;

		public VisioOpener Opener
		{
			set => _Opener = value;
			get => _Opener;
		}

		//----------------------------------------------------------------------//
		//private function
		//----------------------------------------------------------------------//

		private bool IsDocumnetOpen(string _Name)
		{
			if (_Docs == null) return false;

			_Docs.GetNames(out Array _DocNames);
			foreach (string _DocName in _DocNames)
			{
				if (_Name == _DocName)
				{
					return true;
				}
			}

			return false;
		}

		/// <summary>
		/// Application Objectを開放する
		/// </summary>
		private void ReleaseApplication()
		{
			if (_Docs != null)
			{
				while (Marshal.ReleaseComObject(_Docs) > 0);
				_Docs = null;
			}

			if (_App != null)
			{
				while (Marshal.ReleaseComObject(_App) > 0) ;
				_App = null;
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
					this.ReleaseApplication();
				}
				
				_ReferenceCounter -= 1;

				_DisposeValue = true;
			}
		}

		//----------------------------------------------------------------------//
		//Field
		//----------------------------------------------------------------------//

		private VisioOpener _Opener;

		private bool _DisposeValue = false;

		static private int _ReferenceCounter = 0;
		static private Visio.Application _App = null;
		static private Visio.Documents _Docs = null;
	}

	/************************************************************************/
}

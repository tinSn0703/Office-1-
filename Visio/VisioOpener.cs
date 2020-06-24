using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Visio = Microsoft.Office.Interop.Visio;

namespace SuzuOffice
{
	/************************************************************************/

	//----------------------------------------------------------------------//

	/// <summary>
	/// Visioのファイルの開き方を示すオブジェクト。
	/// </summary>
	abstract class VisioOpener
	{
		//----------------------------------------------------------------------//
		//function
		//----------------------------------------------------------------------//

		protected VisioOpener()
		{
			_OpenSaveArgs = 0;
		}

		/// <summary>
		/// Visioの既存のファイルを開く。存在しない場合はエラー
		/// </summary>
		/// <param name="_App"></param>
		/// <param name="_Docs"></param>
		/// <param name="_FilePath"></param>
		/// <returns></returns>
		public virtual Visio.Document Open(ref Visio.Application _App, ref Visio.Documents _Docs, string _FilePath)
		{
			Visio.Document _Doc = null;
			try
			{
				TheOpenAppCapture<Visio.Document> _Capture = new TheOpenAppCapture<Visio.Document>(_FilePath);
				if (_Capture.IsFileOpened())
				{
					_Doc = _Capture.GetRunningApp();
					SecureApp(ref _App, ref _Docs, _Doc.Application);

					return _Doc;
				}
			}
			catch (Exception e)
			{
				if (_Doc != null) while (Marshal.ReleaseComObject(_Doc) > 0);
				if (e.Message.IndexOf(_FilePath) < 0) throw e;
			}

			try
			{
				SecureApp(ref _App, ref _Docs);
				return _OpenSaveArgs == 0 ? _Docs.Open(_FilePath) : _Docs.OpenEx(_FilePath, _OpenSaveArgs);
			}
			catch (Exception e)
			{
				throw e;
			}
		}

		/// <summary>
		/// Visioのファイルを新規作成して開く。既にファイルがある場合も新規作成される。
		/// </summary>
		/// <param name="_App"></param>
		/// <param name="_Docs"></param>
		/// <param name="_FilePath"></param>
		/// <returns></returns>
		public virtual Visio.Document Add(ref Visio.Application _App, ref Visio.Documents _Docs, string _FilePath)
		{
			Visio.Document _Doc = null;
			try
			{
				SecureApp(ref _App, ref _Docs);

				_Doc = _OpenSaveArgs == 0 ? _Docs.Add("") : _Docs.AddEx("", Visio.VisMeasurementSystem.visMSMetric, _OpenSaveArgs);
				_Doc.SaveAs(_FilePath);

				return _Doc;
			}
			catch (Exception e)
			{
				if (_Doc != null) while (Marshal.ReleaseComObject(_Doc) > 0);
				throw e;
			}
		}

		//----------------------------------------------------------------------//
		//propaty
		//----------------------------------------------------------------------//

		public short OpenSaveArgs
		{
			set => _OpenSaveArgs = value;
			get => _OpenSaveArgs;
		}

		//----------------------------------------------------------------------//
		//private function
		//----------------------------------------------------------------------//

		private void SecureApp(ref Visio.Application _App, ref Visio.Documents _Docs, Visio.Application _SecuredApp = null)
		{
			try
			{
				if (_App != null) return;

				if (_SecuredApp == null) _App = new Visio.Application();
				else _App = _SecuredApp;

				_Docs = _App.Documents;
			}
			catch (Exception e)
			{
				throw e;
			}
		}

		//----------------------------------------------------------------------//
		//Field
		//----------------------------------------------------------------------//

		private short _OpenSaveArgs;
	}

	//----------------------------------------------------------------------//

	class VisioDocumentOpener : VisioOpener
	{
		//----------------------------------------------------------------------//
		//function
		//----------------------------------------------------------------------//

		public VisioDocumentOpener() {}

		public override Visio.Document Open(ref Visio.Application _App, ref Visio.Documents _Docs, string _FilePath)
		{
			CheckFilePath(_FilePath);
			return base.Open(ref _App, ref _Docs, _FilePath);
		}

		public override Visio.Document Add(ref Visio.Application _App, ref Visio.Documents _Docs, string _FilePath)
		{
			CheckFilePath(_FilePath);
			return base.Add(ref _App, ref _Docs, _FilePath);
		}

		//----------------------------------------------------------------------//
		//private function
		//----------------------------------------------------------------------//

		private void CheckFilePath(string _FilePath)
		{
			if (!_Checker.IsExtensionCorrect(_FilePath)) throw new Exception(_FilePath + " はこのオブジェクトの対応外の拡張子です。");
			if (!File.Exists(_FilePath)) throw new FileNotFoundException(_FilePath);
		}

		//----------------------------------------------------------------------//
		//Field
		//----------------------------------------------------------------------//

		private readonly ExtensionChecker _Checker = new ExtensionChecker(new List<string> { ".vsd", ".vdx" });
	}

	//----------------------------------------------------------------------//

	class VisioMasterOpener : VisioOpener
	{
		//----------------------------------------------------------------------//
		//function
		//----------------------------------------------------------------------//

		public VisioMasterOpener()
		{
			OpenSaveArgs = (short)Visio.VisOpenSaveArgs.visOpenDocked;
		}

		public override Visio.Document Open(ref Visio.Application _App, ref Visio.Documents _Docs, string _FilePath)
		{
			CheckFilePath(_FilePath);
			return base.Open(ref _App, ref _Docs, _FilePath);
		}

		public override Visio.Document Add(ref Visio.Application _App, ref Visio.Documents _Docs, string _FilePath)
		{
			CheckFilePath(_FilePath);
			return base.Add(ref _App, ref _Docs, _FilePath);
		}

		//----------------------------------------------------------------------//
		//private function
		//----------------------------------------------------------------------//

		private void CheckFilePath(string _FilePath)
		{
			if (!_Checker.IsExtensionCorrect(_FilePath)) throw new Exception(_FilePath + " はこのオブジェクトの対応外の拡張子です。");
		}

		//----------------------------------------------------------------------//
		//Field
		//----------------------------------------------------------------------//

		private readonly ExtensionChecker _Checker = new ExtensionChecker(new List<string> { ".vss", ".vsx", ".VSS" });
	}

	//----------------------------------------------------------------------//

	/************************************************************************/
}

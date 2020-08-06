using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace SuzuOffice
{
	/// <summary>
	/// アプリが実行されていた場合、そのオブジェクトを取得する
	/// </summary>
	/// <typeparam name="AppType"></typeparam>
	class TheOpenAppCapture<AppType> where AppType : class
	{
		public TheOpenAppCapture()
		{
			_FilePath = "";
		}

		public TheOpenAppCapture(string _FilePath)
		{
			this.Set(_FilePath);
		}

		public void Set(string _FilePath)
		{
			if (!File.Exists(_FilePath)) throw new FileNotFoundException(this.GetType().Name + "\n\"" + _FilePath + "\" は存在しないディレクトリです");

			this._FilePath = _FilePath;
		}

		/// <summary>指定したファイルは、既に開かれていますか?</summary>
		/// <returns>開かれているかどうか</returns>
		public bool IsFileOpened()
		{
			if (_FilePath == "") throw new Exception(this.GetType().Name + "\nファイル名が登録されていない状態で呼び出されました");

			try
			{
				string _FileName = Path.GetFileName(_FilePath); //ファイル名を取り出す
				foreach (Process _Process in Process.GetProcesses())
				{
					//関係ないプロセス。スキップ
					if (_Process.MainWindowTitle.Length == 0) continue;

					//現在開かれているプロセス名と比較し、ファイルが開かれているか確認する
					if (_Process.MainWindowTitle.IndexOf(_FileName) >= 0)
					{
						_WasFileOpned = true;
						return true;
					}
				}

				_WasFileOpned = false;
				return false;
			}
			catch (Exception e)
			{
				throw new Exception(this.GetType().Name + "\nプロセスの確認に失敗しました", e);
			}
		}

		/// <summary>実行中のアプリを取得する</summary>
		/// <returns>実行中のアプリ</returns>
		public AppType GetRunningApp()
		{
			AppType _App = null;

			try
			{
				if (!_WasFileOpned) { if (!IsFileOpened()) return null; }

				_App = Marshal.BindToMoniker(_FilePath) as AppType;
				//if (_App == null) throw new Exception(_FilePath + "\n確保に失敗");

				return _App;
			}
			catch (Exception e)
			{
				if(_App != null) { while (Marshal.ReleaseComObject(_App) > 0);	}
				
				throw new Exception(this.GetType().Name + "実行中のアプリの確保に失敗しました", e);
			}
		}
		
		private string _FilePath;
		private bool _WasFileOpned;
	}
}

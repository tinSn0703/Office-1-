using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visio = Microsoft.Office.Interop.Visio;

namespace SuzuOffice
{
	class VisioAppFeatures
	{
		//----------------------------------------------------------------------//
		//const
		//----------------------------------------------------------------------//

		//----------------------------------------------------------------------//
		//public function
		//----------------------------------------------------------------------//

		public VisioAppFeatures()
		{
			this._App = null;
		}

		public VisioAppFeatures(VisioAppAccessor _App)
		{
			this._App = _App;
		}

		/// <summary>
		/// 単位系を変換する
		/// </summary>
		/// <param name="StringOrNumber">変換する値.数値もしくは文字列で指定する</param>
		/// <param name="UnitsIn">StringOrNumber の測定単位</param>
		/// <param name="UnitsOut">変換先の測定単位</param>
		/// <returns>変換した値</returns>
		public double ConvertResult(object StringOrNumber, object UnitsIn, object UnitsOut)
		{
			return _App.Application.ConvertResult(StringOrNumber, UnitsIn, UnitsOut);
		}

		/// <summary>
		/// Formatの書式に従って、変換した文字列を返す。
		/// </summary>
		/// <param name="StringOrNumber">変換する値.数値もしくは文字列で指定する</param>
		/// <param name="UnitsIn">StringOrNumber の測定単位</param>
		/// <param name="UnitsOut">変換先の測定単位</param>
		/// <param name="Format">変換する書式</param>
		/// <returns>指定の書式に変換した文字列</returns>
		public string FormatResult(object StringOrNumber, object UnitsIn, object UnitsOut, string Format)
		{
			return _App.Application.FormatResult(StringOrNumber, UnitsIn, UnitsOut, Format);
		}

		/// <summary>
		/// Formatの書式に従って、変換した文字列を返す。
		/// </summary>
		/// <param name="StringOrNumber">変換する値.数値もしくは文字列で指定する</param>
		/// <param name="UnitsIn">StringOrNumber の測定単位</param>
		/// <param name="UnitsOut">変換先の測定単位</param>
		/// <param name="Format">変換する書式</param>
		/// <param name="LangID">結果に用いる言語</param>
		/// <param name="CalendarID">結果に用いる日時設定</param>
		/// <returns>指定の書式に変換した文字列</returns>
		public string FormatResultEx(object StringOrNumber, object UnitsIn, object UnitsOut, string Format, int LangID = 0, int CalendarID = -1)
		{
			return _App.Application.FormatResultEx(StringOrNumber, UnitsIn, UnitsOut, Format, LangID, CalendarID);
		}

		//----------------------------------------------------------------------//
		//propaty
		//----------------------------------------------------------------------//

		public VisioAppAccessor AppAccesor
		{
			set { this._App = value; }
			get { return this._App; }
		}

		//----------------------------------------------------------------------//
		//Field
		//----------------------------------------------------------------------//

		private VisioAppAccessor _App;
	}
}

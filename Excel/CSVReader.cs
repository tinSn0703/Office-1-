using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuzuOffice
{
	class CSVReader
	{
		public CSVReader(System.IO.Stream _Stream)
		{
			this._Stream = _Stream;
		}

		public CSVReader(System.IO.Stream _Stream, System.Text.Encoding _Encoding)
		{
			this._Stream = _Stream;
			this._Encoding = _Encoding;
		}

		public void Open()
		{
			using (var _Reader = new System.IO.StreamReader(_Stream, _Encoding))
			{
				while ( ! _Reader.EndOfStream)
				{
					
				}
			}
		}

		private System.IO.Stream _Stream;
		private System.Text.Encoding _Encoding;
	}
}

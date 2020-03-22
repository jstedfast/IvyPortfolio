//
// YahooStockData.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2016-2020 Jeffrey Stedfast
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Globalization;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace IvyPortfolio
{
	public class YahooStockData : IStockData
	{
		const NumberStyles CsvValueStyle = NumberStyles.AllowDecimalPoint | NumberStyles.Integer;
		static readonly char[] CsvDelimeters = { ',' };

		readonly List<object[]> values;
		string[] headers;

		YahooStockData ()
		{
			values = new List<object[]> ();
		}

		public int Columns {
			get {
				return headers.Length;
			}
		}

		public int Rows {
			get {
				return values.Count;
			}
		}

		public string GetHeader (int column)
		{
			return headers[column];
		}

		public object GetValue (int row, int column)
		{
			return values[row][column];
		}

		public static async Task<YahooStockData> LoadAsync (Stream stream, CancellationToken cancellationToken)
		{
			var stockData = new YahooStockData ();

			using (var reader = new StreamReader (stream, Encoding.ASCII, false, 4096, true)) {
				var line = await reader.ReadLineAsync ();
				var tokens = line.Split (CsvDelimeters);

				stockData.headers = tokens;

				while ((line = await reader.ReadLineAsync ()) != null) {
					tokens = line.Split (CsvDelimeters);

					if (tokens.Length != stockData.headers.Length)
						Console.WriteLine ("Inconsistent number of columns: {0} vs {1}", tokens.Length, stockData.headers.Length);

					var values = new object[tokens.Length];

					if (tokens[0] != "null")
						values[0] = DateTime.Parse (tokens[0], CultureInfo.InvariantCulture);

					for (int i = 1; i < tokens.Length; i++) {
						if (tokens[i] == "null")
							continue;

						if (!double.TryParse (tokens[i], CsvValueStyle, CultureInfo.InvariantCulture, out var value)) {
							Console.WriteLine ("Failed to parse CSV double value: {0}", tokens[i]);
						} else {
							values[i] = value;
						}
					}

					stockData.values.Add (values);
				}
			}

			return stockData;
		}
	}
}

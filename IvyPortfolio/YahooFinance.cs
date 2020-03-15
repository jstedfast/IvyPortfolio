//
// YahooFinance.cs
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
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace IvyPortfolio
{
	public class YahooFinance : IFinancialService
	{
		static readonly Regex CrumbRegex = new Regex ("CrumbStore\":{\"crumb\":\"(?<crumb>.+?)\"}", RegexOptions.CultureInvariant | RegexOptions.Compiled);
		static readonly DateTime UnixEpoch = new DateTime (1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);

		HttpClient client;
		string crumb;

		public YahooFinance ()
		{
			var handler = new HttpClientHandler ();
			handler.CookieContainer = new CookieContainer ();
			handler.UseCookies = true;

			client = new HttpClient (handler);
		}

		public async Task<string> GetStockDescriptionAsync (string symbol, CancellationToken cancellationToken)
		{
			const string format = "https://finance.yahoo.com/quote/{0}?p={0}";
			var requestUri = string.Format (format, symbol);

			var html = await client.GetStringAsync (requestUri).ConfigureAwait (false);
			int startIndex, endIndex;

			// extract the cookie crumb
			var crumbs = CrumbRegex.Matches (html);
			string crumb;

			if (crumbs.Count > 0) {
				crumb = crumbs[0].Groups["crumb"].Value;
				crumb = crumb.Replace ("\\u002F", "/");
			} else {
				crumb = "xxxxxxxxxxx";
			}

			if ((endIndex = html.IndexOf ("(" + symbol + ")", StringComparison.Ordinal)) <= 0) {
				Console.WriteLine ("Failed to locate \"({0})\" in:", symbol);
				Console.WriteLine ("{0}", html);
				Console.WriteLine ();
				return string.Empty;
			}

			if (html[endIndex - 1] == ' ')
				endIndex--;

			startIndex = endIndex;

			while (startIndex > 0 && html[startIndex - 1] != '>')
				startIndex--;

			return html.Substring (startIndex, endIndex - startIndex).Replace ("&amp;", "&");
		}

		public async Task<string> GetStockDataAsync (string symbol, DateTime start, DateTime end, CancellationToken cancellationToken)
		{
			const string format = "https://query1.finance.yahoo.com/v7/finance/download/{0}?period1={1}&period2={2}&interval=1d&events=history&crumb={3}";
			var requestUri = string.Format (format, symbol, (start - UnixEpoch).TotalSeconds, (end - UnixEpoch).TotalSeconds, crumb);

			var data = await client.GetStringAsync (requestUri).ConfigureAwait (false);

			return data;
		}

		public void Dispose ()
		{
			client?.Dispose ();
			client = null;
		}
	}
}

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

using System.Net;

using Newtonsoft.Json.Linq;

namespace IvyPortfolio
{
	public class YahooFinance : IFinancialService
	{
		static readonly DateTime UnixEpoch = new DateTime (1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);

		HttpClient client;

		public YahooFinance ()
		{
			var handler = new HttpClientHandler ();
			handler.CookieContainer = new CookieContainer ();
			handler.UseCookies = true;

			client = new HttpClient (handler);
		}

		public async Task<string> GetStockDescriptionAsync (string symbol, CancellationToken cancellationToken)
		{
			var requestUri = $"https://finance.yahoo.com/quote/{symbol}/history";
			int startIndex, endIndex;
			string html;

			using (var request = new HttpRequestMessage (HttpMethod.Get, requestUri)) {
				request.Headers.Add ("Accept-Language", "en-US");
				request.Headers.Add ("Connection", "keep-alive");

				using (var response = await client.SendAsync (request, cancellationToken).ConfigureAwait (false)) {
					html = await response.Content.ReadAsStringAsync (cancellationToken).ConfigureAwait (false);

					if (!response.IsSuccessStatusCode) {

					}
				}
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

		static long SecondsSinceEpoch (DateTime dt)
		{
			return (long) (dt - UnixEpoch).TotalSeconds;
		}

		public async Task<IStockData> GetStockDataAsync (string symbol, DateTime start, DateTime end, CancellationToken cancellationToken)
		{
			const string format = "https://query1.finance.yahoo.com/v7/finance/download/{0}?period1={1}&period2={2}&interval=1d&events=history&includeAdjustedClose=true";
			var requestUri = string.Format (format, symbol, SecondsSinceEpoch (start), SecondsSinceEpoch (end));
			int retries = 0;

			// GET https://query1.finance.yahoo.com/v7/finance/download/AAPL?period1=1622906679&period2=1654442679&interval=1d&events=history&includeAdjustedClose=true HTTP/1.1
			// Host: query1.finance.yahoo.com
			// User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:100.0) Gecko/20100101 Firefox/100.0
			// Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8
			// Accept-Language: en-US,en;q=0.5
			// Accept-Encoding: gzip, deflate, br
			// Referer: https://finance.yahoo.com/quote/AAPL/history
			// Connection: keep-alive
			// Upgrade-Insecure-Requests: 1
			// Sec-Fetch-Dest: document
			// Sec-Fetch-Mode: navigate
			// Sec-Fetch-Site: same-site
			// Sec-Fetch-User: ?1
			// Sec-GPC: 1

			do {
				using (var request = new HttpRequestMessage (HttpMethod.Get, requestUri)) {
					request.Headers.Add ("Accept-Language", "en-US");
					request.Headers.Add ("Referer", $"https://https://finance.yahoo.com/quote/{symbol}/history");
					request.Headers.Add ("Connection", "keep-alive");

					using (var response = await client.SendAsync (request, cancellationToken).ConfigureAwait (false)) {
						using (var stream = await response.Content.ReadAsStreamAsync (cancellationToken).ConfigureAwait (false)) {
							if (response.IsSuccessStatusCode)
								return await YahooStockData.LoadAsync (stream, cancellationToken).ConfigureAwait (false);

							if (response.StatusCode == HttpStatusCode.Unauthorized && retries < 5) {
								await Task.Delay (1000).ConfigureAwait (false);
								retries++;
								continue;
							}

							using (var reader = new StreamReader (stream)) {
								var text = reader.ReadToEnd ();
								var json = JObject.Parse (text);

								var code = json.SelectToken ("finance.error.code").ToString ();
								var description = json.SelectToken ("finance.error.description").ToString ();

								throw new Exception ($"{code}: {description}");
							}
						}
					}
				}
			} while (true);
		}

		public void Dispose ()
		{
			client?.Dispose ();
			client = null;
		}
	}
}

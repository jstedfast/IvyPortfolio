//
// Program.cs
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
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

using Newtonsoft.Json;

namespace IvyPortfolio
{
	class Program
	{
		public static void Main (string[] args)
		{
			string dataDir, fileName;

			if (args.Length > 0) {
				dataDir = Path.GetDirectoryName (args[0]);
				fileName = args[0];
			} else {
				var home = Environment.GetFolderPath (Environment.SpecialFolder.Personal);
				dataDir = Path.Combine (home, "Dropbox", "IvyPortfolio");
				fileName = Path.Combine (dataDir, "portfolio.json");
			}

			var json = File.ReadAllText (fileName);
			var portfolio = JsonConvert.DeserializeObject<Portfolio> (json, new MovingAverageConverter ());

			if (portfolio.Documents == null || portfolio.Documents.Length == 0)
				return;

			GeneratePortfolioAsync (portfolio, dataDir, CancellationToken.None).GetAwaiter ().GetResult ();
		}

		static async Task GeneratePortfolioAsync (Portfolio portfolio, string dataDir, CancellationToken cancellationToken)
		{
			var accounts = new Dictionary<string, Account> ();

			foreach (var account in portfolio.Accounts) {
				accounts.Add (account.Name, account);
				await account.InitializeAsync (dataDir, cancellationToken).ConfigureAwait (false);
			}

			using (var client = new YahooFinance ()) {
				foreach (var document in portfolio.Documents) {
					if (string.IsNullOrEmpty (document.FileName) || document.Symbols == null || document.Symbols.Length == 0)
						continue;

					Array.Sort (document.MovingAverages);

					try {
						var workbook = await Excel.CreateSpreadsheetAsync (client, document, cancellationToken).ConfigureAwait (false);

						if (document.RemoteDocuments == null)
							continue;

						foreach (var remote in document.RemoteDocuments) {
							if (string.IsNullOrEmpty (remote.Account) || !accounts.TryGetValue (remote.Account, out var account))
								continue;

							await account.UpdateRemoteDocumentAsync (workbook, remote.Identifier, cancellationToken).ConfigureAwait (false);
						}
					} catch {
						continue;
					}
				}
			}
		}
	}
}

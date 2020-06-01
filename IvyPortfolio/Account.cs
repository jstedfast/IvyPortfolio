//
// Account.cs
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

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

using Newtonsoft.Json;

using NPOI.SS.UserModel;

namespace IvyPortfolio
{
	public class Account
	{
		[JsonProperty ("name")]
		public string Name { get; set; }

		[JsonProperty ("type")]
		public AccountType Type { get; set; }

		[JsonProperty ("credentials")]
		public string Credentials { get; set; }

		[JsonIgnore]
		public SheetsService GoogleSheetsService { get; private set; }

		public async Task InitializeAsync (string dataDir, CancellationToken cancellationToken)
		{
			if (string.IsNullOrEmpty (Credentials))
				return;

			switch (Type) {
			case AccountType.Google:
				GoogleSheetsService = await GetGoogleSheetServiceAsync (dataDir, cancellationToken).ConfigureAwait (false);
				break;
			case AccountType.Office365:
				// TODO: add support for this
				break;
			}
		}

		async Task<SheetsService> GetGoogleSheetServiceAsync (string dataDir, CancellationToken cancellationToken)
		{
			var filename = Path.Combine (dataDir, Credentials);

			if (!File.Exists (filename)) {
				filename = Credentials;

				if (!File.Exists (filename))
					return null;
			}

			using (var stream = File.OpenRead (filename)) {
				string[] scopes = { SheetsService.Scope.Spreadsheets };

				var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync (
					GoogleClientSecrets.Load (stream).Secrets, scopes, Name,
					cancellationToken).ConfigureAwait (false);

				// Create Google Sheets API service.
				return new SheetsService (new BaseClientService.Initializer {
					HttpClientInitializer = credential,
					ApplicationName = "IvyPortfolio"
				});
			}
		}

		public async Task UpdateRemoteDocumentAsync (IWorkbook workbook, string identifier, CancellationToken cancellationToken)
		{
			switch (Type) {
			case AccountType.Google:
				if (GoogleSheetsService != null)
					await UpdateGoogleSpreadsheetAsync (workbook, identifier, cancellationToken);
				break;
			case AccountType.Office365:
				// TODO: add support for this
				break;
			}
		}

		async Task UpdateGoogleSpreadsheetAsync (IWorkbook workbook, string identifier, CancellationToken cancellationToken)
		{
			if (string.IsNullOrEmpty (identifier))
				return;

			Console.WriteLine ("\tUpdating Google Spreadsheet: {0}", identifier);

			// Sheet1 = Dashboard, Sheet2 = Charts
			for (int index = 2; index < workbook.NumberOfSheets; index++) {
				var sheet = workbook.GetSheetAt (index);
				var values = new List<IList<object>> (sheet.LastRowNum);
				short maxColumnIndex = 0;

				for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++) {
					var row = sheet.GetRow (rowIndex);

					maxColumnIndex = Math.Max (maxColumnIndex, row.LastCellNum);
					values.Add (new object[row.LastCellNum]);

					for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++) {
						var cell = row.GetCell (columnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
						object value;

						if (cell != null) {
							switch (cell.CellType) {
							case CellType.String: value = cell.StringCellValue; break;
							case CellType.Numeric: value = cell.NumericCellValue; break;
							case CellType.Formula: value = "=" + cell.CellFormula; break;
							default: value = string.Empty; break;
							}
						} else {
							value = string.Empty;
						}

						values[rowIndex][columnIndex] = value;
					}
				}

				// Update the values in the Google Sheet
				var range = string.Format ("{0}!A1:{1}{2}", sheet.SheetName, (char) ('A' + maxColumnIndex), sheet.LastRowNum);
				var valueRange = new ValueRange { Values = values };

				var updateValuesRequest = GoogleSheetsService.Spreadsheets.Values.Update (valueRange, identifier, range);
				updateValuesRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

				try {
					var response = await updateValuesRequest.ExecuteAsync (cancellationToken).ConfigureAwait (false);
				} catch (Exception ex) {
					Console.WriteLine ("\tFailed: {0}", ex);
				}
			}
		}
	}
}
